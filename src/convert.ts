import {
  ApiError,
  type Conversion,
  type ConversionOptions,
  type ConversionStatus,
  type ConverterApi,
} from './api';

/**
 * One conversion from start to finish: the file goes up to an upload slot, Converter Studio
 * converts it on the account's credits (asking first above a price), and the finished
 * conversion comes back with its download link. Fetching the result is a separate step
 * (`fetchResult`), because a download in the browser needs no copy in the pane.
 */

/** Conversions up to this many credits go ahead without asking (most files cost 1). */
export const AUTO_CONFIRM_CREDITS = 3;
/** How long one status request waits on the server for the conversion to settle. */
const WAIT_SECONDS = 20;
/** How long one conversion is followed before the pane gives up. */
export const TIMEOUT_MS = 30 * 60_000;

export interface Job {
  readonly file: Blob;
  /** The original's name: the server reads the format from it. */
  readonly name: string;
  readonly to: string;
  readonly options: ConversionOptions;
}

export type Step =
  | { readonly step: 'uploading'; readonly sent: number; readonly total: number }
  | {
      readonly step: 'converting';
      readonly status: ConversionStatus;
      readonly progress: number | null;
    };

export interface Hooks {
  readonly onStep?: (step: Step) => void;
  /** The price is above `AUTO_CONFIRM_CREDITS`: resolve `true` to pay it and go ahead. */
  readonly confirmPrice: (credits: number) => Promise<boolean>;
  readonly signal?: AbortSignal;
  /** Tests make time pass faster. */
  readonly now?: () => number;
  readonly sleep?: (ms: number, signal?: AbortSignal) => Promise<void>;
  readonly idempotencyKey?: () => string;
}

export interface Result {
  readonly conversion: Conversion;
  readonly downloadUrl: string;
  readonly outputName: string;
  readonly size: number;
  readonly credits: number;
}

export async function convert(api: ConverterApi, job: Job, hooks: Hooks): Promise<Result> {
  const now = hooks.now ?? Date.now;
  const sleep = hooks.sleep ?? delay;
  const deadline = now() + TIMEOUT_MS;
  const key = (hooks.idempotencyKey ?? newIdempotencyKey)();

  hooks.onStep?.({ step: 'uploading', sent: 0, total: job.file.size });
  const upload = await api.createUpload(job.name, job.file.size, hooks.signal);
  await api.upload(upload, job.file, {
    signal: hooks.signal,
    onProgress: (sent) => hooks.onStep?.({ step: 'uploading', sent, total: job.file.size }),
  });
  const request = {
    uploadId: upload.id,
    to: job.to,
    options: job.options,
    maxCredits: AUTO_CONFIRM_CREDITS,
  };
  let conversion: Conversion;
  try {
    conversion = await api.convert(request, key, hooks.signal);
  } catch (error) {
    // The answer was lost on the way: the same key gives the same conversion, never a second.
    if (!(error instanceof ApiError) || error.code !== 'network') throw error;
    await sleep(2000, hooks.signal);
    conversion = await api.convert(request, key, hooks.signal);
  }

  let failures = 0;
  for (;;) {
    hooks.onStep?.({
      step: 'converting',
      status: conversion.status,
      progress: conversion.progress,
    });
    if (conversion.status === 'quoted') {
      const credits = conversion.credits ?? 0;
      if (!(await hooks.confirmPrice(credits)))
        throw new ApiError('declined', 'Canceled. Nothing was charged.', 0);
      conversion = await api.confirm(conversion.id, hooks.signal);
      continue;
    }
    if (settled(conversion.status)) break;
    if (hooks.signal?.aborted) throw new ApiError('canceled', 'Canceled.', 0);
    if (now() > deadline)
      throw new ApiError(
        'timeout',
        'The conversion is taking too long. Try again later, or with a smaller file.',
        0,
      );
    try {
      conversion = await api.status(conversion.id, WAIT_SECONDS, hooks.signal);
      failures = 0;
    } catch (error) {
      // A dropped connection while waiting: wait longer each time, then give up.
      if (error instanceof ApiError && error.code === 'network' && failures < 4) {
        await sleep(2000 * 2 ** failures, hooks.signal);
        failures += 1;
        continue;
      }
      throw error;
    }
  }

  if (conversion.status !== 'done' || !conversion.downloadUrl) {
    const error = conversion.error;
    if (error) throw new ApiError(error.code, error.message, 0, error.params ?? {});
    throw new ApiError(
      conversion.status,
      conversion.status === 'expired'
        ? 'The result expired before it could be used. Convert the file again.'
        : 'The conversion did not finish. No credits were charged for it.',
      0,
    );
  }
  return {
    conversion,
    downloadUrl: conversion.downloadUrl,
    outputName: conversion.outputName ?? job.name,
    size: conversion.outputSize ?? 0,
    credits: conversion.credits ?? 0,
  };
}

const settled = (status: ConversionStatus) =>
  status === 'done' || status === 'failed' || status === 'expired' || status === 'canceled';

/** The result's file. */
export function fetchResult(api: ConverterApi, result: Result, signal?: AbortSignal) {
  return api.download(result.downloadUrl, signal);
}

/**
 * Whether the result is one file of the format asked for. Some conversions give a ZIP of several
 * files (every page of a PDF as a picture): fine as a mail attachment or a download, but not
 * something to insert into a document.
 */
export function isSingleFile(result: Result, to: string): boolean {
  return to === 'zip' || !result.outputName.toLowerCase().endsWith('.zip');
}

/** A fresh key for one conversion request, so a repeated request does not convert twice. */
export function newIdempotencyKey(): string {
  if (typeof crypto.randomUUID === 'function') return `office-${crypto.randomUUID()}`;
  const bytes = crypto.getRandomValues(new Uint8Array(16));
  return `office-${Array.from(bytes, (b) => b.toString(16).padStart(2, '0')).join('')}`;
}

function delay(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(resolve, ms);
    signal?.addEventListener('abort', () => {
      clearTimeout(timer);
      reject(new ApiError('canceled', 'Canceled.', 0));
    });
  });
}
