/**
 * A small client for the Converter Studio REST API (https://converter.studio/developers).
 * Everything goes through `fetch`, which the caller passes in, so tests can stand in for it.
 *
 * The task pane is served from the same origin as the API (`https://converter.studio`), so these
 * are same-origin requests: no CORS and no preflight. They are sent without cookies
 * (`credentials: 'omit'`): the REST API is authorised by the API key alone, uploads and
 * downloads by the token in their link, and the site's own session cookie has nothing to do
 * here. The key is sent only to the API's own routes; upload and download links are sent
 * without it (the API reads an `Authorization` header before a link's token).
 */

import { VERSION } from './version';

/** Names this add-in, with its version, on its API calls (Converter Studio's connector counts). */
export const CLIENT_HEADER = { 'X-Converter-Studio-Client': `office/${VERSION}` } as const;

/** What the API says when something goes wrong: `{ error: { code, message, params } }`. */
export interface ApiErrorBody {
  readonly code: string;
  readonly message: string;
  readonly params?: Record<string, string | number>;
}

/** A failed request. `code` is the API's error code, or `network` when nothing came back. */
export class ApiError extends Error {
  constructor(
    readonly code: string,
    message: string,
    readonly status: number,
    readonly params: Record<string, string | number> = {},
  ) {
    super(message);
    this.name = 'ApiError';
  }
}

export interface FormatInfo {
  readonly id: string;
  readonly name: string;
  readonly category: string;
  readonly extensions: readonly string[];
  /** What it converts to on the server right now, most used first. */
  readonly to: readonly string[];
  /**
   * The servers make it smaller in its own format (MP4 → MP4, MP3 → MP3…, asked for with a
   * setting such as `emailFit`). Missing from servers before 2026-11-26.
   */
  readonly compress?: boolean;
}

export interface Account {
  readonly email: string;
  readonly plan: string;
  readonly credits: number;
  readonly maxFileBytes: number;
}

export interface Upload {
  readonly id: string;
  readonly uploadUrl: string;
  readonly maxBytes: number;
  readonly partBytes: number;
  readonly expiresAt: string;
}

export type ConversionStatus =
  | 'fetching'
  | 'waiting_for_upload'
  | 'measuring'
  | 'quoted'
  | 'queued'
  | 'processing'
  | 'done'
  | 'failed'
  | 'expired'
  | 'canceled';

export interface Conversion {
  readonly id: string;
  readonly status: ConversionStatus;
  readonly from: string | null;
  readonly to: string | null;
  readonly inputName: string | null;
  readonly credits: number | null;
  readonly progress: number | null;
  readonly outputName: string | null;
  readonly outputSize: number | null;
  /**
   * Done, and the result is the upload itself: it could not be made smaller in its own format,
   * so nothing was charged. Missing from servers before 2026-11-26.
   */
  readonly notSmaller?: boolean;
  readonly downloadUrl: string | null;
  readonly expiresAt: string | null;
  readonly error: ApiErrorBody | null;
}

export type ConversionOptions = Readonly<Record<string, string | number | boolean>>;

export interface ConvertRequest {
  readonly uploadId: string;
  readonly to: string;
  readonly options?: ConversionOptions;
  /** Goes ahead on its own up to this price; above it the conversion waits for `confirm`. */
  readonly maxCredits: number;
}

export type Fetch = (input: string, init?: RequestInit) => Promise<Response>;

export class ConverterApi {
  private readonly base: string;

  /** `base`: the API's root, e.g. `https://converter.studio/api`. */
  constructor(
    base: string,
    private readonly apiKey: string,
    private readonly fetchFn: Fetch = (input, init) => fetch(input, init),
  ) {
    this.base = base.replace(/\/$/, '');
  }

  /** The formats and what each becomes on the server now (no key needed). */
  formats(): Promise<FormatInfo[]> {
    return this.call<{ formats: FormatInfo[] }>('GET', '/v1/formats').then((r) => r.formats);
  }

  account(): Promise<Account> {
    return this.call('GET', '/v1/account');
  }

  createUpload(fileName: string, size: number, signal?: AbortSignal): Promise<Upload> {
    return this.call('POST', '/v1/uploads', { fileName, size }, signal);
  }

  /**
   * Sends a file to an upload slot: in one request when it is small, otherwise in parts
   * (`Upload-Offset`), which also gives progress and lets a broken part be sent again once.
   */
  async upload(
    upload: Upload,
    file: Blob,
    options: { partBytes?: number; onProgress?: (sent: number) => void; signal?: AbortSignal } = {},
  ): Promise<void> {
    const partBytes = Math.min(options.partBytes ?? 8 * 1024 * 1024, upload.partBytes);
    if (file.size <= partBytes) {
      await this.put(upload.uploadUrl, file, {}, options.signal);
      options.onProgress?.(file.size);
      return;
    }
    let offset = 0;
    let retried = false;
    while (offset < file.size) {
      const end = Math.min(offset + partBytes, file.size);
      try {
        await this.put(
          upload.uploadUrl,
          file.slice(offset, end),
          { 'Upload-Offset': String(offset) },
          options.signal,
        );
      } catch (error) {
        // A part the server did not get whole: go on from what it has, once.
        if (!retried && error instanceof ApiError && typeof error.params.received === 'number') {
          retried = true;
          offset = error.params.received;
          continue;
        }
        throw error;
      }
      offset = end;
      options.onProgress?.(offset);
    }
  }

  /**
   * Starts a conversion. `idempotencyKey`: the same key sent again (after a lost answer) gives
   * the same conversion, never a second one or a second charge.
   */
  convert(
    request: ConvertRequest,
    idempotencyKey: string,
    signal?: AbortSignal,
  ): Promise<Conversion> {
    return this.call('POST', '/v1/conversions', request, signal, {
      'Idempotency-Key': idempotencyKey,
    });
  }

  /** The conversion now, or (with `waitSeconds`) once it settles or that time has passed. */
  status(id: string, waitSeconds = 0, signal?: AbortSignal): Promise<Conversion> {
    const wait = waitSeconds > 0 ? `?wait=${Math.min(60, Math.round(waitSeconds))}` : '';
    return this.call('GET', `/v1/conversions/${encodeURIComponent(id)}${wait}`, undefined, signal);
  }

  confirm(id: string, signal?: AbortSignal): Promise<Conversion> {
    return this.call(
      'POST',
      `/v1/conversions/${encodeURIComponent(id)}/confirm`,
      undefined,
      signal,
    );
  }

  /** The result's file. Its link carries its own token (valid until the result expires). */
  async download(url: string, signal?: AbortSignal): Promise<Blob> {
    this.ownLink(url);
    const response = await this.send(url, { method: 'GET', signal });
    if (!response.ok) throw await errorOf(response);
    return response.blob();
  }

  /** Links the API gave out must point back at it: a key or a file never goes anywhere else. */
  ownLink(url: string): void {
    if (!url.startsWith(`${this.base}/`))
      throw new ApiError('invalidRequest', 'The link is not on Converter Studio.', 0);
  }

  private async put(
    url: string,
    body: Blob,
    headers: Record<string, string>,
    signal?: AbortSignal,
  ): Promise<void> {
    this.ownLink(url);
    const response = await this.send(url, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/octet-stream', ...headers },
      body,
      signal,
    });
    if (!response.ok) throw await errorOf(response);
  }

  private async call<T>(
    method: 'GET' | 'POST',
    path: string,
    body?: unknown,
    signal?: AbortSignal,
    extra: Record<string, string> = {},
  ): Promise<T> {
    const headers: Record<string, string> = { ...CLIENT_HEADER, ...extra };
    // The formats need no key (the pane lists them before one is entered).
    if (this.apiKey) headers.Authorization = `Bearer ${this.apiKey}`;
    if (body !== undefined) headers['Content-Type'] = 'application/json';
    const response = await this.send(`${this.base}${path}`, {
      method,
      headers,
      body: body === undefined ? undefined : JSON.stringify(body),
      signal,
    });
    if (!response.ok) throw await errorOf(response);
    return (await response.json()) as T;
  }

  private async send(url: string, init: RequestInit): Promise<Response> {
    try {
      return await this.fetchFn(url, { ...init, credentials: 'omit' });
    } catch (error) {
      if (isAbort(error)) throw error;
      throw new ApiError(
        'network',
        'Converter Studio could not be reached. Check your connection and try again.',
        0,
      );
    }
  }
}

export function isAbort(error: unknown): boolean {
  return (
    typeof error === 'object' &&
    error !== null &&
    (error as { name?: unknown }).name === 'AbortError'
  );
}

async function errorOf(response: Response): Promise<ApiError> {
  let body: { error?: ApiErrorBody | string } | null;
  try {
    body = (await response.json()) as { error?: ApiErrorBody | string };
  } catch {
    // Not JSON (a proxy's error page).
    body = null;
  }
  const error = body?.error;
  if (error && typeof error === 'object' && typeof error.message === 'string')
    return new ApiError(error.code, error.message, response.status, error.params ?? {});
  // The site's own routes (downloads) answer `{ error: 'code' }`.
  const code = typeof error === 'string' ? error : `http${response.status}`;
  return new ApiError(
    code,
    MESSAGES[code] ?? `Converter Studio answered ${response.status}. Try again later.`,
    response.status,
  );
}

const MESSAGES: Record<string, string> = {
  fileExpired: 'The converted file has expired. Convert it again.',
  notFound: 'Converter Studio could not find this conversion.',
  tooManyRequests: 'Too many requests. Wait a minute and try again.',
  serverBusy: 'Converter Studio is busy right now. Try again in a few minutes.',
};
