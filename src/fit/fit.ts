import { ApiError, isAbort, type Conversion, type ConverterApi, type FormatInfo } from '../api';
import type { Catalog } from '../catalog';
import { finish, follow, submit, type Hooks, type Step } from '../convert';
import { blobBytes } from '../office/async';
import { MAX_ATTACHMENT_BYTES, type DraftAttachment, type Mail } from '../office/mail';
import type { Messages } from '../strings';
import {
  MB,
  MIN_SAVING,
  QUALITY_HIGH,
  longSides,
  photoFormat,
  renamed,
  searchEncoding,
  serverSpec,
  shares,
  tooBig,
  type PhotoFormat,
  type ServerSpec,
} from './budget';
import { isAnimated, type PhotoEncoder } from './photo';

/**
 * "Fit attachments for sending": brings a draft's attachments under a total size and replaces
 * them in place, cheapest first.
 *
 * 1. Under the total already: nothing changes.
 * 2. Photos on the computer, largest first, until the draft fits: at most 2048 px, quality 0.85
 *    (free, no account).
 * 3. What still does not fit and the servers can make smaller (PDF, video, sound, a photo the
 *    computer cannot read), only with an API key: the API's price first, sent only on "Convert".
 *    Each gets a share of what is left, by size. Never when nothing could make the draft fit.
 * 4. Still over: the photos smaller still (quality, then 1600 and 1280 px), if that makes it fit.
 * 5. Each result replaces its original only once it is ready: added first, the original removed
 *    after, so a failure leaves the original. Names and order are kept (files after the first
 *    replaced one are added again in order); a name changes only with the type (HEIC → JPG).
 *
 * Inline pictures of the body and attached emails are never touched (they count towards the
 * total); OneDrive links do not travel with the message and are left out.
 */

export interface FitDeps {
  readonly mail: Mail;
  readonly catalog: Catalog;
  readonly photos: PhotoEncoder;
  /** The account's API client; null without a key (the computer's part still runs). */
  readonly api: ConverterApi | null;
  readonly t: Messages;
  readonly say: (text: string) => void;
  /** The servers' price for these files: true goes ahead and charges it. */
  readonly ask: (credits: number, names: readonly string[]) => Promise<boolean>;
  readonly signal?: AbortSignal;
  /** Tests make time pass faster. */
  readonly sleep?: Hooks['sleep'];
}

export interface FitChange {
  readonly name: string;
  readonly newName: string;
  readonly before: number;
  readonly after: number;
  readonly by: 'computer' | 'servers';
}

export interface FitReport {
  readonly kind: 'empty' | 'alreadyFits' | 'fits' | 'over';
  /** In MB, as chosen. */
  readonly target: number;
  /** The attachments that travel with the message. */
  readonly count: number;
  readonly before: number;
  readonly after: number;
  readonly changes: readonly FitChange[];
  /** Over: the largest attachments without which the rest would fit. */
  readonly tooBig: readonly { readonly name: string; readonly size: number }[];
  /** Without a key: what the servers could make smaller. */
  readonly needKey: readonly string[];
  /** The servers' price was declined for these. */
  readonly skipped: readonly string[];
  /**
   * The servers could not make these smaller (`notSmaller`): they stay as they were, and nothing
   * was charged for them. Not failures.
   */
  readonly kept: readonly string[];
  readonly failed: readonly { readonly name: string; readonly message: string }[];
  readonly credits: number;
}

interface Entry {
  readonly attachment: DraftAttachment;
  readonly format: FormatInfo | undefined;
  readonly photo: PhotoFormat | null;
  /** The bytes, read from Outlook once. */
  file?: Blob;
  /** The computer could not read it as a photo. */
  unreadable?: boolean;
  /** A moving PNG or WebP: left as it is (a canvas keeps only the first frame). */
  animated?: boolean;
  result?: { blob: Blob; name: string; by: FitChange['by'] };
}

const sizeOf = (entry: Entry) => entry.result?.blob.size ?? entry.attachment.size;
const sum = (entries: readonly Entry[]) => entries.reduce((total, e) => total + sizeOf(e), 0);

export async function fitDraft(deps: FitDeps, targetMb: number): Promise<FitReport> {
  const { mail, t, say } = deps;
  const target = targetMb * MB;
  say(t.fitChecking);
  const travelling = (await mail.everything()).filter((a) => a.kind !== 'cloud');
  const entries: Entry[] = travelling.map((attachment) => {
    const format = deps.catalog.formatOfName(attachment.name);
    return {
      attachment,
      format,
      photo: attachment.kind === 'file' ? photoFormat(format) : null,
    };
  });
  const before = sum(entries);
  const failed: { name: string; message: string }[] = [];
  const report = (kind: FitReport['kind'], rest: Partial<FitReport> = {}): FitReport => ({
    kind,
    target: targetMb,
    count: entries.length,
    before,
    after: before,
    changes: [],
    tooBig: [],
    needKey: [],
    skipped: [],
    kept: [],
    failed,
    credits: 0,
    ...rest,
  });
  if (entries.length === 0) return report('empty');
  if (before <= target) return report('alreadyFits');

  const total = () => sum(entries);
  const check = () => {
    if (deps.signal?.aborted) throw new ApiError('canceled', 'Canceled.', 0);
  };
  const read = async (entry: Entry): Promise<Blob | null> => {
    try {
      entry.file ??= await mail.content(entry.attachment);
      return entry.file;
    } catch (error) {
      failed.push({ name: entry.attachment.name, message: messageOf(error, t) });
      return null;
    }
  };

  // 2. Photos on the computer, largest first, until the draft fits.
  const photos = entries
    .filter((e) => e.photo)
    .sort((a, b) => b.attachment.size - a.attachment.size);
  for (const [i, entry] of photos.entries()) {
    if (total() <= target) break;
    check();
    say(t.fitPhotos(i + 1, photos.length));
    const file = await read(entry);
    if (!file) continue;
    if (isAnimated(await blobBytes(file.slice(0, 256 * 1024)), entry.photo!)) {
      entry.animated = true;
      continue;
    }
    const opened = await deps.photos.open(file, entry.photo!).catch(() => null);
    if (!opened) {
      entry.unreadable = true;
      continue;
    }
    try {
      const side = longSides(Math.max(opened.width, opened.height))[0]!;
      const blob = await opened.encode(side, opened.output === 'png' ? null : QUALITY_HIGH);
      if (blob.size <= entry.attachment.size * (1 - MIN_SAVING))
        entry.result = { blob, name: photoName(entry, opened.output), by: 'computer' };
    } finally {
      opened.close();
    }
  }

  // 3. The servers, for what the computer cannot make smaller.
  const needKey: string[] = [];
  const skipped: string[] = [];
  const kept: string[] = [];
  let credits = 0;
  const readable = photos.filter((e) => !e.unreadable && !e.animated && e.file);
  if (total() > target) {
    const candidates = entries.filter(
      (e) =>
        e.attachment.kind === 'file' &&
        (!e.photo || e.unreadable) &&
        serverSpec(e.format, e.attachment.size, e.attachment.size) !== null,
    );
    if (candidates.length > 0 && !deps.api) {
      needKey.push(...candidates.map((e) => e.attachment.name));
    } else if (candidates.length > 0 && deps.api) {
      const photoSize = sum(readable);
      const serverSize = sum(candidates);
      const fixed = total() - photoSize - serverSize;
      const room = target - fixed;
      // Nothing makes the draft fit when what cannot change is over already: no credits spent.
      if (room > 0) {
        const budget = Math.max(
          target - fixed - photoSize,
          Math.floor((room * serverSize) / (serverSize + photoSize)),
        );
        const parts = shares(
          candidates.map((e) => e.attachment.size),
          budget,
        );
        const jobs = candidates
          .map((entry, i) => ({
            entry,
            spec: serverSpec(entry.format, entry.attachment.size, parts[i]!),
          }))
          .filter((job): job is { entry: Entry; spec: ServerSpec } => job.spec !== null);
        const outcome = await onServers(deps, jobs, read, failed);
        credits = outcome.credits;
        skipped.push(...outcome.skipped);
        kept.push(...outcome.kept);
      }
    }
  }

  // 4. Still over: the photos smaller still, if that makes it fit.
  if (total() > target && readable.length > 0) {
    const others = total() - sum(readable);
    const room = target - others;
    if (room > 0) {
      const parts = shares(readable.map(sizeOf), room);
      const trial = new Map<Entry, { blob: Blob; output: 'jpg' | 'png' | 'webp' }>();
      for (const [i, entry] of readable.entries()) {
        if (sizeOf(entry) <= parts[i]!) continue;
        check();
        say(t.fitPhotos(i + 1, readable.length));
        const opened = await deps.photos.open(entry.file!, entry.photo!).catch(() => null);
        if (!opened) continue;
        try {
          const blob = await searchEncoding(
            (side, quality) => opened.encode(side, quality),
            parts[i]!,
            longSides(Math.max(opened.width, opened.height)),
            opened.output !== 'png',
          );
          if (blob.size < sizeOf(entry)) trial.set(entry, { blob, output: opened.output });
        } finally {
          opened.close();
        }
      }
      const squeezed =
        others + readable.reduce((s, e) => s + (trial.get(e)?.blob.size ?? sizeOf(e)), 0);
      // Only when it fits: otherwise the photos keep their first, better version.
      if (squeezed <= target)
        for (const [entry, { blob, output }] of trial)
          entry.result = { blob, name: photoName(entry, output), by: 'computer' };
    }
  }

  // 5. Into the draft.
  check();
  const changes = await replace(deps, entries, read, failed);
  const replaced = new Set(changes.map((c) => c.entry));
  const finalSize = (e: Entry) => (replaced.has(e) ? sizeOf(e) : e.attachment.size);
  const after = entries.reduce((s, e) => s + finalSize(e), 0);
  const inline = entries.filter((e) => e.attachment.kind === 'inline');
  return report(after <= target ? 'fits' : 'over', {
    after,
    changes: changes.map((c) => c.change),
    tooBig:
      after <= target
        ? []
        : tooBig(
            entries
              .filter((e) => e.attachment.kind !== 'inline')
              .map((e) => ({ name: e.attachment.name, size: finalSize(e) })),
            target - inline.reduce((s, e) => s + finalSize(e), 0),
          ),
    needKey,
    skipped,
    kept,
    credits,
  });
}

/** The photo's name: the same, or with the new type's extension (HEIC → JPG). */
function photoName(entry: Entry, output: 'jpg' | 'png' | 'webp'): string {
  const same = entry.photo === output;
  return renamed(entry.attachment.name, same ? null : output);
}

/**
 * Prices every file on the servers first (`maxCredits: 0`: each waits in `quoted`), asks once for
 * the total, then converts and fetches them. A file that fails keeps its original; so does one
 * the servers could not make smaller (`notSmaller`: the result is the upload itself, free), which
 * is not fetched.
 */
async function onServers(
  deps: FitDeps,
  jobs: readonly { entry: Entry; spec: ServerSpec }[],
  read: (entry: Entry) => Promise<Blob | null>,
  failed: { name: string; message: string }[],
): Promise<{ credits: number; skipped: string[]; kept: string[] }> {
  const { t, say, signal } = deps;
  const api = deps.api!;
  const quoted: { entry: Entry; spec: ServerSpec; conversion: Conversion }[] = [];
  const steps = (name: string) => (step: Step) => {
    if (step.step === 'uploading')
      say(t.fitSending(name, step.total > 0 ? Math.floor((step.sent / step.total) * 100) : 0));
    else
      say(
        step.status === 'quoted' || step.status === 'measuring'
          ? t.fitPricing(name)
          : t.fitConverting(name),
      );
  };
  for (const { entry, spec } of jobs) {
    const name = entry.attachment.name;
    const file = await read(entry);
    if (!file) continue;
    try {
      const hooks = { signal, sleep: deps.sleep, onStep: steps(name) };
      let conversion = await submit(
        api,
        { file, name, to: spec.to, options: spec.options },
        0,
        hooks,
      );
      conversion = await follow(api, conversion, { ...hooks, untilQuoted: true });
      if (conversion.status === 'quoted') quoted.push({ entry, spec, conversion });
      else finish(conversion, name);
    } catch (error) {
      if (aborted(error, signal)) throw error;
      failed.push({ name, message: messageOf(error, t) });
    }
  }
  if (quoted.length === 0) return { credits: 0, skipped: [], kept: [] };
  const price = quoted.reduce((s, q) => s + (q.conversion.credits ?? 0), 0);
  const names = quoted.map((q) => q.entry.attachment.name);
  if (price > 0 && !(await deps.ask(price, names))) {
    // Left waiting in `quoted`, they expire uncharged.
    if (signal?.aborted) throw new ApiError('canceled', 'Canceled.', 0);
    return { credits: 0, skipped: names, kept: [] };
  }
  let credits = 0;
  const kept: string[] = [];
  for (const { entry, spec, conversion } of quoted) {
    const name = entry.attachment.name;
    try {
      say(t.fitConverting(name));
      const hooks = { signal, sleep: deps.sleep, onStep: steps(name) };
      const result = finish(
        await follow(api, await api.confirm(conversion.id, signal), hooks),
        name,
      );
      credits += result.credits;
      if (result.conversion.notSmaller === true) {
        kept.push(name);
        continue;
      }
      const blob = await api.download(result.downloadUrl, signal);
      if (blob.size < entry.attachment.size)
        entry.result = { blob, name: renamed(name, spec.extension), by: 'servers' };
      else failed.push({ name, message: t.fitNotSmaller });
    } catch (error) {
      if (aborted(error, signal)) throw error;
      failed.push({ name, message: messageOf(error, t) });
    }
  }
  return { credits, skipped: [], kept };
}

/**
 * Puts the results into the draft: from the first replaced attachment on, every file is added
 * again in its order (the result, or the same file) and its original removed only after that; an
 * original whose replacement could not be added stays. Files Outlook would not take back from an
 * add-in (over 25 MB) stay where they are.
 */
async function replace(
  deps: FitDeps,
  entries: readonly Entry[],
  read: (entry: Entry) => Promise<Blob | null>,
  failed: { name: string; message: string }[],
): Promise<{ entry: Entry; change: FitChange }[]> {
  const { mail, t, say } = deps;
  const files = entries.filter((e) => e.attachment.kind === 'file');
  if (!files.some((e) => e.result)) return [];
  say(t.fitReplacing);
  const present = new Set((await mail.everything()).map((a) => a.id));
  const done: { entry: Entry; change: FitChange }[] = [];
  const first = files.findIndex((e) => e.result && present.has(e.attachment.id));
  for (const entry of files) {
    const { attachment, result } = entry;
    if (!present.has(attachment.id)) {
      if (result) failed.push({ name: attachment.name, message: t.fitRemoved });
      continue;
    }
    if (first === -1 || files.indexOf(entry) < first) continue;
    if (!result && attachment.size > MAX_ATTACHMENT_BYTES) continue;
    const blob = result?.blob ?? (await read(entry));
    if (!blob) continue;
    try {
      const id = await mail.attach(blob, result?.name ?? attachment.name);
      try {
        await mail.remove(attachment.id);
      } catch (error) {
        // Never two copies: the one just added goes again.
        await mail.remove(id).catch(() => undefined);
        throw error;
      }
      if (result)
        done.push({
          entry,
          change: {
            name: attachment.name,
            newName: result.name,
            before: attachment.size,
            after: result.blob.size,
            by: result.by,
          },
        });
    } catch (error) {
      if (result) failed.push({ name: attachment.name, message: messageOf(error, t) });
    }
  }
  return done;
}

function aborted(error: unknown, signal: AbortSignal | undefined): boolean {
  return (
    signal?.aborted === true ||
    isAbort(error) ||
    (error instanceof ApiError && error.code === 'canceled')
  );
}

function messageOf(error: unknown, t: Messages): string {
  return error instanceof Error && error.message ? error.message : t.failed;
}
