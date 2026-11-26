import type { ConversionOptions, FormatInfo } from '../api';

/**
 * "Fit attachments for sending": the numbers and rules that decide what happens to each
 * attachment of a draft, with no Office and no browser in them (so the tests can check them all).
 */

/** A megabyte as the site counts sizes and limits: the decimal one, so a limit is always met. */
export const MB = 1_000_000;

/**
 * The site's "Fit for email" size (`EMAIL_ATTACHMENT_MB` in packages/shared/src/options.ts;
 * docs/PLAN.md "E-postaya sığsın"): 16 MB of attachments become about 22 MB once a mail program
 * encodes them (Base64), which Gmail and Yahoo Mail take (25 MB). 10 MB (about 14 MB encoded) also
 * gets through Outlook.com and iCloud Mail (20 MB).
 */
export const EMAIL_ATTACHMENT_MB = 16;
export const FIT_TARGETS = [EMAIL_ATTACHMENT_MB, 10] as const;
export type FitTarget = (typeof FIT_TARGETS)[number];

export function isFitTarget(value: unknown): value is FitTarget {
  return FIT_TARGETS.includes(Number(value) as FitTarget);
}

/** Photos made smaller on the computer: their longer side at most this, then these steps. */
export const PHOTO_LONG_SIDES = [2048, 1600, 1280] as const;
/** JPEG (and WebP) quality: the first try, and the lowest the search goes. */
export const QUALITY_HIGH = 0.85;
export const QUALITY_LOW = 0.5;
/** A re-encoded photo is kept only when it saves at least this share of the original's size. */
export const MIN_SAVING = 0.1;
/** Files smaller than this are not worth a credit on the servers. */
export const MIN_SERVER_BYTES = 200 * 1000;

/** What the computer itself makes smaller: JPEG, PNG, WebP, and HEIC/HEIF where it can read it. */
export type PhotoFormat = 'jpg' | 'png' | 'webp' | 'heic';
const PHOTOS: ReadonlySet<string> = new Set(['jpg', 'png', 'webp', 'heic']);

export function photoFormat(format: FormatInfo | undefined): PhotoFormat | null {
  return format && PHOTOS.has(format.id) ? (format.id as PhotoFormat) : null;
}

/**
 * The longer sides a photo is tried at: at most 2048 px, then 1600 and 1280, never larger than
 * the photo itself.
 */
export function longSides(longSide: number): number[] {
  const first = Math.min(PHOTO_LONG_SIDES[0], Math.max(1, Math.round(longSide)));
  return [first, ...PHOTO_LONG_SIDES.slice(1).filter((side) => side < first)];
}

/** A size at most `longSide` on its longer side, proportions kept, never upscaled. */
export function scaled(
  width: number,
  height: number,
  longSide: number,
): { width: number; height: number } {
  const scale = Math.min(1, longSide / Math.max(width, height));
  return {
    width: Math.max(1, Math.round(width * scale)),
    height: Math.max(1, Math.round(height * scale)),
  };
}

/** One way of writing a photo: at a longer side, at a quality (null: lossless, PNG). */
export type EncodeAt = (longSide: number, quality: number | null) => Promise<Blob>;

/**
 * The best photo under `budget` bytes: at each longer side (2048, then 1600, then 1280) the
 * highest quality that fits, found by halving between `QUALITY_HIGH` and `QUALITY_LOW`; a
 * lossless photo (a PNG with transparency) only steps down in size. When nothing fits, the
 * smallest made.
 */
export async function searchEncoding(
  encodeAt: EncodeAt,
  budget: number,
  sides: readonly number[],
  lossy: boolean,
): Promise<Blob> {
  let smallest: Blob | null = null;
  const tryAt = async (side: number, quality: number | null) => {
    const blob = await encodeAt(side, quality);
    if (!smallest || blob.size < smallest.size) smallest = blob;
    return blob;
  };
  for (const side of sides) {
    if (!lossy) {
      const blob = await tryAt(side, null);
      if (blob.size <= budget) return blob;
      continue;
    }
    const high = await tryAt(side, QUALITY_HIGH);
    if (high.size <= budget) return high;
    const low = await tryAt(side, QUALITY_LOW);
    if (low.size > budget) continue;
    let best = low;
    let [from, to] = [QUALITY_LOW, QUALITY_HIGH];
    for (let i = 0; i < 4; i++) {
      const middle = (from + to) / 2;
      const blob = await tryAt(side, middle);
      if (blob.size <= budget) [best, from] = [blob, middle];
      else to = middle;
    }
    return best;
  }
  return smallest!;
}

/**
 * Shares of `budget` for files of `sizes`: proportional to each file's size, and never more than
 * its size (what one does not need goes to the others). Whole bytes.
 */
export function shares(sizes: readonly number[], budget: number): number[] {
  const result = sizes.map(() => 0);
  let active = sizes.map((_, i) => i).filter((i) => sizes[i]! > 0);
  let left = Math.max(0, budget);
  while (active.length > 0) {
    const weight = active.reduce((sum, i) => sum + sizes[i]!, 0);
    const whole = active.filter((i) => (sizes[i]! * left) / weight >= sizes[i]!);
    if (whole.length === 0) {
      for (const i of active) result[i] = Math.floor((sizes[i]! * left) / weight);
      break;
    }
    for (const i of whole) {
      result[i] = sizes[i]!;
      left -= sizes[i]!;
    }
    active = active.filter((i) => !whole.includes(i));
  }
  return result;
}

/** A conversion on Converter Studio's servers that makes a file smaller. */
export interface ServerSpec {
  readonly to: string;
  readonly options: ConversionOptions;
  /** The extension of the result (the original's name is kept when the type stays). */
  readonly extension: string | null;
}

/** Image formats not worth sending (moving pictures, icons, drawings). */
const KEEP_IMAGES: ReadonlySet<string> = new Set(['gif', 'ico', 'svg']);

/**
 * How the servers make a file of `format` (`size` bytes) at most about `share` bytes, with only
 * targets the API lists for the format now and the settings packages/shared/src/options.ts takes
 * for them; null when nothing would.
 *
 * - PDF → PDF: "fit for email" (`emailFit`), from the strong compression when it must lose more
 *   than half.
 * - Video → MP4: "fit for email", at most `share` (`maxSize`, whole MB). An MP4 stays an MP4
 *   with its name: the servers make it smaller in its own format (`compress` in the format
 *   list, from 2026-11-26); only from servers that do not list that, an MP4 → MOV as before.
 * - Sound → MP3: "fit for email". An MP3 stays an MP3 the same way (→ M4A from older servers).
 * - A picture the computer cannot read (camera RAW, TIFF, BMP, AVIF, a HEIC libheif cannot
 *   decode) → JPG at most `share` (`maxKb`), at most 1920 px on its longer side.
 */
export function serverSpec(
  format: FormatInfo | undefined,
  size: number,
  share: number,
): ServerSpec | null {
  if (!format || size < MIN_SERVER_BYTES) return null;
  const offers = (to: string) => format.to.includes(to);
  const first = (...targets: string[]) =>
    targets.find((to) => to !== format.id && offers(to)) ?? null;
  /** Its own format where the servers make it smaller there, else another one they list. */
  const ownOrFirst = (...targets: string[]) =>
    targets.find((to) => (to === format.id ? format.compress === true : offers(to))) ?? null;
  /** The name stays when the type does. */
  const extension = (to: string) => (to === format.id ? null : to);
  if (format.id === 'pdf' && offers('pdf'))
    return {
      to: 'pdf',
      options: { emailFit: true, compression: share < size / 2 ? 'strong' : 'balanced' },
      extension: null,
    };
  if (format.category === 'video') {
    const to = ownOrFirst('mp4', 'mov');
    if (!to) return null;
    const megabytes = Math.min(EMAIL_ATTACHMENT_MB, Math.max(1, Math.floor(share / MB)));
    return {
      to,
      options: to === 'mp4' ? { emailFit: true, maxSize: megabytes } : { emailFit: true },
      extension: extension(to),
    };
  }
  if (format.category === 'audio') {
    const to = ownOrFirst('mp3', 'm4a');
    return to ? { to, options: { emailFit: true }, extension: extension(to) } : null;
  }
  if (format.category === 'image' && !KEEP_IMAGES.has(format.id)) {
    const to = first('jpg');
    if (!to) return null;
    const kb = Math.min(100_000, Math.max(5, Math.floor(share / 1000)));
    return { to, options: { maxKb: kb, longSide: '1920' }, extension: 'jpg' };
  }
  return null;
}

/** The original's name with another extension (`null`: the name as it is). */
export function renamed(name: string, extension: string | null): string {
  if (!extension) return name;
  const base = name.trim().replace(/\.[A-Za-z0-9]{1,5}$/, '') || 'attachment';
  return `${base}.${extension}`;
}

/**
 * What is still too big when a draft does not fit: the largest attachments, as few as possible,
 * without which the rest fits (those to share as links instead).
 */
export function tooBig<T extends { readonly size: number }>(
  items: readonly T[],
  target: number,
): T[] {
  const sorted = [...items].sort((a, b) => b.size - a.size);
  let total = items.reduce((sum, item) => sum + item.size, 0);
  const out: T[] = [];
  for (const item of sorted) {
    if (total <= target) break;
    out.push(item);
    total -= item.size;
  }
  return out;
}
