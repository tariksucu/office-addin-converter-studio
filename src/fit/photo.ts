import { blobBytes } from '../office/async';
import { longSides, scaled, type PhotoFormat } from './budget';
import { decodeHeic, type Libheif } from './heic';
import { ORIENTATION_PROBE, jpegOrientation, swapsSides, uprightTransform } from './orientation';

/**
 * Photos made smaller on the computer, in the pane itself: decoded by the webview (HEIC by
 * libheif where the webview cannot), turned upright from their EXIF orientation, drawn smaller on
 * a canvas and written again as JPEG (PNG for pictures with transparency, WebP stays WebP where
 * the webview writes it). Nothing leaves the computer. The canvas keeps no metadata: a re-encoded
 * photo loses its EXIF details, the location included.
 */

export type PhotoOutput = 'jpg' | 'png' | 'webp';

/** A photo ready to be written at several sizes. */
export interface OpenedPhoto {
  /** Upright size, in pixels. */
  readonly width: number;
  readonly height: number;
  /** Whether some of it is see-through. */
  readonly alpha: boolean;
  /** What it is written as. */
  readonly output: PhotoOutput;
  /** Writes it at most `longSide` pixels on its longer side, at `quality` (null: lossless). */
  encode(longSide: number, quality: number | null): Promise<Blob>;
  /** Frees its pixels. */
  close(): void;
}

export interface PhotoEncoder {
  /** The photo, or null when this computer cannot read it. */
  open(file: Blob, format: PhotoFormat): Promise<OpenedPhoto | null>;
}

const MIME: Readonly<Record<PhotoOutput, string>> = {
  jpg: 'image/jpeg',
  png: 'image/png',
  webp: 'image/webp',
};

/** The output for a photo: a JPEG stays a JPEG, transparency stays PNG (or WebP), the rest JPEG. */
export function outputFor(format: PhotoFormat, alpha: boolean, writesWebp: boolean): PhotoOutput {
  if (format === 'webp' && writesWebp) return 'webp';
  return alpha && format !== 'jpg' ? 'png' : 'jpg';
}

/**
 * Whether a PNG or WebP moves (APNG's `acTL` chunk before the picture data; WebP's `VP8X`
 * animation flag): a canvas would keep only its first frame, so it is left as it is.
 */
export function isAnimated(bytes: Uint8Array, format: PhotoFormat): boolean {
  const text = (at: number, length: number) =>
    String.fromCharCode(...bytes.subarray(at, at + length));
  if (format === 'webp')
    return (
      text(0, 4) === 'RIFF' &&
      text(8, 4) === 'WEBP' &&
      text(12, 4) === 'VP8X' &&
      ((bytes[20] ?? 0) & 0x02) !== 0
    );
  if (format !== 'png' || bytes.length < 8) return false;
  let at = 8;
  while (at + 8 <= bytes.length) {
    const length =
      ((bytes[at]! << 24) | (bytes[at + 1]! << 16) | (bytes[at + 2]! << 8) | bytes[at + 3]!) >>> 0;
    const type = text(at + 4, 4);
    if (type === 'acTL') return true;
    if (type === 'IDAT' || type === 'IEND') return false;
    at += 12 + length;
  }
  return false;
}

/** Whether any pixel of RGBA data is see-through. */
export function hasAlpha(data: Uint8ClampedArray): boolean {
  for (let i = 3; i < data.length; i += 4) if (data[i]! < 255) return true;
  return false;
}

type Drawable = HTMLImageElement | HTMLCanvasElement;

export function browserPhotos(doc: Document, heic: () => Promise<Libheif>): PhotoEncoder {
  const turnsPhotos = once(() => browserTurnsPhotos(doc));
  const writesWebp = once(() => Promise.resolve(canWrite(doc, 'image/webp')));

  return {
    async open(file, format) {
      const bytes = await blobBytes(file);
      let source: Drawable | null = null;
      let width = 0;
      let height = 0;
      let orientation = 1;
      const image = await loadImage(doc, file);
      if (image) {
        source = image;
        width = image.naturalWidth;
        height = image.naturalHeight;
        // A webview that does not turn photos itself draws them as stored: turn them here.
        if (format === 'jpg' && !(await turnsPhotos())) {
          orientation = jpegOrientation(bytes);
          if (swapsSides(orientation)) [width, height] = [height, width];
        }
      } else if (format === 'heic') {
        let decoded: Awaited<ReturnType<typeof decodeHeic>>;
        try {
          decoded = await decodeHeic(await heic(), bytes);
        } catch {
          return null;
        }
        const canvas = doc.createElement('canvas');
        canvas.width = decoded.width;
        canvas.height = decoded.height;
        const context = canvas.getContext('2d');
        if (!context) return null;
        const pixels = context.createImageData(decoded.width, decoded.height);
        pixels.data.set(decoded.data);
        context.putImageData(pixels, 0, 0);
        source = canvas;
        width = decoded.width;
        height = decoded.height;
      }
      if (!source || width === 0 || height === 0) return null;

      // The largest the photo is ever written at (2048 px or its own size), upright.
      const first = longSides(Math.max(width, height))[0]!;
      const base = draw(doc, source, width, height, first, orientation);
      if (source !== image) (source as HTMLCanvasElement).width = 0;
      const context = base.getContext('2d')!;
      const alpha =
        format !== 'jpg' && hasAlpha(context.getImageData(0, 0, base.width, base.height).data);
      const output = outputFor(format, alpha, await writesWebp());

      return {
        width,
        height,
        alpha,
        output,
        async encode(longSide, quality) {
          const size = scaled(base.width, base.height, longSide);
          // A JPEG of a see-through picture never happens (`outputFor`); white is only a guard.
          const canvas =
            size.width === base.width && size.height === base.height
              ? base
              : draw(doc, base, base.width, base.height, longSide, 1, output === 'jpg');
          try {
            return await canvasBlob(canvas, MIME[output], quality ?? undefined);
          } finally {
            if (canvas !== base) canvas.width = 0;
          }
        },
        close() {
          base.width = 0;
        },
      };
    },
  };
}

/**
 * `source` (stored `width` × `height` once upright) drawn upright at most `longSide` pixels on
 * its longer side; on white when it becomes a JPEG (which has no transparency).
 */
function draw(
  doc: Document,
  source: Drawable,
  width: number,
  height: number,
  longSide: number,
  orientation: number,
  white = false,
): HTMLCanvasElement {
  const size = scaled(width, height, longSide);
  const canvas = doc.createElement('canvas');
  canvas.width = size.width;
  canvas.height = size.height;
  const context = canvas.getContext('2d');
  if (!context) throw new Error('This computer cannot draw pictures here.');
  if (white) {
    context.fillStyle = '#ffffff';
    context.fillRect(0, 0, size.width, size.height);
  }
  context.imageSmoothingEnabled = true;
  context.imageSmoothingQuality = 'high';
  const { matrix, drawWidth, drawHeight } = uprightTransform(orientation, size.width, size.height);
  context.setTransform(...matrix);
  context.drawImage(source, 0, 0, drawWidth, drawHeight);
  context.setTransform(1, 0, 0, 1, 0, 0);
  return canvas;
}

/** The canvas as a file (`toDataURL` where an old webview has no `toBlob`). */
function canvasBlob(canvas: HTMLCanvasElement, type: string, quality?: number): Promise<Blob> {
  return new Promise((resolve, reject) => {
    if (typeof canvas.toBlob === 'function') {
      canvas.toBlob(
        (blob) => (blob ? resolve(blob) : reject(new Error('The photo could not be written.'))),
        type,
        quality,
      );
      return;
    }
    const url = canvas.toDataURL(type, quality);
    const binary = atob(url.slice(url.indexOf(',') + 1));
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    resolve(new Blob([bytes], { type }));
  });
}

/** The picture in an image element, or null when the webview cannot read it. */
function loadImage(doc: Document, file: Blob | string): Promise<HTMLImageElement | null> {
  const url = typeof file === 'string' ? file : URL.createObjectURL(file);
  return new Promise((resolve) => {
    const image = doc.createElement('img');
    const done = (ok: boolean) => {
      if (typeof file !== 'string') URL.revokeObjectURL(url);
      resolve(ok && image.naturalWidth > 0 ? image : null);
    };
    image.onload = () => done(true);
    image.onerror = () => done(false);
    image.src = url;
  });
}

/** Whether the webview draws a JPEG upright from its EXIF orientation (current ones do). */
async function browserTurnsPhotos(doc: Document): Promise<boolean> {
  const probe = await loadImage(doc, ORIENTATION_PROBE);
  return probe !== null && probe.naturalWidth === 1 && probe.naturalHeight === 2;
}

/** Whether the canvas writes this type (Safari writes no WebP and answers with a PNG). */
function canWrite(doc: Document, type: string): boolean {
  const canvas = doc.createElement('canvas');
  canvas.width = canvas.height = 1;
  try {
    return canvas.toDataURL(type).startsWith(`data:${type}`);
  } catch {
    return false;
  }
}

function once<T>(make: () => Promise<T>): () => Promise<T> {
  let value: Promise<T> | null = null;
  return () => (value ??= make());
}
