import { scaled, type PhotoFormat } from '../src/fit/budget';
import { blobBytes } from '../src/office/async';
import { outputFor, type PhotoEncoder } from '../src/fit/photo';

/**
 * Photos as the tests describe them, in place of the browser's canvas (e2e/ runs the real one):
 * a file starting with `PHOTO 4032x3024` (and ` alpha` when see-through) is such a photo; any
 * other file cannot be read. Writing one gives 0.3 byte a pixel at quality 1, 1.2 lossless (a
 * 12 MP photo at 2048 px and 0.85: about 800 KB).
 */
export function fakePhotos() {
  const opened: string[] = [];
  const encoder: PhotoEncoder = {
    async open(file: Blob, format: PhotoFormat) {
      const head = new TextDecoder().decode((await blobBytes(file)).subarray(0, 32));
      const match = /^PHOTO (\d+)x(\d+)( alpha)?/.exec(head);
      opened.push(head.slice(0, 20));
      if (!match) return null;
      const width = Number(match[1]);
      const height = Number(match[2]);
      const alpha = Boolean(match[3]);
      const output = outputFor(format, alpha, true);
      const type = { jpg: 'image/jpeg', png: 'image/png', webp: 'image/webp' }[output];
      return {
        width,
        height,
        alpha,
        output,
        encode(side: number, quality: number | null) {
          const size = scaled(width, height, side);
          const bytes = Math.round(size.width * size.height * (quality ?? 1.2) * 0.3);
          return Promise.resolve(new Blob([new Uint8Array(bytes)], { type }));
        },
        close() {},
      };
    },
  };
  return { encoder, opened };
}

/** A photo file of `bytes` bytes as `fakePhotos` reads it. */
export function photo(width: number, height: number, bytes: number, alpha = false): Uint8Array {
  const file = new Uint8Array(bytes);
  file.set(new TextEncoder().encode(`PHOTO ${width}x${height}${alpha ? ' alpha' : ''}`));
  return file;
}

/** Any other file of `bytes` bytes. */
export function filler(bytes: number, text = 'file'): Uint8Array {
  const file = new Uint8Array(bytes);
  file.set(new TextEncoder().encode(text));
  return file;
}
