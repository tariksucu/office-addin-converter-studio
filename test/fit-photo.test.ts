import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { decodeHeic, startLibheif, type LibheifFactory } from '../src/fit/heic';
import {
  ORIENTATION_PROBE,
  jpegOrientation,
  swapsSides,
  uprightTransform,
} from '../src/fit/orientation';
import { hasAlpha, isAnimated, outputFor } from '../src/fit/photo';

const samples = join(__dirname, '..', '..', '..', 'fixtures', 'samples');
const probe = Uint8Array.from(atob(ORIENTATION_PROBE.split(',')[1]!), (c) => c.charCodeAt(0));

/** A JPEG's start with an EXIF block saying `orientation`, little- or big-endian. */
function jpegWith(orientation: number, little: boolean, extra: number[] = []): Uint8Array {
  const u16 = (n: number) => (little ? [n & 255, n >> 8] : [n >> 8, n & 255]);
  const u32 = (n: number) =>
    little ? [n & 255, (n >> 8) & 255, 0, 0] : [0, 0, (n >> 8) & 255, n & 255];
  const tiff = [
    ...(little ? [0x49, 0x49] : [0x4d, 0x4d]),
    ...u16(42),
    ...u32(8),
    ...u16(2),
    // An unrelated entry first (ImageWidth), then Orientation.
    ...u16(0x0100),
    ...u16(3),
    ...u32(1),
    ...u16(640),
    0,
    0,
    ...u16(0x0112),
    ...u16(3),
    ...u32(1),
    ...u16(orientation),
    0,
    0,
    0,
    0,
    0,
    0,
  ];
  const body = [0x45, 0x78, 0x69, 0x66, 0, 0, ...tiff];
  const length = body.length + 2;
  return Uint8Array.from([
    0xff,
    0xd8,
    ...extra,
    0xff,
    0xe1,
    length >> 8,
    length & 255,
    ...body,
    0xff,
    0xda,
    0,
    2,
  ]);
}

describe('JPEG orientation', () => {
  it('reads it from the EXIF block, either byte order', () => {
    for (let o = 1; o <= 8; o++) {
      expect(jpegOrientation(jpegWith(o, true))).toBe(o);
      expect(jpegOrientation(jpegWith(o, false))).toBe(o);
    }
    expect(jpegOrientation(probe)).toBe(6);
  });

  it('looks past other blocks and fill bytes before the EXIF one', () => {
    const app0 = [0xff, 0xe0, 0, 6, 0x4a, 0x46, 0x49, 0x46];
    expect(jpegOrientation(jpegWith(8, false, [...app0, 0xff]))).toBe(8);
  });

  it('is 1 without EXIF, for what is not a JPEG, and for a cut or nonsense block', () => {
    expect(jpegOrientation(new Uint8Array(readFileSync(join(samples, 'sample.jpg'))))).toBe(1);
    expect(jpegOrientation(new Uint8Array(readFileSync(join(samples, 'sample.png'))))).toBe(1);
    expect(jpegOrientation(jpegWith(6, true).subarray(0, 30))).toBe(1);
    expect(jpegOrientation(jpegWith(9, true))).toBe(1);
    expect(jpegOrientation(new Uint8Array())).toBe(1);
  });

  it('draws each orientation upright: the stored corners land where they belong', () => {
    // A stored 4 × 2 picture; upright it is 4 × 2 (1–4) or 2 × 4 (5–8).
    const stored = { w: 4, h: 2 };
    // Where the stored top-left pixel corner (0, 0) ends up, per EXIF's definitions.
    const topLeftGoesTo: Record<number, [number, number]> = {
      1: [0, 0],
      2: [4, 0],
      3: [4, 2],
      4: [0, 2],
      5: [0, 0],
      6: [2, 0],
      7: [2, 4],
      8: [0, 4],
    };
    for (let o = 1; o <= 8; o++) {
      const [width, height] = swapsSides(o) ? [stored.h, stored.w] : [stored.w, stored.h];
      const { matrix, drawWidth, drawHeight } = uprightTransform(o, width, height);
      expect([drawWidth, drawHeight]).toEqual([stored.w, stored.h]);
      const [a, b, c, d, e, f] = matrix;
      const map = (x: number, y: number) => [a * x + c * y + e, b * x + d * y + f];
      expect(map(0, 0)).toEqual(topLeftGoesTo[o]);
      // Every stored corner stays inside the upright canvas.
      for (const [x, y] of [
        [0, 0],
        [stored.w, 0],
        [0, stored.h],
        [stored.w, stored.h],
      ] as const) {
        const [u, v] = map(x, y);
        expect(u! >= 0 && u! <= width && v! >= 0 && v! <= height).toBe(true);
      }
    }
  });
});

describe('what a photo is written as', () => {
  it('a JPEG stays JPEG, an opaque PNG becomes JPEG, see-through stays PNG, WebP where written', () => {
    expect(outputFor('jpg', false, true)).toBe('jpg');
    expect(outputFor('png', false, true)).toBe('jpg');
    expect(outputFor('png', true, true)).toBe('png');
    expect(outputFor('webp', false, true)).toBe('webp');
    expect(outputFor('webp', true, false)).toBe('png');
    expect(outputFor('webp', false, false)).toBe('jpg');
    expect(outputFor('heic', false, true)).toBe('jpg');
  });

  it('knows a moving PNG or WebP (left as it is: a canvas keeps only the first frame)', () => {
    const png = new Uint8Array(readFileSync(join(samples, 'sample.png')));
    expect(isAnimated(png, 'png')).toBe(false);
    // An APNG: the acTL chunk right after IHDR.
    const ihdrEnd = 8 + 12 + 13;
    const actl = [0, 0, 0, 8, 0x61, 0x63, 0x54, 0x4c, 0, 0, 0, 2, 0, 0, 0, 0, 1, 2, 3, 4];
    const apng = Uint8Array.from([...png.subarray(0, ihdrEnd), ...actl, ...png.subarray(ihdrEnd)]);
    expect(isAnimated(apng, 'png')).toBe(true);
    const webp = new Uint8Array(readFileSync(join(samples, 'sample.webp')));
    expect(isAnimated(webp, 'webp')).toBe(false);
    const header = (flags: number) =>
      Uint8Array.from([
        ...new TextEncoder().encode('RIFF'),
        0,
        0,
        0,
        0,
        ...new TextEncoder().encode('WEBPVP8X'),
        10,
        0,
        0,
        0,
        flags,
        0,
        0,
        0,
      ]);
    expect(isAnimated(header(0x02), 'webp')).toBe(true);
    expect(isAnimated(header(0x10), 'webp')).toBe(false);
    expect(isAnimated(header(0x02), 'jpg')).toBe(false);
  });

  it('finds a see-through pixel', () => {
    expect(hasAlpha(new Uint8ClampedArray([1, 2, 3, 255, 4, 5, 6, 255]))).toBe(false);
    expect(hasAlpha(new Uint8ClampedArray([1, 2, 3, 255, 4, 5, 6, 254]))).toBe(true);
  });
});

describe('HEIC with libheif-js (its WebAssembly build, as the pane loads it)', () => {
  const folder = join(__dirname, '..', 'node_modules', 'libheif-js', 'libheif-wasm') + '/';
  const factory = createRequire(import.meta.url)(`${folder}libheif.js`) as LibheifFactory;
  // As the pane does: the bytes handed over (this build compiles them synchronously).
  const wasm = new Uint8Array(readFileSync(`${folder}libheif.wasm`));

  it('decodes the sample photo upright into RGBA', async () => {
    const lib = await startLibheif(factory, folder, wasm);
    const photo = await decodeHeic(lib, new Uint8Array(readFileSync(join(samples, 'sample.heic'))));
    expect(photo.width).toBeGreaterThan(0);
    expect(photo.height).toBeGreaterThan(0);
    expect(photo.data.length).toBe(photo.width * photo.height * 4);
    expect(hasAlpha(photo.data)).toBe(false);
  });

  it('refuses what is not a HEIC photo', async () => {
    const lib = await startLibheif(factory, folder, wasm);
    await expect(decodeHeic(lib, new Uint8Array(4096).fill(7))).rejects.toThrow(/HEIC/);
  });
});
