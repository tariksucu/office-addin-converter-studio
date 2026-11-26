import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import type { FormatInfo } from '../src/api';
import {
  EMAIL_ATTACHMENT_MB,
  FIT_TARGETS,
  MB,
  QUALITY_HIGH,
  QUALITY_LOW,
  isFitTarget,
  longSides,
  photoFormat,
  renamed,
  scaled,
  searchEncoding,
  serverSpec,
  shares,
  tooBig,
} from '../src/fit/budget';
import { FORMATS } from './fake-api';

const format = (id: string) => FORMATS.find((f) => f.id === id)!;

describe('the target', () => {
  it('is the site’s “Fit for email” size, and 10 MB', () => {
    expect(FIT_TARGETS).toEqual([16, 10]);
    expect(isFitTarget(16) && isFitTarget(10) && !isFitTarget(20) && !isFitTarget('x')).toBe(true);
    // In the monorepo, the same number as packages/shared (absent in the exported repository).
    const shared = join(__dirname, '..', '..', '..', 'packages', 'shared', 'src', 'options.ts');
    if (existsSync(shared))
      expect(readFileSync(shared, 'utf8')).toContain(
        `export const EMAIL_ATTACHMENT_MB = ${EMAIL_ATTACHMENT_MB};`,
      );
  });
});

describe('photos on the computer', () => {
  it('knows JPEG, PNG, WebP and HEIC/HEIF by the API’s formats', () => {
    expect(
      ['jpg', 'png', 'webp', 'heic', 'tiff', 'gif', 'pdf'].map((id) => photoFormat(format(id))),
    ).toEqual(['jpg', 'png', 'webp', 'heic', null, null, null]);
    expect(photoFormat(undefined)).toBeNull();
  });

  it('tries 2048 px, then 1600 and 1280, never larger than the photo', () => {
    expect(longSides(4032)).toEqual([2048, 1600, 1280]);
    expect(longSides(1700)).toEqual([1700, 1600, 1280]);
    expect(longSides(1600)).toEqual([1600, 1280]);
    expect(longSides(900)).toEqual([900]);
  });

  it('scales by the longer side, portrait or landscape, without upscaling', () => {
    expect(scaled(4032, 3024, 2048)).toEqual({ width: 2048, height: 1536 });
    expect(scaled(3024, 4032, 2048)).toEqual({ width: 1536, height: 2048 });
    expect(scaled(800, 600, 2048)).toEqual({ width: 800, height: 600 });
  });
});

/**
 * A photo whose JPEG takes 0.3 byte a pixel at quality 1 (a 12 MP phone photo's 2048 px copy:
 * about 800 KB at 0.85); what it "writes" says how it was made.
 */
function fakeJpeg(width: number, height: number) {
  const calls: [number, number | null][] = [];
  const encodeAt = (side: number, quality: number | null) => {
    calls.push([side, quality]);
    const size = scaled(width, height, side);
    const bytes = Math.round(size.width * size.height * (quality ?? 1.2) * 0.3);
    return Promise.resolve({ size: bytes, side, quality } as unknown as Blob);
  };
  return { encodeAt, calls };
}

describe('searchEncoding', () => {
  it('keeps the first try when it fits', async () => {
    const photo = fakeJpeg(4032, 3024);
    const blob = await searchEncoding(photo.encodeAt, 2 * MB, longSides(4032), true);
    expect(photo.calls).toEqual([[2048, QUALITY_HIGH]]);
    expect(blob.size).toBeLessThan(2 * MB);
  });

  it('finds the highest quality under the budget at 2048 px before going smaller', async () => {
    const photo = fakeJpeg(4032, 3024);
    const budget = 600_000;
    const blob = (await searchEncoding(photo.encodeAt, budget, longSides(4032), true)) as Blob & {
      side: number;
      quality: number;
    };
    expect(blob.size).toBeLessThanOrEqual(budget);
    expect(blob.side).toBe(2048);
    expect(blob.quality).toBeGreaterThan(QUALITY_LOW);
    expect(blob.quality).toBeLessThan(QUALITY_HIGH);
    // 0.85, 0.5, then four halvings.
    expect(photo.calls.length).toBe(6);
    // Within one step of the best quality that fits (0.6375 here).
    expect(blob.quality).toBeGreaterThan(0.6);
  });

  it('steps down to 1600 and 1280 px only while the lowest quality does not fit', async () => {
    const photo = fakeJpeg(4032, 3024);
    const blob = (await searchEncoding(photo.encodeAt, 250_000, longSides(4032), true)) as Blob & {
      side: number;
    };
    expect(blob.side).toBe(1280);
    expect(blob.size).toBeLessThanOrEqual(250_000);
    expect(photo.calls.filter(([side]) => side === 2048)).toEqual([
      [2048, QUALITY_HIGH],
      [2048, QUALITY_LOW],
    ]);
  });

  it('gives the smallest it made when nothing fits', async () => {
    const photo = fakeJpeg(4032, 3024);
    const blob = (await searchEncoding(photo.encodeAt, 1000, longSides(4032), true)) as Blob & {
      side: number;
      quality: number;
    };
    expect([blob.side, blob.quality]).toEqual([1280, QUALITY_LOW]);
  });

  it('only resizes a lossless (see-through PNG) picture', async () => {
    const photo = fakeJpeg(3000, 2000);
    await searchEncoding(photo.encodeAt, 900_000, longSides(3000), false);
    expect(photo.calls).toEqual([
      [2048, null],
      [1600, null],
    ]);
  });
});

describe('shares', () => {
  it('splits a budget by size', () => {
    expect(shares([8 * MB, 2 * MB], 5 * MB)).toEqual([4 * MB, 1 * MB]);
  });

  it('never gives a file more than its size, and the rest goes to the others', () => {
    expect(shares([1 * MB, 9 * MB], 20 * MB)).toEqual([1 * MB, 9 * MB]);
    const parts = shares([3, 5, 1000], 500);
    expect(parts.reduce((a, b) => a + b, 0)).toBeLessThanOrEqual(500);
    expect(parts.every((part, i) => part <= [3, 5, 1000][i]!)).toBe(true);
  });

  it('gives nothing when there is no room', () => {
    expect(shares([4, 5], 0)).toEqual([0, 0]);
    expect(shares([4, 5], -10)).toEqual([0, 0]);
    expect(shares([], 10)).toEqual([]);
  });
});

describe('serverSpec', () => {
  const options = (id: string, size: number, share: number) => serverSpec(format(id), size, share);

  it('PDF: fit for email, from strong compression when more than half must go', () => {
    expect(options('pdf', 12 * MB, 9 * MB)).toEqual({
      to: 'pdf',
      options: { emailFit: true, compression: 'balanced' },
      extension: null,
    });
    expect(options('pdf', 30 * MB, 6 * MB)?.options).toEqual({
      emailFit: true,
      compression: 'strong',
    });
  });

  it('video: an MP4 at most its share (whole MB, never above 16)', () => {
    expect(options('mov', 60 * MB, 9.7 * MB)).toEqual({
      to: 'mp4',
      options: { emailFit: true, maxSize: 9 },
      extension: 'mp4',
    });
    expect(options('mov', 60 * MB, 40 * MB)?.options).toEqual({ emailFit: true, maxSize: 16 });
    expect(options('mov', 60 * MB, 0.3 * MB)?.options).toEqual({ emailFit: true, maxSize: 1 });
  });

  it('an MP4 stays an MP4 with its name: made smaller in its own format', () => {
    // Not a conversion the API lists in `to`: its `compress` says the servers do it.
    expect(format('mp4').to).not.toContain('mp4');
    expect(format('mp4').compress).toBe(true);
    expect(options('mp4', 30 * MB, 9 * MB)).toEqual({
      to: 'mp4',
      options: { emailFit: true, maxSize: 9 },
      extension: null,
    });
  });

  it('sound: MP3, fit for email; an MP3 stays an MP3 with its name', () => {
    expect(options('wav', 40 * MB, 9 * MB)).toEqual({
      to: 'mp3',
      options: { emailFit: true },
      extension: 'mp3',
    });
    expect(options('mp3', 30 * MB, 9 * MB)).toEqual({
      to: 'mp3',
      options: { emailFit: true },
      extension: null,
    });
  });

  it('servers without the same-format compression: an MP4 becomes a MOV, an MP3 an M4A', () => {
    // As servers before 2026-11-26 answered: no `compress`.
    const before = (id: string): FormatInfo => ({ ...format(id), compress: undefined });
    expect(serverSpec(before('mp4'), 30 * MB, 9 * MB)).toEqual({
      to: 'mov',
      options: { emailFit: true },
      extension: 'mov',
    });
    expect(serverSpec(before('mp3'), 30 * MB, 9 * MB)).toEqual({
      to: 'm4a',
      options: { emailFit: true },
      extension: 'm4a',
    });
    expect(serverSpec(before('mov'), 30 * MB, 9 * MB)?.to).toBe('mp4');
    // Neither its own format nor the old target: nothing.
    expect(serverSpec({ ...before('mp4'), to: ['mp3', 'gif'] }, 30 * MB, 9 * MB)).toBeNull();
    // Its own format only where `compress` says so (a MOV is not, today).
    expect(serverSpec({ ...format('mov'), to: ['mp3'] }, 30 * MB, 9 * MB)).toBeNull();
    expect(serverSpec({ ...format('mov'), to: ['mp3'], compress: true }, 30 * MB, 9 * MB)).toEqual({
      to: 'mov',
      options: { emailFit: true },
      extension: null,
    });
  });

  it('a picture the computer cannot read: a JPG at most its share, 1920 px', () => {
    expect(options('heic', 3 * MB, 1.5 * MB)).toEqual({
      to: 'jpg',
      options: { maxKb: 1500, longSide: '1920' },
      extension: 'jpg',
    });
    expect(options('tiff', 20 * MB, 200_000_000)?.options).toEqual({
      maxKb: 100_000,
      longSide: '1920',
    });
    expect(options('dng', 25 * MB, 2000)?.options).toEqual({ maxKb: 5, longSide: '1920' });
  });

  it('nothing for what the servers would not make smaller, or what is too small to matter', () => {
    for (const id of ['docx', 'zip', 'txt', 'gif', 'jpg'])
      expect(options(id, 20 * MB, 5 * MB)).toBeNull();
    expect(options('pdf', 150_000, 100)).toBeNull();
    expect(serverSpec(undefined, 20 * MB, MB)).toBeNull();
  });

  it('asks only for targets the API lists for the format', () => {
    const noMp4: FormatInfo = { ...format('mov'), to: ['mp3', 'gif'] };
    expect(serverSpec(noMp4, 50 * MB, 9 * MB)).toBeNull();
    const pdfWithoutCompression: FormatInfo = { ...format('pdf'), to: ['jpg', 'docx'] };
    expect(serverSpec(pdfWithoutCompression, 50 * MB, 9 * MB)).toBeNull();
  });
});

describe('names and what is still too big', () => {
  it('keeps the name unless the type changes', () => {
    expect(renamed('IMG_0001.HEIC', 'jpg')).toBe('IMG_0001.jpg');
    expect(renamed('Holiday clip.mov', 'mp4')).toBe('Holiday clip.mp4');
    expect(renamed('Report v2.pdf', null)).toBe('Report v2.pdf');
    expect(renamed('.heic', 'jpg')).toBe('attachment.jpg');
  });

  it('names the fewest largest attachments without which the rest fits', () => {
    const items = [
      { name: 'a.pdf', size: 3 * MB },
      { name: 'clip.mov', size: 40 * MB },
      { name: 'b.docx', size: 9 * MB },
    ];
    expect(tooBig(items, 16 * MB).map((x) => x.name)).toEqual(['clip.mov']);
    expect(tooBig(items, 10 * MB).map((x) => x.name)).toEqual(['clip.mov', 'b.docx']);
    expect(tooBig(items, 60 * MB)).toEqual([]);
  });
});
