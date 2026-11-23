import { describe, expect, it } from 'vitest';
import { ApiError, ConverterApi } from '../src/api';
import {
  AUTO_CONFIRM_CREDITS,
  TIMEOUT_MS,
  convert,
  fetchResult,
  isSingleFile,
  newIdempotencyKey,
} from '../src/convert';
import { BASE, KEY, fakeApi } from './fake-api';

const job = {
  file: new Blob(['heic bytes']),
  name: 'IMG_1.heic',
  to: 'jpg',
  options: { quality: 'high' },
};
const noSleep = () => Promise.resolve();

describe('convert', () => {
  it('uploads, converts, waits and returns the download link', async () => {
    const server = fakeApi({ polls: 2 });
    const api = new ConverterApi(BASE, KEY, server.fetch);
    const steps: string[] = [];
    const result = await convert(api, job, {
      confirmPrice: () => Promise.reject(new Error('not asked')),
      onStep: (s) => steps.push(s.step),
      sleep: noSleep,
    });
    expect(result.outputName).toBe('IMG_1.jpg');
    expect(result.credits).toBe(1);
    expect(result.downloadUrl).toMatch(/\/jobs\/up1\/download\?token=t-up1$/);
    expect(steps[0]).toBe('uploading');
    expect(steps).toContain('converting');
    const post = server.requests.find((r) => r.url.endsWith('/v1/conversions'))!;
    expect(post.body).toMatchObject({
      to: 'jpg',
      maxCredits: AUTO_CONFIRM_CREDITS,
      options: { quality: 'high' },
    });
    expect(post.headers['Idempotency-Key']).toMatch(/^office-/);
    const statuses = server.requests.filter((r) => /conversions\/up1\?wait=20$/.test(r.url));
    expect(statuses.length).toBe(2);
    expect(new TextDecoder().decode(await (await fetchResult(api, result)).arrayBuffer())).toBe(
      'converted',
    );
  });

  it('asks before a price above the limit, and confirms when told to', async () => {
    const server = fakeApi({ credits: 7 });
    const api = new ConverterApi(BASE, KEY, server.fetch);
    const asked: number[] = [];
    const result = await convert(api, job, {
      confirmPrice: (credits) => {
        asked.push(credits);
        return Promise.resolve(true);
      },
      sleep: noSleep,
    });
    expect(asked).toEqual([7]);
    expect(server.requests.some((r) => r.url.endsWith('/confirm'))).toBe(true);
    expect(result.credits).toBe(7);
  });

  it('stops without paying when the price is declined', async () => {
    const server = fakeApi({ credits: 7 });
    const api = new ConverterApi(BASE, KEY, server.fetch);
    await expect(
      convert(api, job, { confirmPrice: () => Promise.resolve(false), sleep: noSleep }),
    ).rejects.toMatchObject({ code: 'declined', message: 'Canceled. Nothing was charged.' });
    expect(server.requests.some((r) => r.url.endsWith('/confirm'))).toBe(false);
  });

  it("fails with the conversion's own error", async () => {
    const server = fakeApi({ fail: { code: 'conversionFailed', message: 'The file is damaged.' } });
    const api = new ConverterApi(BASE, KEY, server.fetch);
    const error = await convert(api, job, {
      confirmPrice: () => Promise.resolve(true),
      sleep: noSleep,
    }).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(ApiError);
    expect(error).toMatchObject({ code: 'conversionFailed', message: 'The file is damaged.' });
  });

  it('asks again with the same key when the answer was lost: one conversion, one charge', async () => {
    const server = fakeApi({ dropFirstConvert: true });
    const api = new ConverterApi(BASE, KEY, server.fetch);
    const result = await convert(api, job, {
      confirmPrice: () => Promise.resolve(true),
      sleep: noSleep,
      idempotencyKey: () => 'office-fixed',
    });
    const posts = server.requests.filter((r) => r.url.endsWith('/v1/conversions'));
    expect(posts).toHaveLength(2);
    expect(posts.map((p) => p.headers['Idempotency-Key'])).toEqual([
      'office-fixed',
      'office-fixed',
    ]);
    expect(server.jobs.size).toBe(1);
    expect(result.conversion.id).toBe('up1');
  });

  it('gives up after half an hour', async () => {
    const server = fakeApi({ polls: 1000 });
    const api = new ConverterApi(BASE, KEY, server.fetch);
    let now = 0;
    await expect(
      convert(api, job, {
        confirmPrice: () => Promise.resolve(true),
        sleep: noSleep,
        now: () => (now += TIMEOUT_MS / 3),
      }),
    ).rejects.toMatchObject({ code: 'timeout' });
  });

  it('returns a ZIP of pages as it is, which only inserts refuse', async () => {
    const server = fakeApi({ outputExtension: 'zip' });
    const api = new ConverterApi(BASE, KEY, server.fetch);
    const pdf = { ...job, name: 'Scan.pdf', to: 'jpg' };
    const result = await convert(api, pdf, {
      confirmPrice: () => Promise.resolve(true),
      sleep: noSleep,
    });
    expect(result.outputName).toBe('Scan.zip');
    expect(isSingleFile(result, 'jpg')).toBe(false);
    expect(isSingleFile(result, 'zip')).toBe(true);
    expect(isSingleFile({ ...result, outputName: 'Scan.jpg' }, 'jpg')).toBe(true);
  });

  it('makes a fresh idempotency key each time', () => {
    expect(newIdempotencyKey()).not.toBe(newIdempotencyKey());
    expect(newIdempotencyKey()).toMatch(/^office-[0-9a-f-]{32,36}$/);
  });
});
