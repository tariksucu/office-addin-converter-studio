import { describe, expect, it } from 'vitest';
import { ApiError, CLIENT_HEADER, ConverterApi } from '../src/api';
import { version } from '../package.json';
import { BASE, KEY, fakeApi } from './fake-api';

describe('ConverterApi', () => {
  it('names the add-in and its version on every API call, and sends no cookies', async () => {
    const server = fakeApi();
    const api = new ConverterApi(BASE, KEY, server.fetch);
    await api.account();
    const [request] = server.requests;
    expect(CLIENT_HEADER).toEqual({ 'X-Converter-Studio-Client': `office/${version}` });
    expect(request!.headers['X-Converter-Studio-Client']).toBe(`office/${version}`);
    expect(request!.headers.Authorization).toBe(`Bearer ${KEY}`);
    expect(request!.credentials).toBe('omit');
  });

  it('lists the formats without a key', async () => {
    const server = fakeApi();
    const formats = await new ConverterApi(BASE, '', server.fetch).formats();
    expect(formats.length).toBeGreaterThan(50);
    expect(server.requests[0]!.headers.Authorization).toBeUndefined();
  });

  it('sends the file to its upload link without the key, and asks for an idempotent conversion', async () => {
    const server = fakeApi();
    const api = new ConverterApi(BASE, KEY, server.fetch);
    const upload = await api.createUpload('photo.heic', 4);
    await api.upload(upload, new Blob([new Uint8Array([1, 2, 3, 4])]));
    const put = server.requests.find((r) => r.method === 'PUT')!;
    expect(put.headers.Authorization).toBeUndefined();
    expect(put.url).toBe(upload.uploadUrl);
    const conversion = await api.convert(
      { uploadId: upload.id, to: 'jpg', maxCredits: 3 },
      'office-key-1',
    );
    expect(conversion.status).toBe('queued');
    const post = server.requests.at(-1)!;
    expect(post.headers['Idempotency-Key']).toBe('office-key-1');
    expect(post.body).toEqual({ uploadId: upload.id, to: 'jpg', maxCredits: 3 });
  });

  it('uploads a large file in parts with their offsets', async () => {
    const server = fakeApi();
    const api = new ConverterApi(BASE, KEY, server.fetch);
    const upload = await api.createUpload('big.pdf', 10);
    const seen: number[] = [];
    await api.upload(upload, new Blob([new Uint8Array(10)]), {
      partBytes: 4,
      onProgress: (sent) => seen.push(sent),
    });
    const offsets = server.requests
      .filter((r) => r.method === 'PUT')
      .map((r) => r.headers['Upload-Offset']);
    expect(offsets).toEqual(['0', '4', '8']);
    expect(seen).toEqual([4, 8, 10]);
  });

  it("shows the API's own error message", async () => {
    const server = fakeApi();
    const api = new ConverterApi(BASE, `cvs_live_${'b'.repeat(32)}`, server.fetch);
    const error = await api.account().catch((e: unknown) => e);
    expect(error).toBeInstanceOf(ApiError);
    expect(error).toMatchObject({ code: 'invalidApiKey', status: 401 });
    expect((error as ApiError).message).toMatch(/API key does not work/);
  });

  it('turns a failed fetch into a plain network error', async () => {
    const api = new ConverterApi(BASE, KEY, () => Promise.reject(new TypeError('Failed to fetch')));
    await expect(api.account()).rejects.toMatchObject({ code: 'network' });
  });

  it('never sends the key or a file to a link that is not on the API', async () => {
    const server = fakeApi();
    const api = new ConverterApi(BASE, KEY, server.fetch);
    await expect(api.download('https://evil.example/jobs/1/download')).rejects.toMatchObject({
      code: 'invalidRequest',
    });
    await expect(
      api.upload(
        { id: 'x', uploadUrl: 'https://evil.example/up', maxBytes: 9, partBytes: 9, expiresAt: '' },
        new Blob(['x']),
      ),
    ).rejects.toMatchObject({ code: 'invalidRequest' });
    expect(server.requests).toHaveLength(0);
  });

  it('downloads a result without the key and maps the site route errors', async () => {
    const server = fakeApi();
    const api = new ConverterApi(BASE, KEY, server.fetch);
    await expect(api.download(`${BASE}/jobs/nope/download?token=t-nope`)).rejects.toMatchObject({
      code: 'notFound',
      message: 'Converter Studio could not find this conversion.',
    });
    expect(server.requests[0]!.headers.Authorization).toBeUndefined();
  });
});
