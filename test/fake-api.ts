import type { Conversion, FormatInfo } from '../src/api';
import formatsFixture from './fixtures/formats.json';

/**
 * Converter Studio's REST API in memory, behind a `fetch`: what the add-in sends and how it
 * reacts to the answers, without a server. Shapes follow the API reference
 * (https://converter.studio/developers); the format list is a real answer of `GET /v1/formats`
 * (test/fixtures/formats.json).
 */

export const BASE = 'https://converter.studio/api';
export const KEY = `cvs_live_${'a'.repeat(32)}`;
export const FORMATS = (formatsFixture as { formats: FormatInfo[] }).formats;

export interface Scenario {
  /** The price of each conversion. */
  credits: number;
  /** The result's bytes (its size is theirs). */
  output: Uint8Array;
  /** How many status checks before a conversion is done. */
  polls: number;
  /** A conversion that fails with this error. */
  fail?: { code: string; message: string };
  /** POST /v1/conversions fails like this once (network: the answer is lost on the way). */
  dropFirstConvert?: boolean;
  /** The account's remaining credits. */
  accountCredits: number;
  /** The result's extension when it is not the target's (a ZIP of pages). */
  outputExtension?: string;
}

export interface Seen {
  readonly method: string;
  readonly url: string;
  readonly headers: Record<string, string>;
  readonly credentials: RequestCredentials | undefined;
  readonly body: unknown;
}

interface Job {
  id: string;
  status: Conversion['status'];
  to: string;
  inputName: string;
  credits: number;
  polls: number;
  options: unknown;
  uploadSize: number;
}

export function fakeApi(overrides: Partial<Scenario> = {}) {
  const scenario: Scenario = {
    credits: 1,
    output: new TextEncoder().encode('converted'),
    polls: 1,
    accountCredits: 12,
    ...overrides,
  };
  const requests: Seen[] = [];
  const uploads = new Map<
    string,
    { name: string; size: number; received: number; bytes: Blob[] }
  >();
  const jobs = new Map<string, Job>();
  const idempotency = new Map<string, string>();
  let next = 1;
  let dropConvert = scenario.dropFirstConvert ?? false;

  const json = (status: number, body: unknown) =>
    new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json' } });
  const error = (status: number, code: string, message: string, params?: object) =>
    json(status, { error: { code, message, ...(params ? { params } : {}) } });

  const view = (job: Job): Conversion => {
    const done = job.status === 'done';
    return {
      id: job.id,
      status: job.status,
      from: job.inputName.split('.').pop() ?? null,
      to: job.to,
      inputName: job.inputName,
      credits: job.credits,
      progress: done ? 1 : 0.5,
      outputName: done
        ? job.inputName.replace(/\.\w+$/, `.${scenario.outputExtension ?? job.to}`)
        : null,
      outputSize: done ? scenario.output.length : null,
      downloadUrl: done ? `${BASE}/jobs/${job.id}/download?token=t-${job.id}` : null,
      expiresAt: done ? '2026-11-23T12:00:00.000Z' : null,
      error: job.status === 'failed' && scenario.fail ? scenario.fail : null,
    };
  };

  const advance = (job: Job) => {
    if (job.status !== 'queued' && job.status !== 'processing') return;
    job.polls -= 1;
    if (job.polls <= 0) job.status = scenario.fail ? 'failed' : 'done';
    else job.status = 'processing';
  };

  const answer = (input: string, init: RequestInit): Response => {
    const method = init.method ?? 'GET';
    const headers = { ...(init.headers as Record<string, string> | undefined) };
    let body: unknown = init.body;
    if (typeof init.body === 'string') body = JSON.parse(init.body);
    requests.push({ method, url: input, headers, credentials: init.credentials, body });
    const url = new URL(input);
    if (!input.startsWith(`${BASE}/`)) return error(404, 'notFound', 'Not on this server.');
    const path = url.pathname.replace(/^\/api/, '');
    const keyed = headers.Authorization === `Bearer ${KEY}`;

    if (method === 'GET' && path === '/v1/formats') return json(200, { formats: FORMATS });
    if (method === 'PUT' && path.startsWith('/v1/uploads/')) {
      const id = path.split('/')[3]!;
      const upload = uploads.get(id);
      if (!upload || url.searchParams.get('token') !== `t-${id}` || headers.Authorization)
        return error(404, 'notFound', 'Not found.');
      const blob = init.body as Blob;
      const offset = headers['Upload-Offset'] === undefined ? 0 : Number(headers['Upload-Offset']);
      upload.received = offset + blob.size;
      upload.bytes.push(blob);
      return json(200, { id, status: upload.received >= upload.size ? 'ready' : 'partial' });
    }
    if (method === 'GET' && path.startsWith('/jobs/')) {
      const id = path.split('/')[2]!;
      const job = jobs.get(id);
      if (!job || url.searchParams.get('token') !== `t-${id}` || headers.Authorization)
        return json(404, { error: 'notFound' });
      return new Response(new Blob([scenario.output as Uint8Array<ArrayBuffer>]), { status: 200 });
    }
    if (!keyed)
      return error(
        401,
        'invalidApiKey',
        'This API key does not work. Create one on your account page at converter.studio.',
      );
    if (method === 'GET' && path === '/v1/account')
      return json(200, {
        email: 'office@example.com',
        plan: 'free',
        credits: scenario.accountCredits,
        maxFileBytes: 100_000_000,
      });
    if (method === 'POST' && path === '/v1/uploads') {
      const { fileName, size } = body as { fileName: string; size: number };
      const id = `up${next++}`;
      uploads.set(id, { name: fileName, size, received: 0, bytes: [] });
      return json(201, {
        id,
        uploadUrl: `${BASE}/v1/uploads/${id}?token=t-${id}`,
        maxBytes: 100_000_000,
        partBytes: 64 * 1024 * 1024,
        expiresAt: '2026-11-23T12:00:00.000Z',
      });
    }
    if (method === 'POST' && path === '/v1/conversions') {
      const request = body as {
        uploadId: string;
        to: string;
        options?: unknown;
        maxCredits: number;
      };
      const key = headers['Idempotency-Key'];
      if (key && idempotency.has(key)) return json(200, view(jobs.get(idempotency.get(key)!)!));
      const upload = uploads.get(request.uploadId);
      if (!upload) return error(404, 'notFound', 'Converter Studio could not find this upload.');
      uploads.delete(request.uploadId);
      const job: Job = {
        id: request.uploadId,
        status: scenario.credits > request.maxCredits ? 'quoted' : 'queued',
        to: request.to,
        inputName: upload.name,
        credits: scenario.credits,
        polls: scenario.polls,
        options: request.options,
        uploadSize: upload.received,
      };
      jobs.set(job.id, job);
      if (key) idempotency.set(key, job.id);
      if (dropConvert) {
        dropConvert = false;
        throw new TypeError('Failed to fetch');
      }
      return json(201, view(job));
    }
    const status = /^\/v1\/conversions\/([^/]+)(\/confirm)?$/.exec(path);
    if (status) {
      const job = jobs.get(decodeURIComponent(status[1]!));
      if (!job) return error(404, 'notFound', 'Converter Studio could not find this conversion.');
      if (status[2] && method === 'POST') {
        if (job.status === 'quoted') job.status = 'queued';
        return json(200, view(job));
      }
      if (method === 'GET') {
        advance(job);
        return json(200, view(job));
      }
    }
    return error(404, 'notFound', 'Not found.');
  };

  const fetch = (input: string, init: RequestInit = {}): Promise<Response> => {
    try {
      return Promise.resolve(answer(input, init));
    } catch (reason) {
      return Promise.reject(reason instanceof Error ? reason : new Error(String(reason)));
    }
  };

  return { fetch, requests, jobs, uploads, scenario };
}
