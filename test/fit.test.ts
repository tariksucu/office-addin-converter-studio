import { describe, expect, it } from 'vitest';
import { ConverterApi } from '../src/api';
import { Catalog } from '../src/catalog';
import { MB } from '../src/fit/budget';
import { fitDraft } from '../src/fit/fit';
import { openMail } from '../src/office/mail';
import { messages } from '../src/strings';
import { BASE, FORMATS, KEY, fakeApi, type Scenario } from './fake-api';
import { fakeOffice, type FakeAttachment, type FakeOptions } from './fake-office';
import { fakePhotos, filler, photo } from './fake-photos';

/**
 * "Fit attachments for sending" from start to finish: a draft in the fake Outlook, photos made
 * smaller by the fake encoder (test/fake-photos.ts), the servers' part on the fake API.
 */

interface Run {
  readonly target?: number;
  readonly key?: boolean;
  readonly answer?: boolean;
  readonly scenario?: Partial<Scenario>;
  readonly office?: Partial<FakeOptions>;
  readonly signal?: AbortSignal;
}

async function fit(attachments: FakeAttachment[], options: Run = {}) {
  const fake = fakeOffice({
    host: 'Outlook',
    mail: { mode: 'compose', attachments },
    ...options.office,
  });
  const server = fakeApi(options.scenario);
  const photos = fakePhotos();
  const asked: { credits: number; names: string[] }[] = [];
  const said: string[] = [];
  const report = await fitDraft(
    {
      mail: openMail(fake.Office)!,
      catalog: new Catalog(FORMATS),
      photos: photos.encoder,
      api: options.key === false ? null : new ConverterApi(BASE, KEY, server.fetch),
      t: messages('en-US'),
      say: (text) => said.push(text),
      ask: (credits, names) => {
        asked.push({ credits, names: [...names] });
        return Promise.resolve(options.answer ?? true);
      },
      signal: options.signal,
      sleep: () => Promise.resolve(),
    },
    options.target ?? 10,
  );
  const names = () => fake.state.attachments.map((a) => a.name);
  const bytesOf = (name: string) => fake.state.attachments.find((a) => a.name === name)!.bytes;
  const paths = () => server.requests.map((r) => `${r.method} ${new URL(r.url).pathname}`);
  return { report, fake, server, asked, said, photos, names, bytesOf, paths };
}

const att = (id: string, name: string, bytes: Uint8Array, more: Partial<FakeAttachment> = {}) => ({
  id,
  name,
  bytes,
  ...more,
});

describe('the audit’s first case: a 20 MB limit (10 MB of attachments) and three attachments', () => {
  const draft = () => [
    att('a1', 'IMG_0001.HEIC', photo(4032, 3024, 3 * MB)),
    att('a2', 'IMG_0002.jpg', photo(4032, 3024, 7 * MB)),
    att('a3', 'Contract.pdf', filler(12 * MB, '%PDF-1.7')),
  ];

  it('photos on the computer, the PDF on the servers after its price; all under 10 MB, in order', async () => {
    const run = await fit(draft());
    expect(run.report).toMatchObject({
      kind: 'fits',
      target: 10,
      count: 3,
      before: 22 * MB,
      credits: 1,
    });
    expect(run.report.after).toBeLessThanOrEqual(10 * MB);
    // Names and order kept; only the HEIC's extension changes with its type.
    expect(run.names()).toEqual(['IMG_0001.jpg', 'IMG_0002.jpg', 'Contract.pdf']);
    expect(run.report.changes.map((c) => [c.name, c.newName, c.by])).toEqual([
      ['IMG_0001.HEIC', 'IMG_0001.jpg', 'computer'],
      ['IMG_0002.jpg', 'IMG_0002.jpg', 'computer'],
      ['Contract.pdf', 'Contract.pdf', 'servers'],
    ]);
    expect(new TextDecoder().decode(run.bytesOf('Contract.pdf'))).toBe('converted');
    // Asked once, with the API's price, before anything was charged.
    expect(run.asked).toEqual([{ credits: 1, names: ['Contract.pdf'] }]);
    const create = run.server.requests.find((r) => r.url.endsWith('/v1/conversions'))!;
    expect(create.body).toMatchObject({
      to: 'pdf',
      maxCredits: 0,
      options: { emailFit: true, compression: 'balanced' },
    });
    expect(run.paths()).toContain('POST /api/v1/conversions/up1/confirm');
  });

  it('removes an original only after its replacement is in the draft', async () => {
    const run = await fit(draft());
    const calls = run.fake.state.calls.filter((c) => /^(add|remove)/.test(c));
    expect(calls).toEqual([
      'addFileAttachmentFromBase64Async IMG_0001.jpg',
      'removeAttachmentAsync a1',
      'addFileAttachmentFromBase64Async IMG_0002.jpg',
      'removeAttachmentAsync a2',
      'addFileAttachmentFromBase64Async Contract.pdf',
      'removeAttachmentAsync a3',
    ]);
  });

  it('without a key: the photos still, the PDF named with what a key would do', async () => {
    const run = await fit(draft(), { key: false });
    expect(run.report.kind).toBe('over');
    expect(run.report.needKey).toEqual(['Contract.pdf']);
    expect(run.report.tooBig).toEqual([{ name: 'Contract.pdf', size: 12 * MB }]);
    expect(run.names()).toEqual(['IMG_0001.jpg', 'IMG_0002.jpg', 'Contract.pdf']);
    expect(run.server.requests).toEqual([]);
  });

  it('price declined: nothing charged, the photos still replaced', async () => {
    const run = await fit(draft(), { answer: false });
    expect(run.report).toMatchObject({ kind: 'over', credits: 0, skipped: ['Contract.pdf'] });
    expect(run.paths().some((p) => p.endsWith('/confirm'))).toBe(false);
    expect(run.names()).toEqual(['IMG_0001.jpg', 'IMG_0002.jpg', 'Contract.pdf']);
    expect(run.bytesOf('Contract.pdf').length).toBe(12 * MB);
  });

  it('a conversion that fails leaves its original, with the API’s message', async () => {
    const run = await fit(draft(), {
      scenario: { fail: { code: 'emailTooBig', message: 'Even compressed, it stays over 16 MB.' } },
    });
    expect(run.report.failed).toEqual([
      { name: 'Contract.pdf', message: 'Even compressed, it stays over 16 MB.' },
    ]);
    expect(run.bytesOf('Contract.pdf').length).toBe(12 * MB);
    expect(run.report.kind).toBe('over');
  });
});

describe('cheapest first', () => {
  it('nothing changes when the draft fits already', async () => {
    const run = await fit([att('a1', 'IMG.jpg', photo(4032, 3024, 5 * MB))], { target: 16 });
    expect(run.report).toMatchObject({ kind: 'alreadyFits', before: 5 * MB, changes: [] });
    expect(run.fake.state.calls).toEqual(['getAttachmentsAsync']);
  });

  it('largest photo first, and stops as soon as the draft fits', async () => {
    const run = await fit(
      [
        att('a1', 'notes.txt', filler(2000)),
        att('a2', 'small.jpg', photo(2000, 1500, 2 * MB)),
        att('a3', 'big.jpg', photo(4032, 3024, 9 * MB)),
      ],
      { target: 10 },
    );
    expect(run.report.kind).toBe('fits');
    expect(run.report.changes.map((c) => c.name)).toEqual(['big.jpg']);
    // The small photo was never read, let alone re-encoded (it keeps its metadata).
    expect(run.fake.state.calls).not.toContain('getAttachmentContentAsync a2');
    expect(run.names()).toEqual(['notes.txt', 'small.jpg', 'big.jpg']);
    expect(run.server.requests).toEqual([]);
  });

  it('an opaque PNG becomes JPG, a see-through one stays PNG; smaller still only if that fits', async () => {
    const draft = [
      att('a1', 'scan.pdf', filler(8.8 * MB, '%PDF')),
      att('a2', 'photo.png', photo(4032, 3024, 6 * MB)),
      att('a3', 'logo.png', photo(2500, 2500, 3 * MB, true)),
    ];
    const run = await fit(draft, { key: false, target: 10 });
    expect(run.names()).toEqual(['scan.pdf', 'photo.jpg', 'logo.png']);
    // 8.8 MB of PDF leaves 1.2 MB: both photos had to go below their first version.
    expect(run.report.kind).toBe('fits');
    const sizes = Object.fromEntries(run.report.changes.map((c) => [c.newName, c.after]));
    expect(sizes['photo.jpg']).toBeLessThan(800_000);
    expect(sizes['logo.png']).toBeLessThan(1_500_000);
  });

  it('photos are not degraded further when even that would not make the draft fit', async () => {
    const draft = [
      att('a1', 'scan.pdf', filler(9.9 * MB, '%PDF')),
      att('a2', 'photo.jpg', photo(4032, 3024, 6 * MB)),
    ];
    const run = await fit(draft, { key: false, target: 10 });
    expect(run.report.kind).toBe('over');
    // Its first version (2048 px, quality 0.85), not the smallest.
    expect(run.report.changes).toEqual([
      expect.objectContaining({ name: 'photo.jpg', after: Math.round(2048 * 1536 * 0.85 * 0.3) }),
    ]);
    expect(run.report.tooBig.map((x) => x.name)).toEqual(['scan.pdf']);
  });

  it('spends no credits when nothing could make the draft fit', async () => {
    const run = await fit(
      [
        att('a1', 'huge.docx', filler(20 * MB, 'PK')),
        att('a2', 'scan.pdf', filler(5 * MB, '%PDF')),
        att('a3', 'photo.jpg', photo(4032, 3024, 5 * MB)),
      ],
      { target: 16 },
    );
    expect(run.report.kind).toBe('over');
    expect(run.server.requests).toEqual([]);
    expect(run.asked).toEqual([]);
    expect(run.report.tooBig.map((x) => x.name)).toEqual(['huge.docx']);
  });
});

describe('the servers', () => {
  it('one question for every file, results named by their new type', async () => {
    const run = await fit(
      [
        att('a1', 'clip.mov', filler(30 * MB, 'moov')),
        att('a2', 'IMG_9.HEIC', filler(3 * MB, 'ftypheic')),
        att('a3', 'notes.txt', filler(1000)),
      ],
      { target: 16 },
    );
    expect(run.asked).toEqual([{ credits: 2, names: ['clip.mov', 'IMG_9.HEIC'] }]);
    const jobs = [...run.server.jobs.values()].map((j) => [j.inputName, j.to, j.options]);
    expect(jobs).toEqual([
      ['clip.mov', 'mp4', { emailFit: true, maxSize: 14 }],
      ['IMG_9.HEIC', 'jpg', { maxKb: 1454, longSide: '1920' }],
    ]);
    expect(run.names()).toEqual(['clip.mp4', 'IMG_9.jpg', 'notes.txt']);
    expect(run.report).toMatchObject({ kind: 'fits', credits: 2 });
    // The HEIC the computer could not read was tried there first.
    expect(run.photos.opened).toHaveLength(1);
  });
});

describe('an MP4 and an MP3 on the servers', () => {
  const draft = () => [
    att('a1', 'Holiday.mp4', filler(30 * MB, 'ftypisom')),
    att('a2', 'Interview.mp3', filler(12 * MB, 'ID3')),
    att('a3', 'notes.txt', filler(1000)),
  ];

  it('stay an MP4 and an MP3, smaller, with the same names', async () => {
    const run = await fit(draft(), { target: 16 });
    const jobs = [...run.server.jobs.values()].map((j) => [j.inputName, j.to, j.options]);
    expect(jobs).toEqual([
      ['Holiday.mp4', 'mp4', { emailFit: true, maxSize: 11 }],
      ['Interview.mp3', 'mp3', { emailFit: true }],
    ]);
    expect(run.names()).toEqual(['Holiday.mp4', 'Interview.mp3', 'notes.txt']);
    expect(run.report.changes.map((c) => [c.name, c.newName, c.by])).toEqual([
      ['Holiday.mp4', 'Holiday.mp4', 'servers'],
      ['Interview.mp3', 'Interview.mp3', 'servers'],
    ]);
    expect(run.report).toMatchObject({ kind: 'fits', credits: 2, kept: [], failed: [] });
  });

  it('one the servers could not make smaller stays as it was, free, and is no failure', async () => {
    const run = await fit(draft(), { target: 16, scenario: { notSmaller: ['Interview.mp3'] } });
    expect(run.names()).toEqual(['Holiday.mp4', 'Interview.mp3', 'notes.txt']);
    expect(run.bytesOf('Interview.mp3').length).toBe(12 * MB);
    // The upload itself is the result: not fetched.
    const fetched = run.paths().filter((p) => p.startsWith('GET /api/jobs/'));
    expect(fetched).toEqual(['GET /api/jobs/up1/download']);
    expect(run.report).toMatchObject({
      kind: 'fits',
      credits: 1,
      kept: ['Interview.mp3'],
      failed: [],
    });
    expect(run.report.changes.map((c) => c.name)).toEqual(['Holiday.mp4']);
  });
});

describe('safety', () => {
  const draft = () => [
    att('a1', 'IMG_1.HEIC', photo(4032, 3024, 6 * MB)),
    att('a2', 'IMG_2.jpg', photo(4032, 3024, 6 * MB)),
    att('a3', 'notes.txt', filler(1000)),
  ];

  it('an original whose replacement Outlook refuses stays, the others go ahead', async () => {
    const run = await fit(draft(), { office: { refuseAttach: ['IMG_1.jpg'] }, target: 2 });
    expect(run.names()).toEqual(['IMG_1.HEIC', 'IMG_2.jpg', 'notes.txt']);
    expect(run.bytesOf('IMG_1.HEIC').length).toBe(6 * MB);
    expect(run.report.failed).toEqual([
      { name: 'IMG_1.HEIC', message: 'The attachment could not be added.' },
    ]);
    expect(run.report.changes.map((c) => c.name)).toEqual(['IMG_2.jpg']);
  });

  it('an original Outlook will not remove keeps its place, and the copy goes again', async () => {
    const run = await fit(draft(), { office: { refuseRemove: ['a1'] }, target: 2 });
    expect(run.names().filter((n) => n.startsWith('IMG_1'))).toEqual(['IMG_1.HEIC']);
    expect(run.report.failed[0]).toEqual({
      name: 'IMG_1.HEIC',
      message: 'The attachment could not be removed.',
    });
  });

  it('pictures in the body, attached emails and OneDrive links are never touched', async () => {
    const run = await fit(
      [
        att('a1', 'signature.png', photo(4000, 4000, 2 * MB), { isInline: true }),
        att('a2', 'Fwd: offer.eml', filler(1 * MB), { type: 'item' }),
        att('a3', 'Deck.pptx', filler(30 * MB), { type: 'cloud' }),
        att('a4', 'photo.jpg', photo(4032, 3024, 9 * MB)),
      ],
      { target: 10 },
    );
    // The body's picture and the email count towards the total; the link does not.
    expect(run.report).toMatchObject({ kind: 'fits', count: 3, before: 12 * MB });
    for (const id of ['a1', 'a2', 'a3'])
      expect(run.fake.state.calls).not.toContain(`getAttachmentContentAsync ${id}`);
    expect(run.names()).toEqual(['signature.png', 'Fwd: offer.eml', 'Deck.pptx', 'photo.jpg']);
  });

  it('Cancel before anything is replaced changes nothing', async () => {
    const controller = new AbortController();
    const fake = fakeOffice({ host: 'Outlook', mail: { mode: 'compose', attachments: draft() } });
    const run = fitDraft(
      {
        mail: openMail(fake.Office)!,
        catalog: new Catalog(FORMATS),
        photos: fakePhotos().encoder,
        api: null,
        t: messages('en-US'),
        // Canceled while the first photo is being made smaller.
        say: (text) => text.startsWith('Making photos') && controller.abort(),
        ask: () => Promise.resolve(true),
        signal: controller.signal,
      },
      10,
    );
    await expect(run).rejects.toMatchObject({ code: 'canceled' });
    expect(fake.state.attachments.map((a) => a.name)).toEqual([
      'IMG_1.HEIC',
      'IMG_2.jpg',
      'notes.txt',
    ]);
    expect(fake.state.calls.some((c) => c.startsWith('add'))).toBe(false);
  });

  it('a moving WebP is left as it is, never flattened to its first frame', async () => {
    const moving = photo(4032, 3024, 6 * MB);
    // Not what fakePhotos reads (it would open it): a WebP whose VP8X header says "animated".
    moving.set(new TextEncoder().encode('RIFF\0\0\0\0WEBPVP8X'));
    moving[20] = 0x02;
    const run = await fit(
      [att('a1', 'party.webp', moving), att('a2', 'photo.jpg', photo(4032, 3024, 6 * MB))],
      { target: 2, key: false },
    );
    expect(run.photos.opened).toHaveLength(1);
    expect(run.report.changes.map((c) => c.name)).toEqual(['photo.jpg']);
    expect(run.report.needKey).toEqual([]);
    expect(run.bytesOf('party.webp').length).toBe(6 * MB);
  });

  it('a draft without attachments says so', async () => {
    const run = await fit([]);
    expect(run.report.kind).toBe('empty');
  });
});
