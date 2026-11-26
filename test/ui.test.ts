import { JSDOM } from 'jsdom';
import { describe, expect, it } from 'vitest';
import { mountApp } from '../src/ui/app';
import { BASE, KEY, fakeApi, type Scenario } from './fake-api';
import { fakeOffice, type FakeOptions } from './fake-office';
import { fakePhotos, filler, photo } from './fake-photos';

/**
 * The task pane in a DOM (jsdom), with the fake Office host and the fake API: what a user sees
 * and clicks, from the first run to a converted attachment or an inserted picture.
 */

function storage(): Storage {
  const data = new Map<string, string>();
  return {
    get length() {
      return data.size;
    },
    clear: () => data.clear(),
    getItem: (k) => data.get(k) ?? null,
    key: (i) => [...data.keys()][i] ?? null,
    removeItem: (k) => void data.delete(k),
    setItem: (k, v) => void data.set(k, v),
  };
}

async function until<T>(check: () => T | null | undefined | false, what = 'the pane'): Promise<T> {
  for (let i = 0; i < 200; i++) {
    const value = check();
    if (value) return value;
    await new Promise((r) => setTimeout(r, 5));
  }
  throw new Error(`timed out waiting for ${what}`);
}

async function mount(
  office: FakeOptions,
  options: { key?: string | null; scenario?: Partial<Scenario>; store?: Storage } = {},
) {
  const dom = new JSDOM('<!doctype html><html><body><div id="app"></div></body></html>');
  const win = dom.window;
  const opened: string[] = [];
  win.open = (url: string) => {
    opened.push(url);
    return null;
  };
  const fake = fakeOffice({ partitionKey: 'p1:', ...office });
  const server = fakeApi(options.scenario);
  const store = options.store ?? storage();
  if (options.key !== null) store.setItem('p1:converter-studio.api-key', options.key ?? KEY);
  const root = win.document.getElementById('app')!;
  await mountApp({
    office: fake.Office,
    word: fake.Word,
    excel: fake.Excel,
    host: office.host,
    win: win as unknown as Window,
    root,
    apiBase: BASE,
    fetch: server.fetch,
    storage: store,
    photos: fakePhotos().encoder,
  });
  const text = () => root.textContent ?? '';
  const button = (label: string) =>
    [...root.querySelectorAll('button')].find((b) => b.textContent === label && !b.disabled);
  const click = async (label: string) => (await until(() => button(label), label)).click();
  const status = () => root.querySelector<HTMLElement>('.status');
  return { dom, win, root, fake, server, store, opened, text, button, click, status };
}

function choose(pane: Awaited<ReturnType<typeof mount>>, selector: string, file: File) {
  const input =
    pane.root.querySelectorAll<HTMLInputElement>('input[type=file]')[
      selector === 'picture' ? 0 : 1
    ]!;
  Object.defineProperty(input, 'files', { value: [file], configurable: true });
  input.dispatchEvent(new pane.win.Event('change'));
}

const bytes = (text: string) => new TextEncoder().encode(text);
const heic = { id: 'a1', name: 'IMG_0001.HEIC', bytes: bytes('heic') };
const docx = { id: 'a2', name: 'Report.docx', bytes: bytes('docx') };

describe('first run', () => {
  it('outside Office, says what the page is', async () => {
    const pane = await mount({ host: null });
    expect(pane.text()).toMatch(/Open this page from Outlook, Word, Excel or PowerPoint/);
  });

  it('shows what the add-in does before asking for a key, with the sign-up and key links', async () => {
    const pane = await mount(
      { host: 'Outlook', mail: { mode: 'read', attachments: [] } },
      { key: null },
    );
    expect(pane.text()).toMatch(/Convert the attachments of this message/);
    expect(pane.root.querySelector('input[type=password]')).not.toBeNull();
    await pane.click('Create an account');
    await pane.click('Create an API key');
    await pane.click('Privacy policy');
    expect(pane.opened).toEqual([
      'https://converter.studio/sign-up',
      'https://converter.studio/account#api-keys',
      'https://converter.studio/privacy',
    ]);
    // Nothing but the format list is asked before a key is entered.
    expect(pane.server.requests.map((r) => r.url)).toEqual([`${BASE}/v1/formats`]);
  });

  it('turns away what is not a key, then shows the API’s answer for a wrong one', async () => {
    const pane = await mount({ host: 'Word' }, { key: null });
    const input = pane.root.querySelector<HTMLInputElement>('input[type=password]')!;
    input.value = 'hello';
    await pane.click('Connect');
    expect(pane.text()).toMatch(/That is not a Converter Studio API key/);
    input.value = `cvs_live_${'z'.repeat(32)}`;
    await pane.click('Connect');
    await until(() => pane.text().includes('This API key does not work'), 'the API error');
    expect(pane.store.length).toBe(0);
  });

  it('connects with a key, keeps it under the partition key and shows the account', async () => {
    const pane = await mount({ host: 'Word' }, { key: null });
    pane.root.querySelector<HTMLInputElement>('input[type=password]')!.value = `  ${KEY} `;
    await pane.click('Connect');
    await until(() => pane.text().includes('office@example.com · 12 credits left'), 'the account');
    expect(pane.store.getItem('p1:converter-studio.api-key')).toBe(KEY);
    await pane.click('Disconnect');
    expect(pane.store.length).toBe(0);
    expect(pane.root.querySelector('input[type=password]')).not.toBeNull();
  });

  it('forgets a key the account page has deleted', async () => {
    const pane = await mount({ host: 'Excel' }, { key: `cvs_live_${'q'.repeat(32)}` });
    expect(pane.text()).toMatch(/This API key does not work/);
    expect(pane.store.length).toBe(0);
  });

  it('takes the Office theme’s colours where Office gives them', async () => {
    const pane = await mount({
      host: 'Word',
      theme: {
        bodyBackgroundColor: '#1e1e1e',
        bodyForegroundColor: '#ffffff',
        controlBackgroundColor: '#2d2d2d',
        controlForegroundColor: '#eeeeee',
      },
    });
    const html = pane.win.document.documentElement;
    expect(html.dataset.theme).toBe('office');
    expect(html.style.getPropertyValue('--bg')).toBe('#1e1e1e');
  });
});

describe('Outlook', () => {
  it('while writing: converts a HEIC attachment to JPG and replaces the original', async () => {
    const pane = await mount({
      host: 'Outlook',
      mail: { mode: 'compose', attachments: [heic, docx] },
    });
    await until(() => pane.text().includes('IMG_0001.HEIC'), 'the attachments');
    const select = pane.root.querySelector<HTMLSelectElement>('.attachments select')!;
    expect(select.value).toBe('jpg');
    expect(pane.root.querySelector(`label[for="${select.id}"]`)!.textContent).toBe(
      'Convert IMG_0001.HEIC to',
    );
    pane.root.querySelector<HTMLButtonElement>('.attachments button.primary')!.click();
    await until(() => pane.text().includes('IMG_0001.jpg (9 B). 1 credit used.'), 'the result');
    await pane.click('Replace the original');
    await until(
      () => pane.text().includes('IMG_0001.jpg replaced the original.'),
      'the replacement',
    );
    expect(pane.fake.state.attachments.map((a) => a.name)).toEqual(['Report.docx', 'IMG_0001.jpg']);
    expect(new TextDecoder().decode(pane.fake.state.attachments[1]!.bytes)).toBe('converted');
    // The list follows the draft.
    await until(() => !pane.text().includes('IMG_0001.HEIC'), 'the refreshed list');
  });

  it('while writing: attaches the result beside the original', async () => {
    const pane = await mount({ host: 'Outlook', mail: { mode: 'compose', attachments: [docx] } });
    await until(() => pane.text().includes('Report.docx'));
    pane.root.querySelector<HTMLButtonElement>('.attachments button.primary')!.click();
    await pane.click('Attach it');
    await until(() => pane.text().includes('Report.pdf is attached.'));
    expect(pane.fake.state.attachments.map((a) => a.name)).toEqual(['Report.docx', 'Report.pdf']);
  });

  it('reading: replies with the converted file, or downloads it', async () => {
    const pane = await mount({ host: 'Outlook', mail: { mode: 'read', attachments: [docx] } });
    await until(() => pane.text().includes('Report.docx'));
    pane.root.querySelector<HTMLButtonElement>('.attachments button.primary')!.click();
    await pane.click('Download');
    expect(pane.opened).toEqual([`${BASE}/jobs/up1/download?token=t-up1`]);
    await pane.click('Reply with it');
    await until(() => pane.text().includes('A reply with the file is open.'));
    expect(pane.fake.state.replies.map((r) => r.name)).toEqual(['Report.pdf']);
  });

  it('asks before a dearer conversion and goes ahead when told', async () => {
    const pane = await mount(
      { host: 'Outlook', mail: { mode: 'read', attachments: [docx] } },
      { scenario: { credits: 5 } },
    );
    await until(() => pane.text().includes('Report.docx'));
    pane.root.querySelector<HTMLButtonElement>('.attachments button.primary')!.click();
    await until(() => pane.text().includes('This conversion costs 5 credits. Convert it?'));
    pane.status()!.querySelector<HTMLButtonElement>('button.primary')!.click();
    await until(() => pane.text().includes('5 credits used.'));
  });

  it("shows a failed conversion's own message as an alert", async () => {
    const pane = await mount(
      { host: 'Outlook', mail: { mode: 'read', attachments: [heic] } },
      { scenario: { fail: { code: 'conversionFailed', message: 'This HEIC file is damaged.' } } },
    );
    await until(() => pane.text().includes('IMG_0001.HEIC'));
    pane.root.querySelector<HTMLButtonElement>('.attachments button.primary')!.click();
    await until(() => pane.text().includes('This HEIC file is damaged.'));
    expect(pane.status()!.getAttribute('role')).toBe('alert');
  });

  it('says when there is nothing to convert', async () => {
    const pane = await mount({ host: 'Outlook', mail: { mode: 'read', attachments: [] } });
    await until(() => pane.text().includes('This message has no file attachments.'));
    const other = await mount({
      host: 'Outlook',
      mail: { mode: 'read', attachments: [{ id: 'z', name: 'data.xyz', bytes: bytes('?') }] },
    });
    await until(() => other.text().includes('Converter Studio cannot convert this file.'));
  });
});

describe('Outlook: fit attachments for sending', () => {
  const big = { id: 'b1', name: 'IMG_0007.jpg', bytes: photo(4032, 3024, 12_000_000) };
  const notes = { id: 'b2', name: 'notes.txt', bytes: filler(1000) };
  const contract = { id: 'b3', name: 'Contract.pdf', bytes: filler(12_000_000, '%PDF') };
  const fitButton = (pane: Awaited<ReturnType<typeof mount>>) =>
    pane.root.querySelector<HTMLButtonElement>('section.fit button.primary')!;

  it('works before any key: photos on the computer, the choice of total remembered', async () => {
    const pane = await mount(
      { host: 'Outlook', mail: { mode: 'compose', attachments: [big, notes] } },
      { key: null },
    );
    const section = await until(() => pane.root.querySelector('section.fit'), 'the fit section');
    // Before the key's explanation and form.
    expect(pane.text().indexOf('Fit attachments for sending')).toBeLessThan(
      pane.text().indexOf('You need an account and an API key'),
    );
    expect(section.textContent).toMatch(/lose their metadata, location included/);
    const select = section.querySelector('select')!;
    expect(pane.root.querySelector(`label[for="${select.id}"]`)!.textContent).toBe('Total size');
    expect(select.value).toBe('16');
    select.value = '10';
    select.dispatchEvent(new pane.win.Event('change'));
    expect(pane.store.getItem('p1:converter-studio.fit-target')).toBe('10');
    fitButton(pane).click();
    await until(() => pane.status()?.dataset.kind === 'done', 'the outcome');
    expect(pane.status()!.textContent).toContain(
      '2 attachments, 12 MB → 803 KB. Fits Gmail, Yahoo Mail, Outlook.com and iCloud Mail.',
    );
    expect(pane.fake.state.attachments.map((a) => a.name)).toEqual(['IMG_0007.jpg', 'notes.txt']);
    expect(pane.server.requests.map((r) => r.url)).toEqual([`${BASE}/v1/formats`]);
  });

  it('without a key, says what a key would do and links to the account page', async () => {
    const scan = { ...contract, bytes: filler(15_900_000, '%PDF') };
    const pane = await mount(
      { host: 'Outlook', mail: { mode: 'compose', attachments: [scan, big] } },
      { key: null },
    );
    await until(() => pane.root.querySelector('section.fit'));
    fitButton(pane).click();
    await until(() => pane.status()?.dataset.kind === 'warning', 'the outcome');
    const text = pane.status()!.textContent;
    expect(text).toContain(
      '2 attachments, 27.9 MB → 16.7 MB. Still over 16 MB: Contract.pdf (15.9 MB).',
    );
    expect(text).toContain('Share it as a OneDrive link instead');
    expect(text).toContain(
      'With a Converter Studio API key, Contract.pdf could be made smaller on its servers.',
    );
    pane.status()!.querySelector<HTMLButtonElement>('button.link')!.click();
    expect(pane.opened).toEqual(['https://converter.studio/account#api-keys']);
  });

  it('with a key: the price first, then the result and the credits used', async () => {
    // 10 MB chosen earlier: the photo alone does not bring the draft under it.
    const store = storage();
    store.setItem('p1:converter-studio.fit-target', '10');
    const pane = await mount(
      { host: 'Outlook', mail: { mode: 'compose', attachments: [contract, big, notes] } },
      { store },
    );
    await until(() => pane.text().includes('Contract.pdf'), 'the attachments');
    expect(pane.root.querySelector<HTMLSelectElement>('section.fit select')!.value).toBe('10');
    fitButton(pane).click();
    await until(() => pane.status()?.dataset.kind === 'question', 'the price');
    expect(pane.status()!.textContent).toContain(
      'Making Contract.pdf smaller on Converter Studio’s servers costs 1 credit.',
    );
    expect(pane.server.requests.some((r) => r.url.endsWith('/confirm'))).toBe(false);
    expect([...pane.status()!.querySelectorAll('button')].map((b) => b.textContent)).toEqual([
      'Fit for 1 credit',
      'Cancel',
    ]);
    await pane.click('Fit for 1 credit');
    await until(() => pane.status()?.dataset.kind === 'done', 'the outcome');
    expect(pane.status()!.textContent).toMatch(/3 attachments, 24 MB → 803 KB\. .*1 credit used\./);
    expect(pane.fake.state.attachments.map((a) => a.name)).toEqual([
      'Contract.pdf',
      'IMG_0007.jpg',
      'notes.txt',
    ]);
  });

  it('Cancel at the price charges nothing and keeps the photos made smaller', async () => {
    const store = storage();
    store.setItem('p1:converter-studio.fit-target', '10');
    const pane = await mount(
      { host: 'Outlook', mail: { mode: 'compose', attachments: [contract, big] } },
      { store },
    );
    await until(() => pane.text().includes('Contract.pdf'), 'the attachments');
    fitButton(pane).click();
    await until(() => pane.status()?.dataset.kind === 'question', 'the price');
    await pane.click('Cancel');
    await until(() => pane.status()?.dataset.kind === 'warning', 'the outcome');
    expect(pane.status()!.textContent).toContain(
      'Not converted, so nothing was charged: Contract.pdf.',
    );
    expect(pane.server.requests.some((r) => r.url.endsWith('/confirm'))).toBe(false);
    expect(pane.fake.state.attachments.map((a) => [a.name, a.bytes.length])).toEqual([
      ['Contract.pdf', 12_000_000],
      ['IMG_0007.jpg', 802_161],
    ]);
  });

  it('says which ones the servers could not make smaller: kept, nothing charged', async () => {
    const song = { id: 'b4', name: 'Interview.mp3', bytes: filler(12_000_000, 'ID3') };
    const store = storage();
    store.setItem('p1:converter-studio.fit-target', '10');
    const pane = await mount(
      { host: 'Outlook', mail: { mode: 'compose', attachments: [contract, song] } },
      { store, scenario: { notSmaller: ['Interview.mp3'] } },
    );
    await until(() => pane.text().includes('Interview.mp3'), 'the attachments');
    fitButton(pane).click();
    await until(() => pane.status()?.dataset.kind === 'question', 'the price');
    await pane.click('Fit for 2 credits');
    await until(() => pane.status()?.dataset.kind === 'warning', 'the outcome');
    // In the outcome's one line, after the credits.
    expect(pane.status()!.querySelector('.status-line')!.textContent).toBe(
      '2 attachments, 24 MB → 12 MB. Still over 10 MB: Interview.mp3 (12 MB). 1 credit used. ' +
        'Already as small as it gets, so kept as it was (nothing charged): Interview.mp3.',
    );
    expect(pane.status()!.textContent).not.toContain('stays as it was');
    expect(pane.fake.state.attachments.map((a) => [a.name, a.bytes.length])).toEqual([
      ['Contract.pdf', 9],
      ['Interview.mp3', 12_000_000],
    ]);
  });

  it('says when the attachments fit already, and is not offered on a received message', async () => {
    const pane = await mount({ host: 'Outlook', mail: { mode: 'compose', attachments: [notes] } });
    await until(() => pane.root.querySelector('section.fit'));
    fitButton(pane).click();
    await until(() => pane.status()?.dataset.kind === 'done');
    expect(pane.status()!.textContent).toBe(
      'The attachments are already under 16 MB (1 KB). Nothing was changed.',
    );
    const read = await mount({ host: 'Outlook', mail: { mode: 'read', attachments: [notes] } });
    await until(() => read.text().includes('notes.txt'));
    expect(read.root.querySelector('section.fit')).toBeNull();
  });
});

describe('Word, Excel and PowerPoint', () => {
  it('Word: inserts a HEIC photo as a JPG at the selection', async () => {
    const pane = await mount({ host: 'Word' });
    await until(() => pane.text().includes('Insert a picture'));
    choose(pane, 'picture', new File([bytes('heic')], 'IMG_0001.HEIC'));
    await pane.click('Insert picture');
    await until(() => pane.text().includes('The picture is inserted.'));
    expect(pane.fake.state.pictures).toEqual([
      { base64: btoa('converted'), where: 'selection:Replace' },
    ]);
    const post = pane.server.requests.find((r) => r.url.endsWith('/v1/conversions'))!;
    expect(post.body).toMatchObject({ to: 'jpg', options: { longSide: '1920', quality: 'high' } });
  });

  it('asks for the page of a PDF, and checks it', async () => {
    const pane = await mount({ host: 'PowerPoint' });
    await until(() => pane.text().includes('Insert a picture'));
    const pageField = () =>
      pane.root.querySelector<HTMLInputElement>('input[type=number]')!.closest('.field')!;
    expect((pageField() as HTMLElement).hidden).toBe(true);
    choose(pane, 'picture', new File([bytes('%PDF')], 'Brochure.pdf'));
    expect((pageField() as HTMLElement).hidden).toBe(false);
    pane.root.querySelector<HTMLInputElement>('input[type=number]')!.value = '0';
    await pane.click('Insert picture');
    expect(pane.text()).toMatch(/Write the number of the PDF page/);
    pane.root.querySelector<HTMLInputElement>('input[type=number]')!.value = '4';
    await pane.click('Insert picture');
    await until(() => pane.text().includes('The picture is inserted.'));
    const post = pane.server.requests.find((r) => r.url.endsWith('/v1/conversions'))!;
    expect(post.body).toMatchObject({ to: 'png', options: { pageRange: '4', dpi: '150' } });
    expect(pane.fake.state.pictures[0]!.where).toBe('slide');
  });

  it('does not insert a ZIP of pages', async () => {
    const pane = await mount({ host: 'Word' }, { scenario: { outputExtension: 'zip' } });
    await until(() => pane.text().includes('Insert a picture'));
    choose(pane, 'picture', new File([bytes('heic')], 'IMG_0001.heic'));
    await pane.click('Insert picture');
    await until(() => pane.text().includes('The result is several files in a ZIP'));
    expect(pane.fake.state.pictures).toEqual([]);
  });

  it('attaches a ZIP result in Outlook all the same', async () => {
    const pane = await mount(
      {
        host: 'Outlook',
        mail: {
          mode: 'compose',
          attachments: [{ id: 'p', name: 'Scan.pdf', bytes: bytes('%PDF') }],
        },
      },
      { scenario: { outputExtension: 'zip' } },
    );
    await until(() => pane.text().includes('Scan.pdf'));
    pane.root.querySelector<HTMLButtonElement>('.attachments button.primary')!.click();
    await pane.click('Attach it');
    await until(() => pane.text().includes('Scan.zip is attached.'));
  });

  it('refuses a picture Office opens itself', async () => {
    const pane = await mount({ host: 'Word' });
    await until(() => pane.text().includes('Insert a picture'));
    choose(pane, 'picture', new File([bytes('jpg')], 'photo.jpg'));
    await pane.click('Insert picture');
    expect(pane.text()).toMatch(/Choose a picture or a PDF/);
    expect(pane.server.requests.some((r) => r.method === 'POST')).toBe(false);
  });

  it('Excel: inserts the text of a scan from the selected cell down', async () => {
    const pane = await mount(
      { host: 'Excel' },
      { scenario: { output: bytes('Invoice 42\nTotal 99\n') } },
    );
    await until(() => pane.text().includes('Text from a scan'));
    choose(pane, 'text', new File([bytes('png')], 'scan.png'));
    pane.root.querySelectorAll('select')[1]!.value = 'de';
    await pane.click('Insert text');
    await until(() => pane.text().includes('The text is inserted.'));
    expect(pane.fake.state.cells[0]!.values).toEqual([['Invoice 42'], ['Total 99']]);
    const post = pane.server.requests.find((r) => r.url.endsWith('/v1/conversions'))!;
    expect(post.body).toMatchObject({ to: 'txt', options: { ocrLanguage: 'de' } });
  });

  it('PowerPoint: says to select a text box when there is none', async () => {
    const pane = await mount({ host: 'PowerPoint', noTextSelection: true });
    await until(() => pane.text().includes('Text from a scan'));
    choose(pane, 'text', new File([bytes('%PDF')], 'scan.pdf'));
    await pane.click('Insert text');
    await until(() => pane.text().includes('Select a text box or a shape on the slide'));
  });

  it('exports the document to EPUB and offers the download', async () => {
    const pane = await mount({
      host: 'Word',
      document: {
        url: 'https://contoso.sharepoint.com/Docs/Handbook.docx',
        bytes: bytes('PK docx'),
      },
    });
    await until(() => pane.text().includes('Export this document'));
    const select = pane.root.querySelectorAll('select')[2]!;
    expect([...select.options].map((o) => o.value)).toEqual([
      'epub',
      'txt',
      'doc',
      'html',
      'md',
      'rtf',
      'pdfa',
    ]);
    select.value = 'epub';
    await pane.click('Export');
    await until(() => pane.text().includes('Handbook.epub (9 B) is ready.'));
    const upload = pane.server.requests.find((r) => r.url.endsWith('/v1/uploads'))!;
    expect(upload.body).toEqual({ fileName: 'Handbook.docx', size: 7 });
    expect(pane.fake.state.closedFiles).toBe(1);
    await pane.click('Download');
    expect(pane.opened).toEqual([`${BASE}/jobs/up1/download?token=t-up1`]);
  });

  it('turns every button off while a job runs, and Cancel stops it', async () => {
    const pane = await mount({ host: 'Excel' }, { scenario: { credits: 9 } });
    await until(() => pane.text().includes('Export this document'));
    await pane.click('Export');
    await until(() => pane.text().includes('This conversion costs 9 credits.'));
    expect(pane.button('Insert picture')).toBeUndefined();
    expect(pane.button('Disconnect')).toBeUndefined();
    pane.status()!.querySelectorAll('button')[1]!.click();
    await until(() => pane.text().includes('Canceled. Nothing was charged.'));
    expect(pane.button('Insert picture')).toBeDefined();
  });
});
