/**
 * A stand-in for Office.js: exactly the calls the add-in makes, with Office's own shapes and
 * rules, so the tests (and a stub host for runs in a browser) exercise the same
 * code paths a real Office would.
 *
 * Faithful where it matters:
 * - async calls answer through a callback with `{ status: 'succeeded' | 'failed', value, error }`,
 *   later (a task), never synchronously;
 * - requirement sets follow Microsoft's tables for the platform asked for (Excel on the web has
 *   no ImageCoercion, Outlook before Mailbox 1.15 has no Base64 reply attachments…);
 * - Word.run and Excel.run queue commands and apply them only on `context.sync()`; reading a
 *   property that was not loaded and synced throws, as Office does (PropertyNotLoaded);
 * - Outlook's read-mode item has no compose methods and vice versa;
 * - getFileAsync hands the document over in slices of byte arrays and must be closed.
 */

export type HostName = 'Outlook' | 'Word' | 'Excel' | 'PowerPoint';

export interface FakeAttachment {
  readonly id: string;
  readonly name: string;
  readonly bytes: Uint8Array;
  readonly type?: 'file' | 'item' | 'cloud';
  readonly isInline?: boolean;
}

export interface FakeOptions {
  readonly host: HostName | null;
  /** Supported requirement sets and their highest version (defaults: Microsoft 365 on the web). */
  readonly sets?: Readonly<Record<string, string>>;
  /** More sets on top of the defaults (desktop Office: OpenBrowserWindowApi). */
  readonly extraSets?: Readonly<Record<string, string>>;
  readonly mail?: { readonly mode: 'read' | 'compose'; readonly attachments: FakeAttachment[] };
  readonly document?: { readonly url: string; readonly bytes: Uint8Array };
  readonly partitionKey?: string;
  readonly displayLanguage?: string;
  readonly theme?: Partial<Office.OfficeTheme>;
  /** PowerPoint with no text box selected refuses text. */
  readonly noTextSelection?: boolean;
}

/** What the add-in did to the document or message. */
export interface FakeState {
  readonly attachments: FakeAttachment[];
  readonly replies: { name: string; bytes: Uint8Array }[];
  readonly pictures: { base64: string; where: string }[];
  readonly texts: { text: string; where: string }[];
  readonly cells: { address: string; values: unknown[][]; numberFormat: unknown[][] }[];
  /** Excel's picture shapes, as they stand after the last sync. */
  readonly shapes: { left: number; top: number; width: number; lockAspectRatio: boolean }[];
  readonly opened: string[];
  readonly closedFiles: number;
  readonly calls: string[];
}

const WEB_SETS: Readonly<Record<HostName, Readonly<Record<string, string>>>> = {
  Outlook: { Mailbox: '1.15' },
  Word: {
    WordApi: '1.9',
    ImageCoercion: '1.1',
    TextCoercion: '1.1',
    File: '1.1',
    CompressedFile: '1.1',
  },
  // Excel on the web: no ImageCoercion (Microsoft's table lists Windows, Mac and iPad only).
  Excel: { ExcelApi: '1.17', TextCoercion: '1.1', File: '1.1', CompressedFile: '1.1' },
  PowerPoint: {
    PowerPointApi: '1.5',
    ImageCoercion: '1.1',
    TextCoercion: '1.1',
    File: '1.1',
    CompressedFile: '1.1',
  },
};

const later = (fn: () => void) => void setTimeout(fn, 0);
const ok = <T>(callback: ((r: unknown) => void) | undefined, value: T) =>
  later(() => callback?.({ status: 'succeeded', value }));
const failed = (callback: ((r: unknown) => void) | undefined, code: number, message: string) =>
  later(() =>
    callback?.({ status: 'failed', value: undefined, error: { code, name: 'Error', message } }),
  );
/** The callback is the last argument of Office's async methods. */
const callbackOf = (args: unknown[]) =>
  args.findLast((a) => typeof a === 'function') as ((r: unknown) => void) | undefined;

function atLeast(have: string | undefined, want: string): boolean {
  if (!have) return false;
  const a = have.split('.').map(Number);
  const b = want.split('.').map(Number);
  for (let i = 0; i < Math.max(a.length, b.length); i++) {
    if ((a[i] ?? 0) !== (b[i] ?? 0)) return (a[i] ?? 0) > (b[i] ?? 0);
  }
  return true;
}

function toBase64(bytes: Uint8Array): string {
  let s = '';
  for (const b of bytes) s += String.fromCharCode(b);
  return btoa(s);
}

function fromBase64(base64: string): Uint8Array {
  const s = atob(base64);
  return Uint8Array.from(s, (c) => c.charCodeAt(0));
}

/** A loaded-or-not property bag, as Office's proxy objects behave. */
class Proxy<T extends Record<string, unknown>> {
  private loaded = new Set<string>();
  private requested = new Set<string>();
  constructor(
    private readonly values: T,
    private readonly name: string,
  ) {}
  load(props: string): void {
    for (const p of props.split(',')) this.requested.add(p.trim());
  }
  sync(): void {
    for (const p of this.requested) this.loaded.add(p);
    this.requested.clear();
  }
  get(prop: keyof T & string): T[typeof prop] {
    if (!this.loaded.has(prop))
      throw new Error(
        `PropertyNotLoaded: The property '${prop}' of ${this.name} is not available. Load it and call context.sync() before reading it.`,
      );
    return this.values[prop];
  }
}

export function fakeOffice(options: FakeOptions) {
  const host = options.host;
  const sets = { ...(options.sets ?? (host ? WEB_SETS[host] : {})), ...options.extraSets };
  const state: FakeState = {
    attachments: [...(options.mail?.attachments ?? [])],
    replies: [],
    pictures: [],
    texts: [],
    cells: [],
    shapes: [],
    opened: [],
    closedFiles: 0,
    calls: [],
  };
  const mutable = state as { closedFiles: number };
  const changeHandlers: (() => void)[] = [];
  let nextAttachment = 100;

  const details = (a: FakeAttachment) => ({
    id: a.id,
    name: a.name,
    size: a.bytes.length,
    contentType: 'application/octet-stream',
    attachmentType: a.type ?? 'file',
    isInline: a.isInline ?? false,
  });

  const getAttachmentContentAsync = (id: string, ...rest: unknown[]) => {
    state.calls.push(`getAttachmentContentAsync ${id}`);
    const callback = callbackOf(rest);
    const a = state.attachments.find((x) => x.id === id);
    if (!a) return failed(callback, 9040, 'The attachment could not be found.');
    if (a.type === 'cloud')
      return ok(callback, { format: 'url', content: 'https://example.com/file' });
    if (a.type === 'item') return ok(callback, { format: 'eml', content: 'From: someone' });
    ok(callback, { format: 'base64', content: toBase64(a.bytes) });
  };

  let item: Record<string, unknown> | null = null;
  if (host === 'Outlook' && options.mail?.mode === 'read') {
    item = {
      itemType: 'message',
      get attachments() {
        return state.attachments.map(details);
      },
      getAttachmentContentAsync,
      displayReplyFormAsync(data: Office.ReplyFormData, ...rest: unknown[]) {
        state.calls.push('displayReplyFormAsync');
        const callback = callbackOf(rest);
        for (const a of data.attachments ?? []) {
          const attachment = a as Omit<Office.ReplyFormAttachment, 'type'> & { type: string };
          if (attachment.type === 'base64' && !atLeast(sets.Mailbox, '1.15'))
            return failed(callback, 9000, 'The attachment type is not supported.');
          state.replies.push({
            name: attachment.name,
            bytes: fromBase64(attachment.base64file ?? ''),
          });
        }
        ok(callback, undefined);
      },
    };
  } else if (host === 'Outlook' && options.mail?.mode === 'compose') {
    item = {
      itemType: 'message',
      getAttachmentsAsync(...rest: unknown[]) {
        state.calls.push('getAttachmentsAsync');
        ok(callbackOf(rest), state.attachments.map(details));
      },
      getAttachmentContentAsync,
      addFileAttachmentFromBase64Async(base64: string, name: string, ...rest: unknown[]) {
        state.calls.push(`addFileAttachmentFromBase64Async ${name}`);
        const callback = callbackOf(rest);
        if (base64.length > 34_865_152)
          return failed(callback, 9011, 'The attachment is too large.');
        const id = `att${nextAttachment++}`;
        state.attachments.push({ id, name, bytes: fromBase64(base64) });
        ok(callback, id);
        later(() => changeHandlers.forEach((h) => h()));
      },
      removeAttachmentAsync(id: string, ...rest: unknown[]) {
        state.calls.push(`removeAttachmentAsync ${id}`);
        const callback = callbackOf(rest);
        const index = state.attachments.findIndex((x) => x.id === id);
        if (index === -1) return failed(callback, 9002, 'The attachment ID was invalid.');
        state.attachments.splice(index, 1);
        ok(callback, undefined);
        later(() => changeHandlers.forEach((h) => h()));
      },
      addHandlerAsync(type: string, handler: () => void, ...rest: unknown[]) {
        if (type === 'olkAttachmentsChanged') changeHandlers.push(handler);
        ok(callbackOf(rest), undefined);
      },
    };
  }

  const where = () => (host === 'Word' ? 'selection' : host === 'PowerPoint' ? 'slide' : 'cell');
  const document =
    host && host !== 'Outlook'
      ? {
          url: options.document?.url ?? '',
          setSelectedDataAsync(data: string, opts: { coercionType: string }, ...rest: unknown[]) {
            state.calls.push(`setSelectedDataAsync ${opts.coercionType}`);
            const callback = callbackOf(rest);
            if (opts.coercionType === 'image') {
              if (!atLeast(sets.ImageCoercion, '1.1'))
                return failed(callback, 1000, 'The specified coercion type is not supported');
              state.pictures.push({ base64: data, where: where() });
              return ok(callback, undefined);
            }
            if (opts.coercionType === 'text') {
              if (host === 'PowerPoint' && options.noTextSelection)
                return failed(callback, 2001, 'The current selection is not supported.');
              state.texts.push({ text: data, where: where() });
              return ok(callback, undefined);
            }
            failed(callback, 1000, 'The specified coercion type is not supported');
          },
          getFileAsync(type: string, opts: { sliceSize?: number }, ...rest: unknown[]) {
            state.calls.push(`getFileAsync ${type}`);
            const callback = callbackOf(rest);
            const bytes = options.document?.bytes ?? new Uint8Array();
            if (type !== 'compressed') return failed(callback, 1000, 'Not supported.');
            const sliceSize = Math.min(opts.sliceSize ?? 4194304, 4194304);
            const sliceCount = Math.max(1, Math.ceil(bytes.length / sliceSize));
            ok(callback, {
              size: bytes.length,
              sliceCount,
              getSliceAsync(index: number, cb: (r: unknown) => void) {
                const data = Array.from(bytes.subarray(index * sliceSize, (index + 1) * sliceSize));
                ok(cb, { data, index, size: data.length });
              },
              closeAsync(cb?: (r: unknown) => void) {
                mutable.closedFiles += 1;
                ok(cb, undefined);
              },
            });
          },
        }
      : undefined;

  const OfficeFake = {
    HostType: { Word: 'Word', Excel: 'Excel', PowerPoint: 'PowerPoint', Outlook: 'Outlook' },
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    CoercionType: { Text: 'text', Image: 'image', Matrix: 'matrix', Html: 'html' },
    FileType: { Compressed: 'compressed', Pdf: 'pdf', Text: 'text' },
    EventType: { AttachmentsChanged: 'olkAttachmentsChanged' },
    MailboxEnums: {
      AttachmentType: { File: 'file', Item: 'item', Cloud: 'cloud', Base64: 'base64' },
      AttachmentContentFormat: { Base64: 'base64', Url: 'url', Eml: 'eml', ICalendar: 'iCalendar' },
    },
    context: {
      displayLanguage: options.displayLanguage ?? 'en-US',
      partitionKey: options.partitionKey,
      officeTheme: options.theme,
      requirements: {
        isSetSupported: (name: string, version = '1.1') => atLeast(sets[name], version),
      },
      ui: {
        openBrowserWindow: (url: string) => {
          if (!atLeast(sets.OpenBrowserWindowApi, '1.1'))
            throw new Error('openBrowserWindow is not supported here.');
          state.opened.push(url);
        },
      },
      mailbox: host === 'Outlook' ? { item } : undefined,
      document,
    },
    onReady(callback?: (info: { host: string | null; platform: string | null }) => void) {
      const info = { host, platform: host ? 'OfficeOnline' : null };
      return new Promise((resolve) =>
        later(() => {
          callback?.(info);
          resolve(info);
        }),
      );
    },
  };

  // Word: the selection takes a picture; commands run on sync.
  const WordFake = {
    run: async <T>(batch: (context: unknown) => Promise<T>) => {
      const queue: (() => void)[] = [];
      const context = {
        document: {
          getSelection: () => ({
            insertInlinePictureFromBase64(base64: string, location: string) {
              queue.push(() => state.pictures.push({ base64, where: `selection:${location}` }));
            },
            insertText(text: string, location: string) {
              queue.push(() => state.texts.push({ text, where: `selection:${location}` }));
            },
          }),
        },
        sync: () => {
          queue.splice(0).forEach((op) => op());
          return Promise.resolve();
        },
      };
      return batch(context);
    },
  };

  // Excel: the active cell, a resized range and picture shapes; loads and writes on sync.
  const ExcelFake = {
    run: async <T>(batch: (context: unknown) => Promise<T>) => {
      const queue: (() => void)[] = [];
      const proxies: { sync(): void }[] = [];
      const makeRange = (address: string, rows: number) => {
        const proxy = new Proxy({ left: 96, top: 40, address }, 'Range');
        proxies.push(proxy);
        let values: unknown[][] = [];
        let numberFormat: unknown[][] = [];
        return {
          load: (props: string) => proxy.load(props),
          get left() {
            return proxy.get('left');
          },
          get top() {
            return proxy.get('top');
          },
          getResizedRange(deltaRows: number, deltaColumns: number) {
            if (deltaColumns !== 0) throw new Error('only rows in this fake');
            return makeRange(`B2:B${2 + deltaRows}`, deltaRows + 1);
          },
          set values(v: unknown[][]) {
            if (v.length !== rows) throw new Error(`InvalidArgument: ${v.length} rows for ${rows}`);
            values = v;
            queue.push(() => state.cells.push({ address, values, numberFormat }));
          },
          set numberFormat(v: unknown[][]) {
            if (v.length !== rows) throw new Error(`InvalidArgument: ${v.length} rows for ${rows}`);
            numberFormat = v;
          },
        };
      };
      const context = {
        workbook: {
          getActiveCell: () => makeRange('B2', 1),
          worksheets: {
            getActiveWorksheet: () => ({
              shapes: {
                addImage(base64: string) {
                  if (!atLeast(sets.ExcelApi, '1.9')) throw new Error('ApiNotFound: addImage');
                  const shape = { left: 0, top: 0, width: 1920 * 0.75, lockAspectRatio: false };
                  const proxy = new Proxy(shape, 'Shape');
                  proxies.push(proxy);
                  queue.push(() => {
                    state.pictures.push({ base64, where: 'sheet' });
                    state.shapes.push(shape);
                  });
                  return {
                    set lockAspectRatio(v: boolean) {
                      shape.lockAspectRatio = v;
                    },
                    set left(v: number) {
                      shape.left = v;
                    },
                    set top(v: number) {
                      shape.top = v;
                    },
                    set width(v: number) {
                      queue.push(() => (shape.width = v));
                    },
                    get width() {
                      return proxy.get('width') as number;
                    },
                    load: (props: string) => proxy.load(props),
                  };
                },
              },
            }),
          },
        },
        sync: () => {
          queue.splice(0).forEach((op) => op());
          proxies.forEach((p) => p.sync());
          return Promise.resolve();
        },
      };
      return batch(context);
    },
  };

  return {
    Office: OfficeFake as unknown as typeof Office,
    Word: WordFake as unknown as typeof Word,
    Excel: ExcelFake as unknown as typeof Excel,
    state,
  };
}
