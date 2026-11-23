import { HostError, officeCall } from './async';
import { DOCUMENT_FORMAT, type DocumentHost } from '../catalog';

/**
 * The open document in Word, Excel or PowerPoint. Each call is the one that app supports on the
 * web, on Windows and on Mac (Microsoft's requirement-set tables):
 *
 * - a picture: Word `insertInlinePictureFromBase64` at the selection (WordApi 1.2); Excel a
 *   picture shape at the active cell (`shapes.addImage`, ExcelApi 1.9; Excel on the web has no
 *   image coercion); PowerPoint `setSelectedDataAsync` with image coercion (ImageCoercion 1.1)
 *   on the current slide;
 * - text: Word and PowerPoint `setSelectedDataAsync` with text coercion (TextCoercion 1.1) at the
 *   cursor or into the selected text box; Excel one line a cell from the active cell down,
 *   formatted as text first so no line is read as a formula (ExcelApi 1.7);
 * - the document itself: `getFileAsync(Compressed)` read in slices (File 1.1, CompressedFile).
 */

type OfficeApi = typeof Office;

export interface Hosts {
  readonly office: OfficeApi;
  readonly word?: typeof Word;
  readonly excel?: typeof Excel;
}

export interface OpenDocument {
  readonly host: DocumentHost;
  insertPicture(base64: string): Promise<void>;
  insertText(text: string): Promise<void>;
  /** The document as Office saves it (DOCX, XLSX, PPTX), with a file name for it. */
  file(): Promise<{ readonly blob: Blob; readonly name: string }>;
}

/** The widest a picture is placed on a sheet, in points (it keeps its proportions). */
const SHEET_PICTURE_WIDTH = 480;
/** Office hands the document over in slices of up to 4 MB. */
const SLICE_BYTES = 4 * 1024 * 1024;

export function documentHost(office: OfficeApi, host: unknown): DocumentHost | null {
  if (host === office.HostType.Word) return 'word';
  if (host === office.HostType.Excel) return 'excel';
  if (host === office.HostType.PowerPoint) return 'powerpoint';
  return null;
}

export function openDocument(hosts: Hosts, host: DocumentHost): OpenDocument {
  const { office } = hosts;
  const supports = (set: string, version: string) =>
    office.context.requirements.isSetSupported(set, version);
  const setSelected = (data: string, coercionType: Office.CoercionType) =>
    officeCall<void>((done) =>
      office.context.document.setSelectedDataAsync(data, { coercionType }, done as never),
    );
  const needs = (what: string) =>
    new HostError(`This version of Office cannot ${what}. Update Office and try again.`, 'old');

  return {
    host,
    async insertPicture(base64) {
      if (host === 'word' && hosts.word && supports('WordApi', '1.2')) {
        await hosts.word.run(async (context) => {
          context.document.getSelection().insertInlinePictureFromBase64(base64, 'Replace');
          await context.sync();
        });
        return;
      }
      if (host === 'excel') {
        if (!hosts.excel || !supports('ExcelApi', '1.9')) throw needs('insert pictures');
        await hosts.excel.run(async (context) => {
          const cell = context.workbook.getActiveCell();
          cell.load('left,top');
          await context.sync();
          const shape = context.workbook.worksheets.getActiveWorksheet().shapes.addImage(base64);
          shape.lockAspectRatio = true;
          shape.left = cell.left;
          shape.top = cell.top;
          shape.load('width');
          await context.sync();
          if (shape.width > SHEET_PICTURE_WIDTH) {
            shape.width = SHEET_PICTURE_WIDTH;
            await context.sync();
          }
        });
        return;
      }
      if (!supports('ImageCoercion', '1.1')) throw needs('insert pictures');
      await setSelected(base64, office.CoercionType.Image);
    },
    async insertText(text) {
      if (host === 'excel') {
        if (!hosts.excel || !supports('ExcelApi', '1.7')) throw needs('insert text');
        const rows = text
          .replace(/\r\n?/g, '\n')
          .split('\n')
          .map((line) => [line]);
        await hosts.excel.run(async (context) => {
          const range = context.workbook.getActiveCell().getResizedRange(rows.length - 1, 0);
          range.numberFormat = rows.map(() => ['@']);
          range.values = rows;
          await context.sync();
        });
        return;
      }
      if (!supports('TextCoercion', '1.1')) throw needs('insert text');
      try {
        await setSelected(text, office.CoercionType.Text);
      } catch (error) {
        if (host === 'powerpoint')
          throw new HostError(
            'Select a text box or a shape on the slide, then insert the text again.',
            'noTextBox',
          );
        throw error;
      }
    },
    async file() {
      const name = documentName(office, host);
      const file = await officeCall<Office.File>((done) =>
        office.context.document.getFileAsync(
          office.FileType.Compressed,
          { sliceSize: SLICE_BYTES },
          done as never,
        ),
      );
      try {
        const parts: Uint8Array<ArrayBuffer>[] = [];
        for (let i = 0; i < file.sliceCount; i++) {
          const slice = await officeCall<Office.Slice>((done) =>
            file.getSliceAsync(i, done as never),
          );
          parts.push(Uint8Array.from(slice.data as ArrayLike<number>));
        }
        return { blob: new Blob(parts), name };
      } finally {
        // Office allows only a couple of open files: always let go of it.
        file.closeAsync();
      }
    },
  };
}

/** The document's file name (its URL's last part), or a plain name for an unsaved one. */
export function documentName(office: OfficeApi, host: DocumentHost): string {
  const extension = DOCUMENT_FORMAT[host];
  const url = office.context.document.url ?? '';
  let last = url.split(/[\\/]/).pop() ?? '';
  try {
    last = decodeURIComponent(last.split('?')[0] ?? '');
  } catch {
    // Not URL-encoded: keep it as it is.
  }
  const base = last.replace(/\.[A-Za-z0-9]{1,5}$/, '').trim();
  const fallback = { word: 'Document', excel: 'Workbook', powerpoint: 'Presentation' }[host];
  return `${base || fallback}.${extension}`;
}
