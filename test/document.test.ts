import { describe, expect, it } from 'vitest';
import { blobBytes } from '../src/office/async';
import { documentHost, documentName, openDocument } from '../src/office/document';
import { fakeOffice, type HostName } from './fake-office';

const PNG = 'iVBORw0KGgo=';

function open(host: HostName, extra: Parameters<typeof fakeOffice>[0] = { host }) {
  const fake = fakeOffice({ ...extra, host });
  const kind = documentHost(fake.Office, host)!;
  return {
    ...fake,
    doc: openDocument({ office: fake.Office, word: fake.Word, excel: fake.Excel }, kind),
  };
}

describe('pictures', () => {
  it('Word: an inline picture at the selection, applied on sync', async () => {
    const { doc, state } = open('Word');
    await doc.insertPicture(PNG);
    expect(state.pictures).toEqual([{ base64: PNG, where: 'selection:Replace' }]);
    expect(state.calls).toEqual([]);
  });

  it('Word without WordApi 1.2 (Word 2016): image coercion', async () => {
    const { doc, state } = open('Word', {
      host: 'Word',
      sets: { WordApi: '1.1', ImageCoercion: '1.1' },
    });
    await doc.insertPicture(PNG);
    expect(state.calls).toEqual(['setSelectedDataAsync image']);
    expect(state.pictures).toHaveLength(1);
  });

  it('Excel: a picture shape at the active cell, at most 480 points wide', async () => {
    const { doc, state } = open('Excel');
    await doc.insertPicture(PNG);
    expect(state.pictures).toEqual([{ base64: PNG, where: 'sheet' }]);
    expect(state.shapes).toEqual([{ left: 96, top: 40, width: 480, lockAspectRatio: true }]);
    // Excel on the web has no image coercion: it was never used.
    expect(state.calls).toEqual([]);
  });

  it('Excel before ExcelApi 1.9: asks for a newer Office', async () => {
    const { doc } = open('Excel', { host: 'Excel', sets: { ExcelApi: '1.8' } });
    await expect(doc.insertPicture(PNG)).rejects.toMatchObject({ code: 'old' });
  });

  it('PowerPoint: image coercion on the slide', async () => {
    const { doc, state } = open('PowerPoint');
    await doc.insertPicture(PNG);
    expect(state.pictures).toEqual([{ base64: PNG, where: 'slide' }]);
  });
});

describe('text', () => {
  it('Word: text coercion at the cursor', async () => {
    const { doc, state } = open('Word');
    await doc.insertText('Line one\nLine two');
    expect(state.texts).toEqual([{ text: 'Line one\nLine two', where: 'selection' }]);
  });

  it('Excel: one line a cell from the active cell down, as text so nothing becomes a formula', async () => {
    const { doc, state } = open('Excel');
    await doc.insertText('Total\r\n=SUM(A1:A9)\n42');
    expect(state.cells).toEqual([
      {
        address: 'B2:B4',
        values: [['Total'], ['=SUM(A1:A9)'], ['42']],
        numberFormat: [['@'], ['@'], ['@']],
      },
    ]);
  });

  it('PowerPoint without a text box selected: says what to select', async () => {
    const { doc } = open('PowerPoint', { host: 'PowerPoint', noTextSelection: true });
    await expect(doc.insertText('Hello')).rejects.toMatchObject({ code: 'noTextBox' });
  });
});

describe('the document file', () => {
  it('reads every slice, joins them and closes the file', async () => {
    const bytes = new Uint8Array(9 * 1024 * 1024).map((_, i) => i % 251);
    const { doc, state } = open('Word', {
      host: 'Word',
      document: {
        url: 'https://contoso.sharepoint.com/Shared%20Documents/Quarterly%20report.docx',
        bytes,
      },
    });
    const { blob, name } = await doc.file();
    expect(name).toBe('Quarterly report.docx');
    expect(await blobBytes(blob)).toEqual(bytes);
    expect(state.closedFiles).toBe(1);
    expect(state.calls).toEqual(['getFileAsync compressed']);
  });

  it('names an unsaved document after its app', () => {
    for (const [host, name] of [
      ['Word', 'Document.docx'],
      ['Excel', 'Workbook.xlsx'],
      ['PowerPoint', 'Presentation.pptx'],
    ] as const) {
      const { Office } = fakeOffice({ host, document: { url: '', bytes: new Uint8Array() } });
      expect(documentName(Office, documentHost(Office, host)!)).toBe(name);
    }
    const { Office } = fakeOffice({
      host: 'Excel',
      document: { url: 'C:\\Users\\me\\Budget 2027.xlsx', bytes: new Uint8Array() },
    });
    expect(documentName(Office, 'excel')).toBe('Budget 2027.xlsx');
  });

  it('knows only Word, Excel and PowerPoint as document hosts', () => {
    const { Office } = fakeOffice({ host: 'Word' });
    expect(documentHost(Office, 'Outlook')).toBeNull();
    expect(documentHost(Office, null)).toBeNull();
  });
});
