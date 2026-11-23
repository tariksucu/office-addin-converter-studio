import { describe, expect, it } from 'vitest';
import {
  Catalog,
  pictureOptions,
  resultName,
  textOptions,
  formatBytes,
  PAGE_RANGE,
} from '../src/catalog';
import { FORMATS } from './fake-api';

const catalog = new Catalog(FORMATS);
const format = (id: string) => catalog.get(id)!;

describe('attachment targets (the API list for the format)', () => {
  it('offers what the server makes of a HEIC photo, JPG first', () => {
    const targets = catalog.attachmentTargets(catalog.formatOfName('IMG_0042.HEIC'));
    expect(targets[0]).toMatchObject({ to: 'jpg', label: 'JPG' });
    expect(targets.map((t) => t.to)).toEqual(format('heic').to);
  });

  it('names PDF to PDF as a smaller file, with the compression option', () => {
    const smaller = catalog.attachmentTargets(format('pdf')).find((t) => t.to === 'pdf');
    expect(smaller).toEqual({
      key: 'pdf',
      to: 'pdf',
      label: 'PDF (smaller file)',
      options: { compression: 'balanced' },
    });
  });

  it('names a picture to TXT or DOCX as text recognition', () => {
    const targets = catalog.attachmentTargets(format('jpg'));
    expect(targets.find((t) => t.to === 'txt')?.label).toBe('TXT (text from the picture)');
    expect(targets.find((t) => t.to === 'docx')?.label).toBe('DOCX (text from the picture)');
  });

  it('offers nothing for an unknown file', () => {
    expect(catalog.attachmentTargets(catalog.formatOfName('notes.xyz'))).toEqual([]);
    expect(catalog.attachmentTargets(catalog.formatOfName('no-extension'))).toEqual([]);
  });

  it('offers Office files as PDF', () => {
    for (const name of ['a.docx', 'b.xlsx', 'c.pptx', 'd.doc'])
      expect(catalog.attachmentTargets(catalog.formatOfName(name))[0]?.to).toBe('pdf');
  });
});

describe('pictures for Office', () => {
  it('turns photos into JPG and drawn or see-through pictures into PNG', () => {
    expect(catalog.pictureTarget(format('heic'))).toBe('jpg');
    expect(catalog.pictureTarget(format('dng'))).toBe('jpg');
    expect(catalog.pictureTarget(format('cr2'))).toBe('jpg');
    expect(catalog.pictureTarget(format('avif'))).toBe('png');
    expect(catalog.pictureTarget(format('webp'))).toBe('png');
    expect(catalog.pictureTarget(format('eps'))).toBe('png');
    expect(catalog.pictureTarget(format('ai'))).toBe('png');
    expect(catalog.pictureTarget(format('svg'))).toBe('png');
    expect(catalog.pictureTarget(format('pdf'))).toBe('png');
  });

  it('does not offer what Office inserts itself, or what is not a picture', () => {
    for (const id of ['jpg', 'png', 'gif', 'bmp', 'docx', 'mp4', 'xlsx'])
      expect(catalog.pictureTarget(format(id))).toBeNull();
    expect(catalog.pictureTarget(undefined)).toBeNull();
  });

  it('only offers a target the server lists', () => {
    const noPng = new Catalog([{ ...format('avif'), to: ['jpg'] }]);
    expect(noPng.pictureTarget(noPng.get('avif'))).toBe('jpg');
    const none = new Catalog([{ ...format('heic'), to: ['pdf'] }]);
    expect(none.pictureTarget(none.get('heic'))).toBeNull();
  });

  it('accepts exactly those extensions in the file picker', () => {
    const accept = catalog.accept('picture').split(',');
    expect(accept).toEqual(
      expect.arrayContaining(['.heic', '.heif', '.dng', '.avif', '.eps', '.pdf']),
    );
    expect(accept).not.toContain('.jpg');
    expect(accept).not.toContain('.docx');
  });

  it('asks for a PDF page at 150 dpi, a size for photos, twice the size for SVG', () => {
    expect(pictureOptions(format('pdf'), 'png', { longSide: '1920', page: 3 })).toEqual({
      pageRange: '3',
      dpi: '150',
    });
    expect(pictureOptions(format('heic'), 'jpg', { longSide: '1920', page: 1 })).toEqual({
      longSide: '1920',
      quality: 'high',
    });
    expect(pictureOptions(format('avif'), 'png', { longSide: '', page: 1 })).toEqual({});
    expect(pictureOptions(format('svg'), 'png', { longSide: '1920', page: 1 })).toEqual({
      svgScale: '2',
    });
  });
});

describe('text from scans', () => {
  it('reads scanned PDFs and pictures of pages', () => {
    for (const id of ['pdf', 'jpg', 'png', 'heic', 'tiff', 'webp', 'bmp'])
      expect(catalog.canReadText(format(id))).toBe(true);
    for (const id of ['docx', 'avif', 'mp3']) expect(catalog.canReadText(format(id))).toBe(false);
  });

  it('turns recognition on for every PDF page, or the pages asked for', () => {
    expect(textOptions(format('pdf'), { language: 'auto', pages: '' })).toEqual({
      ocrLanguage: 'auto',
      ocr: true,
    });
    expect(textOptions(format('pdf'), { language: 'de', pages: ' 1-3, 5 ' })).toEqual({
      ocrLanguage: 'de',
      ocr: true,
      pageRange: '1-3, 5',
    });
    expect(textOptions(format('jpg'), { language: 'tr', pages: '2' })).toEqual({
      ocrLanguage: 'tr',
    });
    expect(PAGE_RANGE.test('1-3, 5')).toBe(true);
    expect(PAGE_RANGE.test('one')).toBe(false);
  });
});

describe('export targets (what the app does not save itself)', () => {
  it('Word: the API targets other than DOCX, PDF and ODT, then PDF/A', () => {
    const targets = catalog.exportTargets('word');
    const plain = targets.filter((t) => t.key === t.to).map((t) => t.to);
    expect(plain).toEqual(format('docx').to.filter((to) => !['docx', 'pdf', 'odt'].includes(to)));
    expect(plain).toEqual(expect.arrayContaining(['epub', 'md', 'html', 'rtf', 'txt']));
    expect(targets.at(-1)).toMatchObject({ to: 'pdf', options: { pdfA: true } });
  });

  it('Excel: JSON and XLS, and PDFs with each sheet on one page or PDF/A', () => {
    const targets = catalog.exportTargets('excel');
    expect(targets.filter((t) => t.key === t.to).map((t) => t.to)).toEqual(['json', 'xls']);
    expect(targets.map((t) => t.options)).toContainEqual({ sheetsOnePage: true });
  });

  it('PowerPoint: PPT, and PDFs with notes, handouts or PDF/A', () => {
    const targets = catalog.exportTargets('powerpoint');
    expect(targets.map((t) => t.key)).toEqual(['ppt', 'pdf-notes', 'pdf-handout', 'pdfa']);
    expect(targets[2]!.options).toEqual({ slideLayout: 'handout', handoutSlides: '3' });
  });

  it('follows the server: no PDF variants when it cannot make PDFs', () => {
    const noPdf = new Catalog([{ ...format('docx'), to: ['epub'] }]);
    expect(noPdf.exportTargets('word').map((t) => t.key)).toEqual(['epub']);
    expect(new Catalog([]).exportTargets('excel')).toEqual([]);
  });
});

describe('names and sizes', () => {
  it('keeps the name with the new extension', () => {
    expect(resultName('Holiday photo.HEIC', 'jpg')).toBe('Holiday photo.jpg');
    expect(resultName('.heic', 'jpg')).toBe('converted.jpg');
  });

  it('counts decimal kilobytes and megabytes', () => {
    expect(formatBytes(999)).toBe('999 B');
    expect(formatBytes(25_400)).toBe('25 KB');
    expect(formatBytes(3_000_000)).toBe('3 MB');
    expect(formatBytes(3_450_000)).toBe('3.5 MB');
  });
});
