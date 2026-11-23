import type { ConversionOptions, FormatInfo } from './api';

/**
 * What each file can become, from the API's own format list (`GET /v1/formats`): a target is
 * offered only when the server lists it for that format now. What is decided here is only what
 * Office itself already does, which no API can say: the pictures every Office app inserts on its
 * own and the formats every edition of an app saves on its own (on the web too).
 */

/** Pictures Word, Excel and PowerPoint insert themselves on every platform: never offered. */
export const NATIVE_PICTURES: ReadonlySet<string> = new Set(['jpg', 'png', 'gif', 'bmp']);

/** Pictures that may be drawn or see-through: they become PNG, which keeps that. */
const KEEPS_PNG: ReadonlySet<string> = new Set(['pdf', 'avif', 'webp', 'ico', 'tiff', 'svg']);

export type DocumentHost = 'word' | 'excel' | 'powerpoint';

/** The file Office gives for the open document (`getFileAsync(Compressed)`). */
export const DOCUMENT_FORMAT: Readonly<Record<DocumentHost, string>> = {
  word: 'docx',
  excel: 'xlsx',
  powerpoint: 'pptx',
};

/**
 * What every edition of the app saves or downloads as itself, Office on the web included (Word:
 * a copy, PDF, ODT; Excel: a copy, PDF, ODS, CSV of the sheet; PowerPoint: a copy, PDF, ODP).
 * The export offers the API's other targets, and the PDF layouts the web editions lack.
 */
const HOST_SAVES: Readonly<Record<DocumentHost, ReadonlySet<string>>> = {
  word: new Set(['docx', 'pdf', 'odt']),
  excel: new Set(['xlsx', 'pdf', 'ods', 'csv']),
  powerpoint: new Set(['pptx', 'pdf', 'odp']),
};

/** A result to offer: the API's target and the options that make it what the label says. */
export interface Target {
  /** Unique within one list (a target id, or a PDF variant). */
  readonly key: string;
  readonly to: string;
  readonly label: string;
  readonly options: ConversionOptions;
}

/** OCR targets of a picture or a scan: plain text or a Word file of the recognised text. */
const RECOGNISED = new Set(['txt', 'docx']);

export class Catalog {
  private readonly byId = new Map<string, FormatInfo>();
  private readonly byExtension = new Map<string, FormatInfo>();

  constructor(formats: readonly FormatInfo[]) {
    for (const format of formats) {
      this.byId.set(format.id, format);
      for (const extension of format.extensions)
        this.byExtension.set(extension.toLowerCase(), format);
    }
  }

  get(id: string): FormatInfo | undefined {
    return this.byId.get(id);
  }

  name(id: string): string {
    return this.byId.get(id)?.name ?? id.toUpperCase();
  }

  /** A file's format, by its name's extension. */
  formatOfName(name: string): FormatInfo | undefined {
    const match = /\.([A-Za-z0-9]+)$/.exec(name.trim());
    return match ? this.byExtension.get(match[1]!.toLowerCase()) : undefined;
  }

  /**
   * Mail attachments: everything the server makes of the file, most used first. PDF to PDF is
   * the server's "make it smaller"; pictures to TXT or DOCX are text recognition.
   */
  attachmentTargets(format: FormatInfo | undefined): Target[] {
    if (!format) return [];
    return format.to.map((to): Target => {
      if (to === format.id)
        return {
          key: to,
          to,
          label: `${this.name(to)} (smaller file)`,
          options: to === 'pdf' ? { compression: 'balanced' } : {},
        };
      if (format.category === 'image' && RECOGNISED.has(to))
        return { key: to, to, label: `${this.name(to)} (text from the picture)`, options: {} };
      return { key: to, to, label: this.name(to), options: {} };
    });
  }

  /** The picture to insert for a file Office cannot insert itself, or null. */
  pictureTarget(format: FormatInfo | undefined): 'png' | 'jpg' | null {
    if (!format || NATIVE_PICTURES.has(format.id)) return null;
    if (format.category !== 'image' && format.category !== 'vector' && format.id !== 'pdf')
      return null;
    const first =
      KEEPS_PNG.has(format.id) || format.category === 'vector' ? (['png', 'jpg'] as const) : null;
    for (const to of first ?? (['jpg', 'png'] as const)) if (format.to.includes(to)) return to;
    return null;
  }

  /** A scan's text: from a picture of a page, or a PDF read with text recognition. */
  canReadText(format: FormatInfo | undefined): boolean {
    if (!format || !format.to.includes('txt')) return false;
    return format.category === 'image' || format.id === 'pdf';
  }

  /** The file picker's `accept` for pictures or scans: every extension the server can do. */
  accept(kind: 'picture' | 'text'): string {
    const extensions: string[] = [];
    for (const format of this.byId.values()) {
      const ok = kind === 'picture' ? this.pictureTarget(format) : this.canReadText(format);
      if (ok) extensions.push(...format.extensions.map((e) => `.${e}`));
    }
    return extensions.join(',');
  }

  /** What the open document can be exported as that its app does not save itself. */
  exportTargets(host: DocumentHost): Target[] {
    const format = this.byId.get(DOCUMENT_FORMAT[host]);
    if (!format) return [];
    const saves = HOST_SAVES[host];
    const targets: Target[] = format.to
      .filter((to) => !saves.has(to))
      .map((to) => ({ key: to, to, label: this.name(to), options: {} }));
    if (format.to.includes('pdf')) targets.push(...PDF_VARIANTS[host]);
    return targets;
  }
}

/** PDFs the web editions cannot make: archive PDF/A, and PowerPoint's notes and handouts. */
const PDF_VARIANTS: Readonly<Record<DocumentHost, readonly Target[]>> = {
  word: [{ key: 'pdfa', to: 'pdf', label: 'PDF/A (for archiving)', options: { pdfA: true } }],
  excel: [
    {
      key: 'pdf-fit',
      to: 'pdf',
      label: 'PDF, each sheet on one page',
      options: { sheetsOnePage: true },
    },
    { key: 'pdfa', to: 'pdf', label: 'PDF/A (for archiving)', options: { pdfA: true } },
  ],
  powerpoint: [
    {
      key: 'pdf-notes',
      to: 'pdf',
      label: 'PDF with speaker notes',
      options: { slideLayout: 'notes' },
    },
    {
      key: 'pdf-handout',
      to: 'pdf',
      label: 'PDF handout, 3 slides a page',
      options: { slideLayout: 'handout', handoutSlides: '3' },
    },
    { key: 'pdfa', to: 'pdf', label: 'PDF/A (for archiving)', options: { pdfA: true } },
  ],
};

/** "1-3, 5": pages and ranges of pages, as the API takes them. */
export const PAGE_RANGE = /^\s*\d{1,5}(\s*-\s*\d{1,5})?(\s*,\s*\d{1,5}(\s*-\s*\d{1,5})?){0,49}\s*$/;

/** Text recognition's languages (the API's `ocrLanguage`): found on the page, or one of these. */
export const OCR_LANGUAGES: readonly (readonly [string, string])[] = [
  ['auto', 'Find the language'],
  ['en', 'English'],
  ['ar', 'Arabic'],
  ['zh', 'Chinese'],
  ['cs', 'Czech'],
  ['nl', 'Dutch'],
  ['fr', 'French'],
  ['de', 'German'],
  ['el', 'Greek'],
  ['he', 'Hebrew'],
  ['hi', 'Hindi'],
  ['it', 'Italian'],
  ['ja', 'Japanese'],
  ['ko', 'Korean'],
  ['pl', 'Polish'],
  ['pt', 'Portuguese'],
  ['ru', 'Russian'],
  ['es', 'Spanish'],
  ['tr', 'Turkish'],
  ['vi', 'Vietnamese'],
];

/** Picture sizes (the longer side, in pixels); '' keeps the size. */
export const LONG_SIDES = ['1920', '2560', '1280', ''] as const;

/** The options of a picture conversion for Office. */
export function pictureOptions(
  format: FormatInfo,
  to: 'png' | 'jpg',
  settings: { readonly longSide: string; readonly page: number },
): ConversionOptions {
  if (format.id === 'pdf') return { pageRange: String(settings.page), dpi: '150' };
  if (format.category === 'vector') return format.id === 'svg' ? { svgScale: '2' } : {};
  const options: Record<string, string> = {};
  if (settings.longSide) options.longSide = settings.longSide;
  if (to === 'jpg') options.quality = 'high';
  return options;
}

/** The options of a text recognition: the language, and for a PDF every page read (or some). */
export function textOptions(
  format: FormatInfo,
  settings: { readonly language: string; readonly pages: string },
): ConversionOptions {
  const options: Record<string, string | boolean> = { ocrLanguage: settings.language };
  if (format.id === 'pdf') {
    options.ocr = true;
    if (settings.pages.trim()) options.pageRange = settings.pages.trim();
  }
  return options;
}

/** The result's name: the original's, with the new extension. */
export function resultName(original: string, extension: string): string {
  const base = original.trim().replace(/\.[A-Za-z0-9]{1,5}$/, '') || 'converted';
  return `${base}.${extension}`;
}

export function formatBytes(bytes: number): string {
  if (bytes < 1000) return `${bytes} B`;
  if (bytes < 1000 * 1000) return `${Math.round(bytes / 1000)} KB`;
  return `${(bytes / (1000 * 1000)).toFixed(1).replace(/\.0$/, '')} MB`;
}
