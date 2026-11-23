import {
  LONG_SIDES,
  OCR_LANGUAGES,
  PAGE_RANGE,
  formatBytes,
  pictureOptions,
  textOptions,
} from '../catalog';
import { convert, fetchResult, isSingleFile } from '../convert';
import { blobText, blobToBase64 } from '../office/async';
import type { OpenDocument } from '../office/document';
import type { Context } from './app';
import { h, uid } from './dom';

/**
 * Word, Excel and PowerPoint: a picture Office cannot open, inserted as PNG or JPG; the text of
 * a scan, inserted as text; and the document exported to a format its app does not save.
 */
export function mountDocument(context: Context, main: HTMLElement, document: OpenDocument): void {
  const { doc, t, catalog, tasks } = context;
  const buttons: HTMLButtonElement[] = [];
  tasks.onBusy((busy) => {
    for (const button of buttons) button.disabled = busy;
  });
  const field = (label: string, control: HTMLElement, hint?: string) => {
    const wrapper = h(
      doc,
      'div',
      { className: 'field' },
      h(doc, 'label', { for: control.id }, label),
      control,
    );
    if (hint) wrapper.append(h(doc, 'p', { className: 'hint' }, hint));
    return wrapper;
  };
  const button = (label: string, onClick: () => void) => {
    const element = h(
      doc,
      'button',
      { type: 'button', className: 'primary', onclick: onClick },
      label,
    );
    buttons.push(element);
    return element;
  };
  const chosen = (input: HTMLInputElement) => input.files?.[0] ?? null;

  // A picture.
  const pictureInput = h(doc, 'input', {
    id: uid('picture'),
    type: 'file',
    accept: catalog.accept('picture'),
  });
  const pageInput = h(doc, 'input', {
    id: uid('page'),
    type: 'number',
    min: 1,
    value: 1,
    inputmode: 'numeric',
  });
  const pageField = field(t.pdfPage, pageInput);
  const sizeSelect = h(doc, 'select', { id: uid('size') });
  for (const side of LONG_SIDES)
    sizeSelect.append(
      h(doc, 'option', { value: side }, side ? t.sizePixels(side) : t.sizeOriginal),
    );
  const sizeField = field(t.size, sizeSelect);
  pageField.hidden = true;
  pictureInput.addEventListener('change', () => {
    const format = catalog.formatOfName(chosen(pictureInput)?.name ?? '');
    pageField.hidden = format?.id !== 'pdf';
    sizeField.hidden = format?.id === 'pdf' || format?.category === 'vector';
  });

  const insertPicture = async () => {
    const file = chosen(pictureInput);
    if (!file) return tasks.error(t.chooseFile);
    const format = catalog.formatOfName(file.name);
    const to = catalog.pictureTarget(format);
    if (!format || !to) return tasks.error(t.notPicture);
    const page = Number(pageInput.value);
    if (format.id === 'pdf' && !(Number.isInteger(page) && page >= 1 && page <= 99_999))
      return tasks.error(t.badPage);
    const result = await tasks.run(async (hooks) => {
      const options = pictureOptions(format, to, { longSide: sizeSelect.value, page });
      const done = await convert(context.api, { file, name: file.name, to, options }, hooks);
      tasks.say(t.fetching);
      if (!isSingleFile(done, to)) throw new Error(t.notOneFile);
      const picture = await fetchResult(context.api, done, hooks.signal);
      await document.insertPicture(await blobToBase64(picture));
      return done;
    });
    if (!result) return;
    tasks.done(t.pictureInserted, t.credits(result.credits));
    context.refreshAccount();
  };

  // The text of a scan.
  const textInput = h(doc, 'input', {
    id: uid('scan'),
    type: 'file',
    accept: catalog.accept('text'),
  });
  const languageSelect = h(doc, 'select', { id: uid('language') });
  for (const [code, name] of OCR_LANGUAGES)
    languageSelect.append(h(doc, 'option', { value: code }, name));
  const pagesInput = h(doc, 'input', {
    id: uid('pages'),
    type: 'text',
    placeholder: t.pagesAll,
    inputmode: 'numeric',
  });
  const pagesField = field(t.pages, pagesInput);
  pagesField.hidden = true;
  textInput.addEventListener('change', () => {
    pagesField.hidden = catalog.formatOfName(chosen(textInput)?.name ?? '')?.id !== 'pdf';
  });

  const insertText = async () => {
    const file = chosen(textInput);
    if (!file) return tasks.error(t.chooseFile);
    const format = catalog.formatOfName(file.name);
    if (!format || !catalog.canReadText(format)) return tasks.error(t.notScan);
    const pages = pagesInput.value;
    if (format.id === 'pdf' && pages.trim() && !PAGE_RANGE.test(pages))
      return tasks.error(t.badPages);
    const result = await tasks.run(async (hooks) => {
      const options = textOptions(format, { language: languageSelect.value, pages });
      const done = await convert(context.api, { file, name: file.name, to: 'txt', options }, hooks);
      tasks.say(t.fetching);
      if (!isSingleFile(done, 'txt')) throw new Error(t.notOneFile);
      const text = (await blobText(await fetchResult(context.api, done, hooks.signal))).trim();
      if (text) await document.insertText(text);
      return { done, empty: !text };
    });
    if (!result) return;
    tasks.done(result.empty ? t.noText : t.textInserted, t.credits(result.done.credits));
    context.refreshAccount();
  };

  // The document itself.
  const targets = catalog.exportTargets(document.host);
  const exportSelect = h(doc, 'select', { id: uid('export') });
  for (const target of targets)
    exportSelect.append(h(doc, 'option', { value: target.key }, target.label));

  const exportDocument = async () => {
    const target = targets.find((x) => x.key === exportSelect.value) ?? targets[0];
    if (!target) return;
    const result = await tasks.run(async (hooks) => {
      const { blob, name } = await document.file();
      return convert(
        context.api,
        { file: blob, name, to: target.to, options: target.options },
        hooks,
      );
    });
    if (!result) return;
    const actions = tasks.done(
      t.exportReady(`${result.outputName} (${formatBytes(result.size)})`),
      t.credits(result.credits),
    );
    actions.append(
      h(
        doc,
        'button',
        { type: 'button', className: 'primary', onclick: () => context.open(result.downloadUrl) },
        t.download,
      ),
      h(doc, 'p', { className: 'hint' }, t.expires),
    );
    context.refreshAccount();
  };

  main.append(
    h(
      doc,
      'section',
      { 'aria-labelledby': 'picture-title' },
      h(doc, 'h2', { id: 'picture-title' }, t.pictureTitle),
      h(doc, 'p', { className: 'hint' }, t.pictureHint),
      field(t.pictureFile, pictureInput),
      pageField,
      sizeField,
      button(t.insertPicture, () => void insertPicture()),
    ),
    h(
      doc,
      'section',
      { 'aria-labelledby': 'text-title' },
      h(doc, 'h2', { id: 'text-title' }, t.textTitle),
      h(doc, 'p', { className: 'hint' }, t.textHint[document.host]),
      field(t.textFile, textInput),
      field(t.language, languageSelect),
      pagesField,
      button(t.insertText, () => void insertText()),
    ),
    h(
      doc,
      'section',
      { 'aria-labelledby': 'export-title' },
      h(doc, 'h2', { id: 'export-title' }, t.exportTitle),
      h(doc, 'p', { className: 'hint' }, t.exportHint),
      ...(targets.length > 0
        ? [field(t.exportTo, exportSelect), button(t.exportButton, () => void exportDocument())]
        : [h(doc, 'p', { className: 'hint' }, t.exportNone)]),
    ),
  );
}
