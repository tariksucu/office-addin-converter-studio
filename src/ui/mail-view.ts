import { formatBytes, type Target } from '../catalog';
import { convert, fetchResult, type Result } from '../convert';
import {
  MAX_ATTACHMENT_BYTES,
  canReplyWithFile,
  type Mail,
  type MailAttachment,
} from '../office/mail';
import type { Context } from './app';
import { clear, h, uid } from './dom';

/**
 * Outlook: the message's file attachments, each with what Converter Studio can make of it (the
 * API's list for that format). While writing, the result replaces the original or is attached
 * beside it; in a received message it goes into a reply or is downloaded.
 */
export function mountMail(context: Context, main: HTMLElement, mail: Mail): void {
  const { doc, t, catalog, tasks } = context;
  const list = h(doc, 'ul', { className: 'attachments' });
  const refresh = h(doc, 'button', { type: 'button', onclick: () => void render() }, t.refresh);
  main.append(
    h(doc, 'div', { className: 'section-head' }, h(doc, 'h2', {}, t.attachments), refresh),
    list,
  );
  const buttons = new Set<HTMLButtonElement>();
  tasks.onBusy((busy) => {
    for (const button of buttons) button.disabled = busy;
    refresh.disabled = busy;
  });

  const render = async () => {
    let attachments: MailAttachment[];
    try {
      attachments = await mail.list();
    } catch (error) {
      tasks.error(error instanceof Error ? error.message : t.failed);
      return;
    }
    clear(list);
    buttons.clear();
    if (attachments.length === 0) {
      list.append(h(doc, 'li', { className: 'empty' }, t.noAttachments));
      return;
    }
    for (const attachment of attachments) list.append(row(attachment));
  };

  const row = (attachment: MailAttachment): HTMLElement => {
    const targets = catalog.attachmentTargets(catalog.formatOfName(attachment.name));
    const title = h(
      doc,
      'p',
      { className: 'file' },
      h(doc, 'span', { className: 'file-name' }, attachment.name),
      attachment.size > 0
        ? h(doc, 'span', { className: 'file-size' }, formatBytes(attachment.size))
        : null,
    );
    if (targets.length === 0)
      return h(doc, 'li', {}, title, h(doc, 'p', { className: 'hint' }, t.cannotConvert));
    const select = h(doc, 'select', { id: uid('to') });
    for (const target of targets)
      select.append(h(doc, 'option', { value: target.key }, target.label));
    const button = h(
      doc,
      'button',
      {
        type: 'button',
        className: 'primary',
        onclick: () => {
          const target = targets.find((x) => x.key === select.value) ?? targets[0]!;
          void start(attachment, target);
        },
      },
      t.convert,
    );
    buttons.add(button);
    return h(
      doc,
      'li',
      {},
      title,
      h(
        doc,
        'div',
        { className: 'row' },
        h(doc, 'label', { for: select.id }, t.convertTo(attachment.name)),
        select,
        button,
      ),
    );
  };

  const start = async (attachment: MailAttachment, target: Target) => {
    const result = await tasks.run(async (hooks) => {
      const file = await mail.content(attachment);
      return convert(
        context.api,
        { file, name: attachment.name, to: target.to, options: target.options },
        hooks,
      );
    });
    if (!result) return;
    context.refreshAccount();
    offer(attachment, result);
  };

  /** What to do with a finished conversion. */
  const offer = (attachment: MailAttachment, result: Result) => {
    const name = result.outputName;
    const tooLarge = result.size > MAX_ATTACHMENT_BYTES;
    const actions = tasks.done(`${name} (${formatBytes(result.size)}).`, t.credits(result.credits));
    const action = (
      label: string,
      primary: boolean,
      run: (signal: AbortSignal) => Promise<string>,
    ) => {
      const button = h(
        doc,
        'button',
        {
          type: 'button',
          className: primary ? 'primary' : '',
          onclick: () => {
            void tasks
              .run(({ signal }) => {
                tasks.say(t.fetching);
                return run(signal!);
              })
              .then((message) => {
                if (message === null) return;
                tasks.done(message);
                if (mail.mode === 'compose') void render();
              });
          },
        },
        label,
      );
      actions.append(button);
    };
    const file = (signal: AbortSignal) => fetchResult(context.api, result, signal);
    if (mail.mode === 'compose') {
      if (tooLarge) {
        tasks.error(t.tooLargeForOutlook);
        return;
      }
      action(t.replaceOriginal, true, async (signal) => {
        await mail.attach(await file(signal), name);
        await mail.remove(attachment.id);
        return t.replaced(name);
      });
      action(t.addAttachment, false, async (signal) => {
        await mail.attach(await file(signal), name);
        return t.attached(name);
      });
      return;
    }
    if (canReplyWithFile(context.office) && !tooLarge)
      action(t.replyWith, true, async (signal) => {
        await mail.reply(await file(signal), name);
        return t.replyOpened;
      });
    actions.append(
      h(
        doc,
        'button',
        { type: 'button', onclick: () => context.open(result.downloadUrl) },
        t.download,
      ),
      h(doc, 'p', { className: 'hint' }, t.expires),
    );
  };

  mail.onChange(() => {
    if (!tasks.busy) void render();
  });
  void render();
}
