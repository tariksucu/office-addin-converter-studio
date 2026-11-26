import type { ConverterApi } from '../api';
import { formatBytes, type Catalog } from '../catalog';
import { FIT_TARGETS, isFitTarget } from '../fit/budget';
import { fitDraft, type FitReport } from '../fit/fit';
import type { PhotoEncoder } from '../fit/photo';
import type { Mail } from '../office/mail';
import type { FitTargetStore } from '../settings';
import type { Messages } from '../strings';
import { h, uid } from './dom';
import type { Tasks } from './tasks';

/**
 * "Fit attachments for sending" in a message being written: one total to choose (remembered),
 * one button, and a one-line outcome. Works without a key; with one, the servers' part asks for
 * its price first.
 */

export interface FitView {
  readonly doc: Document;
  readonly t: Messages;
  readonly tasks: Tasks;
  readonly catalog: Catalog;
  readonly mail: Mail;
  readonly photos: PhotoEncoder;
  readonly api: ConverterApi | null;
  readonly targets: FitTargetStore;
  /** Opens the account page's API keys (without a key). */
  readonly openKeys: () => void;
  /** After a change to the draft (the attachment list, the account's credits). */
  readonly onChanged?: () => void;
}

export function mountFit(view: FitView): HTMLElement {
  const { doc, t, tasks } = view;
  const select = h(doc, 'select', { id: uid('fit-target') });
  for (const target of FIT_TARGETS)
    select.append(h(doc, 'option', { value: String(target) }, t.fitTargets[target]));
  select.value = String(view.targets.get());
  select.addEventListener('change', () => {
    const value = Number(select.value);
    if (isFitTarget(value)) view.targets.set(value);
  });
  const button = h(
    doc,
    'button',
    {
      type: 'button',
      className: 'primary',
      onclick: () => {
        const target = Number(select.value);
        void tasks
          .run(({ signal }) =>
            fitDraft(
              {
                mail: view.mail,
                catalog: view.catalog,
                photos: view.photos,
                api: view.api,
                t,
                signal,
                say: (text) => tasks.say(text),
                ask: (credits, names) =>
                  tasks.ask(
                    t.fitPrice(credits, t.list(names)),
                    t.fitPay(credits),
                    t.cancel,
                    signal,
                  ),
              },
              isFitTarget(target) ? target : FIT_TARGETS[0],
            ),
          )
          .then((report) => {
            if (!report) return;
            show(view, report);
            if (report.changes.length > 0 || report.credits > 0) view.onChanged?.();
          });
      },
    },
    t.fitButton,
  );
  tasks.onBusy((busy) => {
    button.disabled = busy;
    select.disabled = busy;
  });
  return h(
    doc,
    'section',
    { className: 'fit' },
    h(doc, 'h2', {}, t.fitTitle),
    h(
      doc,
      'div',
      { className: 'row' },
      h(doc, 'label', { for: select.id }, t.fitTarget),
      select,
      button,
    ),
    h(doc, 'p', { className: 'hint' }, t.fitHint),
    h(doc, 'p', { className: 'hint' }, view.api ? t.fitHintServer : t.fitHintNoKey),
  );
}

/** The outcome: one line, then what is left to do (if anything) below it. */
function show(view: FitView, report: FitReport): void {
  const { doc, t, tasks } = view;
  if (report.kind === 'empty') {
    tasks.done(t.fitEmpty);
    return;
  }
  if (report.kind === 'alreadyFits') {
    tasks.done(t.fitAlready(report.target, formatBytes(report.before)));
    return;
  }
  const target = report.target * 1_000_000;
  const lines = [t.fitSummary(report.count, formatBytes(report.before), formatBytes(report.after))];
  if (report.kind === 'fits') lines.push(t.fitFits(report.after <= 10 * 1_000_000));
  else
    lines.push(
      t.fitStillOver(
        report.target,
        t.list(report.tooBig.map((x) => `${x.name} (${formatBytes(x.size)})`)),
      ),
    );
  if (report.credits > 0) lines.push(t.credits(report.credits));
  if (report.kept.length > 0) lines.push(t.fitKept(t.list(report.kept), report.kept.length));
  const actions = report.kind === 'fits' ? tasks.done(...lines) : tasks.warning(...lines);
  const note = (text: string) => actions.append(h(doc, 'p', { className: 'hint' }, text));
  if (report.kind === 'over' && report.tooBig.length > 0) note(t.fitOneDrive(report.tooBig.length));
  if (report.needKey.length > 0 && report.after > target) {
    note(t.fitNeedKey(t.list(report.needKey)));
    actions.append(
      h(doc, 'button', { type: 'button', className: 'link', onclick: view.openKeys }, t.fitKeyLink),
    );
  }
  if (report.skipped.length > 0) note(t.fitSkipped(t.list(report.skipped)));
  for (const failure of report.failed) note(t.fitFailed(failure.name, failure.message));
}
