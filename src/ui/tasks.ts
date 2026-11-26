import { ApiError, isAbort } from '../api';
import type { Hooks, Step } from '../convert';
import type { Messages } from '../strings';
import { clear, h } from './dom';

/**
 * The pane's one running job at a time: its progress line (a polite live region, so screen
 * readers hear it), its Cancel button, the price question above the automatic limit, and the
 * outcome: a confirmation, or the API's or Office's own error message.
 */

export type Work<T> = (hooks: Hooks) => Promise<T>;

export class Tasks {
  private controller: AbortController | null = null;
  readonly element: HTMLElement;
  private readonly line: HTMLElement;
  private readonly actions: HTMLElement;
  private readonly listeners = new Set<(busy: boolean) => void>();

  constructor(
    private readonly doc: Document,
    private readonly t: Messages,
  ) {
    this.line = h(doc, 'p', { className: 'status-line' });
    this.actions = h(doc, 'div', { className: 'status-actions' });
    this.element = h(
      doc,
      'div',
      { className: 'status', role: 'status', 'aria-live': 'polite' },
      this.line,
      this.actions,
    );
    this.element.hidden = true;
  }

  get busy(): boolean {
    return this.controller !== null;
  }

  /** Called with `true` when a job starts and `false` when it ends (buttons turn off and on). */
  onBusy(listener: (busy: boolean) => void): void {
    this.listeners.add(listener);
  }

  /** Runs one job; resolves its value, or null when it failed or was canceled (shown here). */
  async run<T>(work: Work<T>): Promise<T | null> {
    if (this.controller) return null;
    const controller = new AbortController();
    this.controller = controller;
    this.emit(true);
    this.show(this.t.converting, 'progress');
    try {
      return await work({
        signal: controller.signal,
        onStep: (step) => this.step(step),
        confirmPrice: (credits) => this.askPrice(credits, controller.signal),
      });
    } catch (error) {
      this.fail(error, controller.signal.aborted);
      return null;
    } finally {
      this.controller = null;
      this.emit(false);
    }
  }

  /** A line of progress of the job's own (fetching the result, inserting it). */
  say(text: string): void {
    this.show(text, 'progress');
  }

  /** The job's happy ending. */
  done(...lines: string[]): HTMLElement {
    this.show(lines.join(' '), 'done');
    return this.actions;
  }

  /** An ending short of the goal, with what to do next below it (not an error). */
  warning(...lines: string[]): HTMLElement {
    this.show(lines.join(' '), 'warning');
    return this.actions;
  }

  error(text: string): void {
    this.show(text, 'error');
  }

  private step(step: Step): void {
    if (step.step === 'uploading') {
      const percent = step.total > 0 ? Math.floor((step.sent / step.total) * 100) : 0;
      this.show(this.t.uploading(percent), 'progress');
    } else if (step.progress !== null && step.progress > 0 && step.progress < 1) {
      this.show(this.t.convertingPercent(Math.round(step.progress * 100)), 'progress');
    } else {
      this.show(this.t.converting, 'progress');
    }
  }

  private askPrice(credits: number, signal: AbortSignal): Promise<boolean> {
    return this.ask(this.t.price(credits), this.t.pay, this.t.cancel, signal);
  }

  /**
   * A question with two answers, `yes` first (a price before anything is charged); resolves
   * false when the job is canceled meanwhile.
   */
  ask(question: string, yes: string, no: string, signal?: AbortSignal): Promise<boolean> {
    return new Promise((resolve) => {
      this.show(question, 'question');
      const answer = (value: boolean) => {
        clear(this.actions);
        this.addCancel();
        resolve(value);
      };
      clear(this.actions);
      const first = h(
        this.doc,
        'button',
        { type: 'button', className: 'primary', onclick: () => answer(true) },
        yes,
      );
      this.actions.append(
        first,
        h(this.doc, 'button', { type: 'button', onclick: () => answer(false) }, no),
      );
      signal?.addEventListener('abort', () => resolve(false));
      first.focus();
    });
  }

  private show(text: string, kind: 'progress' | 'question' | 'done' | 'warning' | 'error'): void {
    this.element.hidden = false;
    this.element.dataset.kind = kind;
    // Errors interrupt; progress is read out when the reader is idle.
    this.element.setAttribute('role', kind === 'error' ? 'alert' : 'status');
    this.line.textContent = text;
    if (kind === 'progress') {
      if (!this.actions.querySelector('[data-cancel]')) {
        clear(this.actions);
        this.addCancel();
      }
    } else if (kind !== 'question') {
      clear(this.actions);
    }
  }

  private addCancel(): void {
    this.actions.append(
      h(
        this.doc,
        'button',
        { type: 'button', 'data-cancel': true, onclick: () => this.controller?.abort() },
        this.t.cancel,
      ),
    );
  }

  private fail(error: unknown, aborted: boolean): void {
    if (aborted || isAbort(error) || (error instanceof ApiError && error.code === 'canceled')) {
      this.show(this.t.canceled, 'error');
      return;
    }
    const message = error instanceof Error && error.message ? error.message : this.t.failed;
    this.show(message, 'error');
  }

  private emit(busy: boolean): void {
    for (const listener of this.listeners) listener(busy);
  }
}
