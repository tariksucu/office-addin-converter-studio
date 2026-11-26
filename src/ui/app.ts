import { ApiError, ConverterApi, type Account, type Fetch } from '../api';
import { Catalog } from '../catalog';
import { loadLibheif } from '../fit/heic';
import { browserPhotos, type PhotoEncoder } from '../fit/photo';
import { isApiKey, keyStore } from '../key-store';
import { documentHost, openDocument } from '../office/document';
import { applyTheme, displayLanguage, openLink, partitionKey } from '../office/environment';
import { openMail } from '../office/mail';
import { fitTargetStore, type FitTargetStore } from '../settings';
import { messages, type Messages } from '../strings';
import { mountDocument } from './document-view';
import { clear, h, uid } from './dom';
import { mountFit } from './fit-view';
import { mountMail } from './mail-view';
import { Tasks } from './tasks';

/**
 * The task pane: the first run (what the add-in does, then the API key), the account line, and
 * the view for the app it runs in (Outlook's message, or the Word, Excel or PowerPoint document).
 */

export interface Environment {
  readonly office: typeof Office;
  readonly word?: typeof Word;
  readonly excel?: typeof Excel;
  /** `Office.onReady`'s host: Outlook, Word, Excel, PowerPoint, or null outside Office. */
  readonly host: unknown;
  readonly win: Window;
  readonly root: HTMLElement;
  /** The API's root, on the pane's own origin: `https://converter.studio/api`. */
  readonly apiBase: string;
  readonly fetch?: Fetch;
  readonly storage: Storage | null;
  /** Photos made smaller on the computer (tests stand in for the browser's canvas). */
  readonly photos?: PhotoEncoder;
}

/** What every view gets. */
export interface Context {
  readonly office: typeof Office;
  readonly doc: Document;
  readonly t: Messages;
  readonly api: ConverterApi;
  readonly catalog: Catalog;
  readonly tasks: Tasks;
  readonly open: (url: string) => void;
  /** Reads the account again (its credits after a conversion). */
  readonly refreshAccount: () => void;
  readonly photos: PhotoEncoder;
  /** The total "Fit attachments for sending" aims at, remembered. */
  readonly fitTargets: FitTargetStore;
  /** Opens the account page's API keys. */
  readonly openKeys: () => void;
}

export function siteLinks(apiBase: string) {
  const site = new URL(apiBase).origin;
  return {
    site,
    signUp: `${site}/sign-up`,
    keys: `${site}/account#api-keys`,
    privacy: `${site}/privacy`,
    support: `${site}/support`,
    pricing: `${site}/pricing`,
  };
}

export async function mountApp(env: Environment): Promise<void> {
  const { office, root } = env;
  const doc = root.ownerDocument;
  const t = messages(displayLanguage(office));
  const links = siteLinks(env.apiBase);
  const open = (url: string) => openLink(office, env.win, url);
  if (env.host) applyTheme(office, doc.documentElement);

  const kind = env.host === office.HostType.Outlook ? 'mail' : documentHost(office, env.host);
  if (!env.host || !kind) {
    show(
      root,
      h(doc, 'p', { className: 'notice' }, env.host ? t.unsupportedHost : t.outsideOffice),
    );
    return;
  }

  const store = keyStore(env.storage, partitionKey(office));
  const photos = env.photos ?? browserPhotos(doc, () => loadLibheif(doc));
  const fitTargets = fitTargetStore(env.storage, partitionKey(office));
  const openKeys = () => open(links.keys);
  const api = (key: string) => new ConverterApi(env.apiBase, key, env.fetch);
  show(root, h(doc, 'p', { className: 'notice' }, t.loading));

  let catalog: Catalog;
  try {
    catalog = new Catalog(await api('').formats());
  } catch {
    show(
      root,
      h(doc, 'p', { className: 'notice', role: 'alert' }, t.formatsFailed),
      h(doc, 'button', { type: 'button', onclick: () => void mountApp(env) }, t.tryAgain),
    );
    return;
  }

  const connected = (key: string, account: Account) => {
    const tasks = new Tasks(doc, t);
    const accountLine = h(
      doc,
      'span',
      { className: 'account-line' },
      t.account(account.email, account.credits),
    );
    const client = api(key);
    const context: Context = {
      office,
      doc,
      t,
      api: client,
      catalog,
      tasks,
      open,
      refreshAccount: () => {
        client.account().then(
          (next) => (accountLine.textContent = t.account(next.email, next.credits)),
          () => undefined,
        );
      },
      photos,
      fitTargets,
      openKeys,
    };
    const disconnect = h(
      doc,
      'button',
      {
        type: 'button',
        className: 'link',
        onclick: () => {
          store.clear();
          firstRun(null);
        },
      },
      t.disconnect,
    );
    tasks.onBusy((busy) => (disconnect.disabled = busy));
    const main = h(doc, 'main', { className: 'view' });
    show(
      root,
      h(
        doc,
        'header',
        { className: 'account' },
        accountLine,
        h(
          doc,
          'span',
          { className: 'account-actions' },
          linkButton(doc, t.pricing, () => open(links.pricing)),
          disconnect,
        ),
      ),
      main,
      tasks.element,
      footer(doc, t, links, open),
    );
    if (kind === 'mail') {
      const mail = openMail(office);
      if (!mail) main.append(h(doc, 'p', { className: 'notice' }, t.noMessage));
      else mountMail(context, main, mail);
    } else {
      mountDocument(
        context,
        main,
        openDocument({ office, word: env.word, excel: env.excel }, kind),
      );
    }
  };

  const firstRun = (problem: string | null) => {
    const input = h(doc, 'input', {
      id: uid('key'),
      type: 'password',
      autocomplete: 'off',
      spellcheck: 'false',
      'aria-describedby': 'key-hint',
      placeholder: 'cvs_live_…',
    });
    const error = h(doc, 'p', { className: 'error', role: 'alert' }, problem ?? '');
    error.hidden = !problem;
    const button = h(doc, 'button', { type: 'submit', className: 'primary' }, t.connect);
    // A message being written: its attachments can be fitted for sending before any key.
    const mail = kind === 'mail' ? openMail(office) : null;
    const fit: Node[] = [];
    if (mail?.mode === 'compose') {
      const tasks = new Tasks(doc, t);
      tasks.onBusy((busy) => (button.disabled = busy));
      fit.push(
        mountFit({
          doc,
          t,
          tasks,
          catalog,
          mail,
          photos,
          api: null,
          targets: fitTargets,
          openKeys,
        }),
        tasks.element,
      );
    }
    const form = h(
      doc,
      'form',
      {
        className: 'connect',
        onsubmit: (event: Event) => {
          event.preventDefault();
          const key = input.value.trim();
          if (!isApiKey(key)) {
            error.textContent = t.badKey;
            error.hidden = false;
            input.focus();
            return;
          }
          button.disabled = true;
          error.hidden = true;
          button.textContent = t.connecting;
          api(key)
            .account()
            .then(
              (account) => {
                store.set(key);
                connected(key, account);
              },
              (reason: unknown) => {
                button.disabled = false;
                button.textContent = t.connect;
                error.textContent = reason instanceof Error ? reason.message : t.failed;
                error.hidden = false;
                input.focus();
              },
            );
        },
      },
      h(doc, 'label', { for: input.id }, t.keyLabel),
      input,
      h(doc, 'p', { id: 'key-hint', className: 'hint' }, t.keyHint),
      error,
      button,
    );
    show(
      root,
      h(
        doc,
        'main',
        { className: 'view first-run' },
        h(doc, 'h1', {}, t.appName),
        ...fit,
        h(doc, 'p', {}, kind === 'mail' ? t.introMail : t.introDocument),
        h(doc, 'p', {}, t.introAccount),
        h(
          doc,
          'ol',
          {},
          h(
            doc,
            'li',
            {},
            linkButton(doc, t.stepSignUp, () => open(links.signUp)),
            t.stepSignUpTail,
          ),
          h(
            doc,
            'li',
            {},
            linkButton(doc, t.stepKey, () => open(links.keys)),
            t.stepKeyTail,
          ),
        ),
        form,
      ),
      footer(doc, t, links, open),
    );
  };

  const key = store.get();
  if (!key) {
    firstRun(null);
    return;
  }
  try {
    connected(key, await api(key).account());
  } catch (error) {
    // A key that no longer works (deleted on the account page) is forgotten here too.
    if (error instanceof ApiError && error.status === 401) store.clear();
    firstRun(error instanceof Error ? error.message : t.failed);
  }
}

function show(root: HTMLElement, ...children: Node[]): void {
  clear(root);
  root.append(...children);
}

/** A link that opens outside Office (a button, so it works the same in every Office). */
function linkButton(doc: Document, text: string, onClick: () => void): HTMLButtonElement {
  return h(doc, 'button', { type: 'button', className: 'link', onclick: onClick }, text);
}

function footer(
  doc: Document,
  t: Messages,
  links: ReturnType<typeof siteLinks>,
  open: (url: string) => void,
): HTMLElement {
  return h(
    doc,
    'footer',
    {},
    h(doc, 'span', {}, t.privacyNote, ' '),
    linkButton(doc, t.privacyLink, () => open(links.privacy)),
    ' · ',
    linkButton(doc, t.help, () => open(links.support)),
    ' · ',
    linkButton(doc, 'converter.studio', () => open(links.site)),
  );
}
