/**
 * What the pane needs from its surroundings besides the document or message: opening a link
 * outside Office, the theme, the display language and the partitioned key for local storage.
 */

type OfficeApi = typeof Office;

/**
 * Opens a page or a download in the user's browser, from a click (never on its own). Desktop
 * Office has `openBrowserWindow` (OpenBrowserWindowApi 1.1); Office on the web opens a tab.
 */
export function openLink(office: OfficeApi, win: Window, url: string): void {
  if (
    office.context.requirements.isSetSupported('OpenBrowserWindowApi', '1.1') &&
    typeof office.context.ui?.openBrowserWindow === 'function'
  ) {
    office.context.ui.openBrowserWindow(url);
    return;
  }
  win.open(url, '_blank', 'noopener,noreferrer');
}

const COLOR = /^#[0-9a-f]{6}$/i;

/**
 * The pane takes the colours of the Office theme where Office tells it (`officeTheme`: Office on
 * Windows and Mac); elsewhere the stylesheet follows the system's light or dark mode.
 */
export function applyTheme(office: OfficeApi, root: HTMLElement): void {
  let theme: Office.OfficeTheme | undefined;
  try {
    theme = office.context.officeTheme;
  } catch {
    // Outlook on some platforms throws instead of returning nothing.
    theme = undefined;
  }
  if (!theme) return;
  const colors: [string, string | undefined][] = [
    ['--bg', theme.bodyBackgroundColor],
    ['--fg', theme.bodyForegroundColor],
    ['--control-bg', theme.controlBackgroundColor],
    ['--control-fg', theme.controlForegroundColor],
  ];
  if (!colors.every(([, value]) => value && COLOR.test(value))) return;
  for (const [name, value] of colors) root.style.setProperty(name, value!);
  root.dataset.theme = 'office';
}

/** The language Office shows its own interface in (`en-US`, `tr-TR`…). */
export function displayLanguage(office: OfficeApi): string {
  return office.context.displayLanguage || 'en-US';
}

/** Microsoft's partition key for local storage; empty where storage is not partitioned. */
export function partitionKey(office: OfficeApi): string {
  try {
    return office.context.partitionKey ?? '';
  } catch {
    return '';
  }
}
