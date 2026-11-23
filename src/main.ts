import { mountApp } from './ui/app';

/**
 * The task pane's entry point. Office.js (Microsoft's hosted copy, loaded by taskpane.html)
 * calls back once it knows the app; the API is on the pane's own origin
 * (`https://converter.studio/api` for the pane at `https://converter.studio/integrations/office/`),
 * so the address comes from where the pane was loaded and never from a setting.
 */

function storage(): Storage | null {
  try {
    return window.localStorage;
  } catch {
    // Site data blocked: the key is kept in memory while the pane is open.
    return null;
  }
}

const root = document.getElementById('app');
if (root) {
  if (typeof Office === 'undefined') {
    root.textContent =
      'Office.js could not be loaded. Check your connection and reopen the add-in.';
  } else {
    void Office.onReady((info) =>
      mountApp({
        office: Office,
        word: typeof Word === 'undefined' ? undefined : Word,
        excel: typeof Excel === 'undefined' ? undefined : Excel,
        host: info.host,
        win: window,
        root,
        apiBase: `${window.location.origin}/api`,
        storage: storage(),
      }),
    );
  }
}
