/**
 * The user's Converter Studio API key, kept where Microsoft says an add-in keeps private data:
 * the browser storage of the add-in's own origin, under Office's partition key
 * (`Office.context.partitionKey`, "Persist add-in state and settings"). It stays on this device
 * and in this Office app; "Disconnect" deletes it.
 *
 * Deliberately not kept in:
 * - Outlook's roaming settings: they roam with the mailbox, but Microsoft says they "shouldn't be
 *   considered secure storage" and "shouldn't be used to store sensitive information, such as
 *   user credentials or security tokens" (other services such as Microsoft Graph can read them);
 * - a document's settings (`Office.context.document.settings`): those are saved inside the file
 *   and travel with it to everyone it is sent to.
 */

const STORAGE_KEY = 'converter-studio.api-key';
const KEY_PATTERN = /^cvs_live_[A-Za-z0-9]{32}$/;

export interface KeyStore {
  get(): string | null;
  set(key: string): void;
  clear(): void;
}

/** A key as the account page shows it (`cvs_live_` and 32 letters and digits). */
export function isApiKey(value: string): boolean {
  return KEY_PATTERN.test(value);
}

/**
 * Storage that may be missing or refuse (blocked site data, private windows): the key then lives
 * only as long as the pane stays open.
 */
export function keyStore(storage: Storage | null, partition: string): KeyStore {
  const name = `${partition}${STORAGE_KEY}`;
  let memory: string | null = null;
  return {
    get() {
      try {
        const value = storage?.getItem(name) ?? null;
        return value && isApiKey(value) ? value : memory;
      } catch {
        return memory;
      }
    },
    set(key) {
      memory = key;
      try {
        storage?.setItem(name, key);
      } catch {
        // Kept in memory only.
      }
    },
    clear() {
      memory = null;
      try {
        storage?.removeItem(name);
      } catch {
        // Nothing stored.
      }
    },
  };
}
