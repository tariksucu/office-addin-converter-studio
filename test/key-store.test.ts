import { describe, expect, it } from 'vitest';
import { isApiKey, keyStore } from '../src/key-store';

const KEY = `cvs_live_${'A1'.repeat(16)}`;

function memoryStorage(): Storage & { data: Map<string, string> } {
  const data = new Map<string, string>();
  return {
    data,
    get length() {
      return data.size;
    },
    clear: () => data.clear(),
    getItem: (k) => data.get(k) ?? null,
    key: (i) => [...data.keys()][i] ?? null,
    removeItem: (k) => void data.delete(k),
    setItem: (k, v) => void data.set(k, v),
  };
}

describe('the API key', () => {
  it('is stored under Office’s partition key', () => {
    const storage = memoryStorage();
    const store = keyStore(storage, 'partition-1:');
    store.set(KEY);
    expect([...storage.data.keys()]).toEqual(['partition-1:converter-studio.api-key']);
    expect(keyStore(storage, 'partition-1:').get()).toBe(KEY);
    // Another partition (another app or top-level site) does not see it.
    expect(keyStore(storage, 'partition-2:').get()).toBeNull();
  });

  it('is unprefixed where Office has no partitions (desktop webviews)', () => {
    const storage = memoryStorage();
    keyStore(storage, '').set(KEY);
    expect(storage.getItem('converter-studio.api-key')).toBe(KEY);
  });

  it('ignores anything that is not a key, and forgets on disconnect', () => {
    const storage = memoryStorage();
    storage.setItem('converter-studio.api-key', 'not a key');
    const store = keyStore(storage, '');
    expect(store.get()).toBeNull();
    store.set(KEY);
    store.clear();
    expect(store.get()).toBeNull();
    expect(storage.data.size).toBe(0);
  });

  it('lives in memory while the pane is open when storage refuses', () => {
    const refusing = {
      getItem: () => {
        throw new Error('SecurityError');
      },
      setItem: () => {
        throw new Error('QuotaExceededError');
      },
      removeItem: () => {
        throw new Error('SecurityError');
      },
    } as unknown as Storage;
    const store = keyStore(refusing, '');
    store.set(KEY);
    expect(store.get()).toBe(KEY);
    store.clear();
    expect(store.get()).toBeNull();
    expect(keyStore(null, '').get()).toBeNull();
  });

  it('recognises the account page’s keys only', () => {
    expect(isApiKey(KEY)).toBe(true);
    expect(isApiKey(`cvs_live_${'a'.repeat(31)}`)).toBe(false);
    expect(isApiKey(` ${KEY}`)).toBe(false);
    expect(isApiKey('sk-something')).toBe(false);
  });
});
