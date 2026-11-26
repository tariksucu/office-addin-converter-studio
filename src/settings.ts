import { EMAIL_ATTACHMENT_MB, isFitTarget, type FitTarget } from './fit/budget';

/**
 * The pane's one remembered choice besides the key: the total "Fit attachments for sending"
 * aims at. Kept like the key (the add-in's own storage under Office's partition key), and in
 * memory where storage is blocked.
 */

const STORAGE_KEY = 'converter-studio.fit-target';

export interface FitTargetStore {
  get(): FitTarget;
  set(target: FitTarget): void;
}

export function fitTargetStore(storage: Storage | null, partition: string): FitTargetStore {
  const name = `${partition}${STORAGE_KEY}`;
  let memory: FitTarget = EMAIL_ATTACHMENT_MB;
  return {
    get() {
      try {
        const value = Number(storage?.getItem(name));
        return isFitTarget(value) ? value : memory;
      } catch {
        return memory;
      }
    },
    set(target) {
      memory = target;
      try {
        storage?.setItem(name, String(target));
      } catch {
        // Kept in memory only.
      }
    },
  };
}
