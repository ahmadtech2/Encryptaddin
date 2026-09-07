import { META_KEY, parseMeta, type VaultMeta } from "../crypto/vault";

export function loadMeta(): VaultMeta | null {
  return parseMeta(Office.context.document.settings.get(META_KEY));
}

export function saveMeta(meta: VaultMeta): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.settings.set(META_KEY, JSON.stringify(meta));
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error.message));
        return;
      }
      resolve();
    });
  });
}
