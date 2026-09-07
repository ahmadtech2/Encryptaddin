import { asBufferSource, concatBytes, fromBase64Url, fromUtf8, randomBytes, toBase64Url, utf8 } from "./bytes";
import { decryptAesGcm, encryptAesGcm } from "./aes";
import { deriveRawKey, importAesKey, KDF_PARAMS, type KdfParams } from "./kdf";

export const TOKEN_PREFIX = "ENC1.";
export const META_KEY = "secureColumns.meta";
export const AAD = utf8("secure-columns:v1");
const VERIFIER_PLAINTEXT = "SECURE-COLUMNS-OK";
const MIN_PASSPHRASE_LENGTH = 12;

export type ColumnTarget = {
  sheet: string;
  col: number;
  header: string;
};

export type VaultMeta = {
  v: 1;
  kdf: {
    alg: "argon2id";
    t: number;
    m: number;
    p: number;
    dkLen: number;
    salt: string;
  };
  verify: {
    iv: string;
    ct: string;
  };
  columns: ColumnTarget[];
};

export class WrongPassphraseError extends Error {
  constructor() {
    super("Wrong passphrase. The file was not changed.");
    this.name = "WrongPassphraseError";
  }
}

export class TamperedTokenError extends Error {
  constructor() {
    super("Encrypted value was altered and cannot be decrypted.");
    this.name = "TamperedTokenError";
  }
}

export function isEncryptedToken(value: unknown): value is string {
  return typeof value === "string" && value.startsWith(TOKEN_PREFIX);
}

export function assertPassphraseStrength(passphrase: string): void {
  if (passphrase.length < MIN_PASSPHRASE_LENGTH) {
    throw new Error(`Passphrase must be at least ${MIN_PASSPHRASE_LENGTH} characters.`);
  }
}

export function passphraseScore(passphrase: string): { score: number; label: string } {
  let score = 0;
  if (passphrase.length >= 12) score += 1;
  if (passphrase.length >= 16) score += 1;
  if (passphrase.length >= 20) score += 1;
  if (/[a-z]/.test(passphrase) && /[A-Z]/.test(passphrase)) score += 1;
  if (/\d/.test(passphrase)) score += 1;
  if (/[^A-Za-z0-9]/.test(passphrase)) score += 1;
  if (score <= 2) return { score, label: "Weak" };
  if (score <= 4) return { score, label: "Fair" };
  return { score, label: "Strong" };
}

export function createEmptyMeta(): VaultMeta {
  return {
    v: 1,
    kdf: {
      ...KDF_PARAMS,
      salt: toBase64Url(randomBytes(16)),
    },
    verify: { iv: "", ct: "" },
    columns: [],
  };
}

export async function createVaultSession(passphrase: string, existing?: VaultMeta | null): Promise<{
  key: CryptoKey;
  meta: VaultMeta;
}> {
  assertPassphraseStrength(passphrase);
  const meta = existing ?? createEmptyMeta();
  const salt = fromBase64Url(meta.kdf.salt);
  const params: KdfParams = {
    alg: "argon2id",
    t: meta.kdf.t,
    m: meta.kdf.m,
    p: meta.kdf.p,
    dkLen: meta.kdf.dkLen,
  };
  const raw = deriveRawKey(passphrase, salt, params);
  const key = await importAesKey(raw);

  if (meta.verify.iv && meta.verify.ct) {
    try {
      const plain = await decryptAesGcm(key, fromBase64Url(meta.verify.iv), fromBase64Url(meta.verify.ct), AAD);
      if (fromUtf8(asBufferSource(plain)) !== VERIFIER_PLAINTEXT) {
        throw new WrongPassphraseError();
      }
    } catch (error) {
      if (error instanceof WrongPassphraseError) {
        throw error;
      }
      throw new WrongPassphraseError();
    }
  } else {
    const sealed = await encryptAesGcm(key, utf8(VERIFIER_PLAINTEXT), AAD);
    meta.verify = { iv: toBase64Url(sealed.iv), ct: toBase64Url(sealed.ciphertext) };
  }

  return { key, meta };
}

export async function encryptCell(key: CryptoKey, plaintext: string): Promise<string> {
  const sealed = await encryptAesGcm(key, utf8(plaintext), AAD);
  return TOKEN_PREFIX + toBase64Url(concatBytes(sealed.iv, sealed.ciphertext));
}

export async function decryptCell(key: CryptoKey, token: string): Promise<string> {
  if (!isEncryptedToken(token)) {
    throw new Error("Value is not an encrypted token.");
  }
  const blob = fromBase64Url(token.slice(TOKEN_PREFIX.length));
  if (blob.length < 13) {
    throw new TamperedTokenError();
  }
  const iv = blob.slice(0, 12);
  const ciphertext = blob.slice(12);
  try {
    const plain = await decryptAesGcm(key, iv, ciphertext, AAD);
    return fromUtf8(asBufferSource(plain));
  } catch {
    throw new TamperedTokenError();
  }
}

export function mergeColumnTargets(current: ColumnTarget[], selected: ColumnTarget[]): ColumnTarget[] {
  const byKey = new Map(current.map((item) => [`${item.sheet}::${item.col}`, item]));
  for (const item of selected) {
    byKey.set(`${item.sheet}::${item.col}`, item);
  }
  return [...byKey.values()];
}

export function parseMeta(raw: unknown): VaultMeta | null {
  if (!raw || typeof raw !== "string") {
    return null;
  }
  try {
    const parsed = JSON.parse(raw) as VaultMeta;
    if (parsed.v !== 1 || !parsed.kdf?.salt) {
      return null;
    }
    return parsed;
  } catch {
    return null;
  }
}
