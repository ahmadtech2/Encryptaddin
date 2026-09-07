import { argon2id } from "@noble/hashes/argon2";
import { asBufferSource, utf8 } from "./bytes";

export const KDF_PARAMS = {
  alg: "argon2id" as const,
  t: 3,
  m: 19456,
  p: 1,
  dkLen: 32,
};

export type KdfParams = typeof KDF_PARAMS;

export function deriveRawKey(passphrase: string, salt: Uint8Array, params: KdfParams = KDF_PARAMS): Uint8Array {
  if (!passphrase) {
    throw new Error("Passphrase is required.");
  }
  if (salt.length < 16) {
    throw new Error("Salt must be at least 16 bytes.");
  }
  return argon2id(utf8(passphrase), salt, {
    t: params.t,
    m: params.m,
    p: params.p,
    dkLen: params.dkLen,
  });
}

export async function importAesKey(raw: Uint8Array): Promise<CryptoKey> {
  return crypto.subtle.importKey("raw", asBufferSource(raw), { name: "AES-GCM" }, false, [
    "encrypt",
    "decrypt",
  ]);
}
