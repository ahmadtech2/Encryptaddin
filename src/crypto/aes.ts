import { asBufferSource, randomBytes } from "./bytes";

const IV_LENGTH = 12;
const TAG_BITS = 128;

export async function encryptAesGcm(
  key: CryptoKey,
  plaintext: Uint8Array,
  additionalData: Uint8Array
): Promise<{ iv: Uint8Array; ciphertext: Uint8Array }> {
  const iv = randomBytes(IV_LENGTH);
  const ciphertext = new Uint8Array(
    await crypto.subtle.encrypt(
      { name: "AES-GCM", iv: asBufferSource(iv), additionalData: asBufferSource(additionalData), tagLength: TAG_BITS },
      key,
      asBufferSource(plaintext)
    )
  );
  return { iv, ciphertext };
}

export async function decryptAesGcm(
  key: CryptoKey,
  iv: Uint8Array,
  ciphertext: Uint8Array,
  additionalData: Uint8Array
): Promise<Uint8Array> {
  const plaintext = await crypto.subtle.decrypt(
    { name: "AES-GCM", iv: asBufferSource(iv), additionalData: asBufferSource(additionalData), tagLength: TAG_BITS },
    key,
    asBufferSource(ciphertext)
  );
  return new Uint8Array(plaintext);
}
