import { describe, expect, it } from "vitest";
import { fromBase64Url, toBase64Url } from "../src/crypto/bytes";
import {
  createVaultSession,
  decryptCell,
  encryptCell,
  isEncryptedToken,
  mergeColumnTargets,
  passphraseScore,
  TamperedTokenError,
  TOKEN_PREFIX,
  WrongPassphraseError,
} from "../src/crypto/vault";

const PHRASE = "correct horse battery staple";
const OTHER = "wrong horse battery staple!";

describe("secure columns vault", () => {
  it("rejects short passphrases", async () => {
    await expect(createVaultSession("short")).rejects.toThrow(/at least 12/);
  });

  it("round-trips plaintext and uses a unique token each time", async () => {
    const { key } = await createVaultSession(PHRASE);
    const first = await encryptCell(key, "4111111111111111");
    const second = await encryptCell(key, "4111111111111111");
    expect(first).toMatch(new RegExp(`^${TOKEN_PREFIX}`));
    expect(first).not.toBe(second);
    expect(await decryptCell(key, first)).toBe("4111111111111111");
    expect(await decryptCell(key, second)).toBe("4111111111111111");
  });

  it("rejects the wrong passphrase without touching tokens", async () => {
    const created = await createVaultSession(PHRASE);
    await expect(createVaultSession(OTHER, created.meta)).rejects.toBeInstanceOf(WrongPassphraseError);
  });

  it("detects tampered ciphertext", async () => {
    const { key } = await createVaultSession(PHRASE);
    const token = await encryptCell(key, "Ada Lovelace");
    const blob = fromBase64Url(token.slice(TOKEN_PREFIX.length));
    blob[20] ^= 0xff;
    const mutated = TOKEN_PREFIX + toBase64Url(blob);
    await expect(decryptCell(key, mutated)).rejects.toBeInstanceOf(TamperedTokenError);
  });

  it("recognizes encrypted tokens only", () => {
    expect(isEncryptedToken("ENC1.abc")).toBe(true);
    expect(isEncryptedToken("Ada Lovelace")).toBe(false);
    expect(isEncryptedToken(1234)).toBe(false);
  });

  it("scores passphrases", () => {
    expect(passphraseScore("aaaaaaaaaaaa").label).toBe("Weak");
    expect(passphraseScore("Correct-Horse-Battery-Staple-32").label).toBe("Strong");
  });

  it("merges column selections by sheet and index", () => {
    const merged = mergeColumnTargets(
      [{ sheet: "Sheet1", col: 0, header: "Name" }],
      [
        { sheet: "Sheet1", col: 0, header: "Full name" },
        { sheet: "Sheet1", col: 2, header: "Card" },
      ]
    );
    expect(merged).toEqual([
      { sheet: "Sheet1", col: 0, header: "Full name" },
      { sheet: "Sheet1", col: 2, header: "Card" },
    ]);
  });

  it("unlocks an existing workbook meta with the original phrase", async () => {
    const first = await createVaultSession(PHRASE);
    const again = await createVaultSession(PHRASE, first.meta);
    const token = await encryptCell(first.key, "Nora");
    expect(await decryptCell(again.key, token)).toBe("Nora");
  });
});
