# Secure Columns (Encryptaddin)

Excel add-in for a manager who needs to share a workbook without exposing names, card numbers, or any other chosen columns.

The recipient does **not** install anything. Encrypted cells become `ENC1...` tokens. They can edit every other cell and send the file back. Only someone who knows the passphrase can decrypt.

## Manager flow

1. Open the workbook that still has the real values.
2. Home → **Secure Columns** to open the task pane.
3. Enter a long passphrase (12+ characters). It is never written into the file.
4. Refresh columns, check the sensitive ones (any headers, any sheet).
5. Click **Encrypt selected**.
6. Send the workbook. Staff edit unprotected cells only.
7. When the file comes back, open the add-in, enter the same passphrase, and click **Decrypt selected** or **Decrypt entire workbook**.

You can change the passphrase later with **Change passphrase**. That decrypts every token with the old phrase and re-encrypts with a new salt.

## Security model

| Property | Implementation |
| --- | --- |
| Encryption | AES-256-GCM (Web Crypto), 128-bit tag, random 12-byte IV per cell |
| Key derivation | Argon2id (`t=3`, `m=19456 KiB`, `p=1`) via `@noble/hashes` |
| Salt | 16 random bytes, stored in the workbook |
| Passphrase | Typed by the manager only. Not stored. Session key is dropped after 10 minutes idle or **Lock session** |
| Integrity | GCM authentication; a changed token fails closed |
| Tokens | `ENC1.` + base64url(IV \|\| ciphertext+tag). Same plaintext never produces the same token |
| Metadata | Workbook document settings: KDF params, salt, password verifier, last selected columns |

A wrong passphrase is rejected by the verifier before any cell is rewritten.

This is strong cryptography for a spreadsheet workflow. It is **not** a substitute for a bank card vault or HSM.

**Card numbers:** PCI DSS generally does not allow full PANs to live in Excel, even encrypted, and then be emailed around. Confirm with your compliance team. Prefer tokens from your card processor when that is an option.

## Install (any Mac or Windows Excel)

The add-in is hosted on GitHub Pages. Download the production manifest and upload it once per machine:

1. Open [https://ahmadtech2.github.io/Encryptaddin/](https://ahmadtech2.github.io/Encryptaddin/) and download `manifest.xml`.
2. In Excel: **Insert → Add-ins → My Add-ins → Upload My Add-in**.
3. Choose that XML file. **Secure Columns** appears on the Home tab.

Excel must be able to reach `https://ahmadtech2.github.io`. Recipients of an encrypted workbook still do not need the add-in.

For company-wide rollout, deploy the same hosted manifest through Microsoft 365 admin center (Integrated apps).

## Install (developer)

Requirements: Node.js 20+, Excel for Windows, Mac, or Excel on the web.

```bash
npm install
npm run build
npm start
```

`npm start` trusts a local HTTPS cert and sideloads the localhost `manifest.xml`. The first time, accept the office-addin-dev-certs prompt.

To sideload a local build by hand: **Insert → Add-ins → Upload my add-in** and choose `manifest.xml`, with `npm run dev-server` running.

## Recipients

They open the `.xlsx` as usual. Encrypted columns show unreadable `ENC1.` values and a light fill. They do not need the add-in, the passphrase, or macros.

Do not sort-cut only an encrypted column if that would detach it from the rest of the row. Sorting the whole table is fine.

## Scripts

- `npm test` — crypto unit tests (round-trip, wrong passphrase, tamper)
- `npm run typecheck`
- `npm run validate` — Office manifest validation
- `npm run build` — production webpack bundle

## Changing columns and phrases

Columns are not hard-coded. The pane lists the active sheet’s header row. Check whatever you need each time.

The passphrase is also not hard-coded. Each workbook has its own salt, so two files encrypted with the same phrase still have different keys.
