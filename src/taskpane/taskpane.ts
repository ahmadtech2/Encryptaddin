import {
  createVaultSession,
  decryptCell,
  encryptCell,
  mergeColumnTargets,
  passphraseScore,
  type ColumnTarget,
  type VaultMeta,
  WrongPassphraseError,
} from "../crypto/vault";
import {
  clearEncryptedFormatting,
  columnLetter,
  decryptColumns,
  decryptWorkbook,
  encryptColumns,
  markEncryptedColumns,
  scanActiveSheetHeaders,
  type HeaderColumn,
} from "../excel/columns";
import { loadMeta, saveMeta } from "../excel/settings";

type Session = {
  key: CryptoKey;
  meta: VaultMeta;
};

let session: Session | null = null;
let idleTimer: number | undefined;
const IDLE_MS = 10 * 60 * 1000;

function $(id: string): HTMLElement {
  const node = document.getElementById(id);
  if (!node) {
    throw new Error(`Missing #${id}`);
  }
  return node;
}

function setStatus(message: string, kind: "ok" | "error" | "" = ""): void {
  const status = $("status");
  status.textContent = message;
  status.className = `status ${kind}`.trim();
}

function touchSession(): void {
  window.clearTimeout(idleTimer);
  idleTimer = window.setTimeout(() => {
    session = null;
    ($("passphrase") as HTMLInputElement).value = "";
    setStatus("Session locked after 10 minutes idle.", "");
  }, IDLE_MS);
}

function selectedTargets(): ColumnTarget[] {
  return Array.from(document.querySelectorAll<HTMLInputElement>("input[data-col]:checked")).map((input) => ({
    sheet: input.dataset.sheet ?? "",
    col: Number(input.dataset.col),
    header: input.dataset.header ?? "",
  }));
}

function renderColumns(columns: HeaderColumn[], selected: ColumnTarget[]): void {
  const list = $("column-list");
  if (columns.length === 0) {
    list.innerHTML = `<p class="hint">No used range on this sheet. Add headers in row 1, then refresh.</p>`;
    return;
  }
  const selectedKeys = new Set(selected.map((item) => `${item.sheet}::${item.col}`));
  list.innerHTML = columns
    .map((column) => {
      const key = `${column.sheet}::${column.col}`;
      const preview = column.samples[0] ? column.samples[0].slice(0, 42) : "empty";
      const checked = selectedKeys.has(key) ? "checked" : "";
      return `<label class="column-item">
        <input type="checkbox" data-sheet="${escapeAttr(column.sheet)}" data-col="${column.col}" data-header="${escapeAttr(column.header)}" ${checked} />
        <div>
          <strong>${escapeHtml(column.header)} <span>(${columnLetter(column.col)})</span></strong>
          <span>${escapeHtml(preview)}</span>
        </div>
      </label>`;
    })
    .join("");
}

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[char] ?? char);
}

function escapeAttr(value: string): string {
  return escapeHtml(value);
}

async function refreshColumns(): Promise<void> {
  const meta = session?.meta ?? loadMeta();
  const scan = await scanActiveSheetHeaders();
  renderColumns(scan.columns, meta?.columns.filter((item) => item.sheet === scan.sheet) ?? []);
}

async function unlock(): Promise<Session> {
  const passphrase = ($("passphrase") as HTMLInputElement).value;
  const existing = loadMeta();
  setStatus("Deriving key…");
  const next = await createVaultSession(passphrase, existing);
  session = next;
  touchSession();
  return next;
}

async function runEncrypt(): Promise<void> {
  const targets = selectedTargets();
  if (targets.length === 0) {
    setStatus("Select at least one column.", "error");
    return;
  }
  if (!window.confirm(`Encrypt ${targets.length} column(s)? Recipients will see only ENC1 tokens in those cells.`)) {
    return;
  }
  const { key, meta } = await unlock();
  setStatus("Encrypting selected columns…");
  const result = await encryptColumns(targets, (plaintext) => encryptCell(key, plaintext));
  meta.columns = mergeColumnTargets(meta.columns, targets);
  await saveMeta(meta);
  session = { key, meta };
  await markEncryptedColumns(targets);
  setStatus(`Encrypted ${result.changed} cells. Skipped ${result.skipped}. Failed ${result.failed}.`, result.failed ? "error" : "ok");
}

async function runDecrypt(all: boolean): Promise<void> {
  const { key, meta } = await unlock();
  setStatus(all ? "Decrypting workbook…" : "Decrypting selected columns…");
  const result = all
    ? await decryptWorkbook((token) => decryptCell(key, token))
    : await decryptColumns(selectedTargets(), (token) => decryptCell(key, token));
  const targets = all ? meta.columns : selectedTargets();
  await clearEncryptedFormatting(targets);
  setStatus(`Decrypted ${result.changed} cells. Skipped ${result.skipped}. Failed ${result.failed}.`, result.failed ? "error" : "ok");
}

function toggleRekeyPanel(): void {
  const panel = $("rekey-panel");
  panel.hidden = !panel.hidden;
}

async function runRekey(): Promise<void> {
  const oldPhrase = ($("passphrase") as HTMLInputElement).value;
  const nextPhrase = ($("new-passphrase") as HTMLInputElement).value;
  if (!oldPhrase || !nextPhrase) {
    setStatus("Enter the current passphrase above and the new passphrase in the change panel.", "error");
    return;
  }
  if (!window.confirm("Re-encrypt every protected cell with the new passphrase?")) {
    return;
  }
  setStatus("Changing passphrase…");
  const current = await createVaultSession(oldPhrase, loadMeta());
  const next = await createVaultSession(nextPhrase, null);
  const result = await decryptWorkbook(async (token) => {
    const plain = await decryptCell(current.key, token);
    return encryptCell(next.key, plain);
  });
  next.meta.columns = current.meta.columns;
  await saveMeta(next.meta);
  session = { key: next.key, meta: next.meta };
  ($("passphrase") as HTMLInputElement).value = nextPhrase;
  ($("new-passphrase") as HTMLInputElement).value = "";
  $("rekey-panel").hidden = true;
  touchSession();
  setStatus(`Re-encrypted ${result.changed} cells with the new passphrase.`, result.failed ? "error" : "ok");
}

function lockSession(): void {
  session = null;
  ($("passphrase") as HTMLInputElement).value = "";
  setStatus("Session locked. Passphrase cleared from memory.", "ok");
}

function bind(): void {
  $("refresh-columns").addEventListener("click", () => {
    refreshColumns().catch((error) => setStatus(String(error), "error"));
  });
  $("encrypt").addEventListener("click", () => {
    runEncrypt().catch((error) => setStatus(error instanceof WrongPassphraseError ? error.message : String(error), "error"));
  });
  $("decrypt").addEventListener("click", () => {
    runDecrypt(false).catch((error) => setStatus(error instanceof WrongPassphraseError ? error.message : String(error), "error"));
  });
  $("decrypt-all").addEventListener("click", () => {
    runDecrypt(true).catch((error) => setStatus(error instanceof WrongPassphraseError ? error.message : String(error), "error"));
  });
  $("rekey").addEventListener("click", toggleRekeyPanel);
  $("confirm-rekey").addEventListener("click", () => {
    runRekey().catch((error) => setStatus(error instanceof WrongPassphraseError ? error.message : String(error), "error"));
  });
  $("lock").addEventListener("click", lockSession);
  $("toggle-secret").addEventListener("click", () => {
    const input = $("passphrase") as HTMLInputElement;
    const hidden = input.type === "password";
    input.type = hidden ? "text" : "password";
    $("toggle-secret").textContent = hidden ? "Hide" : "Show";
  });
  $("passphrase").addEventListener("input", () => {
    const strength = $("strength");
    const value = ($("passphrase") as HTMLInputElement).value;
    if (!value) {
      strength.textContent = "Use a long phrase you do not reuse. It is never stored in the file.";
      strength.className = "hint";
      return;
    }
    const result = passphraseScore(value);
    strength.textContent = `Strength: ${result.label}`;
    strength.className = `hint ${result.label.toLowerCase()}`;
  });
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    setStatus("Open this add-in in Excel.", "error");
    return;
  }
  bind();
  refreshColumns().catch((error) => setStatus(String(error), "error"));
});
