import { isEncryptedToken, type ColumnTarget } from "../crypto/vault";

export type HeaderColumn = ColumnTarget & { samples: string[] };

export function columnLetter(index: number): string {
  let n = index + 1;
  let label = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}

function cellText(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value);
}

export async function scanActiveSheetHeaders(): Promise<{ sheet: string; columns: HeaderColumn[] }> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    const used = sheet.getUsedRangeOrNullObject();
    used.load(["rowCount", "columnCount", "values", "isNullObject"]);
    await context.sync();

    if (used.isNullObject) {
      return { sheet: sheet.name, columns: [] };
    }

    const values = used.values as unknown[][];
    const headerRow = values[0] ?? [];
    const columns: HeaderColumn[] = headerRow.map((header, col) => {
      const samples = values
        .slice(1, 4)
        .map((row) => cellText(row[col]))
        .filter((text) => text.length > 0);
      return {
        sheet: sheet.name,
        col,
        header: cellText(header) || `Column ${columnLetter(col)}`,
        samples,
      };
    });
    return { sheet: sheet.name, columns };
  });
}

const CHUNK = 1500;

async function mapColumnCells(
  sheetName: string,
  col: number,
  mapper: (text: string, formula: string) => Promise<string | null>
): Promise<{ changed: number; skipped: number; failed: number }> {
  let changed = 0;
  let skipped = 0;
  let failed = 0;

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const used = sheet.getUsedRangeOrNullObject();
    used.load(["rowCount", "isNullObject"]);
    await context.sync();
    if (used.isNullObject || used.rowCount <= 1) {
      return;
    }

    const dataRows = used.rowCount - 1;
    for (let offset = 0; offset < dataRows; offset += CHUNK) {
      const rows = Math.min(CHUNK, dataRows - offset);
      const range = sheet.getRangeByIndexes(1 + offset, col, rows, 1);
      range.load(["text", "formulas", "values"]);
      await context.sync();

      const text = range.text as string[][];
      const formulas = range.formulas as string[][];
      const next = text.map((row) => [...row]);

      for (let i = 0; i < text.length; i += 1) {
        const formula = cellText(formulas[i]?.[0]);
        const current = cellText(text[i]?.[0]);
        try {
          const replacement = await mapper(current, formula);
          if (replacement === null) {
            skipped += 1;
          } else {
            next[i][0] = replacement;
            changed += 1;
          }
        } catch {
          failed += 1;
        }
      }

      range.values = next;
      await context.sync();
    }
  });

  return { changed, skipped, failed };
}

export async function encryptColumns(
  targets: ColumnTarget[],
  encryptValue: (plaintext: string) => Promise<string>
): Promise<{ changed: number; skipped: number; failed: number }> {
  const totals = { changed: 0, skipped: 0, failed: 0 };
  for (const target of targets) {
    const result = await mapColumnCells(target.sheet, target.col, async (text, formula) => {
      if (!text) {
        return null;
      }
      if (isEncryptedToken(text)) {
        return null;
      }
      if (formula.startsWith("=")) {
        return null;
      }
      return encryptValue(text);
    });
    totals.changed += result.changed;
    totals.skipped += result.skipped;
    totals.failed += result.failed;
  }
  return totals;
}

export async function decryptColumns(
  targets: ColumnTarget[],
  decryptValue: (token: string) => Promise<string>
): Promise<{ changed: number; skipped: number; failed: number }> {
  const totals = { changed: 0, skipped: 0, failed: 0 };
  for (const target of targets) {
    const result = await mapColumnCells(target.sheet, target.col, async (text) => {
      if (!isEncryptedToken(text)) {
        return null;
      }
      return decryptValue(text);
    });
    totals.changed += result.changed;
    totals.skipped += result.skipped;
    totals.failed += result.failed;
  }
  return totals;
}

export async function decryptWorkbook(decryptValue: (token: string) => Promise<string>): Promise<{
  changed: number;
  skipped: number;
  failed: number;
}> {
  const totals = { changed: 0, skipped: 0, failed: 0 };
  const sheets = await Excel.run(async (context) => {
    const list = context.workbook.worksheets;
    list.load("items/name");
    await context.sync();
    return list.items.map((item) => item.name);
  });

  for (const sheetName of sheets) {
    const columns = await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      const used = sheet.getUsedRangeOrNullObject();
      used.load(["columnCount", "isNullObject"]);
      await context.sync();
      return used.isNullObject ? 0 : used.columnCount;
    });
    for (let col = 0; col < columns; col += 1) {
      const result = await mapColumnCells(sheetName, col, async (text) => {
        if (!isEncryptedToken(text)) {
          return null;
        }
        return decryptValue(text);
      });
      totals.changed += result.changed;
      totals.skipped += result.skipped;
      totals.failed += result.failed;
    }
  }
  return totals;
}

export async function markEncryptedColumns(targets: ColumnTarget[]): Promise<void> {
  await Excel.run(async (context) => {
    for (const target of targets) {
      const sheet = context.workbook.worksheets.getItem(target.sheet);
      const used = sheet.getUsedRangeOrNullObject();
      used.load(["rowCount", "isNullObject"]);
      await context.sync();
      if (used.isNullObject || used.rowCount <= 1) {
        continue;
      }
      const range = sheet.getRangeByIndexes(1, target.col, used.rowCount - 1, 1);
      range.format.fill.color = "#E8EEF5";
      range.format.font.color = "#334155";
    }
    await context.sync();
  });
}

export async function clearEncryptedFormatting(targets: ColumnTarget[]): Promise<void> {
  await Excel.run(async (context) => {
    for (const target of targets) {
      const sheet = context.workbook.worksheets.getItem(target.sheet);
      const used = sheet.getUsedRangeOrNullObject();
      used.load(["rowCount", "isNullObject"]);
      await context.sync();
      if (used.isNullObject || used.rowCount <= 1) {
        continue;
      }
      const range = sheet.getRangeByIndexes(1, target.col, used.rowCount - 1, 1);
      range.format.fill.clear();
      range.format.font.color = "#000000";
    }
    await context.sync();
  });
}
