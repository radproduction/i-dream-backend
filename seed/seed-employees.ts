import "dotenv/config";
import path from "node:path";
import xlsx from "xlsx";
import mongoose from "mongoose";
import { connectToMongoDB } from "../mongodb";
import { User } from "../models";
import { createEmployee } from "../db";

type RawRow = Array<unknown>;

const DEFAULT_PASSWORD = process.env.SEED_DEFAULT_PASSWORD || "123";
const DEFAULT_SHEET = process.env.SEED_SHEET_NAME || "";

const isNonEmptyString = (value: unknown): value is string =>
  typeof value === "string" && value.trim().length > 0;

const toCleanString = (value: unknown): string =>
  typeof value === "string" ? value.trim() : "";

const isDateLike = (value: unknown): boolean =>
  value instanceof Date ||
  (typeof value === "number" && value > 20000 && value < 80000);

const isNameLike = (value: unknown): boolean =>
  typeof value === "string" && /[a-zA-Z]/.test(value);

const isEmailLike = (value: unknown): boolean =>
  typeof value === "string" && value.includes("@");

const isEmpIdLike = (value: unknown): boolean =>
  typeof value === "string" && /[a-zA-Z]/.test(value) && /\d/.test(value);

function normalizeEmployeeId(raw: string): string {
  const trimmed = raw.trim();
  if (!trimmed) return trimmed;
  // Normalize common patterns like "IDM - 0001" -> "IDM-0001"
  return trimmed.replace(/\s*-\s*/g, "-").replace(/\s+/g, "");
}

function findHeaderRow(rows: RawRow[]): number {
  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i]
      .map(value => (typeof value === "string" ? value.trim() : ""))
      .filter(Boolean);
    const joined = row.join(" ").toLowerCase();
    if (joined.includes("name") && joined.includes("department")) {
      return i;
    }
  }
  return -1;
}

function findColumnIndex(headers: string[], regex: RegExp): number | null {
  const idx = headers.findIndex(value => regex.test(value));
  return idx >= 0 ? idx : null;
}

function getFirstDataRow(rows: RawRow[], headerRow: number): RawRow | null {
  for (let i = headerRow + 1; i < rows.length; i += 1) {
    const row = rows[i];
    if (row.some(value => isNonEmptyString(value) || typeof value === "number")) {
      return row;
    }
  }
  return null;
}

function resolveIndex(baseIdx: number | null, shift: number): number | null {
  if (baseIdx === null) return null;
  return baseIdx + shift;
}

function getCell(row: RawRow, idx: number | null): unknown {
  if (idx === null) return undefined;
  return row[idx];
}

function getCellString(row: RawRow, idx: number | null): string {
  const value = getCell(row, idx);
  return toCleanString(value);
}

async function run() {
  const fileArg = process.argv[2];
  if (!fileArg) {
    console.error("Usage: tsx seed/seed-employees.ts <path-to-xlsx>");
    process.exit(1);
  }

  const filePath = path.resolve(fileArg);
  const workbook = xlsx.readFile(filePath, { cellDates: true });
  const sheetName = DEFAULT_SHEET || workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  if (!sheet) {
    console.error(`[Seed] Sheet not found: ${sheetName}`);
    process.exit(1);
  }

  const rows = xlsx.utils.sheet_to_json<RawRow>(sheet, {
    header: 1,
    defval: "",
  });

  const headerRowIndex = findHeaderRow(rows);
  if (headerRowIndex === -1) {
    console.error("[Seed] Could not locate header row.");
    process.exit(1);
  }

  const rawHeaders = rows[headerRowIndex].map(value =>
    typeof value === "string" ? value.trim() : ""
  );

  const firstDataRow = getFirstDataRow(rows, headerRowIndex);
  if (!firstDataRow) {
    console.error("[Seed] No data rows found.");
    process.exit(1);
  }

  const nameIdxBase = findColumnIndex(rawHeaders, /name of staff|name/i);
  const deptIdxBase = findColumnIndex(rawHeaders, /department/i);
  const desigIdxBase = findColumnIndex(rawHeaders, /designation|position/i);
  const emailIdxBase = findColumnIndex(rawHeaders, /email/i);
  const dojIdxBase = findColumnIndex(rawHeaders, /doj|joining/i);
  const serialIdxBase = findColumnIndex(rawHeaders, /s\.?\s*no|sr\.?\s*no|sno/i);
  const empIdIdxBase =
    findColumnIndex(rawHeaders, /employee id|emp id|staff id|id/i) ?? null;

  let shift = 0;
  if (nameIdxBase !== null) {
    const nameCell = getCell(firstDataRow, nameIdxBase);
    const nextCell = getCell(firstDataRow, nameIdxBase + 1);
    if (isDateLike(nameCell) && isNameLike(nextCell)) {
      shift = 1;
    }
  }

  let employeeIdIdx = resolveIndex(empIdIdxBase, shift);
  const nameIdx = resolveIndex(nameIdxBase, shift);
  const deptIdx = resolveIndex(deptIdxBase, shift);
  const desigIdx = resolveIndex(desigIdxBase, shift);
  const emailIdx = resolveIndex(emailIdxBase, shift);
  const dojIdx = resolveIndex(dojIdxBase, shift);
  const serialIdx = resolveIndex(serialIdxBase, shift);
  const statusIdx = resolveIndex(
    findColumnIndex(rawHeaders, /status/i),
    shift
  );

  if (employeeIdIdx === null && dojIdxBase !== null) {
    const candidate = getCell(firstDataRow, dojIdxBase);
    if (isEmpIdLike(candidate)) {
      employeeIdIdx = dojIdxBase;
    }
  }

  if (employeeIdIdx === null) {
    for (let i = 0; i < rawHeaders.length; i += 1) {
      const header = rawHeaders[i];
      if (header) continue;
      const candidate = getCell(firstDataRow, i);
      if (isEmpIdLike(candidate)) {
        employeeIdIdx = i;
        break;
      }
    }
  }

  const connected = await connectToMongoDB();
  if (!connected) {
    console.error("[Seed] MongoDB not connected. Check MONGODB_URI.");
    process.exit(1);
  }

  let created = 0;
  let updated = 0;
  let skipped = 0;

  for (let i = headerRowIndex + 1; i < rows.length; i += 1) {
    const row = rows[i];
    if (!row || row.length === 0) continue;

    const rawEmployeeId = getCellString(row, employeeIdIdx);
    let employeeId = normalizeEmployeeId(rawEmployeeId);
    const name = getCellString(row, nameIdx);
    const department = getCellString(row, deptIdx);
    const position = getCellString(row, desigIdx);
    const email = getCellString(row, emailIdx);
    const status = getCellString(row, statusIdx).toLowerCase();

    if (!employeeId) {
      const serialValue = getCell(row, serialIdx);
      if (typeof serialValue === "number" && Number.isFinite(serialValue)) {
        employeeId = `IDM-${String(Math.trunc(serialValue)).padStart(4, "0")}`;
      }
    }

    if (!employeeId || !name) {
      skipped += 1;
      continue;
    }

    if (status && !status.includes("employ")) {
      skipped += 1;
      continue;
    }

    const role = department.toLowerCase() === "admin" ? "admin" : "user";

    const existing = await User.findOne({ employeeId }).lean();
    if (!existing) {
      await createEmployee({
        name,
        email: email || undefined,
        employeeId,
        password: DEFAULT_PASSWORD,
        department: department || undefined,
        position: position || undefined,
        role,
      });
      created += 1;
      continue;
    }

    await User.updateOne(
      { _id: existing._id },
      {
        $set: {
          name,
          email: email || existing.email,
          department: department || existing.department,
          position: position || existing.position,
          role,
        },
      }
    );
    updated += 1;
  }

  console.log(
    `[Seed] Employees import complete. Created: ${created}, Updated: ${updated}, Skipped: ${skipped}`
  );

  await mongoose.connection.close();
}

run().catch(error => {
  console.error("[Seed] Failed:", error);
  process.exit(1);
});
