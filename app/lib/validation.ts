// Pure validation logic for the loan-data checker. No DOM/React dependencies
// so it can be unit-tested with plain Node + the `xlsx` package.
//
// Core rule: never guess. A missing value is left blank and flagged; a
// present-but-bad value is kept as-is and flagged as invalid; a value that
// can be safely inferred from the rest of the file is offered as a
// suggestion that staff must explicitly accept — it is never auto-applied.

export type FieldKey =
  | "branch"
  | "customer_name"
  | "loan_amount"
  | "loan_date"
  | "phone"
  | "status";

export type Severity = "high" | "medium" | "low";
export type FlagKind = "missing" | "invalid" | "info";

export interface Config {
  maxLoanAmount: number;
  minLoanAmount: number;
  minYear: number;
  recentClosedDays: number;
  knownBranches: string[]; // empty = check disabled
  branchSuggestionThreshold: number; // e.g. 0.8 = suggest only if one branch covers 80%+ of rows
}

export const DEFAULT_CONFIG: Config = {
  maxLoanAmount: 10_000_000, // ₹1 crore
  minLoanAmount: 1_000,
  minYear: 2000,
  recentClosedDays: 30,
  knownBranches: [],
  branchSuggestionThreshold: 0.8,
};

// How bad a gap in this field is, used to sort the review queue and badge
// each row. Reused for "invalid" flags on the same field too.
export const FIELD_SEVERITY: Record<FieldKey, Severity> = {
  customer_name: "high",
  loan_amount: "high",
  status: "high",
  loan_date: "medium",
  phone: "medium",
  branch: "low",
};

const SEVERITY_RANK: Record<Severity, number> = { high: 3, medium: 2, low: 1 };

export function severityRank(s: Severity | "none"): number {
  return s === "none" ? 0 : SEVERITY_RANK[s];
}

interface FieldDef {
  key: FieldKey;
  label: string;
  variants: string[];
}

// Order matters: more specific fields are matched first so a generic variant
// like "name" (used for customer_name) doesn't steal a column meant for
// another field (e.g. "Branch Name").
const FIELD_DEFS: FieldDef[] = [
  {
    key: "branch",
    label: "Branch",
    variants: ["branch", "branch name", "br", "branch code"],
  },
  {
    key: "loan_amount",
    label: "Loan Amount",
    variants: [
      "loan amount",
      "amount",
      "amt",
      "loan amt",
      "principal",
      "principal amount",
      "sanctioned amount",
      "disbursed amount",
    ],
  },
  {
    key: "loan_date",
    label: "Loan Date",
    variants: [
      "loan date",
      "date",
      "disbursement date",
      "disb date",
      "disburse date",
      "sanction date",
      "date of loan",
    ],
  },
  {
    key: "phone",
    label: "Phone",
    variants: [
      "phone",
      "phone no",
      "phone number",
      "mobile",
      "mobile no",
      "mobile number",
      "contact",
      "contact no",
      "contact number",
      "cell",
      "cell no",
      "cell number",
      "telephone",
      "tel",
      "tel no",
    ],
  },
  {
    key: "status",
    label: "Status",
    variants: ["status", "loan status", "account status"],
  },
  {
    key: "customer_name",
    label: "Customer Name",
    variants: [
      "customer name",
      "cust name",
      "name",
      "client name",
      "borrower name",
      "customer",
      "full name",
    ],
  },
];

const JUNK_NAMES = new Set([
  "na",
  "n a",
  "nil",
  "xx",
  "xxx",
  "xxxx",
  "-",
  ".",
  "test",
  "dummy",
  "unknown",
  "none",
]);

const STATUS_MAP: Record<string, "Active" | "Closed" | "NPA"> = {
  active: "Active",
  ongoing: "Active",
  running: "Active",
  live: "Active",
  open: "Active",
  inprogress: "Active",
  closed: "Closed",
  close: "Closed",
  paid: "Closed",
  paidoff: "Closed",
  settled: "Closed",
  completed: "Closed",
  complete: "Closed",
  npa: "NPA",
  default: "NPA",
  defaulted: "NPA",
  overdue: "NPA",
  bad: "NPA",
  writtenoff: "NPA",
};

export interface CellFlag {
  field: FieldKey | "cross-row" | "row";
  kind: FlagKind;
  severity: Severity;
  message: string;
}

export interface Suggestion {
  field: FieldKey;
  value: string;
  message: string;
}

export interface CleanedRow {
  branch: string;
  customer_name: string;
  loan_amount: number | null;
  loan_amount_raw: string;
  loan_date: Date | null;
  loan_date_raw: string;
  phone: string;
  phone_raw: string;
  status: "Active" | "Closed" | "NPA" | "";
  status_raw: string;
}

export interface RowResult {
  rowNumber: number; // spreadsheet row number (header = row 1)
  raw: Record<FieldKey, unknown>;
  cleaned: CleanedRow;
  flags: CellFlag[];
  suggestions: Suggestion[];
  severity: Severity | "none";
  skipped: boolean;
  approved: boolean;
}

export interface ValidationReport {
  fatalError?: string;
  missingColumns: string[];
  extraColumns: string[];
  totalRows: number;
  cleanCount: number;
  reviewCount: number;
  skippedCount: number;
  rows: RowResult[];
  colIndex: Record<FieldKey, number>;
}

function normalize(s: unknown): string {
  return String(s ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

export function mapColumns(headerRow: unknown[]) {
  const normalizedHeaders = headerRow.map((h) => normalize(h));
  const colIndex = {} as Record<FieldKey, number>;
  const usedIdx = new Set<number>();

  for (const def of FIELD_DEFS) {
    let found = -1;
    for (const variant of def.variants) {
      const nv = normalize(variant);
      const idx = normalizedHeaders.findIndex(
        (h, i) => !usedIdx.has(i) && h === nv
      );
      if (idx >= 0) {
        found = idx;
        break;
      }
    }
    colIndex[def.key] = found;
    if (found >= 0) usedIdx.add(found);
  }

  const missing = FIELD_DEFS.filter((d) => colIndex[d.key] === -1).map(
    (d) => d.label
  );
  const extra = headerRow
    .map((h, i) => ({ h, i }))
    .filter(({ h, i }) => !usedIdx.has(i) && String(h ?? "").trim() !== "")
    .map(({ h }) => String(h));

  return { colIndex, missing, extra };
}

function cleanText(raw: unknown): string {
  return String(raw ?? "").trim();
}

function parseAmount(raw: unknown): { value: number | null; raw: string } {
  if (raw === null || raw === undefined || raw === "")
    return { value: null, raw: "" };
  if (typeof raw === "number") return { value: raw, raw: String(raw) };

  const rawStr = String(raw).trim();
  if (!rawStr) return { value: null, raw: "" };

  let cleaned = rawStr
    .replace(/₹/g, "")
    .replace(/rs\.?/gi, "")
    .replace(/,/g, "")
    .trim();

  let multiplier = 1;
  if (/cr$/i.test(cleaned)) {
    multiplier = 10_000_000;
    cleaned = cleaned.replace(/cr$/i, "").trim();
  } else if (/l$/i.test(cleaned)) {
    multiplier = 100_000;
    cleaned = cleaned.replace(/l$/i, "").trim();
  }

  const num = parseFloat(cleaned);
  if (isNaN(num)) return { value: null, raw: rawStr };
  return { value: num * multiplier, raw: rawStr };
}

function excelSerialToDate(serial: number): Date | null {
  if (!isFinite(serial) || serial <= 0) return null;
  const utcDays = Math.floor(serial - 25569);
  const date = new Date(utcDays * 86400 * 1000);
  return isNaN(date.getTime()) ? null : date;
}

interface DateParseResult {
  date: Date | null;
  raw: string;
  ambiguous: boolean;
  invalidFormat: boolean;
}

function parseLoanDate(raw: unknown): DateParseResult {
  if (raw === null || raw === undefined || raw === "")
    return { date: null, raw: "", ambiguous: false, invalidFormat: false };

  if (raw instanceof Date) {
    return {
      date: raw,
      raw: raw.toISOString().slice(0, 10),
      ambiguous: false,
      invalidFormat: isNaN(raw.getTime()),
    };
  }

  if (typeof raw === "number") {
    const d = excelSerialToDate(raw);
    return {
      date: d,
      raw: String(raw),
      ambiguous: false,
      invalidFormat: !d,
    };
  }

  const s = String(raw).trim();
  if (!s) return { date: null, raw: "", ambiguous: false, invalidFormat: false };

  // DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY (and 2-digit years)
  let m = s.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})$/);
  if (m) {
    const a = parseInt(m[1], 10);
    const b = parseInt(m[2], 10);
    let year = parseInt(m[3], 10);
    if (year < 100) year += 2000;

    let day = a;
    let month = b;
    let invalid = false;

    if (month > 12 && day <= 12) {
      // Only valid as MM/DD — swap.
      day = b;
      month = a;
    }
    if (month < 1 || month > 12 || day < 1 || day > 31) invalid = true;

    const ambiguous = !invalid && a <= 12 && b <= 12;
    const date = invalid ? null : new Date(year, month - 1, day);
    // Reject impossible calendar dates (e.g. 31/02/2024) that JS silently rolls over.
    const rolledOver =
      date !== null &&
      (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day);
    return { date: rolledOver ? null : date, raw: s, ambiguous, invalidFormat: invalid || rolledOver };
  }

  // YYYY-MM-DD
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) {
    const year = parseInt(m[1], 10);
    const month = parseInt(m[2], 10);
    const day = parseInt(m[3], 10);
    const invalid = month < 1 || month > 12 || day < 1 || day > 31;
    const date = invalid ? null : new Date(year, month - 1, day);
    const rolledOver =
      date !== null &&
      (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day);
    return {
      date: rolledOver ? null : date,
      raw: s,
      ambiguous: false,
      invalidFormat: invalid || rolledOver,
    };
  }

  // Fallback: named months, e.g. "4 Jan 2024" / "Jan 4, 2024"
  const parsed = Date.parse(s);
  if (!isNaN(parsed)) {
    return { date: new Date(parsed), raw: s, ambiguous: false, invalidFormat: false };
  }

  return { date: null, raw: s, ambiguous: false, invalidFormat: true };
}

function cleanPhone(raw: unknown): { digits: string; raw: string } {
  if (raw === null || raw === undefined || raw === "")
    return { digits: "", raw: "" };
  const rawStr = String(raw).trim();
  let s = rawStr.replace(/[\s\-()]/g, "");

  // Only strip a country code / trunk prefix when doing so leaves exactly a
  // 10-digit number — otherwise a genuine number that happens to start with
  // "91" or "0" (e.g. 9123456789) gets wrongly truncated.
  if (s.startsWith("+91") && s.length === 13) {
    s = s.slice(3);
  } else if (s.startsWith("91") && s.length === 12) {
    s = s.slice(2);
  } else if (s.startsWith("0") && s.length === 11) {
    s = s.slice(1);
  }

  const digits = s.replace(/\D/g, "");
  return { digits, raw: rawStr };
}

function isFakePhone(digits: string): boolean {
  if (/^(\d)\1{9}$/.test(digits)) return true; // all same digit
  const knownFake = new Set([
    "1234567890",
    "0123456789",
    "9876543210",
    "0987654321",
  ]);
  return knownFake.has(digits);
}

function mapStatus(raw: unknown): { value: "Active" | "Closed" | "NPA" | ""; raw: string } {
  const rawStr = cleanText(raw);
  const key = rawStr.toLowerCase().replace(/[^a-z]/g, "");
  if (!key) return { value: "", raw: rawStr };
  return { value: STATUS_MAP[key] ?? "", raw: rawStr };
}

function addFlag(
  row: RowResult,
  field: CellFlag["field"],
  kind: FlagKind,
  severity: Severity,
  message: string
) {
  if (row.flags.some((f) => f.message === message)) return;
  row.flags.push({ field, kind, severity, message });
}

function getCell(row: unknown[], idx: number): unknown {
  return idx >= 0 ? row[idx] : undefined;
}

function groupKey(parts: (string | number)[]): string {
  return parts.join("||");
}

function pushTo<T>(map: Map<string, T[]>, key: string, value: T) {
  const arr = map.get(key);
  if (arr) arr.push(value);
  else map.set(key, [value]);
}

export function validateSheet(
  headerRow: unknown[],
  dataRows: unknown[][],
  config: Config = DEFAULT_CONFIG
): ValidationReport {
  const empty = (overrides: Partial<ValidationReport> = {}): ValidationReport => ({
    missingColumns: [],
    extraColumns: [],
    totalRows: 0,
    cleanCount: 0,
    reviewCount: 0,
    skippedCount: 0,
    rows: [],
    colIndex: {} as Record<FieldKey, number>,
    ...overrides,
  });

  if (!dataRows || dataRows.length === 0) {
    return empty({ fatalError: "No rows found" });
  }

  const { colIndex, missing, extra } = mapColumns(headerRow);
  if (missing.length > 0) {
    return empty({
      fatalError: `Missing required column${missing.length > 1 ? "s" : ""}: ${missing.join(", ")}`,
      missingColumns: missing,
      extraColumns: extra,
      colIndex,
    });
  }

  const rows: RowResult[] = [];
  let skippedCount = 0;

  dataRows.forEach((rawRow, i) => {
    const rowNumber = i + 2; // header is row 1
    const isFullyBlank = rawRow.every((c) => cleanText(c) === "");
    if (isFullyBlank) {
      skippedCount++;
      return;
    }

    const rawBranch = getCell(rawRow, colIndex.branch);
    const rawName = getCell(rawRow, colIndex.customer_name);
    const rawAmount = getCell(rawRow, colIndex.loan_amount);
    const rawDate = getCell(rawRow, colIndex.loan_date);
    const rawPhone = getCell(rawRow, colIndex.phone);
    const rawStatus = getCell(rawRow, colIndex.status);

    const branch = cleanText(rawBranch);
    const customer_name = cleanText(rawName);
    const amount = parseAmount(rawAmount);
    const date = parseLoanDate(rawDate);
    const phone = cleanPhone(rawPhone);
    const status = mapStatus(rawStatus);

    const row: RowResult = {
      rowNumber,
      raw: {
        branch: rawBranch,
        customer_name: rawName,
        loan_amount: rawAmount,
        loan_date: rawDate,
        phone: rawPhone,
        status: rawStatus,
      },
      cleaned: {
        branch,
        customer_name,
        loan_amount: amount.value,
        loan_amount_raw: amount.raw,
        loan_date: date.date,
        loan_date_raw: date.raw,
        phone: phone.digits,
        phone_raw: phone.raw,
        status: status.value,
        status_raw: status.raw,
      },
      flags: [],
      suggestions: [],
      severity: "none",
      skipped: false,
      approved: false,
    };

    const sev = (field: FieldKey) => FIELD_SEVERITY[field];

    // ---- Cell-level checks ----
    // Missing = leave blank, flag it. Invalid = keep the original, flag it.
    // Never filled in silently either way.

    // branch
    if (!branch) {
      addFlag(row, "branch", "missing", sev("branch"), "Branch is missing");
    } else if (
      config.knownBranches.length > 0 &&
      !config.knownBranches.some((b) => b.toLowerCase() === branch.toLowerCase())
    ) {
      addFlag(row, "branch", "invalid", sev("branch"), `Branch '${branch}' not in known branch list`);
    }

    // customer_name
    if (!customer_name) {
      addFlag(row, "customer_name", "missing", sev("customer_name"), "Customer name is missing");
    } else {
      if (/\d/.test(customer_name) || /[^a-zA-Z\s.'-]/.test(customer_name)) {
        addFlag(
          row,
          "customer_name",
          "invalid",
          sev("customer_name"),
          `Customer name '${customer_name}' is invalid (contains digits or unusual characters)`
        );
      }
      if (customer_name.length === 1 || JUNK_NAMES.has(customer_name.toLowerCase())) {
        addFlag(
          row,
          "customer_name",
          "invalid",
          sev("customer_name"),
          `Customer name '${customer_name}' is invalid (e.g. 'NA', single character)`
        );
      }
    }

    // loan_amount
    if (amount.raw === "") {
      addFlag(row, "loan_amount", "missing", sev("loan_amount"), "Loan amount is missing");
    } else if (amount.value === null) {
      addFlag(row, "loan_amount", "invalid", sev("loan_amount"), `Loan amount '${amount.raw}' is invalid`);
    } else {
      if (amount.value <= 0)
        addFlag(
          row,
          "loan_amount",
          "invalid",
          sev("loan_amount"),
          `Loan amount '${amount.raw}' is invalid (zero or negative)`
        );
      if (amount.value > config.maxLoanAmount)
        addFlag(
          row,
          "loan_amount",
          "invalid",
          sev("loan_amount"),
          `Loan amount '${amount.raw}' exceeds ₹${config.maxLoanAmount.toLocaleString("en-IN")} limit`
        );
      else if (amount.value > 0 && amount.value < config.minLoanAmount)
        addFlag(
          row,
          "loan_amount",
          "invalid",
          sev("loan_amount"),
          `Loan amount '${amount.raw}' is unrealistically small (< ₹${config.minLoanAmount.toLocaleString("en-IN")})`
        );
    }

    // loan_date
    if (date.raw === "") {
      addFlag(row, "loan_date", "missing", sev("loan_date"), "Loan date is missing");
    } else if (date.invalidFormat || !date.date) {
      addFlag(
        row,
        "loan_date",
        "invalid",
        sev("loan_date"),
        `Loan date '${date.raw}' is invalid (unrecognized format)`
      );
    } else {
      if (date.ambiguous)
        addFlag(row, "loan_date", "invalid", sev("loan_date"), `Loan date '${date.raw}' is ambiguous (day/month unclear)`);
      const today = new Date();
      today.setHours(23, 59, 59, 999);
      if (date.date > today)
        addFlag(row, "loan_date", "invalid", sev("loan_date"), `Loan date '${date.raw}' is in the future`);
      if (date.date.getFullYear() < config.minYear)
        addFlag(row, "loan_date", "invalid", sev("loan_date"), `Loan date '${date.raw}' is before ${config.minYear}`);
    }

    // phone
    if (phone.raw === "") {
      addFlag(row, "phone", "missing", sev("phone"), "Phone is missing");
    } else if (phone.digits.length !== 10) {
      addFlag(
        row,
        "phone",
        "invalid",
        sev("phone"),
        `Phone '${phone.raw}' is invalid (${phone.digits.length} digits)`
      );
    } else {
      if (!/^[6789]/.test(phone.digits))
        addFlag(row, "phone", "invalid", sev("phone"), `Phone '${phone.raw}' is invalid (must start with 6-9)`);
      if (isFakePhone(phone.digits))
        addFlag(row, "phone", "invalid", sev("phone"), `Phone '${phone.raw}' looks fake (repeated/sequential digits)`);
    }

    // status
    if (status.raw === "") {
      addFlag(row, "status", "missing", sev("status"), "Status is missing");
    } else if (!status.value) {
      addFlag(
        row,
        "status",
        "invalid",
        sev("status"),
        `Status '${status.raw}' doesn't map to Active, Closed or NPA`
      );
    }

    // ---- Row-level checks ----

    const blanks = {
      branch: !branch,
      customer_name: !customer_name,
      loan_amount: amount.raw === "",
      loan_date: date.raw === "",
      phone: phone.raw === "",
      status: status.raw === "",
    };
    const blankFieldCount = Object.values(blanks).filter(Boolean).length;

    if (blankFieldCount >= 3) {
      // Too much is missing for a field-by-field review to make sense — send
      // the whole row back to the source instead of listing each gap.
      row.flags = row.flags.filter((f) => f.kind !== "missing");
      addFlag(row, "row", "missing", "high", "Incomplete row: check source register");
    }

    if (status.value === "Closed" && date.date && !date.invalidFormat) {
      const diffDays = (Date.now() - date.date.getTime()) / 86_400_000;
      if (diffDays >= 0 && diffDays <= config.recentClosedDays) {
        addFlag(
          row,
          "row",
          "info",
          "low",
          `Status is Closed but loan date is only ${Math.round(diffDays)} day(s) old — please verify`
        );
      }
    }

    // swapped-field heuristic: amount column holds something that looks like a phone number
    const amountDigitsOnly = amount.raw.replace(/\D/g, "");
    if (amountDigitsOnly.length === 10 && /^[6789]/.test(amountDigitsOnly) && amount.value !== null) {
      addFlag(row, "row", "invalid", "high", "Loan amount looks like a phone number — check for swapped fields");
    }
    // swapped-field heuristic: phone column holds a currency-looking value
    if (phone.raw && /[₹,]|rs\.?/i.test(phone.raw) && phone.digits.length !== 10) {
      addFlag(row, "row", "invalid", "high", "Phone looks like a loan amount — check for swapped fields");
    }

    rows.push(row);
  });

  // ---- Cross-row checks ----

  const exactMap = new Map<string, RowResult[]>();
  const phoneMap = new Map<string, RowResult[]>();
  const nameAmountMap = new Map<string, RowResult[]>();
  const nameDateMap = new Map<string, RowResult[]>();
  const sansPhoneMap = new Map<string, RowResult[]>(); // used for the phone suggestion below

  for (const row of rows) {
    const c = row.cleaned;
    const exactSig = groupKey([
      c.branch.toLowerCase(),
      c.customer_name.toLowerCase(),
      c.loan_amount ?? "",
      c.loan_date ? c.loan_date.toISOString().slice(0, 10) : "",
      c.phone,
      c.status,
    ]);
    pushTo(exactMap, exactSig, row);

    if (c.phone) pushTo(phoneMap, c.phone, row);
    if (c.customer_name && c.loan_amount !== null)
      pushTo(nameAmountMap, groupKey([c.customer_name.toLowerCase(), c.loan_amount]), row);
    if (c.customer_name && c.loan_date)
      pushTo(
        nameDateMap,
        groupKey([c.customer_name.toLowerCase(), c.loan_date.toISOString().slice(0, 10)]),
        row
      );

    const sansPhoneSig = groupKey([
      c.branch.toLowerCase(),
      c.customer_name.toLowerCase(),
      c.loan_amount ?? "",
      c.loan_date ? c.loan_date.toISOString().slice(0, 10) : "",
      c.status,
    ]);
    pushTo(sansPhoneMap, sansPhoneSig, row);
  }

  function addGroupFlags(
    map: Map<string, RowResult[]>,
    messageFn: (other: RowResult) => string,
    filterFn?: (row: RowResult, other: RowResult) => boolean
  ) {
    for (const group of map.values()) {
      if (group.length < 2) continue;
      for (const row of group) {
        const others = group.filter((x) => x !== row);
        const candidates = filterFn ? others.filter((o) => filterFn(row, o)) : others;
        if (candidates.length === 0) continue;
        const other = candidates.reduce((a, b) => (a.rowNumber < b.rowNumber ? a : b));
        addFlag(row, "cross-row", "info", "medium", messageFn(other));
      }
    }
  }

  addGroupFlags(exactMap, (other) => `Exact duplicate of row ${other.rowNumber}`);
  addGroupFlags(
    phoneMap,
    (other) => `Same phone as row ${other.rowNumber}, different customer?`,
    (row, other) => other.cleaned.customer_name.toLowerCase() !== row.cleaned.customer_name.toLowerCase()
  );
  addGroupFlags(nameAmountMap, (other) => `Possible duplicate of row ${other.rowNumber}`);
  addGroupFlags(nameDateMap, (other) => `Possible double entry of row ${other.rowNumber}`);

  // ---- Suggestions (evidence from the file itself; never auto-applied) ----

  const nonBlankBranches = rows.map((r) => r.cleaned.branch).filter(Boolean);
  let dominantBranch: string | null = null;
  let dominantShare = 0;
  if (nonBlankBranches.length > 0) {
    const counts = new Map<string, number>();
    for (const b of nonBlankBranches) counts.set(b, (counts.get(b) ?? 0) + 1);
    for (const [b, c] of counts) {
      const share = c / nonBlankBranches.length;
      if (share > dominantShare) {
        dominantShare = share;
        dominantBranch = b;
      }
    }
    if (dominantShare < config.branchSuggestionThreshold) dominantBranch = null;
  }

  for (const row of rows) {
    const c = row.cleaned;

    if (!c.branch && dominantBranch) {
      row.suggestions.push({
        field: "branch",
        value: dominantBranch,
        message: `${Math.round(dominantShare * 100)}% of rows use "${dominantBranch}" — suggest "${dominantBranch}"?`,
      });
    }

    if (!c.phone) {
      const sig = groupKey([
        c.branch.toLowerCase(),
        c.customer_name.toLowerCase(),
        c.loan_amount ?? "",
        c.loan_date ? c.loan_date.toISOString().slice(0, 10) : "",
        c.status,
      ]);
      const group = sansPhoneMap.get(sig) ?? [];
      const withPhone = group.find((o) => o !== row && o.cleaned.phone);
      if (withPhone) {
        row.suggestions.push({
          field: "phone",
          value: withPhone.cleaned.phone,
          message: `Row ${withPhone.rowNumber} is otherwise identical and has phone ${withPhone.cleaned.phone} — use it?`,
        });
      }
    }

    // Deliberately no suggestion for status: a recent date doesn't tell us
    // whether the loan is Active, Closed or NPA — that would be a guess.
  }

  for (const row of rows) {
    const maxRank = row.flags.reduce((acc, f) => Math.max(acc, severityRank(f.severity)), 0);
    row.severity = maxRank === 0 ? "none" : maxRank === 3 ? "high" : maxRank === 2 ? "medium" : "low";
  }

  const cleanCount = rows.filter((r) => r.approved || r.flags.length === 0).length;
  const reviewCount = rows.filter((r) => !r.approved && r.flags.length > 0).length;

  return {
    missingColumns: [],
    extraColumns: extra,
    totalRows: dataRows.length,
    cleanCount,
    reviewCount,
    skippedCount,
    rows,
    colIndex,
  };
}

// ---- Display / canonical value helpers ----
// The single source of truth for "what does this cell actually show": used by
// both the UI table and the export, so what staff see is what they get.

export function formatAmount(v: number | null): string {
  if (v === null) return "";
  return `₹${v.toLocaleString("en-IN")}`;
}

export function formatDate(d: Date | null): string {
  if (!d) return "";
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${dd}/${mm}/${d.getFullYear()}`;
}

export function resolveCell(
  row: RowResult,
  field: FieldKey
): { value: string; state: "missing" | "invalid" | "ok" } {
  const isInvalid = (f: FieldKey) => row.flags.some((fl) => fl.field === f && fl.kind === "invalid");
  switch (field) {
    case "branch":
      if (!row.cleaned.branch) return { value: "", state: "missing" };
      return { value: row.cleaned.branch, state: isInvalid("branch") ? "invalid" : "ok" };
    case "customer_name":
      if (!row.cleaned.customer_name) return { value: "", state: "missing" };
      return { value: row.cleaned.customer_name, state: isInvalid("customer_name") ? "invalid" : "ok" };
    case "loan_amount":
      if (row.cleaned.loan_amount_raw === "") return { value: "", state: "missing" };
      if (row.cleaned.loan_amount !== null)
        return { value: formatAmount(row.cleaned.loan_amount), state: isInvalid("loan_amount") ? "invalid" : "ok" };
      return { value: row.cleaned.loan_amount_raw, state: "invalid" };
    case "loan_date":
      if (row.cleaned.loan_date_raw === "") return { value: "", state: "missing" };
      if (row.cleaned.loan_date !== null)
        return { value: formatDate(row.cleaned.loan_date), state: isInvalid("loan_date") ? "invalid" : "ok" };
      return { value: row.cleaned.loan_date_raw, state: "invalid" };
    case "phone":
      if (row.cleaned.phone_raw === "") return { value: "", state: "missing" };
      return { value: row.cleaned.phone || row.cleaned.phone_raw, state: isInvalid("phone") ? "invalid" : "ok" };
    case "status":
      if (row.cleaned.status_raw === "") return { value: "", state: "missing" };
      return { value: row.cleaned.status || row.cleaned.status_raw, state: isInvalid("status") ? "invalid" : "ok" };
  }
}

// The value to write back into the sheet when reformatting a cell that is
// already valid (never called for missing/invalid cells — those are left untouched).
function canonicalValue(row: RowResult, field: FieldKey): unknown {
  switch (field) {
    case "branch":
      return row.cleaned.branch;
    case "customer_name":
      return row.cleaned.customer_name;
    case "loan_amount":
      return row.cleaned.loan_amount;
    case "loan_date":
      return row.cleaned.loan_date;
    case "phone":
      return row.cleaned.phone;
    case "status":
      return row.cleaned.status;
  }
}

// ---- Auto-normalize ----
// Unlike the review workflow above, this runs without per-row human approval.
// Its scope is deliberately narrow — only actions backed by hard evidence:
//   - exact duplicate rows (every field identical) -> drop the extras
//   - same record split across two rows because only one has a phone number
//     -> merge the phone into the surviving row
//   - fully blank rows -> drop
//   - cells that are already valid -> rewritten in a consistent format
// It never touches missing/invalid values, and never resolves "possible"
// duplicates (same name+amount, same name+date, same phone/different name) —
// those signals aren't strong enough to act on without a person looking.
// Every action taken is recorded in `actions` so it can be shown as a log.

export interface NormalizationAction {
  type: "delete-duplicate" | "delete-blank" | "merge" | "format" | "note";
  rowNumber: number; // original row number this action concerns (0 = whole-file note)
  relatedRowNumber?: number; // the other row involved, if any
  message: string;
}

export interface NormalizeResult {
  headerRow: unknown[];
  dataRows: unknown[][];
  actions: NormalizationAction[];
}

export function autoNormalizeTable(
  headerRow: unknown[],
  dataRows: unknown[][],
  config: Config = DEFAULT_CONFIG
): NormalizeResult {
  const actions: NormalizationAction[] = [];
  const report = validateSheet(headerRow, dataRows, config);

  if (report.fatalError) {
    return {
      headerRow,
      dataRows,
      actions: [{ type: "note", rowNumber: 0, message: `Skipped — ${report.fatalError}` }],
    };
  }

  dataRows.forEach((r, i) => {
    const rowNumber = i + 2;
    const isBlank = r.every((c) => cleanText(c) === "");
    if (isBlank) {
      actions.push({ type: "delete-blank", rowNumber, message: `Row ${rowNumber} removed — fully blank` });
    }
  });

  const toDelete = new Set<number>();
  const phoneFill = new Map<number, string>();

  // Exact duplicates: keep the earliest row, drop the rest.
  const exactGroups = new Map<string, RowResult[]>();
  for (const row of report.rows) {
    const c = row.cleaned;
    const sig = groupKey([
      c.branch.toLowerCase(),
      c.customer_name.toLowerCase(),
      c.loan_amount ?? "",
      c.loan_date ? c.loan_date.toISOString().slice(0, 10) : "",
      c.phone,
      c.status,
    ]);
    pushTo(exactGroups, sig, row);
  }
  for (const group of exactGroups.values()) {
    if (group.length < 2) continue;
    const sorted = [...group].sort((a, b) => a.rowNumber - b.rowNumber);
    const keep = sorted[0];
    for (const dup of sorted.slice(1)) {
      toDelete.add(dup.rowNumber);
      actions.push({
        type: "delete-duplicate",
        rowNumber: dup.rowNumber,
        relatedRowNumber: keep.rowNumber,
        message: `Row ${dup.rowNumber} removed — exact duplicate of row ${keep.rowNumber}`,
      });
    }
  }

  // Same record (branch, name, amount, date, status all match), phone present
  // on only some copies: merge the phone into the earliest row, drop the rest.
  const sansPhoneGroups = new Map<string, RowResult[]>();
  for (const row of report.rows) {
    if (toDelete.has(row.rowNumber)) continue;
    const c = row.cleaned;
    const sig = groupKey([
      c.branch.toLowerCase(),
      c.customer_name.toLowerCase(),
      c.loan_amount ?? "",
      c.loan_date ? c.loan_date.toISOString().slice(0, 10) : "",
      c.status,
    ]);
    pushTo(sansPhoneGroups, sig, row);
  }
  for (const group of sansPhoneGroups.values()) {
    if (group.length < 2) continue;
    const withPhone = group.filter((r) => r.cleaned.phone);
    const withoutPhone = group.filter((r) => !r.cleaned.phone);
    const distinctPhones = new Set(withPhone.map((r) => r.cleaned.phone));
    // Only act when there's exactly one candidate phone number — if copies
    // disagree on the phone, that's a conflict, not a safe merge.
    if (distinctPhones.size === 1 && withoutPhone.length > 0) {
      const sorted = [...group].sort((a, b) => a.rowNumber - b.rowNumber);
      const survivor = sorted[0];
      const phoneValue = [...distinctPhones][0];
      if (!survivor.cleaned.phone) phoneFill.set(survivor.rowNumber, phoneValue);
      for (const r of sorted.slice(1)) {
        toDelete.add(r.rowNumber);
        actions.push({
          type: "merge",
          rowNumber: r.rowNumber,
          relatedRowNumber: survivor.rowNumber,
          message: `Row ${r.rowNumber} merged into row ${survivor.rowNumber} — phone consolidated, row ${r.rowNumber} removed`,
        });
      }
    }
  }

  // Rebuild the matrix: drop deleted/blank rows, reformat already-valid cells
  // into a consistent representation. Missing and invalid cells are left
  // exactly as they were — normalization never guesses.
  const fieldByColIdx = new Map<number, FieldKey>();
  (Object.keys(report.colIndex) as FieldKey[]).forEach((k) => {
    if (report.colIndex[k] >= 0) fieldByColIdx.set(report.colIndex[k], k);
  });

  let formatCount = 0;
  const newDataRows: unknown[][] = [];
  for (const row of report.rows) {
    if (toDelete.has(row.rowNumber)) continue;
    const dataIdx = row.rowNumber - 2;
    const next = [...(dataRows[dataIdx] ?? [])];

    const fillPhone = phoneFill.get(row.rowNumber);
    if (fillPhone !== undefined && report.colIndex.phone >= 0) {
      next[report.colIndex.phone] = fillPhone;
    }

    for (const [colIdx, field] of fieldByColIdx) {
      const cell = resolveCell(row, field);
      if (cell.state !== "ok") continue; // never touch invalid or missing values
      const canonical = canonicalValue(row, field);
      if (canonical !== null && canonical !== undefined && next[colIdx] !== canonical) {
        next[colIdx] = canonical;
        formatCount++;
      }
    }

    newDataRows.push(next);
  }

  if (formatCount > 0) {
    actions.push({
      type: "format",
      rowNumber: 0,
      message: `Reformatted ${formatCount} valid cell(s) to a consistent format (dates, phone digits, amount numbers, status casing)`,
    });
  }

  return {
    headerRow: headerRow.map((h) => String(h ?? "").trim()),
    dataRows: newDataRows,
    actions,
  };
}
