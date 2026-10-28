"use client";

import { useCallback, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  validateSheet,
  autoNormalizeTable,
  severityRank,
  resolveCell,
  type FieldKey,
  type RowResult,
  type ValidationReport,
  type NormalizationAction,
} from "../lib/validation";

const EDITABLE_FIELDS: { key: FieldKey; label: string }[] = [
  { key: "branch", label: "Branch" },
  { key: "customer_name", label: "Customer Name" },
  { key: "loan_amount", label: "Loan Amount" },
  { key: "loan_date", label: "Loan Date" },
  { key: "phone", label: "Phone" },
  { key: "status", label: "Status" },
];

function flagClasses(kind: string): string {
  switch (kind) {
    case "missing":
      return "bg-slate-100 text-slate-700 dark:bg-slate-800 dark:text-slate-300";
    case "invalid":
      return "bg-red-100 text-red-700 dark:bg-red-950/40 dark:text-red-400";
    default:
      return "bg-purple-100 text-purple-700 dark:bg-purple-950/40 dark:text-purple-400";
  }
}

function severityBadgeClasses(sev: string): string {
  switch (sev) {
    case "high":
      return "bg-red-100 text-red-700 dark:bg-red-950/40 dark:text-red-400";
    case "medium":
      return "bg-amber-100 text-amber-700 dark:bg-amber-950/40 dark:text-amber-400";
    default:
      return "bg-gray-100 text-gray-600 dark:bg-gray-800 dark:text-gray-400";
  }
}

export default function DataChecker() {
  const [fileName, setFileName] = useState<string | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [activeSheet, setActiveSheet] = useState(0);
  const [headerRow, setHeaderRow] = useState<unknown[]>([]);
  const [dataRows, setDataRows] = useState<unknown[][]>([]);
  const [approvedRows, setApprovedRows] = useState<Set<number>>(new Set());
  const [editingRow, setEditingRow] = useState<number | null>(null);
  const [fileError, setFileError] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [filterReviewOnly, setFilterReviewOnly] = useState(false);
  const [normalizationLog, setNormalizationLog] = useState<NormalizationAction[] | null>(null);
  const workbookRef = useRef<XLSX.WorkBook | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const loadSheet = useCallback((workbook: XLSX.WorkBook, sheetIndex: number) => {
    const name = workbook.SheetNames[sheetIndex];
    const ws = workbook.Sheets[name];
    const all = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      defval: "",
      raw: true,
    }) as unknown[][];
    setHeaderRow(all[0] ?? []);
    setDataRows(all.slice(1));
    setApprovedRows(new Set());
    setEditingRow(null);
    setNormalizationLog(null);
  }, []);

  const parseFile = useCallback(
    (file: File) => {
      setFileError(null);
      if (!/\.(xlsx|xls|csv)$/i.test(file.name)) {
        setFileError("Please upload an Excel file (.xlsx, .xls) or .csv");
        return;
      }

      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = e.target?.result;
          const workbook = XLSX.read(data, { type: "binary", cellDates: true });
          if (workbook.SheetNames.length === 0) {
            setFileError("No rows found");
            return;
          }
          workbookRef.current = workbook;
          setSheetNames(workbook.SheetNames);
          setActiveSheet(0);
          setFileName(file.name);
          loadSheet(workbook, 0);
        } catch {
          setFileError("Please upload a valid Excel file");
        }
      };
      reader.onerror = () => setFileError("Failed to read the file.");
      reader.readAsBinaryString(file);
    },
    [loadSheet]
  );

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) parseFile(file);
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) parseFile(file);
  };

  const switchSheet = (idx: number) => {
    setActiveSheet(idx);
    if (workbookRef.current) loadSheet(workbookRef.current, idx);
  };

  const reset = () => {
    setFileName(null);
    setSheetNames([]);
    setHeaderRow([]);
    setDataRows([]);
    setApprovedRows(new Set());
    setFileError(null);
    setNormalizationLog(null);
    workbookRef.current = null;
    if (inputRef.current) inputRef.current.value = "";
  };

  const report: ValidationReport = useMemo(
    () => validateSheet(headerRow, dataRows),
    [headerRow, dataRows]
  );

  const rowsWithApproval: RowResult[] = useMemo(() => {
    return report.rows.map((r) => ({ ...r, approved: approvedRows.has(r.rowNumber) }));
  }, [report.rows, approvedRows]);

  const cleanCount = rowsWithApproval.filter((r) => r.approved || r.flags.length === 0).length;
  const reviewCount = rowsWithApproval.filter((r) => !r.approved && r.flags.length > 0).length;

  const toggleApprove = (rowNumber: number) => {
    setApprovedRows((prev) => {
      const next = new Set(prev);
      if (next.has(rowNumber)) next.delete(rowNumber);
      else next.add(rowNumber);
      return next;
    });
  };

  const updateCell = (rowNumber: number, field: FieldKey, value: string) => {
    const colIdx = report.colIndex[field];
    if (colIdx === undefined || colIdx < 0) return;
    setDataRows((prev) => {
      const next = prev.map((r) => [...r]);
      const dataIdx = rowNumber - 2;
      if (next[dataIdx]) next[dataIdx][colIdx] = value;
      return next;
    });
  };

  const runAutoNormalize = () => {
    const result = autoNormalizeTable(headerRow, dataRows);
    setHeaderRow(result.headerRow);
    setDataRows(result.dataRows);
    // Row numbers shift once rows are dropped/merged, so old approvals/edits no longer line up.
    setApprovedRows(new Set());
    setEditingRow(null);
    setNormalizationLog(result.actions);
  };

  const downloadResults = () => {
    const fieldByColIdx = new Map<number, FieldKey>();
    (Object.keys(report.colIndex) as FieldKey[]).forEach((k) => {
      if (report.colIndex[k] >= 0) fieldByColIdx.set(report.colIndex[k], k);
    });

    const outHeader = [...headerRow.map(String), "Severity", "Validation Status", "Reasons", "Suggestions (not applied)"];
    const outRows = rowsWithApproval.map((r) => {
      const dataIdx = r.rowNumber - 2;
      // Canonical columns are exported as the resolved value — cleaned if
      // valid, the original kept as-is if invalid, blank if missing. Extra
      // columns pass through untouched.
      const base = headerRow.map((_, colIdx) => {
        const field = fieldByColIdx.get(colIdx);
        if (field) return resolveCell(r, field).value;
        return dataRows[dataIdx]?.[colIdx] ?? "";
      });
      const statusLabel = r.approved
        ? r.flags.length > 0
          ? "Approved"
          : "Clean"
        : r.flags.length > 0
        ? "Needs review"
        : "Clean";
      const severityLabel = r.severity === "none" ? "" : r.severity[0].toUpperCase() + r.severity.slice(1);
      return [
        ...base,
        severityLabel,
        statusLabel,
        r.flags.map((f) => f.message).join("; "),
        r.suggestions.map((s) => s.message).join("; "),
      ];
    });
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([outHeader, ...outRows]);
    XLSX.utils.book_append_sheet(wb, ws, "Validated");
    XLSX.writeFile(wb, `validated-${fileName ?? "data"}.xlsx`);
  };

  const sortedRows = useMemo(() => {
    const base = filterReviewOnly
      ? rowsWithApproval.filter((r) => !r.approved && r.flags.length > 0)
      : rowsWithApproval;
    return [...base].sort((a, b) => {
      const aReview = !a.approved && a.flags.length > 0;
      const bReview = !b.approved && b.flags.length > 0;
      if (aReview !== bReview) return aReview ? -1 : 1;
      if (aReview && bReview) {
        const diff = severityRank(b.severity) - severityRank(a.severity);
        if (diff !== 0) return diff;
      }
      return a.rowNumber - b.rowNumber;
    });
  }, [rowsWithApproval, filterReviewOnly]);

  return (
    <div className="w-full max-w-6xl mx-auto flex flex-col gap-6">
      {!fileName && (
        <div
          onDragOver={(e) => {
            e.preventDefault();
            setIsDragging(true);
          }}
          onDragLeave={() => setIsDragging(false)}
          onDrop={handleDrop}
          onClick={() => inputRef.current?.click()}
          className={`cursor-pointer border-2 border-dashed rounded-xl p-10 text-center transition-colors ${
            isDragging
              ? "border-blue-500 bg-blue-50 dark:bg-blue-950/30"
              : "border-gray-300 dark:border-gray-700 hover:border-gray-400 dark:hover:border-gray-600"
          }`}
        >
          <input
            ref={inputRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={handleFileChange}
            className="hidden"
          />
          <div className="flex flex-col items-center gap-2">
            <svg
              xmlns="http://www.w3.org/2000/svg"
              className="w-10 h-10 text-gray-400"
              fill="none"
              viewBox="0 0 24 24"
              stroke="currentColor"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth={1.5}
                d="M3 16.5v2.25A2.25 2.25 0 0 0 5.25 21h13.5A2.25 2.25 0 0 0 21 18.75V16.5M16.5 12 12 16.5m0 0L7.5 12m4.5 4.5V3"
              />
            </svg>
            <p className="font-medium">Drag & drop a loan data file here, or click to browse</p>
            <p className="text-sm text-gray-500">Supports .xlsx, .xls, .csv</p>
          </div>
        </div>
      )}

      {fileError && (
        <div className="text-red-600 bg-red-50 dark:bg-red-950/30 border border-red-200 dark:border-red-900 rounded-lg px-4 py-2 text-sm">
          {fileError}
        </div>
      )}

      {fileName && (
        <div className="flex items-center justify-between flex-wrap gap-2">
          <span className="text-sm text-gray-600 dark:text-gray-400">
            Loaded <span className="font-medium">{fileName}</span>
          </span>
          <div className="flex items-center gap-3">
            {sheetNames.length > 1 && (
              <select
                value={activeSheet}
                onChange={(e) => switchSheet(Number(e.target.value))}
                className="text-sm border border-gray-300 dark:border-gray-700 rounded px-2 py-1 bg-white dark:bg-gray-900"
              >
                {sheetNames.map((name, i) => (
                  <option key={name} value={i}>
                    {name}
                  </option>
                ))}
              </select>
            )}
            <button onClick={reset} className="text-sm text-blue-600 hover:underline">
              Upload a different file
            </button>
          </div>
        </div>
      )}

      {fileName && report.fatalError && (
        <div className="text-red-700 bg-red-50 dark:bg-red-950/30 border border-red-200 dark:border-red-900 rounded-lg px-4 py-3 text-sm">
          <p className="font-medium">{report.fatalError}</p>
          {report.missingColumns.length > 0 && (
            <p className="mt-1">
              Add a column for: {report.missingColumns.join(", ")}
            </p>
          )}
        </div>
      )}

      {fileName && !report.fatalError && (
        <>
          <div className="flex flex-wrap items-center gap-4 bg-gray-50 dark:bg-gray-900 border border-gray-200 dark:border-gray-800 rounded-lg px-5 py-4 text-sm">
            <span>
              Total rows: <span className="font-semibold">{report.totalRows}</span>
            </span>
            <span className="text-green-700 dark:text-green-400">
              Clean: <span className="font-semibold">{cleanCount}</span>
            </span>
            <span className="text-amber-700 dark:text-amber-400">
              Needs review: <span className="font-semibold">{reviewCount}</span>
            </span>
            <span className="text-gray-500">
              Skipped (empty): <span className="font-semibold">{report.skippedCount}</span>
            </span>
            {report.extraColumns.length > 0 && (
              <span className="text-gray-500">
                Ignored extra column{report.extraColumns.length > 1 ? "s" : ""}:{" "}
                {report.extraColumns.join(", ")}
              </span>
            )}
            <div className="ml-auto flex items-center gap-3">
              <label className="flex items-center gap-1.5 text-gray-600 dark:text-gray-400 cursor-pointer">
                <input
                  type="checkbox"
                  checked={filterReviewOnly}
                  onChange={(e) => setFilterReviewOnly(e.target.checked)}
                />
                Needs review only
              </label>
              <button
                onClick={runAutoNormalize}
                className="px-3 py-1.5 rounded border border-gray-300 dark:border-gray-700 text-sm hover:bg-gray-100 dark:hover:bg-gray-800"
              >
                Auto-normalize table
              </button>
              <button
                onClick={downloadResults}
                className="px-3 py-1.5 rounded bg-blue-600 text-white text-sm hover:bg-blue-700"
              >
                Download cleaned data
              </button>
            </div>
          </div>

          {normalizationLog && (
            <div className="bg-blue-50 dark:bg-blue-950/20 border border-blue-200 dark:border-blue-900 rounded-lg px-4 py-3 text-sm">
              <div className="flex items-center justify-between gap-2">
                <p className="font-medium text-blue-800 dark:text-blue-300">
                  {normalizationLog.length === 0
                    ? "Auto-normalize ran — nothing to change (no duplicates or blank rows found)."
                    : `Auto-normalize ran — ${normalizationLog.length} action(s) taken:`}
                </p>
                <button
                  onClick={() => setNormalizationLog(null)}
                  className="text-xs text-blue-600 dark:text-blue-400 hover:underline shrink-0"
                >
                  Dismiss
                </button>
              </div>
              {normalizationLog.length > 0 && (
                <ul className="mt-2 space-y-1 text-blue-700 dark:text-blue-400 list-disc list-inside">
                  {normalizationLog.map((a, i) => (
                    <li key={i}>{a.message}</li>
                  ))}
                </ul>
              )}
            </div>
          )}

          <p className="text-xs text-gray-500 -mt-2">
            Worst rows show first (sorted by severity). Blank = missing, never filled in. Red text = a value is
            present but invalid, kept as-is. Dashed “Suggest” buttons only apply a value when you click Accept.
            “Auto-normalize” only removes exact duplicates/blank rows and merges rows that differ solely by a
            missing phone — every action is logged above, nothing else is auto-filled or deleted.
          </p>

          <div className="overflow-auto border border-gray-200 dark:border-gray-800 rounded-lg max-h-[32rem]">
            <table className="min-w-full text-sm">
              <thead className="bg-gray-50 dark:bg-gray-900 sticky top-0 z-10">
                <tr>
                  <th className="px-3 py-2 text-left font-semibold border-b border-gray-200 dark:border-gray-800">
                    Row
                  </th>
                  {EDITABLE_FIELDS.map((f) => (
                    <th
                      key={f.key}
                      className="px-3 py-2 text-left font-semibold border-b border-gray-200 dark:border-gray-800 whitespace-nowrap"
                    >
                      {f.label}
                    </th>
                  ))}
                  <th className="px-3 py-2 text-left font-semibold border-b border-gray-200 dark:border-gray-800">
                    Status
                  </th>
                  <th className="px-3 py-2 text-left font-semibold border-b border-gray-200 dark:border-gray-800 min-w-[16rem]">
                    Reasons
                  </th>
                  <th className="px-3 py-2 text-left font-semibold border-b border-gray-200 dark:border-gray-800">
                    Actions
                  </th>
                </tr>
              </thead>
              <tbody>
                {sortedRows.map((r) => {
                  const isEditing = editingRow === r.rowNumber;
                  const needsReview = !r.approved && r.flags.length > 0;
                  return (
                    <tr
                      key={r.rowNumber}
                      className={
                        needsReview
                          ? "bg-amber-50 dark:bg-amber-950/20"
                          : "odd:bg-white even:bg-gray-50 dark:odd:bg-gray-950 dark:even:bg-gray-900"
                      }
                    >
                      <td className="px-3 py-2 border-b border-gray-100 dark:border-gray-900 text-gray-500 align-top">
                        {r.rowNumber}
                      </td>
                      {EDITABLE_FIELDS.map((f) => {
                        const rawVal =
                          f.key === "loan_amount"
                            ? r.cleaned.loan_amount_raw
                            : f.key === "loan_date"
                            ? r.cleaned.loan_date_raw
                            : f.key === "phone"
                            ? r.cleaned.phone_raw
                            : f.key === "status"
                            ? r.cleaned.status_raw
                            : String(r.raw[f.key] ?? "");
                        const cell = resolveCell(r, f.key);
                        const suggestion = r.suggestions.find((s) => s.field === f.key);
                        return (
                          <td
                            key={f.key}
                            className="px-3 py-2 border-b border-gray-100 dark:border-gray-900 whitespace-nowrap align-top"
                          >
                            {isEditing ? (
                              <input
                                defaultValue={rawVal}
                                onBlur={(e) => updateCell(r.rowNumber, f.key, e.target.value)}
                                className="border border-gray-300 dark:border-gray-700 rounded px-1.5 py-0.5 text-sm w-32 bg-white dark:bg-gray-900"
                              />
                            ) : (
                              <div className="flex flex-col gap-1">
                                <span
                                  className={
                                    cell.state === "missing"
                                      ? "text-gray-400 italic"
                                      : cell.state === "invalid"
                                      ? "text-red-600 dark:text-red-400"
                                      : ""
                                  }
                                >
                                  {cell.value || "—"}
                                </span>
                                {suggestion && (
                                  <button
                                    onClick={() => updateCell(r.rowNumber, f.key, suggestion.value)}
                                    title={suggestion.message}
                                    className="text-xs text-blue-600 dark:text-blue-400 border border-dashed border-blue-300 dark:border-blue-800 rounded px-1.5 py-0.5 hover:bg-blue-50 dark:hover:bg-blue-950/30 w-fit text-left"
                                  >
                                    Suggest: {suggestion.value} · Accept
                                  </button>
                                )}
                              </div>
                            )}
                          </td>
                        );
                      })}
                      <td className="px-3 py-2 border-b border-gray-100 dark:border-gray-900 align-top">
                        {r.approved ? (
                          <span className="inline-flex items-center gap-1 text-green-700 dark:text-green-400 font-medium">
                            ✓ Approved
                          </span>
                        ) : r.flags.length === 0 ? (
                          <span className="text-green-700 dark:text-green-400 font-medium">Clean</span>
                        ) : (
                          <span
                            className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium whitespace-nowrap ${severityBadgeClasses(
                              r.severity
                            )}`}
                          >
                            {r.severity === "high" ? "High" : r.severity === "medium" ? "Medium" : "Low"} · Review
                          </span>
                        )}
                      </td>
                      <td className="px-3 py-2 border-b border-gray-100 dark:border-gray-900 align-top">
                        <div className="flex flex-wrap gap-1 max-w-sm">
                          {r.flags.map((f, i) => (
                            <span key={i} className={`text-xs px-1.5 py-0.5 rounded ${flagClasses(f.kind)}`}>
                              {f.message}
                            </span>
                          ))}
                        </div>
                      </td>
                      <td className="px-3 py-2 border-b border-gray-100 dark:border-gray-900 whitespace-nowrap align-top">
                        {isEditing ? (
                          <button
                            onClick={() => setEditingRow(null)}
                            className="text-sm text-blue-600 hover:underline"
                          >
                            Done
                          </button>
                        ) : (
                          <div className="flex gap-3">
                            <button
                              onClick={() => setEditingRow(r.rowNumber)}
                              className="text-sm text-blue-600 hover:underline"
                            >
                              Edit
                            </button>
                            {r.flags.length > 0 && (
                              <button
                                onClick={() => toggleApprove(r.rowNumber)}
                                className="text-sm text-gray-600 dark:text-gray-400 hover:underline"
                              >
                                {r.approved ? "Unapprove" : "Approve"}
                              </button>
                            )}
                          </div>
                        )}
                      </td>
                    </tr>
                  );
                })}
                {sortedRows.length === 0 && (
                  <tr>
                    <td
                      colSpan={EDITABLE_FIELDS.length + 4}
                      className="px-3 py-6 text-center text-gray-400"
                    >
                      No rows to show.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </>
      )}
    </div>
  );
}
