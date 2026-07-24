import React, { useState, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Upload,
  FileSpreadsheet,
  CheckCircle2,
  XCircle,
  AlertTriangle,
  ArrowRightLeft,
  Download,
  Search,
  RotateCcw,
  Settings2,
  Info,
  Plus,
  Minus,
} from "lucide-react";

const C = {
  bg: "#0A0C10",
  panel: "#12151C",
  panelAlt: "#171B24",
  border: "#232733",
  borderSoft: "#1B1F29",
  text: "#E3E6EB",
  muted: "#7C8494",
  dim: "#4B5261",
  green: "#3DDC84",
  greenDim: "#1E4A34",
  red: "#FF5C5C",
  redDim: "#4A2020",
  amber: "#FFB020",
  amberDim: "#4A3714",
  blue: "#4C9EFF",
  blueDim: "#1C3252",
};

const MONO = "ui-monospace, SFMono-Regular, 'JetBrains Mono', Menlo, Consolas, monospace";
const SANS = "'Inter', ui-sans-serif, system-ui, -apple-system, sans-serif";

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        resolve(XLSX.read(e.target.result, { type: "array" }));
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("Could not read file"));
    reader.readAsArrayBuffer(file);
  });
}

function sheetToData(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  if (!ws) return { headers: [], rows: [] };
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });
  const headerRow = raw[0] || [];
  const headers = headerRow.map((h, i) => (h === "" || h == null ? `Column ${i + 1}` : String(h).trim()));
  const rows = raw.slice(1).filter((r) => r.some((c) => String(c ?? "").trim() !== "")).map((r) => {
    const obj = {};
    headers.forEach((h, i) => (obj[h] = r[i] === undefined ? "" : r[i]));
    return obj;
  });
  return { headers, rows };
}

function normalizeVal(v, trimWs, caseSensitive) {
  let s = v === null || v === undefined ? "" : String(v);
  if (trimWs) s = s.trim();
  if (!caseSensitive) s = s.toLowerCase();
  return s;
}

const GUESS_KEYS = ["id", "ID", "Id", "code", "key", "sku", "employee_id", "emp_id", "order_id", "invoice_id", "uid", "reference", "ref no", "ref_no"];

function guessKeyColumn(common) {
  for (const g of GUESS_KEYS) {
    const found = common.find((c) => c.toLowerCase() === g.toLowerCase());
    if (found) return found;
  }
  return common[0] || "";
}

function StatCard({ label, value, accent, sub }) {
  return (
    <div
      style={{
        background: C.panel,
        border: `1px solid ${C.border}`,
        borderTop: `2px solid ${accent}`,
        padding: "14px 16px",
        minWidth: 0,
      }}
    >
      <div style={{ fontFamily: SANS, fontSize: 10.5, letterSpacing: "0.08em", textTransform: "uppercase", color: C.muted, marginBottom: 6 }}>
        {label}
      </div>
      <div style={{ fontFamily: MONO, fontSize: 26, fontWeight: 600, color: C.text, lineHeight: 1 }}>{value}</div>
      {sub && <div style={{ fontFamily: SANS, fontSize: 11.5, color: C.dim, marginTop: 5 }}>{sub}</div>}
    </div>
  );
}

function FileSlot({ label, accent, slot, onFile, onSheetChange }) {
  const [dragOver, setDragOver] = useState(false);
  const inputId = `file-input-${label}`;

  const handleFiles = (files) => {
    if (files && files[0]) onFile(files[0]);
  };

  return (
    <div
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragOver(false);
        handleFiles(e.dataTransfer.files);
      }}
      style={{
        background: C.panel,
        border: `1px dashed ${dragOver ? accent : C.border}`,
        padding: 22,
        transition: "border-color 120ms ease",
        flex: 1,
        minWidth: 260,
      }}
    >
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
        <span style={{ width: 8, height: 8, background: accent, display: "inline-block" }} />
        <span style={{ fontFamily: SANS, fontSize: 11, letterSpacing: "0.1em", textTransform: "uppercase", color: C.muted }}>
          {label}
        </span>
      </div>

      {!slot ? (
        <label htmlFor={inputId} style={{ cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 10, padding: "26px 10px", color: C.dim }}>
          <Upload size={26} color={C.dim} />
          <div style={{ fontFamily: SANS, fontSize: 13, color: C.muted, textAlign: "center" }}>
            Drop a .xlsx / .csv file here
            <br />
            <span style={{ color: C.dim }}>or click to browse</span>
          </div>
          <input
            id={inputId}
            type="file"
            accept=".xlsx,.xls,.csv"
            style={{ display: "none" }}
            onChange={(e) => handleFiles(e.target.files)}
          />
        </label>
      ) : (
        <div>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
            <FileSpreadsheet size={16} color={accent} />
            <span style={{ fontFamily: MONO, fontSize: 13, color: C.text, wordBreak: "break-all" }}>{slot.fileName}</span>
          </div>

          {slot.sheetNames.length > 1 && (
            <div style={{ marginBottom: 10 }}>
              <select
                value={slot.sheetName}
                onChange={(e) => onSheetChange(e.target.value)}
                style={{
                  width: "100%",
                  background: C.panelAlt,
                  border: `1px solid ${C.border}`,
                  color: C.text,
                  fontFamily: MONO,
                  fontSize: 12.5,
                  padding: "6px 8px",
                }}
              >
                {slot.sheetNames.map((s) => (
                  <option key={s} value={s}>
                    sheet: {s}
                  </option>
                ))}
              </select>
            </div>
          )}

          <div style={{ display: "flex", gap: 16, fontFamily: MONO, fontSize: 12.5 }}>
            <div>
              <span style={{ color: C.dim }}>rows </span>
              <span style={{ color: C.text }}>{slot.rows.length}</span>
            </div>
            <div>
              <span style={{ color: C.dim }}>cols </span>
              <span style={{ color: C.text }}>{slot.headers.length}</span>
            </div>
          </div>

          <label htmlFor={inputId} style={{ display: "inline-block", marginTop: 12, fontFamily: SANS, fontSize: 11.5, color: accent, cursor: "pointer", borderBottom: `1px solid ${accent}`, paddingBottom: 1 }}>
            replace file
          </label>
          <input
            id={inputId}
            type="file"
            accept=".xlsx,.xls,.csv"
            style={{ display: "none" }}
            onChange={(e) => handleFiles(e.target.files)}
          />
        </div>
      )}
    </div>
  );
}

function Pill({ text, tone }) {
  const map = {
    green: [C.greenDim, C.green],
    red: [C.redDim, C.red],
    blue: [C.blueDim, C.blue],
    amber: [C.amberDim, C.amber],
  };
  const [bg, fg] = map[tone] || [C.panelAlt, C.muted];
  return (
    <span
      style={{
        display: "inline-block",
        background: bg,
        color: fg,
        border: `1px solid ${fg}33`,
        fontFamily: MONO,
        fontSize: 12,
        padding: "3px 8px",
        margin: "0 6px 6px 0",
      }}
    >
      {text}
    </span>
  );
}

function TabButton({ active, onClick, children, count, tone }) {
  return (
    <button
      onClick={onClick}
      style={{
        background: active ? C.panelAlt : "transparent",
        border: "none",
        borderBottom: `2px solid ${active ? tone || C.blue : "transparent"}`,
        color: active ? C.text : C.muted,
        fontFamily: SANS,
        fontSize: 13,
        padding: "10px 16px",
        cursor: "pointer",
        display: "flex",
        alignItems: "center",
        gap: 8,
      }}
    >
      {children}
      {count !== undefined && (
        <span style={{ fontFamily: MONO, fontSize: 11.5, color: active ? tone || C.blue : C.dim }}>{count}</span>
      )}
    </button>
  );
}

export default function ReportDiff() {
  const [slotA, setSlotA] = useState(null);
  const [slotB, setSlotB] = useState(null);
  const [wbA, setWbA] = useState(null);
  const [wbB, setWbB] = useState(null);

  const [matchMode, setMatchMode] = useState("key"); // 'key' | 'position'
  const [keyColumn, setKeyColumn] = useState("");
  const [caseSensitive, setCaseSensitive] = useState(false);
  const [trimWs, setTrimWs] = useState(true);

  const [activeTab, setActiveTab] = useState("overview");
  const [search, setSearch] = useState("");
  const [error, setError] = useState("");

  const loadFile = useCallback(async (file, which) => {
    setError("");
    try {
      const wb = await readWorkbook(file);
      const sheetName = wb.SheetNames[0];
      const { headers, rows } = sheetToData(wb, sheetName);
      const slot = { fileName: file.name, sheetNames: wb.SheetNames, sheetName, headers, rows };
      if (which === "A") {
        setWbA(wb);
        setSlotA(slot);
      } else {
        setWbB(wb);
        setSlotB(slot);
      }
    } catch (e) {
      setError("Couldn't read that file. Make sure it's a valid .xlsx or .csv export.");
    }
  }, []);

  const changeSheet = (which, sheetName) => {
    const wb = which === "A" ? wbA : wbB;
    if (!wb) return;
    const { headers, rows } = sheetToData(wb, sheetName);
    const update = (prev) => ({ ...prev, sheetName, headers, rows });
    if (which === "A") setSlotA(update(slotA));
    else setSlotB(update(slotB));
  };

  const commonColumns = useMemo(() => {
    if (!slotA || !slotB) return [];
    const setB = new Set(slotB.headers);
    return slotA.headers.filter((h) => setB.has(h));
  }, [slotA, slotB]);

  const onlyColumnsA = useMemo(() => {
    if (!slotA || !slotB) return [];
    const setB = new Set(slotB.headers);
    return slotA.headers.filter((h) => !setB.has(h));
  }, [slotA, slotB]);

  const onlyColumnsB = useMemo(() => {
    if (!slotA || !slotB) return [];
    const setA = new Set(slotA.headers);
    return slotB.headers.filter((h) => !setA.has(h));
  }, [slotA, slotB]);

  // auto-guess key column whenever common columns change
  React.useEffect(() => {
    if (commonColumns.length && !commonColumns.includes(keyColumn)) {
      setKeyColumn(guessKeyColumn(commonColumns));
    }
  }, [commonColumns]); // eslint-disable-line

  const result = useMemo(() => {
    if (!slotA || !slotB || commonColumns.length === 0) return null;

    const compareCols = matchMode === "key" ? commonColumns.filter((c) => c !== keyColumn) : commonColumns;

    if (matchMode === "key" && keyColumn) {
      const buildMap = (rows) => {
        const map = new Map();
        let dupes = 0;
        rows.forEach((r, idx) => {
          const k = normalizeVal(r[keyColumn], trimWs, caseSensitive);
          if (k === "") return;
          if (map.has(k)) dupes++;
          else map.set(k, { row: r, idx });
        });
        return { map, dupes };
      };

      const { map: mapA, dupes: dupesA } = buildMap(slotA.rows);
      const { map: mapB, dupes: dupesB } = buildMap(slotB.rows);

      const onlyA = [];
      const onlyB = [];
      const mismatches = [];
      let perfectMatches = 0;
      let matchedCount = 0;

      for (const [k, entryA] of mapA.entries()) {
        if (!mapB.has(k)) {
          onlyA.push({ key: k, row: entryA.row });
        }
      }
      for (const [k, entryB] of mapB.entries()) {
        if (!mapA.has(k)) {
          onlyB.push({ key: k, row: entryB.row });
        }
      }
      for (const [k, entryA] of mapA.entries()) {
        const entryB = mapB.get(k);
        if (!entryB) continue;
        matchedCount++;
        let rowHasMismatch = false;
        for (const col of compareCols) {
          const vA = entryA.row[col];
          const vB = entryB.row[col];
          if (normalizeVal(vA, trimWs, caseSensitive) !== normalizeVal(vB, trimWs, caseSensitive)) {
            rowHasMismatch = true;
            mismatches.push({ key: k, column: col, valueA: vA, valueB: vB });
          }
        }
        if (!rowHasMismatch) perfectMatches++;
      }

      return {
        mode: "key",
        rowsA: slotA.rows.length,
        rowsB: slotB.rows.length,
        colsA: slotA.headers.length,
        colsB: slotB.headers.length,
        matchedCount,
        perfectMatches,
        mismatchedRowCount: matchedCount - perfectMatches,
        onlyA,
        onlyB,
        mismatches,
        dupesA,
        dupesB,
        compareCols,
      };
    }

    // position mode
    const lenA = slotA.rows.length;
    const lenB = slotB.rows.length;
    const max = Math.max(lenA, lenB);
    const onlyA = [];
    const onlyB = [];
    const mismatches = [];
    let perfectMatches = 0;
    let matchedCount = 0;

    for (let i = 0; i < max; i++) {
      const label = `row ${i + 2}`;
      if (i < lenA && i < lenB) {
        matchedCount++;
        let rowHasMismatch = false;
        for (const col of compareCols) {
          const vA = slotA.rows[i][col];
          const vB = slotB.rows[i][col];
          if (normalizeVal(vA, trimWs, caseSensitive) !== normalizeVal(vB, trimWs, caseSensitive)) {
            rowHasMismatch = true;
            mismatches.push({ key: label, column: col, valueA: vA, valueB: vB });
          }
        }
        if (!rowHasMismatch) perfectMatches++;
      } else if (i < lenA) {
        onlyA.push({ key: label, row: slotA.rows[i] });
      } else {
        onlyB.push({ key: label, row: slotB.rows[i] });
      }
    }

    return {
      mode: "position",
      rowsA: lenA,
      rowsB: lenB,
      colsA: slotA.headers.length,
      colsB: slotB.headers.length,
      matchedCount,
      perfectMatches,
      mismatchedRowCount: matchedCount - perfectMatches,
      onlyA,
      onlyB,
      mismatches,
      dupesA: 0,
      dupesB: 0,
      compareCols,
    };
  }, [slotA, slotB, commonColumns, matchMode, keyColumn, trimWs, caseSensitive]);

  const filteredMismatches = useMemo(() => {
    if (!result) return [];
    if (!search.trim()) return result.mismatches;
    const q = search.trim().toLowerCase();
    return result.mismatches.filter(
      (m) => String(m.key).toLowerCase().includes(q) || String(m.column).toLowerCase().includes(q)
    );
  }, [result, search]);

  const reset = () => {
    setSlotA(null);
    setSlotB(null);
    setWbA(null);
    setWbB(null);
    setKeyColumn("");
    setSearch("");
    setActiveTab("overview");
    setError("");
  };

  const exportCsv = () => {
    if (!result) return;
    const rows = [["key", "column", "value_a", "value_b"]];
    result.mismatches.forEach((m) => rows.push([m.key, m.column, m.valueA, m.valueB]));
    const csv = rows.map((r) => r.map((v) => `"${String(v ?? "").replace(/"/g, '""')}"`).join(",")).join("\n");
    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "mismatches.csv";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const bothLoaded = slotA && slotB;
  const matchRate = result && result.matchedCount > 0 ? Math.round((result.perfectMatches / result.matchedCount) * 100) : null;

  return (
    <div style={{ background: C.bg, minHeight: "100%", color: C.text, fontFamily: SANS, padding: 24 }}>
      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 22, flexWrap: "wrap", gap: 12 }}>
        <div>
          <div style={{ display: "flex", alignItems: "baseline", gap: 10 }}>
            <span style={{ fontFamily: MONO, fontSize: 20, fontWeight: 700, letterSpacing: "-0.02em" }}>recon</span>
            <span style={{ fontFamily: MONO, fontSize: 20, color: C.dim }}>://</span>
            <span style={{ fontFamily: MONO, fontSize: 20, color: C.blue }}>diff</span>
          </div>
          <div style={{ fontSize: 12.5, color: C.muted, marginTop: 3 }}>
            Line up two spreadsheet exports, key-match the rows, and see exactly where they disagree.
          </div>
        </div>
        {bothLoaded && (
          <button
            onClick={reset}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 6,
              background: "transparent",
              border: `1px solid ${C.border}`,
              color: C.muted,
              fontFamily: SANS,
              fontSize: 12.5,
              padding: "7px 12px",
              cursor: "pointer",
            }}
          >
            <RotateCcw size={13} /> start over
          </button>
        )}
      </div>

      {error && (
        <div style={{ background: C.redDim, border: `1px solid ${C.red}55`, color: C.red, padding: "10px 14px", marginBottom: 16, fontSize: 13, fontFamily: SANS, display: "flex", gap: 8, alignItems: "center" }}>
          <AlertTriangle size={15} /> {error}
        </div>
      )}

      {/* Upload zone */}
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 22 }}>
        <FileSlot label="Report A" accent={C.blue} slot={slotA} onFile={(f) => loadFile(f, "A")} onSheetChange={(s) => changeSheet("A", s)} />
        <div style={{ display: "flex", alignItems: "center", justifyContent: "center", color: C.dim, padding: "0 4px" }}>
          <ArrowRightLeft size={18} />
        </div>
        <FileSlot label="Report B" accent={C.amber} slot={slotB} onFile={(f) => loadFile(f, "B")} onSheetChange={(s) => changeSheet("B", s)} />
      </div>

      {!bothLoaded && (
        <div style={{ border: `1px solid ${C.borderSoft}`, background: C.panel, padding: 18, fontSize: 12.5, color: C.muted, display: "flex", gap: 10, alignItems: "flex-start" }}>
          <Info size={15} style={{ marginTop: 1, flexShrink: 0 }} />
          <div>
            Works with any two report exports &mdash; sales, inventory, HR, finance, whatever. Upload both files above.
            Nothing leaves your browser; the comparison runs locally.
          </div>
        </div>
      )}

      {bothLoaded && commonColumns.length === 0 && (
        <div style={{ border: `1px solid ${C.red}55`, background: C.redDim, padding: 18, fontSize: 13, color: C.red }}>
          These two files share no column names at all, so there's nothing to match on. Check that both reports use the
          same headers (row 1), or that you picked the right sheet.
        </div>
      )}

      {bothLoaded && commonColumns.length > 0 && (
        <>
          {/* Config bar */}
          <div style={{ background: C.panel, border: `1px solid ${C.border}`, padding: "14px 18px", marginBottom: 20, display: "flex", gap: 24, flexWrap: "wrap", alignItems: "center" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <Settings2 size={14} color={C.muted} />
              <span style={{ fontSize: 11.5, textTransform: "uppercase", letterSpacing: "0.08em", color: C.muted }}>match by</span>
            </div>

            <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, cursor: "pointer" }}>
              <input type="radio" checked={matchMode === "key"} onChange={() => setMatchMode("key")} /> key column
            </label>
            <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, cursor: "pointer" }}>
              <input type="radio" checked={matchMode === "position"} onChange={() => setMatchMode("position")} /> row position
            </label>

            {matchMode === "key" && (
              <select
                value={keyColumn}
                onChange={(e) => setKeyColumn(e.target.value)}
                style={{ background: C.panelAlt, border: `1px solid ${C.border}`, color: C.text, fontFamily: MONO, fontSize: 12.5, padding: "6px 10px" }}
              >
                {commonColumns.map((c) => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </select>
            )}

            <span style={{ width: 1, alignSelf: "stretch", background: C.border }} />

            <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, cursor: "pointer", color: C.muted }}>
              <input type="checkbox" checked={trimWs} onChange={(e) => setTrimWs(e.target.checked)} /> trim whitespace
            </label>
            <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, cursor: "pointer", color: C.muted }}>
              <input type="checkbox" checked={caseSensitive} onChange={(e) => setCaseSensitive(e.target.checked)} /> case sensitive
            </label>
          </div>

          {result && (result.dupesA > 0 || result.dupesB > 0) && matchMode === "key" && (
            <div style={{ marginBottom: 16, fontSize: 12.5, color: C.amber, display: "flex", gap: 8, alignItems: "center" }}>
              <AlertTriangle size={14} />
              Heads up: {result.dupesA} duplicate key{result.dupesA === 1 ? "" : "s"} in Report A and {result.dupesB} in
              Report B were collapsed to their first occurrence.
            </div>
          )}

          {result && (
            <>
              {/* Overview stat cards */}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 12, marginBottom: 22 }}>
                <StatCard label="Rows · A" value={result.rowsA} accent={C.blue} />
                <StatCard label="Rows · B" value={result.rowsB} accent={C.amber} />
                <StatCard label="Columns matched" value={commonColumns.length} accent={C.green} sub={`of ${result.colsA} / ${result.colsB}`} />
                <StatCard label="Rows matched" value={result.matchedCount} accent={C.text} />
                <StatCard label="Perfect matches" value={result.perfectMatches} accent={C.green} sub={matchRate !== null ? `${matchRate}% of matched rows` : undefined} />
                <StatCard label="Mismatched rows" value={result.mismatchedRowCount} accent={C.amber} />
                <StatCard label="Only in A" value={result.onlyA.length} accent={C.red} />
                <StatCard label="Only in B" value={result.onlyB.length} accent={C.blue} />
              </div>

              {/* Proportion bar */}
              <div style={{ marginBottom: 26 }}>
                <div style={{ display: "flex", height: 10, width: "100%", overflow: "hidden", border: `1px solid ${C.border}` }}>
                  {(() => {
                    const total = result.perfectMatches + result.mismatchedRowCount + result.onlyA.length + result.onlyB.length || 1;
                    const seg = (n, color) => (
                      <div style={{ width: `${(n / total) * 100}%`, background: color, minWidth: n > 0 ? 2 : 0 }} />
                    );
                    return (
                      <>
                        {seg(result.perfectMatches, C.green)}
                        {seg(result.mismatchedRowCount, C.amber)}
                        {seg(result.onlyA.length, C.red)}
                        {seg(result.onlyB.length, C.blue)}
                      </>
                    );
                  })()}
                </div>
                <div style={{ display: "flex", gap: 18, marginTop: 8, fontSize: 11.5, color: C.muted, flexWrap: "wrap" }}>
                  <span><span style={{ color: C.green }}>■</span> perfect match</span>
                  <span><span style={{ color: C.amber }}>■</span> mismatch</span>
                  <span><span style={{ color: C.red }}>■</span> only in A</span>
                  <span><span style={{ color: C.blue }}>■</span> only in B</span>
                </div>
              </div>

              {/* Tabs */}
              <div style={{ display: "flex", borderBottom: `1px solid ${C.border}`, marginBottom: 18 }}>
                <TabButton active={activeTab === "overview"} onClick={() => setActiveTab("overview")}>
                  columns
                </TabButton>
                <TabButton active={activeTab === "onlyA"} onClick={() => setActiveTab("onlyA")} count={result.onlyA.length} tone={C.red}>
                  only in A
                </TabButton>
                <TabButton active={activeTab === "onlyB"} onClick={() => setActiveTab("onlyB")} count={result.onlyB.length} tone={C.blue}>
                  only in B
                </TabButton>
                <TabButton active={activeTab === "mismatches"} onClick={() => setActiveTab("mismatches")} count={result.mismatches.length} tone={C.amber}>
                  mismatches
                </TabButton>
              </div>

              {/* Columns tab */}
              {activeTab === "overview" && (
                <div>
                  <div style={{ marginBottom: 16 }}>
                    <div style={{ fontSize: 11.5, textTransform: "uppercase", letterSpacing: "0.08em", color: C.muted, marginBottom: 8 }}>
                      common columns ({commonColumns.length})
                    </div>
                    {commonColumns.map((c) => (
                      <Pill key={c} text={c} tone="green" />
                    ))}
                  </div>
                  {onlyColumnsA.length > 0 && (
                    <div style={{ marginBottom: 16 }}>
                      <div style={{ fontSize: 11.5, textTransform: "uppercase", letterSpacing: "0.08em", color: C.muted, marginBottom: 8 }}>
                        only in report A ({onlyColumnsA.length})
                      </div>
                      {onlyColumnsA.map((c) => (
                        <Pill key={c} text={c} tone="red" />
                      ))}
                    </div>
                  )}
                  {onlyColumnsB.length > 0 && (
                    <div>
                      <div style={{ fontSize: 11.5, textTransform: "uppercase", letterSpacing: "0.08em", color: C.muted, marginBottom: 8 }}>
                        only in report B ({onlyColumnsB.length})
                      </div>
                      {onlyColumnsB.map((c) => (
                        <Pill key={c} text={c} tone="blue" />
                      ))}
                    </div>
                  )}
                </div>
              )}

              {/* Only in A / only in B tabs */}
              {(activeTab === "onlyA" || activeTab === "onlyB") &&
                (() => {
                  const list = activeTab === "onlyA" ? result.onlyA : result.onlyB;
                  const tone = activeTab === "onlyA" ? C.red : C.blue;
                  const label = matchMode === "key" ? keyColumn : "position";
                  if (list.length === 0) {
                    return <div style={{ color: C.dim, fontSize: 13, fontFamily: MONO }}>none — every row matched.</div>;
                  }
                  return (
                    <div style={{ maxHeight: 420, overflow: "auto", border: `1px solid ${C.border}` }}>
                      <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: MONO, fontSize: 12.5 }}>
                        <thead>
                          <tr style={{ background: C.panelAlt, position: "sticky", top: 0 }}>
                            <th style={{ textAlign: "left", padding: "8px 12px", color: tone, borderBottom: `1px solid ${C.border}` }}>{label}</th>
                            {result.compareCols.slice(0, 5).map((c) => (
                              <th key={c} style={{ textAlign: "left", padding: "8px 12px", color: C.muted, borderBottom: `1px solid ${C.border}` }}>
                                {c}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {list.slice(0, 300).map((item, i) => (
                            <tr key={i} style={{ borderBottom: `1px solid ${C.borderSoft}` }}>
                              <td style={{ padding: "7px 12px", color: tone }}>{item.key}</td>
                              {result.compareCols.slice(0, 5).map((c) => (
                                <td key={c} style={{ padding: "7px 12px", color: C.text }}>
                                  {String(item.row[c] ?? "")}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                      {list.length > 300 && (
                        <div style={{ padding: 10, fontSize: 12, color: C.dim }}>showing first 300 of {list.length}</div>
                      )}
                    </div>
                  );
                })()}

              {/* Mismatches tab */}
              {activeTab === "mismatches" && (
                <div>
                  <div style={{ display: "flex", gap: 10, marginBottom: 14, alignItems: "center" }}>
                    <div style={{ position: "relative", flex: 1, maxWidth: 320 }}>
                      <Search size={13} color={C.dim} style={{ position: "absolute", left: 10, top: 10 }} />
                      <input
                        value={search}
                        onChange={(e) => setSearch(e.target.value)}
                        placeholder="filter by key or column"
                        style={{
                          width: "100%",
                          background: C.panelAlt,
                          border: `1px solid ${C.border}`,
                          color: C.text,
                          fontFamily: MONO,
                          fontSize: 12.5,
                          padding: "7px 10px 7px 30px",
                          boxSizing: "border-box",
                        }}
                      />
                    </div>
                    <button
                      onClick={exportCsv}
                      disabled={result.mismatches.length === 0}
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 6,
                        background: "transparent",
                        border: `1px solid ${C.border}`,
                        color: result.mismatches.length ? C.text : C.dim,
                        fontFamily: SANS,
                        fontSize: 12.5,
                        padding: "7px 12px",
                        cursor: result.mismatches.length ? "pointer" : "not-allowed",
                      }}
                    >
                      <Download size={13} /> export csv
                    </button>
                  </div>

                  {filteredMismatches.length === 0 ? (
                    <div style={{ color: C.dim, fontSize: 13, fontFamily: MONO, display: "flex", alignItems: "center", gap: 8 }}>
                      <CheckCircle2 size={15} color={C.green} /> no cell-level mismatches{search ? " matching that filter" : " in the matched rows"}.
                    </div>
                  ) : (
                    <div style={{ border: `1px solid ${C.border}`, maxHeight: 460, overflow: "auto" }}>
                      {filteredMismatches.slice(0, 400).map((m, i) => (
                        <div
                          key={i}
                          style={{
                            display: "grid",
                            gridTemplateColumns: "28px 1fr",
                            borderBottom: `1px solid ${C.borderSoft}`,
                            fontFamily: MONO,
                            fontSize: 12.5,
                          }}
                        >
                          <div style={{ background: C.panelAlt, color: C.dim, textAlign: "right", padding: "8px 6px", borderRight: `1px solid ${C.border}` }}>
                            {i + 1}
                          </div>
                          <div style={{ padding: "8px 12px" }}>
                            <div style={{ color: C.muted, marginBottom: 4 }}>
                              <span style={{ color: C.blue }}>{m.key}</span>
                              <span style={{ color: C.dim }}> · </span>
                              <span>{m.column}</span>
                            </div>
                            <div style={{ display: "flex", gap: 6, alignItems: "center", color: C.red }}>
                              <Minus size={11} /> {String(m.valueA ?? "") || "(empty)"}
                            </div>
                            <div style={{ display: "flex", gap: 6, alignItems: "center", color: C.green }}>
                              <Plus size={11} /> {String(m.valueB ?? "") || "(empty)"}
                            </div>
                          </div>
                        </div>
                      ))}
                      {filteredMismatches.length > 400 && (
                        <div style={{ padding: 10, fontSize: 12, color: C.dim }}>
                          showing first 400 of {filteredMismatches.length} &mdash; narrow with the filter above
                        </div>
                      )}
                    </div>
                  )}
                </div>
              )}
            </>
          )}
        </>
      )}
    </div>
  );
}
