import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ─── Storage ───────────────────────────────────────────────────────
const STORAGE_KEY = "zept_quote_v3";
const loadSettings = () => { try { return JSON.parse(localStorage.getItem(STORAGE_KEY)) || {}; } catch { return {}; } };
const saveSettings = (s) => localStorage.setItem(STORAGE_KEY, JSON.stringify(s));

const DEFAULT_SETTINGS = {
  companyName: "Zept合同会社", address: "", phone: "", email: "", website: "",
  logoDataUrl: "", coverBgDataUrl: "", endBgDataUrl: "", defaultMarkup: 1.8, accentColor: "#1a56db",
  footerLines: [
    "※ 上記金額は消費税抜き価格です。お支払いの際は消費税を付加してお支払いください。",
    "※ お支払い条件：納品後30日以内",
    "※ 見積有効期限：発行日より30日間",
  ],
};

// ─── Sheet types ───────────────────────────────────────────────────
const SHEET_TYPES = {
  cost:        { label: "見積コスト",     icon: "💰", keywords: ["コスト","全体工数","見積コスト","cost","費用"] },
  wbs:         { label: "WBS・工数",      icon: "📋", keywords: ["wbs","工数","開発詳細","詳細見積","task","タスク"] },
  requirement: { label: "前提条件",       icon: "📌", keywords: ["前提","requirement","条件","スコープ"] },
  schedule:    { label: "スケジュール",   icon: "📅", keywords: ["スケジュール","schedule","マスタ","マイルストーン","milestone","計画","請求"] },
  deliverable: { label: "成果物",         icon: "📦", keywords: ["成果物","deliverable","納品"] },
  overview:    { label: "概要",           icon: "📄", keywords: ["概要","overview","表紙","cover","summary"] },
  history:     { label: "変更履歴",       icon: "🔄", keywords: ["変更履歴","history","changelog","revision"] },
  license:     { label: "ライセンス費用", icon: "🔑", keywords: ["license","ライセンス","account","アカウント"] },
  function:    { label: "機能一覧",       icon: "⚙️", keywords: ["function","機能","origin"] },
  other:       { label: "その他",         icon: "📎", keywords: [] },
};

function classifySheet(name) {
  const lower = name.toLowerCase();
  for (const [type, { keywords }] of Object.entries(SHEET_TYPES)) {
    if (type === "other") continue;
    if (keywords.some((k) => lower.includes(k))) return type;
  }
  return "other";
}

// ─── KV fields that can be auto-filled from quoteInfo ─────────────
const KV_AUTO_MAP = {
  "顧客": "clientName", "client": "clientName", "宛先": "clientName",
  "プロジェクト名": "projectName", "project": "projectName", "件名": "projectName",
  "開始予定日": "startDate", "終了予定日": "endDate", "開発期間": "duration",
  "チーム規模": "teamSize",
  "見積額（人月）": "_totalMM", "総工数": "_totalMM",
  "見積額（円）": "_totalCost", "見積額": "_totalCost",
  "担当": "ownerName", "担当者": "ownerName",
};

// ─── Excel parser ──────────────────────────────────────────────────
function parseAllSheets(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
        const sections = [];

        for (const sheetName of wb.SheetNames) {
          const ws = wb.Sheets[sheetName];
          const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
          const nonEmpty = raw.filter((r) => r.some((c) => c !== null && c !== undefined && String(c).trim() !== ""));
          if (nonEmpty.length < 2) continue;

          const rows = nonEmpty.map((r) => r.map((c) => c !== null && c !== undefined ? String(c).trim() : ""));
          let headerIdx = 0, maxFilled = 0;
          for (let i = 0; i < Math.min(rows.length, 12); i++) {
            const filled = rows[i].filter((c) => c !== "").length;
            if (filled > maxFilled) { maxFilled = filled; headerIdx = i; }
          }

          const type = classifySheet(sheetName);

          // Detect KV layout: rows where col[0] or col[1] is a label and col[1] or col[2] is a value
          const kvPairs = extractKVPairs(rows);
          const isKV = kvPairs.length >= 2;

          sections.push({ id: sheetName, name: sheetName, type, rows, headerIdx, enabled: type !== "other", isKV, kvPairs });
        }

        // Extract cost rows
        let costRows = [];
        const costSection = sections.find((s) => s.type === "cost");
        if (costSection) {
          const { rows, headerIdx } = costSection;
          const header = rows[headerIdx];
          const mmI = header.findIndex((c) => c.includes("工数") || c.includes("人月") || c.includes("MM"));
          const costI = header.findIndex((c) => (c.includes("コスト") || c.includes("単価")) && !c.includes("合計"));
          const amtI = header.findIndex((c) => c.includes("金額") || c.includes("コスト") || c.includes("amount"));
          const itemI = header.findIndex((c) => c.includes("項目") || c.includes("役割") || c.includes("タスク") || c.includes("item") || c.includes("Category"));
          const noI = header.findIndex((c) => c === "No" || c === "NO" || c === "#");
          if (mmI !== -1 && (itemI !== -1 || noI !== -1)) {
            for (let i = headerIdx + 1; i < rows.length; i++) {
              const r = rows[i];
              const item = r[itemI >= 0 ? itemI : noI + 1];
              const mm = parseFloat(r[mmI]);
              if (!item || isNaN(mm) || mm <= 0) continue;
              if (item.includes("合計") || item.includes("注記") || item.includes("TOTAL")) continue;
              const ucRaw = parseFloat(r[costI >= 0 ? costI : mmI + 1]);
              const amtRaw = parseFloat(r[amtI >= 0 ? amtI : mmI + 2]);
              const unitCost = !isNaN(ucRaw) ? ucRaw : (!isNaN(amtRaw) && mm > 0 ? Math.round(amtRaw / mm) : 0);
              costRows.push({ no: String(r[noI >= 0 ? noI : 0] ?? costRows.length + 1), item: item.trim(), manMonth: mm, unitCost });
            }
          }
        }
        if (costRows.length === 0) costRows = [
          { no: "1", item: "要件の明確化", manMonth: 0.95, unitCost: 550000 },
          { no: "2", item: "開発", manMonth: 3.95, unitCost: 550000 },
          { no: "3", item: "テスト", manMonth: 1.55, unitCost: 500000 },
          { no: "4", item: "UAT Go-Live サポート", manMonth: 1.25, unitCost: 550000 },
          { no: "5", item: "BrSE", manMonth: 1.155, unitCost: 800000 },
          { no: "6", item: "管理", manMonth: 1.155, unitCost: 550000 },
        ];

        resolve({ sections, costRows });
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(new Error("読み込みエラー"));
    reader.readAsArrayBuffer(file);
  });
}

// Detect key-value pairs from any sheet layout
function extractKVPairs(rows) {
  const pairs = [];
  for (const row of rows) {
    const nonEmpty = row.filter((c) => c !== "");
    if (nonEmpty.length === 0) continue;
    // Find the first non-empty cell as key, next non-empty as value
    let keyIdx = row.findIndex((c) => c !== "");
    if (keyIdx === -1) continue;
    const key = row[keyIdx];
    // Look for value in same row (next non-empty cell)
    let val = "";
    for (let i = keyIdx + 1; i < row.length; i++) {
      if (row[i] !== "") { val = row[i]; break; }
    }
    // Skip rows that look like section headers (very long values, no real key)
    if (key.length > 40) continue;
    pairs.push({ key, value: val });
  }
  return pairs;
}

// ─── Inline KV Editor ─────────────────────────────────────────────
function KVEditor({ kvPairs, onChange, quoteInfo, costRows, markup, acc }) {
  const totalMM = parseFloat(costRows.reduce((s, r) => s + r.manMonth * markup, 0).toFixed(2));
  const totalCost = Math.round(costRows.reduce((s, r) => s + r.manMonth * r.unitCost * markup, 0));

  const autoFillValue = (key) => {
    const mapKey = Object.keys(KV_AUTO_MAP).find((k) => key.includes(k) || k.includes(key));
    if (!mapKey) return null;
    const field = KV_AUTO_MAP[mapKey];
    if (field === "_totalMM") return String(totalMM);
    if (field === "_totalCost") return String(totalCost.toLocaleString("ja-JP"));
    return quoteInfo[field] || null;
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 8 }}>
        <button onClick={() => {
          const filled = kvPairs.map((p) => {
            const auto = autoFillValue(p.key);
            return auto !== null ? { ...p, value: auto } : p;
          });
          onChange(filled);
        }} style={{ background: acc, border: "none", color: "#fff", borderRadius: 6, padding: "5px 12px", fontSize: 12, cursor: "pointer", fontWeight: 600 }}>
          ⚡ 見積情報から自動入力
        </button>
      </div>
      <div style={{ background: "#0f1117", borderRadius: 8, overflow: "hidden", border: "1px solid #2a2d3e" }}>
        {kvPairs.map((pair, i) => {
          const canAuto = autoFillValue(pair.key) !== null;
          return (
            <div key={i} style={{ display: "flex", alignItems: "center", borderBottom: "1px solid #1e2235", background: i % 2 === 0 ? "#0f1117" : "#12151f" }}>
              <div style={{ width: 180, padding: "8px 12px", fontSize: 12, color: "#9ca3af", borderRight: "1px solid #1e2235", flexShrink: 0, display: "flex", alignItems: "center", gap: 6 }}>
                {canAuto && <span title="自動入力可能" style={{ color: acc, fontSize: 10 }}>●</span>}
                {pair.key}
              </div>
              <input
                value={pair.value}
                onChange={(e) => onChange(kvPairs.map((p, j) => j === i ? { ...p, value: e.target.value } : p))}
                placeholder={canAuto ? `例: ${autoFillValue(pair.key)}` : "入力してください"}
                style={{ flex: 1, background: "transparent", border: "none", color: "#e8eaf0", padding: "8px 12px", fontSize: 13, outline: "none" }}
              />
              {canAuto && !pair.value && (
                <button onClick={() => onChange(kvPairs.map((p, j) => j === i ? { ...p, value: autoFillValue(pair.key) } : p))}
                  title="自動入力" style={{ background: "none", border: "none", color: acc, cursor: "pointer", padding: "0 10px", fontSize: 11 }}>↓入力</button>
              )}
            </div>
          );
        })}
        {/* Add row */}
        <div style={{ padding: "8px 12px", borderTop: "1px solid #2a2d3e" }}>
          <button onClick={() => onChange([...kvPairs, { key: "", value: "" }])}
            style={{ background: "none", border: "1px dashed #3a3d50", color: "#6b7280", borderRadius: 6, padding: "4px 12px", fontSize: 12, cursor: "pointer" }}>
            ＋ 行を追加
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── Table Editor (for WBS, schedule, etc.) ───────────────────────
function TableEditor({ rows, headerIdx, onChange }) {
  const header = rows[headerIdx] || [];
  const dataRows = rows.slice(headerIdx + 1).filter((r) => r.some((c) => c !== ""));
  const colCount = Math.max(...rows.map((r) => r.length), 1);

  const updateCell = (rowI, colI, val) => {
    const newRows = [...rows];
    const actualIdx = headerIdx + 1 + rowI;
    if (!newRows[actualIdx]) return;
    const newRow = [...newRows[actualIdx]];
    newRow[colI] = val;
    newRows[actualIdx] = newRow;
    onChange(newRows);
  };

  return (
    <div style={{ overflowX: "auto" }}>
      <table style={{ borderCollapse: "collapse", fontSize: 12, width: "100%" }}>
        <thead>
          <tr style={{ background: "#12151f" }}>
            {header.map((h, i) => h !== "" && (
              <th key={i} style={{ padding: "8px 10px", textAlign: "left", color: "#9ca3af", fontWeight: 600, fontSize: 11, borderBottom: "1px solid #2a2d3e", whiteSpace: "nowrap" }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {dataRows.map((row, ri) => (
            <tr key={ri} style={{ borderBottom: "1px solid #1e2235", background: ri % 2 === 0 ? "#0f1117" : "#12151f" }}>
              {header.map((h, ci) => h !== "" && (
                <td key={ci} style={{ padding: "4px 6px" }}>
                  <input value={row[ci] || ""} onChange={(e) => updateCell(ri, ci, e.target.value)}
                    style={{ width: "100%", background: "transparent", border: "none", color: "#e8eaf0", fontSize: 12, padding: "4px 6px", outline: "none", minWidth: 80 }} />
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// ─── Sheet renderer for preview ───────────────────────────────────
function SheetTable({ rows, headerIdx, kvPairs, isKV, acc }) {
  if (isKV && kvPairs) {
    const displayPairs = kvPairs.filter((p) => p.key !== "");
    return (
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
        <tbody>
          {displayPairs.map((p, i) => (
            <tr key={i} style={{ borderBottom: "1px solid #f0f0f0", background: i % 2 === 0 ? "#fafafa" : "#fff" }}>
              <td style={{ padding: "8px 14px", fontWeight: 600, color: "#444", width: 200, background: `${acc}08` }}>{p.key}</td>
              <td style={{ padding: "8px 14px", color: "#222" }}>{p.value}</td>
            </tr>
          ))}
        </tbody>
      </table>
    );
  }
  if (!rows || rows.length === 0) return null;
  const header = rows[headerIdx] || [];
  const dataRows = rows.slice(headerIdx + 1).filter((r) => r.some((c) => c !== ""));
  const colCount = header.filter((h) => h !== "").length;
  if (colCount === 0) return null;
  return (
    <div style={{ overflowX: "auto" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
        <thead>
          <tr style={{ background: acc }}>
            {header.filter((h) => h !== "").map((h, i) => (
              <th key={i} style={{ padding: "8px 10px", textAlign: "left", color: "#fff", fontWeight: 600, fontSize: 11, whiteSpace: "nowrap" }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {dataRows.map((r, i) => (
            <tr key={i} style={{ background: i % 2 === 0 ? "#fafafa" : "#fff", borderBottom: "1px solid #f0f0f0" }}>
              {r.slice(0, header.filter((h) => h !== "").length).map((cell, ci) => (
                <td key={ci} style={{ padding: "7px 10px", verticalAlign: "top", color: "#333", whiteSpace: "pre-wrap", maxWidth: 320 }}>{cell}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// ─── Cost table for preview ────────────────────────────────────────
function CostTable({ rows, markup, acc }) {
  const fmt = (n) => "¥ " + Math.round(n).toLocaleString("ja-JP");
  const computed = rows.map((r) => ({ ...r, newAmt: r.manMonth * r.unitCost * markup, newMM: parseFloat((r.manMonth * markup).toFixed(3)), newUC: Math.round(r.unitCost * markup) }));
  const total = computed.reduce((s, r) => s + r.newAmt, 0);
  return (
    <>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
        <thead><tr style={{ background: acc }}>
          {["No", "項目", "工数（人月）", "単価（JPY）", "金額（JPY）"].map((h, i) => (
            <th key={h} style={{ padding: "9px 12px", textAlign: i >= 2 ? "right" : "left", color: "#fff", fontWeight: 600, fontSize: 11 }}>{h}</th>
          ))}
        </tr></thead>
        <tbody>
          {computed.map((r, i) => (
            <tr key={i} style={{ background: i % 2 === 0 ? "#fafafa" : "#fff", borderBottom: "1px solid #f0f0f0" }}>
              <td style={{ padding: "8px 12px", color: "#999", fontSize: 11 }}>{r.no}</td>
              <td style={{ padding: "8px 12px", fontWeight: 500 }}>{r.item}</td>
              <td style={{ padding: "8px 12px", textAlign: "right" }}>{r.newMM}</td>
              <td style={{ padding: "8px 12px", textAlign: "right" }}>{r.newUC.toLocaleString()}</td>
              <td style={{ padding: "8px 12px", textAlign: "right", fontWeight: 600 }}>{fmt(r.newAmt)}</td>
            </tr>
          ))}
          <tr style={{ background: `${acc}15`, fontWeight: 700, borderTop: `2px solid ${acc}` }}>
            <td colSpan={4} style={{ padding: "10px 12px" }}>合計（税抜）</td>
            <td style={{ padding: "10px 12px", textAlign: "right", color: acc, fontSize: 14 }}>{fmt(total)}</td>
          </tr>
          <tr style={{ background: "#fafafa" }}>
            <td colSpan={4} style={{ padding: "7px 12px", color: "#888", fontSize: 11 }}>消費税（10%）</td>
            <td style={{ padding: "7px 12px", textAlign: "right", color: "#888", fontSize: 11 }}>{fmt(total * 0.1)}</td>
          </tr>
          <tr style={{ background: `${acc}20`, fontWeight: 700 }}>
            <td colSpan={4} style={{ padding: "10px 12px" }}>合計（税込）</td>
            <td style={{ padding: "10px 12px", textAlign: "right", color: acc, fontSize: 14 }}>{fmt(total * 1.1)}</td>
          </tr>
        </tbody>
      </table>
      <div style={{ marginTop: 8, fontSize: 11, color: "#888" }}>※ 上記金額は消費税抜き価格です</div>
    </>
  );
}

// ─── Quote Preview ─────────────────────────────────────────────────
function QuotePreview({ sections, costRows, markup, settings, quoteInfo }) {
  const acc = settings.accentColor;
  const fmt = (n) => "¥ " + Math.round(n).toLocaleString("ja-JP");
  const total = costRows.reduce((s, r) => s + r.manMonth * r.unitCost * markup, 0);

  return (
    <div id="quote-preview" style={{ background: "#fff", color: "#111", fontFamily: "'Hiragino Sans', 'Meiryo', sans-serif" }}>
      <div style={{ background: acc, padding: "20px 36px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div>
          {settings.logoDataUrl
            ? <img src={settings.logoDataUrl} alt="logo" style={{ height: 44, maxWidth: 160, objectFit: "contain", filter: "brightness(0) invert(1)" }} />
            : <div style={{ color: "#fff", fontWeight: 800, fontSize: 22 }}>{settings.companyName}</div>}
        </div>
        <div style={{ textAlign: "right", color: "rgba(255,255,255,0.85)", fontSize: 11, lineHeight: 1.9 }}>
          {settings.address && <div>{settings.address}</div>}
          {settings.phone && <div>TEL: {settings.phone}</div>}
          {settings.email && <div>{settings.email}</div>}
          {settings.website && <div>{settings.website}</div>}
        </div>
      </div>
      <div style={{ padding: "28px 36px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 20 }}>
          <div>
            <div style={{ fontSize: 24, fontWeight: 800, letterSpacing: "0.1em", marginBottom: 4 }}>御　見　積　書</div>
            <div style={{ fontSize: 13, color: "#555" }}>{quoteInfo.projectName}</div>
          </div>
          <div style={{ textAlign: "right", fontSize: 12, color: "#666", lineHeight: 2 }}>
            <div>見積番号：{quoteInfo.quoteNo}</div>
            <div>発行日：{quoteInfo.date}</div>
            <div>有効期限：{quoteInfo.expiry}</div>
          </div>
        </div>
        <div style={{ borderBottom: `3px solid ${acc}`, paddingBottom: 10, marginBottom: 16, display: "flex", alignItems: "baseline", gap: 10 }}>
          <span style={{ fontSize: 17, fontWeight: 700 }}>{quoteInfo.clientName}</span>
          <span style={{ fontSize: 13, color: "#666" }}>御中</span>
        </div>
        <div style={{ background: "#f8f9ff", border: `1px solid ${acc}30`, borderRadius: 8, padding: "14px 22px", marginBottom: 24, display: "flex", alignItems: "baseline", gap: 14 }}>
          <span style={{ fontSize: 13, color: "#666" }}>御見積金額（税抜）</span>
          <span style={{ fontSize: 30, fontWeight: 800, color: acc }}>{fmt(total)}</span>
          <span style={{ fontSize: 12, color: "#aaa" }}>（消費税別途）</span>
        </div>
        {sections.filter((s) => s.enabled).map((section) => {
          const meta = SHEET_TYPES[section.type] || SHEET_TYPES.other;
          return (
            <div key={section.id} style={{ marginBottom: 28 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, background: `${acc}12`, borderLeft: `4px solid ${acc}`, padding: "9px 14px", marginBottom: 12, borderRadius: "0 6px 6px 0" }}>
                <span style={{ fontSize: 15 }}>{meta.icon}</span>
                <span style={{ fontWeight: 700, fontSize: 14, color: acc }}>{meta.label}</span>
                <span style={{ fontSize: 11, color: "#888", marginLeft: 4 }}>（{section.name}）</span>
              </div>
              {section.type === "cost"
                ? <CostTable rows={costRows} markup={markup} acc={acc} />
                : <SheetTable rows={section.rows} headerIdx={section.headerIdx} kvPairs={section.kvPairs} isKV={section.isKV} acc={acc} />
              }
            </div>
          );
        })}
        {settings.footerLines.length > 0 && (
          <div style={{ borderTop: "1px solid #e5e7eb", paddingTop: 14, marginTop: 8 }}>
            {settings.footerLines.map((line, i) => (
              <div key={i} style={{ fontSize: 11, color: "#888", lineHeight: 2 }}>{line}</div>
            ))}
          </div>
        )}
        <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 28 }}>
          <div style={{ textAlign: "center", border: "1px solid #ddd", borderRadius: 4, padding: "12px 32px", fontSize: 12, color: "#666" }}>
            <div style={{ fontWeight: 600, marginBottom: 4 }}>{settings.companyName}</div>
            <div style={{ color: "#bbb", fontSize: 11 }}>担当者</div>
            <div style={{ height: 36 }} />
          </div>
        </div>
      </div>
      <div style={{ background: acc, height: 6 }} />
    </div>
  );
}

// ─── Settings Panel ────────────────────────────────────────────────
function SettingsPanel({ settings, onChange, onClose }) {
  const logoRef = useRef();
  const coverBgRef = useRef();
  const endBgRef = useRef();
  const [footerText, setFooterText] = useState(settings.footerLines.join("\n"));
  const handleLogo = (e) => {
    const file = e.target.files?.[0]; if (!file) return;
    const r = new FileReader(); r.onload = (ev) => onChange({ ...settings, logoDataUrl: ev.target.result }); r.readAsDataURL(file);
  };
  const fld = (label, key, ph, type = "text") => (
    <div style={{ marginBottom: 11 }}>
      <label style={{ display: "block", fontSize: 11, color: "#9ca3af", marginBottom: 3 }}>{label}</label>
      <input type={type} placeholder={ph} value={settings[key] || ""} onChange={(e) => onChange({ ...settings, [key]: e.target.value })}
        style={{ width: "100%", background: "#0f1117", border: "1px solid #2a2d3e", color: "#e8eaf0", borderRadius: 6, padding: "7px 10px", fontSize: 13, boxSizing: "border-box" }} />
    </div>
  );
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.7)", zIndex: 1000, display: "flex", justifyContent: "flex-end" }} onClick={onClose}>
      <div style={{ width: 380, background: "#1a1d2e", height: "100%", overflowY: "auto", padding: 22, boxSizing: "border-box", borderLeft: "1px solid #2a2d3e" }} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
          <span style={{ fontWeight: 700, fontSize: 15 }}>⚙️ 会社設定</span>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#9ca3af", fontSize: 20, cursor: "pointer" }}>×</button>
        </div>
        <div style={{ marginBottom: 14 }}>
          <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 7 }}>ロゴ画像</div>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            {settings.logoDataUrl ? <img src={settings.logoDataUrl} alt="logo" style={{ height: 38, maxWidth: 110, objectFit: "contain", background: "#fff", borderRadius: 4, padding: 4 }} />
              : <div style={{ width: 80, height: 38, background: "#12151f", borderRadius: 4, border: "1px dashed #3a3d50", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 10, color: "#6b7280" }}>No logo</div>}
            <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
              <button onClick={() => logoRef.current.click()} style={{ background: "#1e2235", border: "1px solid #3a3d50", color: "#9ca3af", borderRadius: 6, padding: "5px 10px", fontSize: 12, cursor: "pointer" }}>📂 アップロード</button>
              {settings.logoDataUrl && <button onClick={() => onChange({ ...settings, logoDataUrl: "" })} style={{ background: "none", border: "none", color: "#ef4444", fontSize: 11, cursor: "pointer" }}>削除</button>}
            </div>
            <input ref={logoRef} type="file" accept="image/*" onChange={handleLogo} style={{ display: "none" }} />
          </div>
        </div>

        {/* Cover & End BG */}
        <div style={{ borderTop: "1px solid #2a2d3e", paddingTop: 14, marginBottom: 14 }}>
          <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 10, letterSpacing: "0.05em" }}>
            スライド背景テンプレート
            <span style={{ fontSize: 10, color: "#4b5563", marginLeft: 6 }}>既存PPTXのスライドをPNG保存してアップロード</span>
          </div>
          {[
            { label: "表紙スライド背景", key: "coverBgDataUrl", hint: "表紙スライドを右クリック→「図として保存」→PNG" },
            { label: "締めスライド背景", key: "endBgDataUrl", hint: "最終スライドを右クリック→「図として保存」→PNG" },
          ].map(({ label, key, hint }) => {
            const bgRef = { coverBgDataUrl: coverBgRef, endBgDataUrl: endBgRef }[key];
            return (
              <div key={key} style={{ marginBottom: 12 }}>
                <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 5 }}>{label}</div>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  {settings[key]
                    ? <img src={settings[key]} alt={label} style={{ height: 40, width: 70, objectFit: "cover", borderRadius: 4, border: "1px solid #3a3d50" }} />
                    : <div style={{ width: 70, height: 40, background: "#12151f", borderRadius: 4, border: "1px dashed #3a3d50", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 9, color: "#6b7280", textAlign: "center", padding: 2 }}>未設定</div>
                  }
                  <div style={{ flex: 1 }}>
                    <button onClick={() => bgRef.current.click()} style={{ background: "#1e2235", border: "1px solid #3a3d50", color: "#9ca3af", borderRadius: 6, padding: "5px 10px", fontSize: 11, cursor: "pointer", display: "block", marginBottom: 4 }}>📂 PNG選択</button>
                    <div style={{ fontSize: 9, color: "#4b5563", lineHeight: 1.4 }}>{hint}</div>
                  </div>
                  {settings[key] && <button onClick={() => onChange({ ...settings, [key]: "" })} style={{ background: "none", border: "none", color: "#ef4444", fontSize: 11, cursor: "pointer", flexShrink: 0 }}>削除</button>}
                </div>
                <input ref={bgRef} type="file" accept="image/png,image/jpeg,image/*" onChange={(e) => {
                  const file = e.target.files?.[0]; if (!file) return;
                  const r = new FileReader(); r.onload = (ev) => onChange({ ...settings, [key]: ev.target.result }); r.readAsDataURL(file);
                  e.target.value = "";
                }} style={{ display: "none" }} />
              </div>
            );
          })}
          <div style={{ background: "#12151f", borderRadius: 6, padding: "8px 12px", fontSize: 10, color: "#6b7280", lineHeight: 1.8 }}>
            💡 設定方法：PowerPointで既存の見積書を開く → 表紙スライドを右クリック →「図として保存」→ PNG形式で保存 → ここにアップロード
          </div>
        </div>

        <div style={{ borderTop: "1px solid #2a2d3e", paddingTop: 12, marginBottom: 12 }}>
          {fld("会社名", "companyName", "Zept合同会社")}
          {fld("住所", "address", "東京都〇〇区〇〇 1-2-3")}
          {fld("電話番号", "phone", "03-XXXX-XXXX")}
          {fld("メールアドレス", "email", "info@zept.com", "email")}
          {fld("ウェブサイト", "website", "https://zept.com")}
        </div>
        <div style={{ borderTop: "1px solid #2a2d3e", paddingTop: 12, marginBottom: 12 }}>
          <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 8 }}>デフォルト掛け率</div>
          <div style={{ display: "flex", gap: 5, flexWrap: "wrap", marginBottom: 10 }}>
            {[1.2, 1.3, 1.5, 1.8, 2.0].map((m) => (
              <button key={m} onClick={() => onChange({ ...settings, defaultMarkup: m })}
                style={{ background: settings.defaultMarkup === m ? settings.accentColor : "#12151f", border: `1px solid ${settings.defaultMarkup === m ? settings.accentColor : "#2a2d3e"}`, color: settings.defaultMarkup === m ? "#fff" : "#9ca3af", borderRadius: 6, padding: "4px 10px", fontSize: 12, cursor: "pointer" }}>×{m}</button>
            ))}
            <input type="number" step="0.01" min="1" value={settings.defaultMarkup} onChange={(e) => onChange({ ...settings, defaultMarkup: parseFloat(e.target.value) || 1.8 })}
              style={{ width: 54, background: "#12151f", border: "1px solid #2a2d3e", color: "#e8eaf0", borderRadius: 6, padding: "4px 7px", fontSize: 12, textAlign: "center" }} />
          </div>
          <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 7 }}>アクセントカラー</div>
          <div style={{ display: "flex", gap: 7, alignItems: "center" }}>
            {["#1a56db", "#0d9488", "#7c3aed", "#dc2626", "#d97706", "#0f766e"].map((c) => (
              <div key={c} onClick={() => onChange({ ...settings, accentColor: c })}
                style={{ width: 22, height: 22, borderRadius: "50%", background: c, cursor: "pointer", border: settings.accentColor === c ? "3px solid #fff" : "3px solid transparent" }} />
            ))}
            <input type="color" value={settings.accentColor} onChange={(e) => onChange({ ...settings, accentColor: e.target.value })}
              style={{ width: 22, height: 22, borderRadius: "50%", border: "none", cursor: "pointer", padding: 0 }} />
          </div>
        </div>
        <div style={{ borderTop: "1px solid #2a2d3e", paddingTop: 12, marginBottom: 14 }}>
          <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 7 }}>フッター文言</div>
          <textarea rows={5} value={footerText} onChange={(e) => { setFooterText(e.target.value); onChange({ ...settings, footerLines: e.target.value.split("\n").filter(Boolean) }); }}
            style={{ width: "100%", background: "#0f1117", border: "1px solid #2a2d3e", color: "#e8eaf0", borderRadius: 6, padding: "8px 10px", fontSize: 12, boxSizing: "border-box", resize: "vertical" }} />
        </div>
        <button onClick={onClose} style={{ width: "100%", background: settings.accentColor, border: "none", color: "#fff", borderRadius: 8, padding: "11px 0", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>✓ 保存して閉じる</button>
      </div>
    </div>
  );
}


// ─── PPTX Generator ───────────────────────────────────────────────
function hexC(color) { return color.replace("#", ""); }

async function generatePPTX({ sections, costRows, markup, settings, quoteInfo }) {
  const PptxGenJS = (await import("pptxgenjs")).default;
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.title = quoteInfo.projectName;

  const acc = hexC(settings.accentColor);
  const W = 10, H = 5.625, M = 0.4;
  const fmt = (n) => "¥ " + Math.round(n).toLocaleString("ja-JP");

  const addHeader = (slide, title) => {
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.65, fill: { color: acc }, line: { color: acc } });
    slide.addText(title, { x: M, y: 0.05, w: W - M * 2, h: 0.56, fontSize: 16, bold: true, color: "FFFFFF", valign: "middle" });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: H - 0.3, w: W, h: 0.3, fill: { color: "F3F4F6" }, line: { color: "E5E7EB" } });
    slide.addText(settings.companyName, { x: M, y: H - 0.28, w: W / 2, h: 0.25, fontSize: 8, color: "9CA3AF" });
    slide.addText(quoteInfo.date, { x: W / 2, y: H - 0.28, w: W / 2 - M, h: 0.25, fontSize: 8, color: "9CA3AF", align: "right" });
  };

  // ── Slide 1: Cover ──────────────────────────────────────────────
  const cover = pres.addSlide();
  if (settings.coverBgDataUrl) {
    cover.background = { data: settings.coverBgDataUrl };
    cover.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: "000000", transparency: 55 }, line: { color: "000000", transparency: 55 } });
  } else {
    cover.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: acc }, line: { color: acc } });
    cover.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: "000000", transparency: 45 }, line: { color: "000000", transparency: 45 } });
    cover.addShape(pres.shapes.RECTANGLE, { x: 0, y: H - 1.4, w: W, h: 1.4, fill: { color: "000000", transparency: 62 }, line: { color: "000000", transparency: 62 } });
  }

  if (!settings.coverBgDataUrl) {
    if (settings.logoDataUrl) {
      try { cover.addImage({ data: settings.logoDataUrl, x: M, y: 0.35, w: 2.8, h: 0.75, sizing: { type: "contain", w: 2.8, h: 0.75 } }); } catch(e) {}
    } else {
      cover.addText(settings.companyName, { x: M, y: 0.28, w: 5, h: 0.7, fontSize: 16, bold: true, color: "FFFFFF" });
    }
  }

  cover.addText("御　見　積　書", { x: M, y: 1.3, w: W - M * 2, h: 1.3, fontSize: 44, bold: true, color: "FFFFFF", align: "center", charSpacing: 10, valign: "middle" });
  cover.addText(quoteInfo.projectName, { x: M, y: 2.7, w: W - M * 2, h: 0.55, fontSize: 15, color: "FFFFFF", align: "center" });
  cover.addText(`${quoteInfo.clientName}　御中`, { x: M, y: 3.3, w: W - M * 2, h: 0.5, fontSize: 13, color: "FFFFFF", align: "center" });

  cover.addText([
    { text: `見積番号：${quoteInfo.quoteNo}`, options: { breakLine: true } },
    { text: `発行日：${quoteInfo.date}`, options: { breakLine: true } },
    { text: `有効期限：${quoteInfo.expiry}` }
  ], { x: M, y: H - 1.3, w: W - M * 2, h: 1.1, fontSize: 11, color: "DDDDDD", align: "center" });

  // ── Slide 2: Cost ───────────────────────────────────────────────
  const total = costRows.reduce((s, r) => s + r.manMonth * r.unitCost * markup, 0);
  const costSlide = pres.addSlide();
  costSlide.background = { color: "FFFFFF" };
  addHeader(costSlide, "💰 見積コスト");

  costSlide.addShape(pres.shapes.RECTANGLE, { x: M, y: 0.75, w: W - M * 2, h: 0.6, fill: { color: acc, transparency: 88 }, line: { color: acc, transparency: 88 } });
  costSlide.addText(`御見積金額（税抜）：${fmt(total)}　　（税込：${fmt(total * 1.1)}）`,
    { x: M + 0.1, y: 0.78, w: W - M * 2, h: 0.52, fontSize: 13, bold: true, color: acc, valign: "middle" });

  const hOpts = (text) => ({ text, options: { fill: { color: acc }, color: "FFFFFF", bold: true, fontSize: 10 } });
  const cOpts = (text, right, i) => ({ text, options: { fill: { color: i % 2 === 0 ? "F8F9FF" : "FFFFFF" }, fontSize: 10, align: right ? "right" : "left" } });

  const cRows = costRows.map((r, i) => [
    cOpts(r.no, true, i),
    cOpts(r.item, false, i),
    cOpts(String(parseFloat((r.manMonth * markup).toFixed(3))), true, i),
    cOpts(Math.round(r.unitCost * markup).toLocaleString("ja-JP"), true, i),
    { text: fmt(r.manMonth * r.unitCost * markup), options: { fill: { color: i % 2 === 0 ? "F8F9FF" : "FFFFFF" }, fontSize: 10, align: "right", bold: true } },
  ]);
  const totRow = (label, val, bg) => [
    { text: "", options: { fill: { color: bg } } },
    { text: label, options: { fill: { color: bg }, bold: true } },
    { text: "", options: { fill: { color: bg } } },
    { text: "", options: { fill: { color: bg } } },
    { text: val, options: { fill: { color: bg }, bold: true, align: "right", color: acc } },
  ];

  costSlide.addTable([
    [hOpts("No"), hOpts("項目"), hOpts("工数（人月）"), hOpts("単価（JPY）"), hOpts("金額（JPY）")],
    ...cRows,
    totRow("合計（税抜）", fmt(total), "EFF6FF"),
    totRow("消費税（10%）", fmt(total * 0.1), "F9FAFB"),
    totRow("合計（税込）", fmt(total * 1.1), "DBEAFE"),
  ], {
    x: M, y: 1.42, w: W - M * 2, h: H - 1.88,
    border: { pt: 0.5, color: "E5E7EB" }, colW: [0.5, 3.6, 1.2, 1.3, 1.5]
  });

  // ── Other sections ──────────────────────────────────────────────
  for (const section of sections.filter(s => s.enabled && s.type !== "cost")) {
    const meta = SHEET_TYPES[section.type] || SHEET_TYPES.other;
    const slideTitle = `${meta.icon} ${meta.label}`;
    const CONTENT_H = H - 1.15;
    const CONTENT_Y = 0.75;

    if (section.isKV && section.kvPairs) {
      const pairs = section.kvPairs.filter(p => p.key !== "");
      if (pairs.length === 0) continue;
      const PER = 16;
      for (let ci = 0; ci * PER < pairs.length; ci++) {
        const sl = pres.addSlide();
        sl.background = { color: "FFFFFF" };
        addHeader(sl, slideTitle + (pairs.length > PER ? ` (${ci + 1}/${Math.ceil(pairs.length / PER)})` : ""));
        const chunk = pairs.slice(ci * PER, (ci + 1) * PER);
        sl.addTable(
          chunk.map((p, i) => [
            { text: p.key, options: { fill: { color: i % 2 === 0 ? "F5F7FF" : "FFFFFF" }, bold: true, color: "374151", fontSize: 10 } },
            { text: p.value || "", options: { fill: { color: i % 2 === 0 ? "F5F7FF" : "FFFFFF" }, color: "111827", fontSize: 10 } },
          ]),
          { x: M, y: CONTENT_Y, w: W - M * 2, h: CONTENT_H, border: { pt: 0.5, color: "E5E7EB" }, colW: [3, W - M * 2 - 3] }
        );
      }
    } else if (section.rows && section.rows.length > 1) {
      const header = section.rows[section.headerIdx] || [];
      const visCols = header.map((h, i) => ({ h, i })).filter(({ h }) => h !== "");
      if (visCols.length === 0) continue;
      const dataRows = section.rows.slice(section.headerIdx + 1).filter(r => r.some(c => c !== ""));
      if (dataRows.length === 0) continue;
      const PER = 14;
      const colW = Array(visCols.length).fill((W - M * 2) / visCols.length);
      for (let ci = 0; ci * PER < dataRows.length; ci++) {
        const sl = pres.addSlide();
        sl.background = { color: "FFFFFF" };
        addHeader(sl, slideTitle + (dataRows.length > PER ? ` (${ci + 1}/${Math.ceil(dataRows.length / PER)})` : ""));
        const chunk = dataRows.slice(ci * PER, (ci + 1) * PER);
        sl.addTable([
          visCols.map(({ h }) => ({ text: h, options: { fill: { color: acc }, color: "FFFFFF", bold: true, fontSize: 9 } })),
          ...chunk.map((r, ri) => visCols.map(({ i: ci2 }) => ({
            text: String(r[ci2] || ""),
            options: { fill: { color: ri % 2 === 0 ? "FAFAFA" : "FFFFFF" }, fontSize: 9 }
          })))
        ], { x: M, y: CONTENT_Y, w: W - M * 2, h: CONTENT_H, border: { pt: 0.5, color: "E5E7EB" }, colW });
      }
    }
  }

  // ── Last slide ──────────────────────────────────────────────────
  const last = pres.addSlide();
  if (settings.endBgDataUrl) {
    last.background = { data: settings.endBgDataUrl };
    last.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: "000000", transparency: 55 }, line: { color: "000000", transparency: 55 } });
  } else {
    last.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: acc }, line: { color: acc } });
    last.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: H, fill: { color: "000000", transparency: 45 }, line: { color: "000000", transparency: 45 } });
  }
  if (!settings.endBgDataUrl) {
    last.addText("以上、よろしくお願い申し上げます", { x: M, y: 1.6, w: W - M * 2, h: 1, fontSize: 22, bold: true, color: "FFFFFF", align: "center" });
    last.addText(settings.companyName, { x: M, y: 2.8, w: W - M * 2, h: 0.6, fontSize: 16, color: "FFFFFF", align: "center" });
    const contactLines = [settings.address, settings.phone ? `TEL: ${settings.phone}` : "", settings.email].filter(Boolean);
    if (contactLines.length > 0) {
      last.addText(contactLines.join("　"), { x: M, y: 3.5, w: W - M * 2, h: 0.4, fontSize: 11, color: "CCCCCC", align: "center" });
    }
  }

  await pres.writeFile({ fileName: `御見積書_${quoteInfo.clientName}_${quoteInfo.date.replace(/\//g, "")}.pptx` });
}

// ─── Main App ──────────────────────────────────────────────────────
export default function App() {
  const saved = loadSettings();
  const [settings, setSettings] = useState({ ...DEFAULT_SETTINGS, ...saved });
  const [sections, setSections] = useState([]);
  const [costRows, setCostRows] = useState([
    { no: "1", item: "要件の明確化", manMonth: 0.95, unitCost: 550000 },
    { no: "2", item: "開発", manMonth: 3.95, unitCost: 550000 },
    { no: "3", item: "テスト", manMonth: 1.55, unitCost: 500000 },
    { no: "4", item: "UAT Go-Live サポート", manMonth: 1.25, unitCost: 550000 },
    { no: "5", item: "BrSE", manMonth: 1.155, unitCost: 800000 },
    { no: "6", item: "管理", manMonth: 1.155, unitCost: 550000 },
  ]);
  const [markup, setMarkup] = useState(saved.defaultMarkup || 1.8);
  const [quoteInfo, setQuoteInfo] = useState({
    clientName: "株式会社エンケイ", projectName: "部品の設計図作成システム 1次フェーズ",
    quoteNo: "QT-" + new Date().toISOString().slice(0, 10).replace(/-/g, ""),
    date: new Date().toLocaleDateString("ja-JP", { year: "numeric", month: "2-digit", day: "2-digit" }),
    expiry: new Date(Date.now() + 30 * 86400000).toLocaleDateString("ja-JP", { year: "numeric", month: "2-digit", day: "2-digit" }),
    startDate: "", endDate: "", duration: "", teamSize: "", ownerName: "",
  });
  const [tab, setTab] = useState("sections");
  const [expandedSection, setExpandedSection] = useState(null);
  const [showSettings, setShowSettings] = useState(false);
  const [loading, setLoading] = useState(false);
  const [pptxGenerating, setPptxGenerating] = useState(false);
  const [error, setError] = useState("");
  const fileRef = useRef();

  const handleSettingsChange = (s) => { setSettings(s); saveSettings(s); };

  const handleFile = useCallback(async (e) => {
    const file = e.target.files?.[0]; if (!file) return;
    setLoading(true); setError("");
    try {
      const result = await parseAllSheets(file);
      setSections(result.sections);
      setCostRows(result.costRows);
    } catch (err) { setError("❌ " + err.message); }
    finally { setLoading(false); e.target.value = ""; }
  }, []);

  const toggleSection = (id) => setSections((p) => p.map((s) => s.id === id ? { ...s, enabled: !s.enabled } : s));
  const moveSection = (id, dir) => setSections((prev) => {
    const idx = prev.findIndex((s) => s.id === id), next = [...prev], swap = idx + dir;
    if (swap < 0 || swap >= next.length) return prev;
    [next[idx], next[swap]] = [next[swap], next[idx]]; return next;
  });
  const updateSectionKV = (id, kvPairs) => setSections((p) => p.map((s) => s.id === id ? { ...s, kvPairs } : s));
  const updateSectionRows = (id, rows) => setSections((p) => p.map((s) => s.id === id ? { ...s, rows } : s));
  const updateCostRow = (idx, field, val) => setCostRows((p) => p.map((r, i) => i === idx ? { ...r, [field]: val } : r));

  const origTotal = costRows.reduce((s, r) => s + r.manMonth * r.unitCost, 0);
  const newTotal = costRows.reduce((s, r) => s + r.manMonth * r.unitCost * markup, 0);
  const fmt = (n) => "¥ " + Math.round(n).toLocaleString("ja-JP");
  const acc = settings.accentColor;

  const handlePPTX = async () => {
    setPptxGenerating(true);
    try {
      await generatePPTX({ sections: sections.length > 0 ? sections : [{ id: "cost", name: "見積コスト", type: "cost", enabled: true, rows: [], headerIdx: 0, isKV: false, kvPairs: [] }], costRows, markup, settings, quoteInfo });
    } catch(err) { alert("PPTX生成エラー: " + err.message); }
    finally { setPptxGenerating(false); }
  };

  const handlePrint = () => {
    const style = document.createElement("style");
    style.textContent = `@media print { body > *:not(#print-root) { display: none !important; } #print-root { display: block !important; } @page { margin: 12mm; } }`;
    const root = document.getElementById("quote-preview").cloneNode(true);
    root.id = "print-root"; document.body.appendChild(style); document.body.appendChild(root);
    window.print(); document.body.removeChild(style); document.body.removeChild(root);
  };

  return (
    <div style={{ minHeight: "100vh", background: "#0f1117", color: "#e8eaf0", fontFamily: "'Hiragino Sans', 'Meiryo', sans-serif" }}>
      {showSettings && <SettingsPanel settings={settings} onChange={handleSettingsChange} onClose={() => setShowSettings(false)} />}
      <header style={{ background: "#1a1d2e", borderBottom: "1px solid #2a2d3e", padding: "12px 24px", display: "flex", alignItems: "center", gap: 12 }}>
        {settings.logoDataUrl ? <img src={settings.logoDataUrl} alt="logo" style={{ height: 28, maxWidth: 90, objectFit: "contain" }} />
          : <div style={{ fontWeight: 800, fontSize: 15, color: acc }}>{settings.companyName}</div>}
        <div style={{ width: 1, height: 20, background: "#2a2d3e" }} />
        <div style={{ fontSize: 13, color: "#9ca3af" }}>見積書コンバーター</div>
        <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
          <button onClick={() => fileRef.current.click()} style={{ background: "#1e2235", border: "1px solid #3a3d50", color: "#9ca3af", borderRadius: 7, padding: "7px 14px", fontSize: 13, cursor: "pointer" }}>📂 Excelを読み込む</button>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
          <button onClick={() => setShowSettings(true)} style={{ background: "#1e2235", border: "1px solid #3a3d50", color: "#9ca3af", borderRadius: 7, padding: "7px 14px", fontSize: 13, cursor: "pointer" }}>⚙️ 会社設定</button>
        </div>
      </header>
      <div style={{ background: "#1a1d2e", borderBottom: "1px solid #2a2d3e", padding: "0 24px", display: "flex" }}>
        {[["sections", "📑 セクション管理"], ["cost", "💰 見積金額"], ["preview", "👁 プレビュー"]].map(([id, label]) => (
          <button key={id} onClick={() => setTab(id)}
            style={{ background: "none", border: "none", borderBottom: tab === id ? `2px solid ${acc}` : "2px solid transparent", color: tab === id ? "#e8eaf0" : "#6b7280", padding: "11px 18px", fontSize: 13, cursor: "pointer", fontWeight: tab === id ? 600 : 400 }}>
            {label}
          </button>
        ))}
      </div>

      {error && <div style={{ background: "#2d1515", margin: "12px 24px", borderRadius: 8, padding: "10px 14px", fontSize: 13, color: "#fca5a5" }}>{error}</div>}
      {loading && <div style={{ textAlign: "center", padding: 40, color: "#6b7280" }}>⏳ 解析中...</div>}

      <div style={{ padding: "20px 24px" }}>

        {/* ── SECTIONS ── */}
        {tab === "sections" && (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 300px", gap: 20 }}>
            <div>
              {sections.length === 0 && (
                <div style={{ background: "#1a1d2e", borderRadius: 12, border: "2px dashed #2a2d3e", padding: 40, textAlign: "center" }}>
                  <div style={{ fontSize: 32, marginBottom: 12 }}>📂</div>
                  <div style={{ fontSize: 14, color: "#6b7280", marginBottom: 16 }}>Excelをアップロードするとシートが自動認識されます</div>
                  <button onClick={() => fileRef.current.click()} style={{ background: acc, border: "none", color: "#fff", borderRadius: 8, padding: "10px 24px", fontSize: 13, cursor: "pointer", fontWeight: 600 }}>📂 ファイルを選択</button>
                </div>
              )}
              {sections.map((section, i) => {
                const meta = SHEET_TYPES[section.type] || SHEET_TYPES.other;
                const isExpanded = expandedSection === section.id;
                return (
                  <div key={section.id} style={{ background: "#1a1d2e", borderRadius: 10, marginBottom: 8, border: `1px solid ${section.enabled ? acc + "44" : "#2a2d3e"}` }}>
                    {/* Row */}
                    <div style={{ padding: "12px 14px", display: "flex", alignItems: "center", gap: 10, opacity: section.enabled ? 1 : 0.5 }}>
                      <input type="checkbox" checked={section.enabled} onChange={() => toggleSection(section.id)}
                        style={{ width: 15, height: 15, accentColor: acc, cursor: "pointer" }} />
                      <span style={{ fontSize: 17 }}>{meta.icon}</span>
                      <div style={{ flex: 1 }}>
                        <div style={{ fontSize: 13, fontWeight: 600, color: section.enabled ? "#e8eaf0" : "#6b7280" }}>{meta.label}</div>
                        <div style={{ fontSize: 11, color: "#6b7280" }}>{section.name} · {section.rows.length}行 {section.isKV ? "· キー/値形式" : "· テーブル形式"}</div>
                      </div>
                      <button onClick={() => setExpandedSection(isExpanded ? null : section.id)}
                        style={{ background: isExpanded ? acc : "#12151f", border: `1px solid ${isExpanded ? acc : "#2a2d3e"}`, color: isExpanded ? "#fff" : "#9ca3af", borderRadius: 6, padding: "5px 12px", fontSize: 12, cursor: "pointer", fontWeight: isExpanded ? 600 : 400 }}>
                        {isExpanded ? "▲ 閉じる" : "✏️ 編集"}
                      </button>
                      <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
                        <button onClick={() => moveSection(section.id, -1)} disabled={i === 0} style={{ background: "#12151f", border: "1px solid #2a2d3e", color: "#9ca3af", borderRadius: 4, width: 22, height: 18, cursor: "pointer", fontSize: 9 }}>▲</button>
                        <button onClick={() => moveSection(section.id, 1)} disabled={i === sections.length - 1} style={{ background: "#12151f", border: "1px solid #2a2d3e", color: "#9ca3af", borderRadius: 4, width: 22, height: 18, cursor: "pointer", fontSize: 9 }}>▼</button>
                      </div>
                    </div>
                    {/* Inline editor */}
                    {isExpanded && (
                      <div style={{ borderTop: "1px solid #2a2d3e", padding: 14 }}>
                        {section.type === "cost" ? (
                          <div style={{ fontSize: 12, color: "#9ca3af" }}>💰 見積金額は「見積金額タブ」で編集できます</div>
                        ) : section.isKV ? (
                          <KVEditor kvPairs={section.kvPairs} onChange={(kv) => updateSectionKV(section.id, kv)} quoteInfo={quoteInfo} costRows={costRows} markup={markup} acc={acc} />
                        ) : (
                          <TableEditor rows={section.rows} headerIdx={section.headerIdx} onChange={(rows) => updateSectionRows(section.id, rows)} />
                        )}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>

            {/* Right: quote info */}
            <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              <div style={{ background: "#1a1d2e", borderRadius: 12, padding: 16, border: "1px solid #2a2d3e" }}>
                <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 12 }}>見積書情報</div>
                {[["宛先", "clientName"], ["件名", "projectName"], ["見積番号", "quoteNo"], ["発行日", "date"], ["有効期限", "expiry"], ["開始予定日", "startDate"], ["終了予定日", "endDate"], ["開発期間", "duration"], ["チーム規模", "teamSize"]].map(([label, key]) => (
                  <div key={key} style={{ marginBottom: 9 }}>
                    <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 2 }}>{label}</div>
                    <input value={quoteInfo[key] || ""} onChange={(e) => setQuoteInfo((p) => ({ ...p, [key]: e.target.value }))}
                      style={{ width: "100%", background: "#0f1117", border: "1px solid #2a2d3e", color: "#e8eaf0", borderRadius: 6, padding: "6px 9px", fontSize: 13, boxSizing: "border-box" }} />
                  </div>
                ))}
              </div>
              <button onClick={() => setTab("preview")} style={{ background: `linear-gradient(135deg, ${acc}, ${acc}99)`, border: "none", color: "#fff", borderRadius: 10, padding: "12px 0", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>
                👁 プレビューへ →
              </button>
            </div>
          </div>
        )}

        {/* ── COST ── */}
        {tab === "cost" && (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 260px", gap: 20 }}>
            <div>
              <div style={{ background: "#1a1d2e", borderRadius: 10, padding: "10px 14px", marginBottom: 12, display: "flex", alignItems: "center", gap: 12, border: "1px solid #2a2d3e", flexWrap: "wrap" }}>
                <span style={{ fontSize: 12, color: "#9ca3af" }}>掛け率</span>
                <div style={{ display: "flex", gap: 4 }}>
                  {[1.2, 1.3, 1.5, 1.8, 2.0].map((m) => (
                    <button key={m} onClick={() => setMarkup(m)} style={{ background: markup === m ? acc : "#12151f", border: `1px solid ${markup === m ? acc : "#2a2d3e"}`, color: markup === m ? "#fff" : "#9ca3af", borderRadius: 6, padding: "4px 10px", fontSize: 12, cursor: "pointer", fontWeight: markup === m ? 700 : 400 }}>×{m}</button>
                  ))}
                  <input type="number" step="0.01" min="1" value={markup} onChange={(e) => setMarkup(parseFloat(e.target.value) || 1)} style={{ width: 56, background: "#12151f", border: "1px solid #2a2d3e", color: "#e8eaf0", borderRadius: 6, padding: "4px 7px", fontSize: 12, textAlign: "center" }} />
                </div>
                <span style={{ fontSize: 12, color: "#9ca3af", marginLeft: "auto" }}>利益率: <span style={{ color: "#4ade80", fontWeight: 700 }}>{(((markup - 1) / markup) * 100).toFixed(1)}%</span></span>
              </div>
              <div style={{ background: "#1a1d2e", borderRadius: 12, border: "1px solid #2a2d3e", overflow: "hidden" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead><tr style={{ background: "#12151f" }}>
                    {["No", "項目", "工数(人月)", "単価(JPY)", "仕入金額", "販売金額"].map((h, i) => (
                      <th key={h} style={{ padding: "9px 12px", textAlign: i < 2 ? "left" : "right", color: "#6b7280", fontWeight: 600, fontSize: 11, borderBottom: "1px solid #2a2d3e", whiteSpace: "nowrap" }}>{h}</th>
                    ))}
                  </tr></thead>
                  <tbody>
                    {costRows.map((r, idx) => {
                      const orig = r.manMonth * r.unitCost, sale = orig * markup;
                      return (
                        <tr key={idx} style={{ borderBottom: "1px solid #1e2235" }}>
                          <td style={{ padding: "8px 12px" }}><input value={r.no} onChange={(e) => updateCostRow(idx, "no", e.target.value)} style={{ width: 28, background: "transparent", border: "none", color: "#6b7280", fontSize: 13 }} /></td>
                          <td style={{ padding: "8px 12px" }}><input value={r.item} onChange={(e) => updateCostRow(idx, "item", e.target.value)} style={{ width: "100%", background: "transparent", border: "none", color: "#e8eaf0", fontSize: 13 }} /></td>
                          <td style={{ padding: "8px 12px", textAlign: "right" }}><input type="number" step="0.01" value={r.manMonth} onChange={(e) => updateCostRow(idx, "manMonth", parseFloat(e.target.value) || 0)} style={{ width: 64, background: "transparent", border: "none", color: "#e8eaf0", fontSize: 13, textAlign: "right" }} /></td>
                          <td style={{ padding: "8px 12px", textAlign: "right" }}><input type="number" step="1000" value={r.unitCost} onChange={(e) => updateCostRow(idx, "unitCost", parseInt(e.target.value) || 0)} style={{ width: 84, background: "transparent", border: "none", color: "#e8eaf0", fontSize: 13, textAlign: "right" }} /></td>
                          <td style={{ padding: "8px 12px", textAlign: "right", color: "#6b7280" }}>{fmt(orig)}</td>
                          <td style={{ padding: "8px 12px", textAlign: "right", color: "#4ade80", fontWeight: 600 }}>{fmt(sale)}</td>
                        </tr>
                      );
                    })}
                    <tr style={{ background: "#12151f", fontWeight: 700 }}>
                      <td colSpan={4} style={{ padding: "10px 12px", color: "#9ca3af", fontSize: 12 }}>合計</td>
                      <td style={{ padding: "10px 12px", textAlign: "right", color: "#6b7280" }}>{fmt(origTotal)}</td>
                      <td style={{ padding: "10px 12px", textAlign: "right", color: "#4ade80", fontSize: 14 }}>{fmt(newTotal)}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                <button onClick={() => setCostRows((p) => [...p, { no: String(p.length + 1), item: "", manMonth: 1, unitCost: 550000 }])} style={{ background: "#1a1d2e", border: "1px dashed #3a3d50", color: "#6b7280", borderRadius: 8, padding: "6px 14px", fontSize: 12, cursor: "pointer" }}>＋ 追加</button>
                {costRows.length > 1 && <button onClick={() => setCostRows((p) => p.slice(0, -1))} style={{ background: "#1a1d2e", border: "1px dashed #3a3d50", color: "#6b7280", borderRadius: 8, padding: "6px 14px", fontSize: 12, cursor: "pointer" }}>－ 削除</button>}
              </div>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              <div style={{ background: "#1a1d2e", borderRadius: 12, padding: 16, border: "1px solid #2a2d3e" }}>
                <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 10 }}>サマリー</div>
                {[["仕入れ合計", fmt(origTotal), "#9ca3af", false], ["販売合計（税抜）", fmt(newTotal), "#4ade80", true], ["消費税（10%）", fmt(newTotal * 0.1), "#9ca3af", false], ["販売合計（税込）", fmt(newTotal * 1.1), acc, true], ["粗利", fmt(newTotal - origTotal), "#60a5fa", false]].map(([l, v, c, b]) => (
                  <div key={l} style={{ display: "flex", justifyContent: "space-between", marginBottom: 7, paddingBottom: b ? 7 : 0, borderBottom: b ? "1px solid #2a2d3e" : "none" }}>
                    <span style={{ fontSize: 12, color: "#6b7280" }}>{l}</span>
                    <span style={{ fontSize: b ? 15 : 13, fontWeight: b ? 700 : 500, color: c }}>{v}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* ── PREVIEW ── */}
        {tab === "preview" && (
          <div>
            <div style={{ display: "flex", gap: 10, marginBottom: 16, justifyContent: "flex-end" }}>
              <button onClick={() => setTab("sections")} style={{ background: "#1a1d2e", border: "1px solid #3a3d50", color: "#9ca3af", borderRadius: 8, padding: "8px 18px", fontSize: 13, cursor: "pointer" }}>← 編集に戻る</button>
              <button onClick={handlePrint} style={{ background: `linear-gradient(135deg, ${acc}, ${acc}cc)`, border: "none", color: "#fff", borderRadius: 8, padding: "8px 22px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>🖨️ PDF / 印刷</button>
              <button onClick={handlePPTX} disabled={pptxGenerating} style={{ background: pptxGenerating ? "#2a2d3e" : "linear-gradient(135deg, #D04423, #C0392B)", border: "none", color: "#fff", borderRadius: 8, padding: "8px 22px", fontSize: 13, fontWeight: 700, cursor: pptxGenerating ? "wait" : "pointer", opacity: pptxGenerating ? 0.7 : 1 }}>
                {pptxGenerating ? "⏳ 生成中..." : "📊 PowerPoint出力"}
              </button>
            </div>
            <QuotePreview sections={sections.length > 0 ? sections : [{ id: "cost", name: "見積コスト", type: "cost", enabled: true, rows: [], headerIdx: 0, isKV: false, kvPairs: [] }]}
              costRows={costRows} markup={markup} settings={settings} quoteInfo={quoteInfo} />
          </div>
        )}
      </div>
    </div>
  );
}
