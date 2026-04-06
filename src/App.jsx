import { useState, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── Storage ───────────────────────────────────────────────────────
const STORAGE_KEY = "zept_quote_settings_v2";
const loadSettings = () => { try { return JSON.parse(localStorage.getItem(STORAGE_KEY)) || {}; } catch { return {}; } };
const saveSettings = (s) => localStorage.setItem(STORAGE_KEY, JSON.stringify(s));

const DEFAULT_SETTINGS = {
  companyName: "Zept合同会社", address: "", phone: "", email: "", website: "",
  logoDataUrl: "", defaultMarkup: 1.8, accentColor: "#1a56db",
  footerLines: [
    "※ 上記金額は消費税抜き価格です。お支払いの際は消費税を付加してお支払いください。",
    "※ お支払い条件：納品後30日以内",
    "※ 見積有効期限：発行日より30日間",
  ],
};

// ─── Sheet classification ──────────────────────────────────────────
const SHEET_TYPES = {
  cost:        { label: "見積コスト",     icon: "💰", keywords: ["コスト","全体工数","見積コスト","cost","費用"] },
  wbs:         { label: "WBS・工数",      icon: "📋", keywords: ["wbs","工数","開発詳細","詳細見積","task","タスク"] },
  requirement: { label: "前提条件",       icon: "📌", keywords: ["前提","requirement","条件","スコープ"] },
  schedule:    { label: "スケジュール",   icon: "📅", keywords: ["スケジュール","schedule","マスタ","マイルストーン","milestone","計画","請求"] },
  deliverable: { label: "成果物",         icon: "📦", keywords: ["成果物","deliverable","納品"] },
  overview:    { label: "概要",           icon: "📄", keywords: ["概要","overview","表紙","cover","summary"] },
  history:     { label: "変更履歴",       icon: "🔄", keywords: ["変更履歴","history","changelog","revision"] },
  license:     { label: "ライセンス費用", icon: "🔑", keywords: ["license","ライセンス","account","アカウント"] },
  function:    { label: "機能一覧",       icon: "⚙️", keywords: ["function","機能","origin","wbs","featur"] },
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

// ─── Excel full parser ─────────────────────────────────────────────
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
          // Remove fully empty rows
          const nonEmpty = raw.filter((r) => r.some((c) => c !== null && c !== undefined && String(c).trim() !== ""));
          if (nonEmpty.length < 2) continue;

          // Find meaningful content range
          const rows = nonEmpty.map((r) => r.map((c) => c !== null && c !== undefined ? String(c).trim() : ""));

          // Detect header row (longest row with non-empty cells)
          let headerIdx = 0;
          let maxFilled = 0;
          for (let i = 0; i < Math.min(rows.length, 12); i++) {
            const filled = rows[i].filter((c) => c !== "").length;
            if (filled > maxFilled) { maxFilled = filled; headerIdx = i; }
          }

          const type = classifySheet(sheetName);
          sections.push({ id: sheetName, name: sheetName, type, rows, headerIdx, enabled: type !== "other" });
        }

        // Extract cost rows from cost sheet
        const costSection = sections.find((s) => s.type === "cost");
        let costRows = [];
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

// ─── Sheet renderer ────────────────────────────────────────────────
function SheetTable({ rows, headerIdx, acc }) {
  if (!rows || rows.length === 0) return null;
  const header = rows[headerIdx] || [];
  const colCount = Math.max(...rows.map((r) => r.length));
  const dataRows = rows.slice(headerIdx + 1).filter((r) => r.some((c) => c !== ""));

  // Detect if it's more of a key-value layout (2 cols, many rows)
  const isKV = colCount <= 3 && dataRows.length > 3;

  if (isKV) {
    return (
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
        <tbody>
          {rows.filter((r) => r.some((c) => c !== "")).map((r, i) => (
            <tr key={i} style={{ borderBottom: "1px solid #f0f0f0" }}>
              {r.filter((_, ci) => ci < 4).map((cell, ci) => (
                <td key={ci} style={{ padding: "7px 12px", verticalAlign: "top", fontWeight: ci === 0 || (i === headerIdx) ? 600 : 400, background: i === headerIdx ? `${acc}15` : ci === 0 ? "#fafafa" : "#fff", color: i === headerIdx ? acc : "#333", whiteSpace: "pre-wrap", minWidth: ci === 0 ? 120 : "auto", maxWidth: 400 }}>
                  {cell}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    );
  }

  return (
    <div style={{ overflowX: "auto" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
        <thead>
          <tr style={{ background: acc }}>
            {header.filter((h) => h !== "").map((h, i) => (
              <th key={i} style={{ padding: "8px 12px", textAlign: "left", color: "#fff", fontWeight: 600, fontSize: 11, whiteSpace: "nowrap" }}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {dataRows.map((r, i) => (
            <tr key={i} style={{ background: i % 2 === 0 ? "#fafafa" : "#fff", borderBottom: "1px solid #f0f0f0" }}>
              {r.filter((_, ci) => (header[ci] !== undefined)).slice(0, header.filter((h) => h !== "").length).map((cell, ci) => (
                <td key={ci} style={{ padding: "7px 12px", verticalAlign: "top", color: "#333", whiteSpace: "pre-wrap", maxWidth: 320, fontSize: 11 }}>{cell}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// ─── Cost table ────────────────────────────────────────────────────
function CostTable({ rows, markup, acc }) {
  const computed = rows.map((r) => ({ ...r, newAmt: r.manMonth * r.unitCost * markup, newMM: parseFloat((r.manMonth * markup).toFixed(3)), newUC: Math.round(r.unitCost * markup) }));
  const total = computed.reduce((s, r) => s + r.newAmt, 0);
  const fmt = (n) => "¥ " + Math.round(n).toLocaleString("ja-JP");
  return (
    <>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
        <thead>
          <tr style={{ background: acc }}>
            {["No", "項目", "工数（人月）", "単価（JPY）", "金額（JPY）"].map((h, i) => (
              <th key={h} style={{ padding: "9px 12px", textAlign: i >= 2 ? "right" : "left", color: "#fff", fontWeight: 600, fontSize: 11 }}>{h}</th>
            ))}
          </tr>
        </thead>
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
            <td colSpan={4} style={{ padding: "8px 12px", color: "#888", fontSize: 11 }}>消費税（10%）</td>
            <td style={{ padding: "8px 12px", textAlign: "right", color: "#888", fontSize: 11 }}>{fmt(total * 0.1)}</td>
          </tr>
          <tr style={{ background: `${acc}20`, fontWeight: 700 }}>
            <td colSpan={4} style={{ padding: "10px 12px" }}>合計（税込）</td>
            <td style={{ padding: "10px 12px", textAlign: "right", color: acc, fontSize: 14 }}>{fmt(total * 1.1)}</td>
          </tr>
        </tbody>
      </table>
      <div style={{ marginTop: 10, fontSize: 11, color: "#888" }}>※ 上記金額は消費税抜き価格です</div>
    </>
  );
}

// ─── Quote Preview ─────────────────────────────────────────────────
function QuotePreview({ sections, costRows, markup, settings, quoteInfo }) {
  const acc = settings.accentColor;
  const fmt = (n) => "¥ " + Math.round(n).toLocaleString("ja-JP");
  const total = costRows.reduce((s, r) => s + r.manMonth * r.unitCost * markup, 0);
  const enabledSections = sections.filter((s) => s.enabled);

  return (
    <div id="quote-preview" style={{ background: "#fff", color: "#111", fontFamily: "'Hiragino Sans', 'Meiryo', sans-serif" }}>
      {/* Header */}
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
        {/* Title block */}
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

        {/* Client */}
        <div style={{ borderBottom: `3px solid ${acc}`, paddingBottom: 10, marginBottom: 16, display: "flex", alignItems: "baseline", gap: 10 }}>
          <span style={{ fontSize: 17, fontWeight: 700 }}>{quoteInfo.clientName}</span>
          <span style={{ fontSize: 13, color: "#666" }}>御中</span>
        </div>

        {/* Total amount */}
        <div style={{ background: "#f8f9ff", border: `1px solid ${acc}30`, borderRadius: 8, padding: "14px 22px", marginBottom: 24, display: "flex", alignItems: "baseline", gap: 14 }}>
          <span style={{ fontSize: 13, color: "#666" }}>御見積金額（税抜）</span>
          <span style={{ fontSize: 30, fontWeight: 800, color: acc }}>{fmt(total)}</span>
          <span style={{ fontSize: 12, color: "#aaa" }}>（消費税別途）</span>
        </div>

        {/* Dynamic sections */}
        {enabledSections.map((section) => {
          const meta = SHEET_TYPES[section.type] || SHEET_TYPES.other;
          return (
            <div key={section.id} style={{ marginBottom: 28 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, background: `${acc}12`, borderLeft: `4px solid ${acc}`, padding: "9px 14px", marginBottom: 12, borderRadius: "0 6px 6px 0" }}>
                <span style={{ fontSize: 15 }}>{meta.icon}</span>
                <span style={{ fontWeight: 700, fontSize: 14, color: acc }}>{meta.label}</span>
                <span style={{ fontSize: 12, color: "#888", marginLeft: 4 }}>（{section.name}）</span>
              </div>
              {section.type === "cost"
                ? <CostTable rows={costRows} markup={markup} acc={acc} />
                : <SheetTable rows={section.rows} headerIdx={section.headerIdx} acc={acc} />
              }
            </div>
          );
        })}

        {/* Footer */}
        {settings.footerLines.length > 0 && (
          <div style={{ borderTop: "1px solid #e5e7eb", paddingTop: 14, marginTop: 8 }}>
            {settings.footerLines.map((line, i) => (
              <div key={i} style={{ fontSize: 11, color: "#888", lineHeight: 2 }}>{line}</div>
            ))}
          </div>
        )}

        {/* Signature */}
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
  const [footerText, setFooterText] = useState(settings.footerLines.join("\n"));
  const handleLogo = (e) => {
    const file = e.target.files?.[0]; if (!file) return;
    const r = new FileReader();
    r.onload = (ev) => onChange({ ...settings, logoDataUrl: ev.target.result });
    r.readAsDataURL(file);
  };
  const f = (label, key, ph, type = "text") => (
    <div style={{ marginBottom: 12 }}>
      <label style={{ display: "block", fontSize: 11, color: "#9ca3af", marginBottom: 3 }}>{label}</label>
      <input type={type} placeholder={ph} value={settings[key] || ""} onChange={(e) => onChange({ ...settings, [key]: e.target.value })}
        style={{ width: "100%", background: "#0f1117", border: "1px solid #2a2d3e", color: "#e8eaf0", borderRadius: 6, padding: "7px 10px", fontSize: 13, boxSizing: "border-box" }} />
    </div>
  );
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.7)", zIndex: 1000, display: "flex", justifyContent: "flex-end" }} onClick={onClose}>
      <div style={{ width: 400, background: "#1a1d2e", height: "100%", overflowY: "auto", padding: 24, boxSizing: "border-box", borderLeft: "1px solid #2a2d3e" }} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <span style={{ fontWeight: 700, fontSize: 15 }}>⚙️ 会社設定</span>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#9ca3af", fontSize: 20, cursor: "pointer" }}>×</button>
        </div>
        <div style={{ marginBottom: 16 }}>
          <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 8 }}>ロゴ画像</div>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            {settings.logoDataUrl ? <img src={settings.logoDataUrl} alt="logo" style={{ height: 40, maxWidth: 120, objectFit: "contain", background: "#fff", borderRadius: 4, padding: 4 }} />
              : <div style={{ width: 80, height: 40, background: "#12151f", borderRadius: 4, border: "1px dashed #3a3d50", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 10, color: "#6b7280" }}>No logo</div>}
            <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
              <button onClick={() => logoRef.current.click()} style={{ background: "#1e2235", border: "1px solid #3a3d50", color: "#9ca3af", borderRadius: 6, padding: "5px 10px", fontSize: 12, cursor: "pointer" }}>📂 アップロード</button>
              {settings.logoDataUrl && <button onClick={() => onChange({ ...settings, logoDataUrl: "" })} style={{ background: "none", border: "none", color: "#ef4444", fontSize: 11, cursor: "pointer" }}>削除</button>}
            </div>
            <input ref={logoRef} type="file" accept="image/*" onChange={handleLogo} style={{ display: "none" }} />
          </div>
        </div>
        <div style={{ borderTop: "1px solid #2a2d3e", paddingTop: 14, marginBottom: 14 }}>
          {f("会社名", "companyName", "Zept合同会社")}
          {f("住所", "address", "東京都〇〇区〇〇 1-2-3")}
          {f("電話番号", "phone", "03-XXXX-XXXX")}
          {f("メールアドレス", "email", "info@zept.com", "email")}
          {f("ウェブサイト", "website", "https://zept.com")}
        </div>
        <div style={{ borderTop: "1px solid #2a2d3e", paddingTop: 14, marginBottom: 14 }}>
          <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 10 }}>デフォルト掛け率</div>
          <div style={{ display: "flex", gap: 5, flexWrap: "wrap", marginBottom: 12 }}>
            {[1.2, 1.3, 1.5, 1.8, 2.0].map((m) => (
              <button key={m} onClick={() => onChange({ ...settings, defaultMarkup: m })}
                style={{ background: settings.defaultMarkup === m ? settings.accentColor : "#12151f", border: `1px solid ${settings.defaultMarkup === m ? settings.accentColor : "#2a2d3e"}`, color: settings.defaultMarkup === m ? "#fff" : "#9ca3af", borderRadius: 6, padding: "5px 10px", fontSize: 12, cursor: "pointer" }}>×{m}</button>
            ))}
            <input type="number" step="0.01" min="1" value={settings.defaultMarkup} onChange={(e) => onChange({ ...settings, defaultMarkup: parseFloat(e.target.value) || 1.8 })}
              style={{ width: 56, background: "#12151f", border: "1px solid #2a2d3e", color: "#e8eaf0", borderRadius: 6, padding: "5px 8px", fontSize: 12, textAlign: "center" }} />
          </div>
          <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 8 }}>アクセントカラー</div>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            {["#1a56db", "#0d9488", "#7c3aed", "#dc2626", "#d97706", "#0f766e"].map((c) => (
              <div key={c} onClick={() => onChange({ ...settings, accentColor: c })}
                style={{ width: 24, height: 24, borderRadius: "50%", background: c, cursor: "pointer", border: settings.accentColor === c ? "3px solid #fff" : "3px solid transparent" }} />
            ))}
            <input type="color" value={settings.accentColor} onChange={(e) => onChange({ ...settings, accentColor: e.target.value })}
              style={{ width: 24, height: 24, borderRadius: "50%", border: "none", cursor: "pointer", padding: 0 }} />
          </div>
        </div>
        <div style={{ borderTop: "1px solid #2a2d3e", paddingTop: 14, marginBottom: 16 }}>
          <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 8 }}>フッター文言（1行ずつ）</div>
          <textarea rows={5} value={footerText} onChange={(e) => { setFooterText(e.target.value); onChange({ ...settings, footerLines: e.target.value.split("\n").filter(Boolean) }); }}
            style={{ width: "100%", background: "#0f1117", border: "1px solid #2a2d3e", color: "#e8eaf0", borderRadius: 6, padding: "8px 10px", fontSize: 12, boxSizing: "border-box", resize: "vertical" }} />
        </div>
        <button onClick={onClose} style={{ width: "100%", background: settings.accentColor, border: "none", color: "#fff", borderRadius: 8, padding: "11px 0", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>✓ 保存して閉じる</button>
      </div>
    </div>
  );
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
  });
  const [tab, setTab] = useState("sections"); // sections | cost | preview
  const [showSettings, setShowSettings] = useState(false);
  const [loading, setLoading] = useState(false);
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

  const toggleSection = (id) => setSections((prev) => prev.map((s) => s.id === id ? { ...s, enabled: !s.enabled } : s));
  const moveSection = (id, dir) => {
    setSections((prev) => {
      const idx = prev.findIndex((s) => s.id === id);
      const next = [...prev];
      const swap = idx + dir;
      if (swap < 0 || swap >= next.length) return prev;
      [next[idx], next[swap]] = [next[swap], next[idx]];
      return next;
    });
  };

  const updateCostRow = (idx, field, val) => setCostRows((p) => p.map((r, i) => i === idx ? { ...r, [field]: val } : r));
  const origTotal = costRows.reduce((s, r) => s + r.manMonth * r.unitCost, 0);
  const newTotal = costRows.reduce((s, r) => s + r.manMonth * r.unitCost * markup, 0);
  const fmt = (n) => "¥ " + Math.round(n).toLocaleString("ja-JP");
  const acc = settings.accentColor;

  const handlePrint = () => {
    const style = document.createElement("style");
    style.textContent = `@media print { body > *:not(#print-root) { display: none !important; } #print-root { display: block !important; } @page { margin: 12mm; } }`;
    const root = document.getElementById("quote-preview").cloneNode(true);
    root.id = "print-root";
    document.body.appendChild(style); document.body.appendChild(root);
    window.print();
    document.body.removeChild(style); document.body.removeChild(root);
  };

  return (
    <div style={{ minHeight: "100vh", background: "#0f1117", color: "#e8eaf0", fontFamily: "'Hiragino Sans', 'Meiryo', sans-serif" }}>
      {showSettings && <SettingsPanel settings={settings} onChange={handleSettingsChange} onClose={() => setShowSettings(false)} />}

      {/* Header */}
      <header style={{ background: "#1a1d2e", borderBottom: "1px solid #2a2d3e", padding: "13px 24px", display: "flex", alignItems: "center", gap: 12 }}>
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

      {/* Tabs */}
      <div style={{ background: "#1a1d2e", borderBottom: "1px solid #2a2d3e", padding: "0 24px", display: "flex" }}>
        {[["sections", "📑 セクション管理"], ["cost", "💰 見積金額編集"], ["preview", "👁 プレビュー / 印刷"]].map(([id, label]) => (
          <button key={id} onClick={() => setTab(id)}
            style={{ background: "none", border: "none", borderBottom: tab === id ? `2px solid ${acc}` : "2px solid transparent", color: tab === id ? "#e8eaf0" : "#6b7280", padding: "11px 18px", fontSize: 13, cursor: "pointer", fontWeight: tab === id ? 600 : 400 }}>
            {label}
          </button>
        ))}
      </div>

      {error && <div style={{ background: "#2d1515", border: "1px solid #7f1d1d", margin: "12px 24px", borderRadius: 8, padding: "10px 14px", fontSize: 13, color: "#fca5a5" }}>{error}</div>}
      {loading && <div style={{ textAlign: "center", padding: 40, color: "#6b7280" }}>⏳ Excelを解析中...</div>}

      <div style={{ padding: "20px 24px" }}>

        {/* ── SECTION MANAGER ── */}
        {tab === "sections" && (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 320px", gap: 20 }}>
            <div>
              <div style={{ fontSize: 13, color: "#9ca3af", marginBottom: 14 }}>
                {sections.length === 0
                  ? "Excelをアップロードするとシートが自動認識されます。各セクションのON/OFFと順序を変更できます。"
                  : `${sections.length}シートを検出。チェックしたセクションが見積書に含まれます。`}
              </div>
              {sections.length === 0 && (
                <div style={{ background: "#1a1d2e", borderRadius: 12, border: "2px dashed #2a2d3e", padding: 40, textAlign: "center" }}>
                  <div style={{ fontSize: 32, marginBottom: 12 }}>📂</div>
                  <div style={{ fontSize: 14, color: "#6b7280", marginBottom: 16 }}>右上の「Excelを読み込む」から<br />仕入れ見積のExcelをアップロード</div>
                  <button onClick={() => fileRef.current.click()} style={{ background: acc, border: "none", color: "#fff", borderRadius: 8, padding: "10px 24px", fontSize: 13, cursor: "pointer", fontWeight: 600 }}>📂 ファイルを選択</button>
                </div>
              )}
              {sections.map((section, i) => {
                const meta = SHEET_TYPES[section.type] || SHEET_TYPES.other;
                return (
                  <div key={section.id} style={{ background: "#1a1d2e", borderRadius: 10, padding: "13px 16px", marginBottom: 8, border: `1px solid ${section.enabled ? acc + "44" : "#2a2d3e"}`, display: "flex", alignItems: "center", gap: 12, opacity: section.enabled ? 1 : 0.5 }}>
                    <input type="checkbox" checked={section.enabled} onChange={() => toggleSection(section.id)}
                      style={{ width: 16, height: 16, accentColor: acc, cursor: "pointer" }} />
                    <span style={{ fontSize: 18 }}>{meta.icon}</span>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontSize: 13, fontWeight: 600, color: section.enabled ? "#e8eaf0" : "#6b7280" }}>{meta.label}</div>
                      <div style={{ fontSize: 11, color: "#6b7280", marginTop: 2 }}>{section.name} · {section.rows.length}行</div>
                    </div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
                      <button onClick={() => moveSection(section.id, -1)} disabled={i === 0} style={{ background: "#12151f", border: "1px solid #2a2d3e", color: "#9ca3af", borderRadius: 4, width: 24, height: 20, cursor: "pointer", fontSize: 10, display: "flex", alignItems: "center", justifyContent: "center" }}>▲</button>
                      <button onClick={() => moveSection(section.id, 1)} disabled={i === sections.length - 1} style={{ background: "#12151f", border: "1px solid #2a2d3e", color: "#9ca3af", borderRadius: 4, width: 24, height: 20, cursor: "pointer", fontSize: 10, display: "flex", alignItems: "center", justifyContent: "center" }}>▼</button>
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Right panel */}
            <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
              <div style={{ background: "#1a1d2e", borderRadius: 12, padding: 18, border: "1px solid #2a2d3e" }}>
                <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 12 }}>見積書情報</div>
                {[["宛先", "clientName"], ["件名", "projectName"], ["見積番号", "quoteNo"], ["発行日", "date"], ["有効期限", "expiry"]].map(([label, key]) => (
                  <div key={key} style={{ marginBottom: 10 }}>
                    <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 3 }}>{label}</div>
                    <input value={quoteInfo[key]} onChange={(e) => setQuoteInfo((p) => ({ ...p, [key]: e.target.value }))}
                      style={{ width: "100%", background: "#0f1117", border: "1px solid #2a2d3e", color: "#e8eaf0", borderRadius: 6, padding: "7px 10px", fontSize: 13, boxSizing: "border-box" }} />
                  </div>
                ))}
              </div>
              <button onClick={() => setTab("preview")} style={{ background: `linear-gradient(135deg, ${acc}, ${acc}99)`, border: "none", color: "#fff", borderRadius: 10, padding: "12px 0", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>
                👁 プレビューへ →
              </button>
            </div>
          </div>
        )}

        {/* ── COST EDITOR ── */}
        {tab === "cost" && (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 280px", gap: 20 }}>
            <div>
              <div style={{ background: "#1a1d2e", borderRadius: 10, padding: "11px 16px", marginBottom: 14, display: "flex", alignItems: "center", gap: 14, border: "1px solid #2a2d3e", flexWrap: "wrap" }}>
                <span style={{ fontSize: 12, color: "#9ca3af" }}>掛け率</span>
                <div style={{ display: "flex", gap: 5 }}>
                  {[1.2, 1.3, 1.5, 1.8, 2.0].map((m) => (
                    <button key={m} onClick={() => setMarkup(m)} style={{ background: markup === m ? acc : "#12151f", border: `1px solid ${markup === m ? acc : "#2a2d3e"}`, color: markup === m ? "#fff" : "#9ca3af", borderRadius: 6, padding: "4px 10px", fontSize: 12, cursor: "pointer", fontWeight: markup === m ? 700 : 400 }}>×{m}</button>
                  ))}
                  <input type="number" step="0.01" min="1" value={markup} onChange={(e) => setMarkup(parseFloat(e.target.value) || 1)} style={{ width: 58, background: "#12151f", border: "1px solid #2a2d3e", color: "#e8eaf0", borderRadius: 6, padding: "4px 8px", fontSize: 12, textAlign: "center" }} />
                </div>
                <span style={{ fontSize: 12, color: "#9ca3af", marginLeft: "auto" }}>利益率: <span style={{ color: "#4ade80", fontWeight: 700 }}>{(((markup - 1) / markup) * 100).toFixed(1)}%</span></span>
              </div>
              <div style={{ background: "#1a1d2e", borderRadius: 12, border: "1px solid #2a2d3e", overflow: "hidden" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                  <thead>
                    <tr style={{ background: "#12151f" }}>
                      {["No", "項目", "工数(人月)", "単価(JPY)", "仕入金額", "販売金額"].map((h, i) => (
                        <th key={h} style={{ padding: "10px 12px", textAlign: i < 2 ? "left" : "right", color: "#6b7280", fontWeight: 600, fontSize: 11, borderBottom: "1px solid #2a2d3e", whiteSpace: "nowrap" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {costRows.map((r, idx) => {
                      const orig = r.manMonth * r.unitCost, sale = orig * markup;
                      return (
                        <tr key={idx} style={{ borderBottom: "1px solid #1e2235" }}>
                          <td style={{ padding: "8px 12px" }}><input value={r.no} onChange={(e) => updateCostRow(idx, "no", e.target.value)} style={{ width: 30, background: "transparent", border: "none", color: "#6b7280", fontSize: 13 }} /></td>
                          <td style={{ padding: "8px 12px" }}><input value={r.item} onChange={(e) => updateCostRow(idx, "item", e.target.value)} style={{ width: "100%", background: "transparent", border: "none", color: "#e8eaf0", fontSize: 13 }} /></td>
                          <td style={{ padding: "8px 12px", textAlign: "right" }}><input type="number" step="0.01" value={r.manMonth} onChange={(e) => updateCostRow(idx, "manMonth", parseFloat(e.target.value) || 0)} style={{ width: 65, background: "transparent", border: "none", color: "#e8eaf0", fontSize: 13, textAlign: "right" }} /></td>
                          <td style={{ padding: "8px 12px", textAlign: "right" }}><input type="number" step="1000" value={r.unitCost} onChange={(e) => updateCostRow(idx, "unitCost", parseInt(e.target.value) || 0)} style={{ width: 85, background: "transparent", border: "none", color: "#e8eaf0", fontSize: 13, textAlign: "right" }} /></td>
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
                <button onClick={() => setCostRows((p) => [...p, { no: String(p.length + 1), item: "", manMonth: 1, unitCost: 550000 }])} style={{ background: "#1a1d2e", border: "1px dashed #3a3d50", color: "#6b7280", borderRadius: 8, padding: "6px 14px", fontSize: 12, cursor: "pointer" }}>＋ 行を追加</button>
                {costRows.length > 1 && <button onClick={() => setCostRows((p) => p.slice(0, -1))} style={{ background: "#1a1d2e", border: "1px dashed #3a3d50", color: "#6b7280", borderRadius: 8, padding: "6px 14px", fontSize: 12, cursor: "pointer" }}>－ 最終行を削除</button>}
              </div>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              <div style={{ background: "#1a1d2e", borderRadius: 12, padding: 16, border: "1px solid #2a2d3e" }}>
                <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 12 }}>サマリー</div>
                {[["仕入れ合計", fmt(origTotal), "#9ca3af", false], ["販売合計（税抜）", fmt(newTotal), "#4ade80", true], ["消費税（10%）", fmt(newTotal * 0.1), "#9ca3af", false], ["販売合計（税込）", fmt(newTotal * 1.1), acc, true], ["粗利", fmt(newTotal - origTotal), "#60a5fa", false]].map(([l, v, c, b]) => (
                  <div key={l} style={{ display: "flex", justifyContent: "space-between", marginBottom: 8, paddingBottom: b ? 8 : 0, borderBottom: b ? "1px solid #2a2d3e" : "none" }}>
                    <span style={{ fontSize: 12, color: "#6b7280" }}>{l}</span>
                    <span style={{ fontSize: b ? 15 : 13, fontWeight: b ? 700 : 500, color: c }}>{v}</span>
                  </div>
                ))}
              </div>
              <button onClick={() => setTab("preview")} style={{ background: `linear-gradient(135deg, ${acc}, ${acc}99)`, border: "none", color: "#fff", borderRadius: 10, padding: "12px 0", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>
                👁 プレビューへ →
              </button>
            </div>
          </div>
        )}

        {/* ── PREVIEW ── */}
        {tab === "preview" && (
          <div>
            <div style={{ display: "flex", gap: 10, marginBottom: 16, justifyContent: "flex-end" }}>
              <button onClick={() => setTab("sections")} style={{ background: "#1a1d2e", border: "1px solid #3a3d50", color: "#9ca3af", borderRadius: 8, padding: "8px 18px", fontSize: 13, cursor: "pointer" }}>← 編集に戻る</button>
              <button onClick={handlePrint} style={{ background: `linear-gradient(135deg, ${acc}, ${acc}cc)`, border: "none", color: "#fff", borderRadius: 8, padding: "8px 22px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>🖨️ PDF / 印刷</button>
            </div>
            <QuotePreview sections={sections.length > 0 ? sections : [{ id: "cost", name: "見積コスト", type: "cost", enabled: true, rows: [], headerIdx: 0 }]}
              costRows={costRows} markup={markup} settings={settings} quoteInfo={quoteInfo} />
          </div>
        )}
      </div>
    </div>
  );
}
