import { useState, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── Storage helpers ───────────────────────────────────────────────
const STORAGE_KEY = "zept_quote_settings";
const loadSettings = () => {
  try { return JSON.parse(localStorage.getItem(STORAGE_KEY)) || {}; } catch { return {}; }
};
const saveSettings = (s) => localStorage.setItem(STORAGE_KEY, JSON.stringify(s));

const DEFAULT_SETTINGS = {
  companyName: "Zept合同会社",
  address: "",
  phone: "",
  email: "",
  website: "",
  logoDataUrl: "",
  defaultMarkup: 1.8,
  footerLines: [
    "※ 上記金額は消費税抜き価格です。お支払いの際は消費税を付加してお支払いください。",
    "※ お支払い条件：納品後30日以内",
    "※ 見積有効期限：発行日より30日間",
  ],
  accentColor: "#1a56db",
};

const DEFAULT_ROWS = [
  { no: "1", item: "要件の明確化", manMonth: 0.95, unitCost: 550000 },
  { no: "2", item: "開発", manMonth: 3.95, unitCost: 550000 },
  { no: "3", item: "テスト", manMonth: 1.55, unitCost: 500000 },
  { no: "4", item: "UAT Go-Live サポート", manMonth: 1.25, unitCost: 550000 },
  { no: "5", item: "BrSE", manMonth: 1.155, unitCost: 800000 },
  { no: "6", item: "管理", manMonth: 1.155, unitCost: 550000 },
];

const fmt = (n) => "¥ " + Math.round(n).toLocaleString("ja-JP");

// ─── Excel parser ──────────────────────────────────────────────────
function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const priority = wb.SheetNames.slice().sort((a, b) => {
          const score = (n) => {
            if (n.includes("コスト") || n.includes("全体工数") || n.includes("見積")) return 3;
            if (n.includes("工数") || n.includes("cost")) return 2;
            return 0;
          };
          return score(b) - score(a);
        });
        let headerIdx = -1, noCol, itemCol, manMonthCol, costCol, amtCol, rows;
        for (const sn of priority) {
          rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: null });
          for (let i = 0; i < rows.length; i++) {
            const r = rows[i].map((c) => (c ? String(c).replace(/\n/g, "").trim() : ""));
            const noI = r.findIndex((c) => c === "No" || c === "NO" || c === "#");
            const mmI = r.findIndex((c) => c.includes("工数") || c.includes("人月") || c.includes("MM"));
            const costI = r.findIndex((c) => (c.includes("コスト") || c.includes("単価")) && !c.includes("合計"));
            const amtI = r.findIndex((c) => c.includes("金額") || c.includes("コスト") || c.includes("amount"));
            const itemI = r.findIndex((c) => c.includes("項目") || c.includes("タスク") || c.includes("役割") || c.includes("item") || c.includes("Category"));
            if (mmI !== -1 && (itemI !== -1 || noI !== -1)) {
              headerIdx = i; noCol = noI >= 0 ? noI : 0; itemCol = itemI >= 0 ? itemI : noI + 1;
              manMonthCol = mmI; costCol = costI >= 0 ? costI : mmI + 1;
              amtCol = amtI >= 0 ? amtI : costI >= 0 ? costI + 1 : mmI + 2;
              break;
            }
          }
          if (headerIdx !== -1) break;
        }
        if (headerIdx === -1) { reject(new Error("見積コストのシートが見つかりませんでした")); return; }
        const parsed = [];
        for (let i = headerIdx + 1; i < rows.length; i++) {
          const r = rows[i];
          const no = r[noCol], item = r[itemCol];
          const mm = parseFloat(r[manMonthCol]);
          const ucRaw = parseFloat(r[costCol]), amtRaw = parseFloat(r[amtCol]);
          if (!item || isNaN(mm) || mm <= 0) continue;
          if (String(item).includes("合計") || String(item).includes("注記") || String(item).includes("TOTAL")) continue;
          const unitCost = !isNaN(ucRaw) ? ucRaw : (!isNaN(amtRaw) && mm > 0 ? Math.round(amtRaw / mm) : 0);
          parsed.push({ no: String(no ?? parsed.length + 1), item: String(item).trim(), manMonth: mm, unitCost });
        }
        if (parsed.length === 0) reject(new Error("データ行が見つかりませんでした"));
        else resolve(parsed);
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(new Error("ファイル読み込みエラー"));
    reader.readAsArrayBuffer(file);
  });
}

// ─── Settings Panel ────────────────────────────────────────────────
function SettingsPanel({ settings, onChange, onClose }) {
  const logoRef = useRef();
  const [footerText, setFooterText] = useState(settings.footerLines.join("\n"));

  const handleLogo = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => onChange({ ...settings, logoDataUrl: ev.target.result });
    reader.readAsDataURL(file);
  };

  const field = (label, key, placeholder, type = "text") => (
    <div style={{ marginBottom: 14 }}>
      <label style={{ display: "block", fontSize: 11, color: "#9ca3af", marginBottom: 4, letterSpacing: "0.05em" }}>{label}</label>
      <input type={type} placeholder={placeholder} value={settings[key] || ""} onChange={(e) => onChange({ ...settings, [key]: e.target.value })}
        style={{ width: "100%", background: "#0f1117", border: "1px solid #2a2d3e", color: "#e8eaf0", borderRadius: 6, padding: "8px 10px", fontSize: 13, boxSizing: "border-box" }} />
    </div>
  );

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.7)", zIndex: 1000, display: "flex", justifyContent: "flex-end" }} onClick={onClose}>
      <div style={{ width: 420, background: "#1a1d2e", height: "100%", overflowY: "auto", padding: 28, boxSizing: "border-box", borderLeft: "1px solid #2a2d3e" }} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 24 }}>
          <div style={{ fontWeight: 700, fontSize: 16 }}>⚙️ 会社設定</div>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#9ca3af", fontSize: 20, cursor: "pointer" }}>×</button>
        </div>

        {/* Logo */}
        <div style={{ marginBottom: 20 }}>
          <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 8, letterSpacing: "0.05em" }}>ロゴ画像</div>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            {settings.logoDataUrl
              ? <img src={settings.logoDataUrl} alt="logo" style={{ height: 48, maxWidth: 120, objectFit: "contain", background: "#fff", borderRadius: 6, padding: 4 }} />
              : <div style={{ width: 80, height: 48, background: "#12151f", borderRadius: 6, border: "1px dashed #3a3d50", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, color: "#6b7280" }}>No logo</div>
            }
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              <button onClick={() => logoRef.current.click()} style={{ background: "#1e2235", border: "1px solid #3a3d50", color: "#9ca3af", borderRadius: 6, padding: "6px 12px", fontSize: 12, cursor: "pointer" }}>
                📂 アップロード
              </button>
              {settings.logoDataUrl && (
                <button onClick={() => onChange({ ...settings, logoDataUrl: "" })} style={{ background: "none", border: "none", color: "#ef4444", fontSize: 11, cursor: "pointer", textAlign: "left" }}>削除</button>
              )}
            </div>
            <input ref={logoRef} type="file" accept="image/*" onChange={handleLogo} style={{ display: "none" }} />
          </div>
        </div>

        <div style={{ borderTop: "1px solid #2a2d3e", paddingTop: 18, marginBottom: 18 }}>
          <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 14, letterSpacing: "0.05em" }}>会社情報</div>
          {field("会社名", "companyName", "Zept合同会社")}
          {field("住所", "address", "東京都〇〇区〇〇 1-2-3")}
          {field("電話番号", "phone", "03-XXXX-XXXX")}
          {field("メールアドレス", "email", "info@zept.com", "email")}
          {field("ウェブサイト", "website", "https://zept.com")}
        </div>

        <div style={{ borderTop: "1px solid #2a2d3e", paddingTop: 18, marginBottom: 18 }}>
          <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 14, letterSpacing: "0.05em" }}>見積書の設定</div>
          <div style={{ marginBottom: 14 }}>
            <label style={{ display: "block", fontSize: 11, color: "#9ca3af", marginBottom: 4 }}>デフォルト掛け率</label>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              {[1.2, 1.3, 1.5, 1.8, 2.0].map((m) => (
                <button key={m} onClick={() => onChange({ ...settings, defaultMarkup: m })}
                  style={{ background: settings.defaultMarkup === m ? "#1a56db" : "#12151f", border: "1px solid " + (settings.defaultMarkup === m ? "#1a56db" : "#2a2d3e"), color: settings.defaultMarkup === m ? "#fff" : "#9ca3af", borderRadius: 6, padding: "5px 12px", fontSize: 12, cursor: "pointer" }}>
                  ×{m}
                </button>
              ))}
              <input type="number" step="0.01" min="1" value={settings.defaultMarkup}
                onChange={(e) => onChange({ ...settings, defaultMarkup: parseFloat(e.target.value) || 1.8 })}
                style={{ width: 64, background: "#12151f", border: "1px solid #2a2d3e", color: "#e8eaf0", borderRadius: 6, padding: "5px 8px", fontSize: 12, textAlign: "center" }} />
            </div>
          </div>
          <div style={{ marginBottom: 14 }}>
            <label style={{ display: "block", fontSize: 11, color: "#9ca3af", marginBottom: 4 }}>アクセントカラー</label>
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              {["#1a56db", "#0d9488", "#7c3aed", "#dc2626", "#d97706"].map((c) => (
                <div key={c} onClick={() => onChange({ ...settings, accentColor: c })}
                  style={{ width: 28, height: 28, borderRadius: "50%", background: c, cursor: "pointer", border: settings.accentColor === c ? "3px solid #fff" : "3px solid transparent" }} />
              ))}
              <input type="color" value={settings.accentColor} onChange={(e) => onChange({ ...settings, accentColor: e.target.value })}
                style={{ width: 28, height: 28, borderRadius: "50%", border: "none", cursor: "pointer", padding: 0, background: "none" }} />
            </div>
          </div>
        </div>

        <div style={{ borderTop: "1px solid #2a2d3e", paddingTop: 18, marginBottom: 18 }}>
          <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 10, letterSpacing: "0.05em" }}>フッター文言（1行ずつ入力）</div>
          <textarea rows={5} value={footerText}
            onChange={(e) => { setFooterText(e.target.value); onChange({ ...settings, footerLines: e.target.value.split("\n").filter(Boolean) }); }}
            style={{ width: "100%", background: "#0f1117", border: "1px solid #2a2d3e", color: "#e8eaf0", borderRadius: 6, padding: "8px 10px", fontSize: 12, boxSizing: "border-box", resize: "vertical", lineHeight: 1.7 }} />
        </div>

        <button onClick={onClose} style={{ width: "100%", background: "#1a56db", border: "none", color: "#fff", borderRadius: 8, padding: "12px 0", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>
          ✓ 保存して閉じる
        </button>
      </div>
    </div>
  );
}

// ─── Quote Preview (print target) ─────────────────────────────────
function QuotePreview({ rows, markup, settings, quoteInfo }) {
  const computedRows = rows.map((r) => ({ ...r, newAmt: r.manMonth * r.unitCost * markup, newManMonth: parseFloat((r.manMonth * markup).toFixed(3)), newUnitCost: Math.round(r.unitCost * markup) }));
  const newTotal = computedRows.reduce((s, r) => s + r.newAmt, 0);
  const acc = settings.accentColor;

  return (
    <div id="quote-preview" style={{ background: "#fff", color: "#111", fontFamily: "'Hiragino Sans', 'Meiryo', sans-serif", maxWidth: 800, margin: "0 auto", boxShadow: "0 4px 40px rgba(0,0,0,0.15)", borderRadius: 4 }}>
      {/* Header bar */}
      <div style={{ background: acc, padding: "20px 36px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div>
          {settings.logoDataUrl
            ? <img src={settings.logoDataUrl} alt="logo" style={{ height: 44, maxWidth: 160, objectFit: "contain", filter: "brightness(0) invert(1)" }} />
            : <div style={{ color: "#fff", fontWeight: 800, fontSize: 22, letterSpacing: "0.04em" }}>{settings.companyName}</div>
          }
        </div>
        <div style={{ textAlign: "right", color: "rgba(255,255,255,0.85)", fontSize: 12, lineHeight: 1.8 }}>
          {settings.address && <div>{settings.address}</div>}
          {settings.phone && <div>TEL: {settings.phone}</div>}
          {settings.email && <div>{settings.email}</div>}
          {settings.website && <div>{settings.website}</div>}
        </div>
      </div>

      <div style={{ padding: "32px 36px" }}>
        {/* Title */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 28 }}>
          <div>
            <div style={{ fontSize: 26, fontWeight: 800, letterSpacing: "0.12em", marginBottom: 6 }}>御　見　積　書</div>
            <div style={{ fontSize: 14, color: "#555", fontWeight: 600 }}>{quoteInfo.projectName}</div>
          </div>
          <div style={{ textAlign: "right", fontSize: 13, color: "#555", lineHeight: 2 }}>
            <div>見積番号：{quoteInfo.quoteNo}</div>
            <div>発行日：{quoteInfo.date}</div>
            <div>有効期限：{quoteInfo.expiry}</div>
          </div>
        </div>

        {/* Client */}
        <div style={{ borderBottom: `3px solid ${acc}`, paddingBottom: 12, marginBottom: 20, display: "flex", alignItems: "baseline", gap: 12 }}>
          <span style={{ fontSize: 18, fontWeight: 700 }}>{quoteInfo.clientName}</span>
          <span style={{ fontSize: 14, color: "#666" }}>御中</span>
        </div>

        {/* Total amount highlight */}
        <div style={{ background: "#f8f9ff", border: `1px solid ${acc}22`, borderRadius: 8, padding: "16px 24px", marginBottom: 24, display: "flex", alignItems: "baseline", gap: 16 }}>
          <span style={{ fontSize: 13, color: "#666" }}>御見積金額（税抜）</span>
          <span style={{ fontSize: 32, fontWeight: 800, color: acc, letterSpacing: "0.02em" }}>{fmt(newTotal)}</span>
          <span style={{ fontSize: 13, color: "#999" }}>（消費税別途）</span>
        </div>

        {/* Table */}
        <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 20, fontSize: 13 }}>
          <thead>
            <tr style={{ background: acc, color: "#fff" }}>
              {["No", "項目", "工数（人月）", "単価（JPY）", "金額（JPY）"].map((h, i) => (
                <th key={h} style={{ padding: "10px 14px", textAlign: i >= 2 ? "right" : "left", fontWeight: 600, fontSize: 12, whiteSpace: "nowrap" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {computedRows.map((r, i) => (
              <tr key={i} style={{ background: i % 2 === 0 ? "#fafafa" : "#fff", borderBottom: "1px solid #e5e7eb" }}>
                <td style={{ padding: "10px 14px", color: "#888", fontSize: 12 }}>{r.no}</td>
                <td style={{ padding: "10px 14px", fontWeight: 500 }}>{r.item}</td>
                <td style={{ padding: "10px 14px", textAlign: "right" }}>{r.newManMonth}</td>
                <td style={{ padding: "10px 14px", textAlign: "right" }}>{r.newUnitCost.toLocaleString()}</td>
                <td style={{ padding: "10px 14px", textAlign: "right", fontWeight: 600 }}>{fmt(r.newAmt)}</td>
              </tr>
            ))}
            <tr style={{ background: `${acc}11`, fontWeight: 700, borderTop: `2px solid ${acc}` }}>
              <td colSpan={4} style={{ padding: "12px 14px", fontSize: 14 }}>合　計（税抜）</td>
              <td style={{ padding: "12px 14px", textAlign: "right", fontSize: 16, color: acc }}>{fmt(newTotal)}</td>
            </tr>
            <tr style={{ background: "#fafafa" }}>
              <td colSpan={4} style={{ padding: "10px 14px", fontSize: 12, color: "#888" }}>消費税（10%）</td>
              <td style={{ padding: "10px 14px", textAlign: "right", fontSize: 13, color: "#888" }}>{fmt(newTotal * 0.1)}</td>
            </tr>
            <tr style={{ background: `${acc}18`, fontWeight: 700 }}>
              <td colSpan={4} style={{ padding: "12px 14px", fontSize: 14 }}>合　計（税込）</td>
              <td style={{ padding: "12px 14px", textAlign: "right", fontSize: 16, color: acc }}>{fmt(newTotal * 1.1)}</td>
            </tr>
          </tbody>
        </table>

        {/* Footer */}
        {settings.footerLines.length > 0 && (
          <div style={{ borderTop: "1px solid #e5e7eb", paddingTop: 16, marginTop: 8 }}>
            {settings.footerLines.map((line, i) => (
              <div key={i} style={{ fontSize: 11, color: "#888", lineHeight: 2 }}>{line}</div>
            ))}
          </div>
        )}

        {/* Signature block */}
        <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 32 }}>
          <div style={{ textAlign: "center", border: "1px solid #ddd", borderRadius: 4, padding: "12px 32px", fontSize: 12, color: "#666" }}>
            <div style={{ marginBottom: 4, fontWeight: 600 }}>{settings.companyName}</div>
            <div style={{ color: "#bbb", fontSize: 11 }}>担当者</div>
            <div style={{ height: 40 }} />
          </div>
        </div>
      </div>

      {/* Footer stripe */}
      <div style={{ background: acc, height: 6, borderRadius: "0 0 4px 4px" }} />
    </div>
  );
}

// ─── Main App ──────────────────────────────────────────────────────
export default function App() {
  const saved = loadSettings();
  const [settings, setSettings] = useState({ ...DEFAULT_SETTINGS, ...saved });
  const [rows, setRows] = useState(DEFAULT_ROWS);
  const [markup, setMarkup] = useState(saved.defaultMarkup || DEFAULT_SETTINGS.defaultMarkup);
  const [quoteInfo, setQuoteInfo] = useState({
    clientName: "株式会社エンケイ",
    projectName: "部品の設計図作成システム 1次フェーズ",
    quoteNo: "QT-" + new Date().toISOString().slice(0, 10).replace(/-/g, ""),
    date: new Date().toLocaleDateString("ja-JP", { year: "numeric", month: "2-digit", day: "2-digit" }),
    expiry: new Date(Date.now() + 30 * 86400000).toLocaleDateString("ja-JP", { year: "numeric", month: "2-digit", day: "2-digit" }),
  });
  const [tab, setTab] = useState("edit"); // edit | preview
  const [showSettings, setShowSettings] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const fileRef = useRef();

  const handleSettingsChange = (s) => { setSettings(s); saveSettings(s); };

  const handleFile = useCallback(async (e) => {
    const file = e.target.files?.[0]; if (!file) return;
    setLoading(true); setError("");
    try { const parsed = await parseExcelFile(file); setRows(parsed); setTab("edit"); }
    catch (err) { setError("❌ " + err.message); }
    finally { setLoading(false); e.target.value = ""; }
  }, []);

  const updateRow = (idx, field, val) => setRows((p) => p.map((r, i) => i === idx ? { ...r, [field]: val } : r));

  const computedRows = rows.map((r) => ({ ...r, origAmt: r.manMonth * r.unitCost, newAmt: r.manMonth * r.unitCost * markup, newManMonth: parseFloat((r.manMonth * markup).toFixed(3)), newUnitCost: Math.round(r.unitCost * markup) }));
  const origTotal = computedRows.reduce((s, r) => s + r.origAmt, 0);
  const newTotal = computedRows.reduce((s, r) => s + r.newAmt, 0);

  const handlePrint = () => {
    const style = document.createElement("style");
    style.textContent = `@media print { body > *:not(#print-root) { display: none !important; } #print-root { display: block !important; } @page { margin: 10mm; } }`;
    const root = document.getElementById("quote-preview").cloneNode(true);
    root.id = "print-root";
    document.body.appendChild(style);
    document.body.appendChild(root);
    window.print();
    document.body.removeChild(style);
    document.body.removeChild(root);
  };

  const exportExcel = () => {
    const header = ["No", "項目", "工数 (人月)", "単価 (JPY)", "金額 (JPY)"];
    const data = [header, ...computedRows.map((r) => [r.no, r.item, r.newManMonth, r.newUnitCost, Math.round(r.newAmt)]),
      ["合計", "", parseFloat(computedRows.reduce((s, r) => s + r.newManMonth, 0).toFixed(2)), "", Math.round(newTotal)]];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws["!cols"] = [{ wch: 6 }, { wch: 24 }, { wch: 12 }, { wch: 14 }, { wch: 14 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "見積コスト");
    XLSX.writeFile(wb, `御見積書_${quoteInfo.clientName}_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const acc = settings.accentColor;

  return (
    <div style={{ minHeight: "100vh", background: "#0f1117", color: "#e8eaf0", fontFamily: "'Hiragino Sans', 'Meiryo', sans-serif" }}>
      {showSettings && <SettingsPanel settings={settings} onChange={handleSettingsChange} onClose={() => setShowSettings(false)} />}

      {/* Header */}
      <header style={{ background: "#1a1d2e", borderBottom: "1px solid #2a2d3e", padding: "14px 28px", display: "flex", alignItems: "center", gap: 14 }}>
        {settings.logoDataUrl
          ? <img src={settings.logoDataUrl} alt="logo" style={{ height: 32, maxWidth: 100, objectFit: "contain" }} />
          : <div style={{ fontWeight: 800, fontSize: 16, color: acc }}>{settings.companyName}</div>
        }
        <div style={{ width: 1, height: 24, background: "#2a2d3e" }} />
        <div style={{ fontSize: 13, color: "#9ca3af" }}>見積書コンバーター</div>
        <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
          <button onClick={() => fileRef.current.click()} style={{ background: "#1e2235", border: "1px solid #3a3d50", color: "#9ca3af", borderRadius: 8, padding: "7px 14px", fontSize: 13, cursor: "pointer" }}>
            📂 Excelを読み込む
          </button>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
          <button onClick={() => setShowSettings(true)} style={{ background: "#1e2235", border: "1px solid #3a3d50", color: "#9ca3af", borderRadius: 8, padding: "7px 14px", fontSize: 13, cursor: "pointer" }}>
            ⚙️ 会社設定
          </button>
        </div>
      </header>

      {/* Tabs */}
      <div style={{ background: "#1a1d2e", borderBottom: "1px solid #2a2d3e", padding: "0 28px", display: "flex", gap: 0 }}>
        {[["edit", "✏️ 編集"], ["preview", "👁 プレビュー / 印刷"]].map(([id, label]) => (
          <button key={id} onClick={() => setTab(id)} style={{ background: "none", border: "none", borderBottom: tab === id ? `2px solid ${acc}` : "2px solid transparent", color: tab === id ? "#e8eaf0" : "#6b7280", padding: "12px 20px", fontSize: 13, cursor: "pointer", fontWeight: tab === id ? 600 : 400 }}>
            {label}
          </button>
        ))}
      </div>

      {error && <div style={{ background: "#2d1515", border: "1px solid #7f1d1d", margin: "16px 28px", borderRadius: 8, padding: "12px 16px", fontSize: 13, color: "#fca5a5" }}>{error}</div>}
      {loading && <div style={{ textAlign: "center", padding: 40, color: "#6b7280" }}>⏳ Excelを解析中...</div>}

      <div style={{ padding: "20px 28px" }}>
        {tab === "edit" && (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 300px", gap: 20 }}>
            {/* Left: Quote info + table */}
            <div>
              {/* Quote info */}
              <div style={{ background: "#1a1d2e", borderRadius: 12, padding: "16px 20px", marginBottom: 16, border: "1px solid #2a2d3e", display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                <div style={{ fontSize: 11, color: "#6b7280", gridColumn: "1/-1", letterSpacing: "0.05em", marginBottom: 4 }}>見積書情報</div>
                {[["宛先（会社名）", "clientName"], ["件名", "projectName"], ["見積番号", "quoteNo"], ["発行日", "date"], ["有効期限", "expiry"]].map(([label, key]) => (
                  <div key={key} style={key === "projectName" ? { gridColumn: "1/-1" } : {}}>
                    <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 3 }}>{label}</div>
                    <input value={quoteInfo[key]} onChange={(e) => setQuoteInfo((p) => ({ ...p, [key]: e.target.value }))}
                      style={{ width: "100%", background: "#0f1117", border: "1px solid #2a2d3e", color: "#e8eaf0", borderRadius: 6, padding: "7px 10px", fontSize: 13, boxSizing: "border-box" }} />
                  </div>
                ))}
              </div>

              {/* Markup bar */}
              <div style={{ background: "#1a1d2e", borderRadius: 10, padding: "12px 18px", marginBottom: 14, display: "flex", alignItems: "center", gap: 16, border: "1px solid #2a2d3e", flexWrap: "wrap" }}>
                <span style={{ fontSize: 12, color: "#9ca3af", whiteSpace: "nowrap" }}>掛け率</span>
                <div style={{ display: "flex", gap: 5 }}>
                  {[1.2, 1.3, 1.5, 1.8, 2.0].map((m) => (
                    <button key={m} onClick={() => setMarkup(m)} style={{ background: markup === m ? acc : "#12151f", border: `1px solid ${markup === m ? acc : "#2a2d3e"}`, color: markup === m ? "#fff" : "#9ca3af", borderRadius: 6, padding: "4px 11px", fontSize: 12, cursor: "pointer", fontWeight: markup === m ? 700 : 400 }}>×{m}</button>
                  ))}
                  <input type="number" step="0.01" min="1" value={markup} onChange={(e) => setMarkup(parseFloat(e.target.value) || 1)} style={{ width: 60, background: "#12151f", border: "1px solid #2a2d3e", color: "#e8eaf0", borderRadius: 6, padding: "4px 8px", fontSize: 12, textAlign: "center" }} />
                </div>
                <span style={{ fontSize: 12, color: "#9ca3af", marginLeft: "auto" }}>利益率: <span style={{ color: "#4ade80", fontWeight: 700 }}>{(((markup - 1) / markup) * 100).toFixed(1)}%</span></span>
              </div>

              {/* Table */}
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
                    {computedRows.map((r, idx) => (
                      <tr key={idx} style={{ borderBottom: "1px solid #1e2235" }}>
                        <td style={{ padding: "9px 12px" }}><input value={r.no} onChange={(e) => updateRow(idx, "no", e.target.value)} style={{ width: 30, background: "transparent", border: "none", color: "#6b7280", fontSize: 13 }} /></td>
                        <td style={{ padding: "9px 12px" }}><input value={r.item} onChange={(e) => updateRow(idx, "item", e.target.value)} style={{ width: "100%", background: "transparent", border: "none", color: "#e8eaf0", fontSize: 13 }} /></td>
                        <td style={{ padding: "9px 12px", textAlign: "right" }}><input type="number" step="0.01" value={r.manMonth} onChange={(e) => updateRow(idx, "manMonth", parseFloat(e.target.value) || 0)} style={{ width: 65, background: "transparent", border: "none", color: "#e8eaf0", fontSize: 13, textAlign: "right" }} /></td>
                        <td style={{ padding: "9px 12px", textAlign: "right" }}><input type="number" step="1000" value={r.unitCost} onChange={(e) => updateRow(idx, "unitCost", parseInt(e.target.value) || 0)} style={{ width: 85, background: "transparent", border: "none", color: "#e8eaf0", fontSize: 13, textAlign: "right" }} /></td>
                        <td style={{ padding: "9px 12px", textAlign: "right", color: "#6b7280" }}>{fmt(r.origAmt)}</td>
                        <td style={{ padding: "9px 12px", textAlign: "right", color: "#4ade80", fontWeight: 600 }}>{fmt(r.newAmt)}</td>
                      </tr>
                    ))}
                    <tr style={{ background: "#12151f", fontWeight: 700 }}>
                      <td colSpan={4} style={{ padding: "11px 12px", color: "#9ca3af", fontSize: 12 }}>合計</td>
                      <td style={{ padding: "11px 12px", textAlign: "right", color: "#6b7280" }}>{fmt(origTotal)}</td>
                      <td style={{ padding: "11px 12px", textAlign: "right", color: "#4ade80", fontSize: 14 }}>{fmt(newTotal)}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
              <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                <button onClick={() => setRows((p) => [...p, { no: String(p.length + 1), item: "", manMonth: 1, unitCost: 550000 }])} style={{ background: "#1a1d2e", border: "1px dashed #3a3d50", color: "#6b7280", borderRadius: 8, padding: "6px 14px", fontSize: 12, cursor: "pointer" }}>＋ 行を追加</button>
                {rows.length > 1 && <button onClick={() => setRows((p) => p.slice(0, -1))} style={{ background: "#1a1d2e", border: "1px dashed #3a3d50", color: "#6b7280", borderRadius: 8, padding: "6px 14px", fontSize: 12, cursor: "pointer" }}>－ 最終行を削除</button>}
              </div>
            </div>

            {/* Right: summary + export */}
            <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
              <div style={{ background: "#1a1d2e", borderRadius: 12, padding: 18, border: "1px solid #2a2d3e" }}>
                <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 12, letterSpacing: "0.05em" }}>サマリー</div>
                {[["仕入れ合計", fmt(origTotal), "#9ca3af", false], ["販売合計（税抜）", fmt(newTotal), "#4ade80", true], ["消費税（10%）", fmt(newTotal * 0.1), "#9ca3af", false], ["販売合計（税込）", fmt(newTotal * 1.1), acc, true], ["粗利", fmt(newTotal - origTotal), "#60a5fa", false], ["利益率", `${(((markup - 1) / markup) * 100).toFixed(1)}%`, "#a78bfa", false]].map(([label, value, color, big]) => (
                  <div key={label} style={{ display: "flex", justifyContent: "space-between", marginBottom: 8, borderBottom: big ? `1px solid #2a2d3e` : "none", paddingBottom: big ? 8 : 0 }}>
                    <span style={{ fontSize: 12, color: "#6b7280" }}>{label}</span>
                    <span style={{ fontSize: big ? 16 : 13, fontWeight: big ? 700 : 500, color }}>{value}</span>
                  </div>
                ))}
              </div>
              <button onClick={() => setTab("preview")} style={{ background: `linear-gradient(135deg, ${acc}, ${acc}aa)`, border: "none", color: "#fff", borderRadius: 10, padding: "12px 0", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>
                👁 プレビューへ →
              </button>
              <button onClick={exportExcel} style={{ background: "#1a2e1a", border: "1px solid #2a4a2a", color: "#4ade80", borderRadius: 10, padding: "11px 0", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
                📥 Excelでダウンロード
              </button>
            </div>
          </div>
        )}

        {tab === "preview" && (
          <div>
            <div style={{ display: "flex", gap: 10, marginBottom: 20, justifyContent: "flex-end" }}>
              <button onClick={() => setTab("edit")} style={{ background: "#1a1d2e", border: "1px solid #3a3d50", color: "#9ca3af", borderRadius: 8, padding: "9px 20px", fontSize: 13, cursor: "pointer" }}>
                ← 編集に戻る
              </button>
              <button onClick={handlePrint} style={{ background: `linear-gradient(135deg, ${acc}, ${acc}cc)`, border: "none", color: "#fff", borderRadius: 8, padding: "9px 24px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                🖨️ PDF / 印刷
              </button>
            </div>
            <QuotePreview rows={rows} markup={markup} settings={settings} quoteInfo={quoteInfo} />
          </div>
        )}
      </div>
    </div>
  );
}
