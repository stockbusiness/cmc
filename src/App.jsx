import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

const DEFAULT_ROWS = [
  { no: "1", item: "要件の明確化", manMonth: 0.95, unitCost: 550000 },
  { no: "2", item: "開発", manMonth: 3.95, unitCost: 550000 },
  { no: "3", item: "テスト", manMonth: 1.55, unitCost: 500000 },
  { no: "4", item: "UAT Go-Live サポート", manMonth: 1.25, unitCost: 550000 },
  { no: "5", item: "BrSE", manMonth: 1.155, unitCost: 800000 },
  { no: "6", item: "管理", manMonth: 1.155, unitCost: 550000 },
];

const COMPANY_INFO = {
  from: "Zept合同会社",
  to: "株式会社エンケイ",
  projectName: "部品の設計図作成システム 1次フェーズ",
  date: new Date().toLocaleDateString("ja-JP", { year: "numeric", month: "2-digit", day: "2-digit" }),
};

const fmt = (n) =>
  "¥ " + Math.round(n).toLocaleString("ja-JP");

function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        // Find cost sheet
        const sheetName = wb.SheetNames.find(
          (n) => n.includes("見積") || n.includes("コスト") || n.includes("cost")
        ) || wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

        // Find header row with 工数 and 金額
        let headerIdx = -1;
        let noCol, itemCol, manMonthCol, costCol, amtCol;
        for (let i = 0; i < rows.length; i++) {
          const r = rows[i].map((c) => (c ? String(c).trim() : ""));
          const noI = r.findIndex((c) => c === "No" || c === "NO" || c === "#");
          const mmI = r.findIndex((c) => c.includes("工数") || c.includes("人月"));
          const costI = r.findIndex((c) => (c.includes("コスト") || c.includes("単価")) && !c.includes("合計"));
          const amtI = r.findIndex((c) => c.includes("金額") || c.includes("amount"));
          const itemI = r.findIndex((c) => c.includes("項目") || c.includes("タスク") || c.includes("item"));
          if (mmI !== -1 && amtI !== -1) {
            headerIdx = i;
            noCol = noI >= 0 ? noI : 0;
            itemCol = itemI >= 0 ? itemI : noI + 1;
            manMonthCol = mmI;
            costCol = costI >= 0 ? costI : mmI + 1;
            amtCol = amtI;
            break;
          }
        }

        if (headerIdx === -1) {
          reject(new Error("見積コストのシートが見つかりませんでした"));
          return;
        }

        const parsed = [];
        for (let i = headerIdx + 1; i < rows.length; i++) {
          const r = rows[i];
          const no = r[noCol];
          const item = r[itemCol];
          const mm = parseFloat(r[manMonthCol]);
          const uc = parseFloat(r[costCol]);
          if (!item || isNaN(mm) || mm <= 0) continue;
          if (String(item).includes("合計") || String(item).includes("注記")) continue;
          parsed.push({
            no: String(no ?? parsed.length + 1),
            item: String(item).trim(),
            manMonth: mm,
            unitCost: isNaN(uc) ? 0 : uc,
          });
        }

        if (parsed.length === 0) reject(new Error("データ行が見つかりませんでした"));
        else resolve(parsed);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("ファイル読み込みエラー"));
    reader.readAsArrayBuffer(file);
  });
}

export default function App() {
  const [rows, setRows] = useState(DEFAULT_ROWS);
  const [markup, setMarkup] = useState(1.8);
  const [markupInput, setMarkupInput] = useState("1.8");
  const [company, setCompany] = useState(COMPANY_INFO);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [step, setStep] = useState("edit"); // edit | preview
  const fileRef = useRef();

  const handleFile = useCallback(async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setLoading(true);
    setError("");
    try {
      const parsed = await parseExcelFile(file);
      setRows(parsed);
      setStep("edit");
    } catch (err) {
      setError("❌ " + err.message + "  ※ 「2.見積コスト」などのシートを含むExcelが対象です");
    } finally {
      setLoading(false);
      e.target.value = "";
    }
  }, []);

  const updateRow = (idx, field, val) => {
    setRows((prev) =>
      prev.map((r, i) => (i === idx ? { ...r, [field]: val } : r))
    );
  };

  const applyMarkup = (m) => {
    setMarkup(m);
    setMarkupInput(String(m));
  };

  const handleMarkupInput = (v) => {
    setMarkupInput(v);
    const n = parseFloat(v);
    if (!isNaN(n) && n > 0) setMarkup(n);
  };

  const computedRows = rows.map((r) => ({
    ...r,
    origAmt: r.manMonth * r.unitCost,
    newAmt: r.manMonth * r.unitCost * markup,
    newManMonth: parseFloat((r.manMonth * markup).toFixed(3)),
    newUnitCost: Math.round(r.unitCost * markup),
  }));
  const origTotal = computedRows.reduce((s, r) => s + r.origAmt, 0);
  const newTotal = computedRows.reduce((s, r) => s + r.newAmt, 0);

  const exportExcel = () => {
    const header = ["No", "項目", "工数 (人月)", "コスト (JPY)", "金額 (JPY)"];
    const data = [
      header,
      ...computedRows.map((r) => [
        r.no,
        r.item,
        r.newManMonth,
        r.newUnitCost,
        Math.round(r.newAmt),
      ]),
      ["合計", "", parseFloat(computedRows.reduce((s, r) => s + r.newManMonth, 0).toFixed(2)), "", Math.round(newTotal)],
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws["!cols"] = [{ wch: 6 }, { wch: 24 }, { wch: 12 }, { wch: 14 }, { wch: 14 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "見積コスト");
    XLSX.writeFile(wb, `御見積書_${company.to}_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  return (
    <div style={{ minHeight: "100vh", background: "#0f1117", color: "#e8eaf0", fontFamily: "'Hiragino Sans', 'Meiryo', sans-serif" }}>
      {/* Header */}
      <header style={{ background: "linear-gradient(135deg, #1a1d2e 0%, #12151f 100%)", borderBottom: "1px solid #2a2d3e", padding: "20px 32px", display: "flex", alignItems: "center", gap: 16 }}>
        <div style={{ width: 36, height: 36, borderRadius: 8, background: "linear-gradient(135deg, #4f8ef7, #7b5ea7)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18 }}>📋</div>
        <div>
          <div style={{ fontWeight: 700, fontSize: 16, letterSpacing: "0.02em" }}>見積書 コンバーター</div>
          <div style={{ fontSize: 11, color: "#6b7280", marginTop: 2 }}>仕入れ見積 → 自社見積書 自動変換</div>
        </div>
        <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
          <button
            onClick={() => fileRef.current.click()}
            style={{ background: "#1e2235", border: "1px solid #3a3d50", color: "#9ca3af", borderRadius: 8, padding: "8px 16px", fontSize: 13, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 }}
          >
            📂 Excelを読み込む
          </button>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
        </div>
      </header>

      <div style={{ maxWidth: 1100, margin: "0 auto", padding: "24px 24px" }}>
        {error && (
          <div style={{ background: "#2d1515", border: "1px solid #7f1d1d", borderRadius: 10, padding: "12px 16px", marginBottom: 16, fontSize: 13, color: "#fca5a5" }}>
            {error}
          </div>
        )}
        {loading && (
          <div style={{ textAlign: "center", padding: 40, color: "#6b7280" }}>⏳ Excelを解析中...</div>
        )}

        <div style={{ display: "grid", gridTemplateColumns: "1fr 320px", gap: 20 }}>
          {/* Main table */}
          <div>
            {/* Markup control */}
            <div style={{ background: "#1a1d2e", borderRadius: 12, padding: "16px 20px", marginBottom: 16, display: "flex", alignItems: "center", gap: 20, border: "1px solid #2a2d3e" }}>
              <div style={{ fontSize: 13, color: "#9ca3af", whiteSpace: "nowrap" }}>掛け率 (マージン)</div>
              <div style={{ display: "flex", gap: 6 }}>
                {[1.2, 1.3, 1.5, 1.8, 2.0].map((m) => (
                  <button key={m} onClick={() => applyMarkup(m)} style={{ background: markup === m ? "linear-gradient(135deg, #4f8ef7, #7b5ea7)" : "#12151f", border: "1px solid " + (markup === m ? "#4f8ef7" : "#2a2d3e"), color: markup === m ? "#fff" : "#9ca3af", borderRadius: 6, padding: "5px 12px", fontSize: 13, cursor: "pointer", fontWeight: markup === m ? 700 : 400 }}>
                    ×{m}
                  </button>
                ))}
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ fontSize: 13, color: "#6b7280" }}>カスタム</span>
                <input
                  type="number"
                  step="0.01"
                  min="1"
                  value={markupInput}
                  onChange={(e) => handleMarkupInput(e.target.value)}
                  style={{ width: 70, background: "#12151f", border: "1px solid #3a3d50", color: "#e8eaf0", borderRadius: 6, padding: "5px 10px", fontSize: 13, textAlign: "center" }}
                />
              </div>
              <div style={{ marginLeft: "auto", fontSize: 13, color: "#6b7280" }}>
                利益率: <span style={{ color: "#4ade80", fontWeight: 700 }}>{(((markup - 1) / markup) * 100).toFixed(1)}%</span>
              </div>
            </div>

            {/* Table */}
            <div style={{ background: "#1a1d2e", borderRadius: 12, border: "1px solid #2a2d3e", overflow: "hidden" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                <thead>
                  <tr style={{ background: "#12151f" }}>
                    {["No", "項目", "工数(人月)", "単価(JPY)", "仕入金額", "販売金額"].map((h, i) => (
                      <th key={h} style={{ padding: "11px 14px", textAlign: i < 2 ? "left" : "right", color: "#6b7280", fontWeight: 600, fontSize: 11, letterSpacing: "0.05em", borderBottom: "1px solid #2a2d3e", whiteSpace: "nowrap" }}>
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {computedRows.map((r, idx) => (
                    <tr key={idx} style={{ borderBottom: "1px solid #1e2235" }}>
                      <td style={{ padding: "10px 14px", color: "#6b7280", width: 40 }}>
                        <input value={r.no} onChange={(e) => updateRow(idx, "no", e.target.value)} style={{ width: 32, background: "transparent", border: "none", color: "#6b7280", fontSize: 13, textAlign: "center" }} />
                      </td>
                      <td style={{ padding: "10px 14px" }}>
                        <input value={r.item} onChange={(e) => updateRow(idx, "item", e.target.value)} style={{ width: "100%", background: "transparent", border: "none", color: "#e8eaf0", fontSize: 13 }} />
                      </td>
                      <td style={{ padding: "10px 14px", textAlign: "right" }}>
                        <input type="number" step="0.01" value={r.manMonth} onChange={(e) => updateRow(idx, "manMonth", parseFloat(e.target.value) || 0)} style={{ width: 70, background: "transparent", border: "none", color: "#e8eaf0", fontSize: 13, textAlign: "right" }} />
                      </td>
                      <td style={{ padding: "10px 14px", textAlign: "right" }}>
                        <input type="number" step="1000" value={r.unitCost} onChange={(e) => updateRow(idx, "unitCost", parseInt(e.target.value) || 0)} style={{ width: 90, background: "transparent", border: "none", color: "#e8eaf0", fontSize: 13, textAlign: "right" }} />
                      </td>
                      <td style={{ padding: "10px 14px", textAlign: "right", color: "#6b7280" }}>
                        {fmt(r.origAmt)}
                      </td>
                      <td style={{ padding: "10px 14px", textAlign: "right", color: "#4ade80", fontWeight: 600 }}>
                        {fmt(r.newAmt)}
                      </td>
                    </tr>
                  ))}
                  <tr style={{ background: "#12151f", fontWeight: 700 }}>
                    <td colSpan={4} style={{ padding: "12px 14px", color: "#9ca3af", fontSize: 12 }}>合計</td>
                    <td style={{ padding: "12px 14px", textAlign: "right", color: "#6b7280" }}>{fmt(origTotal)}</td>
                    <td style={{ padding: "12px 14px", textAlign: "right", color: "#4ade80", fontSize: 15 }}>{fmt(newTotal)}</td>
                  </tr>
                </tbody>
              </table>
            </div>

            {/* Add / Remove row */}
            <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
              <button
                onClick={() => setRows((p) => [...p, { no: String(p.length + 1), item: "", manMonth: 1, unitCost: 550000 }])}
                style={{ background: "#1a1d2e", border: "1px dashed #3a3d50", color: "#6b7280", borderRadius: 8, padding: "7px 16px", fontSize: 12, cursor: "pointer" }}
              >
                ＋ 行を追加
              </button>
              {rows.length > 1 && (
                <button
                  onClick={() => setRows((p) => p.slice(0, -1))}
                  style={{ background: "#1a1d2e", border: "1px dashed #3a3d50", color: "#6b7280", borderRadius: 8, padding: "7px 16px", fontSize: 12, cursor: "pointer" }}
                >
                  － 最終行を削除
                </button>
              )}
            </div>
          </div>

          {/* Right panel */}
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            {/* Summary */}
            <div style={{ background: "#1a1d2e", borderRadius: 12, padding: 20, border: "1px solid #2a2d3e" }}>
              <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 14, letterSpacing: "0.05em" }}>サマリー</div>
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                <SummaryLine label="仕入れ合計" value={fmt(origTotal)} color="#9ca3af" />
                <SummaryLine label="販売合計" value={fmt(newTotal)} color="#4ade80" big />
                <div style={{ borderTop: "1px solid #2a2d3e", paddingTop: 10 }}>
                  <SummaryLine label="粗利" value={fmt(newTotal - origTotal)} color="#60a5fa" />
                  <SummaryLine label="掛け率" value={`×${markup}`} color="#f59e0b" />
                  <SummaryLine label="利益率" value={`${(((markup - 1) / markup) * 100).toFixed(1)}%`} color="#a78bfa" />
                </div>
              </div>
            </div>

            {/* Client info */}
            <div style={{ background: "#1a1d2e", borderRadius: 12, padding: 20, border: "1px solid #2a2d3e" }}>
              <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 14, letterSpacing: "0.05em" }}>見積書情報</div>
              {[
                { label: "宛先", field: "to" },
                { label: "自社名", field: "from" },
                { label: "件名", field: "projectName" },
                { label: "日付", field: "date" },
              ].map(({ label, field }) => (
                <div key={field} style={{ marginBottom: 10 }}>
                  <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 3 }}>{label}</div>
                  <input
                    value={company[field]}
                    onChange={(e) => setCompany((p) => ({ ...p, [field]: e.target.value }))}
                    style={{ width: "100%", background: "#12151f", border: "1px solid #2a2d3e", color: "#e8eaf0", borderRadius: 6, padding: "7px 10px", fontSize: 13, boxSizing: "border-box" }}
                  />
                </div>
              ))}
            </div>

            {/* Export */}
            <div style={{ background: "linear-gradient(135deg, #1e2a1e, #1a2e2a)", borderRadius: 12, padding: 20, border: "1px solid #2a4a3a" }}>
              <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 14, letterSpacing: "0.05em" }}>出力</div>
              <button
                onClick={exportExcel}
                style={{ width: "100%", background: "linear-gradient(135deg, #22c55e, #16a34a)", border: "none", color: "#fff", borderRadius: 8, padding: "11px 0", fontSize: 14, fontWeight: 700, cursor: "pointer", letterSpacing: "0.02em" }}
              >
                📥 Excelでダウンロード
              </button>
              <div style={{ fontSize: 11, color: "#4b5563", marginTop: 8, textAlign: "center" }}>
                ※ PDF出力は印刷→PDF保存で対応
              </div>
            </div>

            {/* Preview button */}
            <button
              onClick={() => setStep(step === "preview" ? "edit" : "preview")}
              style={{ background: "#1a1d2e", border: "1px solid #3a3d50", color: "#9ca3af", borderRadius: 8, padding: "10px 0", fontSize: 13, cursor: "pointer", width: "100%" }}
            >
              {step === "preview" ? "← 編集に戻る" : "👁 プレビュー"}
            </button>
          </div>
        </div>

        {/* Preview */}
        {step === "preview" && (
          <div style={{ marginTop: 24, background: "#fff", borderRadius: 12, padding: 40, color: "#111", maxWidth: 800, margin: "24px auto 0" }}>
            <div style={{ textAlign: "right", fontSize: 12, color: "#666", marginBottom: 4 }}>作成日：{company.date}</div>
            <div style={{ textAlign: "center", fontSize: 22, fontWeight: 700, marginBottom: 4 }}>御　見　積　書</div>
            <div style={{ textAlign: "center", color: "#666", fontSize: 13, marginBottom: 24 }}>{company.projectName}</div>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 32 }}>
              <div>
                <div style={{ fontWeight: 700, fontSize: 15 }}>{company.to}　御中</div>
              </div>
              <div style={{ textAlign: "right", fontSize: 13, color: "#444" }}>
                <div style={{ fontWeight: 700, fontSize: 15 }}>{company.from}</div>
              </div>
            </div>
            <div style={{ marginBottom: 8, fontWeight: 700, fontSize: 28, color: "#111" }}>
              見積金額合計：{fmt(newTotal)} <span style={{ fontSize: 14, fontWeight: 400 }}>(税抜)</span>
            </div>
            <div style={{ marginBottom: 24, fontSize: 12, color: "#888" }}>※上記金額は消費税抜き価格です。</div>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr style={{ background: "#1a1d2e", color: "#fff" }}>
                  {["No", "項目", "工数（人月）", "単価（JPY）", "金額（JPY）"].map((h) => (
                    <th key={h} style={{ padding: "10px 12px", textAlign: "left", fontWeight: 600, fontSize: 12 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {computedRows.map((r, i) => (
                  <tr key={i} style={{ background: i % 2 === 0 ? "#f9fafb" : "#fff" }}>
                    <td style={{ padding: "9px 12px", borderBottom: "1px solid #e5e7eb" }}>{r.no}</td>
                    <td style={{ padding: "9px 12px", borderBottom: "1px solid #e5e7eb" }}>{r.item}</td>
                    <td style={{ padding: "9px 12px", borderBottom: "1px solid #e5e7eb", textAlign: "right" }}>{r.newManMonth}</td>
                    <td style={{ padding: "9px 12px", borderBottom: "1px solid #e5e7eb", textAlign: "right" }}>{r.newUnitCost.toLocaleString()}</td>
                    <td style={{ padding: "9px 12px", borderBottom: "1px solid #e5e7eb", textAlign: "right", fontWeight: 600 }}>{fmt(r.newAmt)}</td>
                  </tr>
                ))}
                <tr style={{ background: "#f0f4ff", fontWeight: 700 }}>
                  <td colSpan={4} style={{ padding: "11px 12px" }}>合計</td>
                  <td style={{ padding: "11px 12px", textAlign: "right", fontSize: 15 }}>{fmt(newTotal)}</td>
                </tr>
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

function SummaryLine({ label, value, color, big }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 6 }}>
      <span style={{ fontSize: 12, color: "#6b7280" }}>{label}</span>
      <span style={{ fontSize: big ? 18 : 14, fontWeight: big ? 700 : 600, color }}>{value}</span>
    </div>
  );
}
