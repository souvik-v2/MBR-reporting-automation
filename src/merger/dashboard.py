"""
Generates a single, self-contained HTML dashboard from the same merged data
that goes into the master workbook. No server required — data is embedded
as JSON directly in the file, so it can be opened by double-clicking.
"""
from __future__ import annotations

import json
from datetime import datetime
from typing import List

from .writer import order_columns

DASHBOARD_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Master Data Dashboard</title>
<style>
:root {
  --ink: #1b1f24;
  --ink-soft: #5b6470;
  --paper: #fbfaf8;
  --panel: #ffffff;
  --line: #e6e2da;
  --accent: #2c5f4f;
  --accent-soft: #e3ece8;
  --warn: #9a5b13;
  --warn-soft: #fbeede;
  --mono: "IBM Plex Mono", "SFMono-Regular", Consolas, monospace;
  --sans: "Inter", -apple-system, "Segoe UI", sans-serif;
}
* { box-sizing: border-box; }
body {
  margin: 0;
  font-family: var(--sans);
  background: var(--paper);
  color: var(--ink);
  -webkit-font-smoothing: antialiased;
}
header {
  padding: 28px 32px 20px;
  border-bottom: 1px solid var(--line);
  background: var(--panel);
  position: sticky;
  top: 0;
  z-index: 20;
}
.title-row {
  display: flex;
  align-items: baseline;
  justify-content: space-between;
  flex-wrap: wrap;
  gap: 8px;
}
h1 {
  font-size: 22px;
  font-weight: 650;
  letter-spacing: -0.01em;
  margin: 0;
}
.meta {
  font-family: var(--mono);
  font-size: 12px;
  color: var(--ink-soft);
}
.summary {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
  gap: 12px;
  padding: 20px 32px;
  max-width: 1400px;
}
.card {
  background: var(--panel);
  border: 1px solid var(--line);
  border-radius: 10px;
  padding: 14px 16px;
}
.card .num {
  font-family: var(--mono);
  font-size: 26px;
  font-weight: 600;
  line-height: 1.1;
}
.card .label {
  font-size: 12px;
  color: var(--ink-soft);
  margin-top: 4px;
}
.card.accent { border-color: var(--accent); background: var(--accent-soft); }
.card.warn { border-color: var(--warn); background: var(--warn-soft); }

.controls {
  display: flex;
  gap: 10px;
  align-items: center;
  padding: 0 32px 16px;
  flex-wrap: wrap;
  max-width: 1400px;
}
.controls input[type="search"] {
  flex: 1;
  min-width: 220px;
  padding: 9px 14px;
  border: 1px solid var(--line);
  border-radius: 8px;
  font-size: 14px;
  background: var(--panel);
}
.controls select {
  padding: 9px 12px;
  border: 1px solid var(--line);
  border-radius: 8px;
  font-size: 13px;
  background: var(--panel);
  color: var(--ink);
}
.tabs {
  display: flex;
  gap: 4px;
  padding: 0 32px;
  border-bottom: 1px solid var(--line);
  max-width: 1400px;
}
.tab {
  padding: 10px 16px;
  font-size: 13px;
  font-weight: 560;
  color: var(--ink-soft);
  cursor: pointer;
  border-bottom: 2px solid transparent;
  user-select: none;
}
.tab.active {
  color: var(--accent);
  border-bottom-color: var(--accent);
}
main {
  padding: 18px 32px 60px;
  max-width: 1400px;
}
.table-wrap {
  border: 1px solid var(--line);
  border-radius: 10px;
  overflow: auto;
  max-height: 65vh;
  background: var(--panel);
}
table {
  border-collapse: collapse;
  width: 100%;
  font-size: 13px;
}
thead th {
  position: sticky;
  top: 0;
  background: #f3f1eb;
  text-align: left;
  padding: 10px 14px;
  font-weight: 620;
  border-bottom: 1px solid var(--line);
  white-space: nowrap;
  cursor: pointer;
}
thead th:hover { color: var(--accent); }
thead th .arrow { font-size: 10px; opacity: 0.5; margin-left: 4px; }
tbody td {
  padding: 9px 14px;
  border-bottom: 1px solid var(--line);
  white-space: nowrap;
  max-width: 320px;
  overflow: hidden;
  text-overflow: ellipsis;
}
tbody tr:hover { background: #f7f5f0; }
tbody td.num { font-family: var(--mono); text-align: right; }
.source-tag {
  font-family: var(--mono);
  font-size: 11px;
  color: var(--ink-soft);
}
.empty-state {
  padding: 40px;
  text-align: center;
  color: var(--ink-soft);
  font-size: 14px;
}
.log-list {
  font-family: var(--mono);
  font-size: 12.5px;
  line-height: 1.7;
  background: var(--panel);
  border: 1px solid var(--line);
  border-radius: 10px;
  padding: 16px 20px;
  max-height: 65vh;
  overflow: auto;
}
.log-list div { padding: 2px 0; border-bottom: 1px dashed var(--line); }
.log-list div:last-child { border-bottom: none; }
.row-count {
  font-size: 12px;
  color: var(--ink-soft);
  padding: 10px 2px 0;
}
footer {
  padding: 18px 32px 36px;
  font-size: 12px;
  color: var(--ink-soft);
}
</style>
</head>
<body>

<header>
  <div class="title-row">
    <h1>Master Data Dashboard</h1>
    <div class="meta" id="generatedAt"></div>
  </div>
</header>

<div class="summary" id="summaryCards"></div>

<div class="tabs" id="tabs"></div>

<div class="controls" id="controlsRow">
  <input type="search" id="searchBox" placeholder="Search all columns…">
  <select id="sourceFilter"></select>
</div>

<main>
  <div class="table-wrap" id="tableWrap"></div>
  <div class="row-count" id="rowCount"></div>
</main>

<footer>Generated by the Excel / CSV / PPTX merger. Open this file directly in any browser — no server needed.</footer>

<script>
const DATA = __DATA_JSON__;

let activeTab = "master";
let sortCol = null;
let sortDir = 1;

const tabsEl = document.getElementById("tabs");
const tableWrap = document.getElementById("tableWrap");
const searchBox = document.getElementById("searchBox");
const sourceFilter = document.getElementById("sourceFilter");
const rowCountEl = document.getElementById("rowCount");

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
}

function renderSummary() {
  const cards = [
    { num: DATA.master.rows.length, label: "Master records (latest only)", cls: "accent" },
    { num: DATA.all.rows.length, label: "Total rows found" },
    { num: DATA.pptx_text.rows.length, label: "PPTX text blocks" },
    { num: DATA.pptx_charts.rows.length, label: "PPTX chart data points" },
    { num: DATA.log.filter(l => l.startsWith("[WARN]") || l.startsWith("[ERROR]")).length, label: "Warnings / errors", cls: "warn" },
  ];
  document.getElementById("summaryCards").innerHTML = cards.map(c =>
    `<div class="card ${c.cls||''}"><div class="num">${c.num}</div><div class="label">${c.label}</div></div>`
  ).join("");
}

function renderTabs() {
  const tabs = [
    { id: "master", label: `Master (${DATA.master.rows.length})` },
    { id: "all", label: `All Records (${DATA.all.rows.length})` },
    { id: "pptx_text", label: `PPTX Text (${DATA.pptx_text.rows.length})` },
    { id: "pptx_charts", label: `PPTX Charts (${DATA.pptx_charts.rows.length})` },
    { id: "log", label: `Run Log (${DATA.log.length})` },
  ];
  tabsEl.innerHTML = tabs.map(t =>
    `<div class="tab ${t.id===activeTab?'active':''}" data-tab="${t.id}">${t.label}</div>`
  ).join("");
  tabsEl.querySelectorAll(".tab").forEach(el => {
    el.addEventListener("click", () => {
      activeTab = el.dataset.tab;
      sortCol = null;
      renderTabs();
      renderControls();
      renderTable();
    });
  });
}

function renderControls() {
  const controlsRow = document.getElementById("controlsRow");
  if (activeTab === "log") {
    controlsRow.style.display = "none";
    return;
  }
  controlsRow.style.display = "flex";
  const dataset = DATA[activeTab];
  const sourceColIdx = dataset.columns.indexOf("Source File");
  if (sourceColIdx === -1) {
    sourceFilter.style.display = "none";
  } else {
    sourceFilter.style.display = "inline-block";
    const sources = [...new Set(dataset.rows.map(r => r[sourceColIdx]))].filter(Boolean).sort();
    sourceFilter.innerHTML = `<option value="">All source files</option>` +
      sources.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join("");
  }
}

function getFilteredRows() {
  if (activeTab === "log") return null;
  const dataset = DATA[activeTab];
  const q = searchBox.value.trim().toLowerCase();
  const sourceColIdx = dataset.columns.indexOf("Source File");
  const sourceVal = sourceFilter.value;

  let rows = dataset.rows;
  if (sourceVal && sourceColIdx !== -1) {
    rows = rows.filter(r => r[sourceColIdx] === sourceVal);
  }
  if (q) {
    rows = rows.filter(r => r.some(cell => String(cell ?? "").toLowerCase().includes(q)));
  }
  if (sortCol !== null) {
    rows = [...rows].sort((a, b) => {
      const av = a[sortCol] ?? "";
      const bv = b[sortCol] ?? "";
      const an = parseFloat(av), bn = parseFloat(bv);
      let cmp;
      if (!isNaN(an) && !isNaN(bn) && String(an) === String(av).trim() && String(bn) === String(bv).trim()) {
        cmp = an - bn;
      } else {
        cmp = String(av).localeCompare(String(bv));
      }
      return cmp * sortDir;
    });
  }
  return rows;
}

function renderTable() {
  if (activeTab === "log") {
    tableWrap.innerHTML = `<div class="log-list">${
      DATA.log.map(l => `<div>${escapeHtml(l)}</div>`).join("")
    }</div>`;
    rowCountEl.textContent = `${DATA.log.length} log entries`;
    return;
  }

  const dataset = DATA[activeTab];
  const rows = getFilteredRows();

  if (rows.length === 0) {
    tableWrap.innerHTML = `<div class="empty-state">No rows match your search/filter.</div>`;
    rowCountEl.textContent = "0 rows";
    return;
  }

  const numericCols = dataset.columns.map((_, i) =>
    rows.every(r => r[i] === "" || r[i] === null || !isNaN(parseFloat(r[i])))
  );

  let html = "<table><thead><tr>";
  dataset.columns.forEach((col, i) => {
    const arrow = sortCol === i ? (sortDir === 1 ? "▲" : "▼") : "";
    html += `<th data-col="${i}">${escapeHtml(col)} <span class="arrow">${arrow}</span></th>`;
  });
  html += "</tr></thead><tbody>";
  for (const row of rows.slice(0, 2000)) {
    html += "<tr>";
    row.forEach((cell, i) => {
      const cls = numericCols[i] && cell !== "" ? "num" : "";
      html += `<td class="${cls}" title="${escapeHtml(cell ?? '')}">${escapeHtml(cell ?? '')}</td>`;
    });
    html += "</tr>";
  }
  html += "</tbody></table>";
  tableWrap.innerHTML = html;

  tableWrap.querySelectorAll("th").forEach(th => {
    th.addEventListener("click", () => {
      const col = parseInt(th.dataset.col, 10);
      if (sortCol === col) { sortDir *= -1; } else { sortCol = col; sortDir = 1; }
      renderTable();
    });
  });

  const shown = Math.min(rows.length, 2000);
  rowCountEl.textContent = rows.length > 2000
    ? `Showing first ${shown.toLocaleString()} of ${rows.length.toLocaleString()} rows — narrow your search to see more precisely.`
    : `${rows.length.toLocaleString()} row${rows.length === 1 ? "" : "s"}`;
}

searchBox.addEventListener("input", renderTable);
sourceFilter.addEventListener("change", renderTable);

document.getElementById("generatedAt").textContent = "Generated " + DATA.generated_at;
renderSummary();
renderTabs();
renderControls();
renderTable();
</script>
</body>
</html>
"""


def _records_to_table(records: List[dict], priority: List[str]) -> dict:
    if not records:
        return {"columns": [], "rows": []}
    columns = order_columns(records, priority)
    display_headers = [
        c.replace("__", "").replace("_", " ").strip().title() if c.startswith("__") else c
        for c in columns
    ]
    rows = []
    for r in records:
        row = []
        for col in columns:
            val = r.get(col)
            if val is None:
                val = ""
            elif hasattr(val, "isoformat"):
                val = val.isoformat(sep=" ")
            row.append(val)
        rows.append(row)
    return {"columns": display_headers, "rows": rows}


def write_dashboard(
    output_path: str,
    master_records: List[dict],
    all_records: List[dict],
    pptx_text_rows: List[dict],
    pptx_chart_rows: List[dict],
    log: List[str],
    id_col: str,
    date_col: str,
):
    data = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "master": _records_to_table(master_records, [id_col, date_col]),
        "all": _records_to_table(all_records, [id_col, date_col]),
        "pptx_text": _records_to_table(pptx_text_rows, ["Slide", "Placeholder", "Text"]),
        "pptx_charts": _records_to_table(
            pptx_chart_rows, ["Slide", "Chart Title", "Series", "Category", "Value"]
        ),
        "log": log,
    }
    html = DASHBOARD_TEMPLATE.replace("__DATA_JSON__", json.dumps(data, default=str))
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
