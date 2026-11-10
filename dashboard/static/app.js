/* Hardware Lot dashboard - read-only view of Hardware Inventory.xlsx. */
"use strict";

const POLL_MS = 60000;

const state = {
  data: null,
  error: null,
  tab: "overview",
  openAsset: null,
  inv: { q: "", type: "", model: "", status: "", sold: "unsold", win11: "", drive: "", sort: "asset_id", dir: 1 },
  sales: { q: "", platform: "", type: "", sort: "date", dir: -1 },
  listings: { q: "", status: "", kind: "", sort: "status", dir: 1 },
  fleetScope: "unsold",
  hwFleetScope: "unsold",
  hw: { q: "", type: "", condition: "", listing: "", sold: "unsold", sort: "id", dir: 1 },
};

/* ---------------- helpers ---------------- */

const $ = (sel, root = document) => root.querySelector(sel);

function esc(v) {
  if (v === null || v === undefined) return "";
  return String(v)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function money(v, opts = {}) {
  if (v === null || v === undefined || v === "") return opts.blank ?? "";
  const n = Number(v);
  if (!isFinite(n)) return esc(v);
  const whole = Number.isInteger(n) && !opts.cents;
  const s = Math.abs(n).toLocaleString("en-US", {
    minimumFractionDigits: whole ? 0 : 2, maximumFractionDigits: whole ? 0 : 2,
  });
  return (n < 0 ? "−$" : "$") + s;
}

function pct(v) {
  if (v === null || v === undefined) return "";
  return (v * 100).toFixed(1).replace(/\.0$/, "") + "%";
}

function fmtDateTime(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d)) return esc(iso);
  return d.toLocaleString(undefined, { year: "numeric", month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });
}

function fmtDate(s) {
  if (!s) return "";
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (!m) return esc(s);
  const d = new Date(+m[1], +m[2] - 1, +m[3]);
  return d.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
}

function naturalKey(s) {
  return String(s ?? "").split(/(\d+)/).map(p => (/^\d+$/.test(p) ? p.padStart(8, "0") : p.toLowerCase())).join("");
}

function cmp(a, b) {
  const an = a === null || a === undefined || a === "";
  const bn = b === null || b === undefined || b === "";
  if (an && bn) return 0;
  if (an) return 1;      // blanks last
  if (bn) return -1;
  if (typeof a === "number" && typeof b === "number") return a - b;
  if (typeof a === "boolean" || typeof b === "boolean") return (a ? 1 : 0) - (b ? 1 : 0);
  return naturalKey(a) < naturalKey(b) ? -1 : naturalKey(a) > naturalKey(b) ? 1 : 0;
}

function uniq(arr) {
  return [...new Set(arr.filter(v => v !== null && v !== undefined && v !== ""))].sort((a, b) => cmp(a, b));
}

function options(values, selected, allLabel) {
  return `<option value="">${esc(allLabel)}</option>` +
    values.map(v => `<option value="${esc(v)}"${String(v) === String(selected) ? " selected" : ""}>${esc(v)}</option>`).join("");
}

function assetLink(id) {
  if (!id) return "";
  const d = state.data;
  const known = d && (d.computers.some(c => c.asset_id === id) ||
    (d.hardware || []).some(h => h.id_unique && h.id === String(id)));
  return known ? `<a href="#inventory/${encodeURIComponent(id)}" data-asset="${esc(id)}">${esc(id)}</a>` : esc(id);
}

function statusBadge(level, text) {
  return `<span class="badge lv-${level}"><span class="status-dot"></span>${esc(text)}</span>`;
}

function typeMeaning(code) {
  const d = state.data;
  if (!d || !code) return null;
  const t = (d.type_codes || []).find(x => x.code.toUpperCase() === String(code).toUpperCase());
  return t ? t.meaning : null;
}

// A type code with its meaning on hover.
function typeCell(code) {
  if (code === null || code === undefined || code === "") return "";
  const m = typeMeaning(code);
  return m ? `<span class="code" data-tip="${esc(code)}: ${esc(m)}">${esc(code)}</span>` : esc(code);
}

// Small marker for a value taken from the Master note instead of the scan.
function noteMark(c, field) {
  return c.note_overrides && c.note_overrides[field] ? ` <span class="dim small">(note)</span>` : "";
}

function ramLabel(gb) {
  return gb === null || gb === undefined ? "" : `${gb} GB`;
}

/* ---------------- charts ---------------- */

// Horizontal bars, one series (one color). rows: [{label, value}]
// typeLabels: the labels are type codes (meaning pops up on hover).
function hbars(rows, fmt = v => v, opts = {}) {
  if (!rows.length) return `<div class="empty">Nothing to show.</div>`;
  const max = Math.max(...rows.map(r => r.value), 0) || 1;
  return `<div class="bars">` + rows.map(r => {
    const w = Math.max(0, (r.value / max) * 100);
    return `<div class="bl">${opts.typeLabels ? typeCell(r.label) : esc(r.label)}</div>` +
      `<div class="bt"><div class="bf" style="width:${w}%"></div></div>` +
      `<div class="bv">${esc(fmt(r.value))}</div>`;
  }).join("") + `</div>`;
}

function countBy(list, fn) {
  const m = new Map();
  for (const x of list) {
    const k = fn(x);
    const key = k === null || k === undefined || k === "" ? "(not recorded)" : String(k);
    m.set(key, (m.get(key) || 0) + 1);
  }
  return [...m.entries()].map(([label, value]) => ({ label, value }));
}

// Cumulative net line with a dashed lot-cost reference line.
function cumulativeChart(days, lot) {
  if (!days.length) return `<div class="empty">No sales recorded yet.</div>`;
  const W = 640, H = 240, L = 56, R = 70, T = 16, B = 30;
  const pts = days.map(d => ({ x: d.date, y: d.cumulative_net, d }));
  const yMax = Math.max(lot || 0, ...pts.map(p => p.y), 1) * 1.08;
  const n = pts.length;
  const xs = i => (n === 1 ? L + (W - L - R) / 2 : L + (i * (W - L - R)) / (n - 1));
  const ys = v => T + (H - T - B) * (1 - v / yMax);
  const step = niceStep(yMax / 4);
  let grid = "";
  for (let v = 0; v <= yMax; v += step) {
    grid += `<line class="${v === 0 ? "axis" : "grid-line"}" x1="${L}" x2="${W - R}" y1="${ys(v)}" y2="${ys(v)}"/>` +
      `<text class="tick" x="${L - 8}" y="${ys(v) + 4}" text-anchor="end">${esc(money(v))}</text>`;
  }
  const every = Math.ceil(n / 8);
  let xt = "";
  pts.forEach((p, i) => {
    if (i % every === 0 || i === n - 1) {
      xt += `<text class="tick" x="${xs(i)}" y="${H - 8}" text-anchor="middle">${esc(shortDate(p.x))}</text>`;
    }
  });
  const path = pts.map((p, i) => `${i ? "L" : "M"}${xs(i)},${ys(p.y)}`).join(" ");
  const dots = pts.map((p, i) => `<circle class="dot" cx="${xs(i)}" cy="${ys(p.y)}" r="4"/>`).join("");
  const ref = lot ? `<line class="ref" x1="${L}" x2="${W - R}" y1="${ys(lot)}" y2="${ys(lot)}"/>` +
    `<text class="ref-label" x="${W - R + 6}" y="${ys(lot) + 4}">Lot cost ${esc(money(lot))}</text>` : "";
  const last = pts[n - 1];
  const endLabel = `<text class="end-label" x="${xs(n - 1) + 8}" y="${ys(last.y) - 8}">${esc(money(last.y))}</text>`;
  return `<div class="linechart"><svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Cumulative net sales by day">
    ${grid}${ref}<path class="series" d="${path}"/>${dots}${endLabel}${xt}</svg></div>`;
}

function niceStep(raw) {
  const p = Math.pow(10, Math.floor(Math.log10(raw || 1)));
  for (const m of [1, 2, 2.5, 5, 10]) if (m * p >= raw) return m * p;
  return 10 * p;
}

function shortDate(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || "");
  return m ? `${+m[2]}/${+m[3]}` : s;
}

/* ---------------- sortable table ---------------- */

function sortHeader(cols, sortState) {
  return "<tr>" + cols.map(c => {
    const on = sortState.sort === c.key;
    const arrow = on ? (sortState.dir > 0 ? "▲" : "▼") : "";
    return `<th class="sortable${c.num ? " num" : ""}" data-sort="${esc(c.key)}">${esc(c.label)}<span class="arrow">${arrow}</span></th>`;
  }).join("") + "</tr>";
}

function applySort(rows, sortState, getter) {
  const g = getter || ((r, k) => r[k]);
  const blank = v => v === null || v === undefined || v === "";
  return rows.slice().sort((a, b) => {
    const av = g(a, sortState.sort), bv = g(b, sortState.sort);
    // Blanks stay at the bottom in both directions.
    if (blank(av) !== blank(bv)) return blank(av) ? 1 : -1;
    return cmp(av, bv) * sortState.dir || cmp(a.asset_id, b.asset_id);
  });
}

/* ---------------- views ---------------- */

function viewOverview(d) {
  const k = d.kpis;
  const lotKnown = k.lot_price !== null && k.lot_price !== undefined;
  const recovered = lotKnown ? Math.min(Math.max(k.recovered || 0, 0), 1) : 0;
  const profitCls = (k.profit || 0) >= 0 ? "pos" : "";
  let breakEven = "";
  if (lotKnown) {
    if (k.break_even_remaining === 0) breakEven = "Lot cost fully recovered.";
    else if (k.break_even_units) breakEven = `≈ ${k.break_even_units} more sales at the current average net of ${money(k.avg_net)}.`;
  }
  const tiles = `
  <div class="tiles">
    <div class="tile"><div class="label">Net sales</div><div class="value">${money(k.net)}</div>
      <div class="sub">Gross ${money(k.gross)} · fees ${money(k.fees)} · shipping ${money(k.shipping)}</div></div>
    ${lotKnown ? `<div class="tile"><div class="label">Lot cost recovered</div><div class="value">${pct(k.recovered)}</div>
      <div class="meter"><span style="width:${recovered * 100}%"></span></div>
      <div class="sub">of ${money(k.lot_price)} paid for the lot</div></div>
    <div class="tile"><div class="label">${(k.profit || 0) >= 0 ? "Profit" : "Still to recover"}</div>
      <div class="value ${profitCls}">${money(Math.abs(k.profit))}</div><div class="sub">${esc(breakEven)}</div></div>` : ""}
    ${potentialTiles(d)}
    <div class="tile"><div class="label">Units sold</div><div class="value">${k.units_sold}</div>
      <div class="sub">${k.orders} order${k.orders === 1 ? "" : "s"}${k.first_sale ? ` · ${fmtDate(k.first_sale)} – ${fmtDate(k.last_sale)}` : ""}</div></div>
    <div class="tile"><div class="label">Average net per unit</div><div class="value">${money(k.avg_net, { blank: "–" })}</div>
      <div class="sub">Average sale price ${money(k.avg_price, { blank: "–" })}</div></div>
    <div class="tile"><div class="label">Machines left</div><div class="value">${k.machines_unsold}</div>
      <div class="sub">of ${k.machines_total} (${k.machines_scanned} scanned)</div></div>
  </div>`;

  const typeRows = d.by_type.map(t => `<tr>
      <td>${typeCell(t.type)}</td>
      <td class="num">${t.total ?? "–"}</td><td class="num">${t.sold}</td><td class="num">${t.unsold ?? "–"}</td>
      <td class="num">${money(t.net)}</td><td class="num">${money(t.avg_net, { blank: "–" })}</td></tr>`).join("");
  const byType = `<div class="card"><div class="card-head"><h2>By machine type</h2></div>
    <div class="table-wrap"><table><thead><tr><th>Type</th><th class="num">Machines</th><th class="num">Sold</th><th class="num">Left</th><th class="num">Net</th><th class="num">Avg net</th></tr></thead>
    <tbody>${typeRows}</tbody></table></div></div>`;

  const attn = attentionCard(d.attention);

  return `${tiles}
  <div class="grid cols-2" style="margin-top:16px">
    <div class="card"><div class="card-head"><h2>Cumulative net vs lot cost</h2><span class="muted small">by sale date</span></div>
      ${cumulativeChart(d.by_day, k.lot_price)}</div>
    ${byType}
  </div>
  <div class="grid cols-3" style="margin-top:16px">
    <div class="card"><h2>Net by platform</h2>${hbars(d.by_platform.map(p => ({ label: p.label, value: p.net, tip: `${p.label}\n${p.count} sold · gross ${money(p.gross)} · net ${money(p.net)}` })), money)}</div>
    <div class="card"><h2>Net by payment method</h2>${hbars(d.by_payment.map(p => ({ label: p.label, value: p.net, tip: `${p.label}\n${p.count} sold · net ${money(p.net)}` })), money)}</div>
    <div class="card"><h2>Net by model sold</h2>${hbars(d.by_model_sold.map(p => ({ label: p.label, value: p.net, tip: `${p.label}\n${p.count} sold · net ${money(p.net)}` })), money)}</div>
  </div>
  ${attn}`;
}

// Net sales potential = actual net of everything sold + asking prices of
// unsold units whose listing status is Listed or Pending.
function potentialTiles(d) {
  const L = d.listings || {};
  if (!L.present) return "";
  const k = d.kpis;
  const unpriced = L.active_unpriced ? ` · ${L.active_unpriced} without a price not counted` : "";
  const profit = L.profit_potential;
  return `<div class="tile"><div class="label">Net sales potential</div><div class="value">${money(L.net_sales_potential)}</div>
      <div class="sub">Sold net ${money(L.sold_net)} + asking ${money(L.asking_total)} (${L.active_priced} listed or pending)${unpriced}</div></div>
    ${profit !== null && profit !== undefined ? `<div class="tile"><div class="label">Total profit potential</div>
      <div class="value ${profit >= 0 ? "pos" : ""}">${money(profit)}</div>
      <div class="sub">Net sales potential minus ${money(k.lot_price)} lot cost</div></div>` : ""}`;
}

const STATUS_RANK = { "Sold": 0, "Pending": 1, "Listed": 2, "Needs listed": 3, "Not listed": 4, "Removed": 5 };

function listingBadge(status) {
  const level = { "Sold": "good", "Pending": "warning", "Listed": "info", "Needs listed": "serious" }[status];
  if (status === "Sold") return `<span class="badge sold">Sold</span>`;
  return level ? statusBadge(level, status) : `<span class="badge">${esc(status)}</span>`;
}

function viewListings(d) {
  const L = d.listings || {};
  if (!L.present) {
    return `<div class="card"><h2>Listings</h2><div class="empty">No Listings tab found in the workbook. Add one to see listing status, asking prices and the sales potential.</div></div>`;
  }
  const f = state.listings;
  const q = f.q.trim().toLowerCase();
  const all = d.listing_units;
  let rows = all.filter(u =>
    (!f.status || u.status === f.status) &&
    (!f.kind || u.kind === f.kind) &&
    (!q || [u.id, u.serial, u.type, u.model, u.detail, u.platform, u.listing_notes]
      .some(v => v !== null && v !== undefined && String(v).toLowerCase().includes(q))));
  const getter = (u, key) => key === "status" ? (STATUS_RANK[u.status] ?? 9) : key === "asset_id" ? u.id : u[key];
  rows = applySort(rows.map(u => ({ ...u, asset_id: u.id })), f, getter);
  const cols = [
    { key: "asset_id", label: "ID" }, { key: "kind", label: "Kind" }, { key: "type", label: "Type" },
    { key: "model", label: "Model" }, { key: "detail", label: "Spec" }, { key: "status", label: "Status" },
    { key: "asking", label: "Asking", num: true }, { key: "sale_net", label: "Sold net", num: true },
    { key: "date_listed", label: "Listed on" }, { key: "platform", label: "Platform" },
    { key: "counts", label: "In potential" }, { key: "listing_notes", label: "Notes" },
  ];
  const body = rows.map(u => {
    const idCell = u.id ? assetLink(u.id) : '<span class="dim">no ID</span>';
    const inPot = u.status === "Sold" ? `<span class="dim">at sale net</span>` : u.counts ? "Yes" : "";
    return `<tr>
      <td class="nowrap">${idCell}</td><td>${esc(u.kind)}</td><td>${typeCell(u.type)}</td><td class="nowrap">${esc(u.model)}</td>
      <td class="nowrap small">${esc(u.detail)}</td><td class="nowrap">${listingBadge(u.status)}</td>
      <td class="num">${money(u.asking)}</td><td class="num">${money(u.sale_net)}</td>
      <td class="nowrap">${fmtDate(u.date_listed)}</td><td class="nowrap">${esc(u.platform)}${u.url ? ` <a href="${esc(u.url)}" target="_blank" rel="noopener noreferrer">link</a>` : ""}</td>
      <td class="nowrap">${inPot}</td><td class="wrap small">${esc(u.listing_notes)}</td></tr>`;
  }).join("");
  const counted = rows.filter(u => u.counts);
  const foot = rows.length ? `<tfoot><tr><td colspan="6">Shown (${rows.length})</td>
      <td class="num">${money(counted.reduce((a, u) => a + u.asking, 0))}</td>
      <td class="num">${money(rows.reduce((a, u) => a + (Number(u.sale_net) || 0), 0))}</td><td colspan="4" class="muted small">Asking total counts Listed and Pending only</td></tr></tfoot>` : "";
  const statuses = L.status_counts.map(s => s.label);
  const split = x => `${x.machines} machine${x.machines === 1 ? "" : "s"}, ${x.other} hardware`;
  const by = s => (L.by_status && L.by_status[s]) || { machines: 0, other: 0 };
  const num = s => by(s).machines + by(s).other;
  const tiles = `<div class="tiles">
    ${potentialTiles(d)}
    <div class="tile"><div class="label">Listed or pending</div><div class="value">${L.active}</div>
      <div class="sub">${split(L.active_split)}</div>
      <div class="sub">${L.listed} listed · ${L.pending} pending · asking ${money(L.asking_total)}</div></div>
    <div class="tile"><div class="label">Needs listed</div><div class="value">${num("Needs listed")}</div>
      <div class="sub">${split(by("Needs listed"))}</div></div>
    <div class="tile"><div class="label">Not listed</div><div class="value">${num("Not listed")}</div>
      <div class="sub">${split(by("Not listed"))}</div></div>
    <div class="tile"><div class="label">Sold</div><div class="value">${num("Sold")}</div>
      <div class="sub">${split(by("Sold"))}</div></div>
  </div>`;
  return `${tiles}
  <div class="grid cols-2" style="margin-top:16px">
    <div class="card"><h2>Units by listing status</h2>${hbars(L.status_counts.map(s => ({ label: s.label, value: s.count, tip: `${s.label}: ${s.count} unit${s.count === 1 ? "" : "s"}\n${s.machines} machine${s.machines === 1 ? "" : "s"}, ${s.other} hardware` })), v => v)}
      <p class="muted small">Every machine and every Items-tab entry (hardware) appears once. Sold comes from the Sales tab and overrides the Listings status.</p></div>
    <div class="card"><h2>How the potential is worked out</h2><dl class="kv">
      <dt>Sold</dt><dd>Actual net from the Sales tab, after fees and shipping: ${money(L.sold_net)}</dd>
      <dt>Listed and Pending</dt><dd>Asking price as entered, no fees or shipping taken off: ${money(L.asking_total)}</dd>
      <dt>Not counted</dt><dd>Needs listed, Not listed, Removed, and Listed or Pending rows with no asking price</dd>
      <dt>Profit potential</dt><dd>Net sales potential minus the lot cost</dd></dl></div>
  </div>
  <div class="card">
    <div class="filters">
      <input type="search" id="l-q" placeholder="Search ID, serial, model, spec, notes…" value="${esc(f.q)}">
      <select id="l-status">${options(statuses, f.status, "Any status")}</select>
      <select id="l-kind">${options(uniq(all.map(u => u.kind)), f.kind, "Machines and hardware")}</select>
      <span class="count">${rows.length} of ${all.length} units (${all.filter(u => u.kind === "Computer").length} machines, ${all.filter(u => u.kind === "Hardware").length} hardware)</span>
    </div>
    <div class="table-wrap"><table><thead>${sortHeader(cols, f)}</thead><tbody>${body || `<tr><td colspan="12" class="empty">Nothing matches these filters.</td></tr>`}</tbody>${foot}</table></div>
  </div>`;
}

function attentionCard(list) {
  const order = { critical: 0, serious: 1, warning: 2, info: 3 };
  const label = { critical: "Critical", serious: "Problem", warning: "Check", info: "Note" };
  const items = list.slice().sort((a, b) => order[a.level] - order[b.level] || cmp(a.asset_id, b.asset_id));
  const body = items.length ? `<ul class="attn">${items.map(a => `<li>
      ${statusBadge(a.level, label[a.level] || a.level)}
      <span class="kind">${a.asset_id ? assetLink(a.asset_id) + " · " : ""}${esc(a.kind)}</span>
      <span class="text">${esc(a.text)}</span></li>`).join("")}</ul>`
    : `<div class="empty">${statusBadge("good", "OK")} Nothing needs attention.</div>`;
  return `<div class="card"><div class="card-head"><h2>Needs attention</h2><span class="muted small">${items.length} item${items.length === 1 ? "" : "s"}</span></div>${body}</div>`;
}

function viewSales(d) {
  const f = state.sales;
  const q = f.q.trim().toLowerCase();
  let rows = d.sales.filter(s =>
    (!f.platform || s.platform === f.platform) &&
    (!f.type || s.type === f.type) &&
    (!q || [s.asset_id, s.order_id, s.model, s.buyer, s.notes, s.sold_with, s.payment, s.platform]
      .some(v => v && String(v).toLowerCase().includes(q))));
  rows = applySort(rows, f);
  // Band alternate orders so multi-machine orders read as one block.
  let band = false, prev = null;
  const cols = [
    { key: "date", label: "Date" }, { key: "order_id", label: "Order" }, { key: "asset_id", label: "Asset" },
    { key: "type", label: "Type" }, { key: "model", label: "Model" },
    { key: "price", label: "Price", num: true }, { key: "fees", label: "Fees", num: true },
    { key: "shipping", label: "Shipping", num: true }, { key: "net", label: "Net", num: true },
    { key: "platform", label: "Platform" }, { key: "payment", label: "Payment" }, { key: "buyer", label: "Buyer" },
    { key: "sold_with", label: "Sold with" }, { key: "notes", label: "Notes" }, { key: "check", label: "Check" },
  ];
  const body = rows.map(s => {
    const key = s.order_id || "_" + s.asset_id;
    if (key !== prev) band = !band;
    prev = key;
    const check = s.check === "OK" ? statusBadge("good", "OK") : statusBadge("serious", s.check);
    return `<tr class="${band ? "" : "band"}">
      <td class="nowrap">${fmtDate(s.date)}</td><td class="nowrap">${esc(s.order_id)}${s.machines_in_order > 1 ? ` <span class="dim small">(${s.machines_in_order})</span>` : ""}</td>
      <td class="nowrap">${assetLink(s.asset_id)}</td><td>${typeCell(s.type)}</td><td class="nowrap">${esc(s.model)}</td>
      <td class="num">${money(s.price)}</td><td class="num">${money(s.fees)}</td><td class="num">${money(s.shipping)}</td>
      <td class="num">${money(s.net)}</td><td class="nowrap">${esc(s.platform)}</td><td>${esc(s.payment)}</td>
      <td>${esc(s.buyer)}</td><td class="wrap">${esc(s.sold_with)}</td><td class="wrap">${esc(s.notes)}</td><td>${check}</td></tr>`;
  }).join("");
  const sum = k => rows.reduce((a, s) => a + (Number(s[k]) || 0), 0);
  const foot = rows.length ? `<tfoot><tr><td colspan="5">Total (${rows.length})</td>
      <td class="num">${money(sum("price"), { cents: !Number.isInteger(sum("price")) })}</td><td class="num">${money(sum("fees"))}</td>
      <td class="num">${money(sum("shipping"))}</td><td class="num">${money(sum("net"))}</td><td colspan="6"></td></tr></tfoot>` : "";
  const dayRows = d.by_day.slice().reverse().map(x => `<tr><td>${fmtDate(x.date)}</td><td class="num">${x.count}</td><td class="num">${money(x.gross)}</td><td class="num">${money(x.net)}</td><td class="num">${money(x.cumulative_net)}</td></tr>`).join("");
  return `<div class="card">
    <div class="filters">
      <input type="search" id="s-q" placeholder="Search asset, order, buyer, notes…" value="${esc(f.q)}">
      <select id="s-platform">${options(uniq(d.sales.map(s => s.platform)), f.platform, "All platforms")}</select>
      <select id="s-type">${options(uniq(d.sales.map(s => s.type)), f.type, "All types")}</select>
      <span class="count">${rows.length} of ${d.sales.length} sales</span>
    </div>
    ${d.sales.length ? `<div class="table-wrap"><table id="sales-table"><thead>${sortHeader(cols, f)}</thead><tbody>${body}</tbody>${foot}</table></div>`
      : `<div class="empty">No sales recorded yet.</div>`}
  </div>
  <div class="card"><h2>By day</h2><div class="table-wrap"><table><thead><tr><th>Date</th><th class="num">Units</th><th class="num">Gross</th><th class="num">Net</th><th class="num">Cumulative net</th></tr></thead><tbody>${dayRows || `<tr><td colspan="5" class="empty">No sales yet.</td></tr>`}</tbody></table></div></div>`;
}

function inventoryRows(d) {
  const f = state.inv;
  const q = f.q.trim().toLowerCase();
  return d.computers.filter(c =>
    (!f.type || c.type === f.type) &&
    (!f.model || c.model === f.model) &&
    (!f.status || (f.status === "__final" ? !!c.final_status : c.status === f.status)) &&
    (!f.win11 || c.win11_eff === f.win11) &&
    (f.sold === "" || (f.sold === "sold" ? c.is_sold : !c.is_sold)) &&
    (f.drive === "" || (f.drive === "yes" ? c.has_drive === true : c.has_drive !== true)) &&
    (!q || [c.asset_id, c.service_tag, c.model, c.cpu, c.drives, c.manual_notes, c.mac_addresses, c.bios_version]
      .some(v => v && String(v).toLowerCase().includes(q))));
}

function viewInventory(d) {
  const f = state.inv;
  const rows = applySort(inventoryRows(d), f, (c, k) => k === "listing_status"
    ? (c.listing && !c.is_sold ? STATUS_RANK[c.listing.status] ?? 9 : null) : c[k]);
  const cols = [
    { key: "asset_id", label: "Asset" }, { key: "type", label: "Type" }, { key: "model", label: "Model" },
    { key: "cpu_short", label: "CPU" }, { key: "cpu_gen", label: "Gen", num: true }, { key: "ram_gb", label: "RAM", num: true },
    { key: "drives", label: "Drive" }, { key: "win11_eff", label: "Win11" }, { key: "tpm_visible", label: "TPM" },
    { key: "status", label: "Status" }, { key: "is_sold", label: "Sold" }, { key: "listing_status", label: "Listing" },
    { key: "manual_notes", label: "Notes" },
  ];
  const body = rows.map(c => {
    const st = c.status === "OK" ? statusBadge("good", "OK") : statusBadge(c.source === "master-only" ? "info" : "warning", c.status || "–");
    const tpm = c.tpm_visible === null || c.tpm_visible === undefined ? "" : c.tpm_visible ? "Yes" : "No";
    return `<tr class="click" data-asset="${esc(c.asset_id)}">
      <td class="nowrap"><a href="#inventory/${encodeURIComponent(c.asset_id)}" data-asset="${esc(c.asset_id)}">${esc(c.asset_id)}</a></td>
      <td>${typeCell(c.type)}</td><td class="nowrap">${esc(c.model)}</td><td class="nowrap">${esc(c.cpu_short)}${noteMark(c, "cpu")}</td>
      <td class="num">${esc(c.cpu_gen)}</td><td class="num">${esc(ramLabel(c.ram_gb))}${noteMark(c, "ram")}</td>
      <td class="nowrap">${c.has_drive ? esc(c.drives) : c.source === "master-only" && !(c.note_overrides || {}).drives ? "" : '<span class="dim">none</span>'}${noteMark(c, "drives")}</td>
      <td class="nowrap">${esc(c.win11_eff)}</td><td>${tpm}</td><td>${st}</td>
      <td>${c.is_sold ? `<span class="badge sold">Sold${c.sale && c.sale.net !== null ? " · " + money(c.sale.net) : ""}</span>` : ""}</td>
      <td class="nowrap">${c.listing && !c.is_sold ? listingBadge(c.listing.status) + (c.listing.asking !== null ? " " + money(c.listing.asking) : "") : ""}</td>
      <td class="wrap small">${esc(c.manual_notes)}</td></tr>`;
  }).join("");
  const all = d.computers;
  return `<div class="card">
    <div class="filters">
      <input type="search" id="i-q" placeholder="Search asset, service tag, CPU, drive, MAC, notes…" value="${esc(f.q)}">
      <div class="seg" id="i-sold">
        ${[["unsold", "Unsold"], ["sold", "Sold"], ["", "All"]].map(([v, l]) => `<button type="button" data-v="${v}" class="${f.sold === v ? "on" : ""}">${l}</button>`).join("")}
      </div>
      <select id="i-type">${options(uniq(all.map(c => c.type)), f.type, "All types")}</select>
      <select id="i-model">${options(uniq(all.map(c => c.model)), f.model, "All models")}</select>
      <select id="i-win11">${options(uniq(all.map(c => c.win11_eff)), f.win11, "Win11: any")}</select>
      <select id="i-status">${options(uniq(all.map(c => c.status)), f.status, "Any status")}<option value="__final"${f.status === "__final" ? " selected" : ""}>Final status (DEAD, BIOS-PW)</option></select>
      <select id="i-drive"><option value="">Drive: any</option><option value="yes"${f.drive === "yes" ? " selected" : ""}>Has drive</option><option value="no"${f.drive === "no" ? " selected" : ""}>No drive</option></select>
      <button type="button" id="i-reset">Reset</button>
      <span class="count">${rows.length} of ${all.length} machines</span>
    </div>
    <div class="table-wrap"><table><thead>${sortHeader(cols, f)}</thead><tbody>${body || `<tr><td colspan="13" class="empty">No machines match these filters.</td></tr>`}</tbody></table></div>
    <p class="muted small">Click a row for every scanned field, Master notes, the sale, and the full scan file.</p>
  </div>`;
}

function viewFleet(d) {
  const scope = state.fleetScope;
  const list = d.computers.filter(c => scope === "all" || !c.is_sold);
  const sortDesc = rows => rows.sort((a, b) => b.value - a.value || cmp(a.label, b.label));
  const sortLabel = rows => rows.sort((a, b) => cmp(a.label, b.label));
  const scanned = list.filter(c => c.source === "scan");
  const charts = [
    ["Model", sortDesc(countBy(list, c => c.model))],
    ["CPU", sortDesc(countBy(scanned, c => c.cpu_short))],
    ["CPU generation", sortLabel(countBy(scanned, c => c.cpu_gen))],
    ["RAM", countBy(scanned, c => c.ram_gb).sort((a, b) => (parseFloat(a.label) || 0) - (parseFloat(b.label) || 0)).map(r => ({ ...r, label: r.label === "(not recorded)" ? r.label : r.label + " GB" }))],
    ["Windows 11 (effective)", sortDesc(countBy(list, c => c.win11_eff))],
    ["TPM visible to the scan", sortDesc(countBy(scanned, c => (c.tpm_visible ? "Yes" : "No")))],
    ["Drive installed", sortDesc(countBy(scanned, c => (c.has_drive ? "Yes" : "No")))],
    ["Type", sortDesc(countBy(list, c => c.type)), true],
    ["Scan status", sortDesc(countBy(list, c => c.status))],
  ];
  const chartCards = charts.map(([title, rows, typeLabels]) =>
    `<div class="card"><h2>${esc(title)}</h2>${hbars(rows, v => v, { typeLabels })}</div>`).join("");

  // Groups of identical machines - handy for writing one listing per group.
  const groups = new Map();
  for (const c of scanned) {
    const key = [c.model, c.cpu_short, c.ram_gb, c.has_drive ? c.drives : "no drive"].join("|");
    if (!groups.has(key)) groups.set(key, { model: c.model, cpu: c.cpu_short, ram: c.ram_gb, drive: c.has_drive ? c.drives : null, win11: new Set(), ids: [] });
    const g = groups.get(key);
    g.ids.push(c.asset_id);
    if (c.win11_eff) g.win11.add(c.win11_eff);
  }
  const grows = [...groups.values()].sort((a, b) => b.ids.length - a.ids.length || cmp(a.model, b.model));
  const gtable = `<div class="card"><div class="card-head"><h2>Identical-spec groups</h2><span class="muted small">model + CPU + RAM + drive</span></div>
    <div class="table-wrap"><table><thead><tr><th class="num">Count</th><th>Model</th><th>CPU</th><th class="num">RAM</th><th>Drive</th><th>Win11</th><th>Assets</th></tr></thead><tbody>
    ${grows.map(g => `<tr><td class="num">${g.ids.length}</td><td class="nowrap">${esc(g.model)}</td><td class="nowrap">${esc(g.cpu)}</td><td class="num">${esc(ramLabel(g.ram))}</td>
      <td class="nowrap">${g.drive ? esc(g.drive) : '<span class="dim">none</span>'}</td><td class="nowrap">${esc([...g.win11].join(", "))}</td>
      <td class="small">${g.ids.map(assetLink).join(", ")}</td></tr>`).join("") || `<tr><td colspan="7" class="empty">No machines.</td></tr>`}
    </tbody></table></div></div>`;

  return `<div class="filters">
      <div class="seg" id="f-scope">
        <button type="button" data-v="unsold" class="${scope === "unsold" ? "on" : ""}">Unsold machines</button>
        <button type="button" data-v="all" class="${scope === "all" ? "on" : ""}">All machines</button>
      </div>
      <span class="count">${list.length} machines${list.length !== scanned.length ? ` (${scanned.length} scanned)` : ""}</span>
    </div>
    <div class="grid cols-3">${chartCards}</div>
    ${gtable}`;
}

function hardwareRows(d) {
  const f = state.hw;
  const q = f.q.trim().toLowerCase();
  return d.hardware.filter(h =>
    (!f.type || h.type === f.type) &&
    (!f.condition || h.condition === f.condition) &&
    (!f.listing || (h.listing && !h.is_sold ? h.listing.status : "") === f.listing) &&
    (f.sold === "" || (f.sold === "sold" ? h.is_sold : !h.is_sold)) &&
    (!q || [h.id, h.serial, h.asset_id, h.type, h.make, h.model, h.key_spec, h.condition, h.notes]
      .some(v => v !== null && v !== undefined && String(v).toLowerCase().includes(q))));
}

function viewHardware(d) {
  const f = state.hw;
  const all = d.hardware;
  const getter = (h, k) => k === "listing" ? (h.listing && !h.is_sold ? STATUS_RANK[h.listing.status] ?? 9 : null)
    : k === "asking" ? (h.listing && !h.is_sold ? h.listing.asking : null)
    : k === "sold_net" ? (h.sale ? h.sale.net : null) : h[k];
  const rows = applySort(hardwareRows(d).map(h => ({ ...h, asset_id: h.id })), f, getter);
  const cols = [
    { key: "id", label: "ID" }, { key: "type", label: "Type" }, { key: "make", label: "Make" },
    { key: "model", label: "Model" }, { key: "key_spec", label: "Key spec" }, { key: "condition", label: "Condition" },
    { key: "listing", label: "Listing" }, { key: "asking", label: "Asking", num: true },
    { key: "sold_net", label: "Sold", num: true }, { key: "notes", label: "Notes" },
  ];
  const body = rows.map(h => {
    const clickable = h.id_unique;
    const idCell = h.id ? (clickable ? `<a href="#hardware/${encodeURIComponent(h.id)}" data-asset="${esc(h.id)}">${esc(h.id)}</a>` : esc(h.id))
      : '<span class="dim">no ID</span>';
    const l = h.listing && !h.is_sold ? h.listing : null;
    return `<tr${clickable ? ` class="click" data-asset="${esc(h.id)}"` : ""}>
      <td class="nowrap">${idCell}</td><td>${typeCell(h.type)}</td><td class="nowrap">${esc(h.make)}</td>
      <td class="nowrap">${esc(h.model)}</td><td class="nowrap small">${esc(h.key_spec)}</td><td>${esc(h.condition)}</td>
      <td class="nowrap">${l ? listingBadge(l.status) : ""}</td><td class="num">${l ? money(l.asking) : ""}</td>
      <td class="nowrap">${h.is_sold ? `<span class="badge sold">Sold${h.sale && h.sale.net !== null ? " · " + money(h.sale.net) : ""}</span>` : ""}</td>
      <td class="wrap small">${esc(h.notes)}</td></tr>`;
  }).join("");
  const noId = all.filter(h => !h.id).length;
  const bulk = d.bulk.length ? `<div class="table-wrap"><table><thead><tr>${d.bulk_cols.map(c => `<th>${esc(c)}</th>`).join("")}</tr></thead><tbody>
    ${d.bulk.map(b => `<tr>${d.bulk_cols.map(c => `<td>${esc(b[c])}</td>`).join("")}</tr>`).join("")}</tbody></table></div>`
    : `<div class="empty">No bins recorded on the Bulk tab yet.</div>`;
  return `<div class="card">
    <div class="filters">
      <input type="search" id="h-q" placeholder="Search ID, serial, make, model, spec, notes…" value="${esc(f.q)}">
      <div class="seg" id="h-sold">
        ${[["unsold", "Unsold"], ["sold", "Sold"], ["", "All"]].map(([v, l]) => `<button type="button" data-v="${v}" class="${f.sold === v ? "on" : ""}">${l}</button>`).join("")}
      </div>
      <select id="h-type">${options(uniq(all.map(h => h.type)), f.type, "All types")}</select>
      <select id="h-condition">${options(uniq(all.map(h => h.condition)), f.condition, "Any condition")}</select>
      <select id="h-listing">${options(uniq(all.filter(h => h.listing && !h.is_sold).map(h => h.listing.status)), f.listing, "Any listing status")}</select>
      <button type="button" id="h-reset">Reset</button>
      <span class="count">${rows.length} of ${all.length} items</span>
    </div>
    ${all.length ? `<div class="table-wrap"><table><thead>${sortHeader(cols, f)}</thead><tbody>${body || `<tr><td colspan="10" class="empty">No items match these filters.</td></tr>`}</tbody></table></div>`
      : `<div class="empty">No items recorded on the Items tab.</div>`}
    <p class="muted small">From the Items tab. Click a row for its details, listing and sale.${noId ? ` ${noId} item${noId === 1 ? " has" : "s have"} no asset ID or serial, so ${noId === 1 ? "it" : "they"} can't be matched to a sale or listing.` : ""}</p>
  </div>
  <div class="card"><h2>Bulk tab</h2>${bulk}</div>`;
}

function viewHwFleet(d) {
  const scope = state.hwFleetScope;
  const list = d.hardware.filter(h => scope === "all" || !h.is_sold);
  const ram = list.filter(h => h.ram);
  const sortDesc = rows => rows.sort((a, b) => b.value - a.value || cmp(a.label, b.label));
  const charts = [
    ["Type", sortDesc(countBy(list, h => h.type)), true],
    ["Make", sortDesc(countBy(list, h => h.make))],
    ["Condition", sortDesc(countBy(list, h => h.condition))],
    ["RAM by kind", sortDesc(countBy(ram, h => h.ram_group))],
    ["RAM by size", countBy(ram, h => h.ram.size_gb).sort((a, b) => (parseFloat(a.label) || 0) - (parseFloat(b.label) || 0))
      .map(r => ({ ...r, label: r.label === "(not recorded)" ? r.label : r.label + " GB" }))],
    ["RAM by generation", sortDesc(countBy(ram, h => h.ram.gen))],
    ["RAM by form factor", sortDesc(countBy(ram, h => h.ram.form))],
    ["RAM: ECC or not", sortDesc(countBy(ram, h => (h.ram.ecc ? "ECC (server)" : "Non-ECC")))],
  ];
  const bulkByType = new Map();
  for (const b of d.bulk) {
    const k = b["Item type"] || "(not recorded)";
    bulkByType.set(k, (bulkByType.get(k) || 0) + (Number(b["Quantity"]) || 0));
  }
  const bulkRows = [...bulkByType.entries()].map(([label, value]) => ({ label, value })).sort((a, b) => b.value - a.value);
  const cards = charts.map(([title, rows, typeLabels]) =>
    `<div class="card"><h2>${esc(title)}</h2>${hbars(rows, v => v, { typeLabels })}</div>`).join("") +
    `<div class="card"><h2>Bulk quantity by item type</h2>${d.bulk.length ? hbars(bulkRows, v => v)
      : `<div class="empty">No bins recorded on the Bulk tab yet.</div>`}<p class="muted small">From the Bulk tab (all bins; bins are not marked sold).</p></div>`;

  // Identical parts: same type, make, model and key spec.
  const groups = new Map();
  for (const h of list) {
    const key = [h.type, h.make, h.model, h.key_spec].join("|");
    if (!groups.has(key)) groups.set(key, { type: h.type, make: h.make, model: h.model, spec: h.key_spec, items: [] });
    groups.get(key).items.push(h);
  }
  const grows = [...groups.values()].sort((a, b) => b.items.length - a.items.length || cmp(a.model, b.model));
  const gtable = `<div class="card"><div class="card-head"><h2>Identical-part groups</h2><span class="muted small">type + make + model + key spec</span></div>
    <div class="table-wrap"><table><thead><tr><th class="num">Count</th><th>Type</th><th>Make</th><th>Model</th><th>Key spec</th><th>IDs</th></tr></thead><tbody>
    ${grows.map(g => `<tr><td class="num">${g.items.length}</td><td>${typeCell(g.type)}</td><td class="nowrap">${esc(g.make)}</td><td class="nowrap">${esc(g.model)}</td>
      <td class="nowrap small">${esc(g.spec)}</td>
      <td class="small">${g.items.map(h => h.id ? (h.id_unique ? assetLink(h.id) : esc(h.id)) : '<span class="dim">no ID</span>').join(", ")}</td></tr>`).join("")
      || `<tr><td colspan="6" class="empty">No items.</td></tr>`}
    </tbody></table></div></div>`;

  return `<div class="filters">
      <div class="seg" id="hf-scope">
        <button type="button" data-v="unsold" class="${scope === "unsold" ? "on" : ""}">Unsold items</button>
        <button type="button" data-v="all" class="${scope === "all" ? "on" : ""}">All items</button>
      </div>
      <span class="count">${list.length} items (${ram.length} RAM)</span>
    </div>
    <div class="grid cols-3">${cards}</div>
    ${gtable}`;
}

function viewNotes(d) {
  const cols = d.master_cols;
  const rows = d.master.slice().sort((a, b) => cmp(a["Asset ID"], b["Asset ID"]));
  return `<div class="card"><div class="card-head"><h2>Master tab</h2><span class="muted small">${rows.length} rows · exceptions, overrides and notes</span></div>
    ${rows.length ? `<div class="table-wrap"><table><thead><tr>${cols.map(c => `<th>${esc(c)}</th>`).join("")}</tr></thead><tbody>
    ${rows.map(m => `<tr>${cols.map(c => `<td class="${c === "Notes" ? "wrap" : "nowrap"}">${c === "Asset ID" ? assetLink(m[c]) : esc(m[c])}</td>`).join("")}</tr>`).join("")}
    </tbody></table></div>` : `<div class="empty">The Master tab is empty.</div>`}</div>
    ${attentionCard(d.attention)}`;
}

function viewData(d) {
  const m = d.meta;
  const sheets = Object.entries(m.sheets).map(([k, v]) => `<tr><td>${esc(k)}</td><td class="num">${v === null ? '<span class="dim">–</span>' : v}</td></tr>`).join("");
  const checks = m.checks.map(c => `<tr><td>${esc(c.label)}</td><td class="num">${esc(c.dashboard)}</td><td class="num">${esc(c.workbook)}</td><td>${c.ok ? statusBadge("good", "Match") : statusBadge("serious", "Mismatch")}</td></tr>`).join("");
  const summary = d.summary_rows.map(r => {
    const isHead = r.length === 1;
    return isHead ? `<tr><td colspan="2"><strong>${esc(r[0])}</strong></td></tr>`
      : `<tr><td>${esc(r[0])}</td><td class="num">${esc(r.slice(1).filter(v => v !== null).join(" · "))}</td></tr>`;
  }).join("");
  const health = healthCard(d.health);
  const lists = Object.entries(d.lists).map(([k, v]) => `<tr><td class="nowrap">${esc(k)}</td><td>${esc(v.join(", "))}</td></tr>`).join("");
  return `${health}<div class="grid cols-2">
    <div class="card"><h2>Source file</h2><dl class="kv">
      <dt>File</dt><dd>${esc(m.file)}</dd>
      <dt>Last saved</dt><dd>${fmtDateTime(m.mtime)}</dd>
      <dt>Size</dt><dd>${(m.size / 1024).toFixed(0)} KB</dd>
      <dt>Read by dashboard</dt><dd>${fmtDateTime(m.read_at)}</dd>
      <dt>Scan files folder</dt><dd>${m.raw_dir_found ? "Found" : "Not found (scan text view is unavailable)"}</dd>
    </dl>
    <h2 style="margin-top:16px">Rows per tab</h2><table><tbody>${sheets}</tbody></table></div>
    <div class="card"><h2>Cross-check with the Selling Dashboard tab</h2>
      ${checks ? `<table><thead><tr><th>Figure</th><th class="num">This dashboard</th><th class="num">Workbook</th><th></th></tr></thead><tbody>${checks}</tbody></table>`
        : `<div class="empty">No totals found to compare.</div>`}
      <p class="muted small">Totals here are added up from the Sales tab's calculated columns. They should always match the workbook.</p>
      <h2 style="margin-top:16px">Warnings</h2>
      ${m.warnings.length ? `<ul>${m.warnings.map(w => `<li>${esc(w)}</li>`).join("")}</ul>` : `<div class="empty">${statusBadge("good", "OK")} None.</div>`}
    </div>
  </div>
  <div class="grid cols-2">
    <div class="card"><h2>Summary tab</h2><table><tbody>${summary}</tbody></table></div>
    <div class="card"><h2>Lists tab (drop-down values)</h2><table><tbody>${lists}</tbody></table></div>
  </div>`;
}

function healthCard(h) {
  if (!h || !h.available) {
    return `<div class="card"><h2>Workbook health</h2><div class="empty">${statusBadge("warning", "Not checked")} ${esc(h && h.error ? "The check failed: " + h.error : "No reference copy (baseline.json) was found, so formulas and drop-downs were not checked.")}</div></div>`;
  }
  const rows = h.sections.map(sct => `<tr><td>${esc(sct.area)}</td><td class="num">${sct.checked.toLocaleString("en-US")}</td>
    <td>${sct.problems ? statusBadge("serious", `${sct.problems} problem${sct.problems === 1 ? "" : "s"}`) : statusBadge("good", "All in place")}</td></tr>`).join("");
  const label = { serious: "Problem", warning: "Check" };
  const probs = h.problems.length ? `<ul class="attn">${h.problems.map(p => `<li>${statusBadge(p.level, label[p.level] || p.level)}
      <span class="kind">${esc(p.where)}</span><span class="text">${esc(p.text)}</span></li>`).join("")}</ul>` : "";
  return `<div class="card"><div class="card-head"><h2>Workbook health</h2>
      <span class="muted small">${h.checked.toLocaleString("en-US")} checks · ${h.ok ? "everything in place" : `${h.problems.length} problem${h.problems.length === 1 ? "" : "s"}`}</span></div>
    <div class="table-wrap"><table><thead><tr><th>What is checked</th><th class="num">Cells or rules</th><th>Result</th></tr></thead><tbody>${rows}</tbody></table></div>
    ${probs}
    <p class="muted small">Compared with a reference copy of the workbook's formulas, drop-downs and shading rules (formula text only, no data). Fill-down columns must hold the same formula in every row from 2 to 500.</p></div>`;
}

/* ---------------- instructions ---------------- */

let guideHtml = null, guideError = null, guideLoading = false;

function viewInstructions(d) {
  if (guideHtml === null) {
    if (!guideLoading && !guideError) {
      guideLoading = true;
      fetch("/static/instructions.html", { cache: "no-store" })
        .then(r => { if (!r.ok) throw new Error(`HTTP ${r.status}`); return r.text(); })
        .then(t => { guideHtml = t; })
        .catch(e => { guideError = e.message; })
        .finally(() => { guideLoading = false; if (state.tab === "instructions") render(); });
    }
    return guideError ? `<div class="card">${statusBadge("serious", "Error")} Could not load the instructions: ${esc(guideError)}</div>`
      : `<p class="muted pad">Loading…</p>`;
  }
  return guideHtml;
}

/* ---------------- drawer ---------------- */

const SCAN_FIELDS = [
  ["service_tag", "Service tag"], ["vendor", "Vendor"], ["model", "Model"], ["cpu", "CPU"], ["cpu_gen", "CPU gen"],
  ["ram_mb", "RAM (scanned, MB)"], ["ram_mb_effective", "RAM (effective, MB)"], ["dimms", "DIMMs"], ["drives", "Drives"],
  ["mac_addresses", "MAC addresses"], ["bios_version", "BIOS version"], ["tpm", "TPM"], ["bios_tpm_settings", "BIOS TPM settings"],
  ["win11", "Win11 (scan)"], ["win11_override", "Win11 override"], ["win11_effective", "Win11 (effective)"],
  ["low_ram_under_4gb", "Low RAM (< 4 GB)"], ["no_drive_64gb_plus", "No drive ≥ 64 GB"], ["status", "Status"],
  ["problems", "Problems"], ["detail_file", "Scan file"],
];

function openHardwareDrawer(h) {
  state.openAsset = h.id;
  $("#drawer-title").innerHTML = `<h2>${esc(h.id)} · ${esc([h.make, h.model].filter(Boolean).join(" ") || "Unknown model")}</h2>
    <div style="margin-top:6px;display:flex;gap:6px;flex-wrap:wrap">
      ${h.is_sold ? `<span class="badge sold">Sold</span>` : `<span class="badge">In stock</span>`}
      ${h.type ? `<span class="badge">${typeCell(h.type)}</span>` : ""}
      ${h.condition ? `<span class="badge">${esc(h.condition)}</span>` : ""}
    </div>`;
  const kv = pairs => `<dl class="kv">${pairs.filter(([, v]) => v !== null && v !== undefined && v !== "").map(([k, v]) => `<dt>${esc(k)}</dt><dd>${v}</dd>`).join("")}</dl>`;
  let html = "";
  if (h.notes) html += `<div class="note">${esc(h.notes)}</div>`;
  html += `<h3>Items tab (row ${esc(h.row)})</h3>` + kv([
    ["Asset ID", esc(h.asset_id)], ["Serial", esc(h.serial)], ["Type", typeCell(h.type)], ["Make", esc(h.make)],
    ["Model", esc(h.model)], ["Key spec", esc(h.key_spec)], ["RAM kind", esc(h.ram_group)], ["Condition", esc(h.condition)],
    ...Object.entries(h.extra || {}).map(([k, v]) => [k, esc(v)]),
  ]);
  if (h.sale) {
    const s = h.sale;
    html += `<h3>Sale</h3>` + kv([
      ["Date", fmtDate(s.date)], ["Order", esc(s.order_id)], ["Price", money(s.price)], ["Fees", money(s.fees)],
      ["Shipping", money(s.shipping)], ["Net", `<strong>${money(s.net)}</strong>`], ["Platform", esc(s.platform)],
      ["Payment", esc(s.payment)], ["Buyer", esc(s.buyer)], ["Notes", esc(s.notes)],
      ["Check", s.check === "OK" ? statusBadge("good", "OK") : statusBadge("serious", s.check)],
    ]);
  }
  html += listingSection(h.listing, h.is_sold);
  $("#drawer-body").innerHTML = html;
  $("#drawer-body").scrollTop = 0;
  $("#drawer").hidden = false;
  $("#scrim").hidden = false;
}

function listingSection(l, isSold) {
  if (!l || (!l.has_listing_row && isSold)) return "";
  const kv = pairs => `<dl class="kv">${pairs.filter(([, v]) => v !== null && v !== undefined && v !== "").map(([k, v]) => `<dt>${esc(k)}</dt><dd>${v}</dd>`).join("")}</dl>`;
  return `<h3>Listing</h3>` + (l.has_listing_row ? kv([
    ["Status", listingBadge(l.status)], ["Status on the Listings tab", esc(l.listing_status)],
    ["Asking price", money(l.asking)], ["Listed on", fmtDate(l.date_listed)], ["Platform", esc(l.platform)],
    ["Listing", l.url ? `<a href="${esc(l.url)}" target="_blank" rel="noopener noreferrer">${esc(l.url)}</a>` : ""],
    ["Notes", esc(l.listing_notes)],
    ["In the potential", l.status === "Sold" ? "At its sale net" : l.counts ? "Yes, at the asking price" : "No"],
  ]) : `<div class="empty">Not on the Listings tab, so it shows as Not listed.</div>`);
}

function openDrawer(id) {
  const d = state.data;
  const c = d && d.computers.find(x => x.asset_id === id);
  if (!c) {
    const h = d && (d.hardware || []).find(x => x.id_unique && x.id === id);
    if (h) openHardwareDrawer(h); else closeDrawer();
    return;
  }
  state.openAsset = id;
  $("#drawer-title").innerHTML = `<h2>${esc(c.asset_id)} · ${esc(c.model || "Unknown model")}</h2>
    <div style="margin-top:6px;display:flex;gap:6px;flex-wrap:wrap">
      ${c.is_sold ? `<span class="badge sold">Sold</span>` : `<span class="badge">In stock</span>`}
      ${c.status === "OK" ? statusBadge("good", "Status OK") : statusBadge(c.source === "master-only" ? "info" : "warning", c.status || "No status")}
      ${c.type ? `<span class="badge">${typeCell(c.type)}</span>` : ""}
      ${c.source === "master-only" ? `<span class="badge">Master tab only, not scanned</span>` : ""}
      ${c.final_status && c.source === "scan" ? `<span class="badge">Master status ${esc(c.final_status)} (final)</span>` : ""}
    </div>`;

  const kv = pairs => `<dl class="kv">${pairs.filter(([, v]) => v !== null && v !== undefined && v !== "").map(([k, v]) => `<dt>${esc(k)}</dt><dd>${v}</dd>`).join("")}</dl>`;

  let html = "";
  const m = c.master;
  if (m) {
    html += `<h3>Master tab</h3>`;
    if (m.Notes) html += `<div class="note">${esc(m.Notes)}</div>`;
    html += kv(Object.entries(m).filter(([k]) => k !== "Notes" && k !== "Asset ID").map(([k, v]) => [k, esc(v)]));
  }
  if (c.sale) {
    const s = c.sale;
    html += `<h3>Sale</h3>` + kv([
      ["Date", fmtDate(s.date)], ["Order", esc(s.order_id) + (s.machines_in_order > 1 ? ` (${s.machines_in_order} machines)` : "")],
      ["Price", money(s.price)], ["Fees", money(s.fees)], ["Shipping", money(s.shipping)], ["Net", `<strong>${money(s.net)}</strong>`],
      ["Platform", esc(s.platform)], ["Payment", esc(s.payment)], ["Buyer", esc(s.buyer)], ["Sold with", esc(s.sold_with)],
      ["Notes", esc(s.notes)], ["Check", s.check === "OK" ? statusBadge("good", "OK") : statusBadge("serious", s.check)],
    ]);
  }
  html += listingSection(c.listing, c.is_sold);
  const ov = c.note_overrides || {};
  if (Object.keys(ov).length) {
    html += `<h3>From the Master note (used instead of the scan)</h3>` + kv([
      ["CPU", ov.cpu ? esc(ov.cpu) + (c.source === "scan" ? ` <span class="dim small">scan: ${esc(c.scan_cpu || "none")}</span>` : "") : ""],
      ["RAM", ov.ram ? esc(ov.ram) + (c.source === "scan" ? ` <span class="dim small">scan: ${esc(ramLabel(c.scan_ram_gb) || "none")}</span>` : "") : ""],
      ["Drives", ov.drives ? esc(ov.drives) + (c.source === "scan" ? ` <span class="dim small">scan: ${esc(c.scan_drives || "none")}</span>` : "") : ""],
    ]);
  }
  if (c.source === "scan") {
    const scanVal = k => (k === "drives" && "scan_drives" in c ? c.scan_drives : k === "cpu" && "scan_cpu" in c ? c.scan_cpu : c[k]);
    html += `<h3>Scan row (Computers tab)</h3>` + kv(SCAN_FIELDS.map(([k, label]) => [label, esc(scanVal(k))]));
  }
  html += `<h3>Scan file</h3>`;
  if (!c.scan_files || !c.scan_files.length) {
    html += `<div class="empty">No scan file in the inventory-raw folder for this asset.</div>`;
  } else {
    const def = c.scan_files.includes(c.detail_file) ? c.detail_file : c.scan_files[c.scan_files.length - 1];
    html += (c.scan_files.length > 1
      ? `<div class="filters"><select id="scan-file">${c.scan_files.map(f => `<option${f === def ? " selected" : ""}>${esc(f)}</option>`).join("")}</select>
         <span class="muted small">${c.scan_files.length} scans. The Computers tab uses ${esc(c.detail_file || "?")}.</span></div>`
      : `<div class="muted small" style="margin-bottom:6px">${esc(def)}</div>`) + `<div id="scan-body" class="muted">Loading…</div>`;
    setTimeout(() => loadScan(c.asset_id, def), 0);
  }
  $("#drawer-body").innerHTML = html;
  $("#drawer-body").scrollTop = 0;
  $("#drawer").hidden = false;
  $("#scrim").hidden = false;
}

async function loadScan(id, file) {
  const box = $("#scan-body");
  if (!box) return;
  box.textContent = "Loading…";
  try {
    const r = await fetch(`/api/raw/${encodeURIComponent(id)}?file=${encodeURIComponent(file)}`);
    if (!r.ok) throw new Error(r.status === 404 ? "File not found in the inventory-raw folder." : `HTTP ${r.status}`);
    const text = await r.text();
    if (state.openAsset !== id) return;
    // Files are split into sections that start with a line "== NAME".
    const parts = [];
    let cur = { title: "Header", lines: [] };
    for (const line of text.split(/\r?\n/)) {
      const mm = /^==\s*(.+?)\s*$/.exec(line);
      if (mm) { if (cur.lines.some(l => l.trim())) parts.push(cur); cur = { title: mm[1], lines: [] }; }
      else cur.lines.push(line);
    }
    if (cur.lines.some(l => l.trim())) parts.push(cur);
    box.className = "";
    box.innerHTML = parts.map((p, i) => `<details class="scan"${i < 2 || /WIN11/i.test(p.title) ? " open" : ""}><summary>${esc(p.title)}</summary><pre>${esc(p.lines.join("\n").replace(/^\n+|\n+$/g, ""))}</pre></details>`).join("");
  } catch (e) {
    box.className = "muted";
    box.textContent = "Could not load the scan file: " + e.message;
  }
}

function closeDrawer() {
  state.openAsset = null;
  $("#drawer").hidden = true;
  $("#scrim").hidden = true;
}

/* ---------------- render & routing ---------------- */

const VIEWS = { overview: viewOverview, sales: viewSales, listings: viewListings, inventory: viewInventory, fleet: viewFleet, hardware: viewHardware, hwfleet: viewHwFleet, notes: viewNotes, data: viewData, instructions: viewInstructions };

function render() {
  for (const a of document.querySelectorAll("#tabs a")) a.classList.toggle("active", a.dataset.tab === state.tab);
  const main = $("#main");
  if (!state.data) {
    main.innerHTML = `<div class="card">${statusBadge("critical", "Error")} <strong>The workbook could not be loaded.</strong><p>${esc(state.error || "Unknown error")}</p>
      <p class="muted">Check that the workbook is mounted into the container at the path set by WORKBOOK_PATH and that the container user can read it. This page retries every minute.</p></div>`;
    return;
  }
  const d = state.data;
  $("#stamp").textContent = `Workbook saved ${fmtDateTime(d.meta.mtime)}`;
  const banner = $("#banner");
  const msgs = [];
  if (d.meta.error) msgs.push(statusBadge("serious", "Read error") + " " + esc(d.meta.error) + " Showing the last version that loaded.");
  for (const w of d.meta.warnings) msgs.push(statusBadge("warning", "Warning") + " " + esc(w));
  banner.hidden = !msgs.length;
  banner.innerHTML = msgs.map(x => `<div class="item">${x}</div>`).join("");

  // Keep focus and caret in search boxes across re-renders.
  const active = document.activeElement;
  const focusId = active && active.id;
  const caret = active && "selectionStart" in active ? active.selectionStart : null;
  const scrollY = window.scrollY;
  main.innerHTML = (VIEWS[state.tab] || viewOverview)(d);
  if (focusId) {
    const el = document.getElementById(focusId);
    if (el && main.contains(el)) { el.focus(); if (caret !== null && el.setSelectionRange) { try { el.setSelectionRange(caret, caret); } catch (_) { /* not a text input */ } } }
  }
  window.scrollTo(0, scrollY);
}

function route() {
  const h = decodeURIComponent(location.hash.replace(/^#/, ""));
  const [tab, asset] = h.split("/");
  const newTab = VIEWS[tab] ? tab : "overview";
  if (newTab !== state.tab) { state.tab = newTab; render(); window.scrollTo(0, 0); }
  if (asset) openDrawer(asset);
  else if (state.openAsset) closeDrawer();
}

/* ---------------- events ---------------- */

document.addEventListener("input", e => {
  const id = e.target.id;
  if (id === "s-q") { state.sales.q = e.target.value; render(); }
  if (id === "i-q") { state.inv.q = e.target.value; render(); }
  if (id === "l-q") { state.listings.q = e.target.value; render(); }
  if (id === "h-q") { state.hw.q = e.target.value; render(); }
});

document.addEventListener("change", e => {
  const id = e.target.id, v = e.target.value;
  const map = { "s-platform": ["sales", "platform"], "s-type": ["sales", "type"], "i-type": ["inv", "type"], "i-model": ["inv", "model"],
    "i-win11": ["inv", "win11"], "i-status": ["inv", "status"], "i-drive": ["inv", "drive"],
    "l-status": ["listings", "status"], "l-kind": ["listings", "kind"] };
  if (map[id]) { state[map[id][0]][map[id][1]] = v; render(); }
  const hw = { "h-type": "type", "h-condition": "condition", "h-listing": "listing" };
  if (hw[id]) { state.hw[hw[id]] = v; render(); }
  if (id === "scan-file" && state.openAsset) loadScan(state.openAsset, v);
});

document.addEventListener("click", e => {
  const t = e.target;
  const seg = t.closest(".seg button");
  if (seg) {
    const group = seg.parentElement.id;
    if (group === "i-sold") state.inv.sold = seg.dataset.v;
    if (group === "f-scope") state.fleetScope = seg.dataset.v;
    if (group === "h-sold") state.hw.sold = seg.dataset.v;
    if (group === "hf-scope") state.hwFleetScope = seg.dataset.v;
    render();
    return;
  }
  if (t.id === "h-reset") {
    Object.assign(state.hw, { q: "", type: "", condition: "", listing: "", sold: "unsold" });
    render();
    return;
  }
  if (t.id === "i-reset") {
    Object.assign(state.inv, { q: "", type: "", model: "", status: "", sold: "unsold", win11: "", drive: "" });
    render();
    return;
  }
  const th = t.closest("th.sortable");
  if (th) {
    const s = { sales: state.sales, listings: state.listings, hardware: state.hw }[state.tab] || state.inv;
    if (s.sort === th.dataset.sort) s.dir = -s.dir; else { s.sort = th.dataset.sort; s.dir = 1; }
    render();
    return;
  }
  const link = t.closest("a[data-asset]");
  if (link) {
    e.preventDefault();
    // Open the machine panel over whichever tab is showing.
    location.hash = `#${state.tab}/${encodeURIComponent(link.dataset.asset)}`;
    return;
  }
  const row = t.closest("tr.click[data-asset]");
  if (row) { location.hash = `#${state.tab}/${encodeURIComponent(row.dataset.asset)}`; return; }
  if (t.id === "drawer-close" || t.id === "scrim") { location.hash = `#${state.tab}`; return; }
  if (t.id === "refresh") load(true);
  const jump = t.closest("a[data-jump]");
  if (jump) {
    e.preventDefault();
    const el = document.getElementById(jump.dataset.jump);
    if (el) el.scrollIntoView({ behavior: "smooth", block: "start" });
  }
});

document.addEventListener("keydown", e => {
  if (e.key === "Escape" && state.openAsset) location.hash = `#${state.tab}`;
});

window.addEventListener("hashchange", route);

// Pop-up for type codes (the only elements with data-tip).
const tip = $("#tip");
document.addEventListener("mousemove", e => {
  const el = e.target.closest && e.target.closest("[data-tip]");
  if (!el) { tip.hidden = true; return; }
  tip.textContent = el.dataset.tip;
  tip.hidden = false;
  const pad = 14, w = tip.offsetWidth, h = tip.offsetHeight;
  let x = e.clientX + pad, y = e.clientY + pad;
  if (x + w > window.innerWidth - 4) x = e.clientX - w - pad;
  if (y + h > window.innerHeight - 4) y = e.clientY - h - pad;
  tip.style.left = x + "px";
  tip.style.top = y + "px";
});
document.addEventListener("mouseleave", () => { tip.hidden = true; });

/* ---------------- data loading ---------------- */

async function load(manual) {
  const btn = $("#refresh");
  if (manual) btn.textContent = "Refreshing…";
  try {
    const r = await fetch("/api/data", { cache: "no-store" });
    const body = await r.json().catch(() => ({ error: `HTTP ${r.status}` }));
    if (!r.ok) throw new Error(body.error || `HTTP ${r.status}`);
    state.data = body;
    state.error = null;
  } catch (e) {
    state.error = e.message;
    if (state.data) {
      state.data.meta = Object.assign({}, state.data.meta, { error: "Lost contact with the dashboard server: " + e.message });
    }
  }
  if (manual) btn.textContent = "Refresh";
  render();
  if (state.openAsset) openDrawer(state.openAsset);
}

async function poll() {
  try {
    const r = await fetch("/api/status", { cache: "no-store" });
    const s = await r.json();
    const d = state.data;
    if (!d || s.mtime !== d.meta.mtime || (s.error || null) !== (d.meta.error || null)) await load();
  } catch (_) {
    if (state.data && !state.data.meta.error) {
      state.data.meta = Object.assign({}, state.data.meta, { error: "Lost contact with the dashboard server." });
      render();
    }
  }
}

(async function start() {
  const h = decodeURIComponent(location.hash.replace(/^#/, ""));
  const tab = h.split("/")[0];
  if (VIEWS[tab]) state.tab = tab;
  await load();
  route();
  setInterval(poll, POLL_MS);
})();
