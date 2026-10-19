/* AI-Augmented Sector Teardown — dashboard logic.
   Loads the pipeline-generated companies.json, renders the leaderboard and
   company table, then layers live quotes from Stooq (~15 min delayed) on top,
   recomputing implied upside in the browser. Falls back to stored prices. */

const STOOQ_URL = "https://stooq.com/q/l/?s={syms}&f=sd2t2ohlcv&h&e=csv";

let COMPANIES = [];
let META = {};
let activeFilter = "All";

const fmt$ = (n) =>
  n == null || isNaN(n) ? "—" : "$" + Number(n).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtPct = (n) =>
  n == null || isNaN(n) ? "—" : (n > 0 ? "+" : "") + Number(n).toFixed(1) + "%";

function median(arr) {
  const s = [...arr].sort((a, b) => a - b);
  const m = Math.floor(s.length / 2);
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
}

async function init() {
  try {
    const res = await fetch("data/companies.json");
    if (!res.ok) throw new Error("companies.json not found");
    const data = await res.json();
    META = data.meta || {};
    COMPANIES = (data.companies || []).map((c) => ({ ...c, livePrice: null }));
  } catch (e) {
    document.getElementById("live-label").textContent = "Data not yet published — check back soon.";
    return;
  }
  renderMeta();
  renderStats();
  renderFilters();
  renderLeaderboard();
  renderTable();
  loadLiveQuotes();
}

function renderMeta() {
  document.getElementById("meta-model-asof").textContent = META.universe_asof || "—";
  document.getElementById("meta-price-asof").textContent = META.price_asof || "—";
}

function renderStats() {
  const n = COMPANIES.length;
  const buys = COMPANIES.filter((c) => c.recommendation === "BUY").length;
  const holds = COMPANIES.filter((c) => c.recommendation === "HOLD").length;
  const passes = COMPANIES.filter((c) => c.recommendation === "PASS").length;
  const sectors = new Set(COMPANIES.map((c) => c.sector)).size;
  document.getElementById("stats").innerHTML = `
    <div class="stat"><div class="k">Companies</div><div class="v">${n}</div></div>
    <div class="stat"><div class="k">Sectors</div><div class="v">${sectors}</div></div>
    <div class="stat"><div class="k">Buy</div><div class="v buy">${buys}</div></div>
    <div class="stat"><div class="k">Hold</div><div class="v hold">${holds}</div></div>
    <div class="stat"><div class="k">Pass</div><div class="v pass">${passes}</div></div>`;
}

function sectorStats() {
  const bySector = {};
  COMPANIES.forEach((c) => {
    (bySector[c.sector] = bySector[c.sector] || []).push(c);
  });
  return Object.entries(bySector)
    .map(([sector, cs]) => ({
      sector,
      medianUpside: median(cs.map((c) => c.upside_pct)),
      buys: cs.filter((c) => c.recommendation === "BUY").length,
      holds: cs.filter((c) => c.recommendation === "HOLD").length,
      passes: cs.filter((c) => c.recommendation === "PASS").length,
      tickers: cs.map((c) => c.ticker).join(", "),
    }))
    .sort((a, b) => b.medianUpside - a.medianUpside);
}

function renderLeaderboard() {
  const rows = sectorStats();
  const maxAbs = Math.max(...rows.map((r) => Math.abs(r.medianUpside)), 1);
  document.getElementById("leaderboard").innerHTML = rows
    .map((r, i) => {
      const w = Math.min(100, (Math.abs(r.medianUpside) / maxAbs) * 100);
      const cls = r.medianUpside >= 0 ? "pos" : "neg";
      return `
      <div class="lb-row">
        <div class="lb-rank">${i + 1}</div>
        <div>
          <div class="lb-name">${r.sector}</div>
          <div class="lb-sub">${r.tickers} · ${r.buys} buy / ${r.holds} hold / ${r.passes} pass</div>
          <div class="lb-bar"><div class="lb-fill ${r.medianUpside < 0 ? "neg" : ""}" style="width:${w}%"></div></div>
        </div>
        <div class="lb-upside ${cls}">${fmtPct(r.medianUpside)}<small>median upside</small></div>
      </div>`;
    })
    .join("");
}

function renderFilters() {
  const sectors = ["All", ...new Set(COMPANIES.map((c) => c.sector))];
  document.getElementById("filters").innerHTML = sectors
    .map((s) => `<button class="${s === activeFilter ? "active" : ""}" data-sector="${s}">${s}</button>`)
    .join("");
  document.querySelectorAll("#filters button").forEach((b) =>
    b.addEventListener("click", () => {
      activeFilter = b.dataset.sector;
      renderFilters();
      renderTable();
    })
  );
}

function effectivePrice(c) {
  return c.livePrice != null ? c.livePrice : c.price;
}
function effectiveUpside(c) {
  const p = effectivePrice(c);
  return p ? ((c.dcf_base - p) / p) * 100 : null;
}

function renderTable() {
  const rows = COMPANIES.filter((c) => activeFilter === "All" || c.sector === activeFilter)
    .sort((a, b) => effectiveUpside(b) - effectiveUpside(a));
  // scale the bull/bear range bars across the visible set
  const allVals = rows.flatMap((c) => [c.dcf_bear, c.dcf_bull]).filter((v) => v != null);
  const lo = Math.min(...allVals), hi = Math.max(...allVals);
  const span = hi - lo || 1;

  document.getElementById("company-rows").innerHTML = rows
    .map((c) => {
      const price = effectivePrice(c);
      const upside = effectiveUpside(c);
      const live = c.livePrice != null;
      const left = ((c.dcf_bear - lo) / span) * 100;
      const width = Math.max(2, ((c.dcf_bull - c.dcf_bear) / span) * 100);
      const basePos = ((c.dcf_base - lo) / span) * 100;
      return `
      <tr>
        <td><span class="ticker">${c.ticker}</span><br><span class="cname">${c.name}</span></td>
        <td>${c.sector}</td>
        <td class="num">${fmt$(price)}${live ? "" : '<br><span class="stale">as of ' + c.price_asof + "</span>"}</td>
        <td class="num">${fmt$(c.dcf_base)}<br><span class="stale">model ${c.model_asof}</span></td>
        <td class="num upside ${upside >= 0 ? "pos" : "neg"}">${fmtPct(upside)}</td>
        <td>
          <div class="range-bar">
            <div class="range-fill" style="left:${left}%;width:${width}%"></div>
            <div class="range-base" style="left:${basePos}%"></div>
          </div>
          <div class="range-lbl">${fmt$(c.dcf_bear)} – ${fmt$(c.dcf_bull)}</div>
        </td>
        <td><span class="badge ${c.recommendation.toLowerCase()}">${c.recommendation}</span></td>
      </tr>`;
    })
    .join("");
}

function parseStooq(csv) {
  const out = {};
  const lines = csv.trim().split("\n");
  for (let i = 1; i < lines.length; i++) {
    const cols = lines[i].split(",");
    if (cols.length < 8) continue;
    const sym = cols[0].replace(".US", "").toUpperCase();
    const close = parseFloat(cols[6]);
    if (sym && !isNaN(close) && close > 0) out[sym] = close;
  }
  return out;
}

async function loadLiveQuotes() {
  const label = document.getElementById("live-label");
  const dot = document.getElementById("live-dot");
  try {
    const syms = COMPANIES.map((c) => c.ticker.toLowerCase() + ".us").join(",");
    const res = await fetch(STOOQ_URL.replace("{syms}", encodeURIComponent(syms)));
    if (!res.ok) throw new Error("quote fetch failed");
    const live = parseStooq(await res.text());
    let n = 0;
    COMPANIES.forEach((c) => {
      if (live[c.ticker] != null) { c.livePrice = live[c.ticker]; n++; }
    });
    if (n > 0) {
      dot.classList.add("on");
      const now = new Date();
      label.textContent = `Live quotes · ~15 min delayed · updated ${now.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" })} (${n}/${COMPANIES.length} tickers)`;
      renderTable();
      renderLeaderboard();
    } else {
      throw new Error("no quotes parsed");
    }
  } catch (e) {
    label.textContent = `Live feed unavailable — showing stored prices as of ${META.price_asof || "—"}`;
  }
}

document.addEventListener("DOMContentLoaded", init);
