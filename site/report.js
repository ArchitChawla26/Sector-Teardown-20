/* Sector Report — live tables.
   Renders the leaderboard and full scorecard from the same
   pipeline-generated companies.json the dashboard uses, so the paper's
   outcomes refresh daily with prices. DCF values stay at the model date. */

const SHORT_SECTOR = {
  "Diversified Financials": "Div. Financials",
  "Enterprise SaaS": "Ent. SaaS",
};

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
  let data;
  try {
    const res = await fetch("data/companies.json");
    if (!res.ok) throw new Error("companies.json not found");
    data = await res.json();
  } catch (e) {
    console.error("companies.json failed to load", e);
    return;
  }
  const meta = data.meta || {};
  const companies = data.companies || [];

  document.getElementById("meta-model-asof").textContent = meta.universe_asof || "—";
  document.getElementById("meta-price-asof").textContent = meta.price_asof || "—";
  const through = document.getElementById("note-through");
  if (through && meta.price_asof) through.textContent = meta.price_asof;

  const buys = companies.filter((c) => c.recommendation === "BUY").length;
  const holds = companies.filter((c) => c.recommendation === "HOLD").length;
  const passes = companies.filter((c) => c.recommendation === "PASS").length;
  document.getElementById("sum-counts").textContent =
    `(live count: ${buys} BUY, ${holds} HOLD, ${passes} PASS at latest prices)`;

  // ---- leaderboard: median upside by sector ----
  const bySector = {};
  companies.forEach((c) => {
    (bySector[c.sector] = bySector[c.sector] || []).push(c);
  });
  const lb = Object.entries(bySector)
    .map(([sector, cs]) => ({
      sector,
      medianUpside: median(cs.map((c) => c.upside_pct)),
      members: [...cs].sort((a, b) => b.upside_pct - a.upside_pct),
    }))
    .sort((a, b) => b.medianUpside - a.medianUpside);

  document.getElementById("lb-body").innerHTML = lb
    .map((r, i) => {
      const members = r.members
        .map((c) => `${c.ticker} ${fmtPct(c.upside_pct)} ${c.recommendation}`)
        .join(" · ");
      return `
      <tr>
        <td><strong>${i + 1}</strong></td>
        <td><strong>${r.sector}</strong></td>
        <td class="num"><span class="upside-chip ${r.medianUpside >= 0 ? "pos" : "neg"}">${fmtPct(r.medianUpside)}</span></td>
        <td>${members}</td>
      </tr>`;
    })
    .join("");

  // ---- full scorecard: ranked by upside ----
  const rows = [...companies].sort((a, b) => b.upside_pct - a.upside_pct);
  document.getElementById("score-body").innerHTML = rows
    .map((c) => `
      <tr>
        <td><span class="ticker">${c.ticker}</span></td>
        <td>${c.name}</td>
        <td>${SHORT_SECTOR[c.sector] || c.sector}</td>
        <td class="num"><span class="price">${fmt$(c.price)}</span></td>
        <td class="num"><span class="dcf">${fmt$(c.dcf_base)}</span></td>
        <td class="num">${fmt$(c.dcf_bull)}</td>
        <td class="num">${fmt$(c.dcf_bear)}</td>
        <td class="num"><span class="upside-chip ${c.upside_pct >= 0 ? "pos" : "neg"}">${fmtPct(c.upside_pct)}</span></td>
        <td><span class="badge ${c.recommendation.toLowerCase()}">${c.recommendation}</span></td>
      </tr>`)
    .join("");
}

document.addEventListener("DOMContentLoaded", init);
