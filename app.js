const storageKey = "portfolio-app-data-v2";

class ManualPriceProvider {
  getCurrentPrice(symbol, state) {
    return state.prices[symbol]?.at(-1)?.price ?? 0;
  }
}

class ApiPriceProvider {
  async getCurrentPrice() {
    throw new Error("未実装: 将来のAPI連携用");
  }
}

const priceProvider = new ManualPriceProvider();
const state = loadState();
hydrateDateDefaults();
attachEvents();
render();

function loadState() {
  const base = { masters: {}, trades: [], alerts: {}, prices: {} };
  try {
    const loaded = JSON.parse(localStorage.getItem(storageKey) || "{}");
    return { ...base, ...loaded };
  } catch {
    return base;
  }
}

function saveState() {
  localStorage.setItem(storageKey, JSON.stringify(state));
}

function hydrateDateDefaults() {
  const today = new Date().toISOString().slice(0, 10);
  for (const id of ["tradeForm", "priceForm"]) {
    const input = document.querySelector(`#${id} input[name=date]`);
    if (input) {
      input.value = today;
    }
  }
}

function attachEvents() {
  document.getElementById("masterForm").addEventListener("submit", onMasterSubmit);
  document.getElementById("tradeForm").addEventListener("submit", onTradeSubmit);
  document.getElementById("priceForm").addEventListener("submit", onPriceSubmit);
  document.getElementById("alertForm").addEventListener("submit", onAlertSubmit);
  document.getElementById("exportCsv").addEventListener("click", exportCsv);
  document.getElementById("importCsv").addEventListener("change", importCsv);
}

function onMasterSubmit(e) {
  e.preventDefault();
  const data = Object.fromEntries(new FormData(e.target).entries());
  const symbol = normalizeSymbol(data.symbol);
  state.masters[symbol] = {
    symbol,
    name: (data.name || "").trim(),
    market: (data.market || "").trim(),
    sector: (data.sector || "").trim(),
  };
  saveState();
  e.target.reset();
  render();
}

function onTradeSubmit(e) {
  e.preventDefault();
  const data = Object.fromEntries(new FormData(e.target).entries());
  state.trades.push({
    symbol: normalizeSymbol(data.symbol),
    broker: (data.broker || "").trim(),
    side: data.side,
    quantity: Number(data.quantity),
    price: Number(data.price),
    fee: Number(data.fee),
    date: data.date,
  });
  saveState();
  e.target.reset();
  hydrateDateDefaults();
  render();
}

function onPriceSubmit(e) {
  e.preventDefault();
  const data = Object.fromEntries(new FormData(e.target).entries());
  const symbol = normalizeSymbol(data.symbol);
  state.prices[symbol] ??= [];
  state.prices[symbol].push({ date: data.date, price: Number(data.price) });
  state.prices[symbol].sort((a, b) => a.date.localeCompare(b.date));
  saveState();
  e.target.reset();
  hydrateDateDefaults();
  render();
}

function onAlertSubmit(e) {
  e.preventDefault();
  const data = Object.fromEntries(new FormData(e.target).entries());
  const symbol = normalizeSymbol(data.symbol);
  state.alerts[symbol] = {
    targetPrice: data.targetPrice ? Number(data.targetPrice) : null,
    stopPrice: data.stopPrice ? Number(data.stopPrice) : null,
    useThreeSigma: Boolean(data.useThreeSigma),
    trailingPercent: data.trailingPercent ? Number(data.trailingPercent) : null,
    highestPrice: state.alerts[symbol]?.highestPrice ?? 0,
    trailingStop: state.alerts[symbol]?.trailingStop ?? null,
  };
  saveState();
  e.target.reset();
  render();
}

function calcPositions() {
  const map = {};
  for (const t of state.trades) {
    const key = `${t.symbol}__${t.broker || "未設定"}`;
    map[key] ??= { symbol: t.symbol, broker: t.broker || "未設定", qty: 0, cost: 0, realized: 0 };
    const p = map[key];
    if (t.side === "BUY") {
      p.qty += t.quantity;
      p.cost += t.quantity * t.price + t.fee;
    } else {
      const avg = p.qty > 0 ? p.cost / p.qty : 0;
      p.realized += t.quantity * (t.price - avg) - t.fee;
      p.qty -= t.quantity;
      p.cost -= avg * t.quantity;
    }
  }

  return Object.values(map).map((p) => {
    const current = priceProvider.getCurrentPrice(p.symbol, state);
    const avg = p.qty > 0 ? p.cost / p.qty : 0;
    const unrealized = p.qty * (current - avg);
    const meta = state.masters[p.symbol] || {};
    return { ...p, avg, current, unrealized, name: meta.name || "-", market: meta.market || "-", sector: meta.sector || "-" };
  });
}

function calcSigma(symbol) {
  const series = state.prices[symbol] || [];
  if (series.length < 20) {
    return null;
  }

  const closes = series.slice(-20).map((entry) => entry.price);
  const mean = closes.reduce((sum, value) => sum + value, 0) / closes.length;
  const variance = closes.reduce((sum, value) => sum + (value - mean) ** 2, 0) / closes.length;
  const std = Math.sqrt(variance);
  return { upper: mean + 3 * std, lower: mean - 3 * std };
}

function calcReports() {
  const monthly = {};
  const yearly = {};

  for (const t of state.trades.filter((trade) => trade.side === "SELL")) {
    const month = t.date.slice(0, 7);
    const year = t.date.slice(0, 4);
    monthly[month] = (monthly[month] || 0) + t.quantity * t.price - t.fee;
    yearly[year] = (yearly[year] || 0) + t.quantity * t.price - t.fee;
  }

  return { monthly, yearly };
}

function evaluateAlerts(positions) {
  const messages = [];

  for (const p of positions) {
    const a = state.alerts[p.symbol];
    if (!a) {
      continue;
    }

    if (a.targetPrice && p.current >= a.targetPrice) {
      messages.push(`${p.symbol}: 目標株価 ${a.targetPrice} 到達`);
    }

    if (a.stopPrice && p.current <= a.stopPrice) {
      messages.push(`${p.symbol}: 損切りライン ${a.stopPrice} 到達`);
    }

    const sigma = a.useThreeSigma ? calcSigma(p.symbol) : null;
    if (sigma && p.current >= sigma.upper) {
      messages.push(`${p.symbol}: 3σ上限 (${sigma.upper.toFixed(2)}) 到達`);
    }

    if (sigma && p.current <= sigma.lower) {
      messages.push(`${p.symbol}: 3σ下限 (${sigma.lower.toFixed(2)}) 到達`);
    }

    if (a.trailingPercent && p.current > 0) {
      a.highestPrice = Math.max(a.highestPrice || 0, p.current);
      a.trailingStop = a.highestPrice * (1 - a.trailingPercent / 100);
      if (p.current <= a.trailingStop) {
        messages.push(`${p.symbol}: トレーリングストップ (${a.trailingStop.toFixed(2)}) 到達`);
      }
    }
  }

  saveState();
  return messages;
}

function render() {
  const positions = calcPositions();
  const totalUnrealized = positions.reduce((sum, p) => sum + p.unrealized, 0);
  const totalRealized = positions.reduce((sum, p) => sum + p.realized, 0);

  document.getElementById("summary").innerHTML = `
    <p>評価損益合計: <span class="${totalUnrealized >= 0 ? "profit" : "loss"}">${totalUnrealized.toFixed(2)}</span></p>
    <p>実現損益合計: <span class="${totalRealized >= 0 ? "profit" : "loss"}">${totalRealized.toFixed(2)}</span></p>
  `;

  const alerts = evaluateAlerts(positions);
  document.getElementById("alerts").innerHTML = alerts.length
    ? alerts.map((line) => `<div class="alert">${line}</div>`).join("")
    : "<p>アラートはありません。</p>";

  document.getElementById("mastersTable").innerHTML = table(
    ["銘柄コード", "銘柄名", "市場", "業種"],
    Object.values(state.masters)
      .sort((a, b) => a.symbol.localeCompare(b.symbol))
      .map((m) => [m.symbol, m.name || "-", m.market || "-", m.sector || "-"])
  );

  document.getElementById("positionsTable").innerHTML = table(
    ["銘柄", "証券会社", "銘柄名", "市場", "業種", "保有株数", "平均取得", "現在値", "評価損益", "実現損益"],
    positions.map((p) => [
      p.symbol,
      p.broker,
      p.name,
      p.market,
      p.sector,
      p.qty,
      p.avg.toFixed(2),
      p.current.toFixed(2),
      classed(p.unrealized),
      classed(p.realized),
    ])
  );

  document.getElementById("tradesTable").innerHTML = table(
    ["日付", "銘柄", "証券会社", "銘柄名", "区分", "数量", "単価", "手数料"],
    [...state.trades]
      .sort((a, b) => b.date.localeCompare(a.date))
      .map((t) => [
        t.date,
        t.symbol,
        t.broker || "未設定",
        state.masters[t.symbol]?.name || "-",
        t.side === "BUY" ? "買付" : "売却",
        t.quantity,
        t.price,
        t.fee,
      ])
  );

  const reports = calcReports();
  const rows = [
    ...Object.entries(reports.monthly).map(([period, value]) => ["月次", period, classed(value)]),
    ...Object.entries(reports.yearly).map(([period, value]) => ["年次", period, classed(value)]),
  ];
  document.getElementById("reportTable").innerHTML = table(["区分", "期間", "損益"], rows);
}

function exportCsv() {
  const rows = [
    ["type", "symbol", "broker", "name", "market", "sector", "side", "quantity", "price", "fee", "date", "targetPrice", "stopPrice", "useThreeSigma", "trailingPercent", "highestPrice", "trailingStop"],
  ];

  for (const m of Object.values(state.masters)) {
    rows.push(["MASTER", m.symbol, "", m.name || "", m.market || "", m.sector || "", "", "", "", "", "", "", "", "", "", "", ""]);
  }

  for (const t of state.trades) {
    rows.push(["TRADE", t.symbol, t.broker || "", "", "", "", t.side, t.quantity, t.price, t.fee, t.date, "", "", "", "", "", ""]);
  }

  for (const [symbol, series] of Object.entries(state.prices)) {
    for (const item of series) {
      rows.push(["PRICE", symbol, "", "", "", "", "", "", item.price, "", item.date, "", "", "", "", "", ""]);
    }
  }

  for (const [symbol, a] of Object.entries(state.alerts)) {
    rows.push([
      "ALERT",
      symbol,
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      a.targetPrice ?? "",
      a.stopPrice ?? "",
      a.useThreeSigma ? "1" : "0",
      a.trailingPercent ?? "",
      a.highestPrice ?? "",
      a.trailingStop ?? "",
    ]);
  }

  const csv = rows.map((row) => row.map(csvEscape).join(",")).join("\n");
  downloadBlob(csv, `portfolio-export-${new Date().toISOString().slice(0, 10)}.csv`, "text/csv;charset=utf-8");
}

function importCsv(e) {
  const [file] = e.target.files || [];
  if (!file) {
    return;
  }

  const reader = new FileReader();
  reader.onload = () => {
    const parsed = parseCsv(String(reader.result || ""));
    if (!parsed.length) {
      return;
    }

    const [header, ...records] = parsed;
    const index = Object.fromEntries(header.map((h, i) => [h, i]));
    const next = { masters: {}, trades: [], alerts: {}, prices: {} };

    for (const record of records) {
      const type = cell(record, index, "type");
      const symbol = normalizeSymbol(cell(record, index, "symbol"));
      if (!type || !symbol) {
        continue;
      }

      if (type === "MASTER") {
        next.masters[symbol] = {
          symbol,
          name: cell(record, index, "name"),
          market: cell(record, index, "market"),
          sector: cell(record, index, "sector"),
        };
      }

      if (type === "TRADE") {
        next.trades.push({
          symbol,
          broker: cell(record, index, "broker") || "未設定",
          side: cell(record, index, "side") || "BUY",
          quantity: Number(cell(record, index, "quantity") || 0),
          price: Number(cell(record, index, "price") || 0),
          fee: Number(cell(record, index, "fee") || 0),
          date: cell(record, index, "date"),
        });
      }

      if (type === "PRICE") {
        next.prices[symbol] ??= [];
        next.prices[symbol].push({
          date: cell(record, index, "date"),
          price: Number(cell(record, index, "price") || 0),
        });
      }

      if (type === "ALERT") {
        next.alerts[symbol] = {
          targetPrice: toNullableNumber(cell(record, index, "targetPrice")),
          stopPrice: toNullableNumber(cell(record, index, "stopPrice")),
          useThreeSigma: cell(record, index, "useThreeSigma") === "1",
          trailingPercent: toNullableNumber(cell(record, index, "trailingPercent")),
          highestPrice: toNullableNumber(cell(record, index, "highestPrice")) || 0,
          trailingStop: toNullableNumber(cell(record, index, "trailingStop")),
        };
      }
    }

    state.masters = next.masters;
    state.trades = next.trades;
    state.alerts = next.alerts;
    state.prices = next.prices;
    saveState();
    render();
    e.target.value = "";
  };
  reader.readAsText(file, "utf-8");
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cellValue = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        cellValue += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === "," && !inQuotes) {
      row.push(cellValue);
      cellValue = "";
      continue;
    }

    if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") {
        i += 1;
      }
      row.push(cellValue);
      if (row.some((value) => value !== "")) {
        rows.push(row);
      }
      row = [];
      cellValue = "";
      continue;
    }

    cellValue += char;
  }

  if (cellValue !== "" || row.length) {
    row.push(cellValue);
    if (row.some((value) => value !== "")) {
      rows.push(row);
    }
  }

  return rows;
}

function csvEscape(value) {
  const raw = String(value ?? "");
  if (/[",\n]/.test(raw)) {
    return `"${raw.replace(/"/g, '""')}"`;
  }
  return raw;
}

function downloadBlob(text, filename, mime) {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function cell(record, index, key) {
  const position = index[key];
  return position === undefined ? "" : (record[position] || "").trim();
}

function toNullableNumber(raw) {
  return raw === "" ? null : Number(raw);
}

function normalizeSymbol(symbol) {
  return String(symbol || "").trim().toUpperCase();
}

function classed(value) {
  return `<span class="${value >= 0 ? "profit" : "loss"}">${Number(value).toFixed(2)}</span>`;
}

function table(headers, rows) {
  return `
    <thead><tr>${headers.map((h) => `<th>${h}</th>`).join("")}</tr></thead>
    <tbody>
      ${rows.length ? rows.map((r) => `<tr>${r.map((c) => `<td>${c}</td>`).join("")}</tr>`).join("") : `<tr><td colspan="${headers.length}">データなし</td></tr>`}
    </tbody>
  `;
}
