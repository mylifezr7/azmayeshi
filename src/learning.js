const fs = require('node:fs');
const path = require('node:path');
const DATA = path.join(__dirname, '..', 'data', 'learning.json');
function read() { try { return JSON.parse(fs.readFileSync(DATA, 'utf8')); } catch { return []; } }
function record(entry) {
  fs.mkdirSync(path.dirname(DATA), { recursive: true });
  const rows = read();
  rows.push({ strategyId: String(entry.strategyId), returnPct: Number(entry.returnPct) || 0, drawdownPct: Number(entry.drawdownPct) || 0, winRatePct: Number(entry.winRatePct) || 0, trades: Number(entry.trades) || 0, recordedAt: new Date().toISOString() });
  fs.writeFileSync(DATA, JSON.stringify(rows.slice(-250), null, 2));
  return summary();
}
function summary() {
  const rows = read();
  const groups = new Map();
  for (const row of rows) { if (!groups.has(row.strategyId)) groups.set(row.strategyId, []); groups.get(row.strategyId).push(row); }
  return [...groups.entries()].map(([strategyId, all]) => {
    const recent = all.slice(-20);
    return { strategyId, samples: all.length, recentSamples: recent.length, avgReturnPct: recent.reduce((s, x) => s + x.returnPct, 0) / recent.length, avgDrawdownPct: recent.reduce((s, x) => s + x.drawdownPct, 0) / recent.length, avgWinRatePct: recent.reduce((s, x) => s + x.winRatePct, 0) / recent.length, lastUpdated: all.at(-1).recordedAt };
  });
}
module.exports = { record, summary };
