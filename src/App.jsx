import React, { useState, useMemo, useRef, useEffect } from 'react';
import {
  BarChart3, LayoutDashboard, Layers, TrendingUp, TrendingDown, SlidersHorizontal, Table2, Search,
  ArrowUpDown, ChevronLeft, ChevronRight, RotateCcw, Upload, X, Plus, Settings2,
  FileSpreadsheet, ChevronDown, ChevronUp, AlertCircle, Sparkles, CheckCircle2,
  AlertTriangle, Target, Loader2, Download, FileDown, Printer, Save, BookmarkCheck,
  GitCompare, Minus, Zap, UploadCloud, Grid3x3, MessageCircleQuestion, Send, FileUp
} from 'lucide-react';
import {
  ResponsiveContainer, AreaChart, Area, BarChart, Bar, XAxis, YAxis,
  CartesianGrid, Tooltip, PieChart, Pie, Cell, RadialBarChart, RadialBar, PolarAngleAxis, Brush
} from 'recharts';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';

/* ===================== Design tokens — "Ledger" theme ===================== */
const COLORS = {
  ink: '#0F1424',
  inkLight: '#1B2340',
  inkSoft: '#2A3356',
  accent: '#5850EC',
  accentSoft: '#EDECFE',
  cyan: '#06B6D4',
  emerald: '#047857',
  rose: '#F43F5E',
  amber: '#F59E0B',
  amberText: '#B45309',
  violet: '#8B5CF6',
  muted: '#64748B',
  border: '#E4E7EE',
  bg: '#F6F7FB',
  panel: '#FFFFFF',
};
const CHART_PALETTE = [COLORS.accent, COLORS.cyan, COLORS.emerald, COLORS.amber, COLORS.rose, COLORS.violet];

const MONO = 'ui-monospace, SFMono-Regular, "Roboto Mono", Menlo, Consolas, monospace';
const SANS = 'ui-sans-serif, -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif';
const tooltipStyle = { fontFamily: SANS, fontSize: 12, borderRadius: 10, border: `1px solid ${COLORS.border}`, boxShadow: '0 8px 24px rgba(15,20,36,0.10)' };

function fmtUSD(v, decimals) {
  const d = decimals === undefined ? 0 : decimals;
  if (v === null || v === undefined || isNaN(v)) return '$0';
  const sign = v < 0 ? '-' : '';
  return sign + '$' + Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: d, maximumFractionDigits: d });
}
function fmtPct(v, decimals) {
  const d = decimals === undefined ? 1 : decimals;
  if (v === null || v === undefined || isNaN(v)) return '0.0%';
  return (v * 100).toFixed(d) + '%';
}
function fmtNum(v, decimals) {
  const d = decimals === undefined ? 4 : decimals;
  if (v === null || v === undefined || isNaN(v)) return '0';
  return v.toFixed(d);
}
function fmtByFormat(v, format) {
  if (v === null || v === undefined || isNaN(v)) v = 0;
  if (format === 'currency') return fmtUSD(v, Math.abs(v) < 10 ? 2 : 0);
  if (format === 'percent') return fmtPct(v, 1);
  return Number(v).toLocaleString('en-US', { maximumFractionDigits: Math.abs(v) < 10 ? 2 : 0 });
}
// Light heuristic so obviously-monetary columns default to $ display instead of a
// plain number — purely a starting guess, always changeable per-metric in Configure analysis.
function guessMetricFormat(key) {
  const k = String(key).toLowerCase();
  if (/(^|[^a-z])(pct|percent|margin\s*%|rate)([^a-z]|$)/.test(k) || k.includes('%')) return 'percent';
  if (/revenue|cost|profit|price|budget|sales|income|expense|amount|spend|salary|payout|fee\b/.test(k)) return 'currency';
  return 'number';
}
// Finds the first numeric column matching an exact (case-insensitive) name, falling
// back to a substring match — used to auto-detect Revenue/Cost-style column pairs.
function findColumnByName(cols, exactNames, substrings) {
  const exact = cols.find(c => exactNames.includes(String(c.key).toLowerCase()));
  if (exact) return exact;
  return cols.find(c => substrings.some(s => String(c.key).toLowerCase().includes(s)));
}
const MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
function monthLabel(m) {
  const [y, mo] = m.split('-');
  return MONTH_NAMES[parseInt(mo, 10) - 1] + " '" + y.slice(2);
}
// KPI cards show a grand total across the (filtered) dataset, so label them as such
// ("Total Revenue") unless the column name already says so.
function kpiLabel(key) {
  return /^total\b/i.test(key) ? key : 'Total ' + key;
}

/* ===================== Generic CSV analysis engine ===================== */

// Strict-ish date pattern gate: Date.parse() alone is too lenient (e.g. it happily
// parses "ORD-1" or "TXN-1000" as a valid date), so we only trust it on strings that
// already look like a real date.
const DATE_LIKE_PATTERN = /^\d{4}-\d{1,2}-\d{1,2}([ T]\d{1,2}:\d{2}(:\d{2})?(\.\d+)?(Z|[+-]\d{2}:?\d{2})?)?$|^\d{1,2}\/\d{1,2}\/\d{2,4}$|^\d{1,2}-\d{1,2}-\d{2,4}$|^\d{4}\/\d{1,2}\/\d{1,2}$|^\d{1,2}\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{2,4}$/i;

export function parseNumericValue(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  if (typeof value !== 'string') return null;
  let text = value.trim();
  if (!text) return null;
  const negative = /^\(.*\)$/.test(text);
  if (negative) text = text.slice(1, -1).trim();
  const percent = text.endsWith('%');
  text = text.replace(/^[\$£€₹]\s*/, '').replace(/\s*[\$£€₹%]$/, '');
  if (!/^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?$/.test(text)) return null;
  const parsed = Number(text.replace(/,/g, ''));
  return Number.isFinite(parsed) ? (negative ? -parsed : parsed) / (percent ? 100 : 1) : null;
}

export function parseDateValue(value, format = 'mdy') {
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value.toISOString().slice(0, 10);
  if (typeof value !== 'string' || !DATE_LIKE_PATTERN.test(value.trim())) return null;
  const text = value.trim();
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:$|[ T])/.exec(text);
  if (iso) {
    const [, year, month, day] = iso.map(Number);
    const check = new Date(Date.UTC(year, month - 1, day));
    if (check.getUTCFullYear() !== year || check.getUTCMonth() + 1 !== month || check.getUTCDate() !== day) return null;
    return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }
  const yearFirst = /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/.exec(text);
  const separated = /^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/.exec(text);
  if (yearFirst || separated) {
    let year, month, day;
    if (yearFirst) [, year, month, day] = yearFirst.map(Number);
    else {
      const first = Number(separated[1]), second = Number(separated[2]);
      year = Number(separated[3]);
      if (year < 100) year += 2000;
      month = format === 'dmy' ? second : first;
      day = format === 'dmy' ? first : second;
    }
    const check = new Date(Date.UTC(year, month - 1, day));
    if (check.getUTCFullYear() !== year || check.getUTCMonth() + 1 !== month || check.getUTCDate() !== day) return null;
    return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed.toISOString().slice(0, 10);
}

export function normalizeRows(rows, columnTypes) {
  return rows.map(row => Object.fromEntries(Object.entries(row).map(([key, value]) => {
    const type = columnTypes[key];
    if (type === 'number') return [key, parseNumericValue(value)];
    if (type === 'date' || type === 'date-dmy') return [key, parseDateValue(value, type === 'date-dmy' ? 'dmy' : 'mdy')];
    return [key, value];
  })));
}

// Classify each column as number / date / category / text(high-cardinality, excluded by default)
export function detectSchema(rows, fields) {
  return fields.map(f => {
    const raw = rows.map(r => r[f]).filter(v => v !== null && v !== undefined && v !== '');
    const total = raw.length || 1;
    const numericCount = raw.filter(v => parseNumericValue(v) !== null).length;
    const dateCount = raw.filter(v => {
      return parseDateValue(v) !== null;
    }).length;
    const uniqueVals = new Set(raw.map(v => String(v)));
    let type = 'category';
    if (numericCount / total >= 0.85) type = 'number';
    else if (dateCount / total >= 0.7) type = 'date';
    else if (uniqueVals.size > Math.min(40, Math.max(15, total * 0.6))) type = 'text';
    return { key: f, type, unique: uniqueVals.size, total: raw.length, sample: raw.slice(0, 4) };
  });
}

export function aggregateMetrics(rows, metricKeys) {
  const out = { count: rows.length };
  metricKeys.forEach(metric => {
    const k = typeof metric === 'string' ? metric : metric.key;
    const aggregation = typeof metric === 'string' ? 'sum' : (metric.aggregation || 'sum');
    const vals = rows.map(r => parseNumericValue(r[k])).filter(v => v !== null);
    const sum = vals.reduce((a, b) => a + b, 0);
    out[k] = aggregation === 'count' ? vals.length :
      aggregation === 'average' ? (vals.length ? sum / vals.length : 0) :
      aggregation === 'min' ? (vals.length ? vals.reduce((a, b) => Math.min(a, b)) : 0) :
      aggregation === 'max' ? (vals.length ? vals.reduce((a, b) => Math.max(a, b)) : 0) : sum;
  });
  return out;
}

// Generalized two-operand calculated metric: sum(A) OP sum(B). Division keeps the
// original "ratio" behavior (aggregate-then-divide); +/- are exact either way since
// sum(a)-sum(b) === sum(a-b). Field names (num/den) are kept from the original
// division-only "ratio" concept for backward compatibility with saved configs.
// Applies a calculated-metric operator to two already-aggregated numbers —
// the scalar half of computeCalc, reused wherever two projected/aggregated
// values (not raw rows) need combining the same way a custom metric is
// combined (e.g. the sensitivity grid, which projects num/den independently
// then recombines them per period).
export function applyCalcOp(numVal, denVal, op) {
  const n = Number(numVal) || 0, d = Number(denVal) || 0;
  if (op === '+') return n + d;
  if (op === '−' || op === '-') return n - d;
  return d ? n / d : 0;
}
export function computeCalc(rows, calc) {
  const numSum = rows.reduce((s, r) => s + (Number(r[calc.num]) || 0), 0);
  const denSum = rows.reduce((s, r) => s + (Number(r[calc.den]) || 0), 0);
  return applyCalcOp(numSum, denSum, calc.op || '÷');
}
// Backward-compatible alias.
export const computeRatio = computeCalc;

export function groupByGeneric(rows, dimKey, metricKeys, ratios) {
  const groups = new Map();
  rows.forEach(r => {
    const raw = r[dimKey];
    const k = (raw === null || raw === undefined || raw === '') ? '(blank)' : String(raw);
    if (!groups.has(k)) groups.set(k, []);
    groups.get(k).push(r);
  });
  return Array.from(groups.keys()).sort().map(k => {
    const subset = groups.get(k);
    const agg = aggregateMetrics(subset, metricKeys);
    const ratioVals = {};
    (ratios || []).forEach(rt => { ratioVals[rt.key] = computeRatio(subset, rt); });
    return { name: k, ...agg, ...ratioVals };
  });
}

export function monthlyGroupGeneric(rows, dateKey, metricKeys, ratios) {
  const withMonth = rows.map(r => {
    const date = parseDateValue(r[dateKey]);
    if (!date) return null;
    const m = date.slice(0, 7);
    return { ...r, __month: m };
  }).filter(Boolean);
  const months = Array.from(new Set(withMonth.map(r => r.__month))).sort();
  return months.map(m => {
    const subset = withMonth.filter(r => r.__month === m);
    const agg = aggregateMetrics(subset, metricKeys);
    const ratioVals = {};
    (ratios || []).forEach(rt => { ratioVals[rt.key] = computeRatio(subset, rt); });
    return { month: m, label: monthLabel(m.length === 7 ? m : m.slice(0, 7)), ...agg, ...ratioVals };
  });
}

export function computeAvgGrowth(trend, key) {
  if (!trend || trend.length < 2) return 0;
  const rates = [];
  for (let i = 1; i < trend.length; i++) {
    const prev = trend[i - 1][key], cur = trend[i][key];
    const previousMonth = trend[i - 1].month, currentMonth = trend[i].month;
    const months = previousMonth && currentMonth ?
      (Number(currentMonth.slice(0, 4)) - Number(previousMonth.slice(0, 4))) * 12 +
      Number(currentMonth.slice(5, 7)) - Number(previousMonth.slice(5, 7)) : 1;
    if (prev > 0 && cur > 0 && months > 0) rates.push(Math.pow(cur / prev, 1 / months) - 1);
  }
  if (!rates.length) return 0;
  const avg = rates.reduce((a, b) => a + b, 0) / rates.length;
  return Math.max(-0.4, Math.min(0.4, avg));
}

export function projectGeneric(baseValue, monthlyGrowth, labels) {
  let v = baseValue;
  return labels.map(label => {
    v = v * (1 + monthlyGrowth);
    return { label, value: v };
  });
}

export function backtestForecast(trend, key) {
  const errors = [];
  for (let i = 3; i < trend.length; i++) {
    const current = trend[i], previous = trend[i - 1];
    const expectedMonth = new Date(Date.UTC(Number(previous.month.slice(0, 4)), Number(previous.month.slice(5, 7)), 1)).toISOString().slice(0, 7);
    if (current.month !== expectedMonth || current[key] <= 0) continue;
    const history = trend.slice(Math.max(0, i - 6), i);
    const projected = previous[key] * (1 + computeAvgGrowth(history, key));
    errors.push(Math.abs(projected - current[key]) / Math.abs(current[key]));
  }
  return errors.length ? { samples: errors.length, meanAbsolutePercentError: errors.reduce((a, b) => a + b, 0) / errors.length } : null;
}

/* ===================== Rule-based insight engine (instant, no API) ===================== */
export function computeInsights(input) {
  const out = [];
  const fmt = (v) => fmtByFormat(v, input.metricFormat);

  (input.breakdowns || []).slice(0, 3).forEach(bd => {
    const groups = bd.groups.filter(g => g.count > 0);
    if (groups.length < 2) return;
    const sorted = [...groups].sort((a, b) => b.value - a.value);
    const top = sorted[0], bottom = sorted[sorted.length - 1];
    const total = groups.reduce((s, g) => s + g.value, 0);
    const topShare = total ? top.value / total : 0;

    out.push({
      type: 'positive',
      title: bd.dimLabel + ': "' + top.name + '" leads',
      detail: top.name + ' delivers ' + fmt(top.value) + ' in ' + input.metricLabel + (total ? ' (' + fmtPct(topShare, 0) + ' of the total across ' + bd.dimLabel.toLowerCase() + ')' : '') + '.',
    });

    if (bottom.name !== top.name) {
      const gap = top.value !== 0 ? (top.value - bottom.value) / Math.abs(top.value) : 0;
      out.push({
        type: 'concern',
        title: bd.dimLabel + ': "' + bottom.name + '" trails the field',
        detail: bottom.name + ' produced ' + fmt(bottom.value) + ', ' + fmtPct(gap, 0) + ' behind the leader (' + top.name + '). Closing even half that gap would meaningfully lift the total.',
      });
    }

    if (groups.length >= 4 && topShare >= 0.4) {
      out.push({
        type: 'opportunity',
        title: bd.dimLabel + ': concentration risk',
        detail: '"' + top.name + '" alone accounts for ' + fmtPct(topShare, 0) + ' of total ' + input.metricLabel + ' across ' + groups.length + ' ' + bd.dimLabel.toLowerCase() + ' segments — diversifying reduces exposure if that segment softens.',
      });
    }
  });

  if (input.trend && input.trend.length >= 3) {
    const first = input.trend[0], last = input.trend[input.trend.length - 1];
    const overallChange = first.value !== 0 ? (last.value - first.value) / Math.abs(first.value) : 0;
    if (overallChange > 0.02) {
      out.push({
        type: 'positive',
        title: input.metricLabel + ' is trending up',
        detail: input.metricLabel + ' moved from ' + fmt(first.value) + ' (' + first.label + ') to ' + fmt(last.value) + ' (' + last.label + '), up ' + fmtPct(overallChange, 0) + ' over the period.',
      });
    } else if (overallChange < -0.02) {
      out.push({
        type: 'concern',
        title: input.metricLabel + ' is trending down',
        detail: input.metricLabel + ' fell from ' + fmt(first.value) + ' (' + first.label + ') to ' + fmt(last.value) + ' (' + last.label + '), down ' + fmtPct(Math.abs(overallChange), 0) + ' — worth investigating the cause before it compounds.',
      });
    }
    let declines = 0;
    for (let i = 1; i < input.trend.length; i++) if (input.trend[i].value < input.trend[i - 1].value) declines++;
    if (declines >= Math.ceil((input.trend.length - 1) * 0.5)) {
      out.push({
        type: 'concern',
        title: 'Inconsistent period-to-period performance',
        detail: input.metricLabel + ' declined in ' + declines + ' of the last ' + (input.trend.length - 1) + ' periods — a choppy pattern that makes forecasting less reliable.',
      });
    }
  }

  if (input.ratio && input.ratio.groups && input.ratio.groups.length >= 2) {
    const rfmt = (v) => fmtByFormat(v, input.ratio.format);
    const sorted = [...input.ratio.groups].sort((a, b) => b.value - a.value);
    const best = sorted[0], worst = sorted[sorted.length - 1];
    const avg = sorted.reduce((s, g) => s + g.value, 0) / sorted.length;
    if (best.name !== worst.name) {
      out.push({
        type: 'opportunity',
        title: input.ratio.label + ': biggest improvement lever',
        detail: '"' + worst.name + '" has the weakest ' + input.ratio.label + ' at ' + rfmt(worst.value) + ', versus ' + rfmt(best.value) + ' for "' + best.name + '". Bringing it to the portfolio average (' + rfmt(avg) + ') is the single highest-leverage fix available.',
      });
    }
  }

  return out;
}

/* ===================== Export helpers ===================== */
export function csvEscape(v) {
  if (v === null || v === undefined) return '';
  // Spreadsheet applications can evaluate formulas in imported CSV cells.
  const raw = String(v);
  const s = /^[\s\u0000-\u001f]*[=+@\-]/.test(raw) && typeof v === 'string' ? "'" + raw : raw;
  if (/[",\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}
export function rowsToCSV(fields, rows) {
  const header = fields.map(csvEscape).join(',');
  const lines = rows.map(r => fields.map(f => csvEscape(r[f])).join(','));
  return [header, ...lines].join('\n');
}
function legacyAnchorDownload(filename, content, mime) {
  const blob = new Blob([content], { type: (mime || 'text/plain') + ';charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
export async function downloadTextFile(filename, content, mime) {
  // If this component is ever published as a Cowork/claude.ai Artifact, its
  // frame is sandboxed from normal browser downloads and a plain <a download>
  // link does nothing there. When the platform grants this runtime
  // capability, hand the file to the viewer through it instead (they see a
  // save confirmation); otherwise fall back to the ordinary browser download.
  if (typeof window !== 'undefined' && window.claude && typeof window.claude.use === 'function') {
    try {
      const downloads = await window.claude.use('downloads');
      if (downloads) {
        try {
          await downloads.save({ filename, data: content });
          return;
        } catch (err) {
          if (err && (err.code === 'rejected_extension' || err.code === 'extension_not_enabled')) {
            const base = filename.replace(/\.[a-zA-Z0-9]+$/, '');
            try { await downloads.save({ filename: base + '.txt', data: content }); } catch (e2) { /* viewer declined or unavailable */ }
          }
          return;
        }
      }
    } catch (e) { /* capability not available — fall back below */ }
  }
  legacyAnchorDownload(filename, content, mime);
}

/* ===================== Correlation & outliers ===================== */
export function pearson(xs, ys) {
  const n = xs.length;
  if (n < 2) return 0;
  const mx = xs.reduce((a, b) => a + b, 0) / n, my = ys.reduce((a, b) => a + b, 0) / n;
  let num = 0, dx2 = 0, dy2 = 0;
  for (let i = 0; i < n; i++) { const dx = xs[i] - mx, dy = ys[i] - my; num += dx * dy; dx2 += dx * dx; dy2 += dy * dy; }
  const den = Math.sqrt(dx2 * dy2);
  return den ? num / den : 0;
}
export function computeCorrelations(rows, keys) {
  const out = [];
  for (let i = 0; i < keys.length; i++) {
    for (let j = i + 1; j < keys.length; j++) {
      const pairs = rows.map(r => [parseNumericValue(r[keys[i]]), parseNumericValue(r[keys[j]])]).filter(p => p[0] !== null && p[1] !== null);
      if (pairs.length < 5) continue;
      out.push({ a: keys[i], b: keys[j], r: pearson(pairs.map(p => p[0]), pairs.map(p => p[1])), n: pairs.length });
    }
  }
  return out.sort((x, y) => Math.abs(y.r) - Math.abs(x.r));
}
export function corrStrength(r) {
  const a = Math.abs(r);
  if (a >= 0.7) return 'strong';
  if (a >= 0.4) return 'moderate';
  if (a >= 0.2) return 'weak';
  return 'negligible';
}
export function computeOutliers(rows, key, zThresh) {
  const zt = zThresh || 2.5;
  const vals = rows.map(r => parseNumericValue(r[key])).filter(v => v !== null);
  const n = vals.length;
  if (n < 6) return [];
  const mean = vals.reduce((a, b) => a + b, 0) / n;
  const variance = vals.reduce((s, v) => s + (v - mean) * (v - mean), 0) / n;
  const sd = Math.sqrt(variance);
  if (!sd) return [];
  const out = [];
  rows.forEach((r) => {
    const v = parseNumericValue(r[key]);
    if (v === null) return;
    const z = (v - mean) / sd;
    if (Math.abs(z) >= zt) out.push({ value: v, z, row: r });
  });
  return out.sort((a, b) => Math.abs(b.z) - Math.abs(a.z));
}

/* ===================== Period-over-period ===================== */
// offset=1 compares the latest period to the one immediately before it ("prior
// period"); offset=12 compares it to the same calendar month a year back
// ("year over year") — both use the same monthly trend array, just indexed
// further back, so one function covers both modes.
export function periodDelta(trend, key, offset = 1) {
  if (!trend || trend.length < 2) return null;
  const cur = trend[trend.length - 1];
  const [year, month] = cur.month.split('-').map(Number);
  const target = new Date(Date.UTC(year, month - 1 - offset, 1)).toISOString().slice(0, 7);
  const prev = trend.find(item => item.month === target);
  if (!prev) return null;
  const pct = prev[key] !== 0 ? (cur[key] - prev[key]) / Math.abs(prev[key]) : 0;
  return { curLabel: cur.label, prevLabel: prev.label, cur: cur[key], prev: prev[key], pct };
}

/* ===================== Saved-setup persistence (window.storage) ===================== */
export function schemaSignature(fields) {
  return JSON.stringify(fields.slice().sort());
}
function legacySchemaSignature(fields) {
  return fields.slice().sort().join('|').toLowerCase().slice(0, 150);
}
export async function loadSavedConfig(fields) {
  if (typeof window === 'undefined' || !window.storage) return null;
  try {
    const res = await window.storage.get('mbu_config:' + schemaSignature(fields), false);
    if (res && res.value) return JSON.parse(res.value);
    const old = await window.storage.get('mbu_config:' + legacySchemaSignature(fields), false);
    if (old && old.value) {
      const config = JSON.parse(old.value);
      const savedFields = new Set([
        ...(config.metrics || []).map(metric => metric.key),
        ...(config.dimensions || []).map(dimension => dimension.key),
        ...(config.dateField ? [config.dateField] : []),
      ]);
      if (savedFields.size === fields.length && fields.every(field => savedFields.has(field))) {
        await saveConfig(fields, config);
        return config;
      }
    }
  } catch (e) { /* no saved config for this shape yet, or storage unavailable */ }
  return null;
}
export async function saveConfig(fields, config) {
  if (typeof window === 'undefined' || !window.storage) throw new Error('Storage not available in this environment.');
  await window.storage.set('mbu_config:' + schemaSignature(fields), JSON.stringify(config), false);
}
export async function deleteSavedConfig(fields) {
  if (typeof window === 'undefined' || !window.storage) return;
  try { await window.storage.delete('mbu_config:' + schemaSignature(fields), false); } catch (e) { /* nothing to delete */ }
  try { await window.storage.delete('mbu_config:' + legacySchemaSignature(fields), false); } catch (e) { /* nothing to delete */ }
}

/* ===================== Saved views (named filter/breakdown presets) ===================== */
export async function loadSavedViews(fields) {
  if (typeof window === 'undefined' || !window.storage) return [];
  try {
    const res = await window.storage.get('mbu_views:' + schemaSignature(fields), false);
    if (res && res.value) return JSON.parse(res.value);
    const old = await window.storage.get('mbu_views:' + legacySchemaSignature(fields), false);
    if (old && old.value) {
      const views = JSON.parse(old.value);
      if (Array.isArray(views)) {
        await saveViewsList(fields, views);
        return views;
      }
    }
  } catch (e) { /* none saved yet, or storage unavailable */ }
  return [];
}
export async function saveViewsList(fields, views) {
  if (typeof window === 'undefined' || !window.storage) throw new Error('Storage not available in this environment.');
  await window.storage.set('mbu_views:' + schemaSignature(fields), JSON.stringify(views), false);
}

/* ===================== Budget vs Actual ===================== */
export function computeVariance(actualVal, budgetVal) {
  const a = Number(actualVal) || 0, b = Number(budgetVal) || 0;
  const variance = a - b;
  const pct = b !== 0 ? variance / Math.abs(b) : 0;
  return { variance, pct };
}

/* ===================== Data quality diagnostics ===================== */
export function computeDataQuality(rows, fields, dateField) {
  const total = rows.length;
  const nullStats = fields.map(f => {
    const nulls = rows.filter(r => r[f] === null || r[f] === undefined || r[f] === '').length;
    return { field: f, nulls, pct: total ? nulls / total : 0 };
  }).filter(s => s.nulls > 0).sort((a, b) => b.nulls - a.nulls);

  const seen = new Set();
  let duplicates = 0;
  rows.forEach(r => {
    const key = fields.map(f => String(r[f])).join('|');
    if (seen.has(key)) duplicates++;
    else seen.add(key);
  });

  let badDateCount = 0;
  if (dateField) {
    badDateCount = rows.filter(r => {
      const v = r[dateField];
      if (v === null || v === undefined || v === '') return false;
      return parseDateValue(v) === null;
    }).length;
  }

  return { total, nullStats, duplicates, badDateCount };
}
export function computeNegativeCounts(rows, metricKeys) {
  return metricKeys
    .map(k => ({ key: k, count: rows.filter(r => typeof r[k] === 'number' && r[k] < 0).length }))
    .filter(x => x.count > 0);
}

function NavItem({ icon: Icon, label, active, onClick, vertical }) {
  const base = {
    display: 'flex', alignItems: 'center', gap: 10, fontFamily: SANS, fontSize: 13.5,
    fontWeight: active ? 600 : 500, color: active ? '#FFFFFF' : '#9098B8',
    background: active ? COLORS.inkLight : 'transparent', border: 'none', cursor: 'pointer', whiteSpace: 'nowrap',
  };
  const vStyle = { ...base, width: '100%', padding: '10px 14px', borderRadius: 8, textAlign: 'left' };
  const hStyle = { ...base, padding: '9px 14px', borderRadius: 999, flexShrink: 0 };
  return (
    <button onClick={onClick} style={vertical ? vStyle : hStyle}>
      <Icon size={16} strokeWidth={2} />
      <span>{label}</span>
    </button>
  );
}
// A Power BI-style "slicer": a labeled group of clickable value pills. Multiple
// values can be selected at once (OR'd together within the dimension); no selection
// means "All". Click a pill again to deselect it.
function Slicer({ label, options, selected, onChange, highlight }) {
  const sel = selected || [];
  const isAll = sel.length === 0;
  const shown = options.slice(0, 40);
  const hiddenCount = options.length - shown.length;
  function toggle(opt) {
    onChange(sel.includes(opt) ? sel.filter(v => v !== opt) : [...sel, opt]);
  }
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 7, maxWidth: 360 }}>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
        <span style={{ fontFamily: SANS, fontSize: 11, fontWeight: 700, color: COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.04em' }}>{label}</span>
        {!isAll ? (
          <button onClick={() => onChange([])} style={{ fontFamily: SANS, fontSize: 10.5, fontWeight: 700, color: COLORS.accent, background: 'transparent', border: 'none', cursor: 'pointer', padding: 0 }}>
            Clear ({sel.length})
          </button>
        ) : null}
      </div>
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
        {shown.map(opt => {
          const active = sel.includes(opt);
          return (
            <button key={opt} onClick={() => toggle(opt)} aria-pressed={active}
              style={{
                fontFamily: SANS, fontSize: 11.5, fontWeight: 600, padding: '5px 11px', borderRadius: 999, cursor: 'pointer',
                border: `1px solid ${active ? COLORS.accent : (highlight === opt ? COLORS.ink : COLORS.border)}`,
                background: active ? COLORS.accent : '#FFFFFF', color: active ? '#FFFFFF' : COLORS.ink,
                boxShadow: highlight === opt && !active ? `0 0 0 1px ${COLORS.ink}` : 'none',
              }}>
              {opt}
            </button>
          );
        })}
        {hiddenCount > 0 ? <span style={{ fontFamily: SANS, fontSize: 11, color: COLORS.muted, alignSelf: 'center' }}>+{hiddenCount} more — search in the Data tab</span> : null}
      </div>
    </div>
  );
}
function Select({ value, onChange, options, label, labelize }) {
  return (
    <label style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
      <span style={{ fontFamily: SANS, fontSize: 11, fontWeight: 600, color: COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.04em' }}>{label}</span>
      <select value={value} onChange={e => onChange(e.target.value)}
        style={{ fontFamily: SANS, fontSize: 13.5, color: COLORS.ink, background: '#FFFFFF', border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: '8px 10px', minWidth: 160, cursor: 'pointer' }}>
        {options.map(o => <option key={o} value={o}>{labelize ? labelize(o) : o}</option>)}
      </select>
    </label>
  );
}
// Shared rich tooltip for line/area/bar charts: crosshair-following, lists every
// active series at that point, value leads (bold, large) with the series name
// secondary, and a short line-key (not a filled box) per dataviz convention.
function RichTooltip({ active, payload, label, formatMap, defaultFormat }) {
  if (!active || !payload || !payload.length) return null;
  const fmt = (p) => fmtByFormat(p.value, (formatMap && formatMap[p.dataKey]) || defaultFormat || 'number');
  return (
    <div style={{ ...tooltipStyle, background: '#FFFFFF', padding: '10px 13px', minWidth: 136 }}>
      {label !== undefined && label !== null && label !== '' ? (
        <div style={{ fontFamily: SANS, fontSize: 11, fontWeight: 700, color: COLORS.ink, marginBottom: 7 }}>{label}</div>
      ) : null}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
        {payload.map((p, i) => (
          <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 16, justifyContent: 'space-between' }}>
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontFamily: SANS, fontSize: 11.5, color: COLORS.muted }}>
              <span style={{ width: 11, height: 2.5, borderRadius: 2, background: p.color || COLORS.accent, display: 'inline-block', flexShrink: 0 }} />
              {p.name}
            </span>
            <span style={{ fontFamily: MONO, fontSize: 13, fontWeight: 700, color: COLORS.ink, whiteSpace: 'nowrap' }}>{fmt(p)}</span>
          </div>
        ))}
      </div>
    </div>
  );
}
// Donut/pie tooltip: names the slice, leads with its value, and shows share of total —
// the reader has the category from the slice they're pointing at, so value comes first.
function DonutTooltip({ active, payload, format, total }) {
  if (!active || !payload || !payload.length) return null;
  const p = payload[0];
  const share = total ? (p.value || 0) / total : 0;
  return (
    <div style={{ ...tooltipStyle, background: '#FFFFFF', padding: '10px 13px', minWidth: 130 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 7, fontFamily: SANS, fontSize: 11.5, color: COLORS.muted, marginBottom: 5 }}>
        <span style={{ width: 9, height: 9, borderRadius: 999, background: p.color, flexShrink: 0 }} />
        {p.name}
      </div>
      <div style={{ fontFamily: MONO, fontSize: 15, fontWeight: 700, color: COLORS.ink }}>{fmtByFormat(p.value, format)}</div>
      <div style={{ fontFamily: SANS, fontSize: 11, color: COLORS.muted, marginTop: 2 }}>{fmtPct(share, 1)} of total</div>
    </div>
  );
}
function Sparkline({ data, dataKey, color }) {
  if (!data || data.length < 2) return null;
  const gid = 'spark-' + dataKey + '-' + Math.random().toString(36).slice(2, 8);
  return (
    <div style={{ height: 32, marginTop: 8 }}>
      <ResponsiveContainer width="100%" height="100%">
        <AreaChart data={data} margin={{ top: 2, right: 0, left: 0, bottom: 0 }}>
          <defs>
            <linearGradient id={gid} x1="0" y1="0" x2="0" y2="1">
              <stop offset="5%" stopColor={color} stopOpacity={0.35} />
              <stop offset="95%" stopColor={color} stopOpacity={0} />
            </linearGradient>
          </defs>
          <Area type="monotone" dataKey={dataKey} stroke={color} strokeWidth={1.75} fill={`url(#${gid})`} isAnimationActive={false} />
        </AreaChart>
      </ResponsiveContainer>
    </div>
  );
}
function KpiCard({ label, value, sub, accent, trend, trendKey, delta }) {
  return (
    <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: '14px 16px' }}>
      <div style={{ fontFamily: SANS, fontSize: 10.5, fontWeight: 600, color: COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: 6 }}>{label}</div>
      <div style={{ display: 'flex', alignItems: 'baseline', gap: 8, flexWrap: 'wrap' }}>
        <div style={{ fontFamily: MONO, fontSize: 22, fontWeight: 600, color: accent || COLORS.ink, lineHeight: 1.15 }}>{value}</div>
        {delta ? <DeltaBadge pct={delta.pct} size={11.5} /> : null}
      </div>
      {sub ? <div style={{ fontFamily: SANS, fontSize: 11, color: COLORS.muted, marginTop: 4 }}>{sub}</div> : null}
      {trend && trendKey ? <Sparkline data={trend} dataKey={trendKey} color={accent || COLORS.accent} /> : null}
    </div>
  );
}
function SectionHeader({ title, subtitle }) {
  return (
    <div style={{ marginBottom: 18 }}>
      <h2 style={{ fontFamily: SANS, fontSize: 19, fontWeight: 700, color: COLORS.ink, margin: 0 }}>{title}</h2>
      {subtitle ? <p style={{ fontFamily: SANS, fontSize: 12.5, color: COLORS.muted, marginTop: 4 }}>{subtitle}</p> : null}
      <div style={{ height: 2, marginTop: 10, background: `linear-gradient(90deg, ${COLORS.accent} 0%, ${COLORS.cyan} 45%, transparent 85%)` }} />
    </div>
  );
}
function Panel({ children, style }) {
  return <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 14, padding: 18, ...style }}>{children}</div>;
}
function EmptyState({ icon: Icon, title, body, action }) {
  return (
    <Panel style={{ textAlign: 'center', padding: '48px 24px' }}>
      <Icon size={28} color={COLORS.muted} style={{ margin: '0 auto 12px auto' }} />
      <div style={{ fontFamily: SANS, fontWeight: 700, fontSize: 14.5, color: COLORS.ink, marginBottom: 6 }}>{title}</div>
      <div style={{ fontFamily: SANS, fontSize: 12.5, color: COLORS.muted, maxWidth: 420, margin: '0 auto' }}>{body}</div>
      {action ? <div style={{ marginTop: 16 }}>{action}</div> : null}
    </Panel>
  );
}
const INSIGHT_STYLES = {
  positive: { icon: CheckCircle2, color: COLORS.emerald, bg: '#EAFBF3', label: 'Strength' },
  concern: { icon: AlertTriangle, color: COLORS.rose, bg: '#FDEEF0', label: 'Pain point' },
  opportunity: { icon: Target, color: COLORS.amberText, bg: '#FFF8EB', label: 'Opportunity' },
};
function InsightCard({ type, title, detail }) {
  const s = INSIGHT_STYLES[type] || INSIGHT_STYLES.opportunity;
  const Icon = s.icon;
  return (
    <div style={{ display: 'flex', gap: 12, padding: 14, borderRadius: 12, border: `1px solid ${COLORS.border}`, background: s.bg }}>
      <Icon size={18} color={s.color} style={{ flexShrink: 0, marginTop: 1 }} />
      <div>
        <div style={{ fontFamily: SANS, fontSize: 10, fontWeight: 700, color: s.color, textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: 3 }}>{s.label}</div>
        <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 700, color: COLORS.ink, marginBottom: 3 }}>{title}</div>
        <div style={{ fontFamily: SANS, fontSize: 12.5, color: '#3A4256', lineHeight: 1.45 }}>{detail}</div>
      </div>
    </div>
  );
}
function DeltaBadge({ pct, size }) {
  const up = pct > 0.001, down = pct < -0.001;
  const color = up ? COLORS.emerald : down ? COLORS.rose : COLORS.muted;
  const Icon = up ? TrendingUp : down ? TrendingDown : Minus;
  const fs = size || 12.5;
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 3, fontFamily: MONO, fontSize: fs, fontWeight: 700, color }}>
      <Icon size={fs + 2} /> {fmtPct(Math.abs(pct), 1)}
    </span>
  );
}
function RadialGoal({ pct, color, label }) {
  const clamped = Math.max(0, Math.min(1, pct));
  const data = [{ name: 'goal', value: clamped * 100, fill: color || COLORS.accent }];
  return (
    <div style={{ position: 'relative', width: 132, height: 132, margin: '0 auto' }}>
      <ResponsiveContainer width="100%" height="100%">
        <RadialBarChart innerRadius="72%" outerRadius="100%" data={data} startAngle={90} endAngle={-270}>
          <PolarAngleAxis type="number" domain={[0, 100]} angleAxisId={0} tick={false} />
          <RadialBar background={{ fill: COLORS.border }} dataKey="value" cornerRadius={999} angleAxisId={0} isAnimationActive={false} />
        </RadialBarChart>
      </ResponsiveContainer>
      <div style={{ position: 'absolute', inset: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', flexDirection: 'column' }}>
        <span style={{ fontFamily: MONO, fontSize: 20, fontWeight: 700, color: COLORS.ink }}>{fmtPct(clamped, 0)}</span>
        {label ? <span style={{ fontFamily: SANS, fontSize: 9.5, color: COLORS.muted, marginTop: 2 }}>{label}</span> : null}
      </div>
    </div>
  );
}
// A cross-tab summary of every active slicer/drill-through filter — clicking a
// donut slice or a breakdown bar sets these same genFilters, but until now the
// only place to see *what's* filtered or clear it was back on the Dashboard tab.
// Surfacing it inline on every tab that reflects genFilteredRows closes that gap.
function ActiveFiltersBar({ filters, onRemove, onClearAll }) {
  const entries = [];
  Object.keys(filters || {}).forEach(dimKey => {
    (filters[dimKey] || []).forEach(value => entries.push({ dimKey, value }));
  });
  if (!entries.length) return null;
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', marginBottom: 16, padding: '8px 10px', background: COLORS.panel, border: `1px solid ${COLORS.border}`, borderRadius: 10 }}>
      <span style={{ fontFamily: SANS, fontSize: 11, fontWeight: 700, color: COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Filtered by</span>
      {entries.map(({ dimKey, value }) => (
        <span key={dimKey + ':' + value} style={{ display: 'inline-flex', alignItems: 'center', gap: 5, fontFamily: SANS, fontSize: 12, fontWeight: 600, padding: '4px 8px 4px 10px', borderRadius: 999, background: COLORS.accent, color: '#FFFFFF' }}>
          {dimKey}: {value}
          <button onClick={() => onRemove(dimKey, value)} aria-label={'Remove filter ' + dimKey + ': ' + value}
            style={{ display: 'flex', alignItems: 'center', background: 'transparent', border: 'none', cursor: 'pointer', color: '#FFFFFF', padding: 0, opacity: 0.85 }}>
            <X size={12} />
          </button>
        </span>
      ))}
      <button onClick={onClearAll} style={{ fontFamily: SANS, fontSize: 11.5, fontWeight: 700, color: COLORS.accent, background: 'transparent', border: 'none', cursor: 'pointer', padding: '4px 2px' }}>
        Clear all
      </button>
    </div>
  );
}
function DonutBreakdown({ data, valueKey, title, centerLabel, centerValue, valueFormat, onSliceClick, selectedValues }) {
  const total = data.reduce((s, d) => s + (d[valueKey] || 0), 0);
  const sel = selectedValues || [];
  return (
    <Panel>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 12, gap: 8 }}>
        <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 600, color: COLORS.ink }}>{title}</div>
        {onSliceClick ? <div style={{ fontFamily: SANS, fontSize: 10.5, color: COLORS.muted }}>Click a segment to filter</div> : null}
      </div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 18, flexWrap: 'wrap' }}>
        <div style={{ position: 'relative', width: 168, height: 168, flexShrink: 0 }}>
          <ResponsiveContainer width="100%" height="100%">
            <PieChart>
              <Pie data={data} dataKey={valueKey} nameKey="name" innerRadius="62%" outerRadius="98%" paddingAngle={2} isAnimationActive={false} strokeWidth={0}
                onClick={onSliceClick ? (d) => onSliceClick(d.name) : undefined} style={onSliceClick ? { cursor: 'pointer' } : undefined}>
                {data.map((entry, i) => (
                  <Cell key={i} fill={CHART_PALETTE[i % CHART_PALETTE.length]} opacity={sel.length > 0 && !sel.includes(entry.name) ? 0.35 : 1} />
                ))}
              </Pie>
              <Tooltip content={<DonutTooltip format={valueFormat || 'number'} total={total} />} />
            </PieChart>
          </ResponsiveContainer>
          <div style={{ position: 'absolute', inset: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', flexDirection: 'column', pointerEvents: 'none' }}>
            <span style={{ fontFamily: MONO, fontSize: 15, fontWeight: 700, color: COLORS.ink }}>{centerValue}</span>
            <span style={{ fontFamily: SANS, fontSize: 9, color: COLORS.muted, marginTop: 1 }}>{centerLabel}</span>
          </div>
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 6, flex: 1, minWidth: 140 }}>
          {data.map((d, i) => {
            const active = sel.includes(d.name);
            const Tag = onSliceClick ? 'button' : 'div';
            return (
              <Tag key={d.name} onClick={onSliceClick ? () => onSliceClick(d.name) : undefined}
                style={{
                  display: 'flex', alignItems: 'center', gap: 8, fontFamily: SANS, fontSize: 12, textAlign: 'left', width: '100%',
                  background: active ? COLORS.accentSoft : 'transparent', border: 'none', borderRadius: 6, padding: '3px 5px',
                  cursor: onSliceClick ? 'pointer' : 'default', opacity: (sel.length > 0 && !active) ? 0.5 : 1,
                }}>
                <span style={{ width: 9, height: 9, borderRadius: 999, background: CHART_PALETTE[i % CHART_PALETTE.length], flexShrink: 0 }} />
                <span style={{ color: COLORS.ink, fontWeight: 600, flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{d.name}</span>
                <span style={{ fontFamily: MONO, color: COLORS.muted, fontSize: 11.5 }}>{total ? fmtPct((d[valueKey] || 0) / total, 0) : '0%'}</span>
              </Tag>
            );
          })}
        </div>
      </div>
    </Panel>
  );
}

function nextMonths(lastMonthStr, n) {
  let [y, m] = lastMonthStr.split('-').map(Number);
  const out = [];
  for (let i = 0; i < n; i++) {
    m += 1;
    if (m > 12) { m = 1; y += 1; }
    out.push(monthLabel(y + '-' + String(m).padStart(2, '0')));
  }
  return out;
}

function ConfigPanel(props) {
  const { dataset, open, onToggleOpen, onToggleMetric, onSetMetricFormat, onSetAggregation, onToggleDimension, onSetDateField, onAddRatio, onUpdateRatio, onRemoveRatio, onSwitchSheet, onSaveConfig, onForgetConfig, saveStatus, savedConfigApplied, onAddBudgetPair, onUpdateBudgetPair, onRemoveBudgetPair } = props;
  const dateCandidates = dataset.schema.filter(s => s.type === 'date');
  const otherCols = dataset.schema.filter(s => s.type !== 'date');
  const sheetNames = dataset.sheets ? Object.keys(dataset.sheets) : [];
  return (
    <Panel style={{ marginBottom: 18, border: `1.5px solid ${COLORS.accent}` }}>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10, flexWrap: 'wrap' }}>
        <button onClick={onToggleOpen} style={{ display: 'flex', alignItems: 'center', gap: 8, background: 'transparent', border: 'none', cursor: 'pointer', padding: 0, flex: 1, minWidth: 200 }}>
          <span style={{ display: 'flex', alignItems: 'center', gap: 8, fontFamily: SANS, fontWeight: 700, fontSize: 13.5, color: COLORS.ink }}>
            <Settings2 size={15} /> Configure analysis — {dataset.name}
          </span>
          {open ? <ChevronUp size={16} color={COLORS.muted} /> : <ChevronDown size={16} color={COLORS.muted} />}
        </button>
        <button onClick={onSaveConfig} title="Remember this setup for files shaped like this"
          style={{ display: 'flex', alignItems: 'center', gap: 6, fontFamily: SANS, fontSize: 11.5, fontWeight: 600, color: saveStatus === 'saved' ? COLORS.emerald : COLORS.ink, background: '#FFFFFF', border: `1px solid ${saveStatus === 'saved' ? COLORS.emerald : COLORS.border}`, borderRadius: 8, padding: '5px 10px', cursor: 'pointer' }}>
          {saveStatus === 'saved' ? <BookmarkCheck size={13} /> : <Save size={13} />}
          {saveStatus === 'saved' ? 'Saved' : 'Remember this setup'}
        </button>
      </div>
      {savedConfigApplied ? (
        <div style={{ marginTop: 10, display: 'flex', alignItems: 'center', gap: 8, background: '#EAFBF3', border: `1px solid ${COLORS.emerald}`, borderRadius: 10, padding: '8px 12px', fontFamily: SANS, fontSize: 11.5, color: '#0B6B4C' }}>
          <BookmarkCheck size={14} /> Applied your saved setup for files shaped like this.
          <button onClick={onForgetConfig} style={{ marginLeft: 'auto', fontFamily: SANS, fontSize: 11, fontWeight: 600, color: '#0B6B4C', background: 'transparent', border: 'none', textDecoration: 'underline', cursor: 'pointer' }}>Forget it</button>
        </div>
      ) : null}
      {open ? (
        <div style={{ marginTop: 16, display: 'flex', flexDirection: 'column', gap: 18 }}>
          {sheetNames.length > 1 ? (
            <div>
              <div style={{ fontFamily: SANS, fontSize: 11, fontWeight: 700, color: COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.04em', marginBottom: 8 }}>Sheet</div>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
                {sheetNames.map(name => (
                  <button key={name} onClick={() => onSwitchSheet(name)}
                    style={{ fontFamily: SANS, fontSize: 12, fontWeight: 600, padding: '6px 12px', borderRadius: 999, cursor: 'pointer', border: `1px solid ${dataset.activeSheet === name ? COLORS.ink : COLORS.border}`, background: dataset.activeSheet === name ? COLORS.ink : '#FFFFFF', color: dataset.activeSheet === name ? '#FFFFFF' : COLORS.ink }}>{name}</button>
                ))}
              </div>
            </div>
          ) : null}
          <div>
            <div style={{ fontFamily: SANS, fontSize: 11, fontWeight: 700, color: COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.04em', marginBottom: 8 }}>Metrics to analyze</div>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10 }}>
              {dataset.config.metrics.map(m => (
                <div key={m.key} style={{ display: 'flex', alignItems: 'center', gap: 6, border: `1px solid ${m.on ? COLORS.accent : COLORS.border}`, borderRadius: 8, padding: '6px 10px', background: m.on ? COLORS.accentSoft : '#FFFFFF' }}>
                  <input type="checkbox" checked={m.on} onChange={() => onToggleMetric(m.key)} />
                  <span style={{ fontFamily: SANS, fontSize: 12.5, color: COLORS.ink, fontWeight: 600 }}>{m.key}</span>
                  {m.on ? <select aria-label={`${m.key} aggregation`} value={m.aggregation || 'sum'} onChange={e => onSetAggregation(m.key, e.target.value)} style={{ fontFamily: SANS, fontSize: 11, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '2px 4px' }}>
                    <option value="sum">Sum</option><option value="average">Average</option><option value="count">Count values</option><option value="min">Minimum</option><option value="max">Maximum</option>
                  </select> : null}
                  {m.on ? (
                    <select value={m.format} onChange={e => onSetMetricFormat(m.key, e.target.value)} style={{ fontFamily: SANS, fontSize: 11, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '2px 4px' }}>
                      <option value="number">#</option>
                      <option value="currency">$</option>
                      <option value="percent">%</option>
                    </select>
                  ) : null}
                </div>
              ))}
            </div>
          </div>
          <div>
            <div style={{ fontFamily: SANS, fontSize: 11, fontWeight: 700, color: COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.04em', marginBottom: 8 }}>Filter / breakdown dimensions</div>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10 }}>
              {dataset.config.dimensions.map(d => (
                <div key={d.key} style={{ display: 'flex', alignItems: 'center', gap: 6, border: `1px solid ${d.on ? COLORS.accent : COLORS.border}`, borderRadius: 8, padding: '6px 10px', background: d.on ? COLORS.accentSoft : '#FFFFFF' }}>
                  <input type="checkbox" checked={d.on} onChange={() => onToggleDimension(d.key)} />
                  <span style={{ fontFamily: SANS, fontSize: 12.5, color: COLORS.ink, fontWeight: 600 }}>{d.key}</span>
                  <span style={{ fontFamily: MONO, fontSize: 10, color: d.unique > 15 ? COLORS.rose : COLORS.muted }}>{d.unique}v</span>
                </div>
              ))}
              {dataset.config.dimensions.length === 0 ? <span style={{ fontFamily: SANS, fontSize: 12, color: COLORS.muted }}>No category-like columns detected.</span> : null}
            </div>
          </div>
          <div>
            <div style={{ fontFamily: SANS, fontSize: 11, fontWeight: 700, color: COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.04em', marginBottom: 8 }}>Date column (enables Trends, Forecast &amp; Insights trends)</div>
            <select value={dataset.config.dateField} onChange={e => onSetDateField(e.target.value)}
              style={{ fontFamily: SANS, fontSize: 12.5, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: '6px 10px' }}>
              <option value="">None</option>
              {dateCandidates.map(s => <option key={s.key} value={s.key}>{s.key}</option>)}
              {otherCols.map(s => <option key={s.key} value={s.key}>{s.key} (unverified)</option>)}
            </select>
          </div>
          <div>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
              <div style={{ fontFamily: SANS, fontSize: 11, fontWeight: 700, color: COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Custom metrics (e.g. Profit = Revenue − Cost)</div>
              <button onClick={onAddRatio} disabled={dataset.config.metrics.filter(m => m.on).length < 2}
                style={{ display: 'flex', alignItems: 'center', gap: 4, fontFamily: SANS, fontSize: 11.5, fontWeight: 600, color: COLORS.ink, background: '#FFFFFF', border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '4px 8px', cursor: 'pointer' }}>
                <Plus size={12} /> Add metric
              </button>
            </div>
            {dataset.config.ratios.length === 0 ? <div style={{ fontFamily: SANS, fontSize: 11.5, color: COLORS.muted }}>None yet — build one from two active metrics, e.g. "Profit" = Revenue − Cost, or "Cost per Unit" = Total Cost ÷ Units Sold.</div> : null}
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              {dataset.config.ratios.map(rt => (
                <div key={rt.key} style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                  <input value={rt.label} onChange={e => onUpdateRatio(rt.key, { label: e.target.value })}
                    style={{ fontFamily: SANS, fontSize: 12, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '5px 8px', width: 110 }} />
                  <span style={{ fontFamily: SANS, fontSize: 11, color: COLORS.muted }}>=</span>
                  <select value={rt.num} onChange={e => onUpdateRatio(rt.key, { num: e.target.value })} style={{ fontFamily: SANS, fontSize: 12, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '5px 6px' }}>
                    {dataset.config.metrics.map(m => <option key={m.key} value={m.key}>{m.key}</option>)}
                  </select>
                  <select value={rt.op || '÷'} onChange={e => onUpdateRatio(rt.key, { op: e.target.value })} title="Operator"
                    style={{ fontFamily: MONO, fontSize: 12, fontWeight: 700, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '5px 4px', color: COLORS.accent }}>
                    <option value="+">+</option>
                    <option value="−">−</option>
                    <option value="÷">÷</option>
                  </select>
                  <select value={rt.den} onChange={e => onUpdateRatio(rt.key, { den: e.target.value })} style={{ fontFamily: SANS, fontSize: 12, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '5px 6px' }}>
                    {dataset.config.metrics.map(m => <option key={m.key} value={m.key}>{m.key}</option>)}
                  </select>
                  <select value={rt.format || 'number'} onChange={e => onUpdateRatio(rt.key, { format: e.target.value })} title="Display format"
                    style={{ fontFamily: SANS, fontSize: 11, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '5px 4px' }}>
                    <option value="number">#</option>
                    <option value="currency">$</option>
                    <option value="percent">%</option>
                  </select>
                  <button onClick={() => onRemoveRatio(rt.key)} style={{ background: 'transparent', border: 'none', cursor: 'pointer', color: COLORS.rose }}><X size={14} /></button>
                </div>
              ))}
            </div>
          </div>
          <div>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
              <div style={{ fontFamily: SANS, fontSize: 11, fontWeight: 700, color: COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Budget vs Actual pairs</div>
              <button onClick={onAddBudgetPair} disabled={dataset.config.metrics.filter(m => m.on).length < 2}
                style={{ display: 'flex', alignItems: 'center', gap: 4, fontFamily: SANS, fontSize: 11.5, fontWeight: 600, color: COLORS.ink, background: '#FFFFFF', border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '4px 8px', cursor: 'pointer' }}>
                <Plus size={12} /> Add pair
              </button>
            </div>
            {(dataset.config.budgetPairs || []).length === 0 ? <div style={{ fontFamily: SANS, fontSize: 11.5, color: COLORS.muted }}>None yet — pair a Budget column with an Actual column to track variance.</div> : null}
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              {(dataset.config.budgetPairs || []).map(p => (
                <div key={p.key} style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                  <input value={p.label} onChange={e => onUpdateBudgetPair(p.key, { label: e.target.value })}
                    style={{ fontFamily: SANS, fontSize: 12, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '5px 8px', width: 130 }} />
                  <span style={{ fontFamily: SANS, fontSize: 11, color: COLORS.muted }}>Budget:</span>
                  <select value={p.budget} onChange={e => onUpdateBudgetPair(p.key, { budget: e.target.value })} style={{ fontFamily: SANS, fontSize: 12, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '5px 6px' }}>
                    {dataset.config.metrics.map(m => <option key={m.key} value={m.key}>{m.key}</option>)}
                  </select>
                  <span style={{ fontFamily: SANS, fontSize: 11, color: COLORS.muted }}>Actual:</span>
                  <select value={p.actual} onChange={e => onUpdateBudgetPair(p.key, { actual: e.target.value })} style={{ fontFamily: SANS, fontSize: 12, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '5px 6px' }}>
                    {dataset.config.metrics.map(m => <option key={m.key} value={m.key}>{m.key}</option>)}
                  </select>
                  <button onClick={() => onRemoveBudgetPair(p.key)} style={{ background: 'transparent', border: 'none', cursor: 'pointer', color: COLORS.rose }}><X size={14} /></button>
                </div>
              ))}
            </div>
          </div>
        </div>
      ) : null}
    </Panel>
  );
}

const SUPPORTED_EXTS = ['csv', 'tsv', 'txt', 'json', 'xlsx', 'xls'];

function ImportPreview({ pending, onSheet, onType, onName, onConfirm, onCancel }) {
  const sheet = pending.sheets[pending.activeSheet];
  const names = sheet.fields.map(field => (pending.names[field] || '').trim());
  const hasDuplicateNames = new Set(names.map(name => name.toLowerCase())).size !== names.length;
  const invalid = sheet.fields.map(field => {
    const type = pending.types[field];
    if (type !== 'number' && type !== 'date' && type !== 'date-dmy') return null;
    const count = sheet.rows.filter(row => row[field] !== null && row[field] !== undefined && row[field] !== '' &&
      (type === 'number' ? parseNumericValue(row[field]) : parseDateValue(row[field], type === 'date-dmy' ? 'dmy' : 'mdy')) === null).length;
    return count ? `${field}: ${count} unparseable ${type} value${count === 1 ? '' : 's'}` : null;
  }).filter(Boolean);
  return (
    <Panel style={{ border: `1.5px solid ${COLORS.accent}` }}>
      <SectionHeader title="Review your import" subtitle={`${pending.name} · ${sheet.rows.length} rows · ${pending.activeSheet}`} />
      {Object.keys(pending.sheets).length > 1 ? (
        <label style={{ display: 'block', marginBottom: 16, fontSize: 12 }}>Sheet{' '}
          <select value={pending.activeSheet} onChange={e => onSheet(e.target.value)}>
            {Object.keys(pending.sheets).map(name => <option key={name} value={name}>{name}</option>)}
          </select>
        </label>
      ) : null}
      <p style={{ color: COLORS.muted, fontSize: 12 }}>Confirm column names and types before analysis. Formatted numbers such as $1,000 are converted; invalid values become blank and are counted below.</p>
      <div style={{ overflowX: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
          <thead><tr><th style={{ textAlign: 'left' }}>Source column</th><th style={{ textAlign: 'left' }}>Use as</th><th style={{ textAlign: 'left' }}>Column name</th><th style={{ textAlign: 'left' }}>Sample values</th></tr></thead>
          <tbody>{sheet.fields.map(field => (
            <tr key={field} style={{ borderTop: `1px solid ${COLORS.border}` }}>
              <td style={{ padding: '9px 8px' }}>{field}</td>
              <td style={{ padding: '9px 8px' }}><select aria-label={`${field} type`} value={pending.types[field]} onChange={e => onType(field, e.target.value)}>
                <option value="number">Number</option><option value="date">Date (month/day/year)</option><option value="date-dmy">Date (day/month/year)</option><option value="category">Category / text</option>
              </select></td>
              <td style={{ padding: '9px 8px' }}><input aria-label={`${field} column name`} value={pending.names[field]} onChange={e => onName(field, e.target.value)} /></td>
              <td style={{ padding: '9px 8px', color: COLORS.muted }}>{sheet.rows.slice(0, 3).map(row => String(row[field] ?? '—')).join(' · ')}</td>
            </tr>
          ))}</tbody>
        </table>
      </div>
      {invalid.length ? <p role="status" style={{ color: COLORS.amberText, fontSize: 12 }}>{invalid.join('; ')}</p> : null}
      {hasDuplicateNames || names.some(name => !name) ? <p role="alert" style={{ color: COLORS.rose, fontSize: 12 }}>Column names must be unique and nonempty.</p> : null}
      <div style={{ display: 'flex', gap: 10, marginTop: 18 }}>
        <button onClick={onConfirm} disabled={hasDuplicateNames || names.some(name => !name)} style={{ background: COLORS.accent, color: '#fff', border: 0, borderRadius: 8, padding: '10px 16px', cursor: 'pointer' }}>Analyze data</button>
        <button onClick={onCancel} style={{ background: '#fff', border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: '10px 16px', cursor: 'pointer' }}>Cancel</button>
      </div>
    </Panel>
  );
}

export default function FinalysisApp() {
  const [tab, setTab] = useState('dashboard');
  const [search, setSearch] = useState('');
  const [sortKey, setSortKey] = useState('');
  const [sortDir, setSortDir] = useState('asc');
  const [page, setPage] = useState(0);

  const [customDataset, setCustomDataset] = useState(null);
  const [pendingImport, setPendingImport] = useState(null);
  const [configOpen, setConfigOpen] = useState(true);
  const [uploadError, setUploadError] = useState(null);
  const [isDragOver, setIsDragOver] = useState(false);
  const [genFilters, setGenFilters] = useState({});
  const [genBreakdownDim, setGenBreakdownDim] = useState('');
  const [genForecastMetric, setGenForecastMetric] = useState('');
  const [genScenario, setGenScenario] = useState('Base');
  const [genCustomGrowth, setGenCustomGrowth] = useState(0);
  const [genForecastHorizon, setGenForecastHorizon] = useState(6);
  const [goalSeekTarget, setGoalSeekTarget] = useState('');
  const [hiddenTrendSeries, setHiddenTrendSeries] = useState({});
  const [trendChartType, setTrendChartType] = useState('area');
  const [sensOutcomeKey, setSensOutcomeKey] = useState('');
  const [sensSpread, setSensSpread] = useState(5);
  const [sensSteps, setSensSteps] = useState(5);
  const [aiInsights, setAiInsights] = useState({ status: 'idle', data: null, error: null });
  const [qaHistory, setQaHistory] = useState([]);
  const [qaInput, setQaInput] = useState('');
  const [qaStatus, setQaStatus] = useState('idle');
  const [saveStatus, setSaveStatus] = useState('idle');
  const [savedConfigApplied, setSavedConfigApplied] = useState(false);
  const [savedViews, setSavedViews] = useState([]);
  const [newViewName, setNewViewName] = useState('');
  const [showViewInput, setShowViewInput] = useState(false);
  const [compareOn, setCompareOn] = useState(false);
  const [compareA, setCompareA] = useState('');
  const [compareB, setCompareB] = useState('');
  const [goalTarget, setGoalTarget] = useState('');
  const [alertRule, setAlertRule] = useState({ metricKey: '', op: 'below', value: '' });
  const [periodCompareMode, setPeriodCompareMode] = useState('prior'); // 'prior' | 'yoy'
  const [periodCompareMetric, setPeriodCompareMetric] = useState('');
  const [fileCompareDataset, setCompareDataset] = useState(null);
  const [fileCompareUploadError, setCompareUploadError] = useState(null);
  const [fileCompareDragOver, setCompareDragOver] = useState(false);
  const [fileCompareBreakdownDim, setCompareBreakdownDim] = useState('');
  const fileInputRef = useRef(null);
  const fileCompareInputRef = useRef(null);

  const activeMetrics = customDataset ? customDataset.config.metrics.filter(m => m.on) : [];
  const activeDimensions = customDataset ? customDataset.config.dimensions.filter(d => d.on) : [];
  const activeRatios = customDataset ? customDataset.config.ratios : [];
  const activeBudgetPairs = customDataset ? (customDataset.config.budgetPairs || []) : [];
  const dateField = customDataset ? customDataset.config.dateField : '';
  // Stable string keys (rather than the freshly-filtered arrays themselves) so the
  // sync effects below only fire when the actual set of active keys changes.
  const activeMetricKeys = activeMetrics.map(m => m.key).join('|');
  const activeDimensionKeys = activeDimensions.map(d => d.key).join('|');
  const activeRatioKeys = activeRatios.map(r => r.key).join('|');

  // Slicers: genFilters[dimKey] is an array of selected values (Power BI-style
  // multi-select). An empty/missing array means "All" — no filtering on that dimension.
  const genFilteredRows = useMemo(() => {
    if (!customDataset) return [];
    return customDataset.rows.filter(r => activeDimensions.every(d => {
      const sel = genFilters[d.key];
      if (!sel || !sel.length) return true;
      const raw = r[d.key];
      const val = (raw === null || raw === undefined || raw === '') ? '(blank)' : String(raw);
      return sel.includes(val);
    }));
  }, [customDataset, genFilters, activeDimensions]);

  const genKpi = useMemo(() => {
    if (!customDataset) return { count: 0 };
    const agg = aggregateMetrics(genFilteredRows, activeMetrics);
    const ratioVals = {};
    activeRatios.forEach(rt => { ratioVals[rt.key] = computeRatio(genFilteredRows, rt); });
    return { ...agg, ...ratioVals };
  }, [genFilteredRows, activeMetrics, activeRatios]);

  const genBudgetVariance = useMemo(() => {
    return activeBudgetPairs.map(p => {
      const actualVal = genKpi[p.actual] || 0;
      const budgetVal = genKpi[p.budget] || 0;
      const { variance, pct } = computeVariance(actualVal, budgetVal);
      return { ...p, actualVal, budgetVal, variance, pct };
    });
  }, [activeBudgetPairs, genKpi]);

  const dataQuality = useMemo(() => {
    if (!customDataset) return null;
    return computeDataQuality(customDataset.rows, customDataset.fields, dateField);
  }, [customDataset, dateField]);

  const negativeCounts = useMemo(() => {
    if (!customDataset) return [];
    return computeNegativeCounts(customDataset.rows, activeMetrics.map(m => m.key));
  }, [customDataset, activeMetrics]);

  // All of these scope to genFilteredRows (the slicer-filtered set), not the raw
  // upload, so every chart/stat/table on Breakdown, Trends, Forecast, and Insights
  // agrees with the slicers shown on Dashboard — filtering something there now
  // actually changes what you see everywhere else, instead of only the KPI cards.
  const genBreakdown = useMemo(() => {
    if (!customDataset || !genBreakdownDim) return [];
    return groupByGeneric(genFilteredRows, genBreakdownDim, activeMetrics, activeRatios);
  }, [customDataset, genFilteredRows, genBreakdownDim, activeMetrics, activeRatios]);

  const genTrend = useMemo(() => {
    if (!customDataset || !dateField) return [];
    return monthlyGroupGeneric(genFilteredRows, dateField, activeMetrics, activeRatios);
  }, [customDataset, genFilteredRows, dateField, activeMetrics, activeRatios]);

  const genFirstDimAgg = useMemo(() => {
    if (!customDataset || !activeDimensions[0]) return [];
    return groupByGeneric(genFilteredRows, activeDimensions[0].key, activeMetrics, activeRatios);
  }, [customDataset, genFilteredRows, activeDimensions, activeMetrics, activeRatios]);

  const genForecastBaseGrowth = useMemo(() => computeAvgGrowth(genTrend, genForecastMetric), [genTrend, genForecastMetric]);
  const forecastBacktest = useMemo(() => backtestForecast(genTrend, genForecastMetric), [genTrend, genForecastMetric]);
  const genScenarioRates = {
    Base: genForecastBaseGrowth,
    Optimistic: genForecastBaseGrowth + Math.max(0.02, Math.abs(genForecastBaseGrowth) * 0.6),
    Pessimistic: genForecastBaseGrowth - Math.max(0.02, Math.abs(genForecastBaseGrowth) * 0.6),
    Custom: genCustomGrowth,
  };
  const genActiveGrowth = genScenarioRates[genScenario];
  const genForecastLabels = useMemo(() => genTrend.length ? nextMonths(genTrend[genTrend.length - 1].month, genForecastHorizon) : [], [genTrend, genForecastHorizon]);
  const genBaseValue = genTrend.length ? (genTrend[genTrend.length - 1][genForecastMetric] || 0) : 0;
  const genForecast = useMemo(() => genForecastLabels.length ? projectGeneric(genBaseValue, genActiveGrowth, genForecastLabels) : [],
    [genBaseValue, genActiveGrowth, genForecastLabels]);
  // Memoized so unrelated re-renders elsewhere on the Forecast tab (e.g. the
  // sensitivity grid's controls) don't hand Recharts a new array identity and
  // trigger a pointless replay of the chart's mount-in animation.
  const genForecastChartData = useMemo(() => [
    { label: (genTrend[genTrend.length - 1] || {}).label, value: genBaseValue, actual: true },
    ...genForecast,
  ], [genTrend, genBaseValue, genForecast]);
  const genSnapshots = useMemo(() => (['Base', 'Optimistic', 'Pessimistic']).map(name => {
    const rate = genScenarioRates[name];
    return { name, value: genBaseValue * Math.pow(1 + rate, genForecastHorizon) };
  }), [genBaseValue, genForecastBaseGrowth, genForecastHorizon]);
  // Goal seek: given a target value for the final forecast period, back-solve the
  // constant monthly growth rate that would reach it from the current base period.
  const goalSeekRate = useMemo(() => {
    const target = Number(goalSeekTarget);
    if (!goalSeekTarget || isNaN(target) || target <= 0 || genBaseValue <= 0 || !genForecastHorizon) return null;
    return Math.pow(target / genBaseValue, 1 / genForecastHorizon) - 1;
  }, [goalSeekTarget, genBaseValue, genForecastHorizon]);

  // ---- Sensitivity grid: vary two forecast assumptions at once ----
  // Only meaningful for a calculated metric (e.g. Profit = Revenue − Cost) —
  // a plain metric has just one growth lever, so there's nothing to cross it
  // against. Each axis is centered on that input's OWN historical average
  // growth (same "auto-estimated from your own history" logic as the single-
  // lever forecast above) and spread symmetrically by a user-adjustable range.
  const sensOutcomeRatio = useMemo(() => activeRatios.find(r => r.key === sensOutcomeKey) || null, [activeRatios, sensOutcomeKey]);
  const sensNumBaseGrowth = useMemo(() => sensOutcomeRatio ? computeAvgGrowth(genTrend, sensOutcomeRatio.num) : 0, [genTrend, sensOutcomeRatio]);
  const sensDenBaseGrowth = useMemo(() => sensOutcomeRatio ? computeAvgGrowth(genTrend, sensOutcomeRatio.den) : 0, [genTrend, sensOutcomeRatio]);
  const sensNumBaseValue = (genTrend.length && sensOutcomeRatio) ? (genTrend[genTrend.length - 1][sensOutcomeRatio.num] || 0) : 0;
  const sensDenBaseValue = (genTrend.length && sensOutcomeRatio) ? (genTrend[genTrend.length - 1][sensOutcomeRatio.den] || 0) : 0;
  function sensAxisValues(centerGrowth) {
    const steps = Math.max(3, sensSteps | 0);
    const spread = Math.max(0.5, Number(sensSpread) || 5) / 100; // percentage points -> fraction
    const half = Math.floor(steps / 2);
    return Array.from({ length: steps }, (_, i) => centerGrowth + (i - half) * (spread / half || spread));
  }
  const sensRowGrowths = useMemo(() => sensAxisValues(sensNumBaseGrowth), [sensNumBaseGrowth, sensSpread, sensSteps]);
  const sensColGrowths = useMemo(() => sensAxisValues(sensDenBaseGrowth), [sensDenBaseGrowth, sensSpread, sensSteps]);
  const sensGrid = useMemo(() => {
    if (!sensOutcomeRatio) return [];
    return sensRowGrowths.map(rg => sensColGrowths.map(cg => {
      const numProj = sensNumBaseValue * Math.pow(1 + rg, genForecastHorizon);
      const denProj = sensDenBaseValue * Math.pow(1 + cg, genForecastHorizon);
      return applyCalcOp(numProj, denProj, sensOutcomeRatio.op);
    }));
  }, [sensOutcomeRatio, sensRowGrowths, sensColGrowths, sensNumBaseValue, sensDenBaseValue, genForecastHorizon]);
  const sensFlatValues = useMemo(() => sensGrid.flat(), [sensGrid]);
  const sensMin = sensFlatValues.length ? Math.min(...sensFlatValues) : 0;
  const sensMax = sensFlatValues.length ? Math.max(...sensFlatValues) : 0;
  const sensBaseCellValue = sensOutcomeRatio ? applyCalcOp(
    sensNumBaseValue * Math.pow(1 + sensNumBaseGrowth, genForecastHorizon),
    sensDenBaseValue * Math.pow(1 + sensDenBaseGrowth, genForecastHorizon),
    sensOutcomeRatio.op
  ) : 0;
  function sensCellColor(value) {
    // Diverging shade around the grid's own midpoint — darker green for the
    // higher end of the range shown, darker red for the lower end. This
    // colors by VALUE, not by an assumed "higher is always better" — the
    // panel copy makes that explicit so a ratio where lower is actually
    // better (e.g. a cost-to-revenue ratio) isn't visually mislabeled.
    const mid = (sensMin + sensMax) / 2;
    const span = (sensMax - sensMin) / 2 || 1;
    const t = Math.max(-1, Math.min(1, (value - mid) / span));
    return t >= 0
      ? `rgba(4, 120, 87, ${0.08 + t * 0.30})`   // COLORS.emerald wash
      : `rgba(244, 63, 94, ${0.08 + -t * 0.30})`; // COLORS.rose wash
  }

  const genAllSegments = useMemo(() => {
    if (!customDataset) return [];
    const out = [];
    activeDimensions.forEach(d => {
      groupByGeneric(genFilteredRows, d.key, activeMetrics, activeRatios).forEach(g => out.push({ name: d.key + ': ' + g.name, ...g }));
    });
    return out;
  }, [customDataset, genFilteredRows, activeDimensions, activeMetrics, activeRatios]);
  // 'yoy' needs 13 monthly points (current + the same month a year back); fall
  // back to comparing against the prior period if the dataset is too short.
  const periodCompareOffset = periodCompareMode === 'yoy' && genTrend.length >= 13 ? 12 : 1;
  const genPeriodDeltaData = useMemo(() => {
    if (!genTrend.length) return {};
    const out = {};
    activeMetrics.concat(activeRatios).forEach(m => { out[m.key] = periodDelta(genTrend, m.key, periodCompareOffset); });
    return out;
  }, [genTrend, activeMetrics, activeRatios, periodCompareOffset]);
  const genCorrelations = useMemo(() => {
    if (!customDataset || activeMetrics.length < 2) return [];
    return computeCorrelations(genFilteredRows, activeMetrics.map(m => m.key));
  }, [customDataset, genFilteredRows, activeMetrics]);
  const genOutlierMetric = activeMetrics[0] ? activeMetrics[0].key : '';
  const genOutliers = useMemo(() => {
    if (!customDataset || !genOutlierMetric) return [];
    return computeOutliers(genFilteredRows, genOutlierMetric, 2.5);
  }, [customDataset, genFilteredRows, genOutlierMetric]);
  const genAlertOptions = useMemo(() => [
    ...activeMetrics.map(m => ({ key: m.key, label: m.key, format: m.format })),
    ...activeRatios.map(r => ({ key: r.key, label: r.label, format: r.format })),
  ], [activeMetrics, activeRatios]);

  const activeMetricFormat = (key) => {
    const m = customDataset && customDataset.config.metrics.find(mm => mm.key === key);
    return m ? m.format : 'number';
  };

  const insightInput = useMemo(() => {
    if (!customDataset || activeMetrics.length === 0) return null;
    const primary = activeMetrics[0];
    const breakdowns = activeDimensions.map(d => ({
      dimLabel: d.key,
      groups: groupByGeneric(genFilteredRows, d.key, [primary], []).map(g => ({ name: g.name, value: g[primary.key], count: g.count })),
    }));
    const trend = dateField && genTrend.length ? genTrend.map(t => ({ label: t.label, value: t[primary.key] })) : [];
    const ratio = (activeRatios[0] && activeDimensions[0])
      ? { label: activeRatios[0].label, format: activeRatios[0].format, groups: groupByGeneric(genFilteredRows, activeDimensions[0].key, [], [activeRatios[0]]).map(g => ({ name: g.name, value: g[activeRatios[0].key] })) }
      : null;
    return { metricLabel: primary.key, metricFormat: primary.format, totalValue: genKpi[primary.key], rowCount: genFilteredRows.length, breakdowns, trend, ratio };
  }, [customDataset, genFilteredRows, activeMetrics, activeDimensions, activeRatios, dateField, genTrend, genKpi]);

  const ruleInsights = useMemo(() => insightInput ? computeInsights(insightInput) : [], [insightInput]);

  // A few clickable example questions grounded in this dataset's own field
  // names — shows what's actually answerable instead of a blank box (the same
  // "suggested questions" pattern most NL-query tools lean on, since an empty
  // input gives no signal of scope or phrasing).
  const qaSuggestions = useMemo(() => {
    if (!insightInput) return [];
    const out = [];
    const dim = activeDimensions[0];
    const metric = insightInput.metricLabel;
    if (dim) out.push('Which ' + dim.key + ' has the highest ' + metric + '?');
    if (dateField) out.push('How has ' + metric + ' trended over time?');
    if (activeRatios[0]) out.push('What is total ' + activeRatios[0].label + '?');
    if (dim && activeDimensions[1]) out.push('Compare ' + dim.key + ' and ' + activeDimensions[1].key + ' by ' + metric + '.');
    return out.slice(0, 3);
  }, [insightInput, activeDimensions, activeRatios, dateField]);

  const allSegments = genAllSegments;
  const alertOptions = genAlertOptions;
  const correlations = genCorrelations;
  const outliers = genOutliers;
  const outlierMetricLabel = genOutlierMetric;
  const outlierMetricFormat = activeMetricFormat(genOutlierMetric);

  const alertMatches = useMemo(() => {
    if (!alertRule.metricKey || alertRule.value === '' || isNaN(Number(alertRule.value))) return [];
    const thresh = Number(alertRule.value);
    return allSegments.filter(seg => {
      const v = seg[alertRule.metricKey];
      if (v === undefined || v === null || isNaN(v)) return false;
      return alertRule.op === 'below' ? v < thresh : v > thresh;
    });
  }, [alertRule, allSegments]);

  const compareDimOptions = genBreakdown;
  const compareSegA = compareDimOptions.find(s => s.name === compareA);
  const compareSegB = compareDimOptions.find(s => s.name === compareB);
  const compareMetricList = [...activeMetrics.map(m => ({ key: m.key, label: m.key, format: m.format })), ...activeRatios.map(r => ({ key: r.key, label: r.label, format: r.format }))];
  const periodCompareEntry = compareMetricList.find(m => m.key === periodCompareMetric);
  const periodCompareLabel = periodCompareEntry ? periodCompareEntry.label : periodCompareMetric;
  const periodCompareFormat = periodCompareEntry ? periodCompareEntry.format : 'number';
  const periodCompareDelta = periodCompareMetric ? genPeriodDeltaData[periodCompareMetric] : null;

  // ============ Two-file comparison (exact-column-name matching only) ============
  // Dataset A's active metrics/dimensions drive the comparison; a column only
  // participates if the exact same key exists in dataset B's raw fields — no
  // fuzzy matching, so what's compared is always auditable from the column
  // names alone. Ratios need BOTH underlying columns present in B.
  const comparableNumeric = new Set(fileCompareDataset ? fileCompareDataset.schema.filter(col => col.type === 'number').map(col => col.key) : []);
  const fileCompareMatchedMetrics = fileCompareDataset ? activeMetrics.filter(m => comparableNumeric.has(m.key)) : [];
  const fileCompareUnmatchedMetrics = fileCompareDataset ? activeMetrics.filter(m => !comparableNumeric.has(m.key)) : [];
  const fileCompareMatchedRatios = fileCompareDataset ? activeRatios.filter(r => comparableNumeric.has(r.num) && comparableNumeric.has(r.den)) : [];
  const fileCompareUnmatchedRatios = fileCompareDataset ? activeRatios.filter(r => !(comparableNumeric.has(r.num) && comparableNumeric.has(r.den))) : [];
  const fileCompareMatchedDimensions = fileCompareDataset ? activeDimensions.filter(d => fileCompareDataset.fields.includes(d.key)) : [];
  const fileCompareUnmatchedDimensions = fileCompareDataset ? activeDimensions.filter(d => !fileCompareDataset.fields.includes(d.key)) : [];
  const fileCompareMatchedMetricKeys = fileCompareMatchedMetrics.map(m => m.key).join('|');
  const fileCompareMatchedRatioKeys = fileCompareMatchedRatios.map(r => r.key).join('|');
  const fileCompareMatchedDimKeys = fileCompareMatchedDimensions.map(d => d.key).join('|');

  const fileCompareTotalsA = useMemo(() => {
    if (!customDataset) return null;
    const agg = aggregateMetrics(customDataset.rows, fileCompareMatchedMetrics);
    fileCompareMatchedRatios.forEach(rt => { agg[rt.key] = computeRatio(customDataset.rows, rt); });
    return agg;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [customDataset, fileCompareMatchedMetricKeys, fileCompareMatchedRatioKeys]);
  const fileCompareTotalsB = useMemo(() => {
    if (!fileCompareDataset) return null;
    const agg = aggregateMetrics(fileCompareDataset.rows, fileCompareMatchedMetrics);
    fileCompareMatchedRatios.forEach(rt => { agg[rt.key] = computeRatio(fileCompareDataset.rows, rt); });
    return agg;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fileCompareDataset, fileCompareMatchedMetricKeys, fileCompareMatchedRatioKeys]);

  const fileCompareMetricRows = [...fileCompareMatchedMetrics.map(m => ({ key: m.key, label: m.key, format: m.format })), ...fileCompareMatchedRatios.map(r => ({ key: r.key, label: r.label, format: r.format }))];

  const fileCompareSegmentTable = useMemo(() => {
    if (!customDataset || !fileCompareDataset || !fileCompareBreakdownDim) return [];
    const groupA = groupByGeneric(customDataset.rows, fileCompareBreakdownDim, fileCompareMatchedMetrics, fileCompareMatchedRatios);
    const groupB = groupByGeneric(fileCompareDataset.rows, fileCompareBreakdownDim, fileCompareMatchedMetrics, fileCompareMatchedRatios);
    const byNameA = Object.fromEntries(groupA.map(g => [g.name, g]));
    const byNameB = Object.fromEntries(groupB.map(g => [g.name, g]));
    const names = Array.from(new Set([...groupA.map(g => g.name), ...groupB.map(g => g.name)])).sort();
    return names.map(name => ({ name, a: byNameA[name] || null, b: byNameB[name] || null }));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [customDataset, fileCompareDataset, fileCompareBreakdownDim, fileCompareMatchedMetricKeys, fileCompareMatchedRatioKeys]);

  const goalMetricLabel = activeMetrics[0] ? activeMetrics[0].key : '';
  const goalMetricFormat = activeMetrics[0] ? activeMetrics[0].format : 'number';
  const goalCurrentValue = activeMetrics[0] ? genKpi[activeMetrics[0].key] : 0;
  const goalPct = (goalTarget !== '' && !isNaN(Number(goalTarget)) && Number(goalTarget) !== 0) ? (goalCurrentValue / Number(goalTarget)) : 0;

  useEffect(() => {
    setCompareA(compareDimOptions[0] ? compareDimOptions[0].name : '');
    setCompareB(compareDimOptions[1] ? compareDimOptions[1].name : '');
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [genBreakdownDim]);

  // Bug fix: if the metric/dimension currently selected for forecasting, breakdown,
  // or a threshold alert gets turned off in Configure analysis, these selections used
  // to keep pointing at the now-hidden field — the <select> would silently fall back to
  // displaying a different option than what was actually being computed, showing stale
  // or all-zero results (e.g. Forecast tab) without any indication anything was wrong.
  // Snap each selection back to a still-active field (or clear it) whenever the active
  // set changes.
  useEffect(() => {
    if (genForecastMetric && !activeMetrics.some(m => m.key === genForecastMetric)) {
      setGenForecastMetric(activeMetrics[0] ? activeMetrics[0].key : '');
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeMetricKeys]);

  useEffect(() => {
    if (genBreakdownDim && !activeDimensions.some(d => d.key === genBreakdownDim)) {
      setGenBreakdownDim(activeDimensions[0] ? activeDimensions[0].key : '');
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeDimensionKeys]);

  useEffect(() => {
    if (alertRule.metricKey && !activeMetrics.some(m => m.key === alertRule.metricKey) && !activeRatios.some(r => r.key === alertRule.metricKey)) {
      setAlertRule(a => ({ ...a, metricKey: '' }));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeMetricKeys, activeRatioKeys]);

  // Keep the period-over-period metric picker pointed at a real, currently-active
  // metric or calculated metric — same "snap back rather than go stale" pattern
  // used for forecast/breakdown/alert selections above.
  useEffect(() => {
    if (periodCompareMetric && !activeMetrics.some(m => m.key === periodCompareMetric) && !activeRatios.some(r => r.key === periodCompareMetric)) {
      setPeriodCompareMetric(activeMetrics[0] ? activeMetrics[0].key : (activeRatios[0] ? activeRatios[0].key : ''));
    } else if (!periodCompareMetric && (activeMetrics[0] || activeRatios[0])) {
      setPeriodCompareMetric(activeMetrics[0] ? activeMetrics[0].key : activeRatios[0].key);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeMetricKeys, activeRatioKeys]);

  // A 'year over year' comparison silently falls back to 'prior period' above
  // once the dataset is too short (see periodCompareOffset) — snap the toggle
  // itself back too so the UI doesn't keep showing a selected mode it can't
  // actually deliver.
  useEffect(() => {
    if (periodCompareMode === 'yoy' && genTrend.length < 13) setPeriodCompareMode('prior');
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [genTrend.length]);

  // Keep the two-file comparison's segment dimension pointed at one that's
  // actually matched between both files (same pattern as the other selection
  // guards above) — auto-pick the first matched dimension once one exists.
  useEffect(() => {
    if (fileCompareBreakdownDim && !fileCompareMatchedDimensions.some(d => d.key === fileCompareBreakdownDim)) {
      setCompareBreakdownDim(fileCompareMatchedDimensions[0] ? fileCompareMatchedDimensions[0].key : '');
    } else if (!fileCompareBreakdownDim && fileCompareMatchedDimensions[0]) {
      setCompareBreakdownDim(fileCompareMatchedDimensions[0].key);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fileCompareMatchedDimKeys, !!fileCompareDataset]);

  // The sensitivity grid only makes sense for a calculated metric (it needs two
  // underlying inputs to vary independently) — keep the selection pointing at a
  // real one, and auto-pick the first available calculated metric once one exists.
  useEffect(() => {
    if (sensOutcomeKey && !activeRatios.some(r => r.key === sensOutcomeKey)) {
      setSensOutcomeKey(activeRatios[0] ? activeRatios[0].key : '');
    } else if (!sensOutcomeKey && activeRatios[0]) {
      setSensOutcomeKey(activeRatios[0].key);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeRatioKeys]);

  const genSearchedRows = useMemo(() => {
    if (!customDataset) return [];
    let rows = genFilteredRows;
    const q = search.trim().toLowerCase();
    if (q) rows = rows.filter(r => customDataset.fields.some(f => String(r[f] === undefined || r[f] === null ? '' : r[f]).toLowerCase().includes(q)));
    return [...rows].sort((a, b) => {
      let av = a[sortKey], bv = b[sortKey];
      if (av === null || av === undefined) av = '';
      if (bv === null || bv === undefined) bv = '';
      if (typeof av === 'string' || typeof bv === 'string') { av = String(av).toLowerCase(); bv = String(bv).toLowerCase(); }
      if (av < bv) return sortDir === 'asc' ? -1 : 1;
      if (av > bv) return sortDir === 'asc' ? 1 : -1;
      return 0;
    });
  }, [customDataset, genFilteredRows, search, sortKey, sortDir]);

  const tableRows = genSearchedRows;
  const pageSize = 12;
  const pageCount = Math.max(1, Math.ceil(tableRows.length / pageSize));
  const clampedPage = Math.min(page, pageCount - 1);
  const pageRows = tableRows.slice(clampedPage * pageSize, clampedPage * pageSize + pageSize);

  function toggleSort(key) {
    if (sortKey === key) setSortDir(d => d === 'asc' ? 'desc' : 'asc');
    else { setSortKey(key); setSortDir('asc'); }
    setPage(0);
  }

  // Click-to-drill: clicking a bar, donut slice, or legend row for a dimension value
  // toggles that value into the corresponding slicer, so exploring the chart directly
  // filters the whole dashboard instead of forcing a trip back to Configure analysis.
  function toggleFilterValue(dimKey, value) {
    setGenFilters(f => {
      const cur = f[dimKey] || [];
      const next = cur.includes(value) ? cur.filter(v => v !== value) : [...cur, value];
      return { ...f, [dimKey]: next };
    });
  }

  function exportCSV() {
    if (!tableRows.length) return;
    const fields = Object.keys(tableRows[0]);
    const csv = rowsToCSV(fields, tableRows);
    const base = customDataset.name.replace(/\.[a-zA-Z0-9]+$/, '').replace(/[^a-z0-9\-_]+/gi, '_');
    downloadTextFile(base + '-export.csv', csv, 'text/csv');
  }
  function exportSummaryReport() {
    if (!insightInput || !customDataset) return;
    const lines = [];
    lines.push('# ' + customDataset.name + ' — Analysis Summary');
    lines.push('');
    lines.push('Generated ' + new Date().toISOString().slice(0, 10) + ' · ' + insightInput.rowCount + ' rows analyzed');
    lines.push('');
    lines.push('## ' + insightInput.metricLabel);
    lines.push('Total: ' + fmtByFormat(insightInput.totalValue, insightInput.metricFormat));
    lines.push('');
    lines.push('## Pattern-detected insights');
    if (!ruleInsights.length) lines.push('_None detected yet — enable more dimensions or a date column._');
    ruleInsights.forEach(ins => {
      const label = (INSIGHT_STYLES[ins.type] || {}).label || ins.type;
      lines.push('- **[' + label + ']** ' + ins.title + ' — ' + ins.detail);
    });
    if (aiInsights.status === 'done' && aiInsights.data) {
      lines.push('');
      lines.push('## AI-written summary');
      [['highlights', 'Highlights'], ['painPoints', 'Pain points'], ['opportunities', 'Opportunities']].forEach(([cat, label]) => {
        lines.push('');
        lines.push('### ' + label);
        (aiInsights.data[cat] || []).forEach(item => lines.push('- **' + item.title + '** — ' + item.detail));
      });
    }
    const base = customDataset.name.replace(/\.[a-zA-Z0-9]+$/, '').replace(/[^a-z0-9\-_]+/gi, '_');
    downloadTextFile(base + '-summary.md', lines.join('\n'), 'text/markdown');
  }

  async function handleSaveConfig() {
    if (!customDataset) return;
    try {
      await saveConfig(customDataset.fields, customDataset.config);
      setSaveStatus('saved');
      setTimeout(() => setSaveStatus('idle'), 2500);
    } catch (e) {
      setSaveStatus('error');
      setTimeout(() => setSaveStatus('idle'), 2500);
    }
  }
  async function handleForgetConfig() {
    if (!customDataset) return;
    await deleteSavedConfig(customDataset.fields);
    setSavedConfigApplied(false);
  }

  function applyView(view) {
    setGenFilters(view.filters || {});
    if (view.breakdownDim) setGenBreakdownDim(view.breakdownDim);
    if (view.forecastMetric) setGenForecastMetric(view.forecastMetric);
  }
  async function saveCurrentView() {
    if (!customDataset || !newViewName.trim()) return;
    const view = { name: newViewName.trim(), filters: genFilters, breakdownDim: genBreakdownDim, forecastMetric: genForecastMetric };
    const next = [...savedViews.filter(v => v.name !== view.name), view];
    setSavedViews(next);
    setNewViewName('');
    setShowViewInput(false);
    try { await saveViewsList(customDataset.fields, next); } catch (e) { /* best-effort */ }
  }
  async function removeView(name) {
    if (!customDataset) return;
    const next = savedViews.filter(v => v.name !== name);
    setSavedViews(next);
    try { await saveViewsList(customDataset.fields, next); } catch (e) { /* best-effort */ }
  }

  async function generateAIInsights() {
    if (!insightInput || !customDataset) return;
    setAiInsights({ status: 'loading', data: null, error: null });
    if (typeof window !== 'undefined' && window.__ARTIFACT_BUILD__) {
      // A published Artifact's frame is sandboxed behind a CSP that blocks
      // direct calls to external APIs, so the live request below can't
      // succeed there. This flag is only ever set by that one build's own
      // entry point — everywhere else (this file run standalone, or inside
      // claude.ai's own artifact-creation sandbox) it's simply absent, and
      // this branch never runs.
      setAiInsights({ status: 'error', data: null, error: 'AI-written summaries need a live API call, which this published page cannot make. The pattern detection above (strengths, pain points, opportunities) runs fully offline and stays available.' });
      return;
    }
    try {
      const payload = {
        dataset: customDataset.name,
        rowCount: insightInput.rowCount,
        primaryMetric: { label: insightInput.metricLabel, total: insightInput.totalValue, format: insightInput.metricFormat },
        breakdowns: insightInput.breakdowns.map(b => ({
          dimension: b.dimLabel,
          top3: [...b.groups].sort((x, y) => y.value - x.value).slice(0, 3),
          bottom3: [...b.groups].sort((x, y) => x.value - y.value).slice(0, 3),
        })),
        trend: insightInput.trend,
        ratio: insightInput.ratio,
      };
      const prompt = 'You are a business analyst. Here is a JSON summary of already-aggregated performance data (not raw rows):\n' +
        JSON.stringify(payload) +
        '\n\nRespond with STRICT JSON only — no markdown fences, no prose outside the JSON — matching exactly this shape:\n' +
        '{"highlights":[{"title":"...","detail":"..."}],"painPoints":[{"title":"...","detail":"..."}],"opportunities":[{"title":"...","detail":"..."}]}\n' +
        'Give 3 to 4 items per array. Ground every claim in the numbers provided — name actual segments and figures from the payload. ' +
        'Infer the business domain from the field names and write in that domain’s language. Keep each detail to one or two sentences.';

      const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ model: 'claude-sonnet-4-6', max_tokens: 1000, messages: [{ role: 'user', content: prompt }] }),
      });
      if (!response.ok) throw new Error('API returned status ' + response.status);
      const data = await response.json();
      const textBlock = (data.content || []).find(b => b.type === 'text');
      if (!textBlock) throw new Error('No text in the response.');
      const cleaned = textBlock.text.replace(/```json|```/g, '').trim();
      const parsed = JSON.parse(cleaned);
      setAiInsights({ status: 'done', data: parsed, error: null });
    } catch (err) {
      setAiInsights({ status: 'error', data: null, error: (err && err.message) ? err.message : 'Something went wrong.' });
    }
  }

  async function submitQA(questionText) {
    const question = (questionText !== undefined ? questionText : qaInput).trim();
    if (!question || !insightInput || !customDataset || qaStatus === 'loading') return;
    setQaHistory(h => [...h, { role: 'user', content: question }]);
    setQaInput('');
    setQaStatus('loading');
    if (typeof window !== 'undefined' && window.__ARTIFACT_BUILD__) {
      // Same CSP limitation as generateAIInsights above — a published Artifact
      // can't reach an external API directly, so fail honestly rather than
      // hang or throw a raw network error into the chat thread.
      setQaHistory(h => [...h, { role: 'error', content: "Q&A needs a live API call, which this published page can't make. The rule-based pattern detection and every chart/table on the other tabs stay fully available." }]);
      setQaStatus('idle');
      return;
    }
    try {
      // Never send raw rows — the same aggregated-summary shape generateAIInsights
      // uses, broadened with every active metric/ratio's grand total (not just the
      // primary one) so a question about any of them can still be answered exactly
      // even though deep segment breakdowns above stay scoped to the primary metric.
      const context = {
        dataset: customDataset.name,
        rowCount: insightInput.rowCount,
        metricTotals: activeMetrics.concat(activeRatios).map(m => ({ metric: m.key, total: genKpi[m.key], format: m.format })),
        primaryMetric: { label: insightInput.metricLabel, total: insightInput.totalValue, format: insightInput.metricFormat },
        breakdownsOfPrimaryMetric: insightInput.breakdowns.map(b => ({
          dimension: b.dimLabel,
          top3: [...b.groups].sort((x, y) => y.value - x.value).slice(0, 3),
          bottom3: [...b.groups].sort((x, y) => x.value - y.value).slice(0, 3),
        })),
        trendOfPrimaryMetric: insightInput.trend,
        calculatedMetricBreakdown: insightInput.ratio,
      };
      const priorTurns = qaHistory.slice(-6).map(m => (m.role === 'user' ? 'Q: ' : 'A: ') + m.content).join('\n');
      const prompt = 'You are a business analyst answering questions about a dataset using ONLY the aggregated JSON summary below — never invent a figure that is not derivable from it:\n' +
        JSON.stringify(context) +
        (priorTurns ? '\n\nEarlier in this conversation:\n' + priorTurns : '') +
        '\n\nQuestion: ' + question +
        '\n\nAnswer in 2-4 sentences of plain text — no markdown, no JSON, no bullet points. Cite specific numbers from the summary. ' +
        "If the summary genuinely doesn't contain what's needed to answer, say so plainly instead of guessing or padding with generic advice.";

      const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ model: 'claude-sonnet-4-6', max_tokens: 500, messages: [{ role: 'user', content: prompt }] }),
      });
      if (!response.ok) throw new Error('API returned status ' + response.status);
      const data = await response.json();
      const textBlock = (data.content || []).find(b => b.type === 'text');
      if (!textBlock) throw new Error('No text in the response.');
      setQaHistory(h => [...h, { role: 'assistant', content: textBlock.text.trim() }]);
    } catch (err) {
      setQaHistory(h => [...h, { role: 'error', content: (err && err.message) ? err.message : 'Something went wrong answering that.' }]);
    } finally {
      setQaStatus('idle');
    }
  }

  function finalizeDataset(fileName, sheetsMap, activeSheetName) {
    const sheet = sheetsMap[activeSheetName];
    if (!sheet || !sheet.fields.length || !sheet.rows.length) { setUploadError('"' + activeSheetName + '" has no usable rows or columns.'); return; }
    const { fields, rows } = sheet;
    const schema = detectSchema(rows, fields);
    const numberCols = schema.filter(s => s.type === 'number');
    const categoryCols = schema.filter(s => s.type === 'category' || s.type === 'text');
    if (!numberCols.length) { setUploadError('No numeric columns detected in "' + activeSheetName + '" — nothing to total or chart. Check the file has at least one numeric column.'); return; }
    const metrics = numberCols.map((s, i) => ({ key: s.key, on: i < 6, format: guessMetricFormat(s.key), aggregation: 'sum' }));
    const dimensions = categoryCols.map((s, i) => ({ key: s.key, on: (s.unique >= 2 && s.unique <= 25 && i < 4), unique: s.unique }));
    const dateCandidate = schema.find(s => s.type === 'date');
    // Auto-suggest a "Profit" calculated metric when the shape of the data makes it
    // obvious (a revenue-like column and a cost-like column both present, and no
    // column already named profit/margin) — this is what most people mean by "add a
    // Total Profit metric" without having to build it by hand every time.
    const revenueCol = findColumnByName(numberCols, ['revenue', 'sales', 'income'], ['revenue', 'sales', 'income']);
    const costCol = findColumnByName(numberCols, ['cost', 'costs', 'expense', 'expenses', 'cogs'], ['cost', 'expense', 'cogs']);
    const hasProfitCol = numberCols.some(c => /profit|margin/i.test(c.key));
    const autoRatios = [];
    if (revenueCol && costCol && revenueCol.key !== costCol.key && !hasProfitCol) {
      autoRatios.push({ key: 'calc_profit_auto', label: 'Profit', num: revenueCol.key, den: costCol.key, op: '−', format: 'currency' });
    }
    const config = { metrics, dimensions, dateField: dateCandidate ? dateCandidate.key : '', ratios: autoRatios, budgetPairs: [] };
    setCustomDataset({ name: fileName, rows, fields, schema, config, sheets: sheetsMap, activeSheet: activeSheetName });
    setTab('dashboard');
    setConfigOpen(true);
    setGenFilters({});
    setGenBreakdownDim(dimensions.find(d => d.on) ? dimensions.find(d => d.on).key : (dimensions[0] ? dimensions[0].key : ''));
    setGenForecastMetric(metrics.find(m => m.on) ? metrics.find(m => m.on).key : (metrics[0] ? metrics[0].key : ''));
    setGenScenario('Base'); setGenCustomGrowth(0); setGenForecastHorizon(6); setGoalSeekTarget('');
    setHiddenTrendSeries({}); setTrendChartType('area');
    setSensOutcomeKey(autoRatios.length ? autoRatios[0].key : ''); setSensSpread(5); setSensSteps(5);
    setSearch(''); setPage(0); setSortKey(fields[0]); setSortDir('asc');
    setAiInsights({ status: 'idle', data: null, error: null });
    setQaHistory([]); setQaInput(''); setQaStatus('idle');
    setUploadError(null);
    setSaveStatus('idle');
    setSavedConfigApplied(false);
    setCompareOn(false); setGoalTarget(''); setAlertRule({ metricKey: '', op: 'below', value: '' });
    setSavedViews([]); setNewViewName(''); setShowViewInput(false);
    setCompareDataset(null); setCompareUploadError(null); setCompareBreakdownDim('');
    loadSavedConfig(fields).then(saved => {
      if (!saved) return;
      setCustomDataset(ds => (ds && ds.name === fileName && ds.activeSheet === activeSheetName) ? { ...ds, config: saved } : ds);
      setSavedConfigApplied(true);
      const savedDim = (saved.dimensions || []).find(d => d.on);
      const savedMetric = (saved.metrics || []).find(m => m.on);
      if (savedDim) setGenBreakdownDim(savedDim.key);
      if (savedMetric) setGenForecastMetric(savedMetric.key);
    }).catch(() => {});
    loadSavedViews(fields).then(views => setSavedViews(views || [])).catch(() => {});
  }

  function switchSheet(name) {
    if (!customDataset || !customDataset.sheets || !customDataset.sheets[name]) return;
    finalizeDataset(customDataset.name, customDataset.sheets, name);
  }

  // Shared file->rows parsing for CSV/TSV/JSON/Excel, factored out so the
  // second-file upload (Compare tab) can reuse the exact same parsing rules
  // as the primary upload instead of duplicating them. onSuccess receives
  // (fileName, sheetsMap, activeSheetName) — the same shape finalizeDataset
  // already expects, so the primary upload path is a direct passthrough.
  function parseFileGeneric(file, onSuccess, onError) {
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    if (!SUPPORTED_EXTS.includes(ext)) {
      onError('Unsupported file type ".' + ext + '". Try CSV, TSV, JSON, or Excel (.xlsx / .xls).');
      return;
    }
    const reader = new FileReader();
    reader.onerror = () => onError('Could not read that file.');

    if (ext === 'xlsx' || ext === 'xls') {
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
          const sheetsMap = {};
          wb.SheetNames.forEach(name => {
            const json = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: null, raw: true });
            if (json.length) {
               const fields = Array.from(new Set(json.flatMap(row => Object.keys(row))));
              const rows = json.map(r => {
                const o = {};
                fields.forEach(f => { let v = r[f]; if (v instanceof Date) v = v.toISOString().slice(0, 10); o[f] = v; });
                return o;
              });
              sheetsMap[name] = { fields, rows };
            }
          });
          const sheetNames = Object.keys(sheetsMap);
          if (!sheetNames.length) { onError('That workbook has no data on any sheet.'); return; }
          onSuccess(file.name, sheetsMap, sheetNames[0]);
        } catch (err) {
          onError('Could not read that Excel file: ' + (err && err.message ? err.message : 'unknown error'));
        }
      };
      reader.readAsArrayBuffer(file);
    } else if (ext === 'json') {
      reader.onload = (e) => {
        try {
          const parsed = JSON.parse(e.target.result);
          let arr = null;
          if (Array.isArray(parsed)) arr = parsed;
          else if (parsed && typeof parsed === 'object') {
            const arrKey = Object.keys(parsed).find(k => Array.isArray(parsed[k]) && parsed[k].length && typeof parsed[k][0] === 'object');
            if (arrKey) arr = parsed[arrKey];
          }
          if (!arr || !arr.length) { onError('Could not find an array of records in that JSON file — expected either a top-level array, or an object containing one.'); return; }
          const fieldSet = new Set();
          arr.forEach(row => { if (row && typeof row === 'object') Object.keys(row).forEach(k => fieldSet.add(k)); });
          const fields = Array.from(fieldSet);
          const rows = arr.map(row => { const o = {}; fields.forEach(f => { o[f] = (row && row[f] !== undefined) ? row[f] : null; }); return o; });
          onSuccess(file.name, { 'Sheet1': { fields, rows } }, 'Sheet1');
        } catch (err) {
          onError('Could not parse that JSON file: ' + (err && err.message ? err.message : 'invalid JSON'));
        }
      };
      reader.readAsText(file);
    } else {
      reader.onload = (e) => {
        const text = e.target.result;
        Papa.parse(text, {
          header: true, dynamicTyping: true, skipEmptyLines: true, delimiter: ext === 'tsv' ? '\t' : '',
          complete: (results) => {
            if (results.errors && results.errors.length) {
              const first = results.errors[0];
              onError(`CSV parse error on row ${first.row ?? '?'}: ${first.message}`);
              return;
            }
            const fields = results.meta.fields || [];
            const rows = (results.data || []).filter(r => fields.some(f => r[f] !== null && r[f] !== undefined && r[f] !== ''));
            if (!fields.length || !rows.length) { onError('Could not find columns or rows in that file — make sure it has a header row.'); return; }
            onSuccess(file.name, { 'Sheet1': { fields, rows } }, 'Sheet1');
          },
          error: (err) => onError('Could not parse that file: ' + err.message),
        });
      };
      reader.readAsText(file);
    }
  }
  function importOptions(sheet) {
    const schema = detectSchema(sheet.rows, sheet.fields);
    return {
      types: Object.fromEntries(schema.map(col => [col.key, col.type === 'number' || col.type === 'date' ? col.type : 'category'])),
      names: Object.fromEntries(sheet.fields.map(field => [field, field])),
    };
  }
  function stageImport(fileName, sheets, activeSheet) {
    setPendingImport({ name: fileName, sheets, activeSheet, ...importOptions(sheets[activeSheet]) });
    setTab('dashboard');
  }
  function confirmImport() {
    if (!pendingImport) return;
    const { name, sheets, activeSheet, types, names } = pendingImport;
    const source = sheets[activeSheet];
    const renamedFields = source.fields.map(field => names[field].trim());
    if (renamedFields.some(field => !field) || new Set(renamedFields.map(field => field.toLowerCase())).size !== renamedFields.length) return;
    const normalized = normalizeRows(source.rows, types).map(row => Object.fromEntries(source.fields.map(field => [names[field].trim(), row[field]])));
    const normalizedSheets = { ...sheets, [activeSheet]: { fields: renamedFields, rows: normalized } };
    finalizeDataset(name, normalizedSheets, activeSheet);
    setPendingImport(null);
  }
  function handleFile(file) {
    if (!file) return;
    setUploadError(null);
    parseFileGeneric(file, stageImport, setUploadError);
  }
  function onDrop(e) { e.preventDefault(); setIsDragOver(false); const f = e.dataTransfer.files && e.dataTransfer.files[0]; if (f) handleFile(f); }
  function onDragOver(e) { e.preventDefault(); setIsDragOver(true); }
  function onDragLeave() { setIsDragOver(false); }

  // Second-file upload for the Compare tab: same parser, but the result is
  // just {name, rows, fields, schema} — no config UI of its own. Comparison
  // is driven entirely by dataset A's active metrics/dimensions, matched to
  // dataset B by exact column name (see compareMatchedMetrics etc. below).
  function handleFileCompareUpload(file) {
    if (!file) return;
    setCompareUploadError(null);
    parseFileGeneric(file, (fileName, sheetsMap, activeSheetName) => {
      const { fields, rows } = sheetsMap[activeSheetName];
      const schema = detectSchema(rows, fields);
      const types = Object.fromEntries(schema.map(col => [col.key, col.type]));
      setCompareDataset({ name: fileName, rows: normalizeRows(rows, types), fields, schema });
    }, setCompareUploadError);
  }
  function removeFileCompareDataset() { setCompareDataset(null); setCompareUploadError(null); }
  function onFileCompareDrop(e) { e.preventDefault(); setCompareDragOver(false); const f = e.dataTransfer.files && e.dataTransfer.files[0]; if (f) handleFileCompareUpload(f); }
  function onFileCompareDragOver(e) { e.preventDefault(); setCompareDragOver(true); }
  function onFileCompareDragLeave() { setCompareDragOver(false); }

  function toggleMetric(key) { setCustomDataset(ds => ({ ...ds, config: { ...ds.config, metrics: ds.config.metrics.map(m => m.key === key ? { ...m, on: !m.on } : m) } })); }
  function setMetricFormat(key, format) { setCustomDataset(ds => ({ ...ds, config: { ...ds.config, metrics: ds.config.metrics.map(m => m.key === key ? { ...m, format } : m) } })); }
  function setAggregation(key, aggregation) { setCustomDataset(ds => ({ ...ds, config: { ...ds.config, metrics: ds.config.metrics.map(m => m.key === key ? { ...m, aggregation } : m) } })); }
  function toggleDimension(key) { setCustomDataset(ds => ({ ...ds, config: { ...ds.config, dimensions: ds.config.dimensions.map(d => d.key === key ? { ...d, on: !d.on } : d) } })); }
  function setDateField(key) { setCustomDataset(ds => ({ ...ds, config: { ...ds.config, dateField: key } })); }
  function addRatio() {
    const opts = customDataset.config.metrics.filter(m => m.on);
    if (opts.length < 2) return;
    setCustomDataset(ds => ({ ...ds, config: { ...ds.config, ratios: [...ds.config.ratios, { key: 'ratio_' + Date.now(), label: 'New Metric', num: opts[0].key, den: opts[1].key, op: '−', format: 'number' }] } }));
  }
  function updateRatio(key, patch) { setCustomDataset(ds => ({ ...ds, config: { ...ds.config, ratios: ds.config.ratios.map(r => r.key === key ? { ...r, ...patch } : r) } })); }
  function removeRatio(key) { setCustomDataset(ds => ({ ...ds, config: { ...ds.config, ratios: ds.config.ratios.filter(r => r.key !== key) } })); }
  function addBudgetPair() {
    const opts = customDataset.config.metrics.filter(m => m.on);
    if (opts.length < 2) return;
    setCustomDataset(ds => ({ ...ds, config: { ...ds.config, budgetPairs: [...(ds.config.budgetPairs || []), { key: 'bp_' + Date.now(), label: 'New Pair', budget: opts[0].key, actual: opts[1].key, format: opts[0].format || 'number' }] } }));
  }
  function updateBudgetPair(key, patch) { setCustomDataset(ds => ({ ...ds, config: { ...ds.config, budgetPairs: (ds.config.budgetPairs || []).map(p => p.key === key ? { ...p, ...patch } : p) } })); }
  function removeBudgetPair(key) { setCustomDataset(ds => ({ ...ds, config: { ...ds.config, budgetPairs: (ds.config.budgetPairs || []).filter(p => p.key !== key) } })); }
  function removeDataset() {
    setCustomDataset(null);
    setSearch(''); setPage(0); setSortKey(''); setSortDir('asc');
    setAiInsights({ status: 'idle', data: null, error: null });
    setQaHistory([]); setQaInput(''); setQaStatus('idle');
    setCompareOn(false); setGoalTarget(''); setAlertRule({ metricKey: '', op: 'below', value: '' });
    setSavedViews([]); setNewViewName(''); setShowViewInput(false);
    setGenForecastHorizon(6); setGoalSeekTarget(''); setHiddenTrendSeries({}); setTrendChartType('area');
    setSensOutcomeKey(''); setSensSpread(5); setSensSteps(5);
    setPeriodCompareMode('prior'); setPeriodCompareMetric('');
    setCompareDataset(null); setCompareUploadError(null); setCompareBreakdownDim('');
  }

  const navItems = [
    { key: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
    { key: 'breakdown', label: 'Breakdown', icon: Layers },
    { key: 'trends', label: 'Trends', icon: TrendingUp },
    { key: 'forecast', label: 'Forecast', icon: SlidersHorizontal },
    { key: 'insights', label: 'Insights', icon: Sparkles },
    { key: 'data', label: 'Data', icon: Table2 },
    { key: 'compare', label: 'Compare', icon: GitCompare },
  ];

  return (
    <div onDrop={onDrop} onDragOver={onDragOver} onDragLeave={onDragLeave}
      style={{ position: 'relative', minHeight: '100%', background: COLORS.bg, fontFamily: SANS, display: 'flex' }}>
      <style>{'@media print { .no-print { display: none !important; } }'}</style>
      {isDragOver ? (
        <div style={{ position: 'absolute', inset: 0, zIndex: 50, background: 'rgba(15,20,36,0.90)', display: 'flex', alignItems: 'center', justifyContent: 'center', border: `3px dashed ${COLORS.accent}`, borderRadius: 8 }}>
          <div style={{ textAlign: 'center', color: '#FFFFFF' }}>
            <UploadCloud size={32} color={COLORS.cyan} style={{ margin: '0 auto 10px auto' }} />
            <div style={{ fontFamily: SANS, fontWeight: 700, fontSize: 16 }}>Drop your file to analyze it</div>
            <div style={{ fontFamily: SANS, fontSize: 12, color: '#9098B8', marginTop: 6 }}>CSV · TSV · JSON · Excel (.xlsx / .xls)</div>
          </div>
        </div>
      ) : null}

      <input ref={fileInputRef} type="file" accept=".csv,.tsv,.txt,.json,.xlsx,.xls" style={{ display: 'none' }}
        onChange={e => { handleFile(e.target.files[0]); e.target.value = ''; }} />
      <input ref={fileCompareInputRef} type="file" accept=".csv,.tsv,.txt,.json,.xlsx,.xls" style={{ display: 'none' }}
        onChange={e => { handleFileCompareUpload(e.target.files[0]); e.target.value = ''; }} />

      {/* Sidebar (desktop) */}
      <div className="hidden md:flex no-print" style={{ width: 216, flexShrink: 0, background: COLORS.ink, flexDirection: 'column', padding: '20px 12px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '4px 10px 20px 10px' }}>
          <BarChart3 size={20} color={COLORS.cyan} strokeWidth={2.2} />
          <span style={{ color: '#FFFFFF', fontWeight: 700, fontSize: 14.5, letterSpacing: '0.01em' }}>MBU Finalysis</span>
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
          {navItems.map(it => <NavItem key={it.key} icon={it.icon} label={it.label} active={tab === it.key} onClick={() => setTab(it.key)} vertical />)}
        </div>
        <button onClick={() => fileInputRef.current && fileInputRef.current.click()}
          style={{ display: 'flex', alignItems: 'center', gap: 8, marginTop: 18, padding: '9px 14px', borderRadius: 8, border: `1px dashed #3A4362`, background: 'transparent', color: '#9098B8', fontFamily: SANS, fontSize: 12.5, fontWeight: 600, cursor: 'pointer' }}>
          <Upload size={15} /> Upload data
        </button>
        <div style={{ fontFamily: SANS, fontSize: 10, color: '#5C6584', padding: '6px 10px 0 10px' }}>CSV · TSV · JSON · Excel</div>
        <div style={{ marginTop: 'auto', padding: '14px 10px 4px 10px', color: '#7078A0', fontSize: 10.5, lineHeight: 1.5 }}>
          {customDataset ? customDataset.rows.length + ' rows loaded' : 'No data loaded yet'}
        </div>
      </div>

      <div style={{ flex: 1, minWidth: 0, display: 'flex', flexDirection: 'column' }}>
        {/* Top bar */}
        <div className="no-print" style={{ background: COLORS.ink, padding: '14px 20px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 10 }}>
          <div className="flex md:hidden" style={{ alignItems: 'center', gap: 8 }}>
            <BarChart3 size={18} color={COLORS.cyan} />
            <span style={{ color: '#FFFFFF', fontWeight: 700, fontSize: 14 }}>MBU Finalysis</span>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
            {customDataset ? (
              <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                <span style={{ display: 'flex', alignItems: 'center', gap: 6, fontFamily: SANS, fontSize: 12, fontWeight: 600, padding: '6px 12px', borderRadius: 999, border: `1px solid ${COLORS.accent}`, background: COLORS.accent, color: '#FFFFFF', maxWidth: 220 }}>
                  <FileSpreadsheet size={12} />
                  <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{customDataset.name}</span>
                </span>
                <button onClick={removeDataset} title="Remove this file" style={{ background: 'transparent', border: 'none', cursor: 'pointer', color: '#9098B8' }}><X size={14} /></button>
              </span>
            ) : (
              <button onClick={() => fileInputRef.current && fileInputRef.current.click()}
                style={{ display: 'flex', alignItems: 'center', gap: 6, fontFamily: SANS, fontSize: 12, fontWeight: 600, padding: '6px 12px', borderRadius: 999, cursor: 'pointer', border: '1px dashed #3A4362', background: 'transparent', color: '#9098B8' }}>
                <Upload size={12} /> Upload your data
              </button>
            )}
            <button onClick={() => window.print()} title="Print or save this view as PDF"
              style={{ display: 'flex', alignItems: 'center', gap: 6, fontFamily: SANS, fontSize: 12, fontWeight: 600, padding: '6px 10px', borderRadius: 999, cursor: 'pointer', border: '1px solid #3A4362', background: 'transparent', color: '#9098B8' }}>
              <Printer size={13} />
            </button>
          </div>
        </div>
        {/* Mobile tab strip */}
        <div className="flex md:hidden no-print" style={{ background: COLORS.inkLight, padding: '8px 12px', gap: 8, overflowX: 'auto' }}>
          {navItems.map(it => <NavItem key={it.key} icon={it.icon} label={it.label} active={tab === it.key} onClick={() => setTab(it.key)} />)}
        </div>

        {uploadError ? (
          <div style={{ margin: '14px 24px 0 24px', display: 'flex', alignItems: 'center', gap: 8, background: '#FDEEF0', border: `1px solid ${COLORS.rose}`, color: '#8A1F30', borderRadius: 10, padding: '10px 14px', fontFamily: SANS, fontSize: 12.5 }}>
            <AlertCircle size={15} /> {uploadError}
          </div>
        ) : null}

        <div style={{ padding: '22px 24px 40px 24px', overflowX: 'hidden' }}>

          {tab === 'dashboard' && (
            <div>
              {pendingImport ? <ImportPreview pending={pendingImport}
                onSheet={name => setPendingImport(p => ({ ...p, activeSheet: name, ...importOptions(p.sheets[name]) }))}
                onType={(field, type) => setPendingImport(p => ({ ...p, types: { ...p.types, [field]: type } }))}
                onName={(field, name) => setPendingImport(p => ({ ...p, names: { ...p.names, [field]: name } }))}
                onConfirm={confirmImport} onCancel={() => setPendingImport(null)} /> : !customDataset ? (
                <div>
                  <SectionHeader title="Welcome to MBU Finalysis" subtitle="Upload any spreadsheet or export and get a full financial analysis workspace in seconds." />
                  <Panel style={{ textAlign: 'center', padding: '56px 24px', background: `linear-gradient(180deg, ${COLORS.ink} 0%, ${COLORS.inkLight} 100%)`, border: 'none' }}>
                    <div style={{ width: 60, height: 60, borderRadius: 16, background: 'rgba(88,80,236,0.18)', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 18px auto' }}>
                      <UploadCloud size={28} color={COLORS.cyan} />
                    </div>
                    <div style={{ fontFamily: SANS, fontWeight: 700, fontSize: 19, color: '#FFFFFF', marginBottom: 8 }}>Upload your data to get started</div>
                    <div style={{ fontFamily: SANS, fontSize: 13, color: '#9098B8', maxWidth: 440, margin: '0 auto 22px auto', lineHeight: 1.5 }}>
                      Upload a file, review its columns, and build filterable dashboards, forecasts, and offline insights from your data.
                    </div>
                    <button onClick={() => fileInputRef.current && fileInputRef.current.click()}
                      style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontFamily: SANS, fontSize: 13.5, fontWeight: 700, color: COLORS.ink, background: `linear-gradient(135deg, ${COLORS.cyan}, ${COLORS.accent})`, border: 'none', borderRadius: 10, padding: '12px 22px', cursor: 'pointer' }}>
                      <Upload size={15} /> Choose a file
                    </button>
                    <div style={{ fontFamily: SANS, fontSize: 11, color: '#7078A0', marginTop: 14 }}>CSV · TSV · JSON · Excel (.xlsx / .xls) — or drag a file anywhere on this page</div>
                  </Panel>
                  <div className="grid grid-cols-1 sm:grid-cols-3" style={{ gap: 14, marginTop: 18 }}>
                    {[
                      { icon: Sparkles, title: 'Auto-detects everything', body: 'Numeric metrics, category dimensions, and date columns — no setup required.' },
                      { icon: TrendingUp, title: 'Forecasts from your history', body: 'Growth scenarios estimated from your own trend, not generic assumptions.' },
                      { icon: Zap, title: 'Finds what matters', body: 'Correlations, outliers, and threshold alerts surface automatically.' },
                    ].map((f, i) => (
                      <Panel key={i}>
                        <f.icon size={18} color={COLORS.accent} style={{ marginBottom: 10 }} />
                        <div style={{ fontFamily: SANS, fontWeight: 700, fontSize: 13, color: COLORS.ink, marginBottom: 4 }}>{f.title}</div>
                        <div style={{ fontFamily: SANS, fontSize: 12, color: COLORS.muted, lineHeight: 1.4 }}>{f.body}</div>
                      </Panel>
                    ))}
                  </div>
                </div>
              ) : (
                <>
                  <SectionHeader title="Portfolio Dashboard" subtitle={customDataset.name + ' — ' + customDataset.rows.length + ' rows, auto-detected on upload. Adjust below any time.'} />
                  <ConfigPanel dataset={customDataset} open={configOpen} onToggleOpen={() => setConfigOpen(o => !o)}
                    onSwitchSheet={switchSheet}
                    onToggleMetric={toggleMetric} onSetMetricFormat={setMetricFormat} onSetAggregation={setAggregation} onToggleDimension={toggleDimension}
                    onSetDateField={setDateField} onAddRatio={addRatio} onUpdateRatio={updateRatio} onRemoveRatio={removeRatio}
                    onSaveConfig={handleSaveConfig} onForgetConfig={handleForgetConfig} saveStatus={saveStatus} savedConfigApplied={savedConfigApplied}
                    onAddBudgetPair={addBudgetPair} onUpdateBudgetPair={updateBudgetPair} onRemoveBudgetPair={removeBudgetPair} />
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', marginBottom: 14 }}>
                    <span style={{ fontFamily: SANS, fontSize: 11, fontWeight: 700, color: COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Views:</span>
                    {savedViews.map(v => (
                      <span key={v.name} style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                        <button onClick={() => applyView(v)}
                          style={{ fontFamily: SANS, fontSize: 12, fontWeight: 600, padding: '5px 10px', borderRadius: 999, cursor: 'pointer', border: `1px solid ${COLORS.border}`, background: '#FFFFFF', color: COLORS.ink }}>{v.name}</button>
                        <button onClick={() => removeView(v.name)} style={{ background: 'transparent', border: 'none', cursor: 'pointer', color: COLORS.muted, padding: 0 }}><X size={11} /></button>
                      </span>
                    ))}
                    {showViewInput ? (
                      <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
                        <input value={newViewName} onChange={e => setNewViewName(e.target.value)} placeholder="View name…" autoFocus
                          onKeyDown={e => { if (e.key === 'Enter') saveCurrentView(); if (e.key === 'Escape') { setShowViewInput(false); setNewViewName(''); } }}
                          style={{ fontFamily: SANS, fontSize: 12, border: `1px solid ${COLORS.border}`, borderRadius: 6, padding: '5px 8px', width: 130 }} />
                        <button onClick={saveCurrentView} disabled={!newViewName.trim()}
                          style={{ fontFamily: SANS, fontSize: 12, fontWeight: 600, color: '#FFFFFF', background: COLORS.accent, border: 'none', borderRadius: 6, padding: '5px 10px', cursor: 'pointer' }}>Save</button>
                      </span>
                    ) : (
                      <button onClick={() => setShowViewInput(true)}
                        style={{ display: 'flex', alignItems: 'center', gap: 4, fontFamily: SANS, fontSize: 12, fontWeight: 600, color: COLORS.ink, background: '#FFFFFF', border: `1px solid ${COLORS.border}`, borderRadius: 999, padding: '5px 10px', cursor: 'pointer' }}>
                        <Plus size={11} /> Save current view
                      </button>
                    )}
                  </div>
                  {activeDimensions.length > 0 ? (
                    <Panel style={{ marginBottom: 20 }}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10, flexWrap: 'wrap', marginBottom: 14 }}>
                        <div style={{ fontFamily: SANS, fontSize: 12.5, fontWeight: 700, color: COLORS.ink }}>Slicers</div>
                        <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
                          <span style={{ fontFamily: MONO, fontSize: 12, color: COLORS.muted }}>{genFilteredRows.length} of {customDataset.rows.length} rows</span>
                          <button onClick={() => setGenFilters({})}
                            style={{ display: 'flex', alignItems: 'center', gap: 6, fontFamily: SANS, fontSize: 12, fontWeight: 600, color: COLORS.muted, background: 'transparent', border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: '6px 11px', cursor: 'pointer' }}>
                            <RotateCcw size={12} /> Reset
                          </button>
                        </div>
                      </div>
                      <div style={{ display: 'flex', gap: 28, flexWrap: 'wrap' }}>
                        {activeDimensions.map(d => {
                          const opts = Array.from(new Set(customDataset.rows.map(r => { const raw = r[d.key]; return (raw === null || raw === undefined || raw === '') ? '(blank)' : String(raw); }))).sort();
                          return <Slicer key={d.key} label={d.key} options={opts} selected={genFilters[d.key]} onChange={v => setGenFilters(f => ({ ...f, [d.key]: v }))} />;
                        })}
                      </div>
                    </Panel>
                  ) : null}
                  {genBudgetVariance.length > 0 ? (
                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3" style={{ gap: 12, marginBottom: 20 }}>
                      {genBudgetVariance.map(bv => (
                        <Panel key={bv.key}>
                          <div style={{ fontFamily: SANS, fontSize: 10.5, fontWeight: 600, color: COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: 6 }}>{bv.label}</div>
                          <div style={{ display: 'flex', alignItems: 'baseline', gap: 10 }}>
                            <span style={{ fontFamily: MONO, fontSize: 20, fontWeight: 700, color: COLORS.ink }}>{fmtByFormat(bv.actualVal, bv.format)}</span>
                            <DeltaBadge pct={bv.pct} size={12} />
                          </div>
                          <div style={{ fontFamily: SANS, fontSize: 11, color: COLORS.muted, marginTop: 4 }}>vs budget {fmtByFormat(bv.budgetVal, bv.format)} · variance {fmtByFormat(bv.variance, bv.format)}</div>
                        </Panel>
                      ))}
                    </div>
                  ) : null}
                  <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4" style={{ gap: 12, marginBottom: 24 }}>
                    {activeMetrics.map((m, i) => (
                      <KpiCard key={m.key} label={kpiLabel(m.key)} value={fmtByFormat(genKpi[m.key], m.format)} accent={CHART_PALETTE[i % CHART_PALETTE.length]}
                        trend={dateField && genTrend.length > 1 ? genTrend : null} trendKey={m.key}
                        delta={dateField ? genPeriodDeltaData[m.key] : null} />
                    ))}
                    {activeRatios.map((rt, i) => (
                      <KpiCard key={rt.key} label={rt.label} value={fmtByFormat(genKpi[rt.key], rt.format)} accent={CHART_PALETTE[(activeMetrics.length + i) % CHART_PALETTE.length]}
                        trend={dateField && genTrend.length > 1 ? genTrend : null} trendKey={rt.key}
                        delta={dateField ? genPeriodDeltaData[rt.key] : null} />
                    ))}
                  </div>
                  {activeMetrics.length === 0 ? (
                    <EmptyState icon={AlertCircle} title="No metrics selected" body="Turn on at least one metric in Configure analysis above to see KPI cards and charts." />
                  ) : (
                    <div className="grid grid-cols-1 lg:grid-cols-2" style={{ gap: 16 }}>
                      {dateField && genTrend.length > 1 ? (
                        <Panel>
                          <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 600, color: COLORS.ink, marginBottom: 12 }}>{activeMetrics[0].key} Trend — Full Dataset</div>
                          <div style={{ height: 240 }}>
                            <ResponsiveContainer width="100%" height="100%">
                              <AreaChart data={genTrend} margin={{ top: 4, right: 8, left: -12, bottom: 0 }}>
                                <defs>
                                  <linearGradient id="dashTrendGrad" x1="0" y1="0" x2="0" y2="1">
                                    <stop offset="5%" stopColor={COLORS.accent} stopOpacity={0.32} />
                                    <stop offset="95%" stopColor={COLORS.accent} stopOpacity={0} />
                                  </linearGradient>
                                </defs>
                                <CartesianGrid stroke={COLORS.border} vertical={false} />
                                <XAxis dataKey="label" tick={{ fontSize: 10, fill: COLORS.muted, fontFamily: SANS }} />
                                <YAxis tick={{ fontSize: 10.5, fill: COLORS.muted, fontFamily: SANS }} />
                                <Tooltip content={<RichTooltip formatMap={{ [activeMetrics[0].key]: activeMetrics[0].format }} />} cursor={{ stroke: COLORS.accent, strokeWidth: 1, strokeDasharray: '3 3' }} />
                                <Area type="monotone" dataKey={activeMetrics[0].key} name={activeMetrics[0].key} stroke={COLORS.accent} strokeWidth={2.5} fill="url(#dashTrendGrad)"
                                  activeDot={{ r: 5, stroke: '#FFFFFF', strokeWidth: 2 }} />
                              </AreaChart>
                            </ResponsiveContainer>
                          </div>
                        </Panel>
                      ) : (
                        <EmptyState icon={TrendingUp} title="No trend chart yet" body="Set a valid date column in Configure analysis to chart this metric over time." />
                      )}
                      {activeDimensions[0] && genFirstDimAgg.length > 0 ? (
                        <DonutBreakdown data={genFirstDimAgg} valueKey={activeMetrics[0].key} title={activeMetrics[0].key + ' by ' + activeDimensions[0].key}
                          valueFormat={activeMetrics[0].format} onSliceClick={(name) => toggleFilterValue(activeDimensions[0].key, name)} selectedValues={genFilters[activeDimensions[0].key]}
                          centerLabel="Total" centerValue={fmtByFormat(genFirstDimAgg.reduce((s, d) => s + (d[activeMetrics[0].key] || 0), 0), activeMetrics[0].format)} />
                      ) : (
                        <EmptyState icon={Layers} title="No breakdown chart yet" body="Enable a dimension in Configure analysis to compare segments." />
                      )}
                    </div>
                  )}

                  {insightInput ? (
                    <div className="grid grid-cols-1 lg:grid-cols-2" style={{ gap: 16, marginTop: 16 }}>
                      <Panel>
                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 8, marginBottom: 8 }}>
                          <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 600, color: COLORS.ink }}>Period over period</div>
                          <div style={{ display: 'flex', gap: 4 }}>
                            <button onClick={() => setPeriodCompareMode('prior')}
                              style={{ fontFamily: SANS, fontSize: 11, fontWeight: 600, padding: '4px 9px', borderRadius: 999, cursor: 'pointer', border: `1px solid ${periodCompareMode === 'prior' ? COLORS.accent : COLORS.border}`, background: periodCompareMode === 'prior' ? COLORS.accent : '#FFFFFF', color: periodCompareMode === 'prior' ? '#FFFFFF' : COLORS.muted }}>
                              Prior period
                            </button>
                            {genTrend.length >= 13 ? (
                              <button onClick={() => setPeriodCompareMode('yoy')}
                                style={{ fontFamily: SANS, fontSize: 11, fontWeight: 600, padding: '4px 9px', borderRadius: 999, cursor: 'pointer', border: `1px solid ${periodCompareMode === 'yoy' ? COLORS.accent : COLORS.border}`, background: periodCompareMode === 'yoy' ? COLORS.accent : '#FFFFFF', color: periodCompareMode === 'yoy' ? '#FFFFFF' : COLORS.muted }}>
                                Year over year
                              </button>
                            ) : null}
                          </div>
                        </div>
                        {compareMetricList.length > 1 ? (
                          <select value={periodCompareMetric} onChange={e => setPeriodCompareMetric(e.target.value)}
                            style={{ fontFamily: SANS, fontSize: 12.5, color: COLORS.ink, background: '#FFFFFF', border: `1px solid ${COLORS.border}`, borderRadius: 7, padding: '5px 8px', marginBottom: 6, cursor: 'pointer' }}>
                            {compareMetricList.map(m => <option key={m.key} value={m.key}>{m.label}</option>)}
                          </select>
                        ) : null}
                        {periodCompareDelta ? (
                          <div>
                            <div style={{ display: 'flex', alignItems: 'baseline', gap: 10, marginTop: 4 }}>
                              <span style={{ fontFamily: MONO, fontSize: 22, fontWeight: 700, color: COLORS.ink }}>{fmtByFormat(periodCompareDelta.cur, periodCompareFormat)}</span>
                              <DeltaBadge pct={periodCompareDelta.pct} size={13} />
                            </div>
                            <div style={{ fontFamily: SANS, fontSize: 11.5, color: COLORS.muted, marginTop: 4 }}>{periodCompareLabel} in {periodCompareDelta.curLabel}, vs {fmtByFormat(periodCompareDelta.prev, periodCompareFormat)} in {periodCompareDelta.prevLabel}</div>
                          </div>
                        ) : (
                          <div style={{ fontFamily: SANS, fontSize: 12, color: COLORS.muted, marginTop: 8 }}>
                            {periodCompareMode === 'yoy' ? 'Needs at least 13 months of date-based history to compare year over year.' : 'Needs at least two periods of date-based history to compare.'}
                          </div>
                        )}
                      </Panel>
                      <Panel>
                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
                          <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 600, color: COLORS.ink }}>Goal tracking — {goalMetricLabel}</div>
                          <input value={goalTarget} onChange={e => setGoalTarget(e.target.value)} placeholder="Set a target…" type="number"
                            style={{ width: 120, fontFamily: MONO, fontSize: 12, padding: '5px 8px', borderRadius: 6, border: `1px solid ${COLORS.border}` }} />
                        </div>
                        {goalTarget !== '' && !isNaN(Number(goalTarget)) && Number(goalTarget) !== 0 ? (
                          <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
                            <RadialGoal pct={goalPct} color={goalPct >= 1 ? COLORS.emerald : COLORS.accent} label={goalPct >= 1 ? 'Reached' : 'to goal'} />
                            <div style={{ fontFamily: SANS, fontSize: 11.5, color: COLORS.muted, lineHeight: 1.6 }}>
                              <div><strong style={{ color: COLORS.ink }}>{fmtByFormat(goalCurrentValue, goalMetricFormat)}</strong> so far</div>
                              <div>of <strong style={{ color: COLORS.ink }}>{fmtByFormat(Number(goalTarget), goalMetricFormat)}</strong> target</div>
                            </div>
                          </div>
                        ) : (
                          <div style={{ fontFamily: SANS, fontSize: 12, color: COLORS.muted }}>Set a target above to track progress against it.</div>
                        )}
                      </Panel>
                    </div>
                  ) : null}
                </>
              )}
            </div>
          )}

          {tab === 'breakdown' && (
            <div>
              {!customDataset ? (
                <EmptyState icon={Layers} title="No data loaded" body="Upload a file from the Dashboard tab to break your data down by segment." />
              ) : activeDimensions.length === 0 ? (
                <EmptyState icon={Layers} title="No dimensions enabled" body="Turn on at least one category column in Configure analysis (Dashboard tab) to break your data down by segment." />
              ) : (
                <>
                  <SectionHeader title="Performance Breakdown" subtitle={'Pick a dimension to segment by.' + (Object.values(genFilters).some(v => v && v.length) ? ' Reflects the active filters (' + genFilteredRows.length + ' of ' + customDataset.rows.length + ' rows).' : ' Reflects all ' + customDataset.rows.length + ' rows — click a bar or slice below (or add slicers on the Dashboard tab) to narrow it.')} />
                  <ActiveFiltersBar filters={genFilters} onRemove={toggleFilterValue} onClearAll={() => setGenFilters({})} />
                  <div style={{ display: 'flex', gap: 8, marginBottom: 18, flexWrap: 'wrap' }}>
                    {activeDimensions.map(d => (
                      <button key={d.key} onClick={() => setGenBreakdownDim(d.key)}
                        style={{ fontFamily: SANS, fontSize: 12.5, fontWeight: 600, padding: '8px 14px', borderRadius: 999, cursor: 'pointer', border: `1px solid ${genBreakdownDim === d.key ? COLORS.ink : COLORS.border}`, background: genBreakdownDim === d.key ? COLORS.ink : '#FFFFFF', color: genBreakdownDim === d.key ? '#FFFFFF' : COLORS.ink }}>{d.key}</button>
                    ))}
                  </div>
                  {activeMetrics.length === 0 ? (
                    <EmptyState icon={AlertCircle} title="No metrics selected" body="Turn on at least one metric in Configure analysis to chart and tabulate this breakdown." />
                  ) : (
                    <>
                      <div className="grid grid-cols-1 lg:grid-cols-2" style={{ gap: 16, marginBottom: 16 }}>
                        <Panel>
                          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 12, gap: 8 }}>
                            <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 600, color: COLORS.ink }}>{activeMetrics[0].key} by {genBreakdownDim} — ranked</div>
                            <div style={{ fontFamily: SANS, fontSize: 10.5, color: COLORS.muted }}>Click a bar to filter</div>
                          </div>
                          <div style={{ height: Math.max(220, genBreakdown.length * 34) }}>
                            <ResponsiveContainer width="100%" height="100%">
                              {(() => {
                                const ranked = [...genBreakdown].sort((a, b) => b[activeMetrics[0].key] - a[activeMetrics[0].key]);
                                const sel = genFilters[genBreakdownDim] || [];
                                return (
                                  <BarChart data={ranked} layout="vertical" margin={{ top: 4, right: 24, left: 8, bottom: 0 }}>
                                    <defs>
                                      <linearGradient id="hbarGrad" x1="0" y1="0" x2="1" y2="0">
                                        <stop offset="0%" stopColor={COLORS.accent} stopOpacity={0.55} />
                                        <stop offset="100%" stopColor={COLORS.accent} stopOpacity={1} />
                                      </linearGradient>
                                    </defs>
                                    <CartesianGrid stroke={COLORS.border} horizontal={false} />
                                    <XAxis type="number" tick={{ fontSize: 10, fill: COLORS.muted, fontFamily: SANS }} />
                                    <YAxis type="category" dataKey="name" width={110} tick={{ fontSize: 10.5, fill: COLORS.ink, fontFamily: SANS }} />
                                    <Tooltip content={<RichTooltip formatMap={{ [activeMetrics[0].key]: activeMetrics[0].format }} />} cursor={{ fill: COLORS.bg }} />
                                    <Bar dataKey={activeMetrics[0].key} name={activeMetrics[0].key} radius={[0, 6, 6, 0]}
                                      onClick={(d) => toggleFilterValue(genBreakdownDim, d.name)} style={{ cursor: 'pointer' }}>
                                      {ranked.map((entry) => (
                                        <Cell key={entry.name} fill="url(#hbarGrad)" opacity={sel.length > 0 && !sel.includes(entry.name) ? 0.35 : 1} />
                                      ))}
                                    </Bar>
                                  </BarChart>
                                );
                              })()}
                            </ResponsiveContainer>
                          </div>
                        </Panel>
                        <DonutBreakdown data={genBreakdown} valueKey={activeMetrics[0].key} title={activeMetrics[0].key + ' share by ' + genBreakdownDim}
                          valueFormat={activeMetrics[0].format} onSliceClick={(name) => toggleFilterValue(genBreakdownDim, name)} selectedValues={genFilters[genBreakdownDim]}
                          centerLabel="Total" centerValue={fmtByFormat(genBreakdown.reduce((s, d) => s + (d[activeMetrics[0].key] || 0), 0), activeMetrics[0].format)} />
                      </div>
                      <Panel style={{ overflowX: 'auto' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', fontFamily: SANS, fontSize: 12.5 }}>
                          <thead>
                            <tr style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                              <th style={{ textAlign: 'left', padding: '8px 10px', color: COLORS.muted, fontWeight: 600, fontSize: 11, textTransform: 'uppercase' }}>{genBreakdownDim}</th>
                              {activeMetrics.map(m => <th key={m.key} style={{ textAlign: 'right', padding: '8px 10px', color: COLORS.muted, fontWeight: 600, fontSize: 11, textTransform: 'uppercase' }}>{m.key}</th>)}
                              {activeRatios.map(rt => <th key={rt.key} style={{ textAlign: 'right', padding: '8px 10px', color: COLORS.muted, fontWeight: 600, fontSize: 11, textTransform: 'uppercase' }}>{rt.label}</th>)}
                              <th style={{ textAlign: 'right', padding: '8px 10px', color: COLORS.muted, fontWeight: 600, fontSize: 11, textTransform: 'uppercase' }}>Rows</th>
                            </tr>
                          </thead>
                          <tbody>
                            {genBreakdown.map(row => (
                              <tr key={row.name} style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                                <td style={{ padding: '9px 10px', fontWeight: 600, color: COLORS.ink }}>{row.name}</td>
                                {activeMetrics.map(m => <td key={m.key} style={{ padding: '9px 10px', textAlign: 'right', fontFamily: MONO }}>{fmtByFormat(row[m.key], m.format)}</td>)}
                                {activeRatios.map(rt => <td key={rt.key} style={{ padding: '9px 10px', textAlign: 'right', fontFamily: MONO }}>{fmtByFormat(row[rt.key], rt.format)}</td>)}
                                <td style={{ padding: '9px 10px', textAlign: 'right', fontFamily: MONO, color: COLORS.muted }}>{row.count}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </Panel>
                    </>
                  )}
                </>
              )}

              {customDataset && compareDimOptions.length >= 2 ? (
                <Panel style={{ marginTop: 16 }}>
                  <button onClick={() => setCompareOn(o => !o)} style={{ display: 'flex', alignItems: 'center', gap: 8, background: 'transparent', border: 'none', cursor: 'pointer', padding: 0, marginBottom: compareOn ? 14 : 0 }}>
                    <GitCompare size={15} color={COLORS.ink} />
                    <span style={{ fontFamily: SANS, fontWeight: 700, fontSize: 13.5, color: COLORS.ink }}>Compare two segments head-to-head</span>
                    {compareOn ? <ChevronUp size={16} color={COLORS.muted} /> : <ChevronDown size={16} color={COLORS.muted} />}
                  </button>
                  {compareOn ? (
                    <div>
                      <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap', marginBottom: 16 }}>
                        <Select label="Segment A" value={compareA} onChange={setCompareA} options={compareDimOptions.map(s => s.name)} />
                        <Select label="Segment B" value={compareB} onChange={setCompareB} options={compareDimOptions.map(s => s.name)} />
                      </div>
                      {compareSegA && compareSegB ? (
                        <div style={{ overflowX: 'auto' }}>
                          <table style={{ width: '100%', borderCollapse: 'collapse', fontFamily: SANS, fontSize: 12.5 }}>
                            <thead>
                              <tr style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                                <th style={{ textAlign: 'left', padding: '8px 10px', color: COLORS.muted, fontWeight: 600, fontSize: 11, textTransform: 'uppercase' }}>Metric</th>
                                <th style={{ textAlign: 'right', padding: '8px 10px', color: COLORS.muted, fontWeight: 600, fontSize: 11, textTransform: 'uppercase' }}>{compareA}</th>
                                <th style={{ textAlign: 'right', padding: '8px 10px', color: COLORS.muted, fontWeight: 600, fontSize: 11, textTransform: 'uppercase' }}>{compareB}</th>
                                <th style={{ textAlign: 'right', padding: '8px 10px', color: COLORS.muted, fontWeight: 600, fontSize: 11, textTransform: 'uppercase' }}>Δ</th>
                              </tr>
                            </thead>
                            <tbody>
                              {compareMetricList.map(m => {
                                const va = compareSegA[m.key], vb = compareSegB[m.key];
                                const pct = (typeof va === 'number' && va !== 0) ? (vb - va) / Math.abs(va) : 0;
                                return (
                                  <tr key={m.key} style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                                    <td style={{ padding: '8px 10px', fontWeight: 600, color: COLORS.ink }}>{m.label}</td>
                                    <td style={{ padding: '8px 10px', textAlign: 'right', fontFamily: MONO }}>{fmtByFormat(va, m.format)}</td>
                                    <td style={{ padding: '8px 10px', textAlign: 'right', fontFamily: MONO }}>{fmtByFormat(vb, m.format)}</td>
                                    <td style={{ padding: '8px 10px', textAlign: 'right' }}><DeltaBadge pct={pct} size={11.5} /></td>
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        </div>
                      ) : null}
                    </div>
                  ) : null}
                </Panel>
              ) : null}
            </div>
          )}

          {tab === 'insights' && (
            <div>
              <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap' }}>
                <SectionHeader title="Insights" subtitle="Automatic pattern detection, plus an optional AI-written summary of highlights, pain points, and opportunities." />
                {insightInput ? (
                  <button onClick={exportSummaryReport}
                    style={{ display: 'flex', alignItems: 'center', gap: 6, fontFamily: SANS, fontSize: 12, fontWeight: 600, color: COLORS.ink, background: '#FFFFFF', border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: '7px 12px', cursor: 'pointer', flexShrink: 0, marginTop: 2 }}>
                    <FileDown size={13} /> Export summary report
                  </button>
                ) : null}
              </div>
              {!customDataset ? (
                <EmptyState icon={Sparkles} title="No data loaded" body="Upload a file from the Dashboard tab to generate insights." />
              ) : (
                <>
                  <ActiveFiltersBar filters={genFilters} onRemove={toggleFilterValue} onClearAll={() => setGenFilters({})} />
                  <Panel style={{ marginBottom: 24, border: `1px solid ${COLORS.accent}` }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
                      <MessageCircleQuestion size={16} color={COLORS.accent} />
                      <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 700, color: COLORS.ink }}>Ask a question</div>
                    </div>
                    <div style={{ fontFamily: SANS, fontSize: 12, color: COLORS.muted, marginBottom: 14 }}>
                      Type a plain-English question about this data — answers are grounded in the aggregated totals and breakdowns below, never raw rows.
                    </div>
                    {qaHistory.length > 0 ? (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: 10, marginBottom: 14, maxHeight: 360, overflowY: 'auto', padding: '2px 2px' }}>
                        {qaHistory.map((m, i) => (
                          <div key={i} style={{ display: 'flex', justifyContent: m.role === 'user' ? 'flex-end' : 'flex-start' }}>
                            <div style={{
                              maxWidth: '80%', fontFamily: SANS, fontSize: 12.5, lineHeight: 1.5, padding: '9px 13px', borderRadius: 12,
                              background: m.role === 'user' ? COLORS.ink : (m.role === 'error' ? '#FDEEF0' : COLORS.bg),
                              color: m.role === 'user' ? '#FFFFFF' : (m.role === 'error' ? COLORS.rose : COLORS.ink),
                              border: m.role === 'error' ? `1px solid ${COLORS.rose}` : 'none',
                            }}>{m.content}</div>
                          </div>
                        ))}
                        {qaStatus === 'loading' ? (
                          <div style={{ display: 'flex', justifyContent: 'flex-start' }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontFamily: SANS, fontSize: 12.5, color: COLORS.muted, padding: '9px 13px', borderRadius: 12, background: COLORS.bg }}>
                              <Loader2 size={13} className="animate-spin" /> Thinking…
                            </div>
                          </div>
                        ) : null}
                      </div>
                    ) : (
                      qaSuggestions.length ? (
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, marginBottom: 14 }}>
                          {qaSuggestions.map((q, i) => (
                            <button key={i} onClick={() => submitQA(q)} disabled={qaStatus === 'loading'}
                              style={{ fontFamily: SANS, fontSize: 11.5, color: COLORS.accent, background: COLORS.accentSoft, border: 'none', borderRadius: 999, padding: '7px 12px', cursor: 'pointer', textAlign: 'left' }}>
                              {q}
                            </button>
                          ))}
                        </div>
                      ) : null
                    )}
                    <form onSubmit={e => { e.preventDefault(); submitQA(); }} style={{ display: 'flex', gap: 8 }}>
                      <input value={qaInput} onChange={e => setQaInput(e.target.value)} placeholder="e.g. Which segment lost the most margin last quarter?"
                        disabled={qaStatus === 'loading'}
                        style={{ flex: 1, fontFamily: SANS, fontSize: 13, padding: '9px 12px', borderRadius: 8, border: `1px solid ${COLORS.border}` }} />
                      <button type="submit" disabled={qaStatus === 'loading' || !qaInput.trim()}
                        style={{ display: 'flex', alignItems: 'center', gap: 6, fontFamily: SANS, fontSize: 12.5, fontWeight: 700, color: '#FFFFFF', background: COLORS.accent, border: 'none', borderRadius: 8, padding: '9px 14px', cursor: (qaStatus === 'loading' || !qaInput.trim()) ? 'default' : 'pointer', opacity: (qaStatus === 'loading' || !qaInput.trim()) ? 0.6 : 1 }}>
                        <Send size={13} /> Ask
                      </button>
                    </form>
                  </Panel>
                  {dataQuality ? (
                    <Panel style={{ marginBottom: 24 }}>
                      <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 700, color: COLORS.ink, marginBottom: 12 }}>Data quality</div>
                      <div className="grid grid-cols-2 sm:grid-cols-4" style={{ gap: 12, marginBottom: dataQuality.nullStats.length || negativeCounts.length ? 14 : 0 }}>
                        <div>
                          <div style={{ fontFamily: MONO, fontSize: 20, fontWeight: 700, color: COLORS.ink }}>{dataQuality.total}</div>
                          <div style={{ fontFamily: SANS, fontSize: 11, color: COLORS.muted }}>rows loaded</div>
                        </div>
                        <div>
                          <div style={{ fontFamily: MONO, fontSize: 20, fontWeight: 700, color: dataQuality.duplicates > 0 ? COLORS.rose : COLORS.emerald }}>{dataQuality.duplicates}</div>
                          <div style={{ fontFamily: SANS, fontSize: 11, color: COLORS.muted }}>duplicate rows</div>
                        </div>
                        <div>
                          <div style={{ fontFamily: MONO, fontSize: 20, fontWeight: 700, color: dataQuality.nullStats.length > 0 ? COLORS.amberText : COLORS.emerald }}>{dataQuality.nullStats.length}</div>
                          <div style={{ fontFamily: SANS, fontSize: 11, color: COLORS.muted }}>columns with gaps</div>
                        </div>
                        <div>
                          <div style={{ fontFamily: MONO, fontSize: 20, fontWeight: 700, color: dataQuality.badDateCount > 0 ? COLORS.rose : COLORS.emerald }}>{dataQuality.badDateCount}</div>
                          <div style={{ fontFamily: SANS, fontSize: 11, color: COLORS.muted }}>unparseable dates</div>
                        </div>
                      </div>
                      {dataQuality.nullStats.length > 0 ? (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginBottom: negativeCounts.length ? 10 : 0 }}>
                          {dataQuality.nullStats.slice(0, 4).map(s => (
                            <div key={s.field} style={{ display: 'flex', justifyContent: 'space-between', fontFamily: SANS, fontSize: 11.5, color: '#3A4256' }}>
                              <span>{s.field}</span>
                              <span style={{ fontFamily: MONO }}>{s.nulls} missing ({fmtPct(s.pct, 0)})</span>
                            </div>
                          ))}
                        </div>
                      ) : null}
                      {negativeCounts.length > 0 ? (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                          {negativeCounts.map(n => (
                            <div key={n.key} style={{ display: 'flex', justifyContent: 'space-between', fontFamily: SANS, fontSize: 11.5, color: '#3A4256' }}>
                              <span>{n.key} has negative values</span>
                              <span style={{ fontFamily: MONO }}>{n.count} rows</span>
                            </div>
                          ))}
                        </div>
                      ) : null}
                    </Panel>
                  ) : null}
                  {!insightInput ? (
                    <EmptyState icon={Sparkles} title="Nothing to analyze yet" body="Turn on at least one metric in Configure analysis (Dashboard tab) to generate insights." />
                  ) : (
                <>
                  <div style={{ marginBottom: 28 }}>
                    <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 700, color: COLORS.ink, marginBottom: 12 }}>Instant pattern detection</div>
                    {ruleInsights.length === 0 ? (
                      <EmptyState icon={Sparkles} title="Not enough structure yet" body="Enable a dimension, a date column, or a ratio metric in Configure analysis to surface patterns here." />
                    ) : (
                      <div className="grid grid-cols-1 lg:grid-cols-2" style={{ gap: 12 }}>
                        {ruleInsights.map((ins, i) => <InsightCard key={i} type={ins.type} title={ins.title} detail={ins.detail} />)}
                      </div>
                    )}
                  </div>

                  <div className="grid grid-cols-1 lg:grid-cols-2" style={{ gap: 16, marginBottom: 28 }}>
                    <Panel>
                      <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 700, color: COLORS.ink, marginBottom: 4 }}>What moves together</div>
                      <div style={{ fontFamily: SANS, fontSize: 11.5, color: COLORS.muted, marginBottom: 12 }}>Correlation across per-row numeric fields.</div>
                      {correlations.filter(c => Math.abs(c.r) >= 0.2).length === 0 ? (
                        <div style={{ fontFamily: SANS, fontSize: 12, color: COLORS.muted }}>{correlations.length === 0 ? 'Turn on at least two metrics to check for correlation.' : 'No meaningful correlations (|r| ≥ 0.2) found among the active metrics.'}</div>
                      ) : (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                          {correlations.filter(c => Math.abs(c.r) >= 0.2).slice(0, 5).map((c, i) => (
                            <div key={i} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '8px 10px', borderRadius: 8, background: COLORS.bg }}>
                              <span style={{ fontFamily: SANS, fontSize: 12, color: COLORS.ink, fontWeight: 600 }}>{c.a} ↔ {c.b}</span>
                              <span style={{ fontFamily: MONO, fontSize: 12, fontWeight: 700, color: c.r > 0 ? COLORS.emerald : COLORS.rose }}>r = {c.r.toFixed(2)} ({corrStrength(c.r)})</span>
                            </div>
                          ))}
                        </div>
                      )}
                    </Panel>
                    <Panel>
                      <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 700, color: COLORS.ink, marginBottom: 4 }}>Outliers — {outlierMetricLabel}</div>
                      <div style={{ fontFamily: SANS, fontSize: 11.5, color: COLORS.muted, marginBottom: 12 }}>Rows more than 2.5 standard deviations from the mean.</div>
                      {outliers.length === 0 ? (
                        <div style={{ fontFamily: SANS, fontSize: 12, color: COLORS.muted }}>No statistical outliers detected — the data is fairly consistent.</div>
                      ) : (
                        <div>
                          <div style={{ fontFamily: MONO, fontSize: 22, fontWeight: 700, color: COLORS.amberText, marginBottom: 8 }}>{outliers.length}</div>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                            {outliers.slice(0, 4).map((o, i) => (
                              <div key={i} style={{ fontFamily: SANS, fontSize: 11.5, color: '#3A4256', display: 'flex', justifyContent: 'space-between' }}>
                                <span>{o.row[customDataset.fields[0]] !== undefined ? String(o.row[customDataset.fields[0]]) : 'row'}</span>
                                <span style={{ fontFamily: MONO, fontWeight: 700 }}>{fmtByFormat(o.value, outlierMetricFormat)} (z={o.z.toFixed(1)})</span>
                              </div>
                            ))}
                          </div>
                        </div>
                      )}
                    </Panel>
                  </div>

                  <div style={{ marginBottom: 28 }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12 }}>
                      <Zap size={15} color={COLORS.amberText} />
                      <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 700, color: COLORS.ink }}>Threshold alerts</div>
                    </div>
                    <Panel>
                      <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', alignItems: 'flex-end', marginBottom: alertMatches.length ? 16 : 0 }}>
                        <Select label="Metric" value={alertRule.metricKey} onChange={v => setAlertRule(a => ({ ...a, metricKey: v }))} options={['', ...alertOptions.map(o => o.key)]} />
                        <Select label="Condition" value={alertRule.op} onChange={v => setAlertRule(a => ({ ...a, op: v }))} options={['below', 'above']} />
                        <label style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                          <span style={{ fontFamily: SANS, fontSize: 11, fontWeight: 600, color: COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Value</span>
                          <input type="number" value={alertRule.value} onChange={e => setAlertRule(a => ({ ...a, value: e.target.value }))}
                            style={{ fontFamily: MONO, fontSize: 13.5, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: '8px 10px', width: 120 }} />
                        </label>
                      </div>
                      {!alertRule.metricKey || alertRule.value === '' ? (
                        <div style={{ fontFamily: SANS, fontSize: 12, color: COLORS.muted }}>Pick a metric and a value — every segment across every dimension that crosses it will be flagged below.</div>
                      ) : alertMatches.length === 0 ? (
                        <div style={{ fontFamily: SANS, fontSize: 12, color: COLORS.muted }}>No segments currently {alertRule.op} that value.</div>
                      ) : (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                          {alertMatches.map((seg, i) => (
                            <div key={i} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '8px 10px', borderRadius: 8, background: '#FDEEF0', border: `1px solid ${COLORS.rose}` }}>
                              <span style={{ fontFamily: SANS, fontSize: 12, color: COLORS.ink, fontWeight: 600 }}>{seg.name}</span>
                              <span style={{ fontFamily: MONO, fontSize: 12, fontWeight: 700, color: COLORS.rose }}>{fmtByFormat(seg[alertRule.metricKey], (alertOptions.find(o => o.key === alertRule.metricKey) || {}).format)}</span>
                            </div>
                          ))}
                        </div>
                      )}
                    </Panel>
                  </div>

                  <div>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 12, flexWrap: 'wrap', gap: 10 }}>
                      <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 700, color: COLORS.ink }}>AI-written summary</div>
                      <button onClick={generateAIInsights} disabled={aiInsights.status === 'loading'}
                        style={{ display: 'flex', alignItems: 'center', gap: 6, fontFamily: SANS, fontSize: 12.5, fontWeight: 600, color: '#FFFFFF', background: `linear-gradient(135deg, ${COLORS.accent}, ${COLORS.violet})`, border: 'none', borderRadius: 8, padding: '8px 14px', cursor: aiInsights.status === 'loading' ? 'default' : 'pointer', opacity: aiInsights.status === 'loading' ? 0.7 : 1 }}>
                        {aiInsights.status === 'loading' ? <Loader2 size={14} className="animate-spin" /> : <Sparkles size={14} />}
                        {aiInsights.status === 'loading' ? 'Analyzing…' : (aiInsights.status === 'done' ? 'Regenerate' : 'Generate AI insights')}
                      </button>
                    </div>

                    {aiInsights.status === 'error' ? (
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8, background: '#FDEEF0', border: `1px solid ${COLORS.rose}`, color: '#8A1F30', borderRadius: 8, padding: '10px 14px', fontFamily: SANS, fontSize: 12.5, marginBottom: 12 }}>
                        <AlertCircle size={15} /> Couldn't generate AI insights: {aiInsights.error}
                      </div>
                    ) : null}

                    {aiInsights.status === 'done' && aiInsights.data ? (
                      <div className="grid grid-cols-1 lg:grid-cols-3" style={{ gap: 16 }}>
                        <div>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontFamily: SANS, fontSize: 11.5, fontWeight: 700, color: COLORS.emerald, textTransform: 'uppercase', letterSpacing: '0.04em', marginBottom: 10 }}><CheckCircle2 size={14} /> Highlights</div>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                            {(aiInsights.data.highlights || []).map((h, i) => (
                              <div key={i} style={{ padding: 12, borderRadius: 10, border: `1px solid ${COLORS.border}`, background: '#EAFBF3' }}>
                                <div style={{ fontFamily: SANS, fontWeight: 700, fontSize: 12.5, color: COLORS.ink, marginBottom: 4 }}>{h.title}</div>
                                <div style={{ fontFamily: SANS, fontSize: 12, color: '#3A4256', lineHeight: 1.4 }}>{h.detail}</div>
                              </div>
                            ))}
                          </div>
                        </div>
                        <div>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontFamily: SANS, fontSize: 11.5, fontWeight: 700, color: COLORS.rose, textTransform: 'uppercase', letterSpacing: '0.04em', marginBottom: 10 }}><AlertTriangle size={14} /> Pain points</div>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                            {(aiInsights.data.painPoints || []).map((h, i) => (
                              <div key={i} style={{ padding: 12, borderRadius: 10, border: `1px solid ${COLORS.border}`, background: '#FDEEF0' }}>
                                <div style={{ fontFamily: SANS, fontWeight: 700, fontSize: 12.5, color: COLORS.ink, marginBottom: 4 }}>{h.title}</div>
                                <div style={{ fontFamily: SANS, fontSize: 12, color: '#3A4256', lineHeight: 1.4 }}>{h.detail}</div>
                              </div>
                            ))}
                          </div>
                        </div>
                        <div>
                          <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontFamily: SANS, fontSize: 11.5, fontWeight: 700, color: COLORS.amberText, textTransform: 'uppercase', letterSpacing: '0.04em', marginBottom: 10 }}><Target size={14} /> Opportunities</div>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                            {(aiInsights.data.opportunities || []).map((h, i) => (
                              <div key={i} style={{ padding: 12, borderRadius: 10, border: `1px solid ${COLORS.border}`, background: '#FFF8EB' }}>
                                <div style={{ fontFamily: SANS, fontWeight: 700, fontSize: 12.5, color: COLORS.ink, marginBottom: 4 }}>{h.title}</div>
                                <div style={{ fontFamily: SANS, fontSize: 12, color: '#3A4256', lineHeight: 1.4 }}>{h.detail}</div>
                              </div>
                            ))}
                          </div>
                        </div>
                      </div>
                    ) : (aiInsights.status !== 'loading' ? (
                      <div style={{ fontFamily: SANS, fontSize: 12.5, color: COLORS.muted }}>
                        Click "Generate AI insights" for a narrative summary tailored to this data's domain — uses one API call, takes a few seconds.
                      </div>
                    ) : null)}
                  </div>
                </>
                  )}
                </>
              )}
            </div>
          )}

          {tab === 'trends' && (
            <div>
              {!customDataset ? (
                <EmptyState icon={TrendingUp} title="No data loaded" body="Upload a file from the Dashboard tab to see trends over time." />
              ) : !dateField || genTrend.length < 2 ? (
                <EmptyState icon={TrendingUp} title="Trends need a date column" body="Open Configure analysis on the Dashboard tab and set a Date column with at least two distinct periods of data." />
              ) : activeMetrics.length === 0 ? (
                <EmptyState icon={AlertCircle} title="No metrics selected" body="Turn on at least one metric in Configure analysis to chart it over time." />
              ) : (
                <>
                  <SectionHeader title="Trends Over Time" subtitle={dateField + ', grouped by month' + (Object.values(genFilters).some(v => v && v.length) ? ' — reflects the active filters.' : '.')} />
                  <ActiveFiltersBar filters={genFilters} onRemove={toggleFilterValue} onClearAll={() => setGenFilters({})} />
                  <Panel style={{ marginBottom: 16 }}>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 10, marginBottom: 4 }}>
                      <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 600, color: COLORS.ink }}>Metric trends</div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 14, flexWrap: 'wrap' }}>
                        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
                          {activeMetrics.slice(0, 3).map((m, i) => {
                            const off = !!hiddenTrendSeries[m.key];
                            return (
                              <button key={m.key} onClick={() => setHiddenTrendSeries(h => ({ ...h, [m.key]: !h[m.key] }))}
                                title={off ? 'Click to show' : 'Click to hide'}
                                style={{ display: 'inline-flex', alignItems: 'center', gap: 6, fontFamily: SANS, fontSize: 11.5, fontWeight: 600, color: off ? COLORS.muted : COLORS.ink, background: 'transparent', border: 'none', cursor: 'pointer', padding: 0, opacity: off ? 0.5 : 1, textDecoration: off ? 'line-through' : 'none' }}>
                                <span style={{ width: 9, height: 9, borderRadius: 999, background: CHART_PALETTE[i], flexShrink: 0 }} />
                                {m.key}
                              </button>
                            );
                          })}
                        </div>
                        <div style={{ display: 'flex', gap: 2, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: 2 }}>
                          {['area', 'bar'].map(t => (
                            <button key={t} onClick={() => setTrendChartType(t)}
                              style={{ fontFamily: SANS, fontSize: 11, fontWeight: 600, padding: '4px 10px', borderRadius: 6, cursor: 'pointer', border: 'none', textTransform: 'capitalize', background: trendChartType === t ? COLORS.ink : 'transparent', color: trendChartType === t ? '#FFFFFF' : COLORS.muted }}>{t}</button>
                          ))}
                        </div>
                      </div>
                    </div>
                    <div style={{ fontFamily: SANS, fontSize: 11, color: COLORS.muted, marginBottom: 12 }}>Click a legend item to hide/show it · hover a point for exact values · drag the handles below to zoom into a range.</div>
                    <div style={{ height: 320 }}>
                      <ResponsiveContainer width="100%" height="100%">
                        {trendChartType === 'area' ? (
                          <AreaChart data={genTrend} margin={{ top: 4, right: 8, left: 8, bottom: 0 }}>
                            <defs>
                              {activeMetrics.slice(0, 3).map((m, i) => (
                                <linearGradient key={m.key} id={`trendGrad${i}`} x1="0" y1="0" x2="0" y2="1">
                                  <stop offset="5%" stopColor={CHART_PALETTE[i]} stopOpacity={0.32} />
                                  <stop offset="95%" stopColor={CHART_PALETTE[i]} stopOpacity={0} />
                                </linearGradient>
                              ))}
                            </defs>
                            <CartesianGrid stroke={COLORS.border} vertical={false} />
                            <XAxis dataKey="label" tick={{ fontSize: 10, fill: COLORS.muted, fontFamily: SANS }} />
                            <YAxis tick={{ fontSize: 10.5, fill: COLORS.muted, fontFamily: SANS }} />
                            <Tooltip content={<RichTooltip formatMap={Object.fromEntries(activeMetrics.map(m => [m.key, m.format]))} />} cursor={{ stroke: COLORS.accent, strokeWidth: 1, strokeDasharray: '3 3' }} />
                            {activeMetrics.slice(0, 3).map((m, i) => (
                              hiddenTrendSeries[m.key] ? null : (
                                <Area key={m.key} type="monotone" dataKey={m.key} name={m.key} stroke={CHART_PALETTE[i]} strokeWidth={2.25} fill={`url(#trendGrad${i})`}
                                  activeDot={{ r: 5, stroke: '#FFFFFF', strokeWidth: 2 }} />
                              )
                            ))}
                            <Brush dataKey="label" height={22} stroke={COLORS.accent} fill={COLORS.bg} travellerWidth={8} />
                          </AreaChart>
                        ) : (
                          <BarChart data={genTrend} margin={{ top: 4, right: 8, left: 8, bottom: 0 }}>
                            <CartesianGrid stroke={COLORS.border} vertical={false} />
                            <XAxis dataKey="label" tick={{ fontSize: 10, fill: COLORS.muted, fontFamily: SANS }} />
                            <YAxis tick={{ fontSize: 10.5, fill: COLORS.muted, fontFamily: SANS }} />
                            <Tooltip content={<RichTooltip formatMap={Object.fromEntries(activeMetrics.map(m => [m.key, m.format]))} />} cursor={{ fill: COLORS.bg }} />
                            {activeMetrics.slice(0, 3).map((m, i) => (
                              hiddenTrendSeries[m.key] ? null : (
                                <Bar key={m.key} dataKey={m.key} name={m.key} fill={CHART_PALETTE[i]} radius={[4, 4, 0, 0]} maxBarSize={24} />
                              )
                            ))}
                            <Brush dataKey="label" height={22} stroke={COLORS.accent} fill={COLORS.bg} travellerWidth={8} />
                          </BarChart>
                        )}
                      </ResponsiveContainer>
                    </div>
                  </Panel>
                  <Panel style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontFamily: SANS, fontSize: 12 }}>
                      <thead>
                        <tr style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                          <th style={{ textAlign: 'left', padding: '7px 9px', color: COLORS.muted, fontWeight: 600, fontSize: 10.5, textTransform: 'uppercase' }}>Period</th>
                          {activeMetrics.map(m => <th key={m.key} style={{ textAlign: 'right', padding: '7px 9px', color: COLORS.muted, fontWeight: 600, fontSize: 10.5, textTransform: 'uppercase' }}>{m.key}</th>)}
                          {activeRatios.map(rt => <th key={rt.key} style={{ textAlign: 'right', padding: '7px 9px', color: COLORS.muted, fontWeight: 600, fontSize: 10.5, textTransform: 'uppercase' }}>{rt.label}</th>)}
                        </tr>
                      </thead>
                      <tbody>
                        {genTrend.map(row => (
                          <tr key={row.month} style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                            <td style={{ padding: '7px 9px', fontWeight: 600, color: COLORS.ink }}>{row.label}</td>
                            {activeMetrics.map(m => <td key={m.key} style={{ padding: '7px 9px', textAlign: 'right', fontFamily: MONO }}>{fmtByFormat(row[m.key], m.format)}</td>)}
                            {activeRatios.map(rt => <td key={rt.key} style={{ padding: '7px 9px', textAlign: 'right', fontFamily: MONO }}>{fmtByFormat(row[rt.key], rt.format)}</td>)}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </Panel>
                </>
              )}
            </div>
          )}

          {tab === 'forecast' && (
            <div>
              {!customDataset ? (
                <EmptyState icon={SlidersHorizontal} title="No data loaded" body="Upload a file from the Dashboard tab to forecast a metric forward." />
              ) : !dateField || genTrend.length < 2 ? (
                <EmptyState icon={SlidersHorizontal} title="Forecasting needs trend history" body="Set a date column with at least two periods of data in Configure analysis (Dashboard tab) to unlock forecasting." />
              ) : (
                <>
                  <SectionHeader title="Forecast &amp; What-If Scenarios" subtitle="Growth rates are auto-estimated from your own historical trend — tune Custom to override, or drag the slider." />
                  <div style={{ display: 'flex', gap: 20, flexWrap: 'wrap', marginBottom: 18, alignItems: 'flex-end' }}>
                    <Select label="Metric to forecast" value={genForecastMetric} onChange={setGenForecastMetric} options={activeMetrics.map(m => m.key)} />
                    <label style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                      <span style={{ fontFamily: SANS, fontSize: 11, fontWeight: 600, color: COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.04em' }}>Horizon</span>
                      <div style={{ display: 'flex', gap: 6 }}>
                        {[3, 6, 12, 24].map(h => (
                          <button key={h} onClick={() => setGenForecastHorizon(h)}
                            style={{ fontFamily: SANS, fontSize: 12.5, fontWeight: 600, padding: '8px 12px', borderRadius: 8, cursor: 'pointer', border: `1px solid ${genForecastHorizon === h ? COLORS.ink : COLORS.border}`, background: genForecastHorizon === h ? COLORS.ink : '#FFFFFF', color: genForecastHorizon === h ? '#FFFFFF' : COLORS.ink }}>
                            {h}mo
                          </button>
                        ))}
                      </div>
                    </label>
                  </div>
                  <div style={{ display: 'flex', gap: 8, marginBottom: 18, flexWrap: 'wrap' }}>
                    {['Base', 'Optimistic', 'Pessimistic', 'Custom'].map(s => (
                      <button key={s} onClick={() => setGenScenario(s)}
                        style={{ fontFamily: SANS, fontSize: 12.5, fontWeight: 600, padding: '9px 16px', borderRadius: 999, cursor: 'pointer', border: `1px solid ${genScenario === s ? COLORS.ink : COLORS.border}`, background: genScenario === s ? COLORS.ink : '#FFFFFF', color: genScenario === s ? '#FFFFFF' : COLORS.ink }}>{s}</button>
                    ))}
                  </div>
                  {genScenario === 'Custom' ? (
                    <Panel style={{ marginBottom: 18, background: COLORS.accentSoft, border: `1px solid ${COLORS.accent}` }}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 10, flexWrap: 'wrap', gap: 8 }}>
                        <span style={{ fontFamily: SANS, fontSize: 11, fontWeight: 600, color: COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.04em' }}>{genForecastMetric} monthly growth</span>
                        <input type="number" step="0.1" value={(genCustomGrowth * 100).toFixed(1)}
                          onChange={e => { const v = parseFloat(e.target.value); setGenCustomGrowth(isNaN(v) ? 0 : v / 100); }}
                          style={{ fontFamily: MONO, fontSize: 14, fontWeight: 700, padding: '6px 9px', borderRadius: 8, border: `1px solid ${COLORS.border}`, width: 90, textAlign: 'right' }} />
                      </div>
                      <input type="range" min={-20} max={20} step={0.5} value={(genCustomGrowth * 100).toFixed(1)}
                        onChange={e => setGenCustomGrowth(parseFloat(e.target.value) / 100)}
                        style={{ width: '100%', accentColor: COLORS.accent, cursor: 'pointer' }} />
                      <div style={{ display: 'flex', justifyContent: 'space-between', fontFamily: MONO, fontSize: 10, color: COLORS.muted, marginTop: 2 }}>
                        <span>−20%</span><span>0%</span><span>+20%</span>
                      </div>
                    </Panel>
                  ) : (
                    <Panel style={{ marginBottom: 18 }}>
                      <div style={{ fontFamily: SANS, fontSize: 12, color: COLORS.muted }}>
                        <strong style={{ color: COLORS.ink }}>{genScenario}:</strong> {fmtPct(genScenarioRates[genScenario], 2)} / month on {genForecastMetric}
                        {genScenario === 'Base' ? ' (this dataset’s own historical average)' : ''}
                      </div>
                    </Panel>
                  )}
                  <Panel style={{ marginBottom: 18, border: `1px solid ${COLORS.violet}` }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
                      <Target size={15} color={COLORS.violet} />
                      <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 700, color: COLORS.ink }}>Goal seek — work backward from a target</div>
                    </div>
                    <div style={{ display: 'flex', gap: 12, flexWrap: 'wrap', alignItems: 'flex-end' }}>
                      <label style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                        <span style={{ fontFamily: SANS, fontSize: 11, fontWeight: 600, color: COLORS.muted }}>Target {genForecastMetric} by period {genForecastHorizon}</span>
                        <input type="number" value={goalSeekTarget} onChange={e => setGoalSeekTarget(e.target.value)} placeholder="e.g. 50000"
                          style={{ fontFamily: MONO, fontSize: 13.5, padding: '8px 10px', borderRadius: 8, border: `1px solid ${COLORS.border}`, width: 160 }} />
                      </label>
                      {goalSeekRate !== null ? (
                        <>
                          <div style={{ fontFamily: SANS, fontSize: 12, color: COLORS.muted, paddingBottom: 8 }}>
                            requires <strong style={{ color: COLORS.ink, fontFamily: MONO }}>{fmtPct(goalSeekRate, 2)}</strong> / month
                          </div>
                          <button onClick={() => { setGenScenario('Custom'); setGenCustomGrowth(goalSeekRate); }}
                            style={{ fontFamily: SANS, fontSize: 12, fontWeight: 700, color: '#FFFFFF', background: COLORS.violet, border: 'none', borderRadius: 8, padding: '8px 14px', cursor: 'pointer', marginBottom: 0 }}>
                            Apply this growth rate
                          </button>
                        </>
                      ) : goalSeekTarget ? (
                        <div style={{ fontFamily: SANS, fontSize: 12, color: COLORS.muted, paddingBottom: 8 }}>Enter a positive target above the current base period to solve for a growth rate.</div>
                      ) : null}
                    </div>
                  </Panel>
                  <div className="grid grid-cols-2 sm:grid-cols-3" style={{ gap: 12, marginBottom: 20 }}>
                    <KpiCard label={'Base period (' + (genTrend[genTrend.length - 1] ? genTrend[genTrend.length - 1].label : '') + ')'} value={fmtByFormat(genBaseValue, activeMetricFormat(genForecastMetric))} sub="most recent actual" />
                    <KpiCard label={genForecast.length ? genForecast[genForecast.length - 1].label + ' projected' : 'Projected'} value={genForecast.length ? fmtByFormat(genForecast[genForecast.length - 1].value, activeMetricFormat(genForecastMetric)) : '—'} accent={COLORS.emerald} />
                    <KpiCard label="Monthly growth applied" value={fmtPct(genActiveGrowth, 2)} accent={COLORS.amberText} />
                  </div>
                  <Panel style={{ marginBottom: 16, background: COLORS.accentSoft }}>
                    <div style={{ fontFamily: SANS, fontSize: 12.5, fontWeight: 700, color: COLORS.ink }}>Forecast reality check</div>
                    <div style={{ fontFamily: SANS, fontSize: 12, color: COLORS.muted, marginTop: 5 }}>
                      {forecastBacktest && forecastBacktest.samples >= 3
                        ? `One-month backtest: ${fmtPct(forecastBacktest.meanAbsolutePercentError, 1)} average absolute error across ${forecastBacktest.samples} past predictions. Scenarios are illustrations, not guarantees.`
                        : 'Too few consecutive positive monthly observations to estimate forecast error. Treat these scenarios as illustrations.'}
                    </div>
                  </Panel>
                  <Panel style={{ marginBottom: 16 }}>
                    <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 600, color: COLORS.ink, marginBottom: 12 }}>Forecast — {genForecastMetric} ({genScenario})</div>
                    <div style={{ height: 240 }}>
                      <ResponsiveContainer width="100%" height="100%">
                        <AreaChart data={genForecastChartData} margin={{ top: 4, right: 8, left: -12, bottom: 0 }}>
                          <defs>
                            <linearGradient id="forecastGrad" x1="0" y1="0" x2="0" y2="1">
                              <stop offset="5%" stopColor={COLORS.emerald} stopOpacity={0.35} />
                              <stop offset="95%" stopColor={COLORS.emerald} stopOpacity={0} />
                            </linearGradient>
                          </defs>
                          <CartesianGrid stroke={COLORS.border} vertical={false} />
                          <XAxis dataKey="label" tick={{ fontSize: 11, fill: COLORS.muted, fontFamily: SANS }} />
                          <YAxis tick={{ fontSize: 10.5, fill: COLORS.muted, fontFamily: SANS }} />
                          <Tooltip content={<RichTooltip formatMap={{ value: activeMetricFormat(genForecastMetric) }} />} cursor={{ stroke: COLORS.emerald, strokeWidth: 1, strokeDasharray: '3 3' }} />
                          <Area type="monotone" dataKey="value" name={genForecastMetric} stroke={COLORS.emerald} strokeWidth={2.5} fill="url(#forecastGrad)"
                            dot={{ r: 3 }} activeDot={{ r: 5, stroke: '#FFFFFF', strokeWidth: 2 }} />
                        </AreaChart>
                      </ResponsiveContainer>
                    </div>
                  </Panel>
                  <div className="grid grid-cols-1 lg:grid-cols-2" style={{ gap: 16 }}>
                    <Panel style={{ overflowX: 'auto' }}>
                      <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 600, color: COLORS.ink, marginBottom: 10 }}>Period-by-period detail</div>
                      <table style={{ width: '100%', borderCollapse: 'collapse', fontFamily: SANS, fontSize: 12 }}>
                        <thead>
                          <tr style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                            <th style={{ textAlign: 'left', padding: '7px 8px', color: COLORS.muted, fontWeight: 600, fontSize: 10.5, textTransform: 'uppercase' }}>Period</th>
                            <th style={{ textAlign: 'right', padding: '7px 8px', color: COLORS.muted, fontWeight: 600, fontSize: 10.5, textTransform: 'uppercase' }}>{genForecastMetric}</th>
                          </tr>
                        </thead>
                        <tbody>
                          {genForecast.map(row => (
                            <tr key={row.label} style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                              <td style={{ padding: '7px 8px', fontWeight: 600, color: COLORS.ink }}>{row.label}</td>
                              <td style={{ padding: '7px 8px', textAlign: 'right', fontFamily: MONO }}>{fmtByFormat(row.value, activeMetricFormat(genForecastMetric))}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </Panel>
                    <Panel>
                      <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 600, color: COLORS.ink, marginBottom: 10 }}>Period-{genForecastHorizon} snapshot — all scenarios</div>
                      <div style={{ height: 190 }}>
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={genSnapshots} margin={{ top: 4, right: 8, left: -12, bottom: 0 }}>
                            <defs>
                              <linearGradient id="snapGrad" x1="0" y1="0" x2="0" y2="1">
                                <stop offset="0%" stopColor={COLORS.accent} stopOpacity={1} />
                                <stop offset="100%" stopColor={COLORS.accent} stopOpacity={0.55} />
                              </linearGradient>
                            </defs>
                            <CartesianGrid stroke={COLORS.border} vertical={false} />
                            <XAxis dataKey="name" tick={{ fontSize: 11, fill: COLORS.muted, fontFamily: SANS }} />
                            <YAxis tick={{ fontSize: 10.5, fill: COLORS.muted, fontFamily: SANS }} />
                            <Tooltip content={<RichTooltip formatMap={{ value: activeMetricFormat(genForecastMetric) }} />} cursor={{ fill: COLORS.bg }} />
                            <Bar dataKey="value" name={genForecastMetric} fill="url(#snapGrad)" radius={[6, 6, 0, 0]} />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                      <div style={{ fontFamily: SANS, fontSize: 11, color: COLORS.muted, marginTop: 8 }}>Base/Optimistic/Pessimistic rates are derived from this dataset's own history — not a generic assumption.</div>
                    </Panel>
                  </div>
                  <Panel style={{ marginTop: 16 }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
                      <Grid3x3 size={15} color={COLORS.accent} />
                      <div style={{ fontFamily: SANS, fontSize: 13, fontWeight: 700, color: COLORS.ink }}>Sensitivity grid</div>
                    </div>
                    {!activeRatios.length ? (
                      <div style={{ fontFamily: SANS, fontSize: 12.5, color: COLORS.muted, padding: '10px 0 4px 0' }}>
                        Needs a calculated metric with two inputs (e.g. <strong style={{ color: COLORS.ink }}>Profit = Revenue − Cost</strong>) to vary independently.
                        Add one under <strong style={{ color: COLORS.ink }}>Configure analysis → Custom metrics</strong> to unlock this.
                      </div>
                    ) : (
                      <>
                        <div style={{ fontFamily: SANS, fontSize: 12, color: COLORS.muted, marginBottom: 14 }}>
                          Cross two growth assumptions at once — {sensOutcomeRatio ? (<><strong style={{ color: COLORS.ink }}>{sensOutcomeRatio.num}</strong> growth (rows) against <strong style={{ color: COLORS.ink }}>{sensOutcomeRatio.den}</strong> growth (columns)</>) : null} — instead of testing one lever at a time. Shaded by value: greener is higher {sensOutcomeRatio ? sensOutcomeRatio.label : ''}, redder is lower — not a judgment of "better," since for some metrics lower is the goal.
                        </div>
                        <div style={{ display: 'flex', gap: 14, flexWrap: 'wrap', alignItems: 'flex-end', marginBottom: 16 }}>
                          <Select label="Outcome metric" value={sensOutcomeKey} onChange={setSensOutcomeKey} options={activeRatios.map(r => r.key)} labelize={k => (activeRatios.find(r => r.key === k) || {}).label || k} />
                          <label style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                            <span style={{ fontFamily: SANS, fontSize: 11, fontWeight: 600, color: COLORS.muted }}>Spread (± pp)</span>
                            <input type="number" min={0.5} max={30} step={0.5} value={sensSpread} onChange={e => setSensSpread(e.target.value)}
                              style={{ fontFamily: MONO, fontSize: 13, padding: '7px 9px', borderRadius: 8, border: `1px solid ${COLORS.border}`, width: 90 }} />
                          </label>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                            <span style={{ fontFamily: SANS, fontSize: 11, fontWeight: 600, color: COLORS.muted }}>Steps per axis</span>
                            <div style={{ display: 'flex', gap: 2, border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: 2 }}>
                              {[3, 5, 7].map(n => (
                                <button key={n} onClick={() => setSensSteps(n)}
                                  style={{ fontFamily: SANS, fontSize: 12, fontWeight: 600, padding: '5px 10px', borderRadius: 6, cursor: 'pointer', border: 'none', background: sensSteps === n ? COLORS.ink : 'transparent', color: sensSteps === n ? '#FFFFFF' : COLORS.muted }}>{n}</button>
                              ))}
                            </div>
                          </div>
                        </div>
                        {sensOutcomeRatio ? (
                          <div style={{ overflowX: 'auto' }}>
                            <table style={{ borderCollapse: 'collapse', fontFamily: MONO, fontSize: 11.5, minWidth: '100%' }}>
                              <thead>
                                <tr>
                                  <th style={{ padding: '6px 10px', textAlign: 'right', fontFamily: SANS, fontSize: 10, fontWeight: 700, color: COLORS.muted, textTransform: 'uppercase', borderBottom: `1px solid ${COLORS.border}` }}>
                                    {sensOutcomeRatio.num} ↓ / {sensOutcomeRatio.den} →
                                  </th>
                                  {sensColGrowths.map((cg, ci) => (
                                    <th key={ci} style={{ padding: '6px 10px', textAlign: 'center', fontWeight: 700, color: COLORS.ink, borderBottom: `1px solid ${COLORS.border}` }}>{fmtPct(cg, 1)}</th>
                                  ))}
                                </tr>
                              </thead>
                              <tbody>
                                {sensRowGrowths.map((rg, ri) => (
                                  <tr key={ri}>
                                    <td style={{ padding: '6px 10px', textAlign: 'right', fontWeight: 700, color: COLORS.ink, borderRight: `1px solid ${COLORS.border}` }}>{fmtPct(rg, 1)}</td>
                                    {sensColGrowths.map((cg, ci) => {
                                      const isBaseCell = ri === Math.floor(sensRowGrowths.length / 2) && ci === Math.floor(sensColGrowths.length / 2);
                                      return (
                                        <td key={ci} title={isBaseCell ? 'Base case — both inputs at their own historical average growth' : undefined}
                                          style={{ padding: '9px 12px', textAlign: 'center', background: sensCellColor(sensGrid[ri][ci]), fontWeight: isBaseCell ? 700 : 500, color: COLORS.ink, outline: isBaseCell ? `2px solid ${COLORS.ink}` : 'none', outlineOffset: -2 }}>
                                          {fmtByFormat(sensGrid[ri][ci], sensOutcomeRatio.format)}
                                        </td>
                                      );
                                    })}
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        ) : null}
                        <div style={{ fontFamily: SANS, fontSize: 11, color: COLORS.muted, marginTop: 10 }}>
                          Boxed cell is the base case ({fmtByFormat(sensBaseCellValue, sensOutcomeRatio ? sensOutcomeRatio.format : 'number')}, both inputs at their own historical average growth) — compare it against {sensOutcomeRatio ? sensOutcomeRatio.num : ''} running soft while {sensOutcomeRatio ? sensOutcomeRatio.den : ''} runs hot, or vice versa. All cells compound over the same {genForecastHorizon}-period horizon as the chart above.
                        </div>
                      </>
                    )}
                  </Panel>
                </>
              )}
            </div>
          )}

          {tab === 'data' && (
            <div>
              {!customDataset ? (
                <EmptyState icon={Table2} title="No data loaded" body="Upload a file from the Dashboard tab to browse it here — searchable, sortable, and exportable." />
              ) : (
                <>
                  <SectionHeader title="Transaction Data" subtitle={customDataset.name + ' — ' + (genFilteredRows.length === customDataset.rows.length ? 'all ' + customDataset.rows.length + ' rows.' : genFilteredRows.length + ' of ' + customDataset.rows.length + ' rows (filtered).')} />
                  <ActiveFiltersBar filters={genFilters} onRemove={toggleFilterValue} onClearAll={() => setGenFilters({})} />
                  <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 14, flexWrap: 'wrap' }}>
                    <div style={{ position: 'relative', flex: '0 1 280px' }}>
                      <Search size={14} color={COLORS.muted} style={{ position: 'absolute', left: 10, top: 10 }} />
                      <input value={search} onChange={e => { setSearch(e.target.value); setPage(0); }}
                        placeholder="Search any column…"
                        style={{ width: '100%', fontFamily: SANS, fontSize: 13, padding: '8px 10px 8px 30px', borderRadius: 8, border: `1px solid ${COLORS.border}` }} />
                    </div>
                    <span style={{ fontFamily: MONO, fontSize: 12, color: COLORS.muted }}>{tableRows.length} rows</span>
                    <button onClick={exportCSV} disabled={!tableRows.length}
                      style={{ display: 'flex', alignItems: 'center', gap: 6, fontFamily: SANS, fontSize: 12, fontWeight: 600, color: COLORS.ink, background: '#FFFFFF', border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: '7px 12px', cursor: tableRows.length ? 'pointer' : 'default', opacity: tableRows.length ? 1 : 0.5, marginLeft: 'auto' }}>
                      <Download size={13} /> Export CSV
                    </button>
                  </div>
                  <Panel style={{ overflowX: 'auto', padding: 0 }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontFamily: SANS, fontSize: 12 }}>
                      <thead>
                        <tr style={{ borderBottom: `1px solid ${COLORS.border}`, background: COLORS.bg }}>
                          {customDataset.fields.map(f => (
                            <th key={f} onClick={() => toggleSort(f)} style={{ textAlign: 'left', padding: '9px 10px', color: COLORS.ink, fontWeight: 700, fontSize: 10.5, textTransform: 'uppercase', letterSpacing: '0.03em', cursor: 'pointer', userSelect: 'none', whiteSpace: 'nowrap' }}>
                              <span style={{ display: 'inline-flex', alignItems: 'center', gap: 3 }}>{f}{sortKey === f ? <ArrowUpDown size={11} /> : null}</span>
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {pageRows.map((row, i) => (
                          <tr key={i} style={{ borderBottom: `1px solid ${COLORS.border}`, background: i % 2 ? COLORS.bg : '#FFFFFF' }}>
                            {customDataset.fields.map(f => (
                              <td key={f} style={{ padding: '8px 10px', fontFamily: typeof row[f] === 'number' ? MONO : SANS, whiteSpace: 'nowrap', textAlign: typeof row[f] === 'number' ? 'right' : 'left' }}>
                                {row[f] === null || row[f] === undefined || row[f] === '' ? '—' : (typeof row[f] === 'number' ? row[f].toLocaleString('en-US', { maximumFractionDigits: 2 }) : String(row[f]))}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </Panel>
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 14, marginTop: 14 }}>
                    <button onClick={() => setPage(p => Math.max(0, p - 1))} disabled={clampedPage === 0}
                      style={{ display: 'flex', alignItems: 'center', gap: 4, fontFamily: SANS, fontSize: 12.5, fontWeight: 600, color: clampedPage === 0 ? COLORS.border : COLORS.ink, background: 'transparent', border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: '6px 12px', cursor: clampedPage === 0 ? 'default' : 'pointer' }}>
                      <ChevronLeft size={14} /> Prev
                    </button>
                    <span style={{ fontFamily: MONO, fontSize: 12, color: COLORS.muted }}>Page {clampedPage + 1} of {pageCount}</span>
                    <button onClick={() => setPage(p => Math.min(pageCount - 1, p + 1))} disabled={clampedPage >= pageCount - 1}
                      style={{ display: 'flex', alignItems: 'center', gap: 4, fontFamily: SANS, fontSize: 12.5, fontWeight: 600, color: clampedPage >= pageCount - 1 ? COLORS.border : COLORS.ink, background: 'transparent', border: `1px solid ${COLORS.border}`, borderRadius: 8, padding: '6px 12px', cursor: clampedPage >= pageCount - 1 ? 'default' : 'pointer' }}>
                      Next <ChevronRight size={14} />
                    </button>
                  </div>
                </>
              )}
            </div>
          )}

          {tab === 'compare' && (
            <div>
              {!customDataset ? (
                <EmptyState icon={GitCompare} title="No data loaded" body="Upload your primary file from the Dashboard tab first, then come back here to compare it against a second file." />
              ) : (
                <>
                  <SectionHeader title="Two-File Comparison"
                    subtitle={fileCompareDataset ? customDataset.name + ' vs ' + fileCompareDataset.name + ' — matched by exact column name.' : 'Upload a second file to compare its totals against ' + customDataset.name + '.'} />

                  {!fileCompareDataset ? (
                    <div onDrop={onFileCompareDrop} onDragOver={onFileCompareDragOver} onDragLeave={onFileCompareDragLeave}>
                      <Panel style={{ textAlign: 'center', padding: '40px 24px', border: `2px dashed ${fileCompareDragOver ? COLORS.accent : COLORS.border}`, background: fileCompareDragOver ? COLORS.accentSoft : COLORS.panel }}>
                        <div style={{ width: 52, height: 52, borderRadius: 14, background: COLORS.accentSoft, display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 14px auto' }}>
                          <FileUp size={24} color={COLORS.accent} />
                        </div>
                        <div style={{ fontFamily: SANS, fontWeight: 700, fontSize: 15, color: COLORS.ink, marginBottom: 6 }}>Upload a second file to compare</div>
                        <div style={{ fontFamily: SANS, fontSize: 12.5, color: COLORS.muted, maxWidth: 440, margin: '0 auto 18px auto', lineHeight: 1.5 }}>
                          Its columns are matched to {customDataset.name}'s active metrics and dimensions by exact name — no fuzzy matching, so what gets compared is always auditable from the column names alone.
                        </div>
                        <button onClick={() => fileCompareInputRef.current && fileCompareInputRef.current.click()}
                          style={{ display: 'inline-flex', alignItems: 'center', gap: 8, fontFamily: SANS, fontSize: 13, fontWeight: 700, color: '#FFFFFF', background: COLORS.accent, border: 'none', borderRadius: 10, padding: '10px 20px', cursor: 'pointer' }}>
                          <Upload size={14} /> Choose a file
                        </button>
                        <div style={{ fontFamily: SANS, fontSize: 11, color: COLORS.muted, marginTop: 12 }}>CSV · TSV · JSON · Excel — or drag it onto this card</div>
                      </Panel>
                      {fileCompareUploadError ? (
                        <div style={{ marginTop: 14, display: 'flex', alignItems: 'center', gap: 8, background: '#FDEEF0', border: `1px solid ${COLORS.rose}`, color: '#8A1F30', borderRadius: 10, padding: '10px 14px', fontFamily: SANS, fontSize: 12.5 }}>
                          <AlertCircle size={15} /> {fileCompareUploadError}
                        </div>
                      ) : null}
                    </div>
                  ) : (
                    <>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 16, flexWrap: 'wrap' }}>
                        <span style={{ display: 'flex', alignItems: 'center', gap: 6, fontFamily: SANS, fontSize: 12, fontWeight: 600, padding: '6px 12px', borderRadius: 999, border: `1px solid ${COLORS.border}`, background: '#FFFFFF', color: COLORS.ink }}>
                          <FileSpreadsheet size={12} color={COLORS.muted} /> {customDataset.name}
                        </span>
                        <GitCompare size={14} color={COLORS.muted} />
                        <span style={{ display: 'flex', alignItems: 'center', gap: 6, fontFamily: SANS, fontSize: 12, fontWeight: 600, padding: '6px 12px', borderRadius: 999, border: `1px solid ${COLORS.accent}`, background: COLORS.accent, color: '#FFFFFF' }}>
                          <FileSpreadsheet size={12} /> {fileCompareDataset.name}
                        </span>
                        <button onClick={removeFileCompareDataset} title="Remove second file" style={{ background: 'transparent', border: 'none', cursor: 'pointer', color: COLORS.muted }}><X size={14} /></button>
                        <span style={{ fontFamily: MONO, fontSize: 11.5, color: COLORS.muted, marginLeft: 'auto' }}>{fileCompareDataset.rows.length} rows in {fileCompareDataset.name}</span>
                      </div>

                      {(fileCompareUnmatchedMetrics.length || fileCompareUnmatchedRatios.length || fileCompareUnmatchedDimensions.length) ? (
                        <div style={{ display: 'flex', gap: 10, marginBottom: 16, padding: '10px 14px', background: '#FFF8EB', border: `1px solid ${COLORS.amber}`, borderRadius: 10 }}>
                          <AlertTriangle size={15} color={COLORS.amberText} style={{ flexShrink: 0, marginTop: 1 }} />
                          <div style={{ fontFamily: SANS, fontSize: 12, color: '#6B4A16', lineHeight: 1.5 }}>
                            <strong>Not compared — matching usable column not found in {fileCompareDataset.name}: </strong>
                            {[...fileCompareUnmatchedMetrics.map(m => m.key), ...fileCompareUnmatchedRatios.map(r => r.label + ' (needs ' + r.num + ' & ' + r.den + ')'), ...fileCompareUnmatchedDimensions.map(d => d.key)].join(', ')}
                          </div>
                        </div>
                      ) : null}

                      {!fileCompareMetricRows.length ? (
                        <EmptyState icon={AlertTriangle} title="No matching columns" body={"None of " + customDataset.name + "'s active metrics exist by the same name in " + fileCompareDataset.name + ". Turn on matching columns in Configure analysis on the Dashboard tab, or upload a file that shares column names."} />
                      ) : (
                        <>
                          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3" style={{ gap: 14, marginBottom: 20 }}>
                            {fileCompareMetricRows.map(m => {
                              const va = fileCompareTotalsA ? fileCompareTotalsA[m.key] : 0;
                              const vb = fileCompareTotalsB ? fileCompareTotalsB[m.key] : 0;
                              const pct = (typeof va === 'number' && va !== 0) ? (vb - va) / Math.abs(va) : 0;
                              return (
                                <Panel key={m.key}>
                                  <div style={{ fontFamily: SANS, fontSize: 10.5, fontWeight: 600, color: COLORS.muted, textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: 10 }}>{m.label}</div>
                                  <div style={{ display: 'flex', alignItems: 'baseline', justifyContent: 'space-between', gap: 8 }}>
                                    <div>
                                      <div style={{ fontFamily: SANS, fontSize: 10, color: COLORS.muted, marginBottom: 2, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 120 }}>{customDataset.name}</div>
                                      <div style={{ fontFamily: MONO, fontSize: 18, fontWeight: 600, color: COLORS.ink }}>{fmtByFormat(va, m.format)}</div>
                                    </div>
                                    <div style={{ textAlign: 'right' }}>
                                      <div style={{ fontFamily: SANS, fontSize: 10, color: COLORS.muted, marginBottom: 2, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 120 }}>{fileCompareDataset.name}</div>
                                      <div style={{ fontFamily: MONO, fontSize: 18, fontWeight: 600, color: COLORS.accent }}>{fmtByFormat(vb, m.format)}</div>
                                    </div>
                                  </div>
                                  <div style={{ marginTop: 10, display: 'flex', justifyContent: 'flex-end' }}><DeltaBadge pct={pct} size={12} /></div>
                                </Panel>
                              );
                            })}
                          </div>

                          {fileCompareMatchedDimensions.length ? (
                            <Panel style={{ overflowX: 'auto' }}>
                              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap', marginBottom: 16 }}>
                                <span style={{ fontFamily: SANS, fontWeight: 700, fontSize: 13.5, color: COLORS.ink }}>Per-segment delta</span>
                                <Select label="Breakdown by" value={fileCompareBreakdownDim} onChange={setCompareBreakdownDim} options={fileCompareMatchedDimensions.map(d => d.key)} />
                              </div>
                              {fileCompareSegmentTable.length ? (
                                <div style={{ overflowX: 'auto' }}>
                                  <table style={{ width: '100%', borderCollapse: 'collapse', fontFamily: SANS, fontSize: 12.5 }}>
                                    <thead>
                                      <tr style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                                        <th style={{ textAlign: 'left', padding: '8px 10px', color: COLORS.muted, fontWeight: 600, fontSize: 11, textTransform: 'uppercase' }}>{fileCompareBreakdownDim}</th>
                                        {fileCompareMetricRows.map(m => (
                                          <React.Fragment key={m.key}>
                                            <th style={{ textAlign: 'right', padding: '8px 10px', color: COLORS.muted, fontWeight: 600, fontSize: 11, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{m.label} (A)</th>
                                            <th style={{ textAlign: 'right', padding: '8px 10px', color: COLORS.muted, fontWeight: 600, fontSize: 11, textTransform: 'uppercase', whiteSpace: 'nowrap' }}>{m.label} (B)</th>
                                            <th style={{ textAlign: 'right', padding: '8px 10px', color: COLORS.muted, fontWeight: 600, fontSize: 11, textTransform: 'uppercase' }}>Δ</th>
                                          </React.Fragment>
                                        ))}
                                      </tr>
                                    </thead>
                                    <tbody>
                                      {fileCompareSegmentTable.map(row => (
                                        <tr key={row.name} style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                                          <td style={{ padding: '9px 10px', fontWeight: 600, color: COLORS.ink, whiteSpace: 'nowrap' }}>{row.name}</td>
                                          {fileCompareMetricRows.map(m => {
                                            const va = row.a ? row.a[m.key] : null;
                                            const vb = row.b ? row.b[m.key] : null;
                                            const pct = (typeof va === 'number' && va !== 0 && typeof vb === 'number') ? (vb - va) / Math.abs(va) : 0;
                                            return (
                                              <React.Fragment key={m.key}>
                                                <td style={{ padding: '9px 10px', textAlign: 'right', fontFamily: MONO, color: row.a ? COLORS.ink : COLORS.border }}>{row.a ? fmtByFormat(va, m.format) : '—'}</td>
                                                <td style={{ padding: '9px 10px', textAlign: 'right', fontFamily: MONO, color: row.b ? COLORS.ink : COLORS.border }}>{row.b ? fmtByFormat(vb, m.format) : '—'}</td>
                                                <td style={{ padding: '9px 10px', textAlign: 'right' }}>{(row.a && row.b) ? <DeltaBadge pct={pct} size={11} /> : <span style={{ color: COLORS.border }}>—</span>}</td>
                                              </React.Fragment>
                                            );
                                          })}
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>
                              ) : null}
                            </Panel>
                          ) : (
                            <div style={{ fontFamily: SANS, fontSize: 12, color: COLORS.muted, textAlign: 'center', padding: '12px 0' }}>
                              No matching category columns to segment by — showing overall totals only.
                            </div>
                          )}
                        </>
                      )}
                    </>
                  )}
                </>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
