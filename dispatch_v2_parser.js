/**
 * Dispatch v2 — parser.js
 *
 * Transforms raw SAP delivery report rows into clean Delivery objects.
 *
 * Input:  rawRows  — array of row arrays, header row already removed
 *         cmLookup — optional map of CM codes to full names
 *                    { "FOWLERM": "Mark Fowler", ... }
 *
 * Output: { deliveries: Delivery[], summary: ParseSummary }
 *
 * One Delivery per unique Sales Document.
 * 42,347 rows → 1,480 deliveries on a real SAP export.
 *
 * Notes on SAP data quirks handled here:
 *  - Quantity field contains price (not count) for Silestone and fitting lines.
 *    This is expected — show the value as-is, team understands the convention.
 *  - qty = 0 means a credit or reversal line — flagged on the delivery.
 *  - ZCS and ZFC are excluded from TYPE_LABELS intentionally — they are not
 *    delivery types CMs schedule. They pass through as unknown types and appear
 *    in qualityNotices.
 */

'use strict';

// ─── Column map ───────────────────────────────────────────────────────────────
// Confirmed against Delivery_Report_06_03_2026.xlsx (24 cols, 42,347 rows).

const COL = Object.freeze({
  type:         0,   // SaTy             ZCD / ZCR / ZCA / ZCC / ZBC
  doc:          1,   // Sales doc.        unique SAP document number
  created:      2,   // Created on        date order raised in SAP
  account:      3,   // Sold-to pt        customer account code
  poNo:         4,   // Purchase order    contains plot ref as first token
  deliveryDate: 5,   // Req.dlv.dt        delivery date — drives the calendar
  name:         6,   // Name 1            account display name
  material:     7,   // Material          SAP material code
  description:  8,   // Description       human-readable material name
  qty:          9,   // Order Quantity    unit count OR price (Silestone/fittings)
  cm:           10,  // User              CM code e.g. FOWLERM
  MIN_LENGTH:   11,  // minimum valid row length
});

// ─── Type definitions ─────────────────────────────────────────────────────────

const TYPE_LABELS = Object.freeze({
  ZCD: 'Kitchen',
  ZCR: 'Remedials',
  ZCA: 'Utility / Silestone',
  ZCC: 'Additions',
  ZBC: 'Back Order',
});

const TYPE_COLOURS = Object.freeze({
  ZCD: '#2563eb',
  ZCR: '#ef4444',
  ZCA: '#10b981',
  ZCC: '#8b5cf6',
  ZBC: '#f59e0b',
});

// ─── Pure helpers ─────────────────────────────────────────────────────────────

/**
 * Normalise any value to a clean string.
 * Trims edges, collapses all whitespace including tabs and newlines.
 */
function norm(v) {
  return String(v ?? '').trim().replace(/[\s\t\n\r]+/g, ' ');
}

/**
 * Extract the plot reference from a Purchase Order field.
 * Format: "54 DHL0025/82 MILLER" — plot is the first whitespace-separated token.
 * Returns empty string if the token is absent or implausibly long.
 *
 * Validated against all 42,347 real rows — all known formats pass.
 */
function extractPlot(poNo) {
  if (!poNo) return '';
  const first = String(poNo).trim().split(/\s+/)[0] ?? '';
  return /^[\w.\-/]+$/.test(first) && first.length <= 12 ? first : '';
}

/**
 * Detect the likely date string format from a sample of raw column values.
 * Returns 'UK' (dd/mm/yyyy) or 'ISO' (yyyy-mm-dd).
 *
 * Only considers string values — Date objects and serials are unambiguous.
 * Votes: if first token > 31 it must be a year (ISO); if > 12 it must be a day (UK).
 * Defaults to 'UK' on a tie or insufficient sample — correct for SAP exports.
 */
function detectDateFormat(rawRows, colIndex, sampleSize = 50) {
  let ukVotes = 0, isoVotes = 0, checked = 0;
  for (const row of rawRows) {
    if (checked >= sampleSize) break;
    const val = row?.[colIndex];
    if (!val || typeof val !== 'string') continue;
    const m = val.trim().match(/^(\d{1,4})[\/\-.]/);
    if (!m) continue;
    const a = Number(m[1]);
    if (a > 31)      isoVotes++;
    else if (a > 12) ukVotes++;
    checked++;
  }
  return isoVotes > ukVotes ? 'ISO' : 'UK';
}

/**
 * Parse a single date value into a time-stripped Date object, or null.
 *
 * Handles:
 *   - Native Date objects (from XLSX libraries that parse dates automatically)
 *   - Excel serial numbers (days since 1900-01-00, adjusted for Lotus bug)
 *   - String dates in UK (dd/mm/yyyy) or ISO (yyyy-mm-dd) format
 *
 * The fmt hint is used only when the first token is ≤ 12 (ambiguous).
 * Years < 2000 are rejected as implausible for a delivery scheduling tool.
 */
function parseDate(value, fmt = 'UK') {
  if (value === null || value === undefined || value === '') return null;

  // Native Date from XLSX
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : stripTime(value);
  }

  // Excel serial number
  if (typeof value === 'number') {
    if (value <= 1000 || value >= 200000) return null;
    const ms = Math.round((value - 25569) * 86400 * 1000);
    const d  = new Date(ms);
    return isNaN(d.getTime()) ? null : stripTime(d);
  }

  // String date
  if (typeof value === 'string') {
    const m = value.trim().match(/^(\d{1,4})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
    if (!m) return null;

    const a = Number(m[1]), b = Number(m[2]), c = Number(m[3]);
    let day, month, year;

    if (a > 31) {
      // First token is too large to be a day — must be yyyy
      year = a; month = b; day = c;
    } else if (fmt === 'ISO') {
      // Caller detected ISO format — treat as year, but validate plausibility
      year = a; month = b; day = c;
      if (year < 2000) return null;   // reject implausible year (e.g. "05" as year)
    } else if (a > 12) {
      // First token > 12 — must be a day (UK dd/mm/yyyy)
      day = a; month = b; year = c;
    } else {
      // Ambiguous — use fmt hint, default UK
      if (fmt === 'UK') { day = a; month = b; year = c; }
      else              { year = a; month = b; day = c; if (year < 2000) return null; }
    }

    const fullYear = year < 100 ? 2000 + year : year;
    if (fullYear < 2000) return null;

    const d = new Date(fullYear, month - 1, day);

    // Reject silently wrapped dates — JS overflows invalid days into the next
    // month (e.g. Feb 30 → Mar 2). If the day changed, the input was invalid.
    if (d.getDate() !== day) return null;

    return isNaN(d.getTime()) ? null : d;
  }

  return null;
}

/**
 * Strip hours/minutes/seconds from a Date.
 * Uses local constructor to avoid timezone drift on calendar comparisons.
 */
function stripTime(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/**
 * Format a Date as YYYY-MM-DD string — used as calendar lookup key.
 */
function toISO(d) {
  if (!d) return null;
  const y  = d.getFullYear();
  const m  = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${dd}`;
}

// ─── parseDeliveries ──────────────────────────────────────────────────────────

/**
 * Main entry point.
 * Groups 42k+ raw rows into clean Delivery objects — one per Sales Document.
 *
 * @param   {Array[]}  rawRows    Row arrays with header already removed.
 * @param   {Object}   cmLookup  Optional { "FOWLERM": "Mark Fowler", ... }
 * @returns {{ deliveries: Delivery[], summary: ParseSummary }}
 */
function parseDeliveries(rawRows, cmLookup = {}) {
  const summary = {
    totalRows:      rawRows.length,
    skippedBadRow:  0,
    skippedNoDoc:   0,
    skippedNoType:  0,
    skippedNoDate:  0,
    unknownTypes:   new Set(),
    uniqueDocs:     0,
    dateRange:      { min: null, max: null },
    qualityNotices: [],
  };

  // Detect date formats once per column — not per row
  const dlvFmt     = detectDateFormat(rawRows, COL.deliveryDate, 50);
  const createdFmt = detectDateFormat(rawRows, COL.created, 50);

  // ── Pass 1: group rows by sales document ──────────────────────────────────
  const docMap = new Map();

  for (const row of rawRows) {
    // Reject truncated rows before touching any column
    if (!Array.isArray(row) || row.length < COL.MIN_LENGTH) {
      summary.skippedBadRow++;
      continue;
    }

    const type    = norm(row[COL.type]);
    const doc     = norm(row[COL.doc]);
    const dlvDate = parseDate(row[COL.deliveryDate], dlvFmt);

    if (!doc)     { summary.skippedNoDoc++;  continue; }
    if (!type)    { summary.skippedNoType++; continue; }
    if (!dlvDate) { summary.skippedNoDate++; continue; }

    if (type && !TYPE_LABELS[type]) summary.unknownTypes.add(type);

    const material    = norm(row[COL.material]);
    const description = norm(row[COL.description]);
    // Guard against NaN — Number('abc') passes ?? but fails isNaN
    const rawQty      = Number(row[COL.qty]);
    const qty         = isNaN(rawQty) ? 0 : rawQty;
    const account     = norm(row[COL.account]);
    const poNo        = norm(row[COL.poNo]);
    const name        = norm(row[COL.name]);
    const cmCode      = norm(row[COL.cm]).toUpperCase();
    const cmName      = cmLookup[cmCode] ?? cmCode;
    const created     = parseDate(row[COL.created], createdFmt);

    if (!docMap.has(doc)) {
      docMap.set(doc, {
        doc,
        type,
        typeLabel:       TYPE_LABELS[type] ?? type,
        colour:          TYPE_COLOURS[type] ?? '#64748b',
        account,
        name,
        plot:            extractPlot(poNo),
        poNo,
        deliveryDate:    dlvDate,
        deliveryDateISO: toISO(dlvDate),
        created,
        cmCode,
        cmName,
        lineCount:   0,
        lines:       [],
        hasZeroQty:  false,
        searchBlob:  '',   // built in pass 2
      });
    }

    const entry = docMap.get(doc);
    entry.lineCount++;
    entry.lines.push({ material, description, qty });

    if (qty === 0) entry.hasZeroQty = true;

    // Track date range for summary
    if (!summary.dateRange.min || dlvDate < summary.dateRange.min)
      summary.dateRange.min = dlvDate;
    if (!summary.dateRange.max || dlvDate > summary.dateRange.max)
      summary.dateRange.max = dlvDate;
  }

  // ── Pass 2: finalise, build search blobs, sort ────────────────────────────
  const deliveries = Array.from(docMap.values());

  for (const d of deliveries) {
    // Include both typeLabel and raw type code so searching "zcr" finds Remedials
    d.searchBlob = `${d.name} ${d.plot} ${d.doc} ${d.cmName} ${d.typeLabel} ${d.type}`
      .toLowerCase();
  }

  deliveries.sort((a, b) => {
    const diff = a.deliveryDate - b.deliveryDate;
    return diff !== 0 ? diff : a.name.localeCompare(b.name);
  });

  summary.uniqueDocs = deliveries.length;

  // ── Quality notices ───────────────────────────────────────────────────────
  const skipTotal = summary.skippedBadRow  + summary.skippedNoDoc +
                    summary.skippedNoType  + summary.skippedNoDate;

  if (skipTotal > 0) {
    const parts = [];
    if (summary.skippedBadRow  > 0) parts.push(`${summary.skippedBadRow} truncated rows`);
    if (summary.skippedNoDoc   > 0) parts.push(`${summary.skippedNoDoc} missing document number`);
    if (summary.skippedNoType  > 0) parts.push(`${summary.skippedNoType} missing type`);
    if (summary.skippedNoDate  > 0) parts.push(`${summary.skippedNoDate} missing delivery date`);
    summary.qualityNotices.push({
      kind:  skipTotal > rawRows.length * 0.05 ? 'warning' : 'info',
      title: `${skipTotal} row${skipTotal !== 1 ? 's' : ''} skipped`,
      body:  parts.join(' · '),
    });
  }

  if (summary.unknownTypes.size > 0) {
    summary.qualityNotices.push({
      kind:  'info',
      title: 'Unrecognised document types',
      body:  `Not in category list: ${[...summary.unknownTypes].join(', ')}. ` +
             `Shown on calendar without colour or filter chip.`,
    });
  }

  if (dlvFmt === 'ISO') {
    summary.qualityNotices.push({
      kind:  'info',
      title: 'ISO date format detected in delivery dates',
      body:  'Dates appear as YYYY-MM-DD. If dates look wrong check the SAP export settings.',
    });
  }

  return { deliveries, summary };
}

// ─── buildIndexes ─────────────────────────────────────────────────────────────

/**
 * Build O(N) lookup indexes. Call once after parseDeliveries.
 * Do not call on every filter change — the indexes are stable.
 *
 * byDate is the primary index for calendar dot rendering.
 * byCM, byAccount, byType feed the filter chips.
 */
function buildIndexes(deliveries) {
  const byCM      = new Map();
  const byAccount = new Map();
  const byDate    = new Map();
  const byType    = new Map();

  for (const d of deliveries) {
    if (!byCM.has(d.cmCode))       byCM.set(d.cmCode, []);
    if (!byAccount.has(d.account)) byAccount.set(d.account, []);
    if (!byType.has(d.type))       byType.set(d.type, []);

    byCM.get(d.cmCode).push(d);
    byAccount.get(d.account).push(d);
    byType.get(d.type).push(d);

    if (d.deliveryDateISO) {
      if (!byDate.has(d.deliveryDateISO)) byDate.set(d.deliveryDateISO, []);
      byDate.get(d.deliveryDateISO).push(d);
    }
  }

  return { byCM, byAccount, byDate, byType };
}

// ─── filterDeliveries ─────────────────────────────────────────────────────────

/**
 * Apply filters to a delivery array. All conditions are AND.
 *
 * Uses pre-computed searchBlob — no string building per call.
 * Searching "zcr" finds Remedials. Searching "kitchen" finds ZCD.
 *
 * @param   {Delivery[]} deliveries
 * @param   {Object}     filters
 * @param   {string}     [filters.cm]        CM code e.g. "FOWLERM"
 * @param   {string}     [filters.account]   Account code
 * @param   {string[]}   [filters.types]     Type codes e.g. ["ZCD","ZCR"]
 * @param   {Date}       [filters.dateFrom]
 * @param   {Date}       [filters.dateTo]
 * @param   {string}     [filters.search]    Free text
 * @returns {Delivery[]}
 */
function filterDeliveries(deliveries, filters = {}) {
  const { cm, account, types, dateFrom, dateTo, search } = filters;
  const searchLower = search ? search.toLowerCase().trim() : null;
  const hasTypes    = Array.isArray(types) && types.length > 0;

  return deliveries.filter(d => {
    if (cm          && d.cmCode  !== cm)                   return false;
    if (account     && d.account !== account)              return false;
    if (hasTypes    && !types.includes(d.type))            return false;
    if (dateFrom    && d.deliveryDate < dateFrom)          return false;
    if (dateTo      && d.deliveryDate > dateTo)            return false;
    if (searchLower && !d.searchBlob.includes(searchLower)) return false;
    return true;
  });
}

// ─── Exports ─────────────────────────────────────────────────────────────────

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    parseDeliveries,
    buildIndexes,
    filterDeliveries,
    detectDateFormat,
    extractPlot,
    parseDate,
    toISO,
    norm,
    TYPE_LABELS,
    TYPE_COLOURS,
    COL,
  };
}
