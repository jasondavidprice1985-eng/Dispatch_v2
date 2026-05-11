/**
 * Dispatch v2 — app.js
 *
 * Single source of truth for all application state.
 *
 * Rules:
 *   - Only app.js mutates state
 *   - All other modules read state and call app actions
 *   - UI never holds its own data copies
 *   - Subscribers are notified after every state change
 */

'use strict';

// ─── CM codes excluded from the "Who are you?" picker ────────────────────────
// These appear in the data for analysis but are not real field CMs.

const SYSTEM_CM_CODES = new Set(['ASM105', 'ASM135', 'TBC', 'CONTCC']);

// ─── localStorage key ────────────────────────────────────────────────────────

const STORAGE_KEY_CM = 'dispatch_cm';

// ─── State ───────────────────────────────────────────────────────────────────

const state = {
  // ── Data — set once on load, never mutated ──────────────────────────────
  deliveries:  [],      // Delivery[] from parser.parseDeliveries
  indexes:     {},      // { byCM, byAccount, byDate, byType } from parser.buildIndexes
  cmLookup:    {},      // { "FOWLERM": "Mark Fowler", ... } from cm-lookup.json
  filename:    '',      // "Delivery_Report_06_03_2026.xlsx"
  loadedAt:    null,    // Date — when the report was loaded
  summary:     null,    // ParseSummary from parser

  // ── Active filters ───────────────────────────────────────────────────────
  filters: {
    cm:     null,       // CM code string or null (null = all CMs)
    types:  [],         // [] = all types shown
    search: '',
  },

  // ── Derived — rebuilt whenever filters change ────────────────────────────
  filtered:    [],      // filterDeliveries(deliveries, filters)

  // ── UI ───────────────────────────────────────────────────────────────────
  selectedDate: null,   // ISO string "2026-06-10" — day tapped on calendar
  loading:      false,
  error:        null,
  ready:        false,  // true once deliveries are loaded and parsed
};

// ─── Subscribers ─────────────────────────────────────────────────────────────

const subscribers = new Set();

/**
 * Register a callback to be called after every state change.
 * Returns an unsubscribe function.
 *
 * @param   {Function} fn  — called with a shallow copy of state
 * @returns {Function}     — call to unsubscribe
 */
function subscribe(fn) {
  subscribers.add(fn);
  return () => subscribers.delete(fn);
}

/**
 * Notify all subscribers with a shallow copy of current state.
 * Internal — called after every state mutation.
 */
function notify() {
  const snapshot = { ...state, filters: { ...state.filters } };
  for (const fn of subscribers) {
    try { fn(snapshot); }
    catch (err) { console.error('app.js subscriber error:', err); }
  }
}

// ─── Load ─────────────────────────────────────────────────────────────────────

/**
 * Load and parse a report from a user-selected File.
 * Wires api.loadFromFile → parser.parseDeliveries → state.
 *
 * @param {File} file
 */
async function load(file) {
  _setLoading(true);

  try {
    const { rows, filename } = await loadFromFile(file);
    _commitRows(rows, filename);
  } catch (err) {
    _setError(err.isApiError ? err.message : 'Something went wrong loading the file.');
  } finally {
    _setLoading(false);
  }
}

/**
 * Load and parse a report from SharePoint via N8N.
 * Wires api.loadFromSharePoint → parser.parseDeliveries → state.
 */
async function loadSharePoint() {
  _setLoading(true);

  try {
    const { rows, filename } = await loadFromSharePoint();
    _commitRows(rows, filename);
  } catch (err) {
    _setError(err.isApiError ? err.message : 'Could not load from SharePoint.');
  } finally {
    _setLoading(false);
  }
}

/**
 * Internal — parse rows and commit to state.
 * Called by both load() and loadSharePoint().
 *
 * @param {Array[]} rows
 * @param {string}  filename
 */
function _commitRows(rows, filename) {
  const { deliveries, summary } = parseDeliveries(rows, state.cmLookup);
  const indexes = buildIndexes(deliveries);

  state.deliveries  = deliveries;
  state.indexes     = indexes;
  state.filename    = filename;
  state.loadedAt    = new Date();
  state.summary     = summary;
  state.error       = null;
  state.ready       = true;

  // Re-apply current filters against fresh data
  _rebuildFiltered();
  notify();
}

// ─── CM Identity ──────────────────────────────────────────────────────────────

/**
 * Initialise the app — load CM lookup, restore saved CM, notify subscribers.
 * Call once on DOMContentLoaded.
 */
async function init() {
  try {
    state.cmLookup = await loadCMLookup();
  } catch (_) {
    state.cmLookup = {};
  }

  // Restore saved CM from localStorage.
  // If the lookup loaded successfully, verify the CM exists in it.
  // If the lookup failed entirely (network glitch etc), trust the saved code
  // so the user stays on their deliveries — they'll see their code instead of
  // their name until the lookup loads, which is better than losing their filter.
  const saved = localStorage.getItem(STORAGE_KEY_CM);
  if (saved) {
    const isKnownCM    = Boolean(state.cmLookup[saved]);
    const lookupFailed = Object.keys(state.cmLookup).length === 0;
    if (isKnownCM || lookupFailed) {
      state.filters.cm = saved;
    }
  }

  notify();
}

/**
 * Set the active CM — called when user picks their name for the first time,
 * or changes it later.
 * Persists to localStorage so it survives page reloads.
 *
 * @param {string} cmCode  e.g. "FOWLERM"
 */
function setCM(cmCode) {
  if (!cmCode || !state.cmLookup[cmCode]) return;
  state.filters.cm = cmCode;
  try { localStorage.setItem(STORAGE_KEY_CM, cmCode); } catch (_) {}
  _rebuildFiltered();
  notify();
}

/**
 * Clear the saved CM — shows all deliveries.
 * Removes from localStorage so the user stays on "all" if they reload.
 */
function clearCM() {
  state.filters.cm = null;
  try { localStorage.removeItem(STORAGE_KEY_CM); } catch (_) {}
  _rebuildFiltered();
  notify();
}

/**
 * True if the user has not yet picked their CM.
 * The UI shows the "Who are you?" picker when this is true and data is loaded.
 *
 * @returns {boolean}
 */
function needsCMPicker() {
  return !state.filters.cm;
}

/**
 * Returns the list of real CMs for the "Who are you?" picker.
 * Excludes system codes (Supply Only, Unassigned, Contracts CC).
 *
 * @returns {{ code: string, name: string }[]}  sorted by name
 */
function getPickerCMs() {
  return Object.entries(state.cmLookup)
    .filter(([code]) => !SYSTEM_CM_CODES.has(code))
    .map(([code, name]) => ({ code, name }))
    .sort((a, b) => a.name.localeCompare(b.name));
}

// ─── Filters ──────────────────────────────────────────────────────────────────

/**
 * Toggle a type filter on or off.
 * If types array is empty — all types are shown.
 *
 * @param {string} typeCode  e.g. "ZCR"
 */
function toggleType(typeCode) {
  const types = state.filters.types;
  const idx   = types.indexOf(typeCode);
  if (idx === -1) types.push(typeCode);
  else            types.splice(idx, 1);
  _rebuildFiltered();
  notify();
}

/**
 * Set free-text search string.
 * Matches against name, plot, doc, cmName, typeLabel and raw type code.
 *
 * @param {string} text
 */
function setSearch(text) {
  state.filters.search = text ?? '';
  _rebuildFiltered();
  notify();
}

/**
 * Clear all filters except the active CM.
 * CM is identity, not a filter — use clearCM() to clear that separately.
 */
function clearFilters() {
  state.filters.types  = [];
  state.filters.search = '';
  _rebuildFiltered();
  notify();
}

// ─── Calendar ─────────────────────────────────────────────────────────────────

/**
 * Set the selected calendar date.
 * The day detail panel re-renders based on this.
 *
 * @param {string|null} iso  e.g. "2026-06-10" or null to deselect
 */
function selectDate(iso) {
  state.selectedDate = iso ?? null;
  notify();
}

/**
 * Get deliveries for the currently selected date.
 * Returns from the filtered set — respects active CM and type filters.
 *
 * @returns {Delivery[]}
 */
function getSelectedDateDeliveries() {
  if (!state.selectedDate) return [];
  return state.filtered.filter(d => d.deliveryDateISO === state.selectedDate);
}

// ─── Internal ─────────────────────────────────────────────────────────────────

function _rebuildFiltered() {
  state.filtered = filterDeliveries(state.deliveries, state.filters);
}

function _setLoading(val) {
  state.loading = val;
  if (val) state.error = null;
  notify();
}

function _setError(message) {
  state.error = message;
  state.ready = false;
  notify();
}

// ─── Read-only state accessors ────────────────────────────────────────────────
// Other modules use these rather than reading state directly.
// Keeps mutation control in app.js.

function getState()        { return { ...state, filters: { ...state.filters } }; }
function getDeliveries()   { return state.deliveries; }
function getFiltered()     { return state.filtered; }
function getIndexes()      { return state.indexes; }
function getCMLookup()     { return state.cmLookup; }
function getFilters()      { return { ...state.filters }; }
function getActiveCM()     { return state.filters.cm; }
function getActiveCMName() { return state.cmLookup[state.filters.cm] ?? state.filters.cm ?? 'Everyone'; }
function isLoading()       { return state.loading; }
function isReady()         { return state.ready; }
function getError()        { return state.error; }
function getFilename()     { return state.filename; }
function getLoadedAt()     { return state.loadedAt; }
function getSummary()      { return state.summary; }

// ─── Exports ─────────────────────────────────────────────────────────────────

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    // Lifecycle
    init,
    // Load
    load,
    loadSharePoint,
    // CM identity
    setCM,
    clearCM,
    needsCMPicker,
    getPickerCMs,
    // Filters
    toggleType,
    setSearch,
    clearFilters,
    // Calendar
    selectDate,
    getSelectedDateDeliveries,
    // Subscriptions
    subscribe,
    // Accessors
    getState,
    getDeliveries,
    getFiltered,
    getIndexes,
    getCMLookup,
    getFilters,
    getActiveCM,
    getActiveCMName,
    isLoading,
    isReady,
    getError,
    getFilename,
    getLoadedAt,
    getSummary,
    // Constants
    SYSTEM_CM_CODES,
  };
}
