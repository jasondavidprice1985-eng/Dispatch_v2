/**
 * Dispatch v2 — filters.js
 *
 * Renders and wires the filter chip bar.
 * Chips available:
 *   - One chip per CM (from cm-lookup, excluding system codes)
 *   - One chip per type present in the loaded data
 *   - A clear button when any filter is active
 *
 * Data flow:
 *   app.subscribe → re-renders chips to reflect current state
 *   Chip tap → calls app.setCM / app.toggleType / app.clearFilters
 *   app state change → calendar.js re-renders automatically
 *
 * Chips are built from live data — if no ZCR exists in the report,
 * the Remedials chip does not appear. No hardcoding.
 */

'use strict';

// ─── Module state ─────────────────────────────────────────────────────────────

let _elChipBar   = null;
let _unsubscribe = null;
let _initialised = false;

// ─── Init / destroy ───────────────────────────────────────────────────────────

/**
 * Initialise the filter bar.
 * Resolves DOM ref, subscribes to app state, renders chips.
 * Safe to call multiple times.
 */
function initFilters() {
  if (_unsubscribe) { _unsubscribe(); _unsubscribe = null; }

  _elChipBar = document.getElementById('filter-chips');
  if (!_elChipBar) {
    console.error('filters.js: #filter-chips element not found.');
    return;
  }

  _initialised = true;
  _unsubscribe = subscribe(_onStateChange);
  _render();
}

function destroyFilters() {
  if (_unsubscribe) { _unsubscribe(); _unsubscribe = null; }
  _initialised = false;
}

// ─── State subscription ───────────────────────────────────────────────────────

function _onStateChange(state) {
  if (!_initialised) return;
  _render(state);
}

// ─── Render ───────────────────────────────────────────────────────────────────

function _render(state) {
  if (!_elChipBar) return;
  state = state ?? getState();

  // Nothing to show until data is loaded
  if (!state.ready) {
    _elChipBar.innerHTML = '';
    return;
  }

  const { filters, indexes, cmLookup } = state;
  const html = [];

  // ── CM chips ───────────────────────────────────────────────────────────────
  // One chip per real CM present in the data
  // Ordered by name; active CM chip is highlighted
  const cmsInData = _getCMsInData(indexes, cmLookup);

  for (const { code, name } of cmsInData) {
    const isActive = filters.cm === code;
    html.push(_cmChip(code, name, isActive));
  }

  // ── Separator ─────────────────────────────────────────────────────────────
  if (cmsInData.length > 0) {
    html.push('<div class="chip-sep"></div>');
  }

  // ── Type chips ─────────────────────────────────────────────────────────────
  // Only show chips for types actually present in the loaded data
  const typesInData = _getTypesInData(indexes);
  const typeOrder   = ['ZCD','ZCR','ZCA','ZCC','ZBC'];

  for (const type of typeOrder) {
    if (!typesInData.has(type)) continue;
    const isActive = filters.types.length === 0 || filters.types.includes(type);
    html.push(_typeChip(type, isActive));
  }

  // ── Clear button ───────────────────────────────────────────────────────────
  const hasActiveFilter = filters.cm || filters.types.length > 0 || filters.search;
  if (hasActiveFilter) {
    html.push(_clearChip());
  }

  _elChipBar.innerHTML = html.join('');

  // Wire click handlers
  _elChipBar.querySelectorAll('[data-action]').forEach(el => {
    el.addEventListener('click', _handleChipTap);
  });
}

// ─── Chip builders ────────────────────────────────────────────────────────────

function _cmChip(code, name, isActive) {
  const short = _shortName(name);
  return `
    <button type="button"
      class="chip${isActive ? ' chip--active chip--cm' : ''}"
      data-action="cm"
      data-code="${_esc(code)}"
      aria-pressed="${isActive}"
      aria-label="${isActive ? 'Remove filter: ' : 'Filter by '}${_esc(name)}">
      ${_esc(short)}
    </button>`;
}

function _typeChip(type, isActive) {
  const label  = TYPE_LABELS[type]  ?? type;
  const colour = TYPE_COLOURS[type] ?? '#64748b';
  return `
    <button type="button"
      class="chip${isActive ? ' chip--active' : ''}"
      data-action="type"
      data-type="${_esc(type)}"
      aria-pressed="${isActive}"
      aria-label="${isActive ? 'Remove filter: ' : 'Filter by '}${_esc(label)}">
      <span class="chip-dot" style="background:${colour}" aria-hidden="true"></span>
      ${_esc(label)}
    </button>`;
}

function _clearChip() {
  return `
    <button type="button"
      class="chip chip--clear"
      data-action="clear"
      aria-label="Clear all filters">
      ✕ Clear
    </button>`;
}

// ─── Event handling ───────────────────────────────────────────────────────────

function _handleChipTap(e) {
  const el     = e.currentTarget;
  const action = el.dataset.action;

  switch (action) {
    case 'cm': {
      const code = el.dataset.code;
      if (getActiveCM() === code) {
        clearCM();     // tap active CM chip → deselect (show all)
      } else {
        setCM(code);   // tap inactive CM chip → select
      }
      break;
    }
    case 'type': {
      toggleType(el.dataset.type);
      break;
    }
    case 'clear': {
      clearFilters();
      clearCM();
      break;
    }
  }
}

// ─── Data helpers ─────────────────────────────────────────────────────────────

/**
 * Get real CMs present in the loaded data, sorted by name.
 * Excludes system codes (Supply Only, Unassigned, Contracts CC).
 */
function _getCMsInData(indexes, cmLookup) {
  if (!indexes?.byCM) return [];

  const SYSTEM = new Set(['ASM105','ASM135','TBC','CONTCC']);

  return [...indexes.byCM.keys()]
    .filter(code => !SYSTEM.has(code))
    .map(code => ({ code, name: cmLookup[code] ?? code }))
    .sort((a, b) => a.name.localeCompare(b.name));
}

/**
 * Get document types present in the loaded data.
 * Returns a Set of type codes e.g. { 'ZCD', 'ZCR', 'ZCA' }
 */
function _getTypesInData(indexes) {
  if (!indexes?.byType) return new Set();
  return new Set(indexes.byType.keys());
}

// ─── Utilities ────────────────────────────────────────────────────────────────

/**
 * Shorten a full name for chip display.
 * "Mark Fowler" → "M. Fowler"
 * "Supply Only"  → "Supply Only" (no first name to shorten)
 */
function _shortName(name) {
  const parts = name.trim().split(' ');
  if (parts.length < 2) return name;
  return `${parts[0][0]}. ${parts.slice(1).join(' ')}`;
}

function _esc(str) {
  return String(str ?? '')
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;');
}

// ─── Exports ─────────────────────────────────────────────────────────────────

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    initFilters,
    destroyFilters,
    // Exported for testing
    _getCMsInData,
    _getTypesInData,
    _shortName,
    _cmChip,
    _typeChip,
    _clearChip,
  };
}
