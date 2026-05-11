/**
 * Dispatch v2 — calendar.js
 *
 * Renders the split-view calendar and day detail panel.
 * No FullCalendar dependency — custom grid, full control.
 *
 * Layout:
 *   Top  — compact Mon-Fri grid, one cell per working day
 *          Each cell: heatmap bg + count number + type bar
 *   Bottom — day detail panel, deliveries grouped by type
 *
 * Data flow:
 *   app.subscribe(render) — re-renders on every state change
 *   app.getIndexes().byDate — O(1) day lookup, never scans full array
 *   app.selectDate(iso) — tap on a cell updates app state → triggers re-render
 *
 * Three design decisions baked in:
 *   1. Auto-select today on load; if empty advance to next delivery date
 *   2. Type bar minimum segment width 8% so small types are always visible
 *   3. Empty weeks compress to 24px; active weeks stay 52px
 */

'use strict';

// ─── Constants ────────────────────────────────────────────────────────────────

const MONTH_NAMES = [
  'January','February','March','April','May','June',
  'July','August','September','October','November','December',
];

// Type display order in day detail — urgent types always first
const TYPE_ORDER = ['ZCR','ZBC','ZCC','ZCA','ZCD'];

// Density thresholds calibrated to real data
// Average day ~7-8 deliveries, heavy Friday ~20+
const DENSITY_LEVEL = (count) => {
  if (count === 0)  return 0;
  if (count <= 5)   return 1;
  if (count <= 12)  return 2;
  if (count <= 20)  return 3;
  return 4;
};

// Minimum type bar segment — 8% so minority types are always visible
const BAR_MIN_PCT = 8;

// Cell heights
const CELL_HEIGHT_ACTIVE = 52;  // px — normal week
const CELL_HEIGHT_EMPTY  = 24;  // px — compressed empty week

// ─── Module state ─────────────────────────────────────────────────────────────

let _currentYear  = null;
let _currentMonth = null;  // 0-indexed
let _unsubscribe  = null;
let _initialised  = false;

// DOM refs — resolved once on init
let _elGrid       = null;
let _elTitle      = null;
let _elPanelDate  = null;
let _elPanelMeta  = null;
let _elPanelCount = null;
let _elNextBadge  = null;
let _elPanelList  = null;
let _elPrevBtn    = null;
let _elNextBtn    = null;

// ─── Init / destroy ───────────────────────────────────────────────────────────

/**
 * Initialise the calendar module.
 * Resolves DOM refs, subscribes to app state, renders for the first time.
 * Safe to call multiple times — re-initialises cleanly.
 */
function initCalendar() {
  // Unsubscribe any previous instance
  if (_unsubscribe) { _unsubscribe(); _unsubscribe = null; }

  _elGrid       = document.getElementById('cal-grid');
  _elTitle      = document.getElementById('cal-title');
  _elPanelDate  = document.getElementById('panel-date');
  _elPanelMeta  = document.getElementById('panel-meta');
  _elPanelCount = document.getElementById('panel-count');
  _elNextBadge  = document.getElementById('panel-next-badge');
  _elPanelList  = document.getElementById('panel-list');
  _elPrevBtn    = document.getElementById('cal-prev');
  _elNextBtn    = document.getElementById('cal-next');

  if (!_elGrid || !_elTitle || !_elPanelList) {
    console.error('calendar.js: Required DOM elements not found.');
    return;
  }

  // Wire navigation buttons
  _elPrevBtn?.addEventListener('click', goToPrevMonth);
  _elNextBtn?.addEventListener('click', goToNextMonth);

  // Start on current month
  const now      = new Date();
  _currentYear   = now.getFullYear();
  _currentMonth  = now.getMonth();
  _initialised   = true;

  // Subscribe to app state — re-render on every change
  _unsubscribe = subscribe(_onStateChange);

  // Initial render
  _renderCalendar();
  _autoSelectToday();
}

/**
 * Tear down — remove listeners, unsubscribe.
 * Call when navigating away from the calendar view.
 */
function destroyCalendar() {
  if (_unsubscribe) { _unsubscribe(); _unsubscribe = null; }
  _elPrevBtn?.removeEventListener('click', goToPrevMonth);
  _elNextBtn?.removeEventListener('click', goToNextMonth);
  _initialised = false;
}

// ─── Navigation ───────────────────────────────────────────────────────────────

function goToPrevMonth() {
  _currentMonth--;
  if (_currentMonth < 0) { _currentMonth = 11; _currentYear--; }
  selectDate(null);
  _renderCalendar();
  _resetPanel();
}

function goToNextMonth() {
  _currentMonth++;
  if (_currentMonth > 11) { _currentMonth = 0; _currentYear++; }
  selectDate(null);
  _renderCalendar();
  _resetPanel();
}

/**
 * Jump the calendar to the month containing a given ISO date.
 * Used by auto-advance when today is empty.
 */
function goToMonth(isoDate) {
  if (!isoDate) return;
  const d = new Date(isoDate + 'T00:00:00');
  _currentYear  = d.getFullYear();
  _currentMonth = d.getMonth();
  _renderCalendar();
}

// ─── State subscription ───────────────────────────────────────────────────────

function _onStateChange(state) {
  if (!_initialised) return;
  // Re-render grid to reflect filter changes (filtered array changed)
  // and update the day detail if the selected date changed
  _renderCalendar();
  if (state.selectedDate) {
    _renderDayPanel(state.selectedDate, false);
  }
}

// ─── Calendar grid ────────────────────────────────────────────────────────────

function _renderCalendar() {
  if (!_elGrid || !_elTitle) return;

  _elTitle.textContent = `${MONTH_NAMES[_currentMonth]} ${_currentYear}`;

  const daysInMonth   = new Date(_currentYear, _currentMonth + 1, 0).getDate();
  const firstDow      = new Date(_currentYear, _currentMonth, 1).getDay();
  // Convert Sunday=0 system to Monday=0
  const startOffset   = firstDow === 0 ? 6 : firstDow - 1;
  const emptyWeekISOs = _getEmptyWeekISOs(_currentYear, _currentMonth);
  const selectedDate  = getState().selectedDate;
  const todayISO      = _todayISO();

  _elGrid.innerHTML = '';

  // Leading blank cells to align first day to correct column
  for (let i = 0; i < startOffset; i++) {
    const blank = document.createElement('div');
    blank.className = 'cal-blank';
    _elGrid.appendChild(blank);
  }

  for (let day = 1; day <= daysInMonth; day++) {
    const dow = new Date(_currentYear, _currentMonth, day).getDay();
    if (dow === 0 || dow === 6) continue;  // Mon-Fri only

    const iso       = _makeISO(_currentYear, _currentMonth + 1, day);
    const dlvs      = _getDayDeliveries(iso);
    const count     = dlvs.length;
    const level     = DENSITY_LEVEL(count);
    const isToday   = iso === todayISO;
    const isSelected = iso === selectedDate;
    const isEmpty   = emptyWeekISOs.has(iso);

    const cell = document.createElement('button');
    cell.type  = 'button';
    cell.className = _cellClasses(level, isToday, isSelected, isEmpty);
    cell.dataset.iso = iso;
    cell.setAttribute('aria-label', _cellAriaLabel(day, count, iso));

    if (isEmpty) {
      cell.innerHTML = `<span class="cell-date">${day}</span>`;
      cell.disabled  = true;
    } else {
      cell.innerHTML = `
        <span class="cell-date">${day}</span>
        <span class="cell-count">${count > 0 ? count : ''}</span>
        <div class="cell-bar" aria-hidden="true">${_buildTypeBar(dlvs)}</div>
      `;
      cell.addEventListener('click', () => {
        selectDate(iso);
        _renderDayPanel(iso, false);
        _updateSelectedCell(iso);
      });
    }

    _elGrid.appendChild(cell);
  }
}

function _cellClasses(level, isToday, isSelected, isEmpty) {
  return [
    'cal-cell',
    `h${level}`,
    isToday    ? 'today'      : '',
    isSelected ? 'selected'   : '',
    isEmpty    ? 'week-empty' : '',
  ].filter(Boolean).join(' ');
}

function _cellAriaLabel(day, count, iso) {
  const date = new Date(iso + 'T00:00:00')
    .toLocaleDateString('en-GB', { weekday:'long', day:'numeric', month:'long' });
  if (count === 0) return `${date}, no deliveries`;
  return `${date}, ${count} deliver${count !== 1 ? 'ies' : 'y'}`;
}

function _updateSelectedCell(iso) {
  _elGrid?.querySelectorAll('.cal-cell').forEach(c => {
    c.classList.toggle('selected', c.dataset.iso === iso);
  });
}

// ─── Type bar ─────────────────────────────────────────────────────────────────

/**
 * Build the proportional type bar HTML.
 * Each segment gets a minimum width of BAR_MIN_PCT (8%)
 * so minority types are always visible.
 */
function _buildTypeBar(dlvs) {
  if (!dlvs || dlvs.length === 0) return '';

  // Count by type
  const counts = {};
  for (const d of dlvs) counts[d.type] = (counts[d.type] || 0) + 1;
  const total = dlvs.length;

  // Build segments with raw percentages
  let segs = Object.entries(counts).map(([type, count]) => ({
    color: TYPE_COLOURS[type] ?? '#64748b',
    pct:   (count / total) * 100,
  }));

  // Apply minimum width — redistribute from largest segments
  if (segs.length > 1) {
    const belowMin = segs.filter(s => s.pct < BAR_MIN_PCT);
    const deficit  = belowMin.reduce((sum, s) => sum + (BAR_MIN_PCT - s.pct), 0);

    if (deficit > 0) {
      belowMin.forEach(s => { s.pct = BAR_MIN_PCT; });
      const aboveMin      = segs.filter(s => s.pct > BAR_MIN_PCT);
      const surplusTotal  = aboveMin.reduce((sum, s) => sum + s.pct - BAR_MIN_PCT, 0);
      if (surplusTotal > 0) {
        aboveMin.forEach(s => {
          const share = (s.pct - BAR_MIN_PCT) / surplusTotal;
          s.pct = BAR_MIN_PCT + Math.max(0, surplusTotal - deficit) * share;
        });
      }
    }
  }

  return segs
    .map(s => `<span class="cell-bar-seg" style="width:${s.pct.toFixed(1)}%;background:${s.color}"></span>`)
    .join('');
}

// ─── Empty week detection ─────────────────────────────────────────────────────

/**
 * Return a Set of ISO dates that fall in entirely empty working weeks.
 * An empty week = all Mon-Fri cells have zero filtered deliveries.
 * These cells will be compressed to CELL_HEIGHT_EMPTY.
 */
function _getEmptyWeekISOs(year, month) {
  const daysInMonth = new Date(year, month + 1, 0).getDate();

  // Group working days by week (keyed by that week's Monday)
  const weeks = {};
  for (let day = 1; day <= daysInMonth; day++) {
    const d   = new Date(year, month, day);
    const dow = d.getDay();
    if (dow === 0 || dow === 6) continue;

    const iso     = _makeISO(year, month + 1, day);
    const monDiff = dow === 0 ? 6 : dow - 1;
    const monday  = new Date(year, month, day - monDiff);
    const wk      = `${monday.getFullYear()}-${String(monday.getMonth()+1).padStart(2,'0')}-${String(monday.getDate()).padStart(2,'0')}`;

    if (!weeks[wk]) weeks[wk] = [];
    weeks[wk].push(iso);
  }

  const emptyISOs = new Set();
  for (const isos of Object.values(weeks)) {
    const hasDeliveries = isos.some(iso => _getDayDeliveries(iso).length > 0);
    if (!hasDeliveries) isos.forEach(iso => emptyISOs.add(iso));
  }
  return emptyISOs;
}

// ─── Auto-select today ────────────────────────────────────────────────────────

/**
 * On first load — select today if it has deliveries.
 * If today is empty — silently advance to the next date with deliveries
 * and show the "↳ next delivery" badge.
 * If no future deliveries exist — show an empty prompt.
 */
function _autoSelectToday() {
  const todayISO = _todayISO();
  const todayDlvs = _getDayDeliveries(todayISO);

  if (todayDlvs.length > 0) {
    selectDate(todayISO);
    _renderDayPanel(todayISO, false);
    _updateSelectedCell(todayISO);
    return;
  }

  // Find next date with deliveries on or after today
  const nextISO = _findNextDeliveryDate(todayISO);
  if (!nextISO) {
    _resetPanel();
    return;
  }

  // Jump calendar to the month of the next delivery if necessary
  const nextDate = new Date(nextISO + 'T00:00:00');
  if (nextDate.getMonth() !== _currentMonth || nextDate.getFullYear() !== _currentYear) {
    goToMonth(nextISO);
  }

  selectDate(nextISO);
  _renderDayPanel(nextISO, true);   // true = auto-advance
  _updateSelectedCell(nextISO);
}

/**
 * Find the earliest ISO date on or after `fromISO` that has filtered deliveries.
 * Scans the byDate index keys — fast even for large datasets.
 */
function _findNextDeliveryDate(fromISO) {
  const byDate = getIndexes().byDate;
  const filtered = getFiltered();
  const filteredDocs = new Set(filtered.map(d => d.doc));

  // Get all dates in the index that have at least one delivery passing current filters
  const candidates = [...byDate.keys()]
    .filter(iso => {
      if (iso < fromISO) return false;
      return byDate.get(iso).some(d => filteredDocs.has(d.doc));
    })
    .sort();

  return candidates[0] ?? null;
}

// ─── Day detail panel ─────────────────────────────────────────────────────────

/**
 * Render the day detail panel for a given ISO date.
 * @param {string}  iso          — e.g. "2026-05-09"
 * @param {boolean} isAutoAdvance — true when today was empty and we jumped forward
 */
function _renderDayPanel(iso, isAutoAdvance) {
  if (!_elPanelDate || !_elPanelList) return;

  const dlvs    = _getDayDeliveries(iso);
  const date    = new Date(iso + 'T00:00:00');
  const todayISO = _todayISO();
  const isToday  = iso === todayISO;

  // Header date label
  _elPanelDate.textContent = isToday
    ? 'Today'
    : date.toLocaleDateString('en-GB', { weekday:'long', day:'numeric', month:'long' });

  // Count pill
  if (_elPanelCount) {
    _elPanelCount.textContent = dlvs.length
      ? `${dlvs.length} deliver${dlvs.length !== 1 ? 'ies' : 'y'}`
      : '';
  }

  // "↳ next delivery" badge — shown when auto-advanced past today
  if (_elNextBadge) {
    _elNextBadge.textContent  = '↳ next delivery';
    _elNextBadge.style.display = (isAutoAdvance && dlvs.length > 0) ? '' : 'none';
  }

  // Empty day
  if (dlvs.length === 0) {
    _elPanelList.innerHTML = `
      <div class="panel-empty">
        <div class="panel-empty-icon">📭</div>
        <div class="panel-empty-text">No deliveries this day</div>
      </div>`;
    return;
  }

  // Group by type — TYPE_ORDER ensures urgent types appear first
  const groups = {};
  for (const d of dlvs) {
    if (!groups[d.type]) groups[d.type] = [];
    groups[d.type].push(d);
  }

  const groupCount = Object.keys(groups).length;
  let html = '';

  for (const type of TYPE_ORDER) {
    if (!groups[type]) continue;

    const items   = groups[type];
    const colour  = TYPE_COLOURS[type] ?? '#64748b';
    const label   = TYPE_LABELS[type]  ?? type;
    const gid     = `grp-${type}`;
    // ZCR and ZBC always expanded; others expanded when there are ≤ 2 groups
    const isOpen  = type === 'ZCR' || type === 'ZBC' || groupCount <= 2;

    html += `
      <div class="panel-group-hdr" role="button" tabindex="0"
           onclick="togglePanelGroup('${gid}')"
           onkeydown="if(event.key==='Enter'||event.key===' ')togglePanelGroup('${gid}')">
        <span class="panel-group-dot" style="background:${colour}"></span>
        <span class="panel-group-label">${label}</span>
        <span class="panel-group-count">${items.length}</span>
        <span class="panel-group-chevron${isOpen ? ' open' : ''}" id="chev-${gid}">›</span>
      </div>
      <div class="panel-group-body" id="${gid}" ${isOpen ? '' : 'hidden'}>
    `;

    for (const d of items) {
      const mid = `mat-${d.doc}`;

      // Build materials rows
      const matRows = (d.lines || [])
        .filter(l => l.material && l.material.trim())
        .map(l => `
          <div class="dcard-mat-row">
            <span class="dcard-mat-code">${_esc(l.material)}</span>
            <span class="dcard-mat-desc">${_esc(l.description)}</span>
            <span class="dcard-mat-qty">x${l.qty % 1 === 0 ? l.qty : l.qty.toFixed(2)}</span>
          </div>`
        ).join('');

      html += `
        <div class="dcard" data-doc="${d.doc}">
          <div class="dcard-stripe" style="background:${colour}"></div>
          <div class="dcard-main" role="button" tabindex="0"
               onclick="toggleMaterials('${mid}')"
               onkeydown="if(event.key==='Enter')toggleMaterials('${mid}')">
            <div class="dcard-body">
              <div class="dcard-plot">${_esc(d.plot || d.doc)}</div>
              <div class="dcard-site">${_esc(d.name)}</div>
            </div>
            <div class="dcard-right">
              <span class="dcard-chip" style="background:${colour}">${_esc(label)}</span>
              <span class="dcard-lines">${d.lineCount} lines</span>
              <span class="dcard-mat-chevron" id="chev-${mid}">+</span>
            </div>
          </div>
          <div class="dcard-materials" id="${mid}" hidden>
            ${matRows || '<div class="dcard-mat-empty">No material lines</div>'}
          </div>
        </div>
      `;
    }

    html += '</div>';
  }

  _elPanelList.innerHTML = html;
}

/**
 * Toggle a panel group open/closed.
 * Exposed globally so inline onclick handlers can call it.
 */
function togglePanelGroup(gid) {
  const body = document.getElementById(gid);
  const chev = document.getElementById(`chev-${gid}`);
  if (!body) return;
  const wasHidden = body.hasAttribute('hidden');
  if (wasHidden) body.removeAttribute('hidden');
  else           body.setAttribute('hidden', '');
  chev?.classList.toggle('open', wasHidden);
}

/**
 * Toggle the materials list on a delivery card.
 * @param {string} mid — the materials div ID e.g. "mat-1443687"
 */
function toggleMaterials(mid) {
  const body = document.getElementById(mid);
  const chev = document.getElementById('chev-' + mid);
  if (!body) return;
  const wasHidden = body.hasAttribute('hidden');
  if (wasHidden) body.removeAttribute('hidden');
  else           body.setAttribute('hidden', '');
  if (chev) chev.textContent = wasHidden ? '−' : '+';
}

/**
 * Called when the user taps a delivery card action.
 * Phase 2: open raise-remedial sheet.
 */
function onDeliveryTap(doc) {
  console.log('Delivery tapped:', doc);
}

// ─── Reset panel ─────────────────────────────────────────────────────────────

function _resetPanel() {
  if (_elPanelDate)  _elPanelDate.textContent  = 'Select a day';
  if (_elPanelCount) _elPanelCount.textContent = '';
  if (_elNextBadge)  _elNextBadge.style.display = 'none';
  if (_elPanelList) {
    _elPanelList.innerHTML = `
      <div class="panel-prompt">
        <div class="panel-prompt-icon">👆</div>
        <div class="panel-prompt-text">Tap a date to see deliveries</div>
      </div>`;
  }
}

// ─── Data helpers ─────────────────────────────────────────────────────────────

/**
 * Get filtered deliveries for a single ISO date.
 * Uses the byDate index for O(1) date lookup, then intersects
 * with the current filtered set to respect active CM and type filters.
 */
function _getDayDeliveries(iso) {
  const byDate    = getIndexes().byDate;
  const filtered  = getFiltered();

  if (!byDate || !filtered) return [];

  const dayDlvs = byDate.get(iso);
  if (!dayDlvs || dayDlvs.length === 0) return [];

  // Filtered is already the right subset — just intersect on doc ID
  const filteredDocs = new Set(filtered.map(d => d.doc));
  return dayDlvs.filter(d => filteredDocs.has(d.doc));
}

// ─── Utilities ────────────────────────────────────────────────────────────────

function _todayISO() {
  const n = new Date();
  return _makeISO(n.getFullYear(), n.getMonth() + 1, n.getDate());
}

function _makeISO(year, month, day) {
  return `${year}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
}

/** Escape HTML to prevent XSS from SAP data appearing in the DOM */
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
    initCalendar,
    destroyCalendar,
    goToPrevMonth,
    goToNextMonth,
    goToMonth,
    togglePanelGroup,
    onDeliveryTap,
    // Exported for testing
    _buildTypeBar,
    _getEmptyWeekISOs,
    _getDayDeliveries,
    _findNextDeliveryDate,
    DENSITY_LEVEL,
  };
}
