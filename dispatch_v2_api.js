/**
 * Dispatch v2 — api.js
 *
 * Single responsibility: get raw rows and hand them to parser.js.
 *
 * Two data sources:
 *   loadFromFile(file)        — manual Excel upload (always available)
 *   loadFromSharePoint()      — N8N webhook fetch (when configured)
 *
 * One lookup source:
 *   loadCMLookup()            — fetches cm-lookup.json from the repo
 *
 * The parser never knows or cares which source was used.
 * When N8N is ready — set N8N_WEBHOOK_URL and loadFromSharePoint() works.
 * Nothing else in the app changes.
 */

'use strict';

// ─── Configuration ────────────────────────────────────────────────────────────

// Paste your N8N webhook URL here when ready.
// Leave as null to disable the SharePoint button.
const N8N_WEBHOOK_URL = null;

// Path to the CM lookup file — relative to index.html
const CM_LOOKUP_PATH = './js/cm-lookup.json';

// ─── CM Lookup ────────────────────────────────────────────────────────────────

/**
 * Fetch the CM code → full name lookup from cm-lookup.json.
 * Returns an empty object on failure so the app still works —
 * CMs will show their code instead of their name.
 *
 * @returns {Promise<Object>}  e.g. { "FOWLERM": "Mark Fowler", ... }
 */
async function loadCMLookup() {
  try {
    const res = await fetch(CM_LOOKUP_PATH);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return await res.json();
  } catch (err) {
    console.warn('api.js: Could not load cm-lookup.json —', err.message);
    return {};
  }
}

// ─── Manual file upload ───────────────────────────────────────────────────────

/**
 * Load raw rows from a user-selected Excel file.
 * Uses the SheetJS (XLSX) library — must be loaded before api.js.
 *
 * Validates:
 *   - File is an Excel format (.xlsx / .xls / .xlsm)
 *   - File size is under 50MB
 *   - At least one data row exists after the header
 *
 * @param   {File} file  — from an <input type="file"> or drag-and-drop
 * @returns {Promise<{ rows: Array[], filename: string }>}
 * @throws  {ApiError}   — with a user-readable message
 */
async function loadFromFile(file) {
  if (!file) throw new ApiError('No file provided.');

  // Type check
  const allowed = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel',
    'application/vnd.ms-excel.sheet.macroenabled.12',
  ];
  const ext = file.name.split('.').pop().toLowerCase();
  if (!allowed.includes(file.type) && !['xlsx','xls','xlsm'].includes(ext)) {
    throw new ApiError(`"${file.name}" doesn't look like an Excel file. Please upload an .xlsx file.`);
  }

  // Size check — 50MB
  if (file.size > 50 * 1024 * 1024) {
    throw new ApiError(`File is ${(file.size / 1024 / 1024).toFixed(1)}MB. Maximum is 50MB.`);
  }

  // Read
  const buffer = await readFileAsArrayBuffer(file);

  return {
    rows:     parseExcelBuffer(buffer, file.name),
    filename: file.name,
  };
}

// ─── SharePoint via N8N ───────────────────────────────────────────────────────

/**
 * Fetch the latest SAP report from SharePoint via the N8N webhook.
 *
 * N8N flow returns the Excel file as binary.
 * We parse it exactly as we would a manual upload.
 *
 * Throws if N8N_WEBHOOK_URL is not configured.
 *
 * @returns {Promise<{ rows: Array[], filename: string }>}
 * @throws  {ApiError}
 */
async function loadFromSharePoint() {
  if (!N8N_WEBHOOK_URL) {
    throw new ApiError(
      'SharePoint connection not configured. ' +
      'Set N8N_WEBHOOK_URL in api.js to enable this feature.'
    );
  }

  let response;
  try {
    response = await fetch(N8N_WEBHOOK_URL, {
      method: 'GET',
      headers: { 'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
    });
  } catch (err) {
    throw new ApiError(`Could not reach SharePoint. Check your connection. (${err.message})`);
  }

  if (!response.ok) {
    throw new ApiError(`SharePoint returned an error (HTTP ${response.status}). Try again or load manually.`);
  }

  // Get filename from Content-Disposition header if present
  const disposition = response.headers.get('Content-Disposition') ?? '';
  const nameMatch   = disposition.match(/filename[^;=\n]*=["']?([^"';\n]+)/i);
  const filename    = nameMatch?.[1]?.trim() ?? 'sharepoint-report.xlsx';

  const buffer = await response.arrayBuffer();

  return {
    rows:     parseExcelBuffer(buffer, filename),
    filename,
  };
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

/**
 * Shared helper to parse an ArrayBuffer into raw rows using SheetJS.
 * @param   {ArrayBuffer} buffer     — the file data
 * @param   {string}      sourceName — filename or source label for error messages
 * @returns {Array[]}     data rows with header removed
 * @throws  {ApiError}
 */
function parseExcelBuffer(buffer, sourceName) {
  if (typeof XLSX === 'undefined') {
    throw new ApiError('XLSX library not loaded. Check the script tags in index.html.');
  }

  let workbook;
  try {
    workbook = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: true });
  } catch (err) {
    throw new ApiError(`Could not read "${sourceName}". Is it a valid Excel file?`);
  }

  const sheetName = workbook.SheetNames[0];
  if (!sheetName) throw new ApiError(`"${sourceName}" contains no sheets.`);

  const rawRows = XLSX.utils.sheet_to_json(
    workbook.Sheets[sheetName],
    { header: 1, defval: null }
  );

  if (rawRows.length < 2) {
    throw new ApiError(`"${sourceName}" appears to be empty.`);
  }

  // Row 0 is the header — return data rows only
  return rawRows.slice(1);
}

/**
 * Read a File as an ArrayBuffer — promisified FileReader.
 * @param   {File} file
 * @returns {Promise<ArrayBuffer>}
 */
function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload  = e => resolve(e.target.result);
    reader.onerror = () => reject(new ApiError(`Could not read "${file.name}".`));
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Structured error class so the UI layer can distinguish
 * user-facing messages from unexpected exceptions.
 */
class ApiError extends Error {
  constructor(message) {
    super(message);
    this.name    = 'ApiError';
    this.isApiError = true;
  }
}

/**
 * Check whether SharePoint loading is configured.
 * The UI uses this to show or hide the SharePoint button.
 *
 * @returns {boolean}
 */
function isSharePointConfigured() {
  return Boolean(N8N_WEBHOOK_URL);
}

// ─── Exports ─────────────────────────────────────────────────────────────────

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    loadFromFile,
    loadFromSharePoint,
    loadCMLookup,
    isSharePointConfigured,
    ApiError,
    N8N_WEBHOOK_URL,
  };
}
