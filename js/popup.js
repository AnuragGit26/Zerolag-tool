// ================================================
// Popup.js
// ================================================
// Table of Contents (comments only, no functional impact)
//  1) Imports and External Modules
//  2) Error & Network Interceptors
//  3) Global State & Config
//  4) GHO Cache Keys & Defaults
//  5) Weekend Roster Cache (sessionStorage with TTL)
//  6) Google Sheets API Helper (with optional cache)
//  7) GHO User Map Cache (localStorage)
//  8) Salesforce User Email Cache (localStorage)
//  9) User Activity Tracking (mouse/keyboard)
// 10) Identity & Session (cookies, jsforce)
// 11) DOMContentLoaded Initialization & UI Wiring
// 12) GHO Modal Logic (queries, cache, rendering)
// 13) CIC Managers (weekend roster modal)
// 14) GHO Mapping Modal (view/edit mappings)
// 15) Misc Helpers & Utilities
// ================================================
// Section: 1) Imports and External Modules
// ================================================
import { timeElapsed, addMinutes, isCurrentlyWeekend } from './utils/datetime.js';
import { applyFilter, applySearch, updateWeekendModeIndicator } from './utils/dom.js';
import { trackAction, logUsageDaily } from './utils/api.js';
import { formatDateWithDayOfWeek, getShiftForDate, isToday, getCurrentShift, getPreferredShiftValues, buildPreferredShiftCondition, getWeekendSignatureTemplate, getGHOTemplate } from './modules/shift.js';
import { SHEET_SERVICE_CLOUD, SHEET_SALES_CLOUD, SHEET_INDUSTRY_CLOUD, SHEET_DATA_CLOUD_AF, getCICColumnForShift, getTEColumnForShift, getSwarmLeadServiceColumnForShift, getSwarmLeadSalesColumnForShift, getPremierSalesDevTEColumn, getPremierSalesNonDevTEColumn, getPremierSalesSwarmLeadColumn, getPremierIndustryTEColumn, getPremierIndustrySwarmLeadColumn, getPremierDataCloudColumn, getPremierAgentforceColumn } from './modules/rosterDetails.js';
import { initPremierCounters, resetPremierCountersAll, parseRosterNames, renderPremierCounters, setPremierOverride } from './modules/premierCounters.js';
import { showToast } from './modules/toast.js';
import { buildPendingCardsHtml, getPendingSectionHtml } from './modules/pending.js';
import { attachGhoPreviewTemplateCopy } from './modules/gho.js';
import { logger } from './utils/logging.js';
import {
  googleSheetsGET as _sheetsGET,
  getCellValueWithoutStrikethrough as _getCellNoStrike,
  getWeekendRosterCacheKey as _wrGetKey,
  getWeekendRosterFromCache as _wrGetFromCache,
  setWeekendRosterCache as _wrSetCache,
  cleanupWeekendRosterCache as _wrCleanup,
  getWeekendRosterData as _wrGetData,
  googleSheetsGETWithCache as _sheetsGETWithCache,
  getWeekendRosterCacheStats as _wrCacheStats,
  clearWeekendRosterCache as _wrClearCache,
  refreshWeekendRosterCache as _wrRefreshCache,
  startWeekendRosterCacheCleanup as _wrStartCleanup
} from './utils/sheets.js';

try { logger.setLevel('warn'); logger.installConsoleBeautifier(); } catch { }

// ================================================
// Section: 2) Error and Network Interceptors (Salesforce IP Access)
// ================================================
(() => {
  try {
    const SEEN_KEY = '__ip_access_toast_last_shown_ts';
    const SHOW_COOLDOWN_MS = 5 * 60 * 1000; // at most once every 5 minutes
    const origError = console.error;
    const matchErr = (args) => {
      try {
        const text = args.map(a => {
          if (a == null) return '';
          if (typeof a === 'string') return a;
          if (a && typeof a.message === 'string') return a.message;
          try { return JSON.stringify(a); } catch { return String(a); }
        }).join(' ');
        return /INSUFFICIENT_ACCESS/i.test(text) && /current\s+ip\s+address\s+is\s+not\s+allowed/i.test(text);
      } catch { return false; }
    };
    const maybeToast = () => {
      const now = Date.now();
      const last = Number(localStorage.getItem(SEEN_KEY) || 0);
      if (!last || (now - last) > SHOW_COOLDOWN_MS) {
        localStorage.setItem(SEEN_KEY, String(now));
        showToast('Salesforce blocked requests from this IP. Please turn on Zscaler/VPN or use an allowed network.', { type: 'error', duration: 12000 });
      }
    };
    const isSfUrl = (u) => {
      try { return typeof u === 'string' && /orgcs\.my\.salesforce\.com/i.test(u); } catch { return false; }
    };
    const bodyMatches = (t) => /INSUFFICIENT_ACCESS/i.test(t) && /current\s*ip\s*address\s*is\s*not\s*allowed/i.test(t);

    // Intercept fetch
    try {
      const origFetch = window.fetch.bind(window);
      window.fetch = async function (input, init) {
        const resp = await origFetch(input, init);
        try {
          const url = (resp && resp.url) || (typeof input === 'string' ? input : (input && input.url));
          if (resp && !resp.ok && isSfUrl(url)) {
            const text = await resp.clone().text().catch(() => '');
            if (text && bodyMatches(text)) maybeToast();
          }
        } catch { /* ignore */ }
        return resp;
      };
    } catch { }

    // Intercept XMLHttpRequest
    try {
      const XHR = window.XMLHttpRequest;
      if (XHR && XHR.prototype) {
        const origOpen = XHR.prototype.open;
        const origSend = XHR.prototype.send;
        XHR.prototype.open = function (method, url, async, user, password) {
          try { this.___sf_ip_err_url = url; } catch { }
          return origOpen.apply(this, arguments);
        };
        XHR.prototype.send = function (body) {
          try {
            this.addEventListener('loadend', () => {
              try {
                if (isSfUrl(this.___sf_ip_err_url) && this.status >= 400) {
                  const text = String(this.responseText || '');
                  if (text && bodyMatches(text)) maybeToast();
                }
              } catch { }
            });
          } catch { }
          return origSend.apply(this, arguments);
        };
      }
    } catch { }
    console.error = function (...args) {
      try {
        if (matchErr(args)) maybeToast();
      } catch { }
      try { return origError.apply(this, args); } catch { }
    };

    // Also intercept our structured logger path so messages printed via logger.error are caught too.
    try {
      if (logger && typeof logger.error === 'function') {
        const origLoggerError = logger.error.bind(logger);
        logger.error = (...args) => {
          try { if (matchErr(args)) maybeToast(); } catch { }
          return origLoggerError(...args);
        };
      }
    } catch { }
  } catch { }
})();

// ================================================
// Section: 3) Global State and Configuration
// ================================================
let SESSION_ID;
let currentMode = localStorage.getItem('caseTriageMode') || 'signature';
let currentUserName;
let searchTimeout;
let ghoRecordsGlobal = [];
let ghoConnectionGlobal = null;
export const SPREADSHEET_ID = '1BKxQLGFrczjhcx9rEt-jXGvlcCPQblwBhFJjoiDD7TI';
export const WEEKEND_ROSTER_SPREADSHEET_ID = '19qZi50CzKHm8PmSHiPjgTHogEue_DS70iXfT-MVfhPs';

const DEV_FORCE_SHOW_WEEKEND_ROSTER = false;

// ================================================
// Section: 4) GHO Caching Keys and Defaults
// ================================================
const GHO_CACHE_TTL = 600000;
const GHO_CACHE_STORAGE_KEY = 'gho_cache_v1';
window.__ghoCache = window.__ghoCache || loadGHOCache();
window.__ghoListState = window.__ghoListState || null;

// Initialize GHO cache
console.log('GHO Cache initialized:', window.__ghoCache);
const GHO_PAGE_SIZE = 10;
const GHO_USERMAP_CACHE_KEY = 'gho_user_map_cache_v2';
const GHO_USERMAP_CACHE_TTL = 2 * 24 * 60 * 60 * 1000;
const USER_EMAIL_CACHE_KEY = 'sf_user_email_cache_v1';
const USER_EMAIL_CACHE_TTL = 2 * 24 * 60 * 60 * 1000;

// GHO mapping storage configuration
const GHO_MAPPING_STORAGE_KEY = 'gho_transfer_mappings_v1';
const GHO_MAPPING_CACHE_TTL = 24 * 60 * 60 * 1000;

// ================================================
// Section: 5) Weekend Roster Cache (sessionStorage with TTL)
// ================================================
// Weekend roster caching system with 1-hour TTL
const WEEKEND_ROSTER_CACHE_TTL = 60 * 60 * 1000;
const WEEKEND_ROSTER_CACHE_PREFIX = 'weekend_roster_cache_';

function getWeekendRosterCacheKey(sheetName, shift, weekendDateStr, mode) { return _wrGetKey(WEEKEND_ROSTER_CACHE_PREFIX, sheetName, shift, weekendDateStr, mode); }

function getWeekendRosterFromCache(cacheKey) { return _wrGetFromCache(WEEKEND_ROSTER_CACHE_TTL, cacheKey); }

function setWeekendRosterCache(cacheKey, data) { try { _wrSetCache(cacheKey, data, WEEKEND_ROSTER_CACHE_TTL); } catch { } }

function cleanupWeekendRosterCache() { try { _wrCleanup(WEEKEND_ROSTER_CACHE_PREFIX, WEEKEND_ROSTER_CACHE_TTL); } catch { } }

function getWeekendRosterData(sheetName, shift, weekendDateStr, mode, callback, onError) {
  return _wrGetData(sheetName, shift, weekendDateStr, mode, WEEKEND_ROSTER_SPREADSHEET_ID, WEEKEND_ROSTER_CACHE_PREFIX, WEEKEND_ROSTER_CACHE_TTL, callback, onError);
}

// ================================================
// Section: 6) Google Sheets API Helper (with optional weekend roster cache)
// ================================================
// Enhanced googleSheetsGET with weekend roster caching
function googleSheetsGETWithCache(rangeA1, callback, onError, options = {}) {
  return _sheetsGETWithCache(rangeA1, callback, onError, { ...options, spreadsheetId: WEEKEND_ROSTER_SPREADSHEET_ID, prefix: WEEKEND_ROSTER_CACHE_PREFIX, ttlMs: WEEKEND_ROSTER_CACHE_TTL });
}

// ================================================
// Section: 7) GHO User Map Cache (localStorage)
// ================================================
function loadUserMapCache() {
  try {
    const raw = localStorage.getItem(GHO_USERMAP_CACHE_KEY);
    if (!raw) return { userMap: {}, fetchedAt: 0 };
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') return { userMap: {}, fetchedAt: 0 };
    const age = Date.now() - (parsed.fetchedAt || 0);
    if (age > GHO_USERMAP_CACHE_TTL) return { userMap: {}, fetchedAt: 0 };
    return { userMap: parsed.userMap || {}, fetchedAt: parsed.fetchedAt || 0 };
  } catch {
    return { userMap: {}, fetchedAt: 0 };
  }
}
function saveUserMapCache(userMap) {
  try {
    const payload = { userMap: userMap || {}, fetchedAt: Date.now() };
    localStorage.setItem(GHO_USERMAP_CACHE_KEY, JSON.stringify(payload));
  } catch { }
}

// ================================================
// Section: 8) Salesforce User Email Cache (localStorage)
// ================================================
function loadUserEmailCache() {
  try {
    const raw = localStorage.getItem(USER_EMAIL_CACHE_KEY);
    if (!raw) return { map: {}, fetchedAt: 0 };
    const parsed = JSON.parse(raw);
    const age = Date.now() - (parsed.fetchedAt || 0);
    if (age > USER_EMAIL_CACHE_TTL) return { map: {}, fetchedAt: 0 };
    return { map: parsed.map || {}, fetchedAt: parsed.fetchedAt || 0 };
  } catch { return { map: {}, fetchedAt: 0 }; }
}
function saveUserEmailCache(map) {
  try { localStorage.setItem(USER_EMAIL_CACHE_KEY, JSON.stringify({ map: map || {}, fetchedAt: Date.now() })); } catch { }
}

// ================================================
// Section: GHO Data Cache (per-mode signature/premier)
// ================================================
// GHO Cache Functions
function loadGHOCache() {
  try {
    const raw = localStorage.getItem(GHO_CACHE_STORAGE_KEY);
    if (!raw) return { signature: null, premier: null };

    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') return { signature: null, premier: null };

    const now = Date.now();
    const validCache = {};

    // Check each mode's cache validity
    ['signature', 'premier'].forEach(mode => {
      if (parsed[mode] && parsed[mode].fetchedAt) {
        const age = now - parsed[mode].fetchedAt;
        if (age < GHO_CACHE_TTL) {
          // Cache is still valid, restore it
          validCache[mode] = {
            ...parsed[mode],
            triageCaseIds: new Set(parsed[mode].triageCaseIds || []),
            mondayCaseIds: new Set(parsed[mode].mondayCaseIds || []),
            noWocTemplateCaseIds: new Set(parsed[mode].noWocTemplateCaseIds || [])
          };
          console.log(`GHO cache restored for ${mode} mode (${Math.floor(age / 1000 / 60)}m old)`);
          console.log(`  - triageCaseIds: ${parsed[mode].triageCaseIds?.length || 0} cases`);
          console.log(`  - mondayCaseIds: ${parsed[mode].mondayCaseIds?.length || 0} cases`);
          console.log(`  - mondayCaseIds content:`, parsed[mode].mondayCaseIds || []);
          console.log(`  - noWocTemplateCaseIds: ${parsed[mode].noWocTemplateCaseIds?.length || 0} cases`);
          console.log(`  - noWocTemplateCaseIds content:`, parsed[mode].noWocTemplateCaseIds || []);
        } else {
          console.log(`GHO cache expired for ${mode} mode (${Math.floor(age / 1000 / 60)}m old)`);
        }
      }
    });

    return validCache;
  } catch (error) {
    console.warn('Failed to load GHO cache:', error);
    return { signature: null, premier: null };
  }
}

function saveGHOCache() {
  try {
    if (!window.__ghoCache) return;

    const payload = {};
    ['signature', 'premier'].forEach(mode => {
      if (window.__ghoCache[mode]) {
        payload[mode] = {
          records: window.__ghoCache[mode].records,
          triageCaseIds: Array.from(window.__ghoCache[mode].triageCaseIds || []),
          mondayCaseIds: Array.from(window.__ghoCache[mode].mondayCaseIds || []),
          noWocTemplateCaseIds: Array.from(window.__ghoCache[mode].noWocTemplateCaseIds || []),
          userMap: window.__ghoCache[mode].userMap || {},
          fetchedAt: window.__ghoCache[mode].fetchedAt,
          shift: window.__ghoCache[mode].shift
        };
      }
    });

    localStorage.setItem(GHO_CACHE_STORAGE_KEY, JSON.stringify(payload));
    console.log('GHO cache saved to persistent storage');
    console.log('Cache payload:', payload);
    ['signature', 'premier'].forEach(mode => {
      if (payload[mode]) {
        console.log(`${mode} mode cache:`, {
          records: payload[mode].records?.length || 0,
          triageCaseIds: payload[mode].triageCaseIds?.length || 0,
          mondayCaseIds: payload[mode].mondayCaseIds?.length || 0,
          mondayCaseIdsContent: payload[mode].mondayCaseIds || [],
          noWocTemplateCaseIds: payload[mode].noWocTemplateCaseIds?.length || 0,
          noWocTemplateCaseIdsContent: payload[mode].noWocTemplateCaseIds || []
        });
      }
    });
  } catch (error) {
    console.warn('Failed to save GHO cache:', error);
  }
}
function escapeSoqlString(s) { return String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'"); }
function cleanRosterNameForQuery(name) {
  let s = String(name || '').trim();
  s = s.replace(/^\s*[^:]*:\s*/, '');
  // Remove any content in common brackets: (), [], {}, <>
  s = s.replace(/\([^)]*\)|\[[^\]]*\]|\{[^}]*\}|<[^>]*>/g, '');
  // Remove any single-quoted suffixes like:  Name 'Sev 1(Sales)'
  s = s.replace(/'[^']*'/g, '');
  // Also remove any double-quoted segments just in case
  s = s.replace(/"[^"]*"/g, '');
  // Omit tokens like 'Sev 1' and '_Dev' (case-insensitive)
  s = s.replace(/\bsev\s*1\b/gi, '');
  s = s.replace(/_dev\b/gi, '');
  s = s.replace(/\s{2,}/g, ' ').trim();
  return s;
}
function getEmailsByNames(names) {
  return new Promise((resolve) => {
    try {
      const unique = Array.from(new Set((names || []).map(n => String(n).trim()).filter(Boolean)));
      if (unique.length === 0) return resolve({});
      const pairs = unique.map(n => ({ original: n, cleaned: cleanRosterNameForQuery(n) })).filter(p => p.cleaned);
      const cached = loadUserEmailCache();
      const cache = cached.map || {};
      const cleanedSet = Array.from(new Set(pairs.map(p => p.cleaned.toLowerCase())));
      const missing = cleanedSet.filter(k => !cache[k]);
      if (missing.length === 0) {
        const out = {};
        pairs.forEach(p => { const k = p.cleaned.toLowerCase(); const em = cache[k]; if (em) { out[k] = em; out[p.original.toLowerCase()] = em; } });
        return resolve(out);
      }
      const conn = new jsforce.Connection({ serverUrl: 'https://orgcs.my.salesforce.com', sessionId: SESSION_ID, version: '64.0' });
      const CHUNK = 70;
      let idx = 0; const merged = { ...cache };
      const runNext = () => {
        if (idx >= missing.length) {
          saveUserEmailCache(merged);
          const out = {};
          pairs.forEach(p => { const k = p.cleaned.toLowerCase(); const em = merged[k]; if (em) { out[k] = em; out[p.original.toLowerCase()] = em; } });
          return resolve(out);
        }
        const slice = missing.slice(idx, idx += CHUNK);
        const soql = `SELECT Name, Email, Username FROM User WHERE IsActive = true AND Email != null AND (Username LIKE '%dreamevent.com' OR Username LIKE '%orgcs.com') AND Name IN ('${slice.map(escapeSoqlString).join("','")}')`;
        conn.query(soql, (err, res) => {
          if (!err && res && res.records) {
            res.records.forEach(u => { if (u && u.Name && u.Email) merged[String(u.Name).toLowerCase()] = u.Email; });
          } else {
            console.warn('Email query failed for names chunk', err);
          }
          runNext();
        });
      };
      runNext();
    } catch (e) { console.warn('getEmailsByNames failed', e); resolve({}); }
  });
}

let mouseActivityTimer;
const MOUSE_ACTIVITY_TIMEOUT = 60000;

const USAGE_STATE_KEY = 'usage_state_v1';
let usageState = (function loadUsageState() {
  try {
    const raw = localStorage.getItem(USAGE_STATE_KEY);
    if (!raw) return {};
    return JSON.parse(raw) || {};
  } catch { return {}; }
})();

const IST_OFFSET_MINUTES = 330;

function getISTNow() {
  const nowUtc = Date.now();
  return new Date(nowUtc + IST_OFFSET_MINUTES * 60 * 1000);
}

function formatDateYMD(d) {
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth() + 1).padStart(2, '0');
  const day = String(d.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function addDaysIST(dateLike, delta) {
  return new Date(dateLike.getTime() + delta * 24 * 60 * 60 * 1000);
}

function getISTBusinessDayKey() {
  const istNow = getISTNow();
  const hours = istNow.getUTCHours(); // since we artificially shifted the date into IST, UTC hours on shifted date == IST hours
  const minutes = istNow.getUTCMinutes();
  // If before 05:30 IST, assign to previous calendar date (IST)
  if (hours < 5 || (hours === 5 && minutes < 30)) {
    const prev = addDaysIST(istNow, -1);
    return formatDateYMD(prev);
  }
  return formatDateYMD(istNow);
}

function sheetDisplayDateFromBusinessKey(key) {
  const [y, m, d] = key.split('-');
  return `${m}/${d}/${y}`;
}

let activeSessionStart = null;
let lastActivityTs = Date.now();
let usageTickInterval = null;

let currentBusinessDayKey = getISTBusinessDayKey();

function cleanupOldUsageData(oldKey, newKey) {
  try {
    if (oldKey && usageState[oldKey] && oldKey !== newKey) {
      delete usageState[oldKey];
    }
    const formattedOld = sheetDisplayDateFromBusinessKey(oldKey);
    const prefixes = [`usage_logged_${formattedOld}_`, `usage_logged_${oldKey}_`];
    const toRemove = [];
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (k && prefixes.some(p => k.startsWith(p))) {
        toRemove.push(k);
      }
    }
    toRemove.forEach(k => localStorage.removeItem(k));
    persistUsageState();
    if (toRemove.length) logger.debug('usage.purge', { oldKey, removed: toRemove.length });
  } catch (e) { console.warn('cleanupOldUsageData failed', e); }
}

function ensureUserUsageBucket(userName) {
  const newKey = getISTBusinessDayKey();
  if (newKey !== currentBusinessDayKey) {
    // Day rollover at 05:30 IST - start fresh bucket (prior day retained in usageState for reference until reload)
    const oldKey = currentBusinessDayKey;
    currentBusinessDayKey = newKey;
    cleanupOldUsageData(oldKey, newKey);
  }
  usageState[currentBusinessDayKey] = usageState[currentBusinessDayKey] || {};
  usageState[currentBusinessDayKey][userName] = usageState[currentBusinessDayKey][userName] || { activeSeconds: 0, lastActiveISO: null, actionedCases: 0 };
  return { dayKey: currentBusinessDayKey, bucket: usageState[currentBusinessDayKey][userName] };
}

function persistUsageState() {
  try { localStorage.setItem(USAGE_STATE_KEY, JSON.stringify(usageState)); } catch { }
}

function startUsageTicking() {
  if (usageTickInterval) return;
  usageTickInterval = setInterval(() => {
    if (!currentUserName) return;
    const now = Date.now();
    if (now - lastActivityTs <= MOUSE_ACTIVITY_TIMEOUT) {
      if (!activeSessionStart) activeSessionStart = new Date();
      const { dayKey, bucket } = ensureUserUsageBucket(currentUserName);
      bucket.activeSeconds += 60;
      bucket.lastActiveISO = new Date().toISOString();
      persistUsageState();
      if (bucket.activeSeconds % (5 * 60) === 0 || bucket.activeSeconds <= 600 && bucket.activeSeconds % 300 === 0) {
        logUsageDaily({
          dateLocalString: sheetDisplayDateFromBusinessKey(dayKey),
          userName: currentUserName,
          totalActiveMinutes: Math.round(bucket.activeSeconds / 60),
          lastActiveISO: bucket.lastActiveISO,
          actionedCases: bucket.actionedCases || 0
        });
      }
    } else {
      activeSessionStart = null;
    }
    // Force evaluation of business day key to trigger rollover if needed
    getISTBusinessDayKey();
  }, 60000);
}

function recordImmediateUsageFlush() {
  try {
    if (!currentUserName) return;
    const { dayKey, bucket } = ensureUserUsageBucket(currentUserName);
    if (bucket.activeSeconds > 0) {
      logUsageDaily({
        dateLocalString: sheetDisplayDateFromBusinessKey(dayKey),
        userName: currentUserName,
        totalActiveMinutes: Math.round(bucket.activeSeconds / 60),
        lastActiveISO: bucket.lastActiveISO || new Date().toISOString(),
        actionedCases: bucket.actionedCases || 0
      });
    }
  } catch (e) { console.warn('Immediate usage flush error', e); }
}

['focus', 'mousemove', 'keydown', 'click', 'scroll'].forEach(evt => {
  window.addEventListener(evt, () => { lastActivityTs = Date.now(); }, true);
});

window.addEventListener('beforeunload', () => {
  recordImmediateUsageFlush();
});
startUsageTicking();

function resetMouseActivityTimer() {
  if (mouseActivityTimer) {
    clearTimeout(mouseActivityTimer);
  }
  mouseActivityTimer = setTimeout(() => {
    logger.info('ui.inactive.refresh', { reason: 'mouse inactivity', timeoutMs: MOUSE_ACTIVITY_TIMEOUT });
    window.location.reload();
  }, MOUSE_ACTIVITY_TIMEOUT);
}

// ================================================
// Section: 9) User Activity Tracking (mouse/keyboard)
// ================================================
function initMouseActivityTracking() {
  const events = ['mousemove', 'mousedown', 'mouseup', 'click', 'scroll', 'keydown', 'keyup'];

  events.forEach(eventType => {
    document.addEventListener(eventType, resetMouseActivityTimer, true);
  });
  resetMouseActivityTimer();
  logger.debug('ui.mouseTracking.init', { timeoutMs: MOUSE_ACTIVITY_TIMEOUT });
}

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  try {
    if (request.action === 'dailyCleanupCompleted') {
      // Handle daily cleanup notification
      // Daily cleanup completed
      showToast(`Daily cleanup completed: ${request.cleanedCases} cases removed from tracking`, 'info');
      sendResponse({ success: true });
      return true;
    } else if (request.action === 'trackActionFromBackground') {
      // Handle case tracking from background script
      console.log('Received trackActionFromBackground message:', request.data);
      const { createdDate, caseNumber, severity, actionType, cloud, mode, userName, newValue } = request.data;

      console.log('About to call trackActionAndCount with params:', {
        createdDate, caseNumber, severity, actionType, cloud, mode, userName, newValue
      });

      // Call trackActionAndCount to log to Google Sheets
      try {
        trackActionAndCount(createdDate, caseNumber, severity, actionType, cloud, mode, userName, newValue);
        console.log('trackActionAndCount called successfully');
      } catch (error) {
        console.error('Error calling trackActionAndCount:', error);
      }

      sendResponse({ success: true });
      return true;
    } else if (request.action === 'setMode') {
      // Handle mode setting request from background script
      console.log('Received setMode message:', request.mode);
      try {
        if (request.mode === 'signature' || request.mode === 'premier') {
          localStorage.setItem('caseTriageMode', request.mode);
          console.log(`Mode set to: ${request.mode} in popup localStorage`);

          // Update the UI if the mode selector exists
          const modeSelector = document.getElementById('caseTriageMode');
          if (modeSelector) {
            modeSelector.value = request.mode;
            console.log(`UI updated to show mode: ${request.mode}`);
          }

          sendResponse({ success: true, message: `Mode set to ${request.mode}` });
        } else {
          sendResponse({ success: false, message: 'Invalid mode' });
        }
      } catch (error) {
        console.error('Error setting mode:', error);
        sendResponse({ success: false, message: error.message });
      }
      return true;
    } else if (request.action === 'getMode') {
      // Handle mode retrieval request from background script
      console.log('Received getMode message');
      try {
        const currentMode = localStorage.getItem('caseTriageMode') || 'signature';
        console.log(`Returning mode: ${currentMode}`);
        sendResponse({ success: true, mode: currentMode });
      } catch (error) {
        console.error('Error getting mode:', error);
        sendResponse({ success: false, message: error.message });
      }
      return true;
    } else if (request.action === 'cleanLocalStorage') {
      // Handle localStorage cleaning request
      console.log('Cleaning localStorage as requested...');
      try {
        const rawData = localStorage.getItem('sentToSheets');

        if (!rawData) {
          console.log('No localStorage data to clean');
          sendResponse({ success: true, message: 'No data to clean' });
          return true;
        }

        try {
          const sentEntries = JSON.parse(rawData);
          if (!Array.isArray(sentEntries)) {
            console.log('Invalid localStorage format, clearing completely');
            localStorage.removeItem('sentToSheets');
            sendResponse({ success: true, message: 'Invalid data cleared' });
            return true;
          }

          // Filter out entries older than current business day to prevent localStorage from growing indefinitely
          const now = new Date();
          const istTime = new Date(now.toLocaleString("en-US", { timeZone: "Asia/Kolkata" }));

          // Calculate start of current business day (5:30 AM of previous day to 5:30 AM of current day)
          // If current time is before 5:30 AM, we're still in yesterday's business day
          // If current time is after 5:30 AM, we're in today's business day
          const currentBusinessDayStart = new Date(istTime);
          currentBusinessDayStart.setHours(5, 30, 0, 0); // 5:30 AM of current day

          // If it's before 5:30 AM today, use yesterday's 5:30 AM as business day start
          if (istTime < currentBusinessDayStart) {
            currentBusinessDayStart.setDate(currentBusinessDayStart.getDate() - 1);
          }

          const filteredEntries = sentEntries.filter(entry => {
            try {
              // Extract date from tracking ID format: CASE_NUMBER_DATE_ACTIONTYPE_ASSIGNEDTO
              const parts = entry.split('_');
              if (parts.length >= 3) {
                const dateStr = parts[1]; // Format: YYYY-MM-DD
                const entryDate = new Date(dateStr);
                // Keep entries from current business day onwards
                return entryDate >= currentBusinessDayStart;
              }
              return false; // Remove malformed entries
            } catch (error) {
              console.log(`Error parsing date from entry ${entry}:`, error);
              return false; // Remove entries with invalid dates
            }
          });

          if (filteredEntries.length !== sentEntries.length) {
            const removedCount = sentEntries.length - filteredEntries.length;
            console.log(`Removed ${removedCount} expired entries from localStorage`);
            localStorage.setItem('sentToSheets', JSON.stringify(filteredEntries));
            sendResponse({ success: true, message: `${removedCount} expired entries removed` });
          } else {
            console.log('No expired entries to remove');
            sendResponse({ success: true, message: 'No expired entries to remove' });
          }

        } catch (parseError) {
          console.log('Error parsing localStorage data, clearing completely:', parseError);
          localStorage.removeItem('sentToSheets');
          sendResponse({ success: true, message: 'Corrupted data cleared' });
        }

      } catch (error) {
        console.error('Error cleaning localStorage:', error);
        sendResponse({ success: false, error: error.message });
      }
      return true;
    } else if (request.action === 'clearLocalStorage') {
      // Handle localStorage clearing request (manual override)
      console.log('Manually clearing localStorage as requested...');
      try {
        localStorage.removeItem('sentToSheets');
        console.log('localStorage cleared successfully');
        sendResponse({ success: true, message: 'localStorage manually cleared' });
      } catch (error) {
        console.error('Error clearing localStorage:', error);
        sendResponse({ success: false, error: error.message });
      }
      return true;
    }
  } catch (error) {
    console.error('Error handling background message:', error);
    sendResponse({ success: false, error: error.message });
  }
  return false;
});

// Generate a unique tracking identifier to prevent Google Sheet duplicates
function generateTrackingId(caseNumber, dateofAction, actionType, assignedTo) {
  try {
    const dateStr = dateofAction instanceof Date ? dateofAction.toISOString().split('T')[0] : dateofAction.split('T')[0];
    const uniqueId = `${caseNumber}_${dateStr}_${actionType}_${assignedTo}`;
    return uniqueId.replace(/[^a-zA-Z0-9_-]/g, '_');
  } catch (error) {
    console.error('Error generating tracking ID:', error);
    return `tracking_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }
}

// Check if this exact tracking entry has already been sent to Google Sheets
function checkIfAlreadySentToSheets(trackingId) {
  try {
    console.log(`Checking if tracking ID ${trackingId} was already sent to Google Sheets...`);
    const rawData = localStorage.getItem('sentToSheets');

    // If no data exists, nothing has been sent
    if (!rawData) {
      console.log('No localStorage data found, returning false');
      return false;
    }

    console.log(`Raw localStorage data: ${rawData.substring(0, 200)}...`);

    // Check if rawData exists and is a string before calling includes
    if (typeof rawData === 'string') {
      // Check if data is corrupted (contains Promise objects)
      if (rawData.includes('[object Promise]')) {
        console.warn('Corrupted localStorage data detected (contains Promise objects). Clearing and resetting.');
        localStorage.removeItem('sentToSheets');
        return false;
      }
    } else {
      // Handle non-string data types
      console.warn(`Invalid localStorage data type detected: ${typeof rawData}. Expected string, got: ${rawData}. Clearing and resetting.`);
      localStorage.removeItem('sentToSheets');
      return false;
    }

    const sentEntries = JSON.parse(rawData);
    console.log(`Parsed sentEntries: ${JSON.stringify(sentEntries).substring(0, 200)}...`);

    // Additional validation to ensure entries are strings
    if (!Array.isArray(sentEntries) || sentEntries.some(entry => typeof entry !== 'string')) {
      console.warn('Invalid localStorage data format detected. Clearing and resetting.');
      localStorage.removeItem('sentToSheets');
      return false;
    }

    const isAlreadySent = sentEntries.includes(trackingId);
    console.log(`Tracking ID ${trackingId} already sent: ${isAlreadySent}`);
    return isAlreadySent;
  } catch (error) {
    console.error('Error checking sent entries:', error);
    // Clear corrupted data
    try {
      localStorage.removeItem('sentToSheets');
    } catch (clearError) {
      console.error('Failed to clear corrupted localStorage:', clearError);
    }
    return false;
  }
}

// Mark tracking entry as sent to Google Sheets
function markAsSentToSheets(trackingId) {
  try {
    console.log(`Marking tracking ID ${trackingId} as sent to Google Sheets...`);

    // Validate trackingId is a string
    if (typeof trackingId !== 'string') {
      console.error('Invalid trackingId type:', typeof trackingId, trackingId);
      return;
    }

    const rawData = localStorage.getItem('sentToSheets');
    let sentEntries = [];

    try {
      if (rawData && typeof rawData === 'string') {
        // Check if data is corrupted
        if (rawData.includes('[object Promise]')) {
          console.warn('Corrupted localStorage data detected. Clearing and resetting.');
          localStorage.removeItem('sentToSheets');
        } else {
          sentEntries = JSON.parse(rawData);
        }
      }
    } catch (parseError) {
      console.warn('Failed to parse existing localStorage data. Clearing and resetting.');
      localStorage.removeItem('sentToSheets');
    }

    // Ensure sentEntries is an array
    if (!Array.isArray(sentEntries)) {
      sentEntries = [];
    }

    // Filter out any non-string entries
    sentEntries = sentEntries.filter(entry => typeof entry === 'string');

    console.log(`Before adding: ${sentEntries.length} entries`);
    sentEntries.push(trackingId);
    console.log(`After adding: ${sentEntries.length} entries`);

    if (sentEntries.length > 1000) {
      sentEntries.splice(0, sentEntries.length - 1000);
    }

    localStorage.setItem('sentToSheets', JSON.stringify(sentEntries));
    console.log(`Marked tracking ID ${trackingId} as sent to Google Sheets`);
  } catch (error) {
    console.error('Error marking as sent to sheets:', error);
  }
}

// Clean up old tracking IDs (older than 30 days) to prevent localStorage bloat
function cleanupOldTrackingIds() {
  try {
    const sentEntries = JSON.parse(localStorage.getItem('sentToSheets') || '[]');
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

    const recentEntries = sentEntries.filter(entry => {
      const dateMatch = entry.match(/\d{4}-\d{2}-\d{2}/);
      if (dateMatch) {
        const entryDate = new Date(dateMatch[0]);
        return entryDate >= thirtyDaysAgo;
      }
      return true;
    });

    if (recentEntries.length < sentEntries.length) {
      localStorage.setItem('sentToSheets', JSON.stringify(recentEntries));
      console.log(`Cleaned up ${sentEntries.length - recentEntries.length} old tracking IDs`);
    }
  } catch (error) {
    console.error('Error cleaning up old tracking IDs:', error);
  }
}

setInterval(cleanupOldTrackingIds, 24 * 60 * 60 * 1000);





// Wrapper: increment per-day actioned cases and then forward to API logger
function trackActionAndCount(dateofAction, caseNumber, severity, actionType, cloud, mode, userName, assignedTo = '') {
  try {
    console.log('trackActionAndCount called with params:', {
      dateofAction, caseNumber, severity, actionType, cloud, mode, userName, assignedTo
    });

    const name = userName || currentUserName;
    console.log('Using name:', name);

    const trackingId = generateTrackingId(caseNumber, dateofAction, actionType, assignedTo);
    console.log(`Generated tracking ID: ${trackingId}`);

    if (checkIfAlreadySentToSheets(trackingId)) {
      console.log(`Case ${caseNumber} with tracking ID ${trackingId} has already been sent to Google Sheets. Skipping to prevent duplicate.`);
      return;
    }

    console.log(`About to call trackAction for case ${caseNumber}. Current actionedCases: ${name ? ensureUserUsageBucket(name).bucket.actionedCases || 0 : 'unknown'}`);
    console.log(`Tracking ID for this case: ${trackingId}`);

    console.log('Calling trackAction function...');
    console.log('trackAction params:', { dateofAction, caseNumber, severity, actionType, cloud, mode: mode || currentMode, name, assignedTo });

    trackAction(dateofAction, caseNumber, severity, actionType, cloud, mode || currentMode, name, assignedTo, () => {
      console.log('trackAction success callback executed');
      if (name) {
        const { bucket } = ensureUserUsageBucket(name);
        const oldCount = bucket.actionedCases || 0;
        bucket.actionedCases = oldCount + 1;
        bucket.lastActiveISO = new Date().toISOString();
        persistUsageState();
        console.log(`SUCCESS: Incremented actionedCases for ${name} from ${oldCount} to ${bucket.actionedCases} for case ${caseNumber} (popup processing)`);

        console.log('Call stack for actionedCases increment:', new Error().stack);
      }

      markAsSentToSheets(trackingId);
      console.log(`Marked tracking ID ${trackingId} as sent to prevent duplicates`);
    });
    console.log('trackAction function called successfully');

  } catch (e) {
    console.error('trackActionAndCount failed', e);
    console.error('Error stack:', e.stack);
    try {
      console.log('Attempting fallback trackAction call...');
      trackAction(dateofAction, caseNumber, severity, actionType, cloud, mode || currentMode, userName || currentUserName, assignedTo);
      console.log('Fallback trackAction call completed');
    } catch (fallbackError) {
      console.error('Fallback trackAction also failed:', fallbackError);
    }
  }
}












// ================================================
// Section: 10) Identity & Session (cookies, jsforce)
// ================================================
function getSessionIds() {
  getCookies("https://orgcs.my.salesforce.com", "sid", function (cookie) {
    SESSION_ID = cookie.value;
    getCaseDetails();
  });
}
function getCookies(domain, name, callback) {
  chrome.cookies.get({
    url: domain,
    name: name
  }, function (cookie) {
    if (cookie) {
      if (callback) {
        callback(cookie);
      }
    }
    else {
      showToast(`No session cookie found for ${domain}. Please log in to Salesforce.`);
    }
  });
}

function getCaseDetails() {
  logger.debug('sf.session.set', { hasSession: !!SESSION_ID });
  let conn = new jsforce.Connection({
    serverUrl: 'https://orgcs.my.salesforce.com',
    sessionId: SESSION_ID,
    version: '64.0',
  });
  conn.identity(function (err, res) {
    if (err) {
      const errorMessage = (err.message || err.toString()).toLowerCase();
      if (errorMessage.includes('session') || errorMessage.includes('connection')) {
        showToast('Connection or Session ID error. Please ensure Zscaler is turned on!');
      }

      var data = 'openTabSilent';
      chrome.runtime.sendMessage(data, function (response) {
        if (chrome.runtime.lastError) {
          console.error('Runtime error:', chrome.runtime.lastError.message);
          return;
        }
      });
      return console.error('error---' + err);
    } else {
      currentUserName = res.display_name;
      // Update 3-dots menu header now that we have the user
      try { updateMoreMenuUserHeader(); } catch { }
      try {
        ensureUserUsageBucket(currentUserName);
        lastActivityTs = Date.now();
        recordImmediateUsageFlush();
      } catch (e) { console.warn('Usage init after identity failed', e); }
      let displayedCaseCount = 0;
      let actionTakenCount = 0;
      let currentUserId;

      conn.query(`SELECT Id FROM User WHERE Name = '${currentUserName}' AND IsActive = True AND Username LIKE '%orgcs.com'`, function (err, userResult) {
        if (err) {
          return console.error('Error fetching user ID:', err);
        }
        if (userResult.records.length > 0) {
          currentUserId = userResult.records[0].Id;

          // Resume continuous processing when user becomes active
          resumeContinuousProcessing(currentUserId, currentUserName, currentMode);
        }
        let signatureQuery = "SELECT Id, CreatedDate, Account.Name, Owner.Name, SE_Target_Response__c, Severity_Level__c, CaseNumber, Subject, CaseRoutingTaxonomy__r.Name,SE_Initial_Response_Status__c, Contact.Is_MVP__c, support_available_timezone__c, (SELECT Transfer_Reason__c,Back_To_Queue__c, CreatedDate, Severity_New_Value__c, Severity_Old_Value__c FROM Case_Routing_Logs__r ORDER BY CreatedDate DESC LIMIT 20) FROM Case WHERE (Owner.Name LIKE '%Skills Queue%' OR Owner.Name='Kase Changer' OR Owner.Name='Working in Org62' OR Owner.Name='Data Cloud Queue') AND IsClosed=false AND Account_Support_SBR_Category__c!='JP MCS' AND Account.Name!='BT Test Account - HPA Premier Plus' AND Status='New' AND (((CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Community%' OR CaseRoutingTaxonomy__r.Name LIKE 'Scale Center%' OR CaseRoutingTaxonomy__r.Name LIKE 'Customer Success Score%') AND (Severity_Level__c='Level 1 - Critical' OR Severity_Level__c='Level 2 - Urgent') AND (Case_Support_level__c='Premier Priority' OR Case_Support_level__c='Signature' OR Case_Support_level__c='Signature Success')) OR ((CaseRoutingTaxonomy__r.Name LIKE 'Industry%') AND (Severity_Level__c='Level 1 - Critical' OR Severity_Level__c='Level 2 - Urgent') AND (Case_Support_level__c='Premier Priority' OR Case_Support_level__c='Signature' OR Case_Support_level__c='Signature Success')) OR (Contact.Is_MVP__c=true AND ((Severity_Level__c IN ('Level 1 - Critical', 'Level 2 - Urgent', 'Level 3 - High', 'Level 4 - Medium') AND (CaseRoutingTaxonomy__r.Name LIKE 'Sales%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry%')) OR (Severity_Level__c IN ('Level 3 - High', 'Level 4 - Medium') AND (CaseRoutingTaxonomy__r.Name LIKE 'Data Cloud-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Sales-Agentforce%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-Agentforce%')))) OR (CaseRoutingTaxonomy__r.Name='Sales-Issues Developing for Salesforce Functions (Product)' AND CreatedDate = LAST_N_DAYS:2)) AND (CaseRoutingTaxonomy__r.Name NOT IN ('Sales-Disability and Product Accessibility', 'Service-Disability and Product Accessibility', 'Industry-Disability and Product Accessibility', 'Sales-Quip', 'Sales-Sales Cloud for Slack', 'Industry-Nonprofit Cloud', 'Industry-Education Cloud', 'Industry-Education Data Architecture (EDA)', 'Industry-Education Packages (Other SFDO)', 'Industry-Nonprofit Packages (Other SFDO)', 'Industry-Nonprofit Success Pack (NPSP)', 'Service-Agentforce', 'Service-Agent for setup', 'Service-AgentforEmail', 'Service-Field Service Agentforce', 'Service-Agentforce for Dev', 'Sales-Agentforce', 'Sales-Agentforce for Dev', 'Sales-Agent for Setup', 'Sales-Prompt Builder', 'Data Cloud-Admin', 'Permissions', 'Flows', 'Reports & Dashboards', 'Data Cloud-Model Builder', 'Data Cloud-Connectors & Data Streams', 'Data Cloud-Developer', 'Calculated Insights & Consumption', 'Data Cloud-Segments', 'Activations & Identity Resolution')) ORDER BY CreatedDate DESC";
        let premierQuery = "SELECT Id, CreatedDate, Account.Name, Owner.Name, SE_Target_Response__c, Severity_Level__c, CaseNumber, Subject, CaseRoutingTaxonomy__r.Name, SE_Initial_Response_Status__c, Initial_Case_Severity__c, Contact.Is_MVP__c, (SELECT Transfer_Reason__c, Back_To_Queue__c, CreatedDate, Severity_New_Value__c, Severity_Old_Value__c FROM Case_Routing_Logs__r ORDER BY CreatedDate DESC LIMIT 20) FROM Case WHERE (Owner.Name IN ('Kase Changer', 'Working in Org62', 'Service Cloud Skills Queue', 'Sales Cloud Skills Queue', 'Industry Skills Queue', 'EXP Skills Queue', 'Data Cloud Queue')) AND (RecordType.Name IN ('Support', 'Partner Program Support', 'Platform / Application Support')) AND (Reason != 'Sales Request') AND (CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry-%') AND (Account_Support_SBR_Category__c != 'JP') AND ((Case_Support_level__c IN ('Partner Premier', 'Premier', 'Premier+', 'Premium') AND (SE_Initial_Response_Status__c NOT IN ('Met', 'Completed After Violation', 'Missed', 'Violated')) AND ((Severity_Level__c IN ('Level 1 - Critical', 'Level 2 - Urgent')) OR (Initial_Case_Severity__c IN ('Level 2 - Urgent', 'Level 1 - Critical')))) OR (Case_Support_level__c = 'Standard' AND Initial_Case_Severity__c = 'Level 1 - Critical')) AND (CaseRoutingTaxonomy__r.Name NOT IN ('Service-Agentforce', 'Service-Agent for setup', 'Service-AgentforEmail', 'Service-Field Service Agentforce', 'Service-Agentforce for Dev', 'Sales-Agentforce', 'Sales-Agentforce for Dev', 'Sales-Agent for Setup', 'Sales-Prompt Builder', 'Data Cloud-Admin', 'Permissions', 'Flows', 'Reports & Dashboards', 'Data Cloud-Model Builder', 'Data Cloud-Connectors & Data Streams', 'Data Cloud-Developer', 'Calculated Insights & Consumption', 'Data Cloud-Segments', 'Activations & Identity Resolution')) AND (IsClosed = false) AND (CreatedDate = TODAY) ORDER BY CreatedDate DESC";
        let query = currentMode === 'premier' ? premierQuery : signatureQuery;

        return conn.query(query,
          function (err, result) {
            if (err) {
              try { showToast('Query failed. Please retry or refresh.', 'error'); } catch { }
              return console.error(err);
            }

            if (result.records.length === 0) {
              const noCasesHtml = `
              <div class="no-cases-message">
                <h4 class="no-cases-title">No Cases to Action</h4>
                <p class="no-cases-text">All cases are up to date. Great work!</p>
                <p class="mode-switch-hint">💡 Meanwhile you can switch between Signature and Premier modes to see different case types.</p>
              </div>
            `;
              const container = document.getElementById("parentSigSev2");
              container.classList.remove('is-loading');
              container.innerHTML = noCasesHtml;
              container.classList.add('content-enter');

              var data = 'openTabSilent';
              chrome.runtime.sendMessage(data, function (response) {
                if (chrome.runtime.lastError) {
                  console.error('Runtime error:', chrome.runtime.lastError.message);
                  return;
                }
              });
              return;
            }

            const caseIds = result.records.map(record => record.Id);

            chrome.runtime.sendMessage({
              action: 'syncPersistentCases',
              mode: currentMode,
              caseIds: caseIds
            }, async function (response) {
              if (chrome.runtime.lastError) {
                console.error('Runtime error syncing persistent cases:', chrome.runtime.lastError.message);
              } else if (response && response.success) {
                logger.info('cases.persistent.synced', { mode: currentMode, removed: response.removed, missing: (response.missingCases || []).length, tracking: response.count });
                // Verify assignment by ANY user before removing missing cases
                try {
                  const missingCases = Array.isArray(response.missingCases) ? response.missingCases : [];
                  if (missingCases.length > 0) {
                    const missingIds = missingCases.map(rc => rc.id);
                    const verified = await detectAssignmentsForCases(conn, { caseIds: missingIds });
                    if (verified && verified.processed && verified.processed.length > 0) {
                      // Remove only those confirmed as assigned
                      for (const p of verified.processed) {
                        try {
                          chrome.runtime.sendMessage({ action: 'removeCaseFromPersistentSet', caseId: p.caseId, reason: 'assigned-by-any' }, () => { });
                        } catch { }
                      }
                    }
                  }
                } catch (e) {
                  console.warn('Error while verifying missing persistent cases:', e);
                }
              }
            });

            chrome.runtime.sendMessage({
              action: 'addCasesToPersistentSet',
              cases: result.records,
              currentMode: currentMode,
              currentUserId: currentUserId,
              currentUserName: currentUserName
            }, function (response) {
              if (chrome.runtime.lastError) {
                console.error('Runtime error adding cases to persistent set:', chrome.runtime.lastError.message);
                return;
              }
              if (response && response.success) {
                logger.debug('cases.persistent.added', { message: response.message });

                // Trigger background processing for all incoming cases using new TrackedChange approach
                // Initiating new background processing
                chrome.runtime.sendMessage({
                  action: 'processCasesInBackgroundNew',
                  cases: result.records,
                  connectionInfo: {
                    serverUrl: 'https://orgcs.my.salesforce.com',
                    sessionId: SESSION_ID
                  },
                  currentMode: currentMode,
                  currentUserId: currentUserId,
                  currentUserName: currentUserName
                }, function (bgResponse) {
                  if (chrome.runtime.lastError) {
                    console.error('❌ Runtime error starting new background processing:', chrome.runtime.lastError.message);
                  } else if (bgResponse && bgResponse.success) {
                    // New background processing initiated
                    logger.info('cases.background.new.processing.started', { message: bgResponse.message });

                    // Start continuous CaseFeed processing to ensure all cases are processed
                    chrome.runtime.sendMessage({
                      action: 'startContinuousCaseFeedProcessing',
                      connectionInfo: {
                        serverUrl: 'https://orgcs.my.salesforce.com',
                        sessionId: SESSION_ID
                      },
                      currentMode: currentMode,
                      currentUserId: currentUserId,
                      currentUserName: currentUserName
                    }, function (continuousResponse) {
                      if (chrome.runtime.lastError) {
                        console.error('❌ Runtime error starting continuous processing:', chrome.runtime.lastError.message);
                      } else if (continuousResponse && continuousResponse.success) {
                        logger.info('cases.background.continuous.processing.started', { message: continuousResponse.message });
                      }
                    });
                  }
                });
              }
            });


            trackNewCaseFromHistory(conn, {
              caseIds,
              currentUserId,
              currentUserName,
              currentMode,
              strategy: 'firstManualByUser',
              removeFromPersistent: true
            });

            const commentQuery = `SELECT ParentId, Body, CreatedById, LastModifiedDate FROM CaseFeed WHERE Visibility = 'InternalUsers' AND ParentId IN ('${caseIds.join("','")}') AND Type = 'TextPost'`;

            conn.query(commentQuery, function (commentErr, commentResult) {
              logger.debug('cases.comments.fetched', { total: commentResult.records ? commentResult.records.length : 0 });

              if (commentErr) {
                return console.error(commentErr);
              }

              const actionedCaseIds = new Set();
              if (commentResult.records) {
                commentResult.records.forEach(record => {
                  if (record.Body && record.Body.includes('#SigQBmention')) {
                    logger.debug('cases.actioned.detected', { caseId: record.ParentId });
                    actionedCaseIds.add(record.ParentId);
                    if (record.CreatedById === currentUserId) {
                      const caseRecord = result.records.find(c => c.Id === record.ParentId);
                      logger.trace && logger.trace('cases.actioned.caseRecord', { caseId: caseRecord.Id });
                      if (caseRecord) {
                        const trackingKey = `tracked_${caseRecord.Id}`;
                        if (!localStorage.getItem(trackingKey)) {
                          trackActionAndCount(caseRecord.LastModifiedDate, caseRecord.CaseNumber, caseRecord.Severity_Level__c, 'New Case', caseRecord.CaseRoutingTaxonomy__r.Name.split('-')[0], currentMode, currentUserName, 'QB');
                          localStorage.setItem(trackingKey, 'true');

                          // Remove from persistent set as it's been processed
                          chrome.runtime.sendMessage({
                            action: 'removeCaseFromPersistentSet',
                            caseId: caseRecord.Id
                          }, function (response) {
                            if (chrome.runtime.lastError) {
                              console.error('Runtime error removing case from persistent set:', chrome.runtime.lastError.message);
                            }
                          });
                        }
                      }
                    }
                  }

                  if (record.Body && record.Body.includes('#GHOTriage')) {
                    logger.debug('gho.triage.comment.detected', { caseId: record.ParentId });

                    const commentDate = record.LastModifiedDate || record.CreatedDate;
                    const isCommentFromToday = isToday(commentDate);
                    const commentShift = getShiftForDate(commentDate);
                    const currentShift = getCurrentShift();
                    const isSameGeo = commentShift === currentShift;

                    logger.trace && logger.trace('gho.triage.comment.analysis', {
                      commentDate,
                      isFromToday: isCommentFromToday,
                      commentShift,
                      currentShift,
                      isSameGeo
                    });

                    if (isCommentFromToday && isSameGeo && record.CreatedById === currentUserId) {
                      const caseRecord = result.records.find(c => c.Id === record.ParentId);
                      logger.trace && logger.trace('gho.triage.caseRecord', { caseId: caseRecord.Id });
                      if (caseRecord) {
                        const ghoTrackingKey = `gho_tracked_${caseRecord.Id}`;
                        if (!localStorage.getItem(ghoTrackingKey)) {
                          trackActionAndCount(record.LastModifiedDate, caseRecord.CaseNumber, caseRecord.Severity_Level__c, 'GHO', caseRecord.CaseRoutingTaxonomy__r.Name.split('-')[0], currentMode, currentUserName, 'QB');
                          localStorage.setItem(ghoTrackingKey, 'true');
                          logger.info('gho.triage.tracked', { caseNumber: caseRecord.CaseNumber, date: commentDate, shift: commentShift });

                          chrome.runtime.sendMessage({
                            action: 'removeCaseFromPersistentSet',
                            caseId: caseRecord.Id
                          }, function (response) {
                            if (chrome.runtime.lastError) {
                              console.error('Runtime error removing GHO case from persistent set:', chrome.runtime.lastError.message);
                            }
                          });
                        }
                      }
                    } else {
                      logger.debug('gho.triage.comment.ignored', {
                        isFromToday: isCommentFromToday,
                        isSameGeo: isSameGeo,
                        commentDate,
                        commentShift,
                        currentShift
                      });
                    }
                  }
                });
              }

              var myHtml;
              var mvpWarningHtml = '';
              var isData = false;
              var today = new Date();
              var totalCasesCount = 0;
              var pendingCasesCount = 0;
              var pendingCasesDetails = [];

              // Weekend SLA thresholds
              const isWeekendFlag = isCurrentlyWeekend();
              var minSev1 = isWeekendFlag ? 0 : 5;
              var minSev2 = isWeekendFlag ? 0 : 20;
              if (result.records.length > 0) {
                isData = true;

                logger.info('cases.query.result', { total: result.records.length });

                totalCasesCount = result.records.length;

                const filteredRecords = result.records.filter(x => {
                  if (x.Contact && x.Contact.Is_MVP__c === true && x.SE_Initial_Response_Status__c === 'Met') {
                    logger.debug('cases.filter.mvp.met', { caseNumber: x.CaseNumber, status: x.SE_Initial_Response_Status__c });
                    return false;
                  }
                  return true;
                });

                logger.info('cases.filter.summary', { filtered: filteredRecords.length, original: result.records.length });

                for (var x in filteredRecords) {
                  const caseRecord = filteredRecords[x];
                  if (caseRecord.Contact && caseRecord.Contact.Is_MVP__c === true &&
                    (caseRecord.SE_Initial_Response_Status__c === 'In Warning' || caseRecord.SE_Initial_Response_Status__c === 'Warning')) {

                    const caseId = caseRecord.Id; // Fix: Define caseId
                    const snoozeUntil = localStorage.getItem('snooze_' + caseId);
                    if (snoozeUntil && new Date().getTime() < parseInt(snoozeUntil)) {
                      continue;
                    } else if (snoozeUntil) {
                      localStorage.removeItem('snooze_' + caseId);
                    }

                    if (actionedCaseIds.has(caseId)) {
                      sessionStorage.setItem(caseId, 'true');
                    }

                    const routingLogs = caseRecord.Case_Routing_Logs__r;
                    if (routingLogs && routingLogs.totalSize > 0) {
                      const lastLog = routingLogs.records[0];
                      if (lastLog.Transfer_Reason__c === 'GEO Locate' && caseRecord.SE_Initial_Response_Status__c === 'Met') {
                        sessionStorage.setItem(caseId, 'true');
                      }
                    }

                    const isActionTaken = (sessionStorage.getItem(caseId) || localStorage.getItem(caseId)) === 'true';
                    const caseData = {
                      number: caseRecord.CaseNumber,
                      severity: caseRecord.Severity_Level__c,
                      cloud: caseRecord.CaseRoutingTaxonomy__r.Name.split('-')[0],
                      isMVP: true
                    };
                    displayedCaseCount++;
                    if (isActionTaken) {
                      actionTakenCount++;
                    }

                    let statusColor = 'red';
                    let routingLogHtml = '';

                    if (routingLogs && routingLogs.totalSize > 0) {
                      const lastLog = routingLogs.records[0];
                      if (lastLog.Transfer_Reason__c && lastLog.Transfer_Reason__c !== 'New') {
                        let logText = lastLog.Transfer_Reason__c;

                        routingLogHtml = `
                          <div class="case-info-item">
                            <span class="checkmark">✓</span>
                            <span style="color: #9F2B68;">${logText} (${timeElapsed(new Date(lastLog.CreatedDate))})</span>
                          </div>
                        `;
                      }
                    }

                    const mvpCaseHtml = `
                      <div class="case-card mvp-warning-case">
                        <div class="case-header">
                          <div style="display: flex; align-items: center; gap: 8px;">
                            <span class="badge-soft badge-soft--danger">MVP URGENT</span>
                            <h3 class="case-title">${caseRecord.Subject}</h3>
                          </div>
                          <div class="case-timestamp">${formatDateWithDayOfWeek(caseRecord.CreatedDate)} (${timeElapsed(new Date(caseRecord.CreatedDate))})</div>
                        </div>
                        
                        <div class="case-details">
                          <div class="case-info">
                            <div class="case-info-item">
                              <span class="checkmark">✓</span>
                              <span>${caseRecord.Account.Name}</span>
                            </div>
                            <div class="case-info-item">
                              <span class="checkmark">✓</span>
                              <span>${caseRecord.CaseRoutingTaxonomy__r.Name}</span>
                            </div>
                            <div class="case-info-item">
                              <span class="checkmark">✓</span>
                              <span>${caseRecord.CaseNumber}</span>
                            </div>
                            <div class="case-info-item">
                              <span class="checkmark">✓</span>
                              <span>${caseRecord.Severity_Level__c}</span>
                            </div>
                            ${routingLogHtml}
                          </div>
                          
                          <div class="case-actions">
                            <a target="_blank" href="https://orgcs.my.salesforce.com/lightning/r/Case/${caseId}/view" 
                               class="preview-btn preview-record-btn" 
                               data-case-number="${caseData.number}" 
                               data-severity="${caseData.severity}" 
                               data-cloud="${caseData.cloud}"
                               data-is-mvp="true">
                              Preview Record
                            </a>
                            
                            <div class="action-controls">
                              <input type="checkbox" class="action-checkbox" 
                                     data-case-id="${caseId}" 
                                     ${isActionTaken ? 'checked' : ''} 
                                     data-case-number="${caseData.number}" 
                                     data-severity="${caseData.severity}" 
                                     data-cloud="${caseData.cloud}"
                                     data-is-mvp="true"
                                     disabled>
                              <span class="action-taken-text" style="display: ${isActionTaken ? 'inline' : 'none'};">Action taken</span>
                            </div>
                            
                            <div class="snooze-controls">
                              <select class="snooze-time" data-case-id="${caseId}">
                                <option value="5">5 mins</option>
                                <option value="10">10 mins</option>
                                <option value="15">15 mins</option>
                                <option value="20">20 mins</option>
                                <option value="custom">Custom</option>
                              </select>
                              <input type="number" class="custom-snooze-input" data-case-id="${caseId}" 
                                     placeholder="Minutes" min="1" max="1440" style="display: none;">
                              <button class="snooze-btn" data-case-id="${caseId}">Snooze</button>
                            </div>
                          </div>
                        </div>
                      </div>
                    `;

                    if (mvpWarningHtml) {
                      mvpWarningHtml = mvpWarningHtml + mvpCaseHtml;
                    } else {
                      mvpWarningHtml = mvpCaseHtml;
                    }
                  }
                }

                for (var x in filteredRecords) {
                  const caseId = filteredRecords[x].Id;

                  if (filteredRecords[x].Contact && filteredRecords[x].Contact.Is_MVP__c === true &&
                    (filteredRecords[x].SE_Initial_Response_Status__c === 'In Warning' || filteredRecords[x].SE_Initial_Response_Status__c === 'Warning')) {
                    continue;
                  }

                  if (actionedCaseIds.has(caseId)) {
                    sessionStorage.setItem(caseId, 'true');
                  }

                  const routingLogs = filteredRecords[x].Case_Routing_Logs__r;
                  if (routingLogs && routingLogs.totalSize > 0) {
                    const lastLog = routingLogs.records[0];
                    if (lastLog.Transfer_Reason__c === 'GEO Locate' || lastLog.Transfer_Reason__c === 'Dispatched' && filteredRecords[x].SE_Initial_Response_Status__c === 'Met') {
                      sessionStorage.setItem(caseId, 'true');
                      // Auto-actioned case due to GEO Locate routing with Met SLA
                    }
                    // Track manual action for customer calls when case is sent back to queue
                    if (lastLog.Back_To_Queue__c === true) {
                      const tr = lastLog.Transfer_Reason__c;
                      if (tr === 'New' || tr === 'Chat/Voice') {
                        try {
                          trackActionAndCount(
                            lastLog.CreatedDate,
                            filteredRecords[x].CaseNumber,
                            filteredRecords[x].Severity_Level__c,
                            'Customer Call',
                            filteredRecords[x].CaseRoutingTaxonomy__r.Name.split('-')[0],
                            currentMode,
                            currentUserName,
                            caseRecord.CaseRoutingTaxonomy__r.Name.split('-')[0] + ' Skills Queue'
                          );
                        } catch (e) { console.warn('Customer Call tracking failed', e); }
                      }
                    }
                  }

                  const isSLAM = filteredRecords[x].CaseRoutingTaxonomy__r.Name === 'Sales-Issues Developing for Salesforce Functions (Product)';
                  if (isSLAM && filteredRecords[x].support_available_timezone__c === '(GMT+09:00) Japan Standard Time (Asia/Tokyo)') {
                    sessionStorage.setItem(caseId, 'true');
                    // Auto-actioned SLAM case due to Japan timezone
                  }

                  // Processing case
                  const meetsAlertCriteria = (filteredRecords[x].CaseRoutingTaxonomy__r.Name == 'Sales-Issues Developing for Salesforce Functions (Product)') ||
                    (today >= addMinutes(minSev1, new Date(filteredRecords[x].CreatedDate)) && filteredRecords[x].Severity_Level__c == 'Level 1 - Critical') ||
                    (today >= addMinutes(minSev2, new Date(filteredRecords[x].CreatedDate)) && filteredRecords[x].Severity_Level__c == 'Level 2 - Urgent') ||
                    (filteredRecords[x].Contact.isMVP === true && filteredRecords[x].SE_Initial_Response_Status__c !== 'Met' && today >= addMinutes(minSev1, new Date(filteredRecords[x].CreatedDate)) && filteredRecords[x].Severity_Level__c == 'Level 1 - Critical') ||
                    (filteredRecords[x].Contact.isMVP === true && filteredRecords[x].SE_Initial_Response_Status__c !== 'Met' && today >= addMinutes(minSev2, new Date(filteredRecords[x].CreatedDate)) && filteredRecords[x].Severity_Level__c == 'Level 2 - Urgent') ||
                    (filteredRecords[x].Contact.isMVP === true && filteredRecords[x].SE_Initial_Response_Status__c !== 'Met' && today >= addMinutes((new Date(filteredRecords[x].SE_Target_Response__c) - new Date(filteredRecords[x].CreatedDate)) / (1000 * 60) - 30) && filteredRecords[x].Severity_Level__c == 'Level 3 - High') ||
                    (filteredRecords[x].Contact.isMVP === true && filteredRecords[x].SE_Initial_Response_Status__c !== 'Met' && today >= addMinutes((new Date(filteredRecords[x].SE_Target_Response__c) - new Date(filteredRecords[x].CreatedDate)) / (1000 * 60) - 30) && filteredRecords[x].Severity_Level__c == 'Level 4 - Medium');

                  if (meetsAlertCriteria) {
                    const snoozeUntil = localStorage.getItem('snooze_' + caseId);
                    if (snoozeUntil && new Date().getTime() < parseInt(snoozeUntil)) {
                      continue;
                    } else if (snoozeUntil) {
                      localStorage.removeItem('snooze_' + caseId);
                    }
                    const isActionTaken = (sessionStorage.getItem(caseId) || localStorage.getItem(caseId)) === 'true';
                    const caseData = {
                      number: filteredRecords[x].CaseNumber,
                      severity: filteredRecords[x].Severity_Level__c,
                      cloud: filteredRecords[x].CaseRoutingTaxonomy__r.Name.split('-')[0],
                      isMVP: filteredRecords[x].Contact && filteredRecords[x].Contact.Is_MVP__c === true,
                      isSLAM: isSLAM
                    };
                    displayedCaseCount++;
                    if (isActionTaken) {
                      actionTakenCount++;
                    }
                    let statusColor = '';
                    if (filteredRecords[x].SE_Initial_Response_Status__c === 'Met') {
                      statusColor = 'green';
                    } else if (filteredRecords[x].SE_Initial_Response_Status__c === 'In Warning' || filteredRecords[x].SE_Initial_Response_Status__c === 'Warning') {
                      statusColor = 'red';
                    }
                    let routingLogHtml = '';
                    if (routingLogs && routingLogs.totalSize > 0) {
                      const lastLog = routingLogs.records[0];
                      if (lastLog.Transfer_Reason__c && lastLog.Transfer_Reason__c !== 'New') {
                        let logText = lastLog.Transfer_Reason__c;

                        routingLogHtml = `
                          <div class="case-info-item">
                            <span class="checkmark">✓</span>
                            <span style="color: #9F2B68;">${logText} (${timeElapsed(new Date(lastLog.CreatedDate))})</span>
                          </div>
                        `;
                      }
                    }

                    if (routingLogs && routingLogs.totalSize > 1) {
                      const lastLog2 = routingLogs.records[0];
                      if (lastLog2.Severity_New_Value__c && lastLog2.Severity_Old_Value__c && (!lastLog2.Transfer_Reason__c || lastLog2.Transfer_Reason__c === '')) {

                        routingLogHtml = `
                          <div class="case-info-item">
                            <span class="checkmark">✓</span>
                            <span style="color: #9F2B68;">Severity changed to ${lastLog2.Severity_New_Value__c} (${timeElapsed(new Date(lastLog2.CreatedDate))})</span>
                          </div>
                        `;
                      }
                    }

                    const newHtml = `
                      <div class="case-card ${caseData.isMVP ? 'mvp-case card-accent-purple' : ''}">
                        <div class="case-header">
                          <div style="display: flex; align-items: center; gap: 8px;">
                            ${caseData.isMVP ? '<span class="badge-soft badge-soft--purple">MVP</span>' : ''}
                            ${caseData.isSLAM ? '<span class="badge-soft badge-soft--danger">SLAM</span>' : ''}
                            <h3 class="case-title">${filteredRecords[x].Subject}</h3>
                          </div>
                          <div class="case-timestamp">${formatDateWithDayOfWeek(filteredRecords[x].CreatedDate)} (${timeElapsed(new Date(filteredRecords[x].CreatedDate))})</div>
                        </div>
                        
                        <div class="case-details">
                          <div class="case-info">
                            <div class="case-info-item">
                              <span class="checkmark">✓</span>
                              <span>${filteredRecords[x].Account.Name}</span>
                            </div>
                            <div class="case-info-item">
                              <span class="checkmark">✓</span>
                              <span>${filteredRecords[x].CaseRoutingTaxonomy__r.Name}</span>
                            </div>
                            <div class="case-info-item">
                              <span class="checkmark">✓</span>
                              <span>${filteredRecords[x].CaseNumber}</span>
                            </div>
                            <div class="case-info-item">
                              <span class="checkmark">✓</span>
                              <span>${filteredRecords[x].Severity_Level__c}</span>
                            </div>
                            <div class="case-info-item">
                              <span class="checkmark">✓</span>
                              <span>${filteredRecords[x].SE_Initial_Response_Status__c}</span>
                            </div>
                            ${routingLogHtml}
                          </div>
                          
                          <div class="case-actions">
                            <a target="_blank" href="https://orgcs.my.salesforce.com/lightning/r/Case/${caseId}/view" 
                               class="preview-btn preview-record-btn" 
                               data-case-number="${caseData.number}" 
                               data-severity="${caseData.severity}" 
                               data-cloud="${caseData.cloud}"
                               data-is-mvp="${caseData.isMVP}"
                               data-is-slam="${caseData.isSLAM}">
                              Preview Record
                            </a>
                            
                            <div class="action-controls">
                              <input type="checkbox" class="action-checkbox" 
                                     data-case-id="${caseId}" 
                                     ${isActionTaken ? 'checked' : ''} 
                                     data-case-number="${caseData.number}" 
                                     data-severity="${caseData.severity}" 
                                     data-cloud="${caseData.cloud}"
                                     data-is-mvp="${caseData.isMVP}"
                                     data-is-slam="${caseData.isSLAM}"
                                     disabled>
                              <span class="action-taken-text" style="display: ${isActionTaken ? 'inline' : 'none'};">Action taken</span>
                            </div>
                            
                            <div class="snooze-controls">
                              <select class="snooze-time" data-case-id="${caseId}">
                                <option value="5">5 mins</option>
                                <option value="10">10 mins</option>
                                <option value="15">15 mins</option>
                                <option value="20">20 mins</option>
                                <option value="custom">Custom</option>
                              </select>
                              <input type="number" class="custom-snooze-input" data-case-id="${caseId}" 
                                     placeholder="Minutes" min="1" max="1440" style="display: none;">
                              <button class="snooze-btn" data-case-id="${caseId}">Snooze</button>
                            </div>
                          </div>
                        </div>
                      </div>
                    `;
                    if (myHtml) {
                      myHtml = myHtml + newHtml;
                    } else {
                      myHtml = newHtml;
                    }
                  } else {
                    pendingCasesCount++;
                    const minutesSinceCreation = Math.floor((today - new Date(filteredRecords[x].CreatedDate)) / (1000 * 60));
                    const requiredMinutes = filteredRecords[x].Severity_Level__c === 'Level 1 - Critical' ? minSev1 : filteredRecords[x].Severity_Level__c === 'Level 2 - Urgent' ? minSev2 : (new Date(filteredRecords[x].SE_Target_Response__c) - new Date(filteredRecords[x].CreatedDate)) / (1000 * 60) - 30;
                    const remainingMinutes = Math.round(requiredMinutes - minutesSinceCreation);

                    pendingCasesDetails.push({
                      caseNumber: filteredRecords[x].CaseNumber,
                      severity: filteredRecords[x].Severity_Level__c,
                      account: filteredRecords[x].Account.Name,
                      createdDate: new Date(filteredRecords[x].CreatedDate),
                      minutesSinceCreation: minutesSinceCreation,
                      remainingMinutes: Math.max(0, remainingMinutes),
                      totalRequiredMinutes: Math.max(0, Math.round(requiredMinutes)),
                      isMVP: filteredRecords[x].Contact && filteredRecords[x].Contact.Is_MVP__c === true,
                      initialResponseStatus: filteredRecords[x].SE_Initial_Response_Status__c
                    });

                    // Case pending alert criteria logged
                  }
                }

              }
              if (isData && (mvpWarningHtml || myHtml)) {
                let finalHtml = '';
                if (mvpWarningHtml) {
                  finalHtml += mvpWarningHtml;
                }
                if (myHtml) {
                  finalHtml += myHtml;
                }

                if (pendingCasesCount > 0) {
                  const pendingGridHtml = buildPendingCardsHtml(pendingCasesDetails);
                  const isWeekendSLA_banner = (minSev1 === 0 && minSev2 === 0);
                  const subTitleText = isWeekendSLA_banner
                    ? `${pendingCasesCount} case${pendingCasesCount === 1 ? '' : 's'} being monitored`
                    : `${pendingCasesCount} pending within SLA · SEV1: ${minSev1}m · SEV2: ${minSev2}m`;
                  finalHtml += getPendingSectionHtml({
                    title: 'Cases in Queue',
                    pendingCasesCount,
                    subTitleText,
                    pendingGridHtml
                  });
                }

                const container = document.getElementById("parentSigSev2");
                container.classList.remove('is-loading');
                container.innerHTML = finalHtml;
                container.classList.add('content-enter');

                startPendingLiveUpdater();

                const savedFilter = localStorage.getItem('caseFilter');
                if (savedFilter) {
                  applyFilter(savedFilter);
                }

                const hasUnactionedCases = actionTakenCount < displayedCaseCount;

                if (hasUnactionedCases) {
                  var audio = new Audio('../assets/audio/notification.wav');
                  audio.play();
                  var data = 'openTab';
                  chrome.runtime.sendMessage(data, function (response) {
                    if (chrome.runtime.lastError) {
                      console.error('Runtime error opening tab:', chrome.runtime.lastError.message);
                      return;
                    }
                  });
                } else {
                  var data = 'openTabSilent';
                  chrome.runtime.sendMessage(data, function (response) {
                    if (chrome.runtime.lastError) {
                      console.error('Runtime error opening tab silently:', chrome.runtime.lastError.message);
                      return;
                    }
                  });
                }
              } else {
                let noCasesHtml;

                if (pendingCasesCount > 0) {
                  const pendingGridHtml = buildPendingCardsHtml(pendingCasesDetails);
                  const isWeekendSLA_nc = (minSev1 === 0 && minSev2 === 0);
                  const subTitleText_nc = isWeekendSLA_nc
                    ? `${pendingCasesCount} case${pendingCasesCount === 1 ? '' : 's'} being monitored`
                    : `${pendingCasesCount} pending within SLA · SEV1: ${minSev1}m · SEV2: ${minSev2}m`;
                  noCasesHtml = getPendingSectionHtml({
                    title: 'No Action Required',
                    pendingCasesCount,
                    subTitleText: subTitleText_nc,
                    pendingGridHtml
                  });
                } else if (totalCasesCount > 0 && displayedCaseCount === 0) {
                  noCasesHtml = `
                    <div class="no-cases-message">
                      <h4 class="no-cases-title">No Cases for Now</h4>
                      <p class="no-cases-text">All ${totalCasesCount} case${totalCasesCount === 1 ? ' is' : 's are'} assigned or handled. Check back later!</p>
                      <p class="mode-switch-hint">💡 Try switching between Signature and Premier modes to see different case types.</p>
                    </div>
                  `;
                } else {
                  noCasesHtml = `
                    <div class="no-cases-message">
                      <h4 class="no-cases-title">No Cases to Action</h4>
                      <p class="no-cases-text">All cases are up to date. Great work!</p>
                      <p class="mode-switch-hint">💡 Try switching between Signature and Premier modes to see different case types.</p>
                    </div>
                  `;
                }

                const container = document.getElementById("parentSigSev2");
                container.classList.remove('is-loading');
                container.innerHTML = noCasesHtml;
                container.classList.add('content-enter');
                startPendingLiveUpdater();
                var data = 'openTabSilent';
                chrome.runtime.sendMessage(data, function (response) {
                  if (chrome.runtime.lastError) {
                    console.error('Runtime error opening tab silently (no cases):', chrome.runtime.lastError.message);
                    return;
                  }
                });
              }
            });
          });
      });
    }
  });
}

// Function to resume continuous processing when user becomes active
function resumeContinuousProcessing(currentUserId, currentUserName, currentMode) {
  try {
    if (!SESSION_ID) {
      console.log('No valid session, skipping continuous processing resume');
      return;
    }

    chrome.runtime.sendMessage({
      action: 'startContinuousCaseFeedProcessing',
      connectionInfo: {
        serverUrl: 'https://orgcs.my.salesforce.com',
        sessionId: SESSION_ID
      },
      currentMode: currentMode,
      currentUserId: currentUserId,
      currentUserName: currentUserName
    }, function (response) {
      if (chrome.runtime.lastError) {
        console.error('❌ Runtime error resuming continuous processing:', chrome.runtime.lastError.message);
      } else if (response && response.success) {
        logger.info('cases.background.continuous.processing.resumed', { message: response.message });
      }
    });
  } catch (error) {
    console.error('Error resuming continuous processing:', error);
  }
}


// GHO Alert System Functions
function checkGHOAlert() {
  const now = new Date();
  const currentTime = now.getHours() * 60 + now.getMinutes();
  const today = now.toDateString();

  const apacAlertTime = 7 * 60 + 30;
  const emeaAlertTime = 14 * 60 + 30;
  const amerAlertTime = 22 * 60 + 30;

  const isAPACTime = Math.abs(currentTime - apacAlertTime) <= 5;
  const isEMEATime = Math.abs(currentTime - emeaAlertTime) <= 5;
  const isAMERTime = Math.abs(currentTime - amerAlertTime) <= 5;

  if (!(isAPACTime || isEMEATime || isAMERTime)) {
    return;
  }

  let region = '';
  if (isAPACTime) region = 'APAC';
  else if (isEMEATime) region = 'EMEA';
  else if (isAMERTime) region = 'AMER';

  const alertKey = `gho_alert_${region}_${today}_${currentUserName}`;

  if (localStorage.getItem(alertKey)) {
    return;
  }

  const preferredShiftValues = getPreferredShiftValues(region);
  const shiftCondition = buildPreferredShiftCondition(preferredShiftValues);

  const ghoQuery = `SELECT Id, CreatedDate, Account.Name, Owner.Name, SE_Target_Response__c, Severity_Level__c, CaseNumber, Subject, CaseRoutingTaxonomy__r.Name, SE_Initial_Response_Status__c, Contact.Is_MVP__c, support_available_timezone__c, (SELECT Transfer_Reason__c, CreatedDate, CreatedById, Preferred_Shift_Old_Value__c, Preferred_Shift_New_Value__c, Severity_New_Value__c, Severity_Old_Value__c FROM Case_Routing_Logs__r ORDER BY CreatedDate DESC LIMIT 20) FROM Case WHERE ((Owner.Name IN ('Skills Queue','Kase Changer', 'Working in Org62','GHO Queue') AND ((Case_Support_level__c='Premier Priority' OR Case_Support_level__c='Signature' OR Case_Support_level__c='Signature Success') OR (Case_Support_level__c='Signature' OR Case_Support_level__c='Premier Priority' OR Case_Support_level__c='Signature Success'))) OR (Contact.Is_MVP__c=true AND Owner.Name='GHO Queue')) AND IsClosed=false AND ${shiftCondition} AND ((CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry%' OR CaseRoutingTaxonomy__r.Name LIKE 'Community-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Scale Center%' OR CaseRoutingTaxonomy__r.Name LIKE 'Customer Success Score%' OR CaseRoutingTaxonomy__r.Name LIKE 'Data Cloud-%') AND (Severity_Level__c='Level 1 - Critical' OR Severity_Level__c='Level 2 - Urgent')) AND CaseRoutingTaxonomy__r.Name NOT IN ('Disability and Product Accessibility','DORA')`;

  if (!SESSION_ID) {
    console.error('SESSION_ID not available for GHO alert');
    return;
  }

  let ghoConn = new jsforce.Connection({
    serverUrl: 'https://orgcs.my.salesforce.com',
    sessionId: SESSION_ID,
    version: '64.0',
  });

  ghoConn.query(ghoQuery, function (err, result) {
    if (err) {
      console.error('GHO Alert Query Error:', err);
      return;
    }

    if (!result.records || result.records.length === 0) {
      localStorage.setItem(alertKey, 'true');
      return;
    }

    const caseIds = result.records.map(record => record.Id);
    const commentQuery = `SELECT ParentId, Body, CreatedById, LastModifiedDate FROM CaseFeed WHERE Visibility = 'InternalUsers' AND ParentId IN ('${caseIds.join("','")}') AND Type = 'TextPost'`;

    ghoConn.query(commentQuery, function (commentErr, commentResult) {
      if (commentErr) {
        console.error('GHO Comment Query Error:', commentErr);
        return;
      }

      ghoConn.query(`SELECT Id FROM User WHERE Name = '${currentUserName}' AND IsActive = True AND Username LIKE '%orgcs.com'`, function (userErr, userResult) {
        if (userErr) {
          console.error('GHO User Query Error:', userErr);
          return;
        }

        let currentUserId = null;
        if (userResult.records.length > 0) {
          currentUserId = userResult.records[0].Id;
        }

        const actionedCaseIds = new Set();
        if (commentResult.records) {
          commentResult.records.forEach(comment => {
            if (comment.Body && comment.Body.includes('#GHOTriage')) {
              const commentDate = comment.LastModifiedDate || comment.CreatedDate;
              const isCommentFromToday = isToday(commentDate);
              const commentShift = getShiftForDate(commentDate);
              const isSameGeo = commentShift === region;

              if (isCommentFromToday && isSameGeo) {
                actionedCaseIds.add(comment.ParentId);

                if (currentUserId && comment.CreatedById === currentUserId) {
                  const caseRecord = result.records.find(c => c.Id === comment.ParentId);
                  if (caseRecord) {
                    const ghoTrackingKey = `gho_tracked_${caseRecord.Id}`;
                    if (!localStorage.getItem(ghoTrackingKey)) {
                      trackActionAndCount(comment.LastModifiedDate, caseRecord.CaseNumber, caseRecord.Severity_Level__c, 'GHO', caseRecord.CaseRoutingTaxonomy__r.Name.split('-')[0], currentMode, currentUserName, 'QB');
                      localStorage.setItem(ghoTrackingKey, 'true');
                    }
                  }
                }
              }
            }
          });
        }
        const unactionedCases = result.records.filter(caseRecord => !actionedCaseIds.has(caseRecord.Id));

        if (unactionedCases.length > 0) {
          try {
            showGHOAlert(region, unactionedCases, alertKey);
          } catch (error) {
            console.error('Failed to show GHO alert:', error);
          }
        } else {
          localStorage.setItem(alertKey, 'true');
        }
      });
    });
  });
}

function showGHOAlert(region, ghoRecords, alertKey) {
  try {
    const alertTime = region === 'APAC' ? '7:30 AM' : region === 'EMEA' ? '2:30 PM' : '10:30 PM';

    let existingModal = document.getElementById('gho-alert-modal');
    if (existingModal) {
      existingModal.remove();
    }

    // Inject scoped styles for the GHO alert modal if not already present
    let existingStyles = document.getElementById('gho-alert-styles');
    if (!existingStyles) {
      const styleEl = document.createElement('style');
      styleEl.id = 'gho-alert-styles';
      styleEl.textContent = `
        #gho-alert-modal { background: rgba(17, 24, 39, 0.50); backdrop-filter: blur(6px); }
        #gho-alert-modal .modal-content { width: 100%; max-width: 680px; background: #ffffff; border-radius: 18px; overflow: hidden; box-shadow: 0 24px 55px rgba(2, 6, 23, 0.20), 0 10px 24px rgba(2, 6, 23, 0.12); border: 1px solid #e5e7eb; transform: translateY(10px) scale(.98); opacity: 0; transition: transform .28s cubic-bezier(.22,.61,.36,1), opacity .24s ease; }
        #gho-alert-modal.modal-show .modal-content { transform: translateY(0) scale(1); opacity: 1; }
        #gho-alert-modal .modal-header { display: flex; align-items: center; justify-content: space-between; padding: 14px 18px; background: #fff7ed; color: #9a3412; border-bottom: 1px solid #fed7aa; }
        #gho-alert-modal .modal-header h3 { display: flex; align-items: center; gap: 10px; font-size: 18px; margin: 0; }
        #gho-alert-modal .modal-header h3 i { background: #ffedd5; border: 1px solid #fed7aa; width: 26px; height: 26px; display: inline-flex; align-items: center; justify-content: center; border-radius: 999px; box-shadow: inset 0 1px 0 rgba(255,255,255,.6); color: #b45309; }
        #gho-alert-modal .modal-header .count-pill { background: #ffedd5; border: 1px solid #fed7aa; color: #9a3412; padding: 2px 8px; border-radius: 999px; font-weight: 700; font-size: 12px; }
        #gho-alert-modal .modal-header .modal-close { color: #374151 !important; opacity: .9; transition: transform .15s ease, opacity .15s ease; }
        #gho-alert-modal .modal-header .modal-close:hover { transform: scale(1.08); opacity: 1; }
        #gho-alert-modal .modal-body { padding: 20px; background: linear-gradient(180deg, #ffffff 0%, #fafafa 100%); }
        #gho-alert-modal .intro { text-align: center; margin-bottom: 16px; }
        #gho-alert-modal .intro h4 { color: #b45309; margin: 0 0 6px 0; font-size: 16px; }
        #gho-alert-modal .intro p { color: #6b7280; font-size: 14px; margin: 0; }
        #gho-alert-modal .case-wrap { background: #fff7ed; border: 1px solid #fed7aa; border-radius: 14px; padding: 12px; margin-bottom: 18px; }
        #gho-alert-modal .case-wrap h5 { color: #9a3412; margin: 0 0 10px 0; font-size: 14px; }
        #gho-alert-modal .case-list { max-height: 280px; overflow-y: auto; }
        #gho-alert-modal .case-list::-webkit-scrollbar { width: 8px; }
        #gho-alert-modal .case-list::-webkit-scrollbar-thumb { background: #e5e7eb; border-radius: 999px; }
        #gho-alert-modal .case-item { display: flex; justify-content: space-between; align-items: center; padding: 10px 8px; border-bottom: 1px solid #fed7aa; border-radius: 10px; transition: background .15s ease; }
        #gho-alert-modal .case-item:last-child { border-bottom: none; }
        #gho-alert-modal .case-item:hover { background: rgba(255, 237, 213, .35); }
        #gho-alert-modal .case-left { display: flex; align-items: center; gap: 10px; }
        #gho-alert-modal .case-number { color: #9a3412; font-weight: 700; }
        #gho-alert-modal .meta { color: #6b7280; font-size: 12px; display: flex; align-items: center; gap: 6px; }
        #gho-alert-modal .case-right { display: flex; align-items: center; gap: 12px; }
        #gho-alert-modal .action-link { color: #4b5563; font-size: 12px; text-decoration: none; display: inline-flex; align-items: center; gap: 6px; padding: 6px 8px; border-radius: 8px; border: 1px solid transparent; transition: all .15s ease; }
        #gho-alert-modal .action-link i { color: #0176d3; }
        #gho-alert-modal .action-link:hover { color: #0176d3; background: #eef6fd; border-color: #d8ecfd; }
        #gho-alert-modal .action-link:focus-visible { outline: 2px solid #bfdbfe; outline-offset: 2px; }
        #gho-alert-modal .chip { background-color: #7c3aed; color: #fff; padding: 2px 6px; border-radius: 999px; font-size: 10px; font-weight: 700; margin-left: 4px; }
        #gho-alert-modal .gho-badge { padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 700; margin-left: 8px; border: 1px solid transparent; }
        #gho-alert-modal .gho-badge.sev-1 { background: #fee2e2; color: #b91c1c; border-color: #fecaca; }
        #gho-alert-modal .gho-badge.sev-2 { background: #ffedd5; color: #9a3412; border-color: #fed7aa; }
        #gho-alert-modal .gho-badge.sev-3 { background: #fef3c7; color: #92400e; border-color: #fde68a; }
        #gho-alert-modal .gho-badge.sev-4 { background: #dbeafe; color: #1d4ed8; border-color: #bfdbfe; }
        #gho-alert-modal .footer-actions { display: flex; gap: 12px; justify-content: center; }
        #gho-alert-modal .btn-primary { background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%); color: #fff; border: none; padding: 12px 20px; border-radius: 12px; font-weight: 700; box-shadow: 0 6px 16px rgba(245, 158, 11, 0.18); transition: transform .12s ease, box-shadow .12s ease; }
        #gho-alert-modal .btn-primary:hover { transform: translateY(-1px); box-shadow: 0 10px 22px rgba(245, 158, 11, 0.24); }
        #gho-alert-modal .btn-secondary { background: #6b7280; color: #fff; border: none; padding: 12px 20px; border-radius: 12px; font-weight: 600; transition: transform .12s ease, opacity .12s ease; }
        #gho-alert-modal .btn-secondary:hover { transform: translateY(-1px); opacity: .96; }
      `;
      document.head.appendChild(styleEl);
    }

    const getSeverityClass = (sev) => {
      const value = String(sev || '').toUpperCase();
      if (value.includes('1')) return 'sev-1';
      if (value.includes('2')) return 'sev-2';
      if (value.includes('3')) return 'sev-3';
      return 'sev-4';
    };

    const modalHtml = `
      <div id="gho-alert-modal" class="modal-overlay" style="display: flex; z-index: 1001;">
        <div class="modal-content" style="max-width: 640px;">
          <div class="modal-header">
            <h3>
              <i class="fa fa-exclamation-triangle" aria-hidden="true"></i>
              <span>GHO Alert – ${region} Region</span>
              <span class="count-pill">${ghoRecords.length}</span>
            </h3>
            <span class="modal-close" id="gho-alert-close" style="color: white; cursor: pointer; font-size: 24px;" aria-label="Close">&times;</span>
          </div>
          <div class="modal-body">
            <div class="intro">
              <h4>Daily GHO Check – ${alertTime} IST</h4>
              <p>Found ${ghoRecords.length} GHO case${ghoRecords.length === 1 ? '' : 's'} requiring attention for ${region} shift</p>
            </div>
            <div class="case-wrap">
              <h5>Cases Needing Action</h5>
              <div class="case-list">
                ${ghoRecords.map(caseRecord => `
                  <div class="case-item">
                    <div class="case-left">
                      <i class="fa fa-life-ring" aria-hidden="true" style="color: #f59e0b;"></i>
                      <div>
                        <div>
                          <span class="case-number">${caseRecord.CaseNumber}</span>
                          <span class="gho-badge ${getSeverityClass(caseRecord.Severity_Level__c)}">${caseRecord.Severity_Level__c || ''}</span>
                          ${caseRecord.Contact && caseRecord.Contact.Is_MVP__c ? '<span class="chip">MVP</span>' : ''}
                        </div>
                      </div>
                    </div>
                    <div class="case-right">
                      <div class="meta">
                        <i class="fa fa-clock-o" aria-hidden="true"></i>
                        <span>${timeElapsed(new Date(caseRecord.CreatedDate))}</span>
                      </div>
                      <a class="action-link" target="_blank" href="https://orgcs.my.salesforce.com/lightning/r/Case/${caseRecord.Id}/view" title="View in Salesforce">
                        <i class="fa fa-external-link" aria-hidden="true"></i>
                        <span>View in Salesforce</span>
                      </a>
                    </div>
                  </div>
                `).join('')}
              </div>
            </div>
            <div class="footer-actions">
              <button id="gho-alert-check" class="preview-btn btn-primary">
                Check GHO Cases
              </button>
              <button id="gho-alert-dismiss" class="preview-btn btn-secondary">
                Dismiss
              </button>
            </div>
          </div>
        </div>
      </div>
    `;

    document.body.insertAdjacentHTML('beforeend', modalHtml);

    const modal = document.getElementById('gho-alert-modal');
    if (modal) {
      requestAnimationFrame(() => modal.classList.add('modal-show'));
      modal.addEventListener('click', (e) => { if (e.target === modal) modal.remove(); });
    } else {
      console.error('Failed to add GHO alert modal to DOM');
      return;
    }

    try {
      const audio = new Audio('../assets/audio/ghoalert.wav');
      audio.volume = 1.0;
      audio.play().catch(e => {
        // Audio autoplay blocked or failed
      });
    } catch (e) {
      // Could not play GHO alert sound
    }

    document.getElementById('gho-alert-close').addEventListener('click', () => {
      document.getElementById('gho-alert-modal').remove();
    });

    document.getElementById('gho-alert-dismiss').addEventListener('click', () => {
      document.getElementById('gho-alert-modal').remove();
      localStorage.setItem(alertKey, 'true');
    });

    document.getElementById('gho-alert-check').addEventListener('click', () => {
      document.getElementById('gho-alert-modal').remove();
      checkGHOStatus();
    });
  } catch (error) {
    console.error('Error showing GHO alert:', error);
    throw error;
  }
}

function updateGHOButtonVisibility() {
  const ghoButton = document.getElementById("check-gho-button");
  if (ghoButton) {
    if (currentMode === 'signature' || currentMode === 'premier') {
      ghoButton.style.display = 'inline-block';
      const modeLabel = currentMode === 'signature' ? 'Signature' : 'Premier';
      const iconHtml = ghoButton.querySelector('i') ? ghoButton.querySelector('i').outerHTML + '\n\t\t' : '';
      ghoButton.innerHTML = `${iconHtml}${modeLabel} GHO Cases`;
    } else {
      ghoButton.style.display = 'none';
    }
  }
}

function updateCICButtonVisibility() {
  const cicButton = document.getElementById("cic-button");
  if (cicButton) {
    cicButton.style.display = (DEV_FORCE_SHOW_WEEKEND_ROSTER || isCurrentlyWeekend()) ? 'inline-block' : 'none';
  }
}

// Update the 3-dots menu user header (name + dummy avatar)
function updateMoreMenuUserHeader() {
  try {
    const nameEl = document.getElementById('more-menu-name');
    const avatarEl = document.getElementById('more-menu-avatar');
    if (!nameEl || !avatarEl) return;
    const nm = currentUserName || 'User';
    nameEl.textContent = nm;
    const initial = (nm || 'U').trim().charAt(0).toUpperCase();
    avatarEl.textContent = initial || 'U';
  } catch { }
}

// ================================================
// Voice Calls Modal and Tracking
// ================================================
async function showVoiceCallModal() {
  try {
    const modal = document.getElementById('voicecall-modal');
    const loading = document.getElementById('voicecall-loading');
    const container = document.getElementById('voicecall-container');
    const tbody = document.getElementById('voicecall-tbody');
    const trackBtn = document.getElementById('voicecall-track-btn');
    const selectAll = document.getElementById('voicecall-select-all');

    if (!modal || !loading || !container || !tbody || !trackBtn || !selectAll) return;

    // Reset state
    tbody.innerHTML = '';
    selectAll.checked = false;
    trackBtn.disabled = true;
    container.style.display = 'none';
    loading.style.display = 'block';

    // Show modal
    modal.style.display = 'flex';
    setTimeout(() => modal.classList.add('modal-show'), 0);

    if (!window.jsforce || !SESSION_ID) {
      loading.innerHTML = '<p style="color:#b91c1c;">Unable to query Salesforce: missing session. Please sign in.</p>';
      return;
    }

    // Resolve current user Id
    const conn = new jsforce.Connection({
      serverUrl: 'https://orgcs.my.salesforce.com',
      sessionId: SESSION_ID,
      version: '64.0',
    });
    let userId = null;
    try {
      const ures = await conn.query(`SELECT Id FROM User WHERE Name = '${currentUserName}' AND IsActive = True AND Username LIKE '%orgcs.com'`);
      if (ures && ures.records && ures.records.length) userId = ures.records[0].Id;
    } catch (e) {
    }

    const daysInputEl = document.getElementById('voicecall-days-input');
    let days = 2;
    try {
      const parsed = parseInt(daysInputEl && daysInputEl.value ? daysInputEl.value : '2', 10);
      if (!isNaN(parsed) && parsed >= 1 && parsed <= 30) days = parsed;
    } catch { }
    const soql = `SELECT Account__c, CallDurationInSeconds, Case__c, Contact__c, Contact_Is_DC__c, CreatedById, CreatedDate, DisconnectReason, FromPhoneNumber, Id, Name, OwnerId, QueueName, RecipientId, Related_Issue__c, ToPhoneNumber, UserId,Voice_Issue_Summary__c, Voice_Resolution_Summary__c, Voice_Summary__c FROM VoiceCall WHERE OwnerId = '${userId}' AND QueueName='CS Signature Hotline' AND CreatedDate=LAST_N_DAYS:${days} ORDER BY CreatedDate DESC`;

    let result;
    try {
      result = await conn.query(soql);
    } catch (e) {
      loading.innerHTML = `<p style="color:#b91c1c;">Query failed: ${e.message}</p>`;
      return;
    }

    const records = (result && result.records) ? result.records : [];
    if (!records.length) {
      loading.innerHTML = '<p style="color:#374151;">No recent Voice Calls found.</p>';
      return;
    }

    // Resolve Case Id -> CaseNumber for better display
    const caseNumMap = {};
    try {
      const caseIds = Array.from(new Set(records.map(r => r.Case__c).filter(id => id && typeof id === 'string' && id.startsWith('500'))));
      const CHUNK = 100;
      for (let i = 0; i < caseIds.length; i += CHUNK) {
        const slice = caseIds.slice(i, i + CHUNK);
        const cres = await conn.query(`SELECT Id, CaseNumber FROM Case WHERE Id IN ('${slice.join("','")}')`);
        (cres.records || []).forEach(c => { if (c && c.Id) caseNumMap[c.Id] = c.CaseNumber; });
      }
    } catch { }

    // Render rows
    const fmt = (d) => {
      try { return new Date(d).toLocaleString(); } catch { return d; }
    };
    const dur = (s) => {
      const n = parseInt(s || 0, 10);
      const m = Math.floor(n / 60); const r = n % 60; return `${m}m ${r}s`;
    };
    records.forEach((r) => {
      const tr = document.createElement('tr');
      tr.style.borderBottom = '1px solid #e5e7eb';
      const caseDisplay = r.Case__c ? (caseNumMap[r.Case__c] || r.Case__c) : '';
      tr.innerHTML = `
        <td style="padding:8px; text-align:center;"><input type="checkbox" class="voicecall-row-select" data-case="${r.Case__c || ''}" data-created="${r.CreatedDate || ''}"></td>
        <td>${fmt(r.CreatedDate)}</td>
        <td>${r.FromPhoneNumber || ''}</td>
        <td>${r.ToPhoneNumber || ''}</td>
        <td>${dur(r.CallDurationInSeconds)}</td>
        <td>${caseDisplay}</td>
        <td>${r.QueueName || ''}</td>
      `;
      tbody.appendChild(tr);
    });

    // Enable UI
    loading.style.display = 'none';
    container.style.display = 'block';

    // Selection logic
    const updateTrackBtn = () => {
      const any = !!document.querySelector('.voicecall-row-select:checked');
      trackBtn.disabled = !any;
    };
    selectAll.addEventListener('change', () => {
      document.querySelectorAll('.voicecall-row-select').forEach(cb => cb.checked = selectAll.checked);
      updateTrackBtn();
    });
    tbody.addEventListener('change', (e) => {
      if (e.target && e.target.classList.contains('voicecall-row-select')) updateTrackBtn();
    });

    // Re-run query on Apply
    const applyBtn = document.getElementById('voicecall-days-apply');
    if (applyBtn) {
      applyBtn.onclick = () => {
        // simple reload of modal for new filter
        showVoiceCallModal();
      };
    }

    // Track selected calls
    trackBtn.onclick = async () => {
      try {
        const selected = Array.from(document.querySelectorAll('.voicecall-row-select:checked'));
        if (!selected.length) return;
        for (const cb of selected) {
          const caseIdOrNumber = cb.getAttribute('data-case') || '';
          const created = cb.getAttribute('data-created');

          // If we have a Case Number, use it; if it's an Id, try to resolve number quickly
          let caseNumber = caseIdOrNumber;
          let severity = '';
          let cloud = '';
          try {
            if (caseIdOrNumber && caseIdOrNumber.startsWith('500')) {
              if (caseNumMap[caseIdOrNumber]) {
                caseNumber = caseNumMap[caseIdOrNumber] || caseNumber;
              } else {
                const cres = await conn.query(`SELECT CaseNumber, Severity_Level__c, CaseRoutingTaxonomy__r.Name FROM Case WHERE Id='${caseIdOrNumber}' LIMIT 1`);
                if (cres && cres.records && cres.records.length) {
                  caseNumber = cres.records[0].CaseNumber || caseNumber;
                  severity = cres.records[0].Severity_Level__c || '';
                  cloud = (cres.records[0].CaseRoutingTaxonomy__r && cres.records[0].CaseRoutingTaxonomy__r.Name ? cres.records[0].CaseRoutingTaxonomy__r.Name.split('-')[0] : '') || '';
                }
              }
            }
          } catch { }

          // Default to signature sheet regardless of currentMode per requirement
          const dateOfAction = created ? new Date(created) : new Date();
          trackActionAndCount(dateOfAction, caseNumber || '--', severity || '--', 'Customer Call', cloud || 'Signature', 'signature', currentUserName, '--');
        }
        showToast(`Tracked ${selected.length} call${selected.length > 1 ? 's' : ''}`);
        // Close modal after tracking
        const modal = document.getElementById('voicecall-modal');
        modal.classList.remove('modal-show');
        setTimeout(() => modal.style.display = 'none', 150);
      } catch (err) {
        console.warn('Tracking voice calls failed', err);
        showToast('Failed to track calls');
      }
    };
  } catch (e) {
    console.warn('showVoiceCallModal failed', e);
    try {
      const loading = document.getElementById('voicecall-loading');
      if (loading) loading.innerHTML = `<p style="color:#b91c1c;">${e.message}</p>`;
    } catch { }
  }
}

// Function to ensure only one extension tab exists
async function ensureSingleTab() {
  try {
    chrome.runtime.sendMessage('ensureSingleTab', function (response) {
      if (chrome.runtime.lastError) {
        console.error('Runtime error ensuring single tab:', chrome.runtime.lastError.message);
        return;
      }
    });
  } catch (error) {
    console.error('Error requesting tab cleanup:', error);
  }
}

document.addEventListener('keydown', function (e) {
  if (e.metaKey || e.ctrlKey) {
    switch (e.key.toLowerCase()) {
      case 'f':
        e.preventDefault();
        const searchInput = document.getElementById("search-input");
        if (searchInput) {
          searchInput.focus();
          searchInput.select();
        }
        break;

      case 'g':
        e.preventDefault();
        if (currentMode === 'signature') {
          checkGHOStatus();
        } else {
          showToast('GHO status is only available in Signature mode');
        }
        break;

      case 'm':
        e.preventDefault();
        if (currentMode === 'signature') {
          showGHOMappingModal();
        } else {
          showToast('GHO mapping is only available in Signature mode');
        }
        break;

      case 'r':
        e.preventDefault();
        const refreshIcon = document.querySelector('#refresh-button .fa-refresh');
        if (refreshIcon) {
          refreshIcon.classList.add('fa-spin');
        }
        setTimeout(() => {
          window.location.reload();
        }, 500);
        break;

      case 'R':
        e.preventDefault();
        // If GHO modal is open, refresh GHO data; otherwise refresh page
        const ghoModal = document.getElementById('gho-modal');
        if (ghoModal && ghoModal.style.display === 'flex') {
          forceRefreshGHOData();
        } else {
          const refreshIcon = document.querySelector('#refresh-button .fa-refresh');
          if (refreshIcon) {
            refreshIcon.classList.add('fa-spin');
          }
          setTimeout(() => {
            window.location.reload();
          }, 500);
        }
        break;

      case 'i':
        e.preventDefault();
        const modeSwitchSig2 = document.getElementById("mode-switch");
        if (modeSwitchSig2 && modeSwitchSig2.checked) {
          modeSwitchSig2.checked = false;
          modeSwitchSig2.dispatchEvent(new Event('change'));
        }
        break;

      case 'p':
        e.preventDefault();
        const modeSwitchPrem2 = document.getElementById("mode-switch");
        if (modeSwitchPrem2 && !modeSwitchPrem2.checked) {
          modeSwitchPrem2.checked = true;
          modeSwitchPrem2.dispatchEvent(new Event('change'));
        }
        break;

      case 'a':
        e.preventDefault();
        document.getElementById("search-input").value = "";
        document.getElementById("action-filter").value = "all";
        localStorage.setItem('caseFilter', 'all');
        clearTimeout(searchTimeout);
        applySearch("");
        showToast('Filters and search cleared');
        break;

      case 's':
        e.preventDefault();
        clearSnoozedCases();
        break;
    }
  }
  if (e.key === 'Escape') {
    const ghoModal = document.getElementById('gho-modal');
    if (ghoModal && ghoModal.style.display === 'flex') {
      ghoModal.style.display = 'none';
    }
    const ghoAlertModal = document.getElementById('gho-alert-modal');
    if (ghoAlertModal) {
      ghoAlertModal.remove();
    }
  }
});
// Manual Track Case modal logic
async function showManualTrackModal() {
  try {
    const modal = document.getElementById('manual-track-modal');
    const input = document.getElementById('manual-track-input');
    const fetchBtn = document.getElementById('manual-track-fetch');
    const loading = document.getElementById('manual-track-loading');
    const details = document.getElementById('manual-track-details');
    const confirmBtn = document.getElementById('manual-track-confirm');
    const cancelBtn = document.getElementById('manual-track-cancel');
    const closeX = document.getElementById('manual-track-close');

    if (!modal || !input || !fetchBtn || !loading || !details || !confirmBtn || !cancelBtn || !closeX) return;

    const resetView = () => {
      loading.style.display = 'none';
      details.style.display = 'none';
      details.innerHTML = '';
      confirmBtn.disabled = true;
    };

    resetView();
    input.value = '';
    modal.style.display = 'flex';
    setTimeout(() => modal.classList.add('modal-show'), 0);
    setTimeout(() => input.focus(), 80);

    const closeModal = () => { modal.classList.remove('modal-show'); setTimeout(() => modal.style.display = 'none', 150); };
    cancelBtn.onclick = closeModal; closeX.onclick = closeModal; modal.onclick = (e) => { if (e.target === modal) closeModal(); };

    const doFetch = async () => {
      const num = (input.value || '').trim();
      if (!num) { showToast('Enter a Case Number'); return; }
      if (!window.jsforce || !SESSION_ID) { showToast('Session not ready'); return; }
      try {
        loading.style.display = 'block';
        details.style.display = 'none';
        const conn = new jsforce.Connection({ serverUrl: 'https://orgcs.my.salesforce.com', sessionId: SESSION_ID, version: '64.0' });
        const q = `SELECT Id, CaseNumber, Severity_Level__c, Subject, CaseRoutingTaxonomy__r.Name, Owner.Name, LastModifiedDate FROM Case WHERE CaseNumber='${num}' LIMIT 1`;
        const res = await conn.query(q);
        loading.style.display = 'none';
        if (!res.records || !res.records.length) { details.style.display = 'block'; details.innerHTML = '<div style="color:#b91c1c;">Case not found</div>'; confirmBtn.disabled = true; return; }
        const c = res.records[0];
        const cloud = (c.CaseRoutingTaxonomy__r && c.CaseRoutingTaxonomy__r.Name) ? c.CaseRoutingTaxonomy__r.Name.split('-')[0] : '';
        // Determine assignedTo for tracking:
        // If Routing_Status__c changed to 'Transferred', use Owner->NewValue from nearest subsequent Owner change
        let assignedToForConfirm = 'QB';
        try {
          const hres = await conn.query(`SELECT Field, NewValue, CreatedDate FROM CaseHistory WHERE CaseId='${c.Id}' ORDER BY CreatedDate ASC LIMIT 100`);
          const recs = (hres && hres.records) ? hres.records : [];
          let ownerVal = '';
          for (let i = 0; i < recs.length; i++) {
            const h = recs[i];
            const newValLower = String(h.NewValue || '').toLowerCase();
            if (h.Field === 'Routing_Status__c' && (newValLower.includes('transferred') || newValLower.includes('manually assigned'))) {
              for (let j = i + 1; j < recs.length; j++) {
                const oh = recs[j];
                if (oh.Field === 'Owner' && oh.NewValue) { ownerVal = String(oh.NewValue); break; }
              }
              break;
            }
          }
          if (ownerVal) {
            const isUserId = /^005[\w]{12,15}$/.test(ownerVal);
            const isGroupId = /^00G[\w]{12,15}$/.test(ownerVal);
            if (isUserId) {
              try {
                const u = await conn.query(`SELECT Name FROM User WHERE Id='${ownerVal}' LIMIT 1`);
                assignedToForConfirm = (u && u.records && u.records.length) ? (u.records[0].Name || ownerVal) : ownerVal;
              } catch { assignedToForConfirm = ownerVal; }
            } else if (isGroupId) {
              try {
                const g = await conn.query(`SELECT Name FROM Group WHERE Id='${ownerVal}' LIMIT 1`);
                assignedToForConfirm = (g && g.records && g.records.length) ? (g.records[0].Name || ownerVal) : ownerVal;
              } catch { assignedToForConfirm = ownerVal; }
            } else {
              assignedToForConfirm = ownerVal;
            }
          } else if (recs.some(h => h.Field === 'Routing_Status__c' && (String(h.NewValue || '').toLowerCase().includes('transferred') || String(h.NewValue || '').toLowerCase().includes('manually assigned')))) {
            assignedToForConfirm = (c.Owner && c.Owner.Name) ? c.Owner.Name : 'Case Owner';
          }
        } catch { }

        // Determine ActionType for tracking: GHO if routing logs show reason 'GHO'
        let actionTypeForConfirm = 'New Case';
        try {
          const r = await conn.query(`SELECT Transfer_Reason__c FROM Case_Routing_Log__c WHERE Case__c='${c.Id}' AND CreatedById='${userId}' ORDER BY CreatedDate DESC LIMIT 20`);
          const recs = (r && r.records) ? r.records : [];
          if (recs.some(rr => String(rr.Transfer_Reason__c || '').toUpperCase() === 'GHO')) actionTypeForConfirm = 'GHO';
        } catch { }
        details.innerHTML = `
          <div style="display:grid; grid-template-columns: 1fr 1fr; gap:8px;">
            <div><strong>Case</strong>: ${c.CaseNumber}</div>
            <div><strong>Severity</strong>: ${c.Severity_Level__c || '-'}</div>
            <div style="grid-column: 1 / -1;"><strong>Subject</strong>: ${String(c.Subject || '').replace(/</g, '&lt;')}</div>
            <div><strong>Owner</strong>: ${c.Owner && c.Owner.Name ? c.Owner.Name : '-'}</div>
            <div><strong>Cloud</strong>: ${cloud || '-'}</div>
          </div>
          <div id="manual-track-activity" style="margin-top:12px; padding-top:10px; border-top:1px solid #e5e7eb;">
            <div style="font-weight:700; color:#0f172a; margin-bottom:6px;">Your recent actions</div>
            <div id="manual-track-activity-body" style="display:flex; flex-direction:column; gap:6px;">Loading...</div>
          </div>`;
        details.style.display = 'block';
        confirmBtn.disabled = false;

        confirmBtn.onclick = async () => {
          try {
            trackActionAndCount(new Date(), c.CaseNumber, c.Severity_Level__c || '', actionTypeForConfirm, cloud || '', currentMode, currentUserName, assignedToForConfirm);
            showToast(`Tracked New Case for ${c.CaseNumber}`);
            closeModal();
          } catch (err) {
            console.warn('Manual confirm track failed:', err);
            showToast('Failed to track');
          }
        };

        // Load CaseFeed and CaseHistory entries you created (GHOTriage/QBMention and Owner/Status changes)
        (async () => {
          try {
            let userId = null;
            try {
              const ures = await conn.query(`SELECT Id FROM User WHERE Name = '${currentUserName}' AND IsActive = True AND Username LIKE '%orgcs.com'`);
              if (ures && ures.records && ures.records.length) userId = ures.records[0].Id;
            } catch { }
            const bodyEl = document.getElementById('manual-track-activity-body');
            if (!bodyEl) return;
            if (!userId) { bodyEl.textContent = 'Could not resolve your user id.'; return; }

            const [feedRes, histRes] = await Promise.all([
              conn.query(`SELECT Body, CreatedDate FROM CaseFeed WHERE Visibility='InternalUsers' AND ParentId='${c.Id}' AND Type='TextPost' AND CreatedById='${userId}' ORDER BY CreatedDate DESC LIMIT 20`),
              conn.query(`SELECT Field, NewValue, CreatedDate FROM CaseHistory WHERE CaseId='${c.Id}' AND CreatedById='${userId}' AND (Field='Routing_Status__c' OR Field='Owner') ORDER BY CreatedDate DESC LIMIT 20`)
            ]);
            const items = [];
            (feedRes.records || []).forEach(cm => {
              const body = String(cm.Body || ''); const low = body.toLowerCase();
              if (low.includes('#ghotriage') || low.includes('#sigqbmention')) items.push({ t: 'comment', when: cm.CreatedDate, text: body });
            });
            (histRes.records || []).forEach(h => {
              const f = h.Field; const val = (typeof h.NewValue === 'string') ? h.NewValue : (h.NewValue && h.NewValue.name) ? h.NewValue.name : '';
              items.push({ t: 'history', when: h.CreatedDate, text: `${f}: ${val}` });
            });
            items.sort((a, b) => new Date(b.when) - new Date(a.when));
            if (!items.length) { bodyEl.textContent = 'No recent comments or history by you.'; return; }
            bodyEl.innerHTML = items.slice(0, 10).map(it => {
              const when = new Date(it.when).toLocaleString();
              const badge = it.t === 'comment' ? '<span class="badge-soft badge-soft--info" style="margin-right:6px;">Comment</span>' : '<span class="badge-soft" style="background:#eef2ff; color:#3730a3; border:1px solid #e0e7ff; margin-right:6px;">History</span>';
              const safe = String(it.text || '').replace(/</g, '&lt;');
              return `<div style="font-size:12px; color:#334155;">${badge}<span style="color:#0f172a; font-weight:600;">${when}</span> — ${safe}</div>`;
            }).join('');
          } catch (e) {
            const bodyEl = document.getElementById('manual-track-activity-body');
            if (bodyEl) bodyEl.textContent = 'Failed to load your recent actions.';
          }
        })();
      } catch (e) {
        loading.style.display = 'none';
        details.style.display = 'block';
        details.innerHTML = `<div style="color:#b91c1c;">${String(e.message || e).replace(/</g, '&lt;')}</div>`;
        confirmBtn.disabled = true;
      }
    };

    fetchBtn.onclick = doFetch;
    input.onkeypress = (e) => { if (e.key === 'Enter') doFetch(); };
  } catch (e) { console.warn('showManualTrackModal failed', e); }
}

document.addEventListener('keydown', function (e) {
  if (e.key === '?' && !e.metaKey && !e.ctrlKey && !e.altKey) {
    e.preventDefault();
    showKeyboardShortcutsHelp();
  }
});

// Function to show keyboard shortcuts help
function showKeyboardShortcutsHelp() {
  const helpModal = `
    <div id="shortcuts-help-modal" class="modal-overlay" style="display: flex; z-index: 1002;">
      <div class="modal-content" style="max-width: 500px;">
        <div class="modal-header" style="background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%); color: white;">
          <h3 style="color: white; margin: 0;">⌨️ Keyboard Shortcuts</h3>
          <span class="modal-close" id="shortcuts-help-close" style="color: white; cursor: pointer; font-size: 24px;">&times;</span>
        </div>
        <div class="modal-body" style="padding: 24px;">
          <div style="display: grid; gap: 16px;">
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px 0; border-bottom: 1px solid #e5e7eb;">
              <span style="font-weight: 600; color: #111827; font-size: 15px;">Search Cases</span>
              <kbd style="background: #f3f4f6; padding: 4px 8px; border-radius: 4px; font-family: monospace; font-size: 12px; color: #374151; border: 1px solid #d1d5db;">⌘ + F</kbd>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px 0; border-bottom: 1px solid #e5e7eb;">
              <span style="font-weight: 600; color: #111827; font-size: 15px;">Open GHO Status</span>
              <kbd style="background: #f3f4f6; padding: 4px 8px; border-radius: 4px; font-family: monospace; font-size: 12px; color: #374151; border: 1px solid #d1d5db;">⌘ + G</kbd>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px 0; border-bottom: 1px solid #e5e7eb;">
              <span style="font-weight: 600; color: #111827; font-size: 15px;">Open GHO Mapping</span>
              <kbd style="background: #f3f4f6; padding: 4px 8px; border-radius: 4px; font-family: monospace; font-size: 12px; color: #374151; border: 1px solid #d1d5db;">⌘ + M</kbd>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px 0; border-bottom: 1px solid #e5e7eb;">
              <span style="font-weight: 600; color: #111827; font-size: 15px;">Refresh Data</span>
              <kbd style="background: #f3f4f6; padding: 4px 8px; border-radius: 4px; font-family: monospace; font-size: 12px; color: #374151; border: 1px solid #d1d5db;">⌘ + R</kbd>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px 0; border-bottom: 1px solid #e5e7eb;">
              <span style="font-weight: 600; color: #111827; font-size: 15px;">Refresh GHO Data</span>
              <kbd style="background: #f3f4f6; padding: 4px 8px; border-radius: 4px; font-family: monospace; font-size: 12px; color: #374151; border: 1px solid #d1d5db;">⌘ + Shift + R</kbd>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px 0; border-bottom: 1px solid #e5e7eb;">
              <span style="font-weight: 600; color: #111827; font-size: 15px;">Signature Mode</span>
              <kbd style="background: #f3f4f6; padding: 4px 8px; border-radius: 4px; font-family: monospace; font-size: 12px; color: #374151; border: 1px solid #d1d5db;">⌘ + I</kbd>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px 0; border-bottom: 1px solid #e5e7eb;">
              <span style="font-weight: 600; color: #111827; font-size: 15px;">Premier Mode</span>
              <kbd style="background: #f3f4f6; padding: 4px 8px; border-radius: 4px; font-family: monospace; font-size: 12px; color: #374151; border: 1px solid #d1d5db;">⌘ + P</kbd>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px 0; border-bottom: 1px solid #e5e7eb;">
              <span style="font-weight: 600; color: #111827; font-size: 15px;">Clear All Filters</span>
              <kbd style="background: #f3f4f6; padding: 4px 8px; border-radius: 4px; font-family: monospace; font-size: 12px; color: #374151; border: 1px solid #d1d5db;">⌘ + A</kbd>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px 0; border-bottom: 1px solid #e5e7eb;">
              <span style="font-weight: 600; color: #111827; font-size: 15px;">Clear Snoozed Cases</span>
              <kbd style="background: #f3f4f6; padding: 4px 8px; border-radius: 4px; font-family: monospace; font-size: 12px; color: #374151; border: 1px solid #d1d5db;">⌘ + S</kbd>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px 0; border-bottom: 1px solid #e5e7eb;">
              <span style="font-weight: 600; color: #111827; font-size: 15px;">Track Case (toggle / submit)</span>
              <kbd style="background: #f3f4f6; padding: 4px 8px; border-radius: 4px; font-family: monospace; font-size: 12px; color: #374151; border: 1px solid #d1d5db;">⌘ + L</kbd>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px 0; border-bottom: 1px solid #e5e7eb;">
              <span style="font-weight: 600; color: #111827; font-size: 15px;">Close Modals</span>
              <kbd style="background: #f3f4f6; padding: 4px 8px; border-radius: 4px; font-family: monospace; font-size: 12px; color: #374151; border: 1px solid #d1d5db;">ESC</kbd>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px 0;">
              <span style="font-weight: 600; color: #111827; font-size: 15px;">Show This Help</span>
              <kbd style="background: #f3f4f6; padding: 4px 8px; border-radius: 4px; font-family: monospace; font-size: 12px; color: #374151; border: 1px solid #d1d5db;">?</kbd>
            </div>
          </div>
          
          <div style="margin-top: 20px; padding: 12px; background: #f0f9ff; border: 1px solid #bfdbfe; border-radius: 6px;">
            <p style="margin: 0; color: #1e40af; font-size: 14px;">
              <strong>Tip:</strong> Use these shortcuts to navigate the app more efficiently!
            </p>
          </div>
        </div>
      </div>
    </div>
  `;

  const existingModal = document.getElementById('shortcuts-help-modal');
  if (existingModal) {
    existingModal.remove();
  }

  document.body.insertAdjacentHTML('beforeend', helpModal);
  const helpEl = document.getElementById('shortcuts-help-modal');
  if (helpEl) {
    requestAnimationFrame(() => helpEl.classList.add('modal-show'));
  }
  document.getElementById('shortcuts-help-close').addEventListener('click', () => {
    document.getElementById('shortcuts-help-modal').remove();
  });
  document.getElementById('shortcuts-help-modal').addEventListener('click', function (e) {
    if (e.target === this) {
      this.remove();
    }
  });
}

// ================================================
// Section: 11) DOMContentLoaded Initialization & UI Wiring
// ================================================
document.addEventListener("DOMContentLoaded", function () {
  ensureSingleTab();
  updateWeekendModeIndicator();
  setInterval(updateWeekendModeIndicator, 60000);
  const firstContainer = document.getElementById('parentSigSev2');
  if (firstContainer) {
    firstContainer.classList.add('is-loading');
  }
  getSessionIds();



  initMouseActivityTracking();
  setInterval(checkGHOAlert, 60000);
  setTimeout(checkGHOAlert, 5000);

  // Initialize default GHO mappings
  // initializeDefaultMappings(); (removed)



  const savedFilter = localStorage.getItem('caseFilter');
  if (savedFilter) {
    document.getElementById('action-filter').value = savedFilter;
  }

  // Search input validation (automatic)
  function validateSearchInput() {
    const searchInput = document.getElementById("search-input");
    const inputValue = searchInput.value.trim();
    // No button to enable/disable; keep for compatibility
    void inputValue;
  }

  // Debounced search function for real-time filtering
  function performRealTimeSearch() {
    clearTimeout(searchTimeout);
    searchTimeout = setTimeout(() => {
      const searchValue = document.getElementById("search-input").value.trim();
      if (searchValue.length >= 4) {
        applySearch(searchValue);
      } else if (searchValue.length === 0) {
        // Reset when cleared
        applySearch("");
      }
    }, 300); // 300ms delay to avoid too many calls while typing
  }

  // Add input event listener to validate search input in real-time
  document.getElementById("search-input").addEventListener("input", function () {
    validateSearchInput();
    performRealTimeSearch();
  });

  // Add keyup event listener for better responsiveness (handles backspace, delete, etc.)
  document.getElementById("search-input").addEventListener("keyup", function () {
    validateSearchInput();
    performRealTimeSearch();
  });

  // Add Enter key support for search
  document.getElementById("search-input").addEventListener("keydown", function (event) {
    if (event.key === "Enter") {
      const searchValue = this.value.trim();
      if (searchValue.length >= 4) {
        event.preventDefault();
        clearTimeout(searchTimeout);
        applySearch(searchValue);
      }
    }
  });

  document.getElementById("action-filter").addEventListener("change", function () {
    let filterValue = this.value;
    localStorage.setItem('caseFilter', filterValue);
    applyFilter(filterValue);
  });

  document.getElementById("mode-switch").addEventListener("change", function () {
    const headerTitle = document.querySelector(".header-title");
    const pill = document.getElementById('mode-pill');
    if (this.checked) {
      currentMode = 'premier';
      headerTitle.textContent = 'Zerolag Tool - Premier Mode';
      if (pill) { pill.textContent = 'Premier'; pill.classList.add('flash'); setTimeout(() => pill.classList.remove('flash'), 600); }
    } else {
      currentMode = 'signature';
      headerTitle.textContent = 'Zerolag Tool - Signature Mode';
      if (pill) { pill.textContent = 'Signature'; pill.classList.add('flash'); setTimeout(() => pill.classList.remove('flash'), 600); }
    }
    localStorage.setItem('caseTriageMode', currentMode);

    try {
      const ghoHeader = document.querySelector('#gho-modal .modal-header h3');
      if (ghoHeader) {
        const modeLabel = currentMode === 'premier' ? 'Premier' : 'Signature';
        let badge = ghoHeader.querySelector('.gho-mode-badge');
        if (badge) {
          badge.textContent = modeLabel + ' Mode';
          badge.style.background = currentMode === 'premier' ? '#065f46' : '#1e3a8a';
        }
      }
    } catch (e) { /* silent */ }

    // Update GHO button visibility based on mode
    updateGHOButtonVisibility();
    updateCICButtonVisibility();

    // Update GHO cache status display if modal is open
    if (document.getElementById('gho-modal') && document.getElementById('gho-modal').style.display === 'flex') {
      updateGHOCacheStatus();
    }

    // Clear search and filter when changing modes
    document.getElementById("search-input").value = "";
    document.getElementById("action-filter").value = "all";
    localStorage.setItem('caseFilter', 'all');
    clearTimeout(searchTimeout);

    // Clear existing content and show loading state
    const listContainer = document.getElementById("parentSigSev2");
    listContainer.classList.add('is-loading');
    listContainer.classList.remove('content-enter');
    listContainer.innerHTML = `
      <div class="loading-message">
        <h4 style="color: #374151; font-size: 20px; margin-bottom: 16px; font-weight: 600;">Switching to ${currentMode.charAt(0).toUpperCase() + currentMode.slice(1)} Mode...</h4>
        <p style="color: #6b7280; font-size: 16px; margin: 0;">Loading cases...</p>
      </div>
    `;


    // Fetch new data
    getSessionIds();
  });

  const modeSwitch = document.getElementById("mode-switch");
  const headerTitle = document.querySelector(".header-title");
  const headerContainer = document.querySelector('.header-container');
  if (headerContainer) headerContainer.classList.add('animate-in');
  const pill = document.getElementById('mode-pill');
  if (currentMode === 'premier') {
    modeSwitch.checked = true;
    headerTitle.textContent = 'Zerolag Tool - Premier Mode';
    if (pill) pill.textContent = 'Premier';
  } else {
    headerTitle.textContent = 'Zerolag Tool - Signature Mode';
    if (pill) pill.textContent = 'Signature';
  }

  // Set initial GHO button visibility
  updateGHOButtonVisibility();
  // Set initial CIC button visibility
  updateCICButtonVisibility();
  // Keep CIC button visibility in sync with weekend window
  setInterval(updateCICButtonVisibility, 60000);

  // Set initial Clear Snoozed button visibility
  updateClearSnoozedButtonVisibility();

  // Help Button Event Listener
  document.getElementById("shortcuts-help-btn").addEventListener("click", function () {
    showKeyboardShortcutsHelp();
  });

  // GHO Button Event Listener
  document.getElementById("check-gho-button").addEventListener("click", function () {
    checkGHOStatus();
  });

  // GHO Refresh Button Event Listener
  document.getElementById("gho-refresh-btn").addEventListener("click", function () {
    forceRefreshGHOData();
  });

  // GHO Refresh Button Hover Effects
  document.getElementById("gho-refresh-btn").addEventListener("mouseover", function () {
    this.style.transform = 'scale(1.05)';
  });

  document.getElementById("gho-refresh-btn").addEventListener("mouseout", function () {
    this.style.transform = 'scale(1)';
  });

  // GHO Mapping Button Event Listener
  document.getElementById("gho-mapping-btn").addEventListener("click", function () {
    showGHOMappingModal();
  });

  // More menu toggle (Signature-only option visibility)
  const moreBtn = document.getElementById('more-menu-btn');
  const moreDropdown = document.getElementById('more-menu-dropdown');
  const openVoiceModalBtn = document.getElementById('open-voicecall-modal');
  if (moreBtn && moreDropdown && openVoiceModalBtn) {
    const updateMoreMenuVisibility = () => {
      // Only show item in Signature mode
      openVoiceModalBtn.style.display = (currentMode === 'signature') ? 'flex' : 'none';
    };
    updateMoreMenuVisibility();
    moreBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      const isOpen = moreDropdown.style.display === 'block';
      if (isOpen) {
        moreDropdown.style.display = 'none';
        return;
      }
      // Ensure dropdown is at body level to escape any stacking context
      try {
        if (moreDropdown.parentElement !== document.body) {
          document.body.appendChild(moreDropdown);
        }
      } catch { }
      // Position as fixed overlay near the button to avoid header clipping and stacking issues
      const rect = moreBtn.getBoundingClientRect();
      moreDropdown.style.position = 'fixed';
      moreDropdown.style.top = `${Math.round(rect.bottom + 8)}px`;
      // Align right edge to button's right, with viewport padding
      const rightPx = Math.max(8, Math.round(window.innerWidth - rect.right));
      moreDropdown.style.right = `${rightPx}px`;
      moreDropdown.style.left = 'auto';
      moreDropdown.style.minWidth = '220px';
      moreDropdown.style.zIndex = '5000';
      moreDropdown.style.display = 'block';
      // Hide on scroll/resize
      const hideOnScrollOrResize = () => { moreDropdown.style.display = 'none'; window.removeEventListener('scroll', hideOnScrollOrResize, true); window.removeEventListener('resize', hideOnScrollOrResize, true); };
      window.addEventListener('scroll', hideOnScrollOrResize, true);
      window.addEventListener('resize', hideOnScrollOrResize, true);
    });
    document.addEventListener('click', (e) => {
      if (!moreDropdown.contains(e.target) && e.target !== moreBtn) {
        moreDropdown.style.display = 'none';
      }
    });
    document.getElementById('mode-switch').addEventListener('change', updateMoreMenuVisibility);
    openVoiceModalBtn.addEventListener('click', () => {
      moreDropdown.style.display = 'none';
      showVoiceCallModal();
    });

    // Manual Track modal open
    const openManualBtn = document.getElementById('open-manual-track-modal');
    if (openManualBtn) {
      openManualBtn.addEventListener('click', () => {
        moreDropdown.style.display = 'none';
        showManualTrackModal();
      });
    }

    // Populate user header initially
    updateMoreMenuUserHeader();
  }

  // Weekend Roster Button Event Listener
  document.getElementById("cic-button").addEventListener("click", function () {
    if (!DEV_FORCE_SHOW_WEEKEND_ROSTER && !isCurrentlyWeekend()) {
      showToast('Weekend Roster is for weekends only');
      return;
    }
    showCICManagers();
  });

  // GHO Modal Event Listeners
  document.getElementById("gho-modal-close").addEventListener("click", function () {
    const m = document.getElementById('gho-modal');
    m.classList.remove('modal-show');
    setTimeout(() => {
      m.style.display = 'none';
      // Clear cache status update interval
      if (m.dataset.cacheStatusInterval) {
        clearInterval(parseInt(m.dataset.cacheStatusInterval));
        delete m.dataset.cacheStatusInterval;
      }
    }, 150);
  });

  // GHO Filter Event Listener
  const filterDropdown = document.getElementById("gho-taxonomy-filter");
  if (filterDropdown) {
    console.log('GHO filter dropdown found, adding event listener');
    filterDropdown.addEventListener("change", function () {
      const filterValue = this.value;
      console.log('GHO filter changed to:', filterValue);
      console.log('GHO records available:', ghoRecordsGlobal.length);
      console.log('GHO connection available:', !!ghoConnectionGlobal);
      console.log('Current mode:', currentMode);

      if (ghoRecordsGlobal.length > 0 && ghoConnectionGlobal) {
        // Reset pagination state on filter change
        window.__ghoListState = null;
        const modeCache = window.__ghoCache && window.__ghoCache[currentMode];
        console.log('Mode cache available:', !!modeCache);
        console.log('About to call renderFilteredGHOCases with:', {
          recordsCount: ghoRecordsGlobal.length,
          filterValue: filterValue,
          hasCache: !!modeCache
        });

        // Always use cached data if available, regardless of filter
        if (modeCache && modeCache.records) {
          console.log('Using cached data for filtering');
          renderFilteredGHOCases(modeCache.records, modeCache.conn, filterValue, {
            triageCaseIds: modeCache.triageCaseIds,
            mondayCaseIds: modeCache.mondayCaseIds,
            noWocTemplateCaseIds: modeCache.noWocTemplateCaseIds,
            userMap: modeCache.userMap,
            fromCache: true
          });
        } else {
          renderFilteredGHOCases(ghoRecordsGlobal, ghoConnectionGlobal, filterValue, modeCache ? { triageCaseIds: modeCache.triageCaseIds } : undefined);
        }
      } else {
        console.warn('Cannot filter: missing GHO records or connection');
        console.log('ghoRecordsGlobal:', ghoRecordsGlobal);
        console.log('ghoConnectionGlobal:', ghoConnectionGlobal);
      }
    });
  } else {
    console.warn('GHO filter dropdown not found');
  }



  // Voice Calls modal wiring
  const voiceModal = document.getElementById('voicecall-modal');
  const voiceCloseX = document.getElementById('voicecall-modal-close');
  const voiceCloseBtn = document.getElementById('voicecall-close-btn');
  const voiceTrackBtn = document.getElementById('voicecall-track-btn');
  const voiceSelectAll = document.getElementById('voicecall-select-all');
  if (voiceModal && voiceCloseX && voiceCloseBtn && voiceTrackBtn && voiceSelectAll) {
    const closeVoiceModal = () => {
      voiceModal.classList.remove('modal-show');
      setTimeout(() => voiceModal.style.display = 'none', 150);
    };
    voiceCloseX.addEventListener('click', closeVoiceModal);
    voiceCloseBtn.addEventListener('click', closeVoiceModal);
    voiceModal.addEventListener('click', (e) => { if (e.target === voiceModal) closeVoiceModal(); });
    document.addEventListener('keydown', (e) => { if (voiceModal.style.display === 'flex' && e.key === 'Escape') closeVoiceModal(); });
  }

  document.getElementById("gho-modal").addEventListener("click", function (e) {
    if (e.target === this) {
      this.classList.remove('modal-show');
      setTimeout(() => {
        this.style.display = 'none';
        // Clear cache status update interval
        if (this.dataset.cacheStatusInterval) {
          clearInterval(parseInt(this.dataset.cacheStatusInterval));
          delete this.dataset.cacheStatusInterval;
        }
      }, 150);
    }
  });

  // GHO Mapping Modal Event Listeners
  document.getElementById("gho-mapping-modal-close").addEventListener("click", function () {
    const m = document.getElementById('gho-mapping-modal');
    m.classList.remove('modal-show');
    setTimeout(() => { m.style.display = 'none'; }, 150);
  });

  document.getElementById("gho-mapping-modal").addEventListener("click", function (e) {
    if (e.target === this) {
      this.classList.remove('modal-show');
      setTimeout(() => { this.style.display = 'none'; }, 150);
    }
  });

  document.getElementById("add-mapping-btn").addEventListener("click", function () {
    addNewMapping();
  });

  document.getElementById("save-mappings-btn").addEventListener("click", function () {
    showToast('Mappings are automatically saved when added or deleted', 'info');
  });

  document.getElementById("delete-all-mappings-btn").addEventListener("click", function () {
    deleteAllMappings();
  });

  // Add event listeners for mapping modal links
  document.addEventListener('click', function (e) {
    if (e.target && e.target.matches('a[href="#"]')) {
      e.preventDefault();
      const text = e.target.textContent.trim();
      if (text === 'View/Edit Mappings' || text === 'Configure Mappings') {
        showGHOMappingModal();
      }
    }

    // Handle mapping links with data-action attributes
    if (e.target && e.target.matches('.mapping-link[data-action="view-mappings"]')) {
      e.preventDefault();
      showGHOMappingModal();
    }

    // Handle delete mapping buttons
    if (e.target && e.target.matches('.delete-mapping-btn')) {
      const index = parseInt(e.target.dataset.index);
      if (!isNaN(index)) {
        deleteMapping(index);
      }
    }

    // Handle save mapping buttons
    if (e.target && e.target.matches('.save-mapping-btn')) {
      saveNewMapping(e.target);
    }

    // Handle cancel mapping buttons
    if (e.target && e.target.matches('.cancel-mapping-btn')) {
      cancelNewMapping(e.target);
    }
  });

  // GHO Toggle Button Event Listener
  document.getElementById("gho-cases-container").addEventListener("click", function (e) {
    // Find the button element, whether clicked directly or on a child element
    const button = e.target.closest('.gho-toggle-btn');
    if (button) {
      const caseUniqueId = button.dataset.caseId;
      toggleGhoTransfers(caseUniqueId);
    }

    // Handle GHO preview record clicks - auto-copy GHO template
    // (Handled via gho.js attachGhoPreviewTemplateCopy)
  });
  // Attach delegated GHO template copying (idempotent)
  attachGhoPreviewTemplateCopy();

  // Add hover effects for GHO toggle buttons using event delegation
  document.getElementById("gho-cases-container").addEventListener("mouseover", function (e) {
    const button = e.target.closest('.gho-toggle-btn');
    if (button) {
      button.style.background = 'linear-gradient(135deg, #e2e8f0 0%, #cbd5e1 100%)';
      button.style.transform = 'translateY(-1px)';
      button.style.boxShadow = '0 2px 4px rgba(0, 0, 0, 0.1)';
    }
  });

  document.getElementById("gho-cases-container").addEventListener("mouseout", function (e) {
    const button = e.target.closest('.gho-toggle-btn');
    if (button) {
      button.style.background = 'linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%)';
      button.style.transform = 'translateY(0)';
      button.style.boxShadow = '0 1px 2px rgba(0, 0, 0, 0.05)';
    }
  });
});

// Clear search input and reset filters
document.getElementById("clear-button").addEventListener("click", function () {
  document.getElementById("search-input").value = "";
  // search button removed

  document.getElementById("action-filter").value = "all";
  localStorage.setItem('caseFilter', 'all');

  clearTimeout(searchTimeout);

  applySearch("");
});

document.getElementById("refresh-button").addEventListener("click", function () {
  const refreshIcon = this.querySelector('.fa-refresh');
  refreshIcon.classList.add('fa-spin');
  setTimeout(() => {
    window.location.reload();
  }, 500);
});

document.getElementById("clear-snoozed-button").addEventListener("click", function () {
  clearSnoozedCases();
});

// Manual track by Case Number: removed inline UI + shortcut; moved to 3-dots modal

document.getElementById("parentSigSev2").addEventListener("click", function (e) {
  if (e.target && e.target.classList.contains("preview-record-btn")) {
    const button = e.target;
    const caseDiv = button.closest('.case-card');
    // clicking preview no longer counts toward actioned cases; actions are counted when trackAction is invoked

    const severityInfoItems = caseDiv.querySelectorAll('.case-info-item');
    let severityText = '';
    severityInfoItems.forEach(item => {
      if (item.textContent.includes('Level')) {
        severityText = item.textContent;
      }
    });
    const severity = severityText.includes('Level 1') ? '1' :
      severityText.includes('Level 2') ? '2' :
        severityText.includes('Level 3') ? '3' : '4';

    const isMVPCase = button.dataset.isMvp === 'true';
    let textToCopy = '';

    if (currentMode === 'premier') {
      textToCopy = `Hi\nNew SEV${severity} assigned to you & App is updated...!`;
    } else if (currentMode === 'signature' && isCurrentlyWeekend()) {
      textToCopy = getWeekendSignatureTemplate(severity);
    } else if (isMVPCase && (severity === '1' || severity === '2')) {
      textToCopy = `Hi\nKindly help with the assignment of new SEV${severity} MVP case, as it has not been assigned through OMNI. Thank you!\nFYI: @Susanna Catherine \n#SigQBmention`;
    } else if (isMVPCase) {
      textToCopy = `Hi\nKindly help with the assignment of new SEV${severity} MVP case, as it has not been assigned through OMNI and SLA is in warning status. Thank you!\nFYI: @Susanna Catherine \n#SigQBmention`;
    }
    else {
      textToCopy = `Hi\nKindly help with the assignment of new SEV${severity} case, as it has not been assigned through OMNI. \nThank you!\nFYI: @Susanna Catherine \n#SigQBmention`;
    }

    navigator.clipboard.writeText(textToCopy).then(function () {
      const toast = document.getElementById('toast');
      toast.textContent = 'Template copied to clipboard!';
      toast.style.display = 'block';
      setTimeout(function () {
        toast.style.display = 'none';
      }, 2000);
    });
  }

  if (e.target.classList.contains("snooze-btn")) {
    const button = e.target;
    const caseId = button.dataset.caseId;
    const caseDiv = button.closest('.case-card');
    const snoozeTimeSelect = caseDiv.querySelector('.snooze-time');
    const customInput = caseDiv.querySelector('.custom-snooze-input');

    let snoozeMinutes;

    if (snoozeTimeSelect.value === 'custom') {
      snoozeMinutes = parseInt(customInput.value);
      if (!snoozeMinutes || snoozeMinutes < 1 || snoozeMinutes > 1440) {
        showToast('Please enter a valid number of minutes (1-1440).', 'error');
        return;
      }
    } else {
      snoozeMinutes = parseInt(snoozeTimeSelect.value);
    }

    const snoozeUntil = new Date().getTime() + snoozeMinutes * 60 * 1000;
    localStorage.setItem('snooze_' + caseId, snoozeUntil);
    caseDiv.style.display = 'none';
    updateClearSnoozedButtonVisibility();
    showToast(`Case snoozed for ${snoozeMinutes} minute${snoozeMinutes === 1 ? '' : 's'}`);
  }
});

document.getElementById("parentSigSev2").addEventListener("change", function (e) {
  if (e.target.classList.contains("snooze-time")) {
    const select = e.target;
    const caseDiv = select.closest('.case-card');
    const customInput = caseDiv.querySelector('.custom-snooze-input');

    if (select.value === 'custom') {
      customInput.style.display = 'inline-block';
      customInput.focus();
    } else {
      customInput.style.display = 'none';
      customInput.value = '';
    }
  }
});


// Determine weekend target date based on shift rules
function getWeekendLookupDateForShift(shift) {
  const now = new Date();
  const istNow = new Date(now.toLocaleString('en-US', { timeZone: 'Asia/Kolkata' }));
  const day = istNow.getDay();
  const hours = istNow.getHours();
  const minutes = istNow.getMinutes();
  const totalMin = hours * 60 + minutes;
  const cutoff = 5 * 60 + 30;
  let target = new Date(istNow);

  const lastSat = new Date(istNow);
  lastSat.setDate(istNow.getDate() - ((day + 1) % 7));
  while (lastSat.getDay() !== 6) lastSat.setDate(lastSat.getDate() - 1);

  if (shift === 'AMER') {
    if (day === 0 && totalMin < cutoff) {
      target = lastSat;
    } else if (day === 6 && totalMin >= cutoff) {
      target = lastSat;
    } else if (day === 0 && totalMin >= cutoff) {
      target = istNow;
    } else if (day === 1 && totalMin < cutoff) {
      const sun = new Date(istNow);
      sun.setDate(sun.getDate() - 1);
      target = sun;
    } else {
      target = lastSat;
    }
  } else {
    if (day === 6 && totalMin >= cutoff) {
      target = istNow;
    } else if (day === 0) {
      target = istNow;
    } else if (day === 1 && totalMin < cutoff) {
      const sun = new Date(istNow);
      sun.setDate(sun.getDate() - 1);
      target = sun;
    } else {
      target = lastSat;
    }
  }
  const m = target.getMonth() + 1;
  const dNum = target.getDate();
  const d = dNum < 10 ? `0${dNum}` : `${dNum}`;
  const y = target.getFullYear();
  return `${m}/${d}/${y}`;
}

function getCurrentShiftIST() {
  const now = new Date();
  const ist = new Date(now.toLocaleString('en-US', { timeZone: 'Asia/Kolkata' }));
  const hours = ist.getHours();
  const minutes = ist.getMinutes();
  const totalMinutes = hours * 60 + minutes;
  const apacStart = 5 * 60 + 30;
  const emeaStart = 12 * 60 + 30;
  const apacEnd = 12 * 60 + 30;
  const emeaEnd = 20 * 60;
  if (totalMinutes >= apacStart && totalMinutes < apacEnd) return 'APAC';
  if (totalMinutes >= emeaStart && totalMinutes < emeaEnd) return 'EMEA';
  return 'AMER';
}

function googleSheetsGET(rangeA1, callback, onError) { return _sheetsGET(rangeA1, callback, onError, WEEKEND_ROSTER_SPREADSHEET_ID); }

async function getCellValueWithoutStrikethrough(sheetName, rowIndexZero, colLetter) { return _getCellNoStrike(sheetName, rowIndexZero, colLetter, WEEKEND_ROSTER_SPREADSHEET_ID); }

// ================================================
// Section: 13) CIC Managers (weekend roster modal)
// ================================================
function showCICManagers() {
  const shift = getCurrentShiftIST();
  const weekendDateStr = getWeekendLookupDateForShift(shift);
  const column = getCICColumnForShift(shift);

  // Include up to column X to support AMER Swarm Lead column
  const range = `'${SHEET_SERVICE_CLOUD}'!A:X`;

  const showLoading = () => {
    const loadingHtml = `
      <div id="cic-modal" class="modal-overlay" style="display:flex; z-index:1003;">
        <div class="modal-content" style="max-width: 720px; border-radius:16px; box-shadow: 0 10px 30px rgba(2,132,199,0.15); overflow:hidden;">
          <div class="modal-header" style="background: linear-gradient(135deg,#0ea5e9 0%, #0284c7 100%); color: white; display:flex; align-items:center; justify-content:space-between;">
            <div style="display:flex; align-items:center; gap:10px;">
              <h3 style="color:white; margin:0;">Weekend Roster</h3>
              <span style="background: rgba(255,255,255,0.2); color:#fff; padding:4px 10px; border-radius:999px; font-size:11px; letter-spacing:.4px;">${shift} · ${weekendDateStr}</span>
            </div>
            <span class="modal-close" id="cic-modal-close" style="color:white; cursor:pointer; font-size:24px;">&times;</span>
          </div>
          <div class="modal-body" style="background: #f8fafc;">
            <div class="loading-message" style="animation: fadeUp 300ms ease; text-align:center; padding: 24px;">
              <div class="spinner"></div>
              <h4 style="margin:12px 0 6px; color:#0c4a6e;">Loading Weekend Roster...</h4>
              <div style="display:flex; justify-content:center; gap:8px; flex-wrap:wrap;">
                <span style="background:#e0f2fe; color:#075985; border:1px solid #bae6fd; padding:4px 10px; border-radius:999px; font-size:12px;">Shift: ${shift}</span>
                <span style="background:#e0f2fe; color:#075985; border:1px solid #bae6fd; padding:4px 10px; border-radius:999px; font-size:12px;">Date: ${weekendDateStr}</span>
                <span style="background:${currentMode === 'premier' ? '#d1fae5' : '#e0e7ff'}; color:${currentMode === 'premier' ? '#065f46' : '#3730a3'}; border:1px solid ${currentMode === 'premier' ? '#6ee7b7' : '#c7d2fe'}; padding:4px 10px; border-radius:999px; font-size:12px;">Mode: ${currentMode === 'premier' ? 'Premier' : 'Signature'}</span>
              </div>
            </div>
          </div>
          <div class="modal-footer" style="padding:12px 16px; border-top:1px solid #e5e7eb; background:#ffffff; color:#64748b; font-size:12px;">
            Tip: Click any names area to copy all names as a comma-separated list.
          </div>
        </div>
      </div>`;
    const existing = document.getElementById('cic-modal');
    if (existing) existing.remove();
    document.body.insertAdjacentHTML('beforeend', loadingHtml);
    const modalEl = document.getElementById('cic-modal');
    if (modalEl) requestAnimationFrame(() => modalEl.classList.add('modal-show'));
    document.getElementById('cic-modal-close').addEventListener('click', () => {
      const m = document.getElementById('cic-modal');
      if (m) {
        m.classList.remove('modal-show');
        setTimeout(() => { m.remove(); }, 150);
      }
    });
    document.getElementById('cic-modal').addEventListener('click', function (e) {
      if (e.target === this) {
        this.classList.remove('modal-show');
        setTimeout(() => { this.remove(); }, 150);
      }
    });
  };

  showLoading();
  // Initialize daily counters scope for Premier (per-shift with AMER Sunday 20:30 rule)
  try { initPremierCounters(weekendDateStr, shift); } catch { }

  // Use cached weekend roster data with 1-hour TTL
  getWeekendRosterData(SHEET_SERVICE_CLOUD, shift, weekendDateStr, currentMode, (resp) => {
    try {
      const rows = resp.values || [];
      let foundRowIdx = -1;
      const targetDatePadded = weekendDateStr;
      const dNumTarget = parseInt(targetDatePadded.split('/')[1], 10);
      const mNumTarget = parseInt(targetDatePadded.split('/')[0], 10);
      const yNumTarget = parseInt(targetDatePadded.split('/')[2], 10);
      const targetDateAlt = `${mNumTarget}/${dNumTarget}/${yNumTarget}`;
      for (let i = 0; i < rows.length; i++) {
        const aVal = rows[i][0];
        if (!aVal) continue;
        const aStr = String(aVal).trim();
        if (aStr === targetDatePadded || aStr === targetDateAlt) {
          foundRowIdx = i;
          break;
        }
        const dt = new Date(aStr);
        if (!isNaN(dt.getTime())) {
          const am = dt.getMonth() + 1;
          const ad = dt.getDate();
          const ay = dt.getFullYear();
          if (am === mNumTarget && ad === dNumTarget && ay === yNumTarget) {
            foundRowIdx = i;
            break;
          }
        }
      }

      let namesCell = '';
      if (foundRowIdx >= 0) {
        const colIndex = column.charCodeAt(0) - 'A'.charCodeAt(0);
        namesCell = (rows[foundRowIdx][colIndex] || '').toString();
      }
      const prettyTitle = `CIC Managers — ${shift}`;
      const bodyEl = document.querySelector('#cic-modal .modal-body');
      if (!bodyEl) return;

      if (!namesCell) {
        bodyEl.innerHTML = `<div class="no-cases-message" style="animation: fadeUp 260ms ease;">
          <h4 class="no-cases-title" style="color:#0f766e;">No names found</h4>
          <p class="no-cases-text">No Weekend Roster listed for ${shift} on ${weekendDateStr}.</p>
          <p class="mode-switch-hint">Verify the date in Column A matches M/DD/YYYY.</p>
        </div>`;
        return;
      }

      const namesComma = namesCell
        .replace(/[\r\n]+/g, ', ')     // new lines -> comma
        .replace(/\s*&\s*/g, ', ')     // ampersand -> comma
        .replace(/\s*,\s*/g, ', ')     // collapse spaces around commas
        .replace(/,\s*,+/g, ', ')       // remove duplicate commas
        .replace(/\s{2,}/g, ' ')        // collapse multiple spaces
        .replace(/,\s*$/, '')           // trim trailing comma
        .trim();

      bodyEl.innerHTML = `
        <div style="display:flex; flex-direction:column; gap:12px;">
          <div class="cic-card" style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; box-shadow: 0 1px 2px rgba(0,0,0,0.04); padding:16px; animation: fadeUp 260ms ease;">
            <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;">
              <div style="font-weight:700; color:#0f766e;">${prettyTitle}</div>
              <span style="font-size:12px; color:#64748b; background:#f1f5f9; border:1px solid #e2e8f0; padding:4px 8px; border-radius:999px;">Click to copy</span>
            </div>
            <div id="cic-names" title="Click to copy" style="padding:12px; border:1px dashed #94a3b8; border-radius:8px; background:#f8fafc; cursor:pointer; white-space:normal; word-break:break-word; overflow-wrap:anywhere;">
              ${formatNamesToMultiline(namesCell)}
            </div>
          </div>
      ${currentMode === 'premier' ? `
          <div class="premier-toggle" style="display:flex; align-items:center; gap:8px; margin-top:4px;">
            <label for="premier-view" style="font-size:12px; color:#475569;">Premier view:</label>
            <select id="premier-view" style="padding:6px 10px; border:1px solid #cbd5e1; border-radius:8px; background:#ffffff; color:#0f172a; font-size:13px;">
        <option value="sales">Sales</option>
        <option value="service">Service</option>
              <option value="industry">Industry</option>
              <option value="data">Data</option>
            </select>
            <div style="margin-left:auto; display:flex; gap:8px; align-items:center;">
              <button id="premier-edit" title="Edit names/emails" style="padding:6px 10px; border:1px solid #94a3b8; background:#f8fafc; color:#0f172a; border-radius:8px; font-size:12px; cursor:pointer;">✏️ Edit</button>
              <button id="premier-reset" title="Reset all counters" style="padding:6px 10px; border:1px solid #ef4444; background:#fef2f2; color:#b91c1c; border-radius:8px; font-size:12px; cursor:pointer;">Reset Counts</button>
            </div>
          </div>
          <div id="premier-sections" style="display:flex; flex-direction:column; gap:12px;"></div>
          ` : ''}
        </div>`;

      const cell = document.getElementById('cic-names');
      cell.addEventListener('click', () => {
        navigator.clipboard.writeText(namesComma).then(() => {
          showToast('Weekend Roster copied');
        }).catch(() => showToast('Copy failed'));
      });

      if (currentMode === 'signature') {
        // Insert Sales, Service, Industry sections for better visibility (if not already present)
        try {
          const wrapper = document.querySelector('#cic-modal .modal-body > div');
          if (wrapper && !document.getElementById('sig-sales-section')) {
            wrapper.insertAdjacentHTML('beforeend', `
              <div class="sig-section" id="sig-sales-section" style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; padding:12px;">
                <div class="sig-section-header" style="font-weight:700; color:#0f172a; margin-bottom:8px;">Sales</div>
                <div class="sig-section-content" id="sig-sales-section-content" style="display:flex; flex-direction:column; gap:12px;"></div>
              </div>
              <div class="sig-section" id="sig-service-section" style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; padding:12px; margin-top:12px;">
                <div class="sig-section-header" style="font-weight:700; color:#0f172a; margin-bottom:8px;">Service</div>
                <div class="sig-section-content" id="sig-service-section-content" style="display:flex; flex-direction:column; gap:12px;"></div>
              </div>
              <div class="sig-section" id="sig-industry-section" style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; padding:12px; margin-top:12px;">
                <div class="sig-section-header" style="font-weight:700; color:#0f172a; margin-bottom:8px;">Industry</div>
                <div class="sig-section-content" id="sig-industry-section-content" style="display:flex; flex-direction:column; gap:12px;"></div>
              </div>
              <div class="sig-section" id="sig-data-section" style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; padding:12px; margin-top:12px;">
                <div class="sig-section-header" style="font-weight:700; color:#0f172a; margin-bottom:8px;">Data</div>
                <div class="sig-section-content" id="sig-data-section-content" style="display:flex; flex-direction:column; gap:12px;"></div>
              </div>
            `);
          }
        } catch { }

        // Show the Signature Sales Swarm Lead first (Sales > Service > Industry ordering)
        const bodyElSales = document.getElementById('sig-sales-section-content');
        if (bodyElSales) {
          const salesRange = `'${SHEET_SALES_CLOUD}'!A:Z`;
          const salesSwarmCol = getSwarmLeadSalesColumnForShift(shift);
          // Use cached weekend roster data for sales
          getWeekendRosterData(SHEET_SALES_CLOUD, shift, weekendDateStr, currentMode, (rSales) => {
            try {
              const rowsSales = (rSales && rSales.values) ? rSales.values : [];
              let rowIdxSales = -1;
              const targetDatePadded = weekendDateStr;
              const dNumTarget = parseInt(targetDatePadded.split('/')[1], 10);
              const mNumTarget = parseInt(targetDatePadded.split('/')[0], 10);
              const yNumTarget = parseInt(targetDatePadded.split('/')[2], 10);
              const targetDateAlt = `${mNumTarget}/${dNumTarget}/${yNumTarget}`;
              for (let i = 0; i < rowsSales.length; i++) {
                const aVal = rowsSales[i][0];
                if (!aVal) continue;
                const aStr = String(aVal).trim();
                if (aStr === targetDatePadded || aStr === targetDateAlt) { rowIdxSales = i; break; }
                const dt = new Date(aStr);
                if (!isNaN(dt.getTime())) {
                  const am = dt.getMonth() + 1, ad = dt.getDate(), ay = dt.getFullYear();
                  if (am === mNumTarget && ad === dNumTarget && ay === yNumTarget) { rowIdxSales = i; break; }
                }
              }

              let salesSwarmCell = '';
              if (rowIdxSales >= 0) {
                const colIdxSales = salesSwarmCol.charCodeAt(0) - 'A'.charCodeAt(0);
                salesSwarmCell = (rowsSales[rowIdxSales][colIdxSales] || '').toString().trim();
              }

              const salesSwarmTitle = 'Signature Sales Swarm Lead - ' + shift;
              if (!salesSwarmCell) {
                bodyElSales.insertAdjacentHTML('beforeend', `
                  <div style="margin-top:12px;">
                    <div style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; box-shadow: 0 1px 2px rgba(0,0,0,0.04); padding:16px; animation: fadeUp 260ms ease;">
                      <div style="font-weight:700; color:#0f766e; margin-bottom:4px;">${salesSwarmTitle}</div>
                      <div style="font-size:14px; color:#64748b;">No Designated Name</div>
                    </div>
                  </div>
                `);
              } else {
                const salesSwarmCopy = normalizeNamesForCopy(salesSwarmCell);
                bodyElSales.insertAdjacentHTML('beforeend', `
                  <div style="margin-top:12px;">
                    <div style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; box-shadow: 0 1px 2px rgba(0,0,0,0.04); padding:16px; animation: fadeUp 260ms ease;">
                      <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;">
                        <div style="font-weight:700; color:#0f766e;">${salesSwarmTitle}</div>
                        <span style="font-size:12px; color:#64748b; background:#f1f5f9; border:1px solid #e2e8f0; padding:4px 8px; border-radius:999px;">Click to copy</span>
                      </div>
                      <div id="sig-sales-swarm-lead-names" title="Click to copy" style="padding:12px; border:1px dashed #94a3b8; border-radius:8px; background:#f8fafc; cursor:pointer; white-space:normal; word-break:break-word; overflow-wrap:anywhere;">${formatNamesToMultiline(salesSwarmCell)}</div>
                    </div>
                  </div>
                `);
                const swSalesEl = document.getElementById('sig-sales-swarm-lead-names');
                if (swSalesEl) {
                  swSalesEl.addEventListener('click', () => {
                    navigator.clipboard.writeText(salesSwarmCopy).then(() => showToast('Signature Sales Swarm Lead copied')).catch(() => showToast('Copy failed'));
                  });
                }
              }
            } catch (e) {
              console.error('Failed building Signature Sales Swarm Lead section:', e);
            }
          }, (err) => {
            console.error('Failed loading Sales Cloud sheet for Swarm Lead:', err);
          });
        }
        const swarmCol = getSwarmLeadServiceColumnForShift(shift);
        const swarmIdx = swarmCol.charCodeAt(0) - 'A'.charCodeAt(0);
        const swarmCellRaw = (foundRowIdx >= 0 && rows[foundRowIdx]) ? (rows[foundRowIdx][swarmIdx] || '') : '';
        const swarmCell = String(swarmCellRaw || '').trim();
        const swarmTitle = 'Signature Service Swarm Lead - ' + shift;
        const bodyEl2 = document.getElementById('sig-service-section-content');
        if (bodyEl2) {
          if (!swarmCell) {
            bodyEl2.insertAdjacentHTML('beforeend', `
              <div style="margin-top:12px;">
                <div style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; box-shadow: 0 1px 2px rgba(0,0,0,0.04); padding:16px; animation: fadeUp 260ms ease;">
                  <div style="font-weight:700; color:#0f766e; margin-bottom:4px;">${swarmTitle}</div>
                  <div style="font-size:14px; color:#64748b;">No Designated Name</div>
                </div>
              </div>
            `);
          } else {
            const swarmCopy = normalizeNamesForCopy(swarmCell);
            bodyEl2.insertAdjacentHTML('beforeend', `
              <div style="margin-top:12px;">
                <div style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; box-shadow: 0 1px 2px rgba(0,0,0,0.04); padding:16px; animation: fadeUp 260ms ease;">
                  <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;">
                    <div style="font-weight:700; color:#0f766e;">${swarmTitle}</div>
                    <span style="font-size:12px; color:#64748b; background:#f1f5f9; border:1px solid #e2e8f0; padding:4px 8px; border-radius:999px;">Click to copy</span>
                  </div>
                  <div id="sig-swarm-lead-names" title="Click to copy" style="padding:12px; border:1px dashed #94a3b8; border-radius:8px; background:#f8fafc; cursor:pointer; white-space:normal; word-break:break-word; overflow-wrap:anywhere;">
                    ${formatNamesToMultiline(swarmCell)}
                  </div>
                </div>
              </div>
            `);
            const swEl = document.getElementById('sig-swarm-lead-names');
            if (swEl) {
              swEl.addEventListener('click', () => {
                navigator.clipboard.writeText(swarmCopy).then(() => showToast('Signature Swarm Lead copied')).catch(() => showToast('Copy failed'));
              });
            }
          }
        }

        // Industry Swarm Lead (Signature) — same mapping as Premier Industry, read from Industry sheet
        const industryContainer = document.getElementById('sig-industry-section-content');
        if (industryContainer) {
          const industryRange = `'${SHEET_INDUSTRY_CLOUD}'!A:Z`;
          const industrySwarmCol = getPremierIndustrySwarmLeadColumn(shift);
          // Use cached weekend roster data for industry
          getWeekendRosterData(SHEET_INDUSTRY_CLOUD, shift, weekendDateStr, currentMode, (rInd) => {
            try {
              const rowsInd = (rInd && rInd.values) ? rInd.values : [];
              let rowIdxInd = -1;
              const targetDatePadded = weekendDateStr;
              const dNumTarget = parseInt(targetDatePadded.split('/')[1], 10);
              const mNumTarget = parseInt(targetDatePadded.split('/')[0], 10);
              const yNumTarget = parseInt(targetDatePadded.split('/')[2], 10);
              const targetDateAlt = `${mNumTarget}/${dNumTarget}/${yNumTarget}`;
              for (let i = 0; i < rowsInd.length; i++) {
                const aVal = rowsInd[i][0];
                if (!aVal) continue;
                const aStr = String(aVal).trim();
                if (aStr === targetDatePadded || aStr === targetDateAlt) { rowIdxInd = i; break; }
                const dt = new Date(aStr);
                if (!isNaN(dt.getTime())) {
                  const am = dt.getMonth() + 1, ad = dt.getDate(), ay = dt.getFullYear();
                  if (am === mNumTarget && ad === dNumTarget && ay === yNumTarget) { rowIdxInd = i; break; }
                }
              }
              let indSwarmCell = '';
              if (rowIdxInd >= 0) {
                const colIdxInd = industrySwarmCol.charCodeAt(0) - 'A'.charCodeAt(0);
                indSwarmCell = (rowsInd[rowIdxInd][colIdxInd] || '').toString().trim();
              }
              const indSwarmTitle = 'Signature Industry Swarm Lead - ' + shift;
              if (!indSwarmCell) {
                industryContainer.insertAdjacentHTML('beforeend', `
                  <div style="margin-top:12px;">
                    <div style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; box-shadow: 0 1px 2px rgba(0,0,0,0.04); padding:16px; animation: fadeUp 260ms ease;">
                      <div style="font-weight:700; color:#0f766e; margin-bottom:4px;">${indSwarmTitle}</div>
                      <div style="font-size:14px; color:#64748b;">No Designated Name</div>
                    </div>
                  </div>
                `);
              } else {
                const indSwarmCopy = normalizeNamesForCopy(indSwarmCell);
                industryContainer.insertAdjacentHTML('beforeend', `
                  <div style="margin-top:12px;">
                    <div style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; box-shadow: 0 1px 2px rgba(0,0,0,0.04); padding:16px; animation: fadeUp 260ms ease;">
                      <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;">
                        <div style="font-weight:700; color:#0f766e;">${indSwarmTitle}</div>
                        <span style="font-size:12px; color:#64748b; background:#f1f5f9; border:1px solid #e2e8f0; padding:4px 8px; border-radius:999px;">Click to copy</span>
                      </div>
                      <div id="sig-industry-swarm-lead-names" title="Click to copy" style="padding:12px; border:1px dashed #94a3b8; border-radius:8px; background:#f8fafc; cursor:pointer; white-space:normal; word-break:break-word; overflow-wrap:anywhere;">${formatNamesToMultiline(indSwarmCell)}</div>
                    </div>
                  </div>
                `);
                const swIndEl = document.getElementById('sig-industry-swarm-lead-names');
                if (swIndEl) {
                  swIndEl.addEventListener('click', () => {
                    navigator.clipboard.writeText(indSwarmCopy).then(() => showToast('Signature Industry Swarm Lead copied')).catch(() => showToast('Copy failed'));
                  });
                }
              }
            } catch (e) {
              console.error('Failed building Signature Industry Swarm Lead section:', e);
            }
          }, (err) => {
            console.error('Failed loading Industry Cloud sheet for Swarm Lead:', err);
          });
        }

        // Data Cloud & Agentforce (Signature) — mirror Premier Data mappings, read from Data Cloud and AF sheet
        const dataContainer = document.getElementById('sig-data-section-content');
        if (dataContainer) {
          const dataRange = `'${SHEET_DATA_CLOUD_AF}'!A:Z`;
          // Use cached weekend roster data for data cloud
          getWeekendRosterData(SHEET_DATA_CLOUD_AF, shift, weekendDateStr, currentMode, (rData) => {
            try {
              const rowsD = (rData && rData.values) ? rData.values : [];
              let rowIdxD = -1;
              const targetDatePadded = weekendDateStr;
              const dNumTarget = parseInt(targetDatePadded.split('/')[1], 10);
              const mNumTarget = parseInt(targetDatePadded.split('/')[0], 10);
              const yNumTarget = parseInt(targetDatePadded.split('/')[2], 10);
              const targetDateAlt = `${mNumTarget}/${dNumTarget}/${yNumTarget}`;
              for (let i = 0; i < rowsD.length; i++) {
                const aVal = rowsD[i][0];
                if (!aVal) continue;
                const aStr = String(aVal).trim();
                if (aStr === targetDatePadded || aStr === targetDateAlt) { rowIdxD = i; break; }
                const dt = new Date(aStr);
                if (!isNaN(dt.getTime())) {
                  const am = dt.getMonth() + 1, ad = dt.getDate(), ay = dt.getFullYear();
                  if (am === mNumTarget && ad === dNumTarget && ay === yNumTarget) { rowIdxD = i; break; }
                }
              }

              const blocks = [
                { title: `DataCloud ${shift}`, col: getPremierDataCloudColumn(shift), id: 'sig-data-datacloud' },
                { title: `Agentforce ${shift}`, col: getPremierAgentforceColumn(shift), id: 'sig-data-agentforce' },
              ];
              blocks.forEach(async ({ title, col, id }) => {
                let cellVal = '';
                if (rowIdxD >= 0) {
                  const colIdx = col.charCodeAt(0) - 'A'.charCodeAt(0);
                  cellVal = (rowsD[rowIdxD][colIdx] || '').toString().trim();
                }
                try { cellVal = (await getCellValueWithoutStrikethrough(SHEET_DATA_CLOUD_AF, rowIdxD, col)) || cellVal; } catch { }
                if (!cellVal) {
                  dataContainer.insertAdjacentHTML('beforeend', `
                    <div style="margin-top:12px;">
                      <div style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; box-shadow: 0 1px 2px rgba(0,0,0,0.04); padding:16px; animation: fadeUp 260ms ease;">
                        <div style="font-weight:700; color:#0f766e; margin-bottom:4px;">${title}</div>
                        <div style="font-size:14px; color:#64748b;">No data for ${shift} on ${weekendDateStr}.</div>
                      </div>
                    </div>
                  `);
                } else {
                  const copyText = normalizeNamesForCopy(cellVal);
                  dataContainer.insertAdjacentHTML('beforeend', `
                    <div style="margin-top:12px;">
                      <div style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; box-shadow: 0 1px 2px rgba(0,0,0,0.04); padding:16px; animation: fadeUp 260ms ease;">
                        <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;">
                          <div style="font-weight:700; color:#0f766e;">${title}</div>
                          <span style="font-size:12px; color:#64748b; background:#f1f5f9; border:1px solid #e2e8f0; padding:4px 8px; border-radius:999px;">Click to copy</span>
                        </div>
                        <div id="${id}" title="Click to copy" style="padding:12px; border:1px dashed #94a3b8; border-radius:8px; background:#f8fafc; cursor:pointer; white-space:normal; word-break:break-word; overflow-wrap:anywhere;">${formatNamesToMultiline(cellVal)}</div>
                      </div>
                    </div>
                  `);
                  const copyEl = document.getElementById(id);
                  if (copyEl) {
                    copyEl.addEventListener('click', () => {
                      navigator.clipboard.writeText(copyText).then(() => showToast(`${title} copied`)).catch(() => showToast('Copy failed'));
                    });
                  }
                }
              });
            } catch (e) {
              console.error('Failed building Signature Data section:', e);
            }
          }, (err) => {
            console.error('Failed loading Data Cloud and AF sheet:', err);
          });
        }

      }

      // In Signature mode, also show Sales/Service/Industry Cloud TEs sections
      if (currentMode === 'signature') {
        const salesContainer = document.getElementById('sig-sales-section-content');
        const serviceContainer = document.getElementById('sig-service-section-content');
        const industryContainer = document.getElementById('sig-industry-section-content');
        if (salesContainer && serviceContainer && industryContainer) {
          // Helper to append a TE section for a given sheet and title
          const appendTESection = (sheetName, titlePrefix, idSuffix, containerEl) => {
            const teColumn = (sheetName === SHEET_INDUSTRY_CLOUD)
              ? getPremierIndustryTEColumn(shift)
              : getTEColumnForShift(shift);
            const teRange = `'${sheetName}'!A:Z`;
            // Use cached weekend roster data for TE
            getWeekendRosterData(sheetName, shift, weekendDateStr, currentMode, (r2) => {
              try {
                const rows2 = r2.values || [];
                let rowIdx2 = -1;
                const targetDatePadded = weekendDateStr;
                const dNumTarget = parseInt(targetDatePadded.split('/')[1], 10);
                const mNumTarget = parseInt(targetDatePadded.split('/')[0], 10);
                const yNumTarget = parseInt(targetDatePadded.split('/')[2], 10);
                const targetDateAlt = `${mNumTarget}/${dNumTarget}/${yNumTarget}`;
                for (let i = 0; i < rows2.length; i++) {
                  const aVal2 = rows2[i][0];
                  if (!aVal2) continue;
                  const aStr2 = String(aVal2).trim();
                  if (aStr2 === targetDatePadded || aStr2 === targetDateAlt) { rowIdx2 = i; break; }
                  const dt2 = new Date(aVal2);
                  if (!isNaN(dt2.getTime())) {
                    const am2 = dt2.getMonth() + 1, ad2 = dt2.getDate(), ay2 = dt2.getFullYear();
                    if (am2 === mNumTarget && ad2 === dNumTarget && ay2 === yNumTarget) { rowIdx2 = i; break; }
                  }
                }

                let teNamesCell = '';
                if (rowIdx2 >= 0) {
                  const colIdx2 = teColumn.charCodeAt(0) - 'A'.charCodeAt(0);
                  teNamesCell = (rows2[rowIdx2][colIdx2] || '').toString();
                }

                const sectionTitle = `${titlePrefix} ${shift}`;
                let sectionHtml = '';
                if (!teNamesCell) {
                  sectionHtml = `
                    <div style="margin-top:12px;">
                      <div style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; box-shadow: 0 1px 2px rgba(0,0,0,0.04); padding:16px; animation: fadeUp 260ms ease;">
                        <div style="font-weight:700; color:#0f766e; margin-bottom:4px;">${sectionTitle}</div>
                        <div style="font-size:14px; color:#64748b;">No TEs found for ${shift} on ${weekendDateStr}.</div>
                      </div>
                    </div>`;
                  containerEl.insertAdjacentHTML('beforeend', sectionHtml);
                  return;
                }

                const teNamesComma = normalizeNamesForCopy(teNamesCell);
                sectionHtml = `
                  <div style="margin-top:12px;">
                    <div style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; box-shadow: 0 1px 2px rgba(0,0,0,0.04); padding:16px; animation: fadeUp 260ms ease;">
                      <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;">
                        <div style="font-weight:700; color:#0f766e;">${sectionTitle}</div>
                        <span style="font-size:12px; color:#64748b; background:#f1f5f9; border:1px solid #e2e8f0; padding:4px 8px; border-radius:999px;">Click to copy</span>
                      </div>
                      <div id="${idSuffix}-names" title="Click to copy" style="padding:12px; border:1px dashed #94a3b8; border-radius:8px; background:#f8fafc; cursor:pointer; white-space:normal; word-break:break-word; overflow-wrap:anywhere;">${formatNamesToMultiline(teNamesCell)}</div>
                    </div>
                  </div>`;
                containerEl.insertAdjacentHTML('beforeend', sectionHtml);

                const teDiv = document.getElementById(`${idSuffix}-names`);
                if (teDiv) {
                  teDiv.addEventListener('click', () => {
                    navigator.clipboard.writeText(teNamesComma).then(() => {
                      showToast(`${titlePrefix} copied`);
                    }).catch(() => showToast('Copy failed'));
                  });
                }
              } catch (e) {
                console.error(`Failed building ${titlePrefix} section:`, e);
              }
            }, (err2) => {
              console.error(`Failed loading ${sheetName} sheet:`, err2);
            });
          };

          appendTESection(SHEET_SALES_CLOUD, 'Sales Cloud TEs', 'sales-te', salesContainer);
          appendTESection(SHEET_SERVICE_CLOUD, 'Service Cloud TEs', 'service-te', serviceContainer);
          appendTESection(SHEET_INDUSTRY_CLOUD, 'Industry Cloud TEs', 'industry-te', industryContainer);
        }
      }

      // Premier mode: add dropdown-driven sections
      if (currentMode === 'premier') {
        const container = document.getElementById('premier-sections');
        const select = document.getElementById('premier-view');
        const resetBtn = document.getElementById('premier-reset');
        const editBtn = document.getElementById('premier-edit');
        let currentPremierBlocks = [];
        if (container && select) {
          if (resetBtn) {
            resetBtn.addEventListener('click', () => {
              resetPremierCountersAll(weekendDateStr, shift); // re-render current view
              const v = select.value; if (v === 'sales') return renderSales(); if (v === 'service') return renderService(); if (v === 'industry') return renderIndustry(); return renderData();
            });
          }
          const openPremierEditDialog = () => {
            if (!currentPremierBlocks || currentPremierBlocks.length === 0) { showToast('Nothing to edit'); return; }
            const dlg = document.createElement('div');
            dlg.style.position = 'fixed';
            dlg.style.inset = '0';
            dlg.style.background = 'rgba(0,0,0,0.4)';
            dlg.style.display = 'flex';
            dlg.style.alignItems = 'center';
            dlg.style.justifyContent = 'center';
            dlg.style.zIndex = '1005';
            dlg.innerHTML = `
              <div style="background:#fff; padding:16px; border-radius:8px; width:380px; max-width:90vw; box-shadow:0 10px 25px rgba(0,0,0,0.2);">
                <div style="font-weight:700; margin-bottom:10px; color:#0f172a;">Edit Premier TE</div>
                <div style="display:flex; flex-direction:column; gap:10px;">
                  <label style="font-size:12px; color:#475569;">Block</label>
                  <select id="pe-block" class="slds-select">${currentPremierBlocks.map(b => `<option value="${b.blockId}">${b.title}</option>`).join('')}</select>
                  <label style="font-size:12px; color:#475569;">Select name</label>
                  <select id="pe-name" class="slds-select"></select>
                  <label style="font-size:12px; color:#475569;">Search Email with Full Name</label>
                  <input id="pe-newname" class="slds-input" placeholder="Leave same to keep" />
                  <label style="font-size:12px; color:#475569;">Search Email</label>
                  <div style="display:flex; gap:6px;">
                    <input id="pe-email" class="slds-input" placeholder="Search and select email" style="flex:1;" />
                    <button id="pe-search" class="slds-button slds-button_neutral" title="Search by exact Name in Salesforce">Search</button>
                  </div>
                  <div id="pe-results" style="max-height:140px; overflow:auto; border:1px solid #e5e7eb; border-radius:6px; padding:6px; display:none;"></div>
                </div>
                <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:14px;">
                  <button id="pe-cancel" class="slds-button slds-button_neutral">Cancel</button>
                  <button id="pe-save" class="slds-button slds-button_brand">Save</button>
                </div>
              </div>`;
            document.body.appendChild(dlg);
            const blockSel = dlg.querySelector('#pe-block');
            const nameSel = dlg.querySelector('#pe-name');
            const emailInp = dlg.querySelector('#pe-email');
            const newNameInp = dlg.querySelector('#pe-newname');
            const refreshNames = () => {
              const b = currentPremierBlocks.find(x => x.blockId === blockSel.value);
              const rows = b && b.mount ? Array.from(b.mount.querySelectorAll('.pc-row')) : [];
              const namesArr = rows.map(r => r.getAttribute('data-name'));
              nameSel.innerHTML = namesArr.map(n => `<option value="${n}">${n}</option>`).join('');
              const sel = namesArr[0] || '';
              const row = rows.find(r => (r.getAttribute('data-name') || '') === sel);
              const mailBtn = row ? row.querySelector('.pc-mail') : null;
              const em = mailBtn ? (mailBtn.getAttribute('data-email') || '') : '';
              emailInp.value = em || '';
              newNameInp.value = sel || '';
            };
            blockSel.addEventListener('change', refreshNames);
            nameSel.addEventListener('change', () => {
              const b = currentPremierBlocks.find(x => x.blockId === blockSel.value);
              const sel = nameSel.value;
              const rows = b && b.mount ? Array.from(b.mount.querySelectorAll('.pc-row')) : [];
              const row = rows.find(r => (r.getAttribute('data-name') || '') === sel);
              const mailBtn = row ? row.querySelector('.pc-mail') : null;
              const em = mailBtn ? (mailBtn.getAttribute('data-email') || '') : '';
              emailInp.value = em || '';
              newNameInp.value = sel || '';
            });
            refreshNames();
            dlg.querySelector('#pe-search').addEventListener('click', async () => {
              const qname = (newNameInp.value || '').trim();
              const resultsEl = dlg.querySelector('#pe-results');
              resultsEl.style.display = 'none';
              resultsEl.innerHTML = '';
              if (!qname) return;
              try {
                if (!window.jsforce || !SESSION_ID) {
                  const m = await getEmailsByNames([qname]);
                  const cleaned = cleanRosterNameForQuery(qname).toLowerCase();
                  const fallback = m[qname] || m[qname.toLowerCase()] || m[cleaned] || '';
                  if (fallback) { emailInp.value = fallback; }
                  return;
                }
                const conn = new jsforce.Connection({ serverUrl: 'https://orgcs.my.salesforce.com', sessionId: SESSION_ID, version: '64.0' });
                const soql = `SELECT Name, Email, Title, Username FROM User WHERE IsActive = true AND Email != null AND (Username LIKE '%dreamevent.com' OR Username LIKE '%orgcs.com') AND Name ='${escapeSoqlString(qname)}'`;
                conn.query(soql, (err, res) => {
                  if (err) {
                    resultsEl.style.display = 'block';
                    resultsEl.innerHTML = `<div style=\"color:#b91c1c; font-size:12px;\">Search failed: ${String(err.message || err).replace(/</g, '&lt;')}</div>`;
                    return;
                  }
                  const recs = (res && res.records) ? res.records : [];
                  if (recs.length === 0) {
                    resultsEl.style.display = 'block';
                    resultsEl.innerHTML = `<div style=\"color:#64748b; font-size:12px;\">No results found</div>`;
                    return;
                  }
                  resultsEl.style.display = 'block';
                  resultsEl.innerHTML = recs.map(r => {
                    const nm = String(r.Name || ''); const em = String(r.Email || ''); const un = String(r.Username || ''); const dg = String(r.Title || '');
                    return `<div class=\"pe-result\" data-name=\"${nm.replace(/\"/g, '&quot;')}\" data-email=\"${em.replace(/\"/g, '&quot;')}\" style=\"padding:6px 8px; border-radius:6px; cursor:pointer; display:flex; flex-direction:column; gap:2px;\">`
                      + `<div style=\"font-size:13px; color:#0f172a;\">${nm}</div>`
                      + `<div style=\"font-size:12px; color:#475569;\">${em} · ${un} · ${dg}</div>`
                      + `</div>`;
                  }).join('');
                  resultsEl.querySelectorAll('.pe-result').forEach(item => {
                    item.addEventListener('click', () => {
                      const selName = item.getAttribute('data-name') || '';
                      const selEmail = item.getAttribute('data-email') || '';
                      if (selEmail) emailInp.value = selEmail;
                      resultsEl.querySelectorAll('.pe-result').forEach(el => { el.style.background = 'transparent'; });
                      item.style.background = '#f1f5f9';
                    });
                  });
                });
              } catch (e) {
                resultsEl.style.display = 'block';
                resultsEl.innerHTML = `<div style=\"color:#b91c1c; font-size:12px;\">Search error</div>`;
              }
            });
            dlg.querySelector('#pe-cancel').addEventListener('click', () => dlg.remove());
            dlg.querySelector('#pe-save').addEventListener('click', async () => {
              const bId = blockSel.value;
              const original = nameSel.value;
              const newName = (newNameInp.value || '').trim() || original;
              const email = (emailInp.value || '').trim();
              setPremierOverride(weekendDateStr, shift, bId, original, newName, email);
              try {
                if (email && newName) {
                  const cached = loadUserEmailCache();
                  const map = cached.map || {};
                  const cleaned = cleanRosterNameForQuery(newName).toLowerCase();
                  map[cleaned] = email; map[newName.toLowerCase()] = email;
                  saveUserEmailCache(map);
                }
              } catch { }
              // Re-render current view
              const v = select.value; if (v === 'sales') await renderSales(); else if (v === 'service') await renderService(); else if (v === 'industry') await renderIndustry(); else await renderData();
              dlg.remove();
            });
          };
          if (editBtn) {
            editBtn.addEventListener('click', openPremierEditDialog);
          }
          const renderSales = () => {
            container.innerHTML = '';
            currentPremierBlocks = [];
            // 3 cells from Sales Cloud sheet
            const rangePS = `'${SHEET_SALES_CLOUD}'!A:Z`;
            // Use cached weekend roster data for Premier Sales
            getWeekendRosterData(SHEET_SALES_CLOUD, shift, weekendDateStr, currentMode, (r) => {
              const rowsPS = r.values || [];
              let rowIdx = -1;
              const targetDatePadded = weekendDateStr;
              const dNumTarget = parseInt(targetDatePadded.split('/')[1], 10);
              const mNumTarget = parseInt(targetDatePadded.split('/')[0], 10);
              const yNumTarget = parseInt(targetDatePadded.split('/')[2], 10);
              const targetDateAlt = `${mNumTarget}/${dNumTarget}/${yNumTarget}`;
              for (let i = 0; i < rowsPS.length; i++) {
                const aVal = rowsPS[i][0];
                if (!aVal) continue;
                const aStr = String(aVal).trim();
                if (aStr === targetDatePadded || aStr === targetDateAlt) { rowIdx = i; break; }
                const dt = new Date(aStr);
                if (!isNaN(dt.getTime())) {
                  const am = dt.getMonth() + 1, ad = dt.getDate(), ay = dt.getFullYear();
                  if (am === mNumTarget && ad === dNumTarget && ay === yNumTarget) { rowIdx = i; break; }
                }
              }
              const getCell = (col) => {
                if (rowIdx < 0) return '';
                const idx = col.charCodeAt(0) - 'A'.charCodeAt(0);
                return (rowsPS[rowIdx][idx] || '').toString();
              };
              const blocks = [
                { title: `${shift} Dev TEs`, col: getPremierSalesDevTEColumn(shift), id: 'premier-sales-dev-tes' },
                { title: `${shift} Non-Dev TEs`, col: getPremierSalesNonDevTEColumn(shift), id: 'premier-sales-nondev-tes' },
                { title: `${shift} Swarm Lead`, col: getPremierSalesSwarmLeadColumn(shift), id: 'premier-sales-swarm' },
              ];
              blocks.forEach(async ({ title, col, id }) => {
                let val = getCell(col).trim();
                try { val = (await getCellValueWithoutStrikethrough(SHEET_SALES_CLOUD, rowIdx, col)) || val; } catch { }
                const html = val
                  ? `
                    <div style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; box-shadow:0 1px 2px rgba(0,0,0,0.04); padding:16px; animation: fadeUp 260ms ease;">
                      <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;">
                        <div style="font-weight:700; color:#0f766e;">${title}</div>
                        <span style="font-size:12px; color:#64748b; background:#f1f5f9; border:1px solid #e2e8f0; padding:4px 8px; border-radius:999px;">Click to copy</span>
                      </div>
                      <div id="${id}" title="Click to copy" style="padding:12px; border:1px dashed #94a3b8; border-radius:8px; background:#f8fafc; cursor:pointer; white-space:normal; word-break:break-word; overflow-wrap:anywhere;">${formatNamesToMultiline(val)}</div>
                      <div id="${id}-counters" style="margin-top:10px;"></div>
                    </div>`
                  : `
                    <div style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; box-shadow:0 1px 2px rgba(0,0,0,0.04); padding:16px; animation: fadeUp 260ms ease;">
                      <div style="font-weight:700; color:#0f766e; margin-bottom:4px;">${title}</div>
                      <div style="font-size:14px; color:#64748b;">No data for ${shift} on ${weekendDateStr}.</div>
                    </div>`;
                container.insertAdjacentHTML('beforeend', html);
                if (val) {
                  const copyEl = document.getElementById(id);
                  if (copyEl) {
                    const toCopy = normalizeNamesForCopy(val);
                    copyEl.addEventListener('click', () => navigator.clipboard.writeText(toCopy).then(() => showToast(`${title} copied`)).catch(() => showToast('Copy failed')));
                  }
                  const countersEl = document.getElementById(`${id}-counters`);
                  if (countersEl) {
                    const names = parseRosterNames(val);
                    let emailMap = {};
                    try { emailMap = await getEmailsByNames(names); } catch { }
                    renderPremierCounters(countersEl, names, { dateStr: weekendDateStr, shift, blockId: id, emailMap });
                    currentPremierBlocks.push({ blockId: id, title, names, emailMap, mount: countersEl });
                  }
                }
              });
            }, (e) => console.error('Premier Sales read fail:', e));
          };

          const renderService = () => {
            container.innerHTML = '';
            currentPremierBlocks = [];
            // 3 cells from Service Cloud sheet
            const rangePSe = `'${SHEET_SERVICE_CLOUD}'!A:Z`;
            // Use cached weekend roster data for Premier Service
            getWeekendRosterData(SHEET_SERVICE_CLOUD, shift, weekendDateStr, currentMode, (r) => {
              const rowsPS = r.values || [];
              let rowIdx = -1;
              const targetDatePadded = weekendDateStr;
              const dNumTarget = parseInt(targetDatePadded.split('/')[1], 10);
              const mNumTarget = parseInt(targetDatePadded.split('/')[0], 10);
              const yNumTarget = parseInt(targetDatePadded.split('/')[2], 10);
              const targetDateAlt = `${mNumTarget}/${dNumTarget}/${yNumTarget}`;
              for (let i = 0; i < rowsPS.length; i++) {
                const aVal = rowsPS[i][0];
                if (!aVal) continue;
                const aStr = String(aVal).trim();
                if (aStr === targetDatePadded || aStr === targetDateAlt) { rowIdx = i; break; }
                const dt = new Date(aStr);
                if (!isNaN(dt.getTime())) {
                  const am = dt.getMonth() + 1, ad = dt.getDate(), ay = dt.getFullYear();
                  if (am === mNumTarget && ad === dNumTarget && ay === yNumTarget) { rowIdx = i; break; }
                }
              }
              const getCell = (col) => {
                if (rowIdx < 0) return '';
                const idx = col.charCodeAt(0) - 'A'.charCodeAt(0);
                return (rowsPS[rowIdx][idx] || '').toString();
              };
              const blocks = [
                { title: `${shift} Dev TEs`, col: getPremierSalesDevTEColumn(shift), id: 'premier-service-dev-tes' },
                { title: `${shift} Non-Dev TEs`, col: getPremierSalesNonDevTEColumn(shift), id: 'premier-service-nondev-tes' },
                { title: `${shift} Swarm Lead`, col: getPremierSalesSwarmLeadColumn(shift), id: 'premier-service-swarm' },
              ];
              blocks.forEach(async ({ title, col, id }) => {
                let val = getCell(col).trim();
                try { val = (await getCellValueWithoutStrikethrough(SHEET_SERVICE_CLOUD, rowIdx, col)) || val; } catch { }
                const html = val
                  ? `
                    <div style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; box-shadow:0 1px 2px rgba(0,0,0,0.04); padding:16px; animation: fadeUp 260ms ease;">
                      <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;">
                        <div style="font-weight:700; color:#0f766e;">${title}</div>
                        <span style="font-size:12px; color:#64748b; background:#f1f5f9; border:1px solid #e2e8f0; padding:4px 8px; border-radius:999px;">Click to copy</span>
                      </div>
                      <div id="${id}" title="Click to copy" style="padding:12px; border:1px dashed #94a3b8; border-radius:8px; background:#f8fafc; cursor:pointer; white-space:normal; word-break:break-word; overflow-wrap:anywhere;">${formatNamesToMultiline(val)}</div>
                      <div id="${id}-counters" style="margin-top:10px;"></div>
                    </div>`
                  : `
                    <div style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; box-shadow:0 1px 2px rgba(0,0,0,0.04); padding:16px; animation: fadeUp 260ms ease;">
                      <div style="font-weight:700; color:#0f766e; margin-bottom:4px;">${title}</div>
                      <div style="font-size:14px; color:#64748b;">No data for ${shift} on ${weekendDateStr}.</div>
                    </div>`;
                container.insertAdjacentHTML('beforeend', html);
                if (val) {
                  const copyEl = document.getElementById(id);
                  if (copyEl) {
                    const toCopy = normalizeNamesForCopy(val);
                    copyEl.addEventListener('click', () => navigator.clipboard.writeText(toCopy).then(() => showToast(`${title} copied`)).catch(() => showToast('Copy failed')));
                  }
                  const countersEl = document.getElementById(`${id}-counters`);
                  if (countersEl) {
                    const names = parseRosterNames(val);
                    let emailMap = {};
                    try { emailMap = await getEmailsByNames(names); } catch { }
                    renderPremierCounters(countersEl, names, { dateStr: weekendDateStr, shift, blockId: id, emailMap });
                    currentPremierBlocks.push({ blockId: id, title, names, emailMap, mount: countersEl });
                  }
                }
              });
            }, (e) => console.error('Premier Service read fail:', e));
          };

          const renderIndustry = () => {
            container.innerHTML = '';
            currentPremierBlocks = [];
            // 2 cells: Industry TEs, Swarm Lead from Industry Cloud sheet
            const rangePI = `'${SHEET_INDUSTRY_CLOUD}'!A:Z`;
            // Use cached weekend roster data for Premier Industry
            getWeekendRosterData(SHEET_INDUSTRY_CLOUD, shift, weekendDateStr, currentMode, (r) => {
              const rowsPI = r.values || [];
              let rowIdx = -1;
              const targetDatePadded = weekendDateStr;
              const dNumTarget = parseInt(targetDatePadded.split('/')[1], 10);
              const mNumTarget = parseInt(targetDatePadded.split('/')[0], 10);
              const yNumTarget = parseInt(targetDatePadded.split('/')[2], 10);
              const targetDateAlt = `${mNumTarget}/${dNumTarget}/${yNumTarget}`;
              for (let i = 0; i < rowsPI.length; i++) {
                const aVal = rowsPI[i][0];
                if (!aVal) continue;
                const aStr = String(aVal).trim();
                if (aStr === targetDatePadded || aStr === targetDateAlt) { rowIdx = i; break; }
                const dt = new Date(aStr);
                if (!isNaN(dt.getTime())) {
                  const am = dt.getMonth() + 1, ad = dt.getDate(), ay = dt.getFullYear();
                  if (am === mNumTarget && ad === dNumTarget && ay === yNumTarget) { rowIdx = i; break; }
                }
              }
              const getCell = (col) => {
                if (rowIdx < 0) return '';
                const idx = col.charCodeAt(0) - 'A'.charCodeAt(0);
                return (rowsPI[rowIdx][idx] || '').toString();
              };
              const blocks = [
                { title: `${shift} Industry TEs`, col: getPremierIndustryTEColumn(shift), id: 'premier-industry-tes' },
                { title: `${shift} Swarm Lead`, col: getPremierIndustrySwarmLeadColumn(shift), id: 'premier-industry-swarm' },
              ];
              blocks.forEach(async ({ title, col, id }) => {
                let val = getCell(col).trim();
                try { val = (await getCellValueWithoutStrikethrough(SHEET_INDUSTRY_CLOUD, rowIdx, col)) || val; } catch { }
                const html = val
                  ? `
                    <div style=\"border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; box-shadow:0 1px 2px rgba(0,0,0,0.04); padding:16px; animation: fadeUp 260ms ease;\">\n                      <div style=\"display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;\">\n                        <div style=\"font-weight:700; color:#0f766e;\">${title}</div>\n                        <span style=\"font-size:12px; color:#64748b; background:#f1f5f9; border:1px solid #e2e8f0; padding:4px 8px; border-radius:999px;\">Click to copy</span>\n                      </div>\n                      <div id=\"${id}\" title=\"Click to copy\" style=\"padding:12px; border:1px dashed #94a3b8; border-radius:8px; background:#f8fafc; cursor:pointer; white-space:normal; word-break:break-word; overflow-wrap:anywhere;\">${formatNamesToMultiline(val)}</div>\n                    </div>`
                  : `
                    <div style=\"border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; box-shadow:0 1px 2px rgba(0,0,0,0.04); padding:16px; animation: fadeUp 260ms ease;\">\n                      <div style=\"font-weight:700; color:#0f766e; margin-bottom:4px;\">${title}</div>\n                      <div style=\"font-size:14px; color:#64748b;\">No data for ${shift} on ${weekendDateStr}.</div>\n                    </div>`;
                container.insertAdjacentHTML('beforeend', html);
                if (val) {
                  const copyEl = document.getElementById(id);
                  if (copyEl) {
                    const toCopy = normalizeNamesForCopy(val);
                    copyEl.addEventListener('click', () => navigator.clipboard.writeText(toCopy).then(() => showToast(`${title} copied`)).catch(() => showToast('Copy failed')));
                  }
                  // Add counters under each industry block
                  const countersEl = document.createElement('div');
                  countersEl.id = `${id}-counters`;
                  countersEl.style.marginTop = '10px';
                  const lastBlock = container.lastElementChild;
                  if (lastBlock) lastBlock.appendChild(countersEl);
                  const names = parseRosterNames(val);
                  let emailMap = {};
                  try { emailMap = await getEmailsByNames(names); } catch { }
                  renderPremierCounters(countersEl, names, { dateStr: weekendDateStr, shift, blockId: id, emailMap });
                  currentPremierBlocks.push({ blockId: id, title, names, emailMap, mount: countersEl });
                }
              });
            }, (e) => console.error('Premier Industry read fail:', e));
          };

          const renderData = () => {
            container.innerHTML = '';
            currentPremierBlocks = [];
            // 2 cells from Data Cloud and AF sheet: DataCloud {Shift}, Agentforce {Shift}
            const rangePD = `'${SHEET_DATA_CLOUD_AF}'!A:Z`;
            // Use cached weekend roster data for Premier Data
            getWeekendRosterData(SHEET_DATA_CLOUD_AF, shift, weekendDateStr, currentMode, (r) => {
              const rowsPD = r.values || [];
              let rowIdx = -1;
              const targetDatePadded = weekendDateStr;
              const dNumTarget = parseInt(targetDatePadded.split('/')[1], 10);
              const mNumTarget = parseInt(targetDatePadded.split('/')[0], 10);
              const yNumTarget = parseInt(targetDatePadded.split('/')[2], 10);
              const targetDateAlt = `${mNumTarget}/${dNumTarget}/${yNumTarget}`;
              for (let i = 0; i < rowsPD.length; i++) {
                const aVal = rowsPD[i][0];
                if (!aVal) continue;
                const aStr = String(aVal).trim();
                if (aStr === targetDatePadded || aStr === targetDateAlt) { rowIdx = i; break; }
                const dt = new Date(aStr);
                if (!isNaN(dt.getTime())) {
                  const am = dt.getMonth() + 1, ad = dt.getDate(), ay = dt.getFullYear();
                  if (am === mNumTarget && ad === dNumTarget && ay === yNumTarget) { rowIdx = i; break; }
                }
              }
              const getCell = (col) => {
                if (rowIdx < 0) return '';
                const idx = col.charCodeAt(0) - 'A'.charCodeAt(0);
                return (rowsPD[rowIdx][idx] || '').toString();
              };
              const blocks = [
                { title: `DataCloud ${shift}`, col: getPremierDataCloudColumn(shift), id: 'premier-data-datacloud' },
                { title: `Agentforce ${shift}`, col: getPremierAgentforceColumn(shift), id: 'premier-data-agentforce' },
              ];
              blocks.forEach(async ({ title, col, id }) => {
                let val = getCell(col).trim();
                try { val = (await getCellValueWithoutStrikethrough(SHEET_DATA_CLOUD_AF, rowIdx, col)) || val; } catch { }
                const html = val
                  ? `
                    <div style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; box-shadow:0 1px 2px rgba(0,0,0,0.04); padding:16px; animation: fadeUp 260ms ease;">
                      <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;">
                        <div style="font-weight:700; color:#0f766e;">${title}</div>
                        <span style="font-size:12px; color:#64748b; background:#f1f5f9; border:1px solid #e2e8f0; padding:4px 8px; border-radius:999px;">Click to copy</span>
                      </div>
                      <div id="${id}" title="Click to copy" style="padding:12px; border:1px dashed #94a3b8; border-radius:8px; background:#f8fafc; cursor:pointer; white-space:normal; word-break:break-word; overflow-wrap:anywhere;">${formatNamesToMultiline(val)}</div>
                      <div id="${id}-counters" style="margin-top:10px;"></div>
                    </div>`
                  : `
                    <div style="border:1px solid #e5e7eb; border-radius:12px; background:#ffffff; box-shadow:0 1px 2px rgba(0,0,0,0.04); padding:16px; animation: fadeUp 260ms ease;">
                      <div style="font-weight:700; color:#0f766e; margin-bottom:4px;">${title}</div>
                      <div style="font-size:14px; color:#64748b;">No data for ${shift} on ${weekendDateStr}.</div>
                    </div>`;
                container.insertAdjacentHTML('beforeend', html);
                if (val) {
                  const copyEl = document.getElementById(id);
                  if (copyEl) {
                    const toCopy = normalizeNamesForCopy(val);
                    copyEl.addEventListener('click', () => navigator.clipboard.writeText(toCopy).then(() => showToast(`${title} copied`)).catch(() => showToast('Copy failed')));
                  }
                  const countersEl = document.getElementById(`${id}-counters`);
                  if (countersEl) {
                    const names = parseRosterNames(val);
                    let emailMap = {};
                    try { emailMap = await getEmailsByNames(names); } catch { }
                    renderPremierCounters(countersEl, names, { dateStr: weekendDateStr, shift, blockId: id, emailMap });
                    currentPremierBlocks.push({ blockId: id, title, names, emailMap, mount: countersEl });
                  }
                }
              });
            }, (e) => console.error('Premier Data read fail:', e));
          };

          // initial render
          renderSales();
          select.addEventListener('change', () => {
            if (select.value === 'sales') return renderSales();
            if (select.value === 'service') return renderService();
            if (select.value === 'industry') return renderIndustry();
            return renderData();
          });
        }
      }
    } catch (e) {
      console.error(e);
      const bodyEl = document.querySelector('#cic-modal .modal-body');
      if (bodyEl) {
        bodyEl.innerHTML = `<div class="no-cases-message" style="background: linear-gradient(135deg,#fff1f2 0%,#ffe4e6 100%); border: 1.5px solid #fecaca; animation: fadeUp 300ms ease;">
            <h4 class="no-cases-title" style="color:#dc2626;">Error</h4>
            <p class="no-cases-text" style="color:#dc2626;">Failed parsing sheet response.</p>
          </div>`;
      }
    }
  }, (err) => {
    const bodyEl = document.querySelector('#cic-modal .modal-body');
    if (bodyEl) {
      bodyEl.innerHTML = `<div class="no-cases-message" style="background: linear-gradient(135deg,#fff1f2 0%,#ffe4e6 100%); border: 1.5px solid #fecaca; animation: fadeUp 300ms ease;">
          <h4 class="no-cases-title" style="color:#dc2626;">Error Loading</h4>
          <p class="no-cases-text" style="color:#dc2626;">${(err && err.message) || 'Could not fetch Google Sheet.'}</p>
        </div>`;
    }
  });
}

// Normalize names copied from multi-line cells to comma-separated values
function normalizeNamesForCopy(namesCell) {
  return (namesCell || '')
    .replace(/[\r\n]+/g, ', ')
    .replace(/\s*&\s*/g, ', ')
    .replace(/\s*,\s*/g, ', ')
    .replace(/,\s*,+/g, ', ')
    .replace(/\s{2,}/g, ' ')
    .replace(/,\s*$/, '')
    .trim();
}

// Render helper: visually show each name on its own line (affects display only, not copy)
function formatNamesToMultiline(text) {
  if (!text) return '';
  const parts = String(text)
    .split(/\n|,|&/g)
    .map(s => s.trim())
    .filter(Boolean);
  if (parts.length === 0) return '';
  const escapeHtml = (s) => s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
  return parts.map(p => `<div>${escapeHtml(p)}</div>`).join('');
}

// Function to check if there are snoozed cases and update button visibility
function updateClearSnoozedButtonVisibility() {
  const allKeys = Object.keys(localStorage);
  const snoozeKeys = allKeys.filter(key => key.startsWith('snooze_'));
  const clearSnoozedButton = document.getElementById("clear-snoozed-button");

  if (clearSnoozedButton) {
    if (snoozeKeys.length > 0) {
      clearSnoozedButton.style.display = 'inline-block';
    } else {
      clearSnoozedButton.style.display = 'none';
    }
  }
}

function clearSnoozedCases() {
  const allKeys = Object.keys(localStorage);
  const snoozeKeys = allKeys.filter(key => key.startsWith('snooze_'));

  if (snoozeKeys.length === 0) {
    showToast('No snoozed cases found');
    return;
  }

  snoozeKeys.forEach(key => {
    localStorage.removeItem(key);
  });

  showToast(`Cleared ${snoozeKeys.length} snoozed case${snoozeKeys.length === 1 ? '' : 's'}`);
  updateClearSnoozedButtonVisibility();

  setTimeout(() => {
    window.location.reload();
  }, 1000);
}

// GHO Functions
function forceRefreshGHOData() {
  // Clear the GHO cache for the current mode to force a fresh query
  if (window.__ghoCache && window.__ghoCache[currentMode]) {
    delete window.__ghoCache[currentMode];
    // GHO cache cleared for mode - forcing fresh data fetch
    // Save updated cache to persistent storage
    saveGHOCache();
  }

  // Clear global GHO variables
  ghoRecordsGlobal = [];
  ghoConnectionGlobal = null;

  // Reset GHO list state
  window.__ghoListState = null;

  // Show toast notification
  showToast('GHO cache cleared - fetching fresh data...', 'info');

  // If GHO modal is open, refresh the data
  const modal = document.getElementById('gho-modal');
  if (modal && modal.style.display === 'flex') {
    // Add spinning animation to refresh button
    const refreshBtn = document.getElementById('gho-refresh-btn');
    if (refreshBtn) {
      const icon = refreshBtn.querySelector('.fa-refresh');
      if (icon) {
        icon.classList.add('fa-spin');
      }
      // Add visual feedback that refresh is in progress
      refreshBtn.style.background = 'linear-gradient(135deg, #0ea5e9 0%, #0284c7 100%)';
      refreshBtn.style.color = 'white';
      refreshBtn.style.transform = 'scale(1.05)';
    }

    // Force refresh by calling checkGHOStatus with cache bypass
    checkGHOStatus(true);
  } else {
    // If modal is not open, just clear cache for next time
    showToast('GHO cache cleared - will fetch fresh data on next open', 'info');
  }
}

function stopGHORefreshSpinning() {
  const refreshBtn = document.getElementById('gho-refresh-btn');
  if (refreshBtn) {
    const icon = refreshBtn.querySelector('.fa-refresh');
    if (icon) {
      icon.classList.remove('fa-spin');
    }
    // Reset button styling to default
    refreshBtn.style.background = '';
    refreshBtn.style.color = '';
    refreshBtn.style.transform = '';
  }
}

function updateGHOCacheStatus() {
  const cacheInfo = document.getElementById('gho-cache-info');
  const cacheTime = document.getElementById('gho-cache-time');

  if (!cacheInfo || !cacheTime) return;

  const currentModeCache = window.__ghoCache && window.__ghoCache[currentMode];

  // Build cache status for current mode only
  let statusText = 'Cache status: ';
  let timeText = '';
  let statusColor = '#6b7280';

  if (currentModeCache && currentModeCache.fetchedAt) {
    const now = Date.now();
    const age = now - currentModeCache.fetchedAt;
    const ageMinutes = Math.floor(age / (1000 * 60));
    const ageSeconds = Math.floor((age % (1000 * 60)) / 1000);

    if (age < GHO_CACHE_TTL) {
      const remaining = GHO_CACHE_TTL - age;
      const remainingMinutes = Math.floor(remaining / (1000 * 60));
      const remainingSeconds = Math.floor((remaining % (1000 * 60)) / 1000);

      statusText = `Cache status: ${currentMode.charAt(0).toUpperCase() + currentMode.slice(1)} mode - Fresh (${ageMinutes}m ${ageSeconds}s old)`;
      statusColor = '#059669';
      timeText = `Expires in: ${remainingMinutes}m ${remainingSeconds}s`;
    } else {
      statusText = `Cache status: ${currentMode.charAt(0).toUpperCase() + currentMode.slice(1)} mode - Expired`;
      statusColor = '#dc2626';
      timeText = 'Data needs refresh';
    }
  } else {
    statusText = `Cache status: ${currentMode.charAt(0).toUpperCase() + currentMode.slice(1)} mode - No cached data`;
    statusColor = '#6b7280';
    timeText = 'Data will be fetched fresh';
  }

  // Update the display
  cacheInfo.textContent = statusText;
  cacheInfo.style.color = statusColor;
  cacheTime.textContent = timeText;
  cacheTime.style.color = statusColor;

  // Add persistent cache info
  try {
    const persistentCache = localStorage.getItem(GHO_CACHE_STORAGE_KEY);
    if (persistentCache) {
      const parsed = JSON.parse(persistentCache);
      const hasPersistentData = parsed && (parsed.signature || parsed.premier);
      if (hasPersistentData) {
        cacheInfo.textContent += ' (Persistent)';
      }
    }
  } catch (e) {
    // Ignore errors in cache status display
  }

  // Update detailed cache information
  updateDetailedCacheInfo();
}

// Function to update detailed cache information display
function updateDetailedCacheInfo() {
  const currentModeDetail = document.getElementById('current-mode-cache-detail');

  if (!currentModeDetail) return;

  const currentModeCache = window.__ghoCache && window.__ghoCache[currentMode];

  if (currentModeCache && currentModeCache.fetchedAt) {
    const age = Date.now() - currentModeCache.fetchedAt;
    const ageMinutes = Math.floor(age / (1000 * 60));
    const isExpired = age >= GHO_CACHE_TTL;

    const weekendFlag = isCurrentlyWeekend();
    const mondayCount = weekendFlag ? (currentModeCache.mondayCaseIds?.size || 0) : 0;
    currentModeDetail.innerHTML = `
      <strong>${currentMode.charAt(0).toUpperCase() + currentMode.slice(1)} Mode Cache:</strong> ${isExpired ? 'Expired' : 'Fresh'} 
      (${ageMinutes}m old)<br>
      <small>${currentModeCache.records?.length || 0} cases, ${currentModeCache.triageCaseIds?.size || 0} triage${weekendFlag ? `, ${mondayCount} Monday cases` : ''}</small>
    `;
    currentModeDetail.style.color = isExpired ? '#dc2626' : '#059669';
  } else {
    currentModeDetail.innerHTML = `<strong>${currentMode.charAt(0).toUpperCase() + currentMode.slice(1)} Mode:</strong> No cached data`;
    currentModeDetail.style.color = '#6b7280';
  }
}

// ================================================
// Section: 12) GHO Modal Logic (queries, cache, rendering)
// ================================================
function checkGHOStatus(forceRefresh = false) {
  const modal = document.getElementById('gho-modal');
  const container = document.getElementById('gho-cases-container');
  const filterDropdown = document.getElementById('gho-taxonomy-filter');

  // Don't reset filter when opening modal - keep user's selection
  // This allows filtering to work properly with cached data
  if (filterDropdown) {
    console.log('Current filter value when opening modal:', filterDropdown.value);
  }

  // Show modal and loading state
  modal.style.display = 'flex';
  requestAnimationFrame(() => modal.classList.add('modal-show'));

  // Update cache status display
  updateGHOCacheStatus();

  // Start periodic cache status updates
  const cacheStatusInterval = setInterval(updateGHOCacheStatus, 1000);

  // Store the interval ID to clear it later
  modal.dataset.cacheStatusInterval = cacheStatusInterval;

  container.innerHTML = `
    <div class="loading-message" style="animation: fadeUp 300ms ease;">
      <div class="spinner"></div>
      <h4 style="margin-bottom:8px">Loading GHO cases...</h4>
      <p>Please wait while we fetch the latest GHO cases.</p>
    </div>
  `;

  // Use cached results if within TTL and for same shift to avoid repeated network & comment queries
  // Skip cache if forceRefresh is true
  const nowTs = Date.now();
  const currentShiftLive = getCurrentShift();
  const modeCache = window.__ghoCache[currentMode];
  if (!forceRefresh && modeCache && (nowTs - modeCache.fetchedAt < GHO_CACHE_TTL) && modeCache.shift === currentShiftLive) {
    try {
      ghoRecordsGlobal = modeCache.records;
      ghoConnectionGlobal = modeCache.conn;

      console.log('Using cached GHO data:', {
        records: modeCache.records?.length || 0,
        triageCaseIds: modeCache.triageCaseIds?.size || 0,
        mondayCaseIds: modeCache.mondayCaseIds?.size || 0,
        mondayCaseIdsContent: Array.from(modeCache.mondayCaseIds || []),
        noWocTemplateCaseIds: modeCache.noWocTemplateCaseIds?.size || 0,
        noWocTemplateCaseIdsContent: Array.from(modeCache.noWocTemplateCaseIds || [])
      });

      // Re-render with cached comment/user resolution
      // Use current filter value instead of hardcoded 'All'
      const currentFilterValue = document.getElementById('gho-taxonomy-filter')?.value || 'All';
      console.log('Rendering cached data with current filter:', currentFilterValue);

      renderFilteredGHOCases(modeCache.records, modeCache.conn, currentFilterValue, {
        triageCaseIds: modeCache.triageCaseIds,
        mondayCaseIds: modeCache.mondayCaseIds,
        noWocTemplateCaseIds: modeCache.noWocTemplateCaseIds,
        userMap: modeCache.userMap,
        fromCache: true
      });
      return;
    } catch (e) { console.warn('Failed using GHO cache, falling back to fresh query', e); }
  }

  // Get connection and execute fresh GHO query
  let ghoConn = new jsforce.Connection({
    serverUrl: 'https://orgcs.my.salesforce.com',
    sessionId: SESSION_ID,
  });

  // Get current shift dynamically and use in GHO query
  const currentShift = currentShiftLive;
  const preferredShiftValues = getPreferredShiftValues(currentShiftLive);
  const shiftCondition = buildPreferredShiftCondition(preferredShiftValues);

  const signatureGHOQuery = `SELECT Id, CreatedDate, Account.Name, Owner.Name, SE_Target_Response__c, Severity_Level__c, CaseNumber, Subject, CaseRoutingTaxonomy__r.Name, SE_Initial_Response_Status__c, Contact.Is_MVP__c, support_available_timezone__c, (SELECT Transfer_Reason__c, CreatedDate, CreatedById, Preferred_Shift_Old_Value__c, Preferred_Shift_New_Value__c, Severity_New_Value__c, Severity_Old_Value__c FROM Case_Routing_Logs__r ORDER BY CreatedDate DESC LIMIT 5) FROM Case WHERE ((Owner.Name IN ('Skills Queue','Kase Changer', 'Working in Org62','GHO Queue') AND (Case_Support_level__c IN ('Premier Priority','Signature','Signature Success'))) OR (Contact.Is_MVP__c=true AND Owner.Name='GHO Queue')) AND IsClosed=false AND ${shiftCondition} AND ((CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry%' OR CaseRoutingTaxonomy__r.Name LIKE 'Community-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Scale Center%' OR CaseRoutingTaxonomy__r.Name LIKE 'Customer Success Score%' OR CaseRoutingTaxonomy__r.Name LIKE 'Data Cloud-%') AND (Severity_Level__c='Level 1 - Critical' OR Severity_Level__c='Level 2 - Urgent')) AND CaseRoutingTaxonomy__r.Name NOT IN ('Disability and Product Accessibility','DORA')`;
  const premierGHOQuery = `SELECT Case__c, Transfer_Reason__c, Case__r.Case_Support_level__c, Case__r.IsClosed, Case__r.Account_Support_SBR_Category__c, Case__r.Severity_Level__c, Case__r.OrgCS_Owner__c, Case__r.Contact.Is_MVP__c, Case__r.GHO__c, Case__r.Out_of_Impact_Service_Restored__c, Case__r.AX_Product_Name__c, Case__r.CaseRoutingTaxonomy__r.Name FROM Case_Routing_Log__c WHERE Transfer_Reason__c = 'GHO' AND Case__r.Case_Support_level__c In ('Partner Premier', 'Premier', 'Premier+', 'Premium','Standard','') AND Case__r.IsClosed = false AND Case__r.Account_Support_SBR_Category__c != 'JP' AND Case__r.Severity_Level__c IN ('Level 1 - Critical' ,'Level 2 - Urgent') AND Case__r.OrgCS_Owner__c LIKE '%Queue%' AND Case__r.Contact.Is_MVP__c = false AND Case__r.GHO__c = true AND Case__r.Out_of_Impact_Service_Restored__c = false AND (Case__r.AX_Product_Name__c = 'Sales' OR Case__r.AX_Product_Name__c = 'Service'  OR Case__r.AX_Product_Name__c = 'Industry' ) AND Case__r.CaseRoutingTaxonomy__r.Name NOT IN ('Service-Agentforce','Service-Agent for setup','Service-AgentforEmail','Service-Field Service Agentforce','Service-Agentforce for Dev','Sales-Agentforce','Sales -Agentforce for Dev','Sales-Agent for Setup','Sales-Prompt Builder','Data Cloud-Admin','Permissions','Flows','Reports & Dashboards','Data Cloud-Model Builder','Data Cloud-Connectors & Data Streams','Data Cloud-Developer','Calculated Insights & Consumption','Data Cloud-Segments','Activations & Identity Resolution')`;
  const ghoQuery = currentMode === 'premier' ? premierGHOQuery : signatureGHOQuery;

  ghoConn.query(ghoQuery, function (err, result) {
    if (err) {
      console.error('GHO Query Error:', err);

      // Stop refresh button spinning animation on error
      stopGHORefreshSpinning();

      container.innerHTML = `
        <div class="no-cases-message" style="background: linear-gradient(135deg,#fff1f2 0%,#ffe4e6 100%); border: 1.5px solid #fecaca; animation: fadeUp 300ms ease;">
          <h4 class="no-cases-title" style="color: #dc2626;">Error Loading GHO Cases</h4>
          <p class="no-cases-text" style="color: #dc2626;">Failed to fetch GHO cases. Please check your connection and try again.</p>
          <p style="color: #6b7280; font-size: 12px; margin-top: 8px;">Error: ${err.message}</p>
        </div>
      `;
      return;
    }

    // Hydrate full Case records for Premier flow
    if (currentMode === 'premier' && result && Array.isArray(result.records)) {
      const caseIds = Array.from(new Set(result.records.map(r => r.Case__c))).filter(Boolean);
      if (caseIds.length === 0) {
        // No qualifying cases
        stopGHORefreshSpinning();
        updateGHOCacheStatus();
        showGHOStatusModal([], ghoConn);
        return;
      }
      const detailQuery = `SELECT Id, CreatedDate, Account.Name, Owner.Name, SE_Target_Response__c, Severity_Level__c, CaseNumber, Subject, CaseRoutingTaxonomy__r.Name, SE_Initial_Response_Status__c, Contact.Is_MVP__c, support_available_timezone__c, (SELECT Transfer_Reason__c, CreatedDate, CreatedById, Preferred_Shift_Old_Value__c, Preferred_Shift_New_Value__c, Severity_New_Value__c, Severity_Old_Value__c FROM Case_Routing_Logs__r ORDER BY CreatedDate DESC LIMIT 20) FROM Case WHERE Id IN ('${caseIds.join("','")}')`;
      return ghoConn.query(detailQuery, (detailErr, detailRes) => {
        if (detailErr) {
          console.error('Premier detail fetch error:', detailErr);

          // Stop refresh button spinning animation on error
          stopGHORefreshSpinning();

          container.innerHTML = `<div class=\"no-cases-message\" style=\"background:linear-gradient(135deg,#fff1f2 0%,#ffe4e6 100%); border:1.5px solid #fecaca; padding:20px;\"><h4 style=\"color:#dc2626; margin:0 0 8px;\">Error Hydrating Premier GHO Cases</h4><p style=\"margin:0; font-size:14px; color:#6b7280;\">${detailErr.message}</p></div>`;
          return;
        }
        result.records = (detailRes && detailRes.records) ? detailRes.records : [];
        proceedWithGHOCaseProcessing(result);
      });
    }

    // If Signature continue directly
    proceedWithGHOCaseProcessing(result);
  });

  //Premier GHO processing
  function proceedWithGHOCaseProcessing(result) {
    if (!result) {
      stopGHORefreshSpinning();
      updateGHOCacheStatus();
      showGHOStatusModal([], ghoConn);
      return;
    }
    if (result.records.length === 0) {
      stopGHORefreshSpinning();
      updateGHOCacheStatus();
      showGHOStatusModal(result.records || [], ghoConn);
      return;
    }
    const caseIds = result.records.map(r => r.Id);
    // Query comments for all cases, we'll limit to 5 per case in JavaScript
    const commentQuery = `SELECT ParentId, Body, CreatedById, LastModifiedDate, CreatedDate FROM CaseFeed WHERE Visibility = 'InternalUsers' AND ParentId IN ('${caseIds.join("','")}') AND Type = 'TextPost' ORDER BY CreatedDate DESC`;
    Promise.all([
      new Promise(res => ghoConn.query(commentQuery, (e, r) => res({ e, r }))),
      new Promise(res => ghoConn.query(`SELECT Id FROM User WHERE Name = '${currentUserName}' AND IsActive = True AND Username LIKE '%orgcs.com'`, (e, r) => res({ e, r })))
    ]).then(([commentResp, userResp]) => {
      let triageCaseIds = new Set();
      let mondayCaseIds = new Set();
      let noWocTemplateCaseIds = new Set();
      let currentUserId = null;
      if (!userResp.e && userResp.r.records && userResp.r.records.length > 0) currentUserId = userResp.r.records[0].Id;
      if (!commentResp.e && commentResp.r.records) {
        // Group comments by case ID and limit to 5 most recent per case
        const commentsByCase = {};
        commentResp.r.records.forEach(c => {
          if (!commentsByCase[c.ParentId]) {
            commentsByCase[c.ParentId] = [];
          }
          // Only keep the 5 most recent comments per case
          if (commentsByCase[c.ParentId].length < 5) {
            commentsByCase[c.ParentId].push(c);
          }
        });

        // Check each case for GHO triage and Monday mentions
        const isWeekendNow = isCurrentlyWeekend();
        Object.keys(commentsByCase).forEach(caseId => {
          const caseComments = commentsByCase[caseId];

          // Check for GHO triage comments
          caseComments.forEach(c => {
            if (c.Body && c.Body.includes('#GHOTriage') && c.CreatedById === currentUserId) {
              const cDate = c.LastModifiedDate || c.CreatedDate;
              if (isToday(cDate) && getShiftForDate(cDate) === currentShiftLive) {
                const caseRecord = result.records.find(rec => rec.Id === c.ParentId);
                if (caseRecord) {
                  const gKey = `gho_tracked_${caseRecord.Id}`;
                  if (!localStorage.getItem(gKey)) {
                    trackActionAndCount(c.LastModifiedDate || c.CreatedDate, caseRecord.CaseNumber, caseRecord.Severity_Level__c, 'GHO', caseRecord.CaseRoutingTaxonomy__r.Name.split('-')[0], currentMode, currentUserName, 'QB');
                    localStorage.setItem(gKey, 'true');
                  }
                }
                triageCaseIds.add(c.ParentId);
              }
            }
          });

          // Check for Monday mentions in the 5 most recent comments - weekend only
          const hasMondayInComments = isWeekendNow && caseComments.some(comment => {
            if (!comment.Body) return false;
            // Strip HTML tags and check for Monday (case insensitive)
            const cleanBody = comment.Body.replace(/<[^>]*>/g, '').toLowerCase();
            // Check for "Monday" or "TO: {Any word} Monday" pattern
            const hasMonday = cleanBody.includes('monday') ||
              /to:\s*\w+\s+monday/i.test(cleanBody);
            console.log(`Comment check for case ${caseId}: Original="${comment.Body}", Cleaned="${cleanBody}", HasMonday=${hasMonday}`);
            return hasMonday;
          });
          if (hasMondayInComments) {
            mondayCaseIds.add(caseId);
            console.log(`GHO Monday case detected: ${caseId} - found "monday" in 5 latest comments`);
          } else {
            console.log(`No Monday detected for case ${caseId} - comments:`, caseComments.map(c => c.Body));
          }

          // Check for "No {currentShift} GHO template" mentions in the 5 most recent comments - weekend only
          const hasNoWocTemplateInComments = isWeekendNow && caseComments.some(comment => {
            if (!comment.Body) return false;
            // Strip HTML tags and check for "No {currentShift} GHO template" pattern (case insensitive)
            const cleanBody = comment.Body.replace(/<[^>]*>/g, '').toLowerCase();
            const currentShiftLower = currentShiftLive.toLowerCase();
            // More flexible pattern: "no" + any case shift + "gho" + "template" with word boundaries
            const noWocPattern = new RegExp(`\\bno\\s+${currentShiftLower}\\s+gho\\s+template\\b`, 'i');
            const hasNoWocTemplate = noWocPattern.test(cleanBody);
            console.log(`No WOC Template check for case ${caseId}: Original="${comment.Body}", Cleaned="${cleanBody}", HasNoWocTemplate=${hasNoWocTemplate}, Pattern="${noWocPattern.source}", CurrentShift="${currentShiftLive}"`);
            return hasNoWocTemplate;
          });

          if (hasNoWocTemplateInComments) {
            noWocTemplateCaseIds.add(caseId);
            console.log(`No WOC Template case detected: ${caseId} - found "No ${currentShiftLive} GHO template" in 5 latest comments`);
          } else {
            console.log(`No WOC Template not detected for case ${caseId} - comments:`, caseComments.map(c => c.Body));
          }
        });
      }

      window.__ghoCache[currentMode] = {
        records: result.records,
        triageCaseIds,
        mondayCaseIds,
        noWocTemplateCaseIds,
        userMap: {},
        fetchedAt: Date.now(),
        shift: currentShiftLive,
        conn: ghoConn
      };

      // Save cache to persistent storage
      saveGHOCache();

      // Stop refresh button spinning animation
      stopGHORefreshSpinning();

      // Update cache status display
      updateGHOCacheStatus();

      showGHOStatusModal(result.records, ghoConn, { triageCaseIds, mondayCaseIds, fromCache: false });
    }).catch(errAll => {
      console.warn('Parallel GHO enrichment failed, rendering basic list', errAll);

      // Stop refresh button spinning animation even on error
      stopGHORefreshSpinning();

      // Update cache status display
      updateGHOCacheStatus();

      showGHOStatusModal(result.records, ghoConn);
    });
  }

}

// Helper function to get user names from User IDs
function getUserNames(userIds, conn, callback) {
  try {
    const ids = Array.from(new Set(userIds || [])).filter(Boolean);
    if (ids.length === 0) { callback({}); return; }

    // Load persisted cache and determine missing
    const cached = loadUserMapCache();
    const cacheMap = cached.userMap || {};
    const missing = ids.filter(id => !cacheMap[id]);

    // If nothing missing, return cache directly
    if (missing.length === 0) {
      callback({ ...cacheMap });
      return;
    }

    // Query missing ids in chunks to avoid SOQL limits
    const CHUNK = 100;
    const chunks = [];
    for (let i = 0; i < missing.length; i += CHUNK) chunks.push(missing.slice(i, i + CHUNK));

    const resultsMap = { ...cacheMap };
    let idx = 0;
    const runNext = () => {
      if (idx >= chunks.length) {
        // Save merged map back to localStorage and return
        saveUserMapCache(resultsMap);
        callback(resultsMap);
        return;
      }
      const subset = chunks[idx++];
      const q = `SELECT Id, Name FROM User WHERE Id IN ('${subset.join("','")}') AND Username LIKE '%orgcs.com'`;
      conn.query(q, (err, res) => {
        if (err) {
          console.warn('User query failed for chunk, continuing:', err);
        } else if (res && res.records) {
          res.records.forEach(u => { if (u && u.Id) resultsMap[u.Id] = u.Name; });
        }
        runNext();
      });
    };
    runNext();
  } catch (e) {
    console.error('getUserNames failed, returning empty map', e);
    callback({});
  }
}

function showGHOStatusModal(ghoRecords, conn, opts) {
  // Store globally for filtering
  ghoRecordsGlobal = ghoRecords;
  ghoConnectionGlobal = conn;
  const pre = opts || {};

  const container = document.getElementById('gho-cases-container');
  const currentShift = getCurrentShift();

  // Update mapping button with count
  updateMappingButtonCount();
  try {
    const modalBody = document.querySelector('#gho-modal .modal-body');
    if (modalBody && !modalBody.classList.contains('gho-modern-enabled')) {
      modalBody.classList.add('gho-modern-enabled');
    }
  } catch (e) { }
  const modeLabel = currentMode === 'premier' ? 'Premier' : 'Signature';
  try {
    const header = document.querySelector('#gho-modal .modal-header h3');
    if (header) {
      let badge = header.querySelector('.gho-mode-badge');
      if (!badge) {
        badge = document.createElement('span');
        badge.className = 'gho-mode-badge';
        badge.style.cssText = 'margin-left:8px;font-size:11px; letter-spacing:.5px; background:#1e3a8a; color:#fff; padding:3px 8px; border-radius:999px; text-transform:uppercase; box-shadow:0 2px 4px rgba(0,0,0,0.15);';
        header.appendChild(badge);
      }
      badge.textContent = modeLabel + ' Mode';
      badge.style.background = currentMode === 'premier' ? '#065f46' : '#1e3a8a';

      // Add mapping status indicator
      let mappingIndicator = header.querySelector('.gho-mapping-indicator');
      if (!mappingIndicator) {
        mappingIndicator = document.createElement('span');
        mappingIndicator.className = 'gho-mapping-indicator';
        mappingIndicator.style.cssText = 'margin-left:8px;font-size:10px; letter-spacing:.3px; background:#10b981; color:#fff; padding:2px 6px; border-radius:999px; text-transform:uppercase; box-shadow:0 1px 3px rgba(0,0,0,0.1);';
        header.appendChild(mappingIndicator);
      }

      const mappings = loadGHOMappings();
      if (mappings.length > 0) {
        mappingIndicator.textContent = `${mappings.length} Mapping${mappings.length > 1 ? 's' : ''}`;
        mappingIndicator.style.display = 'inline-block';
      } else {
        mappingIndicator.style.display = 'none';
      }
    }
  } catch (e) { }

  if (!ghoRecords || ghoRecords.length === 0) {
    container.innerHTML = `
      <div class="gho-empty-card" style="background: linear-gradient(135deg, #ecfeff 0%, #e0f2fe 100%); border: 1.5px solid #bae6fd; border-radius: 16px; padding: 28px; text-align: center; animation: fadeUp 320ms ease; box-shadow: 0 6px 20px rgba(2,132,199,0.08);">
        <div style="font-size: 28px; margin-bottom: 8px;">🛰️</div>
        <h4 style="margin: 0 0 6px; font-size: 18px; color: #075985;">No GHO cases for now</h4>
        <p style="margin: 0; color: #0c4a6e; font-size: 14px;">All good for <strong>${currentShift}</strong> shift. Great work!</p>
        <div style="display:flex; justify-content:center; gap:8px; margin-top:12px; flex-wrap: wrap;">
          <span style="background:#e0f2fe; color:#075985; border:1px solid #bae6fd; padding:4px 10px; border-radius:999px; font-size:12px;">Shift: ${currentShift}</span>
          <span style="background:#e0f2fe; color:#075985; border:1px solid #bae6fd; padding:4px 10px; border-radius:999px; font-size:12px;">Filter: All</span>
          <span style="background:${currentMode === 'premier' ? '#d1fae5' : '#e0e7ff'}; color:${currentMode === 'premier' ? '#065f46' : '#3730a3'}; border:1px solid ${currentMode === 'premier' ? '#6ee7b7' : '#c7d2fe'}; padding:4px 10px; border-radius:999px; font-size:12px;">Mode: ${modeLabel}</span>
        </div>
        <div style="margin-top:14px; display:flex; justify-content:center; gap:10px;">
          <button id="gho-empty-refresh" class="preview-btn" style="background:#0ea5e9; padding:10px 16px;">Refresh</button>
        </div>
      </div>
    `;
    const refreshBtn = document.getElementById('gho-empty-refresh');
    if (refreshBtn) {
      refreshBtn.addEventListener('click', () => {
        checkGHOStatus();
      });
    }
    return;
  }

  // Render with current filter
  const filterValue = document.getElementById('gho-taxonomy-filter').value;
  renderFilteredGHOCases(ghoRecords, conn, filterValue, pre);
}

// Function to filter and render GHO cases based on CaseRoutingTaxonomy__r.Name
function renderFilteredGHOCases(ghoRecords, conn, filterValue = 'All', pre = {}) {
  const container = document.getElementById('gho-cases-container');
  const currentShift = getCurrentShift();

  // Filter records based on CaseRoutingTaxonomy__r.Name
  let filteredRecords = ghoRecords;
  if (filterValue !== 'All') {
    const aliases = {
      'Data': ['data', 'data cloud'],
      'Sales': ['sales', 'sales cloud'],
      'Service': ['service', 'service cloud'],
      'Industry': ['industry', 'industry cloud']
    };
    const selected = filterValue; // Keep original case
    const targets = aliases[selected] || [selected.toLowerCase()];

    console.log(`Filtering with value: "${filterValue}", targets: ${targets}`);

    filteredRecords = ghoRecords.filter(record => {
      // Both modes should have CaseRoutingTaxonomy__r.Name after hydration
      const taxonomyName = record.CaseRoutingTaxonomy__r?.Name || '';
      const rawName = taxonomyName.toLowerCase().trim();

      if (!rawName) {
        console.log('No taxonomy name found for record:', record.Id, record);
        return false;
      }

      const group = rawName.split(/[\-–—]/)[0].trim();
      const matches = targets.includes(group);

      if (filterValue !== 'All') {
        console.log(`Filtering record ${record.Id}: taxonomy="${taxonomyName}", group="${group}", targets="${targets}", matches=${matches}`);
      }

      return matches;
    });

    console.log(`Filtered ${ghoRecords.length} records to ${filteredRecords.length} records`);
  }

  if (filteredRecords.length === 0) {
    const displayFilter = filterValue === 'Data' ? 'Data Cloud' : filterValue;
    const isAll = filterValue === 'All';
    container.innerHTML = `
      <div class="gho-empty-card" style="background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%); border: 1.5px solid #bae6fd; border-radius: 16px; padding: 24px; text-align: center; animation: fadeUp 320ms ease; box-shadow: 0 6px 20px rgba(2,132,199,0.06);">
        <div style="font-size: 26px; margin-bottom: 8px;">🧭</div>
        <h4 style="margin: 0 0 6px; font-size: 17px; color: #075985;">No ${isAll ? '' : displayFilter + ' '}GHO cases</h4>
        <p style="margin: 0; color: #0c4a6e; font-size: 14px;">No cases matching ${isAll ? 'GHO criteria' : displayFilter + ' taxonomy'} for <strong>${currentShift}</strong> shift.</p>
        <div style="display:flex; justify-content:center; gap:8px; margin-top:12px; flex-wrap: wrap;">
          <span style="background:#e0f2fe; color:#075985; border:1px solid #bae6fd; padding:4px 10px; border-radius:999px; font-size:12px;">Shift: ${currentShift}</span>
          <span style="background:#e0f2fe; color:#075985; border:1px solid #bae6fd; padding:4px 10px; border-radius:999px; font-size:12px;">Filter: ${isAll ? 'All' : displayFilter}</span>
          <span style="background:${currentMode === 'premier' ? '#d1fae5' : '#e0e7ff'}; color:${currentMode === 'premier' ? '#065f46' : '#3730a3'}; border:1px solid ${currentMode === 'premier' ? '#6ee7b7' : '#c7d2fe'}; padding:4px 10px; border-radius:999px; font-size:12px;">Mode: ${currentMode === 'premier' ? 'Premier' : 'Signature'}</span>
        </div>
        <div style="margin-top:14px; display:flex; justify-content:center; gap:10px; flex-wrap: wrap;">
          ${isAll ? '' : '<button id="gho-empty-showall" class="preview-btn" style="background:#10b981; padding:8px 14px;">Show All</button>'}
          <button id="gho-empty-refresh" class="preview-btn" style="background:#0ea5e9; padding:8px 14px;">Refresh</button>
        </div>
      </div>
    `;
    const refreshBtn = document.getElementById('gho-empty-refresh');
    if (refreshBtn) {
      refreshBtn.addEventListener('click', () => {
        checkGHOStatus();
      });
    }
    const showAllBtn = document.getElementById('gho-empty-showall');
    if (showAllBtn) {
      showAllBtn.addEventListener('click', () => {
        const dd = document.getElementById('gho-taxonomy-filter');
        if (dd) dd.value = 'All';
        renderFilteredGHOCases(ghoRecords, conn, 'All');
      });
    }
    return;
  }

  // If we have precomputed triage ids (from cache or upstream), re-use; else perform a comment query
  if (pre && pre.triageCaseIds instanceof Set) {
    console.log('Using precomputed cache data for rendering:', {
      triageCaseIds: pre.triageCaseIds?.size || 0,
      mondayCaseIds: pre.mondayCaseIds?.size || 0,
      mondayCaseIdsContent: Array.from(pre.mondayCaseIds || []),
      noWocTemplateCaseIds: pre.noWocTemplateCaseIds?.size || 0,
      noWocTemplateCaseIdsContent: Array.from(pre.noWocTemplateCaseIds || [])
    });
    const weekendNow = isCurrentlyWeekend();
    const weekendMondaySet = weekendNow ? (pre.mondayCaseIds || new Set()) : new Set();
    const weekendNoWocSet = weekendNow ? (pre.noWocTemplateCaseIds || new Set()) : new Set();
    renderGHOCasesWithCommentInfo(filteredRecords, conn, currentShift, filterValue, pre.triageCaseIds, weekendMondaySet, weekendNoWocSet);
    return;
  }

  const caseIds = filteredRecords.map(record => record.Id);
  if (caseIds.length === 0) {
    renderGHOCasesWithCommentInfo(filteredRecords, conn, currentShift, filterValue, new Set(), new Set(), new Set());
    return;
  }
  // Query comments for all cases, we'll limit to 5 per case in JavaScript
  const commentQuery = `SELECT ParentId, Body, CreatedById, LastModifiedDate, CreatedDate FROM CaseFeed WHERE Visibility = 'InternalUsers' AND ParentId IN ('${caseIds.join("','")}') AND Type = 'TextPost' ORDER BY CreatedDate DESC`;
  conn.query(commentQuery, function (commentErr, commentResult) {
    if (commentErr) {
      console.error('Error querying GHO comments:', commentErr);
      renderGHOCasesWithCommentInfo(filteredRecords, conn, currentShift, filterValue, new Set(), new Set(), new Set());
      return;
    }
    const ghoTriageCommentCases = new Set();
    const ghoMondayCases = new Set();
    const ghoNoWocTemplateCases = new Set();

    if (commentResult.records) {
      // Group comments by case ID and limit to 5 most recent per case
      const commentsByCase = {};
      commentResult.records.forEach(comment => {
        if (!commentsByCase[comment.ParentId]) {
          commentsByCase[comment.ParentId] = [];
        }
        // Only keep the 5 most recent comments per case
        if (commentsByCase[comment.ParentId].length < 5) {
          commentsByCase[comment.ParentId].push(comment);
        }
      });

      // Check each case for GHO triage and Monday mentions
      const isWeekendNow = isCurrentlyWeekend();
      Object.keys(commentsByCase).forEach(caseId => {
        const caseComments = commentsByCase[caseId];

        // Check for GHO triage comments
        caseComments.forEach(comment => {
          if (comment.Body && comment.Body.includes('#GHOTriage')) {
            const commentDate = comment.LastModifiedDate || comment.CreatedDate;
            const isCommentFromToday = isToday(commentDate);
            const commentShift = getShiftForDate(commentDate);
            const isSameGeo = commentShift === currentShift;
            if (isCommentFromToday && isSameGeo) ghoTriageCommentCases.add(comment.ParentId);
          }
        });

        // Check for Monday mentions in the 5 most recent comments - weekend only
        const hasMondayInComments = isWeekendNow && caseComments.some(comment => {
          if (!comment.Body) return false;
          // Strip HTML tags and check for Monday (case insensitive)
          const cleanBody = comment.Body.replace(/<[^>]*>/g, '').toLowerCase();
          // Check for "Monday" or "TO: {Any word} Monday" pattern
          const hasMonday = cleanBody.includes('monday') ||
            /to:\s*\w+\s+monday/i.test(cleanBody);
          console.log(`Comment check for case ${caseId}: Original="${comment.Body}", Cleaned="${cleanBody}", HasMonday=${hasMonday}`);
          return hasMonday;
        });
        if (hasMondayInComments) {
          ghoMondayCases.add(caseId);
          console.log(`GHO Monday case detected (filtered): ${caseId} - found "monday" in 5 latest comments`);
        } else {
          console.log(`No Monday detected for case ${caseId} - comments:`, caseComments.map(c => c.Body));
        }

        // Check for "No {currentShift} GHO template" mentions in the 5 most recent comments - weekend only
        const hasNoWocTemplateInComments = isWeekendNow && caseComments.some(comment => {
          if (!comment.Body) return false;
          // Strip HTML tags and check for "No {currentShift} GHO template" pattern (case insensitive)
          const cleanBody = comment.Body.replace(/<[^>]*>/g, '').toLowerCase();
          const currentShiftLower = currentShift.toLowerCase();
          // More flexible pattern: "no" + any case shift + "gho" + "template" with word boundaries
          const noWocPattern = new RegExp(`\\bno\\s+${currentShiftLower}\\s+gho\\s+template\\b`, 'i');
          const hasNoWocTemplate = noWocPattern.test(cleanBody);
          console.log(`No WOC Template check for case ${caseId}: Original="${comment.Body}", Cleaned="${cleanBody}", HasNoWocTemplate=${hasNoWocTemplate}, Pattern="${noWocPattern.source}", CurrentShift="${currentShift}"`);
          return hasNoWocTemplate;
        });

        if (hasNoWocTemplateInComments) {
          ghoNoWocTemplateCases.add(caseId);
          console.log(`No WOC Template case detected (filtered): ${caseId} - found "No ${currentShift} GHO template" in 5 latest comments`);
        } else {
          console.log(`No WOC Template not detected for case ${caseId} - comments:`, caseComments.map(c => c.Body));
        }
      });
    }
    renderGHOCasesWithCommentInfo(filteredRecords, conn, currentShift, filterValue, ghoTriageCommentCases, ghoMondayCases, ghoNoWocTemplateCases);
  });
}

// Helper function to render GHO cases with comment information
function renderGHOCasesWithCommentInfo(filteredRecords, conn, currentShift, filterValue, ghoTriageCommentCases, ghoMondayCases = new Set(), ghoNoWocTemplateCases = new Set()) {
  const container = document.getElementById('gho-cases-container');

  // Collect all unique user IDs from routing logs
  const allUserIds = new Set();
  filteredRecords.forEach(caseRecord => {
    if (caseRecord.Case_Routing_Logs__r && caseRecord.Case_Routing_Logs__r.records) {
      caseRecord.Case_Routing_Logs__r.records.forEach(log => {
        if (log.CreatedById) {
          allUserIds.add(log.CreatedById);
        }
      });
    }
  });

  // If cache already holds userMap for same shift reuse it to skip User query
  const perModeCache = window.__ghoCache[currentMode];
  const tryCacheUserMap = (perModeCache && perModeCache.shift === currentShift && perModeCache.userMap && Object.keys(perModeCache.userMap).length > 0)
    ? perModeCache.userMap : null;

  const proceed = (userMap) => {
    // persist userMap into cache for future renders
    try {
      if (window.__ghoCache[currentMode] && window.__ghoCache[currentMode].shift === currentShift) {
        window.__ghoCache[currentMode].userMap = userMap;
        // Save updated cache to persistent storage
        saveGHOCache();
      }
      // also persist to localStorage day-cache
      const existing = loadUserMapCache();
      saveUserMapCache({ ...(existing.userMap || {}), ...(userMap || {}) });
    } catch (e) { }

    // Count cases with QB mentions
    const qbMentionedCount = Array.from(ghoTriageCommentCases).length;

    // Count cases with suggested assignees
    const suggestedAssigneeCount = filteredRecords.filter(record =>
      getSuggestedAssignee(record, userMap)
    ).length;

    // Get mappings for display
    const mappings = loadGHOMappings();

    const headerHtml = `
      <div class="gho-summary-modern" style="animation: fadeUp 260ms ease;">
        <div class="gho-summary-header">
          <div class="gho-summary-icon">
            <i class="fa fa-list-alt" style="font-size: 24px; color: #6366f1;"></i>
          </div>
          <div class="gho-summary-title-section">
            <h4 class="gho-summary-title">${filteredRecords.length} ${filterValue === 'All' ? '' : filterValue + ' '}GHO Case${filteredRecords.length === 1 ? '' : 's'} Found</h4>
            <p class="gho-summary-subtitle">Cases matching ${filterValue === 'All' ? 'GHO criteria' : filterValue + ' taxonomy'} for <strong>${currentShift}</strong> shift</p>
          </div>
          <div class="gho-summary-stats">
            <div class="gho-stat-item">
              <span class="gho-stat-number">${filteredRecords.length}</span>
              <span class="gho-stat-label">Total</span>
            </div>
          </div>
        </div>
        
        <div class="gho-summary-metrics">
          ${qbMentionedCount > 0 ? `
            <div class="gho-metric-card gho-metric-card--success">
              <div class="gho-metric-icon">
                <i class="fa fa-check-circle" style="color: #10b981;"></i>
              </div>
              <div class="gho-metric-content">
                <span class="gho-metric-number">${qbMentionedCount}</span>
                <span class="gho-metric-label">QB Mentioned (#GHOTriage)</span>
              </div>
            </div>
          ` : ''}
          
          ${ghoMondayCases.size > 0 ? `
            <div class="gho-metric-card gho-metric-card--monday gho-metric-card--compact">
              <div class="gho-metric-icon gho-metric-icon--small">
                <i class="fa fa-calendar" style="color: #10b981; font-size: 14px;"></i>
              </div>
              <div class="gho-metric-content">
                <span class="gho-metric-number gho-metric-number--small">${ghoMondayCases.size}</span>
                <span class="gho-metric-label gho-metric-label--small">Monday</span>
              </div>
            </div>
          ` : ''}
          
          ${ghoNoWocTemplateCases.size > 0 ? `
            <div class="gho-metric-card gho-metric-card--woc">
              <div class="gho-metric-icon">
                <i class="fa fa-file-text-o" style="color: #f59e0b;"></i>
              </div>
              <div class="gho-metric-content">
                <span class="gho-metric-number">${ghoNoWocTemplateCases.size}</span>
                <span class="gho-metric-label">No WOC Template</span>
              </div>
            </div>
          ` : ''}
          
          ${suggestedAssigneeCount > 0 ? `
            <div class="gho-metric-card gho-metric-card--info">
              <div class="gho-metric-icon">
                <i class="fa fa-lightbulb-o" style="color: #3b82f6;"></i>
              </div>
              <div class="gho-metric-content">
                <span class="gho-metric-number">${suggestedAssigneeCount}</span>
                <span class="gho-metric-label">Suggested Assignee${suggestedAssigneeCount > 1 ? 's' : ''}</span>
              </div>
            </div>
          ` : ''}
        </div>
        
        <div class="gho-summary-info-grid">
          <div class="gho-info-card gho-info-card--shift">
            <div class="gho-info-icon">
              <i class="fa fa-clock-o" style="color: #f59e0b;"></i>
            </div>
            <div class="gho-info-content">
              <span class="gho-info-label">Current Shift</span>
              <span class="gho-info-value">${currentShift}</span>
            </div>
          </div>
          
          <div class="gho-info-card gho-info-card--time">
            <div class="gho-info-icon">
              <i class="fa fa-calendar-check-o" style="color: #10b981;"></i>
            </div>
            <div class="gho-info-content">
              <span class="gho-info-label">Current Time</span>
              <span class="gho-info-value">${new Date().toLocaleTimeString()}</span>
            </div>
          </div>
          
          <div class="gho-info-card gho-info-card--mappings">
            <div class="gho-info-icon">
              <i class="fa fa-exchange" style="color: #6366f1;"></i>
            </div>
            <div class="gho-info-content">
              <span class="gho-info-label">Active Mappings</span>
              <span class="gho-info-value">${mappings.length}</span>
            </div>
          </div>
        </div>
        

        
        ${mappings.length > 0 ? `
          <div class="gho-summary-mapping-status gho-summary-mapping-status--active">
            <div class="gho-mapping-status-icon">
              <i class="fa fa-check-circle" style="color: #10b981;"></i>
            </div>
            <div class="gho-mapping-status-content">
              <span class="gho-mapping-status-title">Mapping System Active</span>
              <span class="gho-mapping-status-description">
                ${suggestedAssigneeCount > 0 ? `${suggestedAssigneeCount} case${suggestedAssigneeCount > 1 ? 's' : ''} have suggested assignee${suggestedAssigneeCount > 1 ? 's' : ''}` : 'No cases currently match your mappings'}
              </span>
              <a href="#" class="gho-mapping-status-link mapping-link" data-action="view-mappings">
                <i class="fa fa-cog" style="margin-right: 6px;"></i>
                View/Edit Mappings
              </a>
            </div>
          </div>
        ` : `
          <div class="gho-summary-mapping-status gho-summary-mapping-status--inactive">
            <div class="gho-mapping-status-icon">
              <i class="fa fa-lightbulb-o" style="color: #f59e0b;"></i>
            </div>
            <div class="gho-mapping-status-content">
              <span class="gho-mapping-status-title">No Mappings Configured</span>
              <span class="gho-mapping-status-description">
                Set up GHO transfer mappings to get suggested assignees for cases
              </span>
              <a href="#" class="gho-mapping-status-link mapping-link" data-action="view-mappings">
                <i class="fa fa-plus" style="margin-right: 6px;"></i>
                Configure Mappings
              </a>
            </div>
          </div>
        `}
      </div>
      <div id="gho-items"></div>
    `;

    container.innerHTML = headerHtml;

    // Initialize list state for lazy loading
    window.__ghoListState = {
      allRecords: filteredRecords.slice(),
      renderedCount: 0,
      total: filteredRecords.length,
      userMap,
      triageCaseIds: ghoTriageCommentCases,
      mondayCaseIds: ghoMondayCases,
      noWocTemplateCaseIds: ghoNoWocTemplateCases,
      currentShift,
      filterValue,
      conn
    };

    const ensureFillViewport = () => {
      // If content doesn't fill, load more until it does or items exhausted
      let safety = 0;
      while (safety < 5 && window.__ghoListState.renderedCount < window.__ghoListState.total) {
        const loaded = loadMoreGhoCases();
        safety++;
        const items = document.getElementById('gho-items');
        if (!items || !container) break;
        if (container.scrollHeight > container.clientHeight + 200) break;
        if (!loaded) break;
      }
    };

    // Render first page and set up scroll
    loadMoreGhoCases();
    setupGhoInfiniteScroll();
    // Try to fill viewport for small datasets
    ensureFillViewport();
  };

  if (tryCacheUserMap) {
    proceed(tryCacheUserMap);
  } else {
    getUserNames(Array.from(allUserIds), conn, proceed);
  }
}

function setupGhoInfiniteScroll() {
  try {
    const modalBody = document.querySelector('#gho-modal .modal-body');
    if (!modalBody) return;
    // Cleanup prior listener
    if (modalBody.__ghoScrollHandler) {
      modalBody.removeEventListener('scroll', modalBody.__ghoScrollHandler);
    }
    const handler = () => {
      const nearBottom = modalBody.scrollTop + modalBody.clientHeight >= modalBody.scrollHeight - 200;
      if (nearBottom) loadMoreGhoCases();
    };
    modalBody.__ghoScrollHandler = handler;
    modalBody.addEventListener('scroll', handler, { passive: true });
  } catch (e) { console.warn('setupGhoInfiniteScroll failed', e); }
}

function renderSingleGhoCaseHTML(caseRecord, userMap, ghoTriageCommentCases, currentShift, ghoMondayCases = new Set(), ghoNoWocTemplateCases = new Set()) {
  const caseId = caseRecord.Id;
  const isMVP = caseRecord.Contact && caseRecord.Contact.Is_MVP__c === true;
  const hasGHOTriage = ghoTriageCommentCases.has(caseId);
  const hasMondayMention = ghoMondayCases.has(caseId);
  const hasNoWocTemplate = ghoNoWocTemplateCases.has(caseId);

  // Get suggested assignee based on mapping
  const suggestedAssignee = getSuggestedAssignee(caseRecord, userMap);

  // Build routing log HTML and GHO transfer information
  let routingLogHtml = '';
  let ghoTransferHtml = '';
  const routingLogs = caseRecord.Case_Routing_Logs__r;

  if (routingLogs && routingLogs.totalSize > 0) {
    const lastLog = routingLogs.records[0];
    if (lastLog.Transfer_Reason__c && lastLog.Transfer_Reason__c !== 'New') {
      let logText = lastLog.Transfer_Reason__c;
      if (lastLog.Severity_Old_Value__c && lastLog.Severity_New_Value__c) {
        logText += ` | Severity Changed: ${lastLog.Severity_Old_Value__c} → ${lastLog.Severity_New_Value__c}`;
      }
      routingLogHtml = `
        <div class="case-info-item">
          <span class="checkmark">✓</span>
          <span style="color: #9F2B68;">${logText} (${timeElapsed(new Date(lastLog.CreatedDate))})</span>
        </div>
      `;
    }

    const allGhoTransfers = routingLogs.records.filter(log =>
      log.Transfer_Reason__c === 'GHO'
    ).sort((a, b) => new Date(b.CreatedDate) - new Date(a.CreatedDate));

    if (allGhoTransfers.length > 0) {
      const currentShiftTransfers = allGhoTransfers.filter(transfer =>
        transfer.Preferred_Shift_Old_Value__c === currentShift
      );
      const otherTransfers = allGhoTransfers.filter(transfer =>
        transfer.Preferred_Shift_Old_Value__c !== currentShift
      );

      const caseUniqueId = `gho-${caseRecord.Id}`;
      let nestedListHtml = '';

      if (currentShiftTransfers.length > 0) {
        currentShiftTransfers.forEach((transfer) => {
          const userName = userMap[transfer.CreatedById] || 'Unknown User';
          const transferTime = new Date(transfer.CreatedDate);
          const ghoFrom = transfer.Preferred_Shift_Old_Value__c || 'N/A';
          const ghoTo = transfer.Preferred_Shift_New_Value__c || 'N/A';

          // Check if this transfer triggers a mapping
          const mappings = loadGHOMappings();
          const mapping = mappings.find(m => m.from.toLowerCase() === userName.toLowerCase());
          const mappingBadge = mapping ? `
            <span style="background: #dcfce7; color: #166534; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 600; margin-left: 8px;">
              <i class="fa fa-exchange" style="margin-right: 2px;"></i>Maps to ${mapping.to}
            </span>
          ` : '';

          nestedListHtml += `
            <div class="gho-transfer-item gho-transfer-item--current">
              <div class="gho-transfer-title">
                <span class="name" style="color:#92400e;">${userName}</span>
                <span class="time" style="color:#78350f;">${formatDateWithDayOfWeek(transferTime)}</span>
              </div>
              <div class="gho-transfer-fields" style="color:#78350f;">
                <span><strong>GHO FROM:</strong> ${ghoFrom}</span>
                <span><strong>GHO TO:</strong> ${ghoTo}</span>
              </div>
              <div class="gho-transfer-meta">
                ${timeElapsed(transferTime)} ago
                ${mappingBadge}
              </div>
            </div>
          `;
        });
      }

      if (otherTransfers.length > 0) {
        nestedListHtml += `
          <div style="margin: 8px 0;">
            <button class="gho-toggle-btn" data-case-id="${caseUniqueId}">
              <i class="fa fa-chevron-down" aria-hidden="true" style="font-size: 10px; color: #64748b; transition: transform 0.2s ease;"></i>
              <span id="${caseUniqueId}-toggle-text" style="font-weight: 500;">Show ${otherTransfers.length} more GHO transfer${otherTransfers.length > 1 ? 's' : ''}</span>
            </button>
          </div>
          <div id="${caseUniqueId}-expanded" style="display: none;">
        `;

        otherTransfers.forEach(transfer => {
          const userName = userMap[transfer.CreatedById] || 'Unknown User';
          const transferTime = new Date(transfer.CreatedDate);
          const ghoFrom = transfer.Preferred_Shift_Old_Value__c || 'N/A';
          const ghoTo = transfer.Preferred_Shift_New_Value__c || 'N/A';

          nestedListHtml += `
            <div class="gho-transfer-item gho-transfer-item--other">
              <div class="gho-transfer-title">
                <span class="name">${userName}</span>
                <span class="time">${formatDateWithDayOfWeek(transferTime)}</span>
              </div>
              <div class="gho-transfer-fields">
                <span><strong>GHO FROM:</strong> ${ghoFrom}</span>
                <span><strong>GHO TO:</strong> ${ghoTo}</span>
              </div>
              <div class="gho-transfer-meta">${timeElapsed(transferTime)} ago</div>
            </div>
          `;
        });

        nestedListHtml += `</div>`;
      }

      ghoTransferHtml = `
        <div class="case-info-item" style="flex-direction: column; align-items: flex-start;">
          <div class="gho-transfer-heading">
            <span class="checkmark">✓</span>
            <span>GHO Transfer History (${allGhoTransfers.length} transfer${allGhoTransfers.length > 1 ? 's' : ''}):</span>
            ${suggestedAssignee ? `
            <div style="margin-left: 20px; font-size: 12px; color: #166534; font-weight: 600;">
              <i class="fa fa-lightbulb-o" style="margin-right: 4px;"></i>
              Next suggested: ${suggestedAssignee}
            </div>
            ` : ''}
          </div>
          <div class="gho-transfer-section">
            ${nestedListHtml}
          </div>
        </div>
      `;
    } else {
      ghoTransferHtml = `
        <div class="case-info-item">
          <span class="checkmark">✓</span>
          <span class="gho-empty-transfer">No GHO transfers found</span>
        </div>
      `;
    }
  }


  return `
    <div class="case-card gho-case-card ${isMVP ? 'mvp-case card-accent-purple' : ''}" data-case-id="${caseId}" style="margin-bottom:16px;">
      <div class="gho-card-header">
        <div class="gho-card-header-main">
          <div class="gho-card-badges">
            <span class="badge-soft badge-soft--amber">GHO</span>
            ${hasMondayMention ? '<span class="badge-soft badge-soft--monday">GHO Monday</span>' : ''}
            ${hasNoWocTemplate ? '<span class="badge-soft badge-soft--woc">No WOC Template</span>' : ''}
            ${isMVP ? '<span class="badge-soft badge-soft--purple">MVP</span>' : ''}
            ${hasGHOTriage ? '<span class="badge-soft badge-soft--success">QB Mentioned</span>' : ''}
          </div>
          <h3 class="gho-card-subject">${caseRecord.Subject}</h3>
        </div>
        <div class="gho-card-created">${formatDateWithDayOfWeek(caseRecord.CreatedDate)}<br>(${timeElapsed(new Date(caseRecord.CreatedDate))})</div>
      </div>
      <div class="gho-meta-grid">
      <div class="gho-meta-item"><span class="gho-meta-label">Case #</span><span class="gho-meta-value">${caseRecord.CaseNumber}</span></div>
        <div class="gho-meta-item"><span class="gho-meta-label">Account</span><span class="gho-meta-value">${caseRecord.Account.Name}</span></div>
        <div class="gho-meta-item"><span class="gho-meta-label">Owner</span><span class="gho-meta-value">${caseRecord.Owner.Name}</span></div>
        <div class="gho-meta-item"><span class="gho-meta-label">Cloud</span><span class="gho-meta-value">${caseRecord.CaseRoutingTaxonomy__r.Name}</span></div>
        ${suggestedAssignee ? `
        <div class="gho-suggested-assignee" style="grid-column: 1 / -1;" title="Based on GHO transfer mapping: ${getMappingSource(caseRecord, userMap)} → ${suggestedAssignee}">
          <div class="suggested-label">Suggested Assignee</div>
          <div class="suggested-name">${suggestedAssignee}</div>
        </div>
        ` : ''}
        <div class="gho-meta-item--severity-row">
          <div class="gho-severity-wrap">
            <span class="gho-meta-label">Severity</span>
            <span class="gho-meta-value">${caseRecord.Severity_Level__c}</span>
          </div>
          <a target="_blank" href="https://orgcs.my.salesforce.com/lightning/r/Case/${caseId}/view" class="preview-btn gho-preview-btn" data-case-id="${caseId}">View Case Record</a>
        </div>
      </div>
      ${routingLogHtml ? `<div class="gho-routing-log">${routingLogHtml}</div>` : ''}
      <div class="gho-transfer-block">${ghoTransferHtml}</div>
    </div>`;
}

function loadMoreGhoCases() {
  try {
    const state = window.__ghoListState;
    if (!state) return false;
    if (state.renderedCount >= state.total) return false;
    const listEl = document.getElementById('gho-items');
    if (!listEl) return false;

    const nextEnd = Math.min(state.renderedCount + GHO_PAGE_SIZE, state.total);
    let buf = '';
    for (let i = state.renderedCount; i < nextEnd; i++) {
      buf += renderSingleGhoCaseHTML(state.allRecords[i], state.userMap, state.triageCaseIds, state.currentShift, state.mondayCaseIds, state.noWocTemplateCaseIds);
    }
    listEl.insertAdjacentHTML('beforeend', buf);
    state.renderedCount = nextEnd;
    return true;
  } catch (e) {
    console.warn('loadMoreGhoCases failed', e);
    return false;
  }
}

// Function to toggle GHO transfer expansion
function toggleGhoTransfers(caseUniqueId) {
  try {
    const expandedDiv = document.getElementById(`${caseUniqueId}-expanded`);
    const toggleText = document.getElementById(`${caseUniqueId}-toggle-text`);
    const toggleButton = document.querySelector(`[data-case-id="${caseUniqueId}"]`);
    const chevronIcon = toggleButton ? toggleButton.querySelector('.fa-chevron-down') : null;

    if (expandedDiv && toggleText && chevronIcon) {
      if (expandedDiv.style.display === 'none' || expandedDiv.style.display === '') {
        expandedDiv.style.display = 'block';
        toggleText.textContent = 'Hide Transfer Logs';
        chevronIcon.style.transform = 'rotate(180deg)';
      } else {
        expandedDiv.style.display = 'none';
        toggleText.textContent = 'Show Transfer Logs';
        chevronIcon.style.transform = 'rotate(0deg)';
      }
    }
  } catch (error) {
    console.error('Error toggling GHO transfers:', error);
  }
}

// (debug functions removed for production)

// Function to clear GHO cache for a specific mode
function clearGHOCacheForMode(mode) {
  try {
    if (!mode || !['signature', 'premier'].includes(mode)) {
      console.warn('Invalid mode specified for cache clearing:', mode);
      return false;
    }

    // Clear in-memory cache for specific mode
    if (window.__ghoCache && window.__ghoCache[mode]) {
      delete window.__ghoCache[mode];
      console.log(`GHO cache cleared for ${mode} mode`);
    }

    // Update persistent cache (remove the specific mode)
    saveGHOCache();

    // Clear global variables if clearing current mode
    if (mode === currentMode) {
      ghoRecordsGlobal = [];
      ghoConnectionGlobal = null;
      window.__ghoListState = null;
    }

    showToast(`${mode.charAt(0).toUpperCase() + mode.slice(1)} mode GHO cache cleared`, 'info');
    return true;
  } catch (error) {
    console.error(`Failed to clear GHO cache for ${mode} mode:`, error);
    showToast(`Failed to clear ${mode} mode cache`, 'error');
    return false;
  }
}

// Function to clear all GHO cache
// (debug cache clear function removed for production)

// Function to manually save current cache to persistent storage
// (debug cache save function removed for production)

// Function to get comprehensive cache statistics for both modes
// (debug cache stats function removed for production)

// Function to show cache debugging information
// (debug cache inspection functions removed for production)

// Debug helper: open the GHO Alert modal with mock data
// (debug alert helper removed for production)

// Add a quick cache status function for debugging
// (debug cache status function removed for production)

// Test function to verify GHO Monday cache
// (debug test helper removed for production)

// Test function to verify No WOC Template pattern matching
// (debug test helper removed for production)

// Test function for the specific EMEA example
// (debug test helper removed for production)

// GHO Mapping Functions
function loadGHOMappings() {
  try {
    const raw = localStorage.getItem(GHO_MAPPING_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    console.warn('Failed to load GHO mappings:', error);
    return [];
  }
}

// (initializeDefaultMappings removed as unused)

function saveGHOMappings(mappings) {
  try {
    localStorage.setItem(GHO_MAPPING_STORAGE_KEY, JSON.stringify(mappings));
    return true;
  } catch (error) {
    console.error('Failed to save GHO mappings:', error);
    return false;
  }
}

function getSuggestedAssignee(caseRecord, userMap) {
  try {
    if (!caseRecord.Case_Routing_Logs__r || caseRecord.Case_Routing_Logs__r.records.length === 0) {
      return null;
    }

    // Find the latest GHO transfer
    const ghoTransfers = caseRecord.Case_Routing_Logs__r.records.filter(log =>
      log.Transfer_Reason__c === 'GHO'
    ).sort((a, b) => new Date(b.CreatedDate) - new Date(a.CreatedDate));

    if (ghoTransfers.length === 0) {
      return null;
    }

    const latestTransfer = ghoTransfers[0];
    const transferUserId = latestTransfer.CreatedById;
    const transferUserName = userMap[transferUserId];

    if (!transferUserName) {
      return null;
    }

    // Load mappings and find a match
    const mappings = loadGHOMappings();
    const mapping = mappings.find(m =>
      m.from.toLowerCase() === transferUserName.toLowerCase()
    );

    if (mapping) {
      return mapping.to;
    }

    return null;
  } catch (error) {
    console.error('Error getting suggested assignee:', error);
    return null;
  }
}

function getMappingSource(caseRecord, userMap) {
  try {
    if (!caseRecord.Case_Routing_Logs__r || caseRecord.Case_Routing_Logs__r.records.length === 0) {
      return null;
    }

    // Find the latest GHO transfer
    const ghoTransfers = caseRecord.Case_Routing_Logs__r.records.filter(log =>
      log.Transfer_Reason__c === 'GHO'
    ).sort((a, b) => new Date(b.CreatedDate) - new Date(a.CreatedDate));

    if (ghoTransfers.length === 0) {
      return null;
    }

    const latestTransfer = ghoTransfers[0];
    const transferUserId = latestTransfer.CreatedById;
    const transferUserName = userMap[transferUserId];

    return transferUserName || 'Unknown User';
  } catch (error) {
    console.error('Error getting mapping source:', error);
    return null;
  }
}

// ================================================
// Section: 14) GHO Mapping Modal (view/edit mappings)
// ================================================
function showGHOMappingModal() {
  const modal = document.getElementById('gho-mapping-modal');
  if (!modal) return;

  modal.style.display = 'flex';
  requestAnimationFrame(() => modal.classList.add('modal-show'));

  // Load and display existing mappings
  renderMappingList();
}

function renderMappingList() {
  const mappingList = document.getElementById('mapping-list');
  if (!mappingList) return;

  const mappings = loadGHOMappings();

  if (mappings.length === 0) {
    mappingList.innerHTML = `
      <div style="text-align: center; padding: 20px; color: #6b7280;">
        <i class="fa fa-info-circle" style="font-size: 24px; margin-bottom: 8px; display: block;"></i>
        <p>No mappings configured yet.</p>
        <p style="font-size: 12px;">Click "Add New Mapping" to get started.</p>
      </div>
    `;
    return;
  }

  mappingList.innerHTML = mappings.map((mapping, index) => `
      <div class="mapping-item" style="display: flex; align-items: center; gap: 12px; padding: 12px; border: 1px solid #e5e7eb; border-radius: 8px; margin-bottom: 8px; background: #f9fafb;">
        <div style="flex: 1; display: flex; align-items: center; gap: 8px;">
          <span style="font-weight: 600; color: #374151;">${mapping.from}</span>
          <i class="fa fa-arrow-right" style="color: #6b7280; font-size: 12px;"></i>
          <span style="font-weight: 600; color: #374151;">${mapping.to}</span>
        </div>
        <button class="btn-secondary delete-mapping-btn" data-index="${index}" style="padding: 6px 10px; font-size: 11px; background: #fee2e2; color: #dc2626; border-color: #fecaca;">
          <i class="fa fa-trash" aria-hidden="true"></i>
        </button>
      </div>
    `).join('');
}

function addNewMapping() {
  const mappingList = document.getElementById('mapping-list');
  if (!mappingList) return;

  const newMappingHtml = `
    <div class="mapping-item new-mapping" style="display: flex; align-items: center; gap: 12px; padding: 12px; border: 2px solid #3b82f6; border-radius: 8px; margin-bottom: 8px; background: #eff6ff;">
      <div style="flex: 1; display: flex; align-items: center; gap: 8px;">
        <input type="text" class="mapping-from" placeholder="From user name" 
               style="flex: 1; padding: 6px 8px; border: 1px solid #d1d5db; border-radius: 4px; font-size: 12px;">
        <i class="fa fa-arrow-right" style="color: #3b82f6; font-size: 12px;"></i>
        <input type="text" class="mapping-to" placeholder="To user name" 
               style="flex: 1; padding: 6px 8px; border: 1px solid #d1d5db; border-radius: 4px; font-size: 12px;">
      </div>
      <button class="btn-primary save-mapping-btn" style="padding: 6px 10px; font-size: 11px;">
        <i class="fa fa-check" aria-hidden="true"></i>
      </button>
      <button class="btn-secondary cancel-mapping-btn" style="padding: 6px 10px; font-size: 11px; background: #f3f4f6; color: #6b7280; border-color: #d1d5db;">
        <i class="fa fa-times" aria-hidden="true"></i>
      </button>
    </div>
  `;

  mappingList.insertAdjacentHTML('beforeend', newMappingHtml);

  // Focus on the first input
  const firstInput = mappingList.querySelector('.new-mapping .mapping-from');
  if (firstInput) {
    firstInput.focus();
  }
}

function saveNewMapping(button) {
  const mappingItem = button.closest('.mapping-item');
  const fromInput = mappingItem.querySelector('.mapping-from');
  const toInput = mappingItem.querySelector('.mapping-to');

  const from = fromInput.value.trim();
  const to = toInput.value.trim();

  if (!from || !to) {
    showToast('Please fill in both fields', 'warning');
    return;
  }

  const mappings = loadGHOMappings();

  // Check for duplicates
  if (mappings.some(m => m.from.toLowerCase() === from.toLowerCase())) {
    showToast(`Mapping for "${from}" already exists`, 'warning');
    return;
  }

  mappings.push({ from, to });

  if (saveGHOMappings(mappings)) {
    showToast(`Mapping added: ${from} → ${to}`, 'success');
    renderMappingList();
  } else {
    showToast('Failed to save mapping', 'error');
  }
}

function deleteMapping(index) {
  const mappings = loadGHOMappings();
  if (index >= 0 && index < mappings.length) {
    const deleted = mappings.splice(index, 1)[0];
    if (saveGHOMappings(mappings)) {
      showToast(`Mapping deleted: ${deleted.from} → ${deleted.to}`, 'success');
      renderMappingList();
    } else {
      showToast('Failed to delete mapping', 'error');
    }
  }
}

function cancelNewMapping(button) {
  const mappingItem = button.closest('.mapping-item');
  mappingItem.remove();
}

function deleteAllMappings() {
  const mappings = loadGHOMappings();

  if (mappings.length === 0) {
    showToast('No mappings to delete', 'info');
    return;
  }

  // Show confirmation dialog
  if (confirm(`Are you sure you want to delete all ${mappings.length} mapping${mappings.length > 1 ? 's' : ''}? This action cannot be undone.`)) {
    // Clear all mappings
    saveGHOMappings([]);

    // Refresh the mapping list display
    renderMappingList();

    // Update the mapping button count
    updateMappingButtonCount();

    // Show success message
    showToast(`All ${mappings.length} mapping${mappings.length > 1 ? 's' : ''} deleted successfully`, 'success');

    // All GHO mappings deleted
  }
}

// Functions are now handled through event delegation

function updateMappingButtonCount() {
  const mappingBtn = document.getElementById('gho-mapping-btn');
  if (!mappingBtn) return;

  const mappings = loadGHOMappings();
  const count = mappings.length;

  if (count > 0) {
    mappingBtn.innerHTML = `<i class="fa fa-exchange" aria-hidden="true"></i> Mapping (${count})`;
    mappingBtn.style.background = 'linear-gradient(180deg, #10b981 0%, #059669 100%)';
    mappingBtn.style.color = 'white';

    // Add tooltip with mapping preview
    const tooltipText = mappings.slice(0, 3).map(m => `${m.from} → ${m.to}`).join('\n');
    mappingBtn.title = `Active Mappings:\n${tooltipText}${mappings.length > 3 ? `\n... and ${mappings.length - 3} more` : ''}`;
  } else {
    mappingBtn.innerHTML = `<i class="fa fa-exchange" aria-hidden="true"></i> Mapping`;
    mappingBtn.style.background = '';
    mappingBtn.style.color = '';
    mappingBtn.title = 'Configure GHO transfer mappings';
  }
}

// Helper Fn: process CaseHistory and track 'New Case'

async function trackNewCaseFromHistory(conn, params) {
  const { caseIds, currentUserId, currentUserName, currentMode, strategy, removeFromPersistent } = params || {};
  try {
    if (!Array.isArray(caseIds) || caseIds.length === 0) return { processed: [] };

    // Fetch required case details including Owner
    const caseRes = await conn.query(`SELECT Id, CaseNumber, Severity_Level__c, CaseRoutingTaxonomy__r.Name, Owner.Name FROM Case WHERE Id IN ('${caseIds.join("','")}')`);
    const caseMap = new Map((caseRes.records || []).map(r => [r.Id, r]));

    const orderDir = strategy === 'firstManualByUser' ? 'ASC' : 'DESC';
    const histRes = await conn.query(`SELECT CaseId, CreatedById, CreatedDate, Field, NewValue FROM CaseHistory WHERE CaseId IN ('${caseIds.join("','")}') AND CreatedById='${currentUserId}' AND (Field='Routing_Status__c' OR Field='Owner') ORDER BY CreatedDate ${orderDir}`);
    const records = histRes.records || [];

    // Group by CaseId
    const byCase = new Map();
    for (const h of records) {
      if (!byCase.has(h.CaseId)) byCase.set(h.CaseId, []);
      byCase.get(h.CaseId).push(h);
    }

    const processed = [];
    for (const caseId of caseIds) {
      const hist = byCase.get(caseId) || [];
      // Choose candidate per strategy
      let candidate = null;
      for (const h of hist) {
        const isManualAssign = (h.Field === 'Routing_Status__c' && typeof h.NewValue === 'string' && h.NewValue.startsWith('Manually Assigned'));
        const isOwnerChange = (h.Field === 'Owner');
        if (strategy === 'firstManualByUser') {
          if (isManualAssign && h.CreatedById === currentUserId) { candidate = h; break; }
        } else {
          if ((isManualAssign || isOwnerChange) && h.CreatedById === currentUserId) { candidate = h; break; }
        }
      }
      if (!candidate) continue;

      const cRec = caseMap.get(caseId);
      if (!cRec) continue;

      const trackingKey = `tracked_${currentMode}_assignment_${caseId}`;
      if (localStorage.getItem(trackingKey)) continue;

      const cloud = (cRec.CaseRoutingTaxonomy__r && cRec.CaseRoutingTaxonomy__r.Name) ? cRec.CaseRoutingTaxonomy__r.Name.split('-')[0] : '';
      // Prefer latest Owner change with a human-readable name over IDs
      const isSfUserId = (v) => typeof v === 'string' && /^005[\w]{12,15}$/.test(v);
      const scanArr = strategy === 'firstManualByUser' ? hist.slice().reverse() : hist;
      let ownerNameFromHist = '';
      for (const oh of scanArr) {
        if (oh.Field === 'Owner' && oh.NewValue && !isSfUserId(oh.NewValue)) { ownerNameFromHist = oh.NewValue; break; }
      }
      const assignedTo = ownerNameFromHist || ((cRec.Owner && cRec.Owner.Name) ? cRec.Owner.Name : (candidate.NewValue || ''));

      trackActionAndCount(
        candidate.CreatedDate,
        cRec.CaseNumber,
        cRec.Severity_Level__c,
        'New Case',
        cloud,
        currentMode,
        currentUserName,
        assignedTo
      );
      localStorage.setItem(trackingKey, 'true');

      if (removeFromPersistent) {
        try { chrome.runtime.sendMessage({ action: 'removeCaseFromPersistentSet', caseId }, () => { }); } catch { }
      }

      processed.push({ caseId, action: 'New Case (history)', assignedTo });
    }
    return { processed };
  } catch (e) {
    console.warn('trackNewCaseFromHistory failed:', e);
    return { processed: [], error: e.message };
  }
}

// Detect if cases have been assigned by ANY user (no tracking/logging)
async function detectAssignmentsForCases(conn, params) {
  const { caseIds } = params || {};
  try {
    if (!Array.isArray(caseIds) || caseIds.length === 0) return { processed: [] };

    // Pull Case info for context/state checks
    const caseRes = await conn.query(`SELECT Id, CaseNumber, Owner.Name, Status, IsClosed FROM Case WHERE Id IN ('${caseIds.join("','")}')`);
    const caseMap = new Map((caseRes.records || []).map(r => [r.Id, r]));

    // Look for either routing status manual assignment OR owner changes by anyone in recent history (last 24h window)
    const sinceIso = new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString();
    const histRes = await conn.query(`SELECT CaseId, CreatedById, CreatedDate, Field, NewValue FROM CaseHistory WHERE CaseId IN ('${caseIds.join("','")}') AND CreatedDate >= ${sinceIso} AND (Field='Routing_Status__c' OR Field='Owner') ORDER BY CreatedDate DESC`);
    const records = histRes.records || [];
    const byCase = new Map();
    for (const h of records) {
      if (!byCase.has(h.CaseId)) byCase.set(h.CaseId, []);
      byCase.get(h.CaseId).push(h);
    }

    const processed = [];
    for (const caseId of caseIds) {
      const hist = byCase.get(caseId) || [];
      let assigned = false;
      let reason = '';

      // Check current state first (cheap & definitive)
      const cRec = caseMap.get(caseId);
      if (cRec) {
        const ownerName = (cRec.Owner && cRec.Owner.Name) ? String(cRec.Owner.Name) : '';
        const isQueueOwner = /Queue/i.test(ownerName) || ['Kase Changer', 'Working in Org62', 'Data Cloud Queue'].some(q => ownerName.includes(q));
        if (cRec.IsClosed === true) { assigned = true; reason = 'closed'; }
        else if (cRec.Status && cRec.Status !== 'New') { assigned = true; reason = 'status-changed'; }
        else if (!isQueueOwner && ownerName) { assigned = true; reason = 'owner-human'; }
      }

      // If still not decided, look at recent history changes by anyone
      for (const h of hist) {
        const isManualAssign = (h.Field === 'Routing_Status__c' && typeof h.NewValue === 'string' && h.NewValue.startsWith('Manually Assigned'));
        const isOwnerChange = (h.Field === 'Owner');
        if (isManualAssign || isOwnerChange) { assigned = true; reason = isManualAssign ? 'history-manual' : 'history-owner'; break; }
      }
      if (assigned) {
        processed.push({ caseId, action: 'assignment-detected', reason, owner: (cRec && cRec.Owner && cRec.Owner.Name) || '' });
      }
    }
    return { processed };
  } catch (e) {
    console.warn('detectAssignmentsForCases failed:', e);
    return { processed: [], error: e.message };
  }
}

// Force-process a single case by Case Number: checks manual assignments, QB mentions, and GHO triage
async function forceProcessCase(caseNumber) {
  try {
    if (!caseNumber || typeof caseNumber !== 'string') {
      showToast('Provide a valid Case Number');
      return { success: false, message: 'Invalid Case Number' };
    }
    if (!SESSION_ID) {
      showToast('Session not ready. Try again after data loads.');
      return { success: false, message: 'Missing SESSION_ID' };
    }

    const conn = new jsforce.Connection({
      serverUrl: 'https://orgcs.my.salesforce.com',
      sessionId: SESSION_ID,
    });

    // Resolve current user id
    let userId = null;
    try {
      const ures = await conn.query(`SELECT Id FROM User WHERE Name = '${currentUserName}' AND IsActive = True AND Username LIKE '%orgcs.com'`);
      if (ures.records && ures.records.length > 0) userId = ures.records[0].Id;
    } catch (e) {
      console.warn('Failed fetching user id:', e);
    }
    if (!userId) {
      showToast('Could not resolve current user');
      return { success: false, message: 'Missing user id' };
    }

    // Fetch the case by number
    const caseRes = await conn.query(`SELECT Id, CaseNumber, Severity_Level__c, Subject, CaseRoutingTaxonomy__r.Name, Owner.Name, LastModifiedDate FROM Case WHERE CaseNumber='${caseNumber}' LIMIT 5`);
    if (!caseRes.records || caseRes.records.length === 0) {
      showToast(`Case ${caseNumber} not found`);
      return { success: false, message: 'Case not found' };
    }
    const caseRec = caseRes.records[0];
    const caseId = caseRec.Id;

    const cloud = (caseRec.CaseRoutingTaxonomy__r && caseRec.CaseRoutingTaxonomy__r.Name) ? caseRec.CaseRoutingTaxonomy__r.Name.split('-')[0] : '';

    let actions = [];

    // History-based tracking via helper
    try {
      const res = await trackNewCaseFromHistory(conn, {
        caseIds: [caseId],
        currentUserId: userId,
        currentUserName,
        currentMode,
        strategy: 'latestByUser',
        removeFromPersistent: false
      });
      if (res && res.processed && res.processed.length > 0) actions.push('New Case (history)');
    } catch (e) { console.warn('History helper failed in forceProcessCase:', e); }

    // Check CaseFeed comments for #SigQBmention and #GHOTriage
    try {
      const cRes = await conn.query(`SELECT ParentId, Body, CreatedById, LastModifiedDate, CreatedDate FROM CaseFeed WHERE Visibility='InternalUsers' AND ParentId='${caseId}' AND Type='TextPost' ORDER BY CreatedDate DESC LIMIT 20`);
      const comments = cRes.records || [];
      for (const c of comments) {
        if (c.Body && c.Body.includes('#SigQBmention') && c.CreatedById === userId) {
          const trackingKey = `tracked_${caseId}`;
          if (!localStorage.getItem(trackingKey)) {
            trackActionAndCount(c.LastModifiedDate || c.CreatedDate, caseRec.CaseNumber, caseRec.Severity_Level__c, 'New Case', cloud, currentMode, currentUserName, 'QB');
            localStorage.setItem(trackingKey, 'true');
            actions.push('New Case (QB mention)');
          }
          break;
        }
      }
      // GHO triage (only if signature mode makes sense, but keep generic)
      for (const c of comments) {
        if (c.Body && c.Body.includes('#GHOTriage') && c.CreatedById === userId) {
          const commentDate = c.LastModifiedDate || c.CreatedDate;
          if (isToday(commentDate) && getShiftForDate(commentDate) === getCurrentShift()) {
            const ghoKey = `gho_tracked_${caseId}`;
            if (!localStorage.getItem(ghoKey)) {
              trackActionAndCount(commentDate, caseRec.CaseNumber, caseRec.Severity_Level__c, 'GHO', cloud, currentMode, currentUserName, 'QB');
              localStorage.setItem(ghoKey, 'true');
              actions.push('GHO (#GHOTriage)');
            }
          }
          break;
        }
      }
    } catch (e) {
      console.warn('Comments query failed in forceProcessCase:', e);
    }

    // Optionally, remove from persistent set if present
    try {
      chrome.runtime.sendMessage({ action: 'removeCaseFromPersistentSet', caseId }, () => { });
    } catch { }

    if (actions.length === 0) {
      showToast(`No actions tracked for ${caseNumber}`);
      return { success: true, message: 'No actions', caseId, actions };
    } else {
      showToast(`Processed ${actions.length} action(s) for ${caseNumber}`);
      return { success: true, message: 'Processed', caseId, actions };
    }
  } catch (e) {
    console.error('forceProcessCase failed:', e);
    showToast('Force process failed (see console)');
    return { success: false, message: e.message };
  }
}
window.forceProcessCase = forceProcessCase;

// Force-process by Case Id (15/18-char, typically starts with 500)
async function forceProcessCaseById(caseId) {
  try {
    if (!caseId || typeof caseId !== 'string') {
      showToast('Provide a valid Case Id');
      return { success: false, message: 'Invalid Case Id' };
    }
    if (!SESSION_ID) {
      showToast('Session not ready. Try again after data loads.');
      return { success: false, message: 'Missing SESSION_ID' };
    }

    const conn = new jsforce.Connection({
      serverUrl: 'https://orgcs.my.salesforce.com',
      sessionId: SESSION_ID,
    });

    let userId = null;
    try {
      const ures = await conn.query(`SELECT Id FROM User WHERE Name = '${currentUserName}' AND IsActive = True AND Username LIKE '%orgcs.com'`);
      if (ures.records && ures.records.length > 0) userId = ures.records[0].Id;
    } catch (e) {
      console.warn('Failed fetching user id:', e);
    }
    if (!userId) {
      showToast('Could not resolve current user');
      return { success: false, message: 'Missing user id' };
    }

    const caseRes = await conn.query(`SELECT Id, CaseNumber, Severity_Level__c, Subject, CaseRoutingTaxonomy__r.Name, Owner.Name, LastModifiedDate FROM Case WHERE Id='${caseId}' LIMIT 5`);
    if (!caseRes.records || caseRes.records.length === 0) {
      showToast(`Case ${caseId} not found`);
      return { success: false, message: 'Case not found' };
    }
    const caseRec = caseRes.records[0];
    const cloud = (caseRec.CaseRoutingTaxonomy__r && caseRec.CaseRoutingTaxonomy__r.Name) ? caseRec.CaseRoutingTaxonomy__r.Name.split('-')[0] : '';

    let actions = [];

    // History-based tracking via helper
    try {
      const res = await trackNewCaseFromHistory(conn, {
        caseIds: [caseRec.Id],
        currentUserId: userId,
        currentUserName,
        currentMode,
        strategy: 'latestByUser',
        removeFromPersistent: false
      });
      if (res && res.processed && res.processed.length > 0) actions.push('New Case (history)');
    } catch (e) { console.warn('History helper failed in forceProcessCaseById:', e); }

    // Comments-based tracking
    try {
      const cRes = await conn.query(`SELECT ParentId, Body, CreatedById, LastModifiedDate, CreatedDate FROM CaseFeed WHERE Visibility='InternalUsers' AND ParentId='${caseRec.Id}' AND Type='TextPost' ORDER BY CreatedDate DESC LIMIT 20`);
      const comments = cRes.records || [];
      for (const c of comments) {
        if (c.Body && c.Body.includes('#SigQBmention') && c.CreatedById === userId) {
          const trackingKey = `tracked_${caseRec.Id}`;
          if (!localStorage.getItem(trackingKey)) {
            trackActionAndCount(c.LastModifiedDate || c.CreatedDate, caseRec.CaseNumber, caseRec.Severity_Level__c, 'New Case', cloud, currentMode, currentUserName, 'QB');
            localStorage.setItem(trackingKey, 'true');
            actions.push('New Case (QB mention)');
          }
          break;
        }
      }
      for (const c of comments) {
        if (c.Body && c.Body.includes('#GHOTriage') && c.CreatedById === userId) {
          const commentDate = c.LastModifiedDate || c.CreatedDate;
          if (isToday(commentDate) && getShiftForDate(commentDate) === getCurrentShift()) {
            const ghoKey = `gho_tracked_${caseRec.Id}`;
            if (!localStorage.getItem(ghoKey)) {
              trackActionAndCount(commentDate, caseRec.CaseNumber, caseRec.Severity_Level__c, 'GHO', cloud, currentMode, currentUserName, 'QB');
              localStorage.setItem(ghoKey, 'true');
              actions.push('GHO (#GHOTriage)');
            }
          }
          break;
        }
      }
    } catch (e) {
      console.warn('Comments query failed in forceProcessCaseById:', e);
    }

    try {
      chrome.runtime.sendMessage({ action: 'removeCaseFromPersistentSet', caseId: caseRec.Id }, () => { });
    } catch { }

    if (actions.length === 0) {
      showToast(`No actions tracked for ${caseRec.CaseNumber}`);
      return { success: true, message: 'No actions', caseId: caseRec.Id, actions };
    } else {
      showToast(`Processed ${actions.length} action(s) for ${caseRec.CaseNumber}`);
      return { success: true, message: 'Processed', caseId: caseRec.Id, actions };
    }
  } catch (e) {
    console.error('forceProcessCaseById failed:', e);
    showToast('Force process by Id failed (see console)');
    return { success: false, message: e.message };
  }
}
window.forceProcessCaseById = forceProcessCaseById;

async function forceProcess(input) {
  const s = String(input || '').trim();
  if (!s) {
    showToast('Provide a Case Number, Id, or URL');
    return { success: false, message: 'No input' };
  }
  // Try to extract a Case Id from URL or raw Id
  let idMatch = null;
  const m1 = s.match(/Case\/([a-zA-Z0-9]{15,18})/);
  const m2 = s.match(/(500[\w]{12,15})/);
  if (m1 && m1[1]) idMatch = m1[1];
  else if (m2 && m2[1]) idMatch = m2[1];

  if (idMatch) {
    return await forceProcessCaseById(idMatch);
  }
  return await forceProcessCase(s);
}
window.forceProcess = forceProcess;
window.fp = forceProcessCase;
window.fpid = forceProcessCaseById;

let pendingUpdaterInterval = null;
function startPendingLiveUpdater() {
  try {
    const section = document.querySelector('.pending-section');
    if (!section) {
      if (pendingUpdaterInterval) {
        clearInterval(pendingUpdaterInterval);
        pendingUpdaterInterval = null;
      }
      return;
    }

    const updateOnce = () => {
      if (!document.querySelector('.pending-section')) {
        if (pendingUpdaterInterval) {
          clearInterval(pendingUpdaterInterval);
          pendingUpdaterInterval = null;
        }
        return;
      }
      const cards = document.querySelectorAll('.pending-card');
      const now = new Date();
      cards.forEach(card => {
        const createdIso = card.getAttribute('data-created');
        const totalStr = card.getAttribute('data-total');
        if (!createdIso) return;
        const created = new Date(createdIso);
        let total = Math.max(0, Number(totalStr || 0));
        const elapsedMin = Math.max(0, Math.floor((now - created) / (1000 * 60)));
        let remaining = Math.max(0, total - elapsedMin);
        const totalDisplay = total;
        const progressPct = totalDisplay > 0 ? Math.max(0, Math.min(100, Math.round((remaining / totalDisplay) * 100))) : 0;

        const elapsedEl = card.querySelector('.js-elapsed');
        if (elapsedEl) elapsedEl.textContent = `${elapsedMin}m`;

        const remainingEl = card.querySelector('.js-remaining');
        const badge = card.querySelector('.remaining-badge');
        if (badge) {
          if (remaining <= 0) {
            badge.classList.add('due');
            if (remainingEl) remainingEl.textContent = '0';
            badge.textContent = 'Due now';
          } else {
            badge.classList.remove('due');
            if (remainingEl) remainingEl.textContent = `${remaining}`;
            else badge.innerHTML = `<span class="js-remaining">${remaining}</span>m remaining`;
          }
        }

        const bar = card.querySelector('.pending-progress-bar');
        if (bar) {
          bar.style.width = `${progressPct}%`;
        }

        const prRemain = card.querySelector('.js-progress-remaining');
        if (prRemain) prRemain.textContent = `${remaining}`;
        const prTotal = card.querySelector('.js-progress-total');
        if (prTotal) prTotal.textContent = `${totalDisplay}`;
      });
    };

    updateOnce();
    if (pendingUpdaterInterval) clearInterval(pendingUpdaterInterval);
    // Update roughly every 30s so the minute counter feels live without being noisy
    pendingUpdaterInterval = setInterval(updateOnce, 30000);

    // Refresh immediately when tab/popup becomes visible again
    if (!window._pendingVisHandlerAdded) {
      document.addEventListener('visibilitychange', () => {
        if (document.visibilityState === 'visible') {
          updateOnce();
        }
      });
      window._pendingVisHandlerAdded = true;
    }
  } catch (e) {
    console.error('Failed starting pending live updater:', e);
  }
}

// ================= Helper Rendering Functions (extracted for DRY) =================
// (pending section builders moved to modules/pending.js)
// ================================================================================

// Debug functions removed for production

// ================= Weekend Roster Cache Management =================

// Cache management functions (production version - no debug output)
function getWeekendRosterCacheStats() { return _wrCacheStats(WEEKEND_ROSTER_CACHE_PREFIX, WEEKEND_ROSTER_CACHE_TTL); }

function clearWeekendRosterCache() { return _wrClearCache(WEEKEND_ROSTER_CACHE_PREFIX); }

function refreshWeekendRosterCache(sheetName, shift, weekendDateStr, mode) { return _wrRefreshCache(WEEKEND_ROSTER_CACHE_PREFIX, sheetName, shift, weekendDateStr, mode); }

function startWeekendRosterCacheCleanup() { try { _wrStartCleanup(WEEKEND_ROSTER_CACHE_PREFIX, WEEKEND_ROSTER_CACHE_TTL); } catch { } }

// Initialize cache cleanup on page load
startWeekendRosterCacheCleanup();

// ================================================================================
