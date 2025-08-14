import { timeElapsed, addMinutes, isCurrentlyWeekend } from './utils/datetime.js';
import { applyFilter, applySearch, updateWeekendModeIndicator } from './utils/dom.js';
import { trackAction, logUsageDaily } from './utils/api.js';
import { formatDateWithDayOfWeek, getShiftForDate, isToday, getCurrentShift, getPreferredShiftValues, buildPreferredShiftCondition, getWeekendSignatureTemplate, getGHOTemplate } from './modules/shift.js';
import { SHEET_SERVICE_CLOUD, SHEET_SALES_CLOUD, SHEET_INDUSTRY_CLOUD, SHEET_DATA_CLOUD_AF, getCICColumnForShift, getTEColumnForShift, getSwarmLeadColumnForShift, getPremierSalesDevTEColumn, getPremierSalesNonDevTEColumn, getPremierSalesSwarmLeadColumn, getPremierIndustryTEColumn, getPremierIndustrySwarmLeadColumn, getPremierDataCloudColumn, getPremierAgentforceColumn } from './modules/rosterDetails.js';
import { initPremierCounters, resetPremierCountersAll, parseRosterNames, renderPremierCounters } from './modules/premierCounters.js';
import { showToast } from './modules/toast.js';
import { buildPendingCardsHtml, getPendingSectionHtml } from './modules/pending.js';
import { attachGhoPreviewTemplateCopy } from './modules/gho.js';
import { logger } from './utils/logging.js';

try { logger.installConsoleBeautifier(); } catch { /* noop */ }

let SESSION_ID;
let currentMode = localStorage.getItem('caseTriageMode') || 'signature';
let currentUserName;
let searchTimeout;
let ghoRecordsGlobal = [];
let ghoConnectionGlobal = null;
export const SPREADSHEET_ID = '1BKxQLGFrczjhcx9rEt-jXGvlcCPQblwBhFJjoiDD7TI';
export const WEEKEND_ROSTER_SPREADSHEET_ID = '19qZi50CzKHm8PmSHiPjgTHogEue_DS70iXfT-MVfhPs';

// DEV: Force show Weekend Roster button and allow access regardless of actual weekend window
// Set to false to restore original behavior, or gate with a localStorage flag if preferred.
const DEV_FORCE_SHOW_WEEKEND_ROSTER = true;

const GHO_CACHE_TTL = 80000;
window.__ghoCache = window.__ghoCache || { signature: null, premier: null };
window.__ghoListState = window.__ghoListState || null;
const GHO_PAGE_SIZE = 10;
const GHO_USERMAP_CACHE_KEY = 'gho_user_map_cache_v2';
const GHO_USERMAP_CACHE_TTL = 2 * 24 * 60 * 60 * 1000;
const USER_EMAIL_CACHE_KEY = 'sf_user_email_cache_v1';
const USER_EMAIL_CACHE_TTL = 2 * 24 * 60 * 60 * 1000;

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
  } catch { /* noop */ }
}

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
function escapeSoqlString(s) { return String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'"); }
function cleanRosterNameForQuery(name) {
  let s = String(name || '').trim();
  s = s.replace(/^\s*[^:]*:\s*/, '');
  s = s.replace(/\([^)]*\)/g, '');
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
      const conn = new jsforce.Connection({ serverUrl: 'https://orgcs.my.salesforce.com', sessionId: SESSION_ID });
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
        const soql = `SELECT Name, Email FROM User WHERE IsActive = true AND Username LIKE '%orgcs.com' AND Email != null AND Name IN ('${slice.map(escapeSoqlString).join("','")}')`;
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
    currentBusinessDayKey = newKey;
    cleanupOldUsageData(oldKey, newKey);
  }
  usageState[currentBusinessDayKey] = usageState[currentBusinessDayKey] || {};
  usageState[currentBusinessDayKey][userName] = usageState[currentBusinessDayKey][userName] || { activeSeconds: 0, lastActiveISO: null, actionedCases: 0 };
  return { dayKey: currentBusinessDayKey, bucket: usageState[currentBusinessDayKey][userName] };
}

function persistUsageState() {
  try { localStorage.setItem(USAGE_STATE_KEY, JSON.stringify(usageState)); } catch { /* noop */ }
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
    if (request.action === 'trackActionFromBackground') {
      const { createdDate, caseNumber, severity, actionType, cloud, mode, userName, newValue } = request.data;
      trackActionAndCount(createdDate, caseNumber, severity, actionType, cloud, mode, userName, newValue);
      sendResponse({ success: true });
      return true;
    }
  } catch (error) {
    console.error('Error handling background message:', error);
    sendResponse({ success: false, error: error.message });
  }
  return false;
});

// Wrapper: increment per-day actioned cases and then forward to API logger
function trackActionAndCount(dateofAction, caseNumber, severity, actionType, cloud, mode, userName, assignedTo = '') {
  try {
    const name = userName || currentUserName;
    if (name) {
      const { bucket } = ensureUserUsageBucket(name);
      bucket.actionedCases = (bucket.actionedCases || 0) + 1;
      bucket.lastActiveISO = new Date().toISOString();
      persistUsageState();
    }
    trackAction(dateofAction, caseNumber, severity, actionType, cloud, mode || currentMode, name, assignedTo);
  } catch (e) {
    console.warn('trackActionAndCount failed', e);
    // Fallback to original trackAction to avoid losing logs
    try { trackAction(dateofAction, caseNumber, severity, actionType, cloud, mode || currentMode, userName || currentUserName, assignedTo); } catch { }
  }
}

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
  });
}

function getCaseDetails() {
  logger.debug('sf.session.set', { hasSession: !!SESSION_ID });
  let conn = new jsforce.Connection({
    serverUrl: 'https://orgcs.my.salesforce.com',
    sessionId: SESSION_ID,
  });
  conn.identity(function (err, res) {
    if (err) {
      const errorMessage = (err.message || err.toString()).toLowerCase();
      if (errorMessage.includes('session') || errorMessage.includes('connection')) {
        showToast('Connection or Session ID error. Please refresh or try toggling modes.');
      }

      var data = 'openTabSilent';
      chrome.runtime.sendMessage(data, function (response) {
        if (chrome.runtime.lastError) {
          console.error('Runtime error:', chrome.runtime.lastError.message);
          return;
        }
        console.log('response-----' + response);
      });
      return console.error('error---' + err);
    } else {
      currentUserName = res.display_name;
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
        }
        let signatureQuery = "SELECT Id, CreatedDate, Account.Name, Owner.Name, SE_Target_Response__c, Severity_Level__c, CaseNumber, Subject, CaseRoutingTaxonomy__r.Name,SE_Initial_Response_Status__c, Contact.Is_MVP__c, support_available_timezone__c, (SELECT Transfer_Reason__c, CreatedDate, Severity_New_Value__c, Severity_Old_Value__c FROM Case_Routing_Logs__r ORDER BY CreatedDate DESC LIMIT 5) FROM Case WHERE (Owner.Name LIKE '%Skills Queue%' OR Owner.Name='Kase Changer' OR Owner.Name='Working in Org62' OR Owner.Name='Data Cloud Queue') AND IsClosed=false AND Account_Support_SBR_Category__c!='JP MCS' AND Account.Name!='BT Test Account - HPA Premier Plus' AND Status='New' AND (((CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Community%' OR CaseRoutingTaxonomy__r.Name LIKE 'Scale Center%' OR CaseRoutingTaxonomy__r.Name LIKE 'Customer Success Score%') AND (Severity_Level__c='Level 1 - Critical' OR Severity_Level__c='Level 2 - Urgent') AND (Case_Support_level__c='Premier Priority' OR Case_Support_level__c='Signature' OR Case_Support_level__c='Signature Success')) OR ((CaseRoutingTaxonomy__r.Name LIKE 'Industry%') AND (Severity_Level__c='Level 1 - Critical' OR Severity_Level__c='Level 2 - Urgent') AND (Case_Support_level__c='Premier Priority' OR Case_Support_level__c='Signature' OR Case_Support_level__c='Signature Success')) OR (Contact.Is_MVP__c=true AND ((Severity_Level__c IN ('Level 1 - Critical', 'Level 2 - Urgent', 'Level 3 - High', 'Level 4 - Medium') AND (CaseRoutingTaxonomy__r.Name LIKE 'Sales%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry%')) OR (Severity_Level__c IN ('Level 3 - High', 'Level 4 - Medium') AND (CaseRoutingTaxonomy__r.Name LIKE 'Data Cloud-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Sales-Agentforce%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-Agentforce%')))) OR (CaseRoutingTaxonomy__r.Name='Sales-Issues Developing for Salesforce Functions (Product)' AND CreatedDate = LAST_N_DAYS:2)) AND (CaseRoutingTaxonomy__r.Name NOT IN ('Sales-Disability and Product Accessibility', 'Service-Disability and Product Accessibility', 'Industry-Disability and Product Accessibility', 'Sales-Quip', 'Sales-Sales Cloud for Slack', 'Industry-Nonprofit Cloud', 'Industry-Education Cloud', 'Industry-Education Data Architecture (EDA)', 'Industry-Education Packages (Other SFDO)', 'Industry-Nonprofit Packages (Other SFDO)', 'Industry-Nonprofit Success Pack (NPSP)', 'Service-Agentforce', 'Service-Agent for setup', 'Service-AgentforEmail', 'Service-Field Service Agentforce', 'Service-Agentforce for Dev', 'Sales-Agentforce', 'Sales-Agentforce for Dev', 'Sales-Agent for Setup', 'Sales-Prompt Builder', 'Data Cloud-Admin', 'Permissions', 'Flows', 'Reports & Dashboards', 'Data Cloud-Model Builder', 'Data Cloud-Connectors & Data Streams', 'Data Cloud-Developer', 'Calculated Insights & Consumption', 'Data Cloud-Segments', 'Activations & Identity Resolution')) ORDER BY CreatedDate DESC";
        let premierQuery = "SELECT Id, CreatedDate, Account.Name, Owner.Name, SE_Target_Response__c, Severity_Level__c, CaseNumber, Subject, CaseRoutingTaxonomy__r.Name, SE_Initial_Response_Status__c, Initial_Case_Severity__c, Contact.Is_MVP__c, (SELECT Transfer_Reason__c, CreatedDate, Severity_New_Value__c, Severity_Old_Value__c FROM Case_Routing_Logs__r ORDER BY CreatedDate DESC LIMIT 5) FROM Case WHERE (Owner.Name IN ('Kase Changer', 'Working in Org62', 'Service Cloud Skills Queue', 'Sales Cloud Skills Queue', 'Industry Skills Queue', 'EXP Skills Queue', 'Data Cloud Queue')) AND (RecordType.Name IN ('Support', 'Partner Program Support', 'Platform / Application Support')) AND (Reason != 'Sales Request') AND (CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry-%') AND (Account_Support_SBR_Category__c != 'JP') AND (Case_Support_level__c IN ('Partner Premier', 'Premier', 'Premier+', 'Premium')) AND (IsClosed = false) AND (SE_Initial_Response_Status__c NOT IN ('Met', 'Completed After Violation', 'Missed', 'Violated')) AND (Account_Support_SBR_Category__c != 'JP') AND ((Severity_Level__c IN ('Level 1 - Critical', 'Level 2 - Urgent')) OR (Initial_Case_Severity__c IN ('Level 2 - Urgent', 'Level 1 - Critical'))) AND (CaseRoutingTaxonomy__r.Name NOT IN ('Service-Agentforce', 'Service-Agent for setup', 'Service-AgentforEmail', 'Service-Field Service Agentforce', 'Service-Agentforce for Dev', 'Sales-Agentforce', 'Sales-Agentforce for Dev', 'Sales-Agent for Setup', 'Sales-Prompt Builder', 'Data Cloud-Admin', 'Permissions', 'Flows', 'Reports & Dashboards', 'Data Cloud-Model Builder', 'Data Cloud-Connectors & Data Streams', 'Data Cloud-Developer', 'Calculated Insights & Consumption', 'Data Cloud-Segments', 'Activations & Identity Resolution')) AND CreatedDate = TODAY ORDER BY CreatedDate DESC";

        let query = currentMode === 'premier' ? premierQuery : signatureQuery;

        return conn.query(query,
          function (err, result) {
            if (err) {
              alert('Your query has failed');
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
                console.log('response-----' + response);
              });
              return;
            }

            const caseIds = result.records.map(record => record.Id);

            chrome.runtime.sendMessage({
              action: 'syncPersistentCases',
              mode: currentMode,
              caseIds: caseIds
            }, function (response) {
              if (chrome.runtime.lastError) {
                console.error('Runtime error syncing persistent cases:', chrome.runtime.lastError.message);
              } else if (response && response.success) {
                logger.info('cases.persistent.synced', { mode: currentMode, removed: response.removed, tracking: response.count });
                // If cases were removed due to refresh (not in current query), try to track any last-moment assignments
                try {
                  const removedCases = Array.isArray(response.removedCases) ? response.removedCases : [];
                  if (removedCases.length > 0) {
                    const removedIds = removedCases.map(rc => rc.id);
                    trackNewCaseFromHistory(conn, {
                      caseIds: removedIds,
                      currentUserId,
                      currentUserName,
                      currentMode,
                      strategy: 'latestByUser',
                      removeFromPersistent: false
                    });
                  }
                } catch (e) {
                  console.warn('Error while post-processing removed persistent cases:', e);
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
                      localStorage.setItem(caseId, 'true');
                    }

                    const routingLogs = caseRecord.Case_Routing_Logs__r;
                    if (routingLogs && routingLogs.totalSize > 0) {
                      const lastLog = routingLogs.records[0];
                      if (lastLog.Transfer_Reason__c === 'GEO Locate' && caseRecord.SE_Initial_Response_Status__c === 'Met') {
                        localStorage.setItem(caseId, 'true');
                      }
                    }

                    const isActionTaken = localStorage.getItem(caseId) === 'true';
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
                    localStorage.setItem(caseId, 'true');
                  }

                  const routingLogs = filteredRecords[x].Case_Routing_Logs__r;
                  if (routingLogs && routingLogs.totalSize > 0) {
                    const lastLog = routingLogs.records[0];
                    if (lastLog.Transfer_Reason__c === 'GEO Locate' || lastLog.Transfer_Reason__c === 'Dispatched' && filteredRecords[x].SE_Initial_Response_Status__c === 'Met') {
                      localStorage.setItem(caseId, 'true');
                      console.log(`Auto-actioned case ${filteredRecords[x].CaseNumber} due to GEO Locate routing with Met SLA`);
                    }
                  }

                  const isSLAM = filteredRecords[x].CaseRoutingTaxonomy__r.Name === 'Sales-Issues Developing for Salesforce Functions (Product)';
                  if (isSLAM && filteredRecords[x].support_available_timezone__c === '(GMT+09:00) Japan Standard Time (Asia/Tokyo)') {
                    localStorage.setItem(caseId, 'true');
                    console.log(`Auto-actioned SLAM case ${filteredRecords[x].CaseNumber} due to Japan timezone`);
                  }

                  console.log('Processing case:', filteredRecords[x]);

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
                    const isActionTaken = localStorage.getItem(caseId) === 'true';
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

                    console.log('Case pending alert criteria:', filteredRecords[x].CaseNumber,
                      'Severity:', filteredRecords[x].Severity_Level__c,
                      'Created:', new Date(filteredRecords[x].CreatedDate),
                      'Time since creation:', minutesSinceCreation, 'minutes',
                      'Remaining:', remainingMinutes, 'minutes');
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
                    console.log('response-----' + response);
                  });
                } else {
                  var data = 'openTabSilent';
                  chrome.runtime.sendMessage(data, function (response) {
                    if (chrome.runtime.lastError) {
                      console.error('Runtime error opening tab silently:', chrome.runtime.lastError.message);
                      return;
                    }
                    console.log('response-----' + response);
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
                  console.log('response-----' + response);
                });
              }
            });
          });
      });
    }
  });
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

  console.log('GHO Alert Check:', {
    currentTime: `${Math.floor(currentTime / 60)}:${(currentTime % 60).toString().padStart(2, '0')}`,
    isAPACTime,
    isEMEATime,
    isAMERTime,
    today,
    currentUserName
  });

  if (!(isAPACTime || isEMEATime || isAMERTime)) {
    return;
  }

  let region = '';
  if (isAPACTime) region = 'APAC';
  else if (isEMEATime) region = 'EMEA';
  else if (isAMERTime) region = 'AMER';

  const alertKey = `gho_alert_${region}_${today}_${currentUserName}`;
  console.log('GHO Alert Key:', alertKey, 'Already shown:', !!localStorage.getItem(alertKey));

  if (localStorage.getItem(alertKey)) {
    return;
  }

  const preferredShiftValues = getPreferredShiftValues(region);
  const shiftCondition = buildPreferredShiftCondition(preferredShiftValues);
  console.log('GHO Alert Shift Condition:', shiftCondition);

  const ghoQuery = `SELECT Id, CreatedDate, Account.Name, Owner.Name, SE_Target_Response__c, Severity_Level__c, CaseNumber, Subject, CaseRoutingTaxonomy__r.Name, SE_Initial_Response_Status__c, Contact.Is_MVP__c, support_available_timezone__c, (SELECT Transfer_Reason__c, CreatedDate, CreatedById, Preferred_Shift_Old_Value__c, Preferred_Shift_New_Value__c, Severity_New_Value__c, Severity_Old_Value__c FROM Case_Routing_Logs__r ORDER BY CreatedDate DESC LIMIT 10) FROM Case WHERE ((Owner.Name IN ('Skills Queue','Kase Changer', 'Working in Org62','GHO Queue') AND ((Case_Support_level__c='Premier Priority' OR Case_Support_level__c='Signature' OR Case_Support_level__c='Signature Success') OR (Case_Support_level__c='Signature' OR Case_Support_level__c='Premier Priority' OR Case_Support_level__c='Signature Success'))) OR (Contact.Is_MVP__c=true AND Owner.Name='GHO Queue')) AND IsClosed=false AND ${shiftCondition} AND ((CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry%' OR CaseRoutingTaxonomy__r.Name LIKE 'Community-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Scale Center%' OR CaseRoutingTaxonomy__r.Name LIKE 'Customer Success Score%' OR CaseRoutingTaxonomy__r.Name LIKE 'Data Cloud-%') AND (Severity_Level__c='Level 1 - Critical' OR Severity_Level__c='Level 2 - Urgent')) AND CaseRoutingTaxonomy__r.Name NOT IN ('Disability and Product Accessibility','DORA')`;

  console.log('GHO Query:', ghoQuery);

  if (!SESSION_ID) {
    console.error('SESSION_ID not available for GHO alert');
    return;
  }

  let ghoConn = new jsforce.Connection({
    serverUrl: 'https://orgcs.my.salesforce.com',
    sessionId: SESSION_ID,
  });

  ghoConn.query(ghoQuery, function (err, result) {
    console.log('GHO Alert Query Result:', {
      error: err,
      recordCount: result ? result.records.length : 0,
      records: result ? result.records : null
    });

    if (err) {
      console.error('GHO Alert Query Error:', err);
      return;
    }

    if (!result.records || result.records.length === 0) {
      console.log('No GHO records found for region:', region);
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
          console.log('Checking comments for GHO triage actions:', commentResult.records.length, 'comments found');
          commentResult.records.forEach(comment => {
            if (comment.Body && comment.Body.includes('#GHOTriage')) {
              const commentDate = comment.LastModifiedDate || comment.CreatedDate;
              const isCommentFromToday = isToday(commentDate);
              const commentShift = getShiftForDate(commentDate);
              const isSameGeo = commentShift === region;

              console.log('GHO Alert - #GHOTriage comment analysis:', {
                caseId: comment.ParentId,
                commentDate,
                isFromToday: isCommentFromToday,
                commentShift,
                currentRegion: region,
                isSameGeo
              });

              if (isCommentFromToday && isSameGeo) {
                actionedCaseIds.add(comment.ParentId);
                console.log('Found valid #GHOTriage comment for case:', comment.ParentId, 'Date:', commentDate, 'Shift:', commentShift);

                if (currentUserId && comment.CreatedById === currentUserId) {
                  const caseRecord = result.records.find(c => c.Id === comment.ParentId);
                  console.log('GHO alert - Case record for triage comment:', caseRecord);
                  if (caseRecord) {
                    const ghoTrackingKey = `gho_tracked_${caseRecord.Id}`;
                    if (!localStorage.getItem(ghoTrackingKey)) {
                      trackActionAndCount(comment.LastModifiedDate, caseRecord.CaseNumber, caseRecord.Severity_Level__c, 'GHO', caseRecord.CaseRoutingTaxonomy__r.Name.split('-')[0], currentMode, currentUserName, 'QB');
                      localStorage.setItem(ghoTrackingKey, 'true');
                      console.log('GHO alert - triage action tracked for case:', caseRecord.CaseNumber, 'Date:', commentDate, 'Shift:', commentShift);
                    }
                  }
                }
              } else {
                console.log('GHO Alert - #GHOTriage comment ignored - not from today/same GEO:', {
                  caseId: comment.ParentId,
                  isFromToday: isCommentFromToday,
                  isSameGeo: isSameGeo,
                  commentDate,
                  commentShift,
                  currentRegion: region
                });
              }
            }
          });
        } else {
          console.log('No comments found for GHO cases');
        }
        const unactionedCases = result.records.filter(caseRecord => !actionedCaseIds.has(caseRecord.Id));

        console.log('GHO Alert Cases Analysis:', {
          totalCases: result.records.length,
          actionedCaseIds: Array.from(actionedCaseIds),
          unactionedCases: unactionedCases.length,
          unactionedCaseNumbers: unactionedCases.map(c => c.CaseNumber)
        });

        if (unactionedCases.length > 0) {
          console.log('Showing GHO alert for region:', region, 'with', unactionedCases.length, 'cases');
          try {
            showGHOAlert(region, unactionedCases, alertKey);
          } catch (error) {
            console.error('Failed to show GHO alert:', error);
          }
        } else {
          console.log('All GHO cases have been actioned - no alert needed');
          localStorage.setItem(alertKey, 'true');
        }
      });
    });
  });
}

function showGHOAlert(region, ghoRecords, alertKey) {
  try {
    console.log('showGHOAlert called with:', { region, recordCount: ghoRecords.length, alertKey });

    const alertTime = region === 'APAC' ? '7:30 AM' : region === 'EMEA' ? '2:30 PM' : '10:30 PM';

    let existingModal = document.getElementById('gho-alert-modal');
    if (existingModal) {
      console.log('Removing existing GHO alert modal');
      existingModal.remove();
    }

    const modalHtml = `
      <div id="gho-alert-modal" class="modal-overlay" style="display: flex; z-index: 1001;">
        <div class="modal-content" style="max-width: 600px;">
          <div class="modal-header" style="background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%); color: white;">
            <h3 style="color: white; margin: 0;">🚨 GHO Alert - ${region} Region</h3>
            <span class="modal-close" id="gho-alert-close" style="color: white; cursor: pointer; font-size: 24px;">&times;</span>
          </div>
          <div class="modal-body">
            <div style="text-align: center; margin-bottom: 20px;">
              <h4 style="color: #f59e0b; margin-bottom: 8px;">Daily GHO Check - ${alertTime} IST</h4>
              <p style="color: #6b7280; font-size: 14px;">Found ${ghoRecords.length} GHO case${ghoRecords.length === 1 ? '' : 's'} requiring attention for ${region} shift</p>
            </div>
            
            <div style="background: #fff7ed; border: 1px solid #fed7aa; border-radius: 8px; padding: 16px; margin-bottom: 20px;">
              <h5 style="color: #9a3412; margin: 0 0 8px 0;">Cases Needing Action:</h5>
              <div style="max-height: 200px; overflow-y: auto;">
                ${ghoRecords.map(caseRecord => `
                  <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px 0; border-bottom: 1px solid #fed7aa;">
                    <div>
                      <strong style="color: #9a3412;">${caseRecord.CaseNumber}</strong>
                      <span style="color: #6b7280; margin-left: 8px;">${caseRecord.Severity_Level__c}</span>
                      ${caseRecord.Contact && caseRecord.Contact.Is_MVP__c ? '<span style="background-color: #9333ea; color: white; padding: 1px 4px; border-radius: 4px; font-size: 10px; margin-left: 4px;">MVP</span>' : ''}
                    </div>
                    <div style="color: #6b7280; font-size: 12px;">
                      ${timeElapsed(new Date(caseRecord.CreatedDate))}
                    </div>
                  </div>
                `).join('')}
              </div>
            </div>
            
            <div style="display: flex; gap: 12px; justify-content: center;">
              <button id="gho-alert-check" class="preview-btn" style="background: #f59e0b; border: none; padding: 12px 24px;">
                Check GHO Cases
              </button>
              <button id="gho-alert-dismiss" class="preview-btn" style="background: #6b7280; border: none; padding: 12px 24px;">
                Dismiss
              </button>
            </div>
          </div>
        </div>
      </div>
    `;

    console.log('Inserting GHO alert modal into DOM');
    document.body.insertAdjacentHTML('beforeend', modalHtml);

    const modal = document.getElementById('gho-alert-modal');
    if (modal) {
      console.log('GHO alert modal successfully added to DOM');
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
        console.log('Audio autoplay blocked or failed:', e);
      });
    } catch (e) {
      console.log('Could not play GHO alert sound:', e);
    }

    document.getElementById('gho-alert-close').addEventListener('click', () => {
      console.log('GHO alert close button clicked - modal closed but alert not dismissed');
      document.getElementById('gho-alert-modal').remove();
    });

    document.getElementById('gho-alert-dismiss').addEventListener('click', () => {
      console.log('GHO alert dismiss button clicked - marking alert as acknowledged');
      document.getElementById('gho-alert-modal').remove();
      localStorage.setItem(alertKey, 'true');
      console.log('GHO alert dismissed and flagged as shown:', alertKey);
    });

    document.getElementById('gho-alert-check').addEventListener('click', () => {
      console.log('GHO alert check button clicked - opening GHO status without dismissing alert');
      document.getElementById('gho-alert-modal').remove();
      checkGHOStatus();
    });

    console.log('GHO alert modal setup completed successfully');
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

// Function to ensure only one extension tab exists
async function ensureSingleTab() {
  try {
    chrome.runtime.sendMessage('ensureSingleTab', function (response) {
      if (chrome.runtime.lastError) {
        console.error('Runtime error ensuring single tab:', chrome.runtime.lastError.message);
        return;
      }
      console.log('Tab cleanup requested:', response);
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
              <span style="font-weight: 600; color: #111827; font-size: 15px;">Refresh Data</span>
              <kbd style="background: #f3f4f6; padding: 4px 8px; border-radius: 4px; font-family: monospace; font-size: 12px; color: #374151; border: 1px solid #d1d5db;">⌘ + R</kbd>
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
    setTimeout(() => { m.style.display = 'none'; }, 150);
  });

  // GHO Filter Event Listener
  document.getElementById("gho-taxonomy-filter").addEventListener("change", function () {
    const filterValue = this.value;
    if (ghoRecordsGlobal.length > 0 && ghoConnectionGlobal) {
      // Reset pagination state on filter change
      window.__ghoListState = null;
      const modeCache = window.__ghoCache && window.__ghoCache[currentMode];
      renderFilteredGHOCases(ghoRecordsGlobal, ghoConnectionGlobal, filterValue, modeCache ? { triageCaseIds: modeCache.triageCaseIds } : undefined);
    }
  });

  document.getElementById("gho-modal").addEventListener("click", function (e) {
    if (e.target === this) {
      this.classList.remove('modal-show');
      setTimeout(() => { this.style.display = 'none'; }, 150);
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

// Manual track by Case Number UI wiring
try {
  const trackInput = document.getElementById('track-case-input');
  // Button removed; tracking now via Cmd+L
  if (trackInput) {
    const setTrackLoading = (loading) => {
      try {
        if (loading) {
          if (!trackBtn.dataset.origLabel) trackBtn.dataset.origLabel = trackBtn.innerHTML;
          trackBtn.innerHTML = '<i class="fa fa-spinner fa-spin" aria-hidden="true"></i> Tracking...';
          trackBtn.disabled = true;
          trackInput.disabled = true;
        } else {
          trackBtn.innerHTML = trackBtn.dataset.origLabel || 'Track Case';
          trackBtn.disabled = false;
          trackInput.disabled = false;
        }
      } catch (e) { /* noop */ }
    };

    const doTrack = async () => {
      const isHidden = trackInput.classList.contains('hidden');
      if (isHidden) {
        trackInput.classList.remove('hidden');
        trackInput.focus();
        return;
      }
      const num = (trackInput.value || '').trim();
      if (!num) { trackInput.classList.add('hidden'); return; }
      setTrackLoading(true);
      try {
        const res = await window.forceProcessCase(num);
        if (res && res.success) {
          showToast(res.actions && res.actions.length ? `Tracked ${res.actions.join(', ')}` : 'No actions (by you)');
        } else {
          showToast(res && res.message ? res.message : 'Failed to track');
        }
      } catch (err) {
        console.warn('Manual track failed:', err);
        showToast('Failed to track');
      } finally {
        setTrackLoading(false);
        trackInput.value = '';
        trackInput.classList.add('hidden');
      }
    };
    trackInput.addEventListener('keypress', (e) => { if (e.key === 'Enter') doTrack(); });

    document.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter' || e.metaKey || e.ctrlKey || e.altKey) return;
      const active = document.activeElement;
      const isOtherInput = active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA') && active !== trackInput;
      if (isOtherInput) return;
      if (!trackInput.classList.contains('hidden')) {
        const val = (trackInput.value || '').trim();
        if (val) {
          e.preventDefault();
          doTrack();
        }
      }
    });

    document.addEventListener('keydown', (e) => {
      const active = document.activeElement;
      const isTyping = active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA') && active !== trackInput;
      if (isTyping) return;
      const isMac = navigator.platform.toUpperCase().indexOf('MAC') >= 0;
      const metaPressed = isMac ? e.metaKey : e.ctrlKey;
      if (metaPressed && !e.shiftKey && !e.altKey && (e.key === 'l' || e.key === 'L')) {
        e.preventDefault();
        const isHidden = trackInput.classList.contains('hidden');
        if (isHidden) {
          trackInput.classList.remove('hidden');
          trackInput.focus();
          trackInput.select();
        } else {
          const val = (trackInput.value || '').trim();
          if (val) {
            doTrack();
          } else {
            trackInput.classList.add('hidden');
          }
        }
      }
    });
  }
} catch (e) { console.warn('Manual track wiring failed:', e); }

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
        alert('Please enter a valid number of minutes (1-1440).');
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

function googleSheetsGET(rangeA1, callback, onError) {
  chrome.identity.getAuthToken({ interactive: true }, function (token) {
    if (chrome.runtime.lastError || !token) {
      console.error('Auth token error:', chrome.runtime.lastError);
      onError && onError(new Error('Auth token error'));
      return;
    }
    fetch(`https://sheets.googleapis.com/v4/spreadsheets/${WEEKEND_ROSTER_SPREADSHEET_ID}/values/${encodeURIComponent(rangeA1)}?majorDimension=ROWS`, {
      headers: { 'Authorization': 'Bearer ' + token }
    }).then(r => r.json()).then(data => {
      if (data.error) throw new Error(data.error.message || 'Sheets error');
      callback && callback(data);
    }).catch(err => {
      console.error('Sheets read failed:', err);
      onError && onError(err);
    });
  });
}

async function getCellValueWithoutStrikethrough(sheetName, rowIndexZero, colLetter) {
  return new Promise((resolve, reject) => {
    try {
      const a1 = `${colLetter}${rowIndexZero + 1}`;
      chrome.identity.getAuthToken({ interactive: true }, function (token) {
        if (chrome.runtime.lastError || !token) {
          return reject(new Error('Auth token error'));
        }
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${WEEKEND_ROSTER_SPREADSHEET_ID}?includeGridData=true&ranges=${encodeURIComponent(`${sheetName}!${a1}`)}`;
        fetch(url, { headers: { 'Authorization': 'Bearer ' + token } })
          .then(r => r.json())
          .then(j => {
            try {
              const sheets = j.sheets || [];
              for (const sh of sheets) {
                const data = (sh.data || []);
                for (const block of data) {
                  const rows = block.rowData || [];
                  for (const row of rows) {
                    const vals = row.values || [];
                    for (const v of vals) {
                      const text = (v.formattedValue) || (v.effectiveValue && v.effectiveValue.stringValue) || '';
                      const runs = v.textFormatRuns || [];
                      if (!runs || runs.length === 0) return resolve(String(text || ''));
                      const ordered = runs.slice().sort((a, b) => (a.startIndex || 0) - (b.startIndex || 0));
                      let out = '';
                      for (let i = 0; i < ordered.length; i++) {
                        const start = ordered[i].startIndex || 0;
                        const end = (i + 1 < ordered.length) ? (ordered[i + 1].startIndex || text.length) : text.length;
                        const seg = text.substring(start, end);
                        const fmt = ordered[i].format || {};
                        if (fmt.strikethrough) continue;
                        out += seg;
                      }
                      return resolve(out.trim());
                    }
                  }
                }
              }
              resolve('');
            } catch (e) { reject(e); }
          })
          .catch(err => reject(err));
      });
    } catch (e) { reject(e); }
  });
}

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

  googleSheetsGET(range, (resp) => {
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
            <button id="premier-reset" title="Reset all counters" style="margin-left:auto; padding:6px 10px; border:1px solid #ef4444; background:#fef2f2; color:#b91c1c; border-radius:8px; font-size:12px; cursor:pointer;">Reset Counts</button>
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
        const swarmCol = getSwarmLeadColumnForShift(shift);
        const swarmIdx = swarmCol.charCodeAt(0) - 'A'.charCodeAt(0);
        const swarmCellRaw = (foundRowIdx >= 0 && rows[foundRowIdx]) ? (rows[foundRowIdx][swarmIdx] || '') : '';
        const swarmCell = String(swarmCellRaw || '').trim();
        const swarmTitle = 'Signature Swarm Lead - ' + shift;
        const bodyEl2 = document.querySelector('#cic-modal .modal-body');
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
      }

      // In Signature mode, also show Sales/Service Cloud TEs sections
      if (currentMode === 'signature') {
        const bodyEl2 = document.querySelector('#cic-modal .modal-body');
        if (bodyEl2) {
          // Helper to append a TE section for a given sheet and title
          const appendTESection = (sheetName, titlePrefix, idSuffix) => {
            const teColumn = getTEColumnForShift(shift);
            const teRange = `'${sheetName}'!A:Z`;
            googleSheetsGET(teRange, (r2) => {
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
                  const dt2 = new Date(aStr2);
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
                  bodyEl2.insertAdjacentHTML('beforeend', sectionHtml);
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
                bodyEl2.insertAdjacentHTML('beforeend', sectionHtml);

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

          appendTESection(SHEET_SALES_CLOUD, 'Sales Cloud TEs', 'sales-te');
          appendTESection(SHEET_SERVICE_CLOUD, 'Service Cloud TEs', 'service-te');
        }
      }

      // Premier mode: add dropdown-driven sections
      if (currentMode === 'premier') {
        const container = document.getElementById('premier-sections');
        const select = document.getElementById('premier-view');
        const resetBtn = document.getElementById('premier-reset');
        if (container && select) {
          if (resetBtn) {
            resetBtn.addEventListener('click', () => {
              resetPremierCountersAll(weekendDateStr, shift); // re-render current view
              const v = select.value; if (v === 'sales') return renderSales(); if (v === 'service') return renderService(); if (v === 'industry') return renderIndustry(); return renderData();
            });
          }
          const renderSales = () => {
            container.innerHTML = '';
            // 3 cells from Sales Cloud sheet
            const rangePS = `'${SHEET_SALES_CLOUD}'!A:Z`;
            googleSheetsGET(rangePS, (r) => {
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
                  }
                }
              });
            }, (e) => console.error('Premier Sales read fail:', e));
          };

          const renderService = () => {
            container.innerHTML = '';
            // 3 cells from Service Cloud sheet
            const rangePSe = `'${SHEET_SERVICE_CLOUD}'!A:Z`;
            googleSheetsGET(rangePSe, (r) => {
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
                  }
                }
              });
            }, (e) => console.error('Premier Service read fail:', e));
          };

          const renderIndustry = () => {
            container.innerHTML = '';
            // 2 cells: Industry TEs, Swarm Lead from Industry Cloud sheet
            const rangePI = `'${SHEET_INDUSTRY_CLOUD}'!A:Z`;
            googleSheetsGET(rangePI, (r) => {
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
                }
              });
            }, (e) => console.error('Premier Industry read fail:', e));
          };

          const renderData = () => {
            container.innerHTML = '';
            // 2 cells from Data Cloud and AF sheet: DataCloud {Shift}, Agentforce {Shift}
            const rangePD = `'${SHEET_DATA_CLOUD_AF}'!A:Z`;
            googleSheetsGET(rangePD, (r) => {
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
function checkGHOStatus() {
  const modal = document.getElementById('gho-modal');
  const container = document.getElementById('gho-cases-container');
  const filterDropdown = document.getElementById('gho-taxonomy-filter');

  // Reset filter to "All" when opening modal
  if (filterDropdown) {
    filterDropdown.value = 'All';
  }

  // Show modal and loading state
  modal.style.display = 'flex';
  requestAnimationFrame(() => modal.classList.add('modal-show'));
  container.innerHTML = `
    <div class="loading-message" style="animation: fadeUp 300ms ease;">
      <div class="spinner"></div>
      <h4 style="margin-bottom:8px">Loading GHO cases...</h4>
      <p>Please wait while we fetch the latest GHO cases.</p>
    </div>
  `;

  // Use cached results if within TTL and for same shift to avoid repeated network & comment queries
  const nowTs = Date.now();
  const currentShiftLive = getCurrentShift();
  const modeCache = window.__ghoCache[currentMode];
  if (modeCache && (nowTs - modeCache.fetchedAt < GHO_CACHE_TTL) && modeCache.shift === currentShiftLive) {
    try {
      ghoRecordsGlobal = modeCache.records;
      ghoConnectionGlobal = modeCache.conn;
      // Re-render with cached comment/user resolution
      renderFilteredGHOCases(modeCache.records, modeCache.conn, 'All', {
        triageCaseIds: modeCache.triageCaseIds,
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

  const signatureGHOQuery = `SELECT Id, CreatedDate, Account.Name, Owner.Name, SE_Target_Response__c, Severity_Level__c, CaseNumber, Subject, CaseRoutingTaxonomy__r.Name, SE_Initial_Response_Status__c, Contact.Is_MVP__c, support_available_timezone__c, (SELECT Transfer_Reason__c, CreatedDate, CreatedById, Preferred_Shift_Old_Value__c, Preferred_Shift_New_Value__c, Severity_New_Value__c, Severity_Old_Value__c FROM Case_Routing_Logs__r ORDER BY CreatedDate DESC LIMIT 10) FROM Case WHERE ((Owner.Name IN ('Skills Queue','Kase Changer', 'Working in Org62','GHO Queue') AND (Case_Support_level__c IN ('Premier Priority','Signature','Signature Success'))) OR (Contact.Is_MVP__c=true AND Owner.Name='GHO Queue')) AND IsClosed=false AND ${shiftCondition} AND ((CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry%' OR CaseRoutingTaxonomy__r.Name LIKE 'Community-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Scale Center%' OR CaseRoutingTaxonomy__r.Name LIKE 'Customer Success Score%' OR CaseRoutingTaxonomy__r.Name LIKE 'Data Cloud-%') AND (Severity_Level__c='Level 1 - Critical' OR Severity_Level__c='Level 2 - Urgent')) AND CaseRoutingTaxonomy__r.Name NOT IN ('Disability and Product Accessibility','DORA')`;
  const premierGHOQuery = `SELECT Case__c, Transfer_Reason__c, Case__r.Case_Support_level__c, Case__r.IsClosed, Case__r.Account_Support_SBR_Category__c, Case__r.Severity_Level__c, Case__r.OrgCS_Owner__c, Case__r.Contact.Is_MVP__c, Case__r.GHO__c, Case__r.Out_of_Impact_Service_Restored__c, Case__r.AX_Product_Name__c, Case__r.CaseRoutingTaxonomy__r.Name FROM Case_Routing_Log__c WHERE Transfer_Reason__c = 'GHO' AND Case__r.Case_Support_level__c In ('Partner Premier', 'Premier', 'Premier+', 'Premium','Standard','') AND Case__r.IsClosed = false AND Case__r.Account_Support_SBR_Category__c != 'JP' AND Case__r.Severity_Level__c IN ('Level 1 - Critical' ,'Level 2 - Urgent') AND Case__r.OrgCS_Owner__c LIKE '%Queue%' AND Case__r.Contact.Is_MVP__c = false AND Case__r.GHO__c = true AND Case__r.Out_of_Impact_Service_Restored__c = false AND (Case__r.AX_Product_Name__c = 'Sales' OR Case__r.AX_Product_Name__c = 'Service'  OR Case__r.AX_Product_Name__c = 'Industry' ) AND Case__r.CaseRoutingTaxonomy__r.Name NOT IN ('Service-Agentforce','Service-Agent for setup','Service-AgentforEmail','Service-Field Service Agentforce','Service-Agentforce for Dev','Sales-Agentforce','Sales -Agentforce for Dev','Sales-Agent for Setup','Sales-Prompt Builder','Data Cloud-Admin','Permissions','Flows','Reports & Dashboards','Data Cloud-Model Builder','Data Cloud-Connectors & Data Streams','Data Cloud-Developer','Calculated Insights & Consumption','Data Cloud-Segments','Activations & Identity Resolution')`;
  const ghoQuery = currentMode === 'premier' ? premierGHOQuery : signatureGHOQuery;

  ghoConn.query(ghoQuery, function (err, result) {
    if (err) {
      console.error('GHO Query Error:', err);
      container.innerHTML = `
        <div class="no-cases-message" style="background: linear-gradient(135deg,#fff1f2 0%,#ffe4e6 100%); border: 1.5px solid #fecaca; animation: fadeUp 300ms ease;">
          <h4 class="no-cases-title" style="color: #dc2626;">Error Loading GHO Cases</h4>
          <p class="no-cases-text" style="color: #dc2626;">Failed to fetch GHO cases. Please check your connection and try again.</p>
          <p style="color: #6b7280; font-size: 12px; margin-top: 8px;">Error: ${err.message}</p>
        </div>
      `;
      return;
    }

    // For Premier (log-based) flow: hydrate full Case records before continuing existing pipeline logic.
    if (currentMode === 'premier' && result && Array.isArray(result.records)) {
      const caseIds = Array.from(new Set(result.records.map(r => r.Case__c))).filter(Boolean);
      if (caseIds.length === 0) {
        // No qualifying cases
        showGHOStatusModal([], ghoConn);
        return;
      }
      const detailQuery = `SELECT Id, CreatedDate, Account.Name, Owner.Name, SE_Target_Response__c, Severity_Level__c, CaseNumber, Subject, CaseRoutingTaxonomy__r.Name, SE_Initial_Response_Status__c, Contact.Is_MVP__c, support_available_timezone__c, (SELECT Transfer_Reason__c, CreatedDate, CreatedById, Preferred_Shift_Old_Value__c, Preferred_Shift_New_Value__c, Severity_New_Value__c, Severity_Old_Value__c FROM Case_Routing_Logs__r ORDER BY CreatedDate DESC LIMIT 10) FROM Case WHERE Id IN ('${caseIds.join("','")}')`;
      return ghoConn.query(detailQuery, (detailErr, detailRes) => {
        if (detailErr) {
          console.error('Premier detail fetch error:', detailErr);
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
    if (!result) { showGHOStatusModal([], ghoConn); return; }
    if (result.records.length === 0) {
      showGHOStatusModal(result.records || [], ghoConn);
      return;
    }
    const caseIds = result.records.map(r => r.Id);
    const commentQuery = `SELECT ParentId, Body, CreatedById, LastModifiedDate, CreatedDate FROM CaseFeed WHERE Visibility = 'InternalUsers' AND ParentId IN ('${caseIds.join("','")}') AND Type = 'TextPost'`;
    Promise.all([
      new Promise(res => ghoConn.query(commentQuery, (e, r) => res({ e, r }))),
      new Promise(res => ghoConn.query(`SELECT Id FROM User WHERE Name = '${currentUserName}' AND IsActive = True AND Username LIKE '%orgcs.com'`, (e, r) => res({ e, r })))
    ]).then(([commentResp, userResp]) => {
      let triageCaseIds = new Set();
      let currentUserId = null;
      if (!userResp.e && userResp.r.records && userResp.r.records.length > 0) currentUserId = userResp.r.records[0].Id;
      if (!commentResp.e && commentResp.r.records && currentUserId) {
        commentResp.r.records.forEach(c => {
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
      }

      window.__ghoCache[currentMode] = {
        records: result.records,
        triageCaseIds,
        userMap: {},
        fetchedAt: Date.now(),
        shift: currentShiftLive,
        conn: ghoConn
      };
      showGHOStatusModal(result.records, ghoConn, { triageCaseIds, fromCache: false });
    }).catch(errAll => {
      console.warn('Parallel GHO enrichment failed, rendering basic list', errAll);
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
      'data': ['data cloud', 'data'],
      'sales': ['sales'],
      'service': ['service'],
      'industry': ['industry']
    };
    const selected = (filterValue || '').toLowerCase();
    const targets = aliases[selected] || [selected];

    filteredRecords = ghoRecords.filter(record => {
      const rawName = (record.CaseRoutingTaxonomy__r?.Name || '').toLowerCase().trim();
      const group = rawName.split(/[\-–—]/)[0].trim();
      return targets.includes(group);
    });
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
    renderGHOCasesWithCommentInfo(filteredRecords, conn, currentShift, filterValue, pre.triageCaseIds);
    return;
  }

  const caseIds = filteredRecords.map(record => record.Id);
  if (caseIds.length === 0) {
    renderGHOCasesWithCommentInfo(filteredRecords, conn, currentShift, filterValue, new Set());
    return;
  }
  const commentQuery = `SELECT ParentId, Body, CreatedById, LastModifiedDate, CreatedDate FROM CaseFeed WHERE Visibility = 'InternalUsers' AND ParentId IN ('${caseIds.join("','")}') AND Type = 'TextPost'`;
  conn.query(commentQuery, function (commentErr, commentResult) {
    if (commentErr) {
      console.error('Error querying GHO comments:', commentErr);
      renderGHOCasesWithCommentInfo(filteredRecords, conn, currentShift, filterValue, new Set());
      return;
    }
    const ghoTriageCommentCases = new Set();
    if (commentResult.records) {
      commentResult.records.forEach(comment => {
        if (comment.Body && comment.Body.includes('#GHOTriage')) {
          const commentDate = comment.LastModifiedDate || comment.CreatedDate;
          const isCommentFromToday = isToday(commentDate);
          const commentShift = getShiftForDate(commentDate);
          const isSameGeo = commentShift === currentShift;
          if (isCommentFromToday && isSameGeo) ghoTriageCommentCases.add(comment.ParentId);
        }
      });
    }
    renderGHOCasesWithCommentInfo(filteredRecords, conn, currentShift, filterValue, ghoTriageCommentCases);
  });
}

// Helper function to render GHO cases with comment information
function renderGHOCasesWithCommentInfo(filteredRecords, conn, currentShift, filterValue, ghoTriageCommentCases) {
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
      }
      // also persist to localStorage day-cache
      const existing = loadUserMapCache();
      saveUserMapCache({ ...(existing.userMap || {}), ...(userMap || {}) });
    } catch (e) { /* noop */ }

    // Count cases with QB mentions
    const qbMentionedCount = Array.from(ghoTriageCommentCases).length;

    const headerHtml = `
      <div class="gho-summary" style="animation: fadeUp 260ms ease;">
        <h4>Found ${filteredRecords.length} ${filterValue === 'All' ? '' : filterValue + ' '}GHO Case${filteredRecords.length === 1 ? '' : 's'}</h4>
        <p>Cases matching ${filterValue === 'All' ? 'GHO criteria' : filterValue + ' taxonomy'} for <strong>${currentShift}</strong> shift</p>
        ${qbMentionedCount > 0 ? `<p><span class="badge-soft badge-soft--success">${qbMentionedCount}</span> case${qbMentionedCount === 1 ? ' has' : 's have'} QB mentioned (#GHOTriage)</p>` : ''}
        <div class="gho-summary-meta">
          <strong>Current Shift:</strong> ${currentShift} | 
          <strong>Time:</strong> ${new Date().toLocaleTimeString()} |
          <strong>Shifts:</strong> APAC (5:30AM-2:30PM), EMEA (12:30PM-9:30PM), AMER (8:30PM-5:30AM)
        </div>
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

function renderSingleGhoCaseHTML(caseRecord, userMap, ghoTriageCommentCases, currentShift) {
  const caseId = caseRecord.Id;
  const isMVP = caseRecord.Contact && caseRecord.Contact.Is_MVP__c === true;
  const hasGHOTriage = ghoTriageCommentCases.has(caseId);

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
              <div class="gho-transfer-meta">${timeElapsed(transferTime)} ago</div>
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
      buf += renderSingleGhoCaseHTML(state.allRecords[i], state.userMap, state.triageCaseIds, state.currentShift);
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

// Debug function to check persistent case set status
function debugPersistentCases() {
  console.log('Requesting persistent case details...');

  chrome.runtime.sendMessage({ action: 'getPersistentCaseDetails' }, (response) => {
    if (chrome.runtime.lastError) {
      console.error('Runtime error:', chrome.runtime.lastError);
      return;
    }

    console.log('Response received:', response);

    if (response && typeof response.count !== 'undefined') {
      console.log(`✅ Persistent Case Set Status: ${response.count} cases being tracked`);
      if (response.cases && response.cases.length > 0) {
        console.log('📊 Case Details:');
        console.table(response.cases);
      } else if (response.count === 0) {
        console.log('📭 No cases currently being tracked');
      }
    } else {
      console.warn('⚠️ No response received or invalid response format');
    }
  });
}

// Alternative simple debug function
function debugPersistentCasesSimple() {
  chrome.runtime.sendMessage({ action: 'getPersistentCaseCount' }, (response) => {
    if (chrome.runtime.lastError) {
      console.error('Runtime error:', chrome.runtime.lastError);
      return;
    }
    console.log('Persistent cases count:', response ? response.count : 'undefined');
  });
}

// Debug function to clear GHO alert flags
function clearGHOAlertFlags() {
  const today = new Date().toDateString();
  const currentUser = currentUserName || 'unknown';

  console.log('Clearing GHO alert flags for today:', today, 'user:', currentUser);

  const regions = ['APAC', 'EMEA', 'AMER'];
  let clearedCount = 0;

  regions.forEach(region => {
    const alertKey = `gho_alert_${region}_${today}_${currentUser}`;
    if (localStorage.getItem(alertKey)) {
      localStorage.removeItem(alertKey);
      clearedCount++;
      console.log(`Cleared dismissed flag for ${region}:`, alertKey);
    }
  });

  const allKeys = Object.keys(localStorage);
  const ghoKeys = allKeys.filter(key => key.startsWith('gho_alert_') || key.startsWith('gho_tracked_'));

  console.log(`Found ${ghoKeys.length} GHO-related localStorage keys`);
  console.log(`Cleared ${clearedCount} dismissed alert flags for today`);

  if (clearedCount > 0) {
    console.log('✅ GHO alert flags cleared. Alerts will now reappear at designated times until properly dismissed.');
  } else {
    console.log('ℹ️ No dismissed alert flags found for today to clear.');
  }

  console.log('💡 Remember: Alerts will only be marked as "shown" when you click the DISMISS button, not just close or check buttons.');

  return { clearedAlertFlags: clearedCount, totalGHOKeys: ghoKeys.length };
}
window.debugPersistentCases = debugPersistentCases;
window.debugPersistentCasesSimple = debugPersistentCasesSimple;
window.clearGHOAlertFlags = clearGHOAlertFlags;

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
    const caseRes = await conn.query(`SELECT Id, CaseNumber, Severity_Level__c, Subject, CaseRoutingTaxonomy__r.Name, Owner.Name, LastModifiedDate FROM Case WHERE CaseNumber='${caseNumber}' LIMIT 1`);
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
      chrome.runtime.sendMessage({ action: 'removeCaseFromPersistentSet', caseId }, () => { /* noop */ });
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

    const caseRes = await conn.query(`SELECT Id, CaseNumber, Severity_Level__c, Subject, CaseRoutingTaxonomy__r.Name, Owner.Name, LastModifiedDate FROM Case WHERE Id='${caseId}' LIMIT 1`);
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
      chrome.runtime.sendMessage({ action: 'removeCaseFromPersistentSet', caseId: caseRec.Id }, () => { /* noop */ });
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
