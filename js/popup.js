import { timeElapsed, addMinutes, isCurrentlyWeekend } from './utils/datetime.js';
import { applyFilter, applySearch, updateWeekendModeIndicator } from './utils/dom.js';
import { trackAction } from './utils/api.js';

let SESSION_ID;
let currentMode = localStorage.getItem('caseTriageMode') || 'signature';
let currentUserName;
let searchTimeout;
let ghoRecordsGlobal = [];
let ghoConnectionGlobal = null;
export const SPREADSHEET_ID = '1BKxQLGFrczjhcx9rEt-jXGvlcCPQblwBhFJjoiDD7TI';
export const WEEKEND_ROSTER_SPREADSHEET_ID = '19qZi50CzKHm8PmSHiPjgTHogEue_DS70iXfT-MVfhPs';

let mouseActivityTimer;
const MOUSE_ACTIVITY_TIMEOUT = 60000;

function resetMouseActivityTimer() {
  if (mouseActivityTimer) {
    clearTimeout(mouseActivityTimer);
  }
  mouseActivityTimer = setTimeout(() => {
    console.log('No mouse activity for 60 seconds, refreshing window...');
    window.location.reload();
  }, MOUSE_ACTIVITY_TIMEOUT);
}

function initMouseActivityTracking() {
  const events = ['mousemove', 'mousedown', 'mouseup', 'click', 'scroll', 'keydown', 'keyup'];

  events.forEach(eventType => {
    document.addEventListener(eventType, resetMouseActivityTimer, true);
  });
  resetMouseActivityTimer();
  console.log('Mouse activity tracking initialized - auto-refresh after 60 seconds of inactivity');
}

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  try {
    if (request.action === 'trackActionFromBackground') {
      const { createdDate, caseNumber, severity, actionType, cloud, mode, userName, newValue } = request.data;
      trackAction(createdDate, caseNumber, severity, actionType, cloud, mode, userName, newValue);
      sendResponse({ success: true });
      return true;
    }
  } catch (error) {
    console.error('Error handling background message:', error);
    sendResponse({ success: false, error: error.message });
  }
  return false;
});

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
  console.log('SESSION_ID>>>', SESSION_ID);
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

        const currentShift = getCurrentShift();
        const preferredShiftValues = getPreferredShiftValues(currentShift);
        const shiftCondition = buildPreferredShiftCondition(preferredShiftValues);

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
                  <p class="mode-switch-hint">💡 Try switching between Signature and Premier modes to see different case types.</p>
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
                console.log(`Synced persistent cases for mode ${currentMode}. Removed: ${response.removed}. Now tracking: ${response.count}`);
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
                console.log(response.message);
              }
            });


            const caseMap = new Map(result.records.map(record => [record.Id, record]));

            const historyQuery2 = `SELECT CaseId, CreatedById, CreatedDate, Field, NewValue FROM CaseHistory WHERE CaseId IN ('${caseIds.join("','")}') AND (Field = 'Routing_Status__c' OR Field = 'Owner') AND CreatedById = '${currentUserId}' ORDER BY CreatedDate ASC LIMIT 2`;
            conn.query(historyQuery2, function (historyErr, historyResult) {
              if (historyErr) {
                return console.error('Error fetching case history:', historyErr);
              }

              if (!historyResult.records || historyResult.records.length === 0) {
                return;
              }

              const manuallyAssignedHistories = historyResult.records.filter(
                history => history.NewValue && typeof history.NewValue === 'string' && history.NewValue.startsWith('Manually Assigned')
              );

              const firstAssignments = new Map();
              for (const history of manuallyAssignedHistories) {
                if (!firstAssignments.has(history.CaseId)) {
                  firstAssignments.set(history.CaseId, history);
                }
              }

              if (firstAssignments.size > 0) {
                for (const [caseId, history] of firstAssignments.entries()) {
                  if (history.CreatedById === currentUserId) {
                    const trackingKey = `tracked_${currentMode}_assignment_${caseId}`;
                    if (!localStorage.getItem(trackingKey)) {
                      const caseRecord = caseMap.get(caseId);
                      if (caseRecord) {
                        trackAction(
                          history.CreatedDate,
                          caseRecord.CaseNumber,
                          caseRecord.Severity_Level__c,
                          'New Case',
                          caseRecord.CaseRoutingTaxonomy__r.Name.split('-')[0],
                          currentMode,
                          currentUserName,
                          history.NewValue
                        );
                        localStorage.setItem(trackingKey, 'true');

                        chrome.runtime.sendMessage({
                          action: 'removeCaseFromPersistentSet',
                          caseId: caseId
                        }, function (response) {
                          if (chrome.runtime.lastError) {
                            console.error('Runtime error removing case from persistent set:', chrome.runtime.lastError.message);
                          }
                        });
                      }
                    }
                  }
                }
              }
            });

            const commentQuery = `SELECT ParentId, Body, CreatedById, LastModifiedDate FROM CaseFeed WHERE Visibility = 'InternalUsers' AND ParentId IN ('${caseIds.join("','")}') AND Type = 'TextPost'`;

            conn.query(commentQuery, function (commentErr, commentResult) {
              console.log("Case Comments:", commentResult);

              if (commentErr) {
                return console.error(commentErr);
              }

              const actionedCaseIds = new Set();
              if (commentResult.records) {
                commentResult.records.forEach(record => {
                  if (record.Body && record.Body.includes('#SigQBmention')) {
                    console.log('Actioned case found:', record.ParentId);
                    actionedCaseIds.add(record.ParentId);
                    if (record.CreatedById === currentUserId) {
                      const caseRecord = result.records.find(c => c.Id === record.ParentId);
                      console.log('Case record for actioned comment:', caseRecord);
                      if (caseRecord) {
                        const trackingKey = `tracked_${caseRecord.Id}`;
                        if (!localStorage.getItem(trackingKey)) {
                          trackAction(caseRecord.LastModifiedDate, caseRecord.CaseNumber, caseRecord.Severity_Level__c, 'New Case', caseRecord.CaseRoutingTaxonomy__r.Name.split('-')[0], currentMode, currentUserName, 'QB');
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
                    console.log('GHO Triage action found:', record.ParentId);

                    const commentDate = record.LastModifiedDate || record.CreatedDate;
                    const isCommentFromToday = isToday(commentDate);
                    const commentShift = getShiftForDate(commentDate);
                    const currentShift = getCurrentShift();
                    const isSameGeo = commentShift === currentShift;

                    console.log('GHO Triage comment analysis:', {
                      commentDate,
                      isFromToday: isCommentFromToday,
                      commentShift,
                      currentShift,
                      isSameGeo
                    });

                    if (isCommentFromToday && isSameGeo && record.CreatedById === currentUserId) {
                      const caseRecord = result.records.find(c => c.Id === record.ParentId);
                      console.log('Case record for GHO triage comment:', caseRecord);
                      if (caseRecord) {
                        const ghoTrackingKey = `gho_tracked_${caseRecord.Id}`;
                        if (!localStorage.getItem(ghoTrackingKey)) {
                          trackAction(record.LastModifiedDate, caseRecord.CaseNumber, caseRecord.Severity_Level__c, 'GHO', caseRecord.CaseRoutingTaxonomy__r.Name.split('-')[0], currentMode, currentUserName, 'QB');
                          localStorage.setItem(ghoTrackingKey, 'true');
                          console.log('GHO triage action tracked for case:', caseRecord.CaseNumber, 'Date:', commentDate, 'Shift:', commentShift);

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
                      console.log('GHO Triage comment ignored - not from today/same GEO:', {
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
              var totalCasesCount = 0; // Track total cases before filtering
              var pendingCasesCount = 0; // Track cases that exist but haven't met alert criteria
              var pendingCasesDetails = []; // Store details of pending cases


              function isWeekend() {
                return isCurrentlyWeekend();
              }
              var minSev1 = isWeekend() ? 1 : 5;
              var minSev2 = isWeekend() ? 1 : 20;
              if (result.records.length > 0) {
                isData = true;

                console.log('All case records received:', result.records);
                console.log('Total records:', result.records.length);
                console.log(result.records.filter(x => x.Contact));

                totalCasesCount = result.records.length;

                const filteredRecords = result.records.filter(x => {
                  if (x.Contact && x.Contact.Is_MVP__c === true && x.SE_Initial_Response_Status__c === 'Met') {
                    console.log(`Filtering out MVP case with Met status: ${x.CaseNumber} ${x.SE_Initial_Response_Status__c}`);
                    return false;
                  }
                  return true;
                });

                console.log('Filtered records (after removing MVP Met cases):', filteredRecords.length, 'out of', result.records.length);

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
                  let pendingCasesHtml = '';
                  pendingCasesDetails.forEach(caseDetail => {
                    const severityShort = caseDetail.severity.includes('Level 1') ? 'SEV1' :
                      caseDetail.severity.includes('Level 2') ? 'SEV2' :
                        caseDetail.severity.includes('Level 3') ? 'SEV3' : 'SEV4';
                    const mvpBadge = caseDetail.isMVP ? '<span class="badge-soft badge-soft--purple" style="padding:1px 6px; font-size:10px; margin-left:8px;">MVP</span>' : '';
                    pendingCasesHtml += `
                      <div class="pending-item">
                        <div class="pending-left">
                          <div class="pending-title">
                            ${caseDetail.caseNumber} - ${severityShort}${mvpBadge}
                          </div>
                          <div class="pending-sub">
                            ${caseDetail.account}
                          </div>
                          <div class="pending-meta">
                            Created: ${formatDateWithDayOfWeek(caseDetail.createdDate)} (${caseDetail.minutesSinceCreation}m ago)
                          </div>
                        </div>
                        <div class="pending-right">
                          ${caseDetail.remainingMinutes > 0 ? `${caseDetail.remainingMinutes}m remaining` : 'Due now'}
                        </div>
                      </div>
                    `;
                  });

                  const pendingCasesBanner = `
                    <div class="no-cases-message" style="background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%); border: 2px solid #3b82f6; margin-top: 20px;">
                      <h4 class="no-cases-title" style="color: #1d4ed8;">Cases Monitoring - No Action Required Yet</h4>
                      <p class="no-cases-text">${pendingCasesCount} case${pendingCasesCount === 1 ? ' is' : 's are'} within SLA window (SEV1: ${minSev1}min, SEV2: ${minSev2}min). Monitoring in progress...</p>
                      <div style="margin-top: 16px;">
                        <div style="font-weight: 600; color: #1d4ed8; margin-bottom: 8px; font-size: 14px;">Pending Cases:</div>
                        ${pendingCasesHtml}
                      </div>
                      <p class="mode-switch-hint" style="margin-top: 16px;">⏱️ These cases will appear in the action section when they exceed SLA thresholds.</p>
                    </div>
                  `;

                  finalHtml = finalHtml + pendingCasesBanner;
                }

                const container = document.getElementById("parentSigSev2");
                container.classList.remove('is-loading');
                container.innerHTML = finalHtml; // replace loading tile entirely
                container.classList.add('content-enter');

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
                  let pendingCasesHtml = '';
                  pendingCasesDetails.forEach(caseDetail => {
                    const severityShort = caseDetail.severity.includes('Level 1') ? 'SEV1' :
                      caseDetail.severity.includes('Level 2') ? 'SEV2' :
                        caseDetail.severity.includes('Level 3') ? 'SEV3' : 'SEV4';
                    const mvpBadge = caseDetail.isMVP ? '<span class="badge-soft badge-soft--purple" style="padding:1px 6px; font-size:10px; margin-left:8px;">MVP</span>' : '';
                    pendingCasesHtml += `
                      <div class="pending-item">
                        <div class="pending-left">
                          <div class="pending-title">
                            ${caseDetail.caseNumber} - ${severityShort}${mvpBadge}
                          </div>
                          <div class="pending-sub">
                            ${caseDetail.account}
                          </div>
                          <div class="pending-meta">
                            Created: ${formatDateWithDayOfWeek(caseDetail.createdDate)} (${caseDetail.minutesSinceCreation}m ago)
                          </div>
                        </div>
                        <div class="pending-right">
                          ${caseDetail.remainingMinutes > 0 ? `${caseDetail.remainingMinutes}m remaining` : 'Due now'}
                        </div>
                      </div>
                    `;
                  });

                  noCasesHtml = `
                    <div class="no-cases-message" style="background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%); border: 2px solid #3b82f6;">
                      <h4 class="no-cases-title" style="color: #1d4ed8;">No Action Required</h4>
                      <p class="no-cases-text">${pendingCasesCount} case${pendingCasesCount === 1 ? ' is' : 's are'} within SLA window (SEV1: ${minSev1}min, SEV2: ${minSev2}min). Monitoring in progress...</p>
                      <div style="margin-top: 16px;">
                        <div style="font-weight: 600; color: #1d4ed8; margin-bottom: 8px; font-size: 14px;">Pending Cases:</div>
                        ${pendingCasesHtml}
                      </div>
                      <p class="mode-switch-hint" style="margin-top: 16px;">⏱️ Cases will appear here when they exceed SLA thresholds.</p>
                    </div>
                  `;
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
                container.innerHTML = noCasesHtml; // replace loading tile entirely
                container.classList.add('content-enter');
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

function getWeekendSignatureTemplate(caseSeverity = '') {
  return `Hi @,
New Sev${caseSeverity} case assigned to you & App is updated.
Thank You`;
}

function getGHOTemplate() {
  if (isCurrentlyWeekend()) {
    return `Hi @,
GHO (WOC) case assigned to you & App is updated.
Thank You`;
  }
  else {
    return `Hi @QB,
Greetings for the day.!

Kindly assist with GHO case assignment on this case.

Thank You
#GHOTriage`;
  }
}
// Helper function to format date with day of the week
function formatDateWithDayOfWeek(date) {
  const dateObj = new Date(date);
  const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const dayOfWeek = dayNames[dateObj.getDay()];
  return `${dayOfWeek}, ${dateObj.toLocaleString()}`;
}

// Function to determine shift based on a specific date/time
function getShiftForDate(date) {
  const dateObj = new Date(date);
  const hours = dateObj.getHours();
  const minutes = dateObj.getMinutes();
  const totalMinutes = hours * 60 + minutes;

  const apacStart = 5 * 60 + 30;
  const emeaStart = 12 * 60 + 30;
  const apacEnd = 12 * 60 + 30;
  const emeaEnd = 20 * 60;

  if (totalMinutes >= apacStart && totalMinutes < apacEnd) {
    return 'APAC';
  } else if (totalMinutes >= emeaStart && totalMinutes < emeaEnd) {
    return 'EMEA';
  } else {
    return 'AMER';
  }
}

// Function to check if a date is today
function isToday(date) {
  const today = new Date();
  const checkDate = new Date(date);
  return today.toDateString() === checkDate.toDateString();
}

// Function to determine current shift based on time of day
function getCurrentShift() {
  const now = new Date();
  const hours = now.getHours();
  const minutes = now.getMinutes();
  const totalMinutes = hours * 60 + minutes;

  const apacStart = 5 * 60 + 30;
  const emeaStart = 12 * 60 + 30;
  const apacEnd = 12 * 60 + 30;
  const emeaEnd = 20 * 60;

  if (totalMinutes >= apacStart && totalMinutes < apacEnd) {
    return 'APAC';
  } else if (totalMinutes >= emeaStart && totalMinutes < emeaEnd) {
    return 'EMEA';
  } else {
    return 'AMER';
  }
}
function getPreferredShiftValues(currentShift) {
  if (currentShift === 'APAC') {
    return ['APAC', 'IST'];
  } else if (currentShift === 'EMEA') {
    return ['EMEA', 'IST'];
  } else {
    return [currentShift];
  }
}

function buildPreferredShiftCondition(shiftValues) {
  if (shiftValues.length === 1) {
    return `Preferred_Shift__c='${shiftValues[0]}'`;
  } else {
    return `Preferred_Shift__c IN ('${shiftValues.join("','")}')`;
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
                      trackAction(comment.LastModifiedDate, caseRecord.CaseNumber, caseRecord.Severity_Level__c, 'GHO', caseRecord.CaseRoutingTaxonomy__r.Name.split('-')[0], currentMode, currentUserName, 'QB');
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

// Helper function to check if a case matches GHO criteria
function updateGHOButtonVisibility() {
  const ghoButton = document.getElementById("check-gho-button");
  if (ghoButton) {
    if (currentMode === 'signature') {
      ghoButton.style.display = 'inline-block';
    } else {
      ghoButton.style.display = 'none';
    }
  }
}

function updateCICButtonVisibility() {
  const cicButton = document.getElementById("cic-button");
  if (cicButton) {
    cicButton.style.display = isCurrentlyWeekend() ? 'inline-block' : 'none';
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

      case '1':
        e.preventDefault();
        const modeSwitch = document.getElementById("mode-switch");
        if (modeSwitch && modeSwitch.checked) {
          modeSwitch.checked = false;
          modeSwitch.dispatchEvent(new Event('change'));
        }
        break;

      case '2':
        e.preventDefault();
        const modeSwitchPremier = document.getElementById("mode-switch");
        if (modeSwitchPremier && !modeSwitchPremier.checked) {
          modeSwitchPremier.checked = true;
          modeSwitchPremier.dispatchEvent(new Event('change'));
        }
        break;

      case 'q':
        e.preventDefault();
        const modeSwitchSig = document.getElementById("mode-switch");
        if (modeSwitchSig && modeSwitchSig.checked) {
          modeSwitchSig.checked = false;
          modeSwitchSig.dispatchEvent(new Event('change'));
        }
        break;

      case 'w':
        e.preventDefault();
        const modeSwitchPrem = document.getElementById("mode-switch");
        if (modeSwitchPrem && !modeSwitchPrem.checked) {
          modeSwitchPrem.checked = true;
          modeSwitchPrem.dispatchEvent(new Event('change'));
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
        document.getElementById("search-button").disabled = true;
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
  // Ensure visibility per CSS rules (avoid :not(.modal-show) opacity:0)
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

  // Search input validation - enable/disable search button based on input
  function validateSearchInput() {
    const searchInput = document.getElementById("search-input");
    const searchButton = document.getElementById("search-button");
    const inputValue = searchInput.value.trim();

    if (inputValue.length > 0) {
      searchButton.disabled = false;
    } else {
      searchButton.disabled = true;
    }
  }

  // Debounced search function for real-time filtering
  function performRealTimeSearch() {
    clearTimeout(searchTimeout);
    searchTimeout = setTimeout(() => {
      const searchValue = document.getElementById("search-input").value.trim();
      applySearch(searchValue);
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
    if (event.key === "Enter" && !document.getElementById("search-button").disabled) {
      event.preventDefault();
      clearTimeout(searchTimeout); // Cancel any pending real-time search
      const searchValue = this.value.trim();
      applySearch(searchValue);
    }
  });

  document.getElementById("search-button").addEventListener("click", function () {
    // Safety check - prevent execution if button is disabled
    if (this.disabled) {
      return;
    }

    let searchValue = document.getElementById("search-input").value.trim();
    applySearch(searchValue);
  });

  document.getElementById("action-filter").addEventListener("change", function () {
    let filterValue = this.value;
    localStorage.setItem('caseFilter', filterValue);
    applyFilter(filterValue);
  });

  document.getElementById("mode-switch").addEventListener("change", function () {
    const headerTitle = document.querySelector(".header-title");
    if (this.checked) {
      currentMode = 'premier';
      headerTitle.textContent = 'Zerolag Tool - Premier Mode';
    } else {
      currentMode = 'signature';
      headerTitle.textContent = 'Zerolag Tool - Signature Mode';
    }
    localStorage.setItem('caseTriageMode', currentMode);

    // Update GHO button visibility based on mode
    updateGHOButtonVisibility();
    updateCICButtonVisibility();

    // Clear search and filter when changing modes
    document.getElementById("search-input").value = "";
    document.getElementById("search-button").disabled = true;
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
  if (currentMode === 'premier') {
    modeSwitch.checked = true;
    headerTitle.textContent = 'Zerolag Tool - Premier Mode';
  } else {
    headerTitle.textContent = 'Zerolag Tool - Signature Mode';
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
    if (!isCurrentlyWeekend()) {
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
      renderFilteredGHOCases(ghoRecordsGlobal, ghoConnectionGlobal, filterValue);
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
    if (e.target && e.target.classList.contains("gho-preview-btn")) {
      const templateText = getGHOTemplate();
      navigator.clipboard.writeText(templateText).then(function () {
        const toast = document.getElementById('toast');
        toast.textContent = 'GHO template copied to clipboard!';
        toast.style.display = 'block';
        setTimeout(function () {
          toast.style.display = 'none';
        }, 2000);
      }).catch(function (err) {
        console.error('Failed to copy GHO template: ', err);
        showToast('Failed to copy template to clipboard');
      });
    }
  });

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

//clear search input and reset filters
document.getElementById("clear-button").addEventListener("click", function () {
  document.getElementById("search-input").value = "";
  document.getElementById("search-button").disabled = true;

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

document.getElementById("parentSigSev2").addEventListener("click", function (e) {
  if (e.target && e.target.classList.contains("preview-record-btn")) {
    const button = e.target;
    const caseDiv = button.closest('.case-card');

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

// Function to show toast message
function showToast(message) {
  const toast = document.getElementById('toast');
  toast.textContent = message;
  toast.style.display = 'block';
  setTimeout(() => {
    toast.style.display = 'none';
  }, 3000);
}

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

function getCICColumnForShift(shift) {
  if (shift === 'APAC') return 'B';
  if (shift === 'EMEA') return 'J';
  return 'S';
}

function getTEColumnForShift(shift) {
  if (shift === 'APAC') return 'F';
  if (shift === 'EMEA') return 'O';
  return 'W';
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

function showCICManagers() {
  const shift = getCurrentShiftIST();
  const weekendDateStr = getWeekendLookupDateForShift(shift);
  const column = getCICColumnForShift(shift);

  const range = `'Service Cloud'!A:S`;

  const showLoading = () => {
    const loadingHtml = `
      <div id="cic-modal" class="modal-overlay" style="display:flex; z-index:1003;">
        <div class="modal-content" style="max-width: 600px;">
          <div class="modal-header" style="background: linear-gradient(135deg,#14b8a6 0%, #0ea5e9 100%); color: white;">
            <h3 style="color:white; margin:0;">Weekend Roster</h3>
            <span class="modal-close" id="cic-modal-close" style="color:white; cursor:pointer; font-size:24px;">&times;</span>
          </div>
          <div class="modal-body">
            <div class="loading-message" style="animation: fadeUp 300ms ease;">
              <div class="spinner"></div>
              <h4 style="margin-bottom:8px">Loading Weekend Roster...</h4>
              <p>${shift} · ${weekendDateStr}</p>
              <div class="skeleton-list">
                <div class="skeleton-card"></div>
                <div class="skeleton-card"></div>
                <div class="skeleton-card"></div>
              </div>
            </div>
          </div>
          <div class="modal-footer" style="padding:12px 16px; border-top:1px solid #e5e7eb; background:#f8fafc; color:#64748b; font-size:12px;">
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
          <div>
            <div style="font-weight:700; color:#0f766e; margin-bottom:6px;">${prettyTitle}</div>
    <div id="cic-names" title="Click to copy" style="padding:12px; border:1px solid #e2e8f0; border-radius:8px; background:#f8fafc; cursor:pointer; white-space:normal; word-break:break-word; overflow-wrap:anywhere;">
      ${formatNamesToMultiline(namesCell)}
            </div>
          </div>
        </div>`;

      const cell = document.getElementById('cic-names');
      cell.addEventListener('click', () => {
        navigator.clipboard.writeText(namesComma).then(() => {
          showToast('Weekend Roster copied');
        }).catch(() => showToast('Copy failed'));
      });

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
                    <div class="no-cases-message" style="animation: fadeUp 260ms ease; margin-top:12px;">
                      <h4 class="no-cases-title" style="color:#0f766e;">${sectionTitle}</h4>
                      <p class="no-cases-text">No TEs found for ${shift} on ${weekendDateStr}.</p>
                    </div>`;
                  bodyEl2.insertAdjacentHTML('beforeend', sectionHtml);
                  return;
                }

                const teNamesComma = normalizeNamesForCopy(teNamesCell);
                sectionHtml = `
                  <div style="display:flex; flex-direction:column; gap:12px; margin-top:12px;">
                    <div>
                      <div style="font-weight:700; color:#0f766e; margin-bottom:6px;">${sectionTitle}</div>
          <div id="${idSuffix}-names" title="Click to copy" style="padding:12px; border:1px solid #e2e8f0; border-radius:8px; background:#f8fafc; cursor:pointer; white-space:normal; word-break:break-word; overflow-wrap:anywhere;">${formatNamesToMultiline(teNamesCell)}</div>
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

          appendTESection('Sales Cloud', 'Sales Cloud TEs', 'sales-te');
          appendTESection('Service Cloud', 'Service Cloud TEs', 'service-te');
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

// Function to update status indicator based on case status
// Status indicator removed

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
      <div class="skeleton-list">
        <div class="skeleton-card"></div>
        <div class="skeleton-card"></div>
        <div class="skeleton-card"></div>
      </div>
    </div>
  `;

  // Get connection and execute GHO query
  let ghoConn = new jsforce.Connection({
    serverUrl: 'https://orgcs.my.salesforce.com',
    sessionId: SESSION_ID,
  });

  // Get current shift dynamically and use in GHO query
  const currentShift = getCurrentShift();
  const preferredShiftValues = getPreferredShiftValues(currentShift);
  const shiftCondition = buildPreferredShiftCondition(preferredShiftValues);

  const ghoQuery = `SELECT Id, CreatedDate, Account.Name, Owner.Name, SE_Target_Response__c, Severity_Level__c, CaseNumber, Subject, CaseRoutingTaxonomy__r.Name, SE_Initial_Response_Status__c, Contact.Is_MVP__c, support_available_timezone__c, (SELECT Transfer_Reason__c, CreatedDate, CreatedById, Preferred_Shift_Old_Value__c, Preferred_Shift_New_Value__c, Severity_New_Value__c, Severity_Old_Value__c FROM Case_Routing_Logs__r ORDER BY CreatedDate DESC LIMIT 10) FROM Case WHERE ((Owner.Name IN ('Skills Queue','Kase Changer', 'Working in Org62','GHO Queue') AND ((Case_Support_level__c='Premier Priority' OR Case_Support_level__c='Signature' OR Case_Support_level__c='Signature Success') OR (Case_Support_level__c='Signature' OR Case_Support_level__c='Premier Priority' OR Case_Support_level__c='Signature Success'))) OR (Contact.Is_MVP__c=true AND Owner.Name='GHO Queue')) AND IsClosed=false AND ${shiftCondition} AND ((CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry%' OR CaseRoutingTaxonomy__r.Name LIKE 'Community-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Scale Center%' OR CaseRoutingTaxonomy__r.Name LIKE 'Customer Success Score%' OR CaseRoutingTaxonomy__r.Name LIKE 'Data Cloud-%') AND (Severity_Level__c='Level 1 - Critical' OR Severity_Level__c='Level 2 - Urgent')) AND CaseRoutingTaxonomy__r.Name NOT IN ('Disability and Product Accessibility','DORA')`;

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

    // Check for and track any GHO triage actions before showing the modal
    if (result.records && result.records.length > 0) {
      const caseIds = result.records.map(record => record.Id);
      const commentQuery = `SELECT ParentId, Body, CreatedById, LastModifiedDate FROM CaseFeed WHERE Visibility = 'InternalUsers' AND ParentId IN ('${caseIds.join("','")}') AND Type = 'TextPost'`;

      ghoConn.query(commentQuery, function (commentErr, commentResult) {
        if (!commentErr && commentResult.records) {
          // Get current user ID for tracking
          ghoConn.query(`SELECT Id FROM User WHERE Name = '${currentUserName}' AND IsActive = True AND Username LIKE '%orgcs.com'`, function (userErr, userResult) {
            let currentUserId = null;
            if (!userErr && userResult.records.length > 0) {
              currentUserId = userResult.records[0].Id;
            }

            if (currentUserId) {
              commentResult.records.forEach(comment => {
                if (comment.Body && comment.Body.includes('#GHOTriage') && comment.CreatedById === currentUserId) {
                  // Check if comment is from today and same GEO
                  const commentDate = comment.LastModifiedDate || comment.CreatedDate;
                  const isCommentFromToday = isToday(commentDate);
                  const commentShift = getShiftForDate(commentDate);
                  const currentShift = getCurrentShift();
                  const isSameGeo = commentShift === currentShift;

                  console.log('GHO Status Check - #GHOTriage comment analysis:', {
                    caseId: comment.ParentId,
                    commentDate,
                    isFromToday: isCommentFromToday,
                    commentShift,
                    currentShift,
                    isSameGeo
                  });

                  // Only process if comment is from today and same GEO
                  if (isCommentFromToday && isSameGeo) {
                    const caseRecord = result.records.find(c => c.Id === comment.ParentId);
                    if (caseRecord) {
                      const ghoTrackingKey = `gho_tracked_${caseRecord.Id}`;
                      if (!localStorage.getItem(ghoTrackingKey)) {
                        trackAction(comment.LastModifiedDate, caseRecord.CaseNumber, caseRecord.Severity_Level__c, 'GHO', caseRecord.CaseRoutingTaxonomy__r.Name.split('-')[0], currentMode, currentUserName, 'QB');
                        localStorage.setItem(ghoTrackingKey, 'true');
                        console.log('GHO status check - GHO triage action tracked for case:', caseRecord.CaseNumber, 'Date:', commentDate, 'Shift:', commentShift);
                      }
                    }
                  } else {
                    console.log('GHO Status Check - #GHOTriage comment ignored - not from today/same GEO:', {
                      caseId: comment.ParentId,
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

            // Show modal with results
            showGHOStatusModal(result.records, ghoConn);
          });
        } else {
          // Show modal even if comment query fails
          showGHOStatusModal(result.records, ghoConn);
        }
      });
    } else {
      // Show modal with empty results
      showGHOStatusModal(result.records, ghoConn);
    }
  });
}

// Helper function to get user names from User IDs
function getUserNames(userIds, conn, callback) {
  if (!userIds || userIds.length === 0) {
    callback({});
    return;
  }

  const userQuery = `SELECT Id, Name FROM User WHERE Id IN ('${userIds.join("','")}') AND Username LIKE '%orgcs.com'`;

  conn.query(userQuery, function (err, result) {
    if (err) {
      console.error('Error fetching user names:', err);
      callback({});
      return;
    }

    const userMap = {};
    if (result.records) {
      result.records.forEach(user => {
        userMap[user.Id] = user.Name;
      });
    }
    callback(userMap);
  });
}

function showGHOStatusModal(ghoRecords, conn) {
  // Store globally for filtering
  ghoRecordsGlobal = ghoRecords;
  ghoConnectionGlobal = conn;

  const container = document.getElementById('gho-cases-container');
  const currentShift = getCurrentShift();

  if (!ghoRecords || ghoRecords.length === 0) {
    container.innerHTML = `
      <div class="no-cases-message" style="animation: fadeUp 320ms ease;">
        <h4 class="no-cases-title" style="color: #1d4ed8;">No GHO Cases Found</h4>
        <p class="no-cases-text">No cases matching GHO criteria were found for <strong>${currentShift}</strong> shift. Great work!</p>
        <p class="mode-switch-hint">✅ All GHO cases are up to date for current shift.</p>
      </div>
    `;
    return;
  }

  // Render with current filter
  const filterValue = document.getElementById('gho-taxonomy-filter').value;
  renderFilteredGHOCases(ghoRecords, conn, filterValue);
}

// Function to filter and render GHO cases based on CaseRoutingTaxonomy__r.Name
function renderFilteredGHOCases(ghoRecords, conn, filterValue = 'All') {
  const container = document.getElementById('gho-cases-container');
  const currentShift = getCurrentShift();

  // Filter records based on CaseRoutingTaxonomy__r.Name
  let filteredRecords = ghoRecords;
  if (filterValue !== 'All') {
    filteredRecords = ghoRecords.filter(record => {
      const taxonomyName = record.CaseRoutingTaxonomy__r?.Name || '';
      return taxonomyName.toLowerCase().startsWith(filterValue.toLowerCase() + '-');
    });
  }

  if (filteredRecords.length === 0) {
    container.innerHTML = `
      <div class="no-cases-message" style="animation: fadeUp 320ms ease;">
        <h4 class="no-cases-title" style="color: #1d4ed8;">No ${filterValue === 'All' ? '' : filterValue + ' '}GHO Cases Found</h4>
        <p class="no-cases-text">No cases matching ${filterValue === 'All' ? 'GHO criteria' : filterValue + ' taxonomy'} were found for <strong>${currentShift}</strong> shift.</p>
        <p class="mode-switch-hint">✅ All ${filterValue === 'All' ? 'GHO' : filterValue} cases are up to date for current shift.</p>
      </div>
    `;
    return;
  }

  // Query comments for all filtered cases to check for #GHOTriage
  const caseIds = filteredRecords.map(record => record.Id);
  const commentQuery = `SELECT ParentId, Body, CreatedById, LastModifiedDate FROM CaseFeed WHERE Visibility = 'InternalUsers' AND ParentId IN ('${caseIds.join("','")}') AND Type = 'TextPost'`;

  conn.query(commentQuery, function (commentErr, commentResult) {
    if (commentErr) {
      console.error('Error querying GHO comments:', commentErr);
      // Continue rendering without comment info
      renderGHOCasesWithCommentInfo(filteredRecords, conn, currentShift, filterValue, new Set());
      return;
    }

    // Check which cases have #GHOTriage comments
    const ghoTriageCommentCases = new Set();
    if (commentResult.records) {
      commentResult.records.forEach(comment => {
        if (comment.Body && comment.Body.includes('#GHOTriage')) {
          // Check if comment is from today and same GEO
          const commentDate = comment.LastModifiedDate || comment.CreatedDate;
          const isCommentFromToday = isToday(commentDate);
          const commentShift = getShiftForDate(commentDate);
          const isSameGeo = commentShift === currentShift;

          console.log('GHO Render - #GHOTriage comment analysis:', {
            caseId: comment.ParentId,
            commentDate,
            isFromToday: isCommentFromToday,
            commentShift,
            currentShift,
            isSameGeo
          });

          // Only consider #GHOTriage comments from today and same GEO
          if (isCommentFromToday && isSameGeo) {
            ghoTriageCommentCases.add(comment.ParentId);
            console.log('GHO Render - Found valid #GHOTriage comment for case:', comment.ParentId, 'Date:', commentDate, 'Shift:', commentShift);
          } else {
            console.log('GHO Render - #GHOTriage comment ignored - not from today/same GEO:', {
              caseId: comment.ParentId,
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

    console.log('GHO Cases with #GHOTriage comments:', Array.from(ghoTriageCommentCases));

    // Render cases with comment information
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

  // Get user names for all user IDs
  getUserNames(Array.from(allUserIds), conn, function (userMap) {
    // Count cases with QB mentions
    const qbMentionedCount = Array.from(ghoTriageCommentCases).length;

    let ghoHtml = `
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
    `;

    filteredRecords.forEach(caseRecord => {
      const caseId = caseRecord.Id;
      const isMVP = caseRecord.Contact && caseRecord.Contact.Is_MVP__c === true;
      const hasGHOTriage = ghoTriageCommentCases.has(caseId);

      // Determine status color
      let statusColor = '';
      if (caseRecord.SE_Initial_Response_Status__c === 'Met') {
        statusColor = 'green';
      } else if (caseRecord.SE_Initial_Response_Status__c === 'In Warning' || caseRecord.SE_Initial_Response_Status__c === 'Warning') {
        statusColor = 'red';
      }

      // Build routing log HTML and GHO transfer information
      let routingLogHtml = '';
      let ghoTransferHtml = '';
      const routingLogs = caseRecord.Case_Routing_Logs__r;

      if (routingLogs && routingLogs.totalSize > 0) {
        const lastLog = routingLogs.records[0];
        if (lastLog.Transfer_Reason__c && lastLog.Transfer_Reason__c !== 'New') {
          let logText = lastLog.Transfer_Reason__c;

          // Check for severity change
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

        // Process GHO transfers and create nested list
        const allGhoTransfers = routingLogs.records.filter(log =>
          log.Transfer_Reason__c === 'GHO'
        ).sort((a, b) => new Date(b.CreatedDate) - new Date(a.CreatedDate)); // Sort by CreatedDate DESC

        if (allGhoTransfers.length > 0) {
          // Separate transfers by current shift and others
          const currentShiftTransfers = allGhoTransfers.filter(transfer =>
            transfer.Preferred_Shift_Old_Value__c === currentShift
          );
          const otherTransfers = allGhoTransfers.filter(transfer =>
            transfer.Preferred_Shift_Old_Value__c !== currentShift
          );

          const caseUniqueId = `gho-${caseRecord.Id}`;
          let nestedListHtml = '';

          // Default view: Current shift transfers
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

          // Add expansion button if there are other transfers
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

      const caseHtml = `
    <div class="case-card ${isMVP ? 'mvp-case card-accent-purple' : ''}" style="margin-bottom: 16px;">
          <div class="case-header">
            <div style="display: flex; align-items: center; gap: 8px;">
      ${isMVP ? '<span class="badge-soft badge-soft--purple">MVP</span>' : ''}
      <span class="badge-soft badge-soft--amber">GHO</span>
      ${hasGHOTriage ? '<span class="badge-soft badge-soft--success">QB Mentioned</span>' : ''}
              <h3 class="case-title">${caseRecord.Subject}</h3>
            </div>
            <div class="case-timestamp">${formatDateWithDayOfWeek(caseRecord.CreatedDate)}<br/>
             (${timeElapsed(new Date(caseRecord.CreatedDate))})</div>
          </div>
          
          <div class="case-details">
            <div class="case-info">
              <div class="case-info-item">
                <span class="checkmark">✓</span>
                <span>${caseRecord.Account.Name}</span>
              </div>
              <div class="case-info-item">
                <span class="checkmark">✓</span>
                <span>${caseRecord.Owner.Name}</span>
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
              ${hasGHOTriage ? '<div class="case-info-item"><span class="checkmark" style="color: #059669;">✓</span><span style="color: #059669; font-weight: bold;">QB has been mentioned (#GHOTriage found in comments)</span></div>' : ''}
              ${routingLogHtml}
              ${ghoTransferHtml}
            </div>
            
            <div class="case-actions">
          <a target="_blank" href="https://orgcs.my.salesforce.com/lightning/r/Case/${caseId}/view" 
            class="preview-btn gho-preview-btn" 
            data-case-id="${caseId}">
                View Case Record
              </a>
            </div>
          </div>
        </div>
      `;

      ghoHtml += caseHtml;
    });

    container.innerHTML = ghoHtml;
  });
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
}// Add debug functions to global scope for console access
window.debugPersistentCases = debugPersistentCases;
window.debugPersistentCasesSimple = debugPersistentCasesSimple;
window.clearGHOAlertFlags = clearGHOAlertFlags;
