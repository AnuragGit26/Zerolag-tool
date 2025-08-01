import { timeElapsed, addMinutes, isCurrentlyWeekend } from './utils/datetime.js';
import { applyFilter, applySearch, updateWeekendModeIndicator } from './utils/dom.js';
import { trackAction } from './utils/api.js';

let SESSION_ID;
let currentMode = localStorage.getItem('caseTriageMode') || 'signature';
let currentUserName;
let searchTimeout;
export const SPREADSHEET_ID = '1BKxQLGFrczjhcx9rEt-jXGvlcCPQblwBhFJjoiDD7TI';

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

      // Auto-open tab even on error for user to try toggle or troubleshoot
      var data = 'openTabSilent';
      chrome.runtime.sendMessage(data, function (response) {
        console.log('response-----' + response);
      });
      return console.error('error---' + err);
    } else {
      currentUserName = res.display_name;
      let displayedCaseCount = 0;
      let actionTakenCount = 0;

      // Get current shift dynamically
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
            document.getElementById("parentSigSev2").innerHTML = noCasesHtml;
            // Auto-open tab without focus for easy mode switching
            var data = 'openTabSilent';
            chrome.runtime.sendMessage(data, function (response) {
              console.log('response-----' + response);
            });
            return;
          }

          const caseIds = result.records.map(record => record.Id);
          const caseMap = new Map(result.records.map(record => [record.Id, record]));

          // Fetch case history and track actions
          const combinedHistoryQuery = `SELECT Field, CaseId, CreatedById, CreatedDate, NewValue FROM CaseHistory WHERE CaseId IN ('${caseIds.join("','")}') AND (Field = 'Routing_Status__c' OR (Field = 'Owner' AND CreatedById = '${currentUserId}')) ORDER BY CreatedDate DESC`;

          conn.query(combinedHistoryQuery, function (historyErr, historyResult) {
            if (historyErr) {
              return console.error('Error fetching case history:', historyErr);
            }

            if (!historyResult.records || historyResult.records.length === 0) {
              return;
            }

            // Separate the results by field type
            const routingStatusHistories = historyResult.records.filter(history => history.Field === 'Routing_Status__c');
            const ownerAssignmentHistories = historyResult.records.filter(history => history.Field === 'Owner' && history.CreatedById === currentUserId);

            // Process routing status histories (original historyQuery logic)
            const manuallyAssignedHistories = routingStatusHistories.filter(
              history => history.NewValue && typeof history.NewValue === 'string' && history.NewValue.startsWith('Manually Assigned')
            );

            const firstAssignments = new Map();
            for (const history of manuallyAssignedHistories) {
              if (!firstAssignments.has(history.CaseId)) {
                firstAssignments.set(history.CaseId, history);
              }
            }

            // Combine both filtering logics into a single processing loop
            const allTrackableHistories = [];

            // Add routing status histories that meet the criteria
            if (firstAssignments.size > 0) {
              for (const [caseId, history] of firstAssignments.entries()) {
                if (history.CreatedById === currentUserId) {
                  allTrackableHistories.push({
                    ...history,
                    actionType: 'New Case',
                    trackingKey: `tracked_premier_assignment_${caseId}`,
                    assignedTo: '' // Empty for routing status changes
                  });
                }
              }
            }

            // Add owner assignment histories that meet the criteria
            if (ownerAssignmentHistories.length > 0) {
              for (const assignment of ownerAssignmentHistories) {
                if (assignment.NewValue) {
                  allTrackableHistories.push({
                    ...assignment,
                    actionType: 'New Case',
                    trackingKey: `tracked_assignment_${assignment.CaseId}_${assignment.CreatedDate}`,
                    assignedTo: assignment.NewValue
                  });
                }
              }
            }

            // Single loop to process all trackable histories
            for (const historyItem of allTrackableHistories) {
              const caseRecord = caseMap.get(historyItem.CaseId);
              if (caseRecord) {
                if (!localStorage.getItem(historyItem.trackingKey)) {
                  trackAction(
                    historyItem.CreatedDate,
                    caseRecord.CaseNumber,
                    caseRecord.Severity_Level__c,
                    historyItem.actionType,
                    caseRecord.CaseRoutingTaxonomy__r.Name.split('-')[0],
                    currentMode,
                    currentUserName,
                    historyItem.assignedTo
                  );
                  localStorage.setItem(historyItem.trackingKey, 'true');
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
                      }
                    }
                  }
                }

                // Track GHO Triage actions
                if (record.Body && record.Body.includes('#GHOTriage')) {
                  console.log('GHO Triage action found:', record.ParentId);
                  if (record.CreatedById === currentUserId) {
                    const caseRecord = result.records.find(c => c.Id === record.ParentId);
                    console.log('Case record for GHO triage comment:', caseRecord);
                    if (caseRecord) {
                      const ghoTrackingKey = `gho_tracked_${caseRecord.Id}`;
                      if (!localStorage.getItem(ghoTrackingKey)) {
                        // Track to Google Sheet with "GHO" prefix for cloud type to distinguish from regular cases
                        trackAction(record.LastModifiedDate, caseRecord.CaseNumber, caseRecord.Severity_Level__c, 'GHO', caseRecord.CaseRoutingTaxonomy__r.Name.split('-')[0], currentMode, currentUserName, 'QB');
                        localStorage.setItem(ghoTrackingKey, 'true');
                        console.log('GHO triage action tracked for case:', caseRecord.CaseNumber);
                      }
                    }
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

              // Set total cases count based on raw result before filtering for display
              totalCasesCount = result.records.length;

              const filteredRecords = result.records.filter(x => {
                // This filter logic removes MVP cases with 'Met' status from being displayed as actionable alerts
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

                  // Check for GEO Locate routing log with Met SLA and auto-mark as actioned
                  const routingLogs = caseRecord.Case_Routing_Logs__r;
                  if (routingLogs && routingLogs.totalSize > 0) {
                    const lastLog = routingLogs.records[0];
                    if (lastLog.Transfer_Reason__c === 'GEO Locate' && caseRecord.SE_Initial_Response_Status__c === 'Met') {
                      localStorage.setItem(caseId, 'true');
                      console.log(`Auto-actioned case ${caseRecord.CaseNumber} due to GEO Locate routing with Met SLA`);
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

                  let statusColor = 'red'; // Warning status always red
                  let routingLogHtml = '';
                  // routingLogs already declared above, reuse it
                  if (routingLogs && routingLogs.totalSize > 0) {
                    const lastLog = routingLogs.records[0];
                    if (lastLog.Transfer_Reason__c && lastLog.Transfer_Reason__c !== 'New') {
                      routingLogHtml = `
                        <div class="case-info-item">
                          <span class="checkmark">✓</span>
                          <span style="color: #9F2B68;">${lastLog.Transfer_Reason__c} (${timeElapsed(new Date(lastLog.CreatedDate))})</span>
                        </div>
                      `;
                    }
                  }

                  const mvpCaseHtml = `
                    <div class="case-card mvp-warning-case" style="border: 3px solid #ef4444; background-color: #fef2f2;">
                      <div class="case-header">
                        <div style="display: flex; align-items: center; gap: 8px;">
                          <span style="background-color: #ef4444; color: white; padding: 2px 8px; border-radius: 12px; font-size: 12px; font-weight: bold;">MVP URGENT</span>
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
                          <div class="case-info-item">
                            <span class="checkmark">✓</span>
                            <span style="color: ${statusColor}; font-weight: bold;">${caseRecord.SE_Initial_Response_Status__c} - MVP CASE</span>
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

              // Second pass: Process regular cases (using filtered records)
              for (var x in filteredRecords) {
                const caseId = filteredRecords[x].Id;

                // Skip MVP Warning cases as they're already processed above
                if (filteredRecords[x].Contact && filteredRecords[x].Contact.Is_MVP__c === true &&
                  (filteredRecords[x].SE_Initial_Response_Status__c === 'In Warning' || filteredRecords[x].SE_Initial_Response_Status__c === 'Warning')) {
                  continue;
                }

                if (actionedCaseIds.has(caseId)) {
                  localStorage.setItem(caseId, 'true');
                }

                // Check for GEO Locate routing log with Met SLA and auto-mark as actioned
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

                // Check if case meets alert criteria (timing + severity)
                const meetsAlertCriteria = (filteredRecords[x].CaseRoutingTaxonomy__r.Name == 'Sales-Issues Developing for Salesforce Functions (Product)') ||
                  (today >= addMinutes(minSev1, new Date(filteredRecords[x].CreatedDate)) && filteredRecords[x].Severity_Level__c == 'Level 1 - Critical') ||
                  (today >= addMinutes(minSev2, new Date(filteredRecords[x].CreatedDate)) && filteredRecords[x].Severity_Level__c == 'Level 2 - Urgent') ||
                  (filteredRecords[x].Contact.isMVP === true && filteredRecords[x].SE_Initial_Response_Status__c !== 'Met' && today >= addMinutes(minSev1, new Date(filteredRecords[x].CreatedDate)) && filteredRecords[x].Severity_Level__c == 'Level 1 - Critical') ||
                  (filteredRecords[x].Contact.isMVP === true && filteredRecords[x].SE_Initial_Response_Status__c !== 'Met' && today >= addMinutes(minSev2, new Date(filteredRecords[x].CreatedDate)) && filteredRecords[x].Severity_Level__c == 'Level 2 - Urgent') ||
                  (filteredRecords[x].Contact.isMVP === true && filteredRecords[x].SE_Initial_Response_Status__c !== 'Met' && today >= addMinutes(210, new Date(filteredRecords[x].CreatedDate)) && filteredRecords[x].Severity_Level__c == 'Level 3 - High') ||
                  (filteredRecords[x].Contact.isMVP === true && filteredRecords[x].SE_Initial_Response_Status__c !== 'Met' && today >= addMinutes(430, new Date(filteredRecords[x].CreatedDate)) && filteredRecords[x].Severity_Level__c == 'Level 4 - Medium');

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
                      routingLogHtml = `
                        <div class="case-info-item">
                          <span class="checkmark">✓</span>
                          <span style="color: #9F2B68;">${lastLog.Transfer_Reason__c} (${timeElapsed(new Date(lastLog.CreatedDate))})</span>
                        </div>
                      `;
                    }
                  }

                  const newHtml = `
                    <div class="case-card ${caseData.isMVP ? 'mvp-case' : ''}" ${caseData.isMVP ? 'style="border-left: 4px solid #9333ea;"' : ''}>
                      <div class="case-header">
                        <div style="display: flex; align-items: center; gap: 8px;">
                          ${caseData.isMVP ? '<span style="background-color: #9333ea; color: white; padding: 2px 8px; border-radius: 12px; font-size: 12px; font-weight: bold;">MVP</span>' : ''}
                          ${caseData.isSLAM ? '<span style="background-color: #d9534f; color: white; padding: 2px 8px; border-radius: 12px; font-size: 12px; font-weight: bold;">SLAM</span>' : ''}
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
                            <span style="color: ${statusColor}">${filteredRecords[x].SE_Initial_Response_Status__c}${caseData.isMVP ? ' - MVP CASE' : ''}</span>
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
                  // Case exists but doesn't meet alert criteria yet
                  pendingCasesCount++;
                  const minutesSinceCreation = Math.floor((today - new Date(filteredRecords[x].CreatedDate)) / (1000 * 60));
                  const requiredMinutes = filteredRecords[x].Severity_Level__c === 'Level 1 - Critical' ? minSev1 : filteredRecords[x].Severity_Level__c === 'Level 2 - Urgent' ? minSev2 : filteredRecords[x].Severity_Level__c === 'Level 3 - High' ? 210 : 430;
                  const remainingMinutes = requiredMinutes - minutesSinceCreation;

                  pendingCasesDetails.push({
                    caseNumber: filteredRecords[x].CaseNumber,
                    severity: filteredRecords[x].Severity_Level__c,
                    account: filteredRecords[x].Account.Name,
                    createdDate: new Date(filteredRecords[x].CreatedDate),
                    minutesSinceCreation: minutesSinceCreation,
                    remainingMinutes: Math.max(0, remainingMinutes),
                    isMVP: filteredRecords[x].Contact && filteredRecords[x].Contact.Is_MVP__c === true
                  });

                  console.log('Case pending alert criteria:', filteredRecords[x].CaseNumber,
                    'Severity:', filteredRecords[x].Severity_Level__c,
                    'Created:', new Date(filteredRecords[x].CreatedDate),
                    'Time since creation:', minutesSinceCreation, 'minutes',
                    'Remaining:', remainingMinutes, 'minutes');
                }
              }

            }
            // The condition now checks if there is actual HTML content to display, not just if variables are defined.
            if (isData && (mvpWarningHtml || myHtml)) {
              // Display MVP warning cases first, then regular cases
              let finalHtml = '';
              if (mvpWarningHtml) {
                finalHtml += mvpWarningHtml;
              }
              if (myHtml) {
                finalHtml += myHtml;
              }

              // Always show pending cases section if there are any pending cases
              if (pendingCasesCount > 0) {
                let pendingCasesHtml = '';
                pendingCasesDetails.forEach(caseDetail => {
                  const severityShort = caseDetail.severity.includes('Level 1') ? 'SEV1' :
                    caseDetail.severity.includes('Level 2') ? 'SEV2' :
                      caseDetail.severity.includes('Level 3') ? 'SEV3' : 'SEV4';
                  const mvpBadge = caseDetail.isMVP ? '<span style="background-color: #9333ea; color: white; padding: 1px 6px; border-radius: 8px; font-size: 10px; margin-left: 8px;">MVP</span>' : '';

                  pendingCasesHtml += `
                    <div style="background: rgba(59, 130, 246, 0.1); border: 1px solid rgba(59, 130, 246, 0.3); border-radius: 8px; padding: 12px; margin: 8px 0; display: flex; justify-content: space-between; align-items: center;">
                      <div style="flex: 1;">
                        <div style="font-weight: 600; color: #1d4ed8; font-size: 14px;">
                          ${caseDetail.caseNumber} - ${severityShort}${mvpBadge}
                        </div>
                        <div style="color: #374151; font-size: 12px; margin-top: 2px;">
                          ${caseDetail.account}
                        </div>
                        <div style="color: #6b7280; font-size: 11px; margin-top: 2px;">
                          Created: ${formatDateWithDayOfWeek(caseDetail.createdDate)} (${caseDetail.minutesSinceCreation}m ago)
                        </div>
                      </div>
                      <div style="text-align: right; color: #1d4ed8; font-weight: 600; font-size: 12px;">
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

              document.getElementById("parentSigSev2").innerHTML += finalHtml;

              const savedFilter = localStorage.getItem('caseFilter');
              if (savedFilter) {
                applyFilter(savedFilter);
              }

              // Check if there are cases without action taken
              const hasUnactionedCases = actionTakenCount < displayedCaseCount;
              updateStatusIndicator(hasUnactionedCases, displayedCaseCount, actionTakenCount);

              if (hasUnactionedCases) {
                // There are unactioned cases - play sound and focus tab
                var audio = new Audio('../assets/audio/notification.wav');
                audio.play();
                var data = 'openTab';
                chrome.runtime.sendMessage(data, function (response) {
                  console.log('response-----' + response);
                });
              } else {
                // All cases have action taken - silent mode
                var data = 'openTabSilent';
                chrome.runtime.sendMessage(data, function (response) {
                  console.log('response-----' + response);
                });
              }
            } else {
              // No cases meet alert criteria - determine what to show
              let noCasesHtml;

              if (pendingCasesCount > 0) {
                // Cases exist but none meet alert criteria yet (still within SLA window)
                let pendingCasesHtml = '';
                pendingCasesDetails.forEach(caseDetail => {
                  const severityShort = caseDetail.severity.includes('Level 1') ? 'SEV1' :
                    caseDetail.severity.includes('Level 2') ? 'SEV2' :
                      caseDetail.severity.includes('Level 3') ? 'SEV3' : 'SEV4';
                  const mvpBadge = caseDetail.isMVP ? '<span style="background-color: #9333ea; color: white; padding: 1px 6px; border-radius: 8px; font-size: 10px; margin-left: 8px;">MVP</span>' : '';

                  pendingCasesHtml += `
                    <div style="background: rgba(59, 130, 246, 0.1); border: 1px solid rgba(59, 130, 246, 0.3); border-radius: 8px; padding: 12px; margin: 8px 0; display: flex; justify-content: space-between; align-items: center;">
                      <div style="flex: 1;">
                        <div style="font-weight: 600; color: #1d4ed8; font-size: 14px;">
                          ${caseDetail.caseNumber} - ${severityShort}${mvpBadge}
                        </div>
                        <div style="color: #374151; font-size: 12px; margin-top: 2px;">
                          ${caseDetail.account}
                        </div>
                        <div style="color: #6b7280; font-size: 11px; margin-top: 2px;">
                            Created: ${formatDateWithDayOfWeek(caseDetail.createdDate)} (${caseDetail.minutesSinceCreation}m ago)
                        </div>
                      </div>
                      <div style="text-align: right; color: #1d4ed8; font-weight: 600; font-size: 12px;">
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
                // Cases exist but none require action (all are assigned/handled or filtered out)
                noCasesHtml = `
                  <div class="no-cases-message">
                    <h4 class="no-cases-title">No Cases for Now</h4>
                    <p class="no-cases-text">All ${totalCasesCount} case${totalCasesCount === 1 ? ' is' : 's are'} assigned or handled. Check back later!</p>
                    <p class="mode-switch-hint">💡 Try switching between Signature and Premier modes to see different case types.</p>
                  </div>
                `;
              } else {
                // Truly no cases exist
                noCasesHtml = `
                  <div class="no-cases-message">
                    <h4 class="no-cases-title">No Cases to Action</h4>
                    <p class="no-cases-text">All cases are up to date. Great work!</p>
                    <p class="mode-switch-hint">💡 Try switching between Signature and Premier modes to see different case types.</p>
                  </div>
                `;
              }

              document.getElementById("parentSigSev2").innerHTML += noCasesHtml;
              updateStatusIndicator(false, 0, 0);
              // No cases - silent mode
              var data = 'openTabSilent';
              chrome.runtime.sendMessage(data, function (response) {
                console.log('response-----' + response);
              });
            }
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
    return 'APAC'; // 5:30 AM - 12:30 PM IST
  } else if (totalMinutes >= emeaStart && totalMinutes < emeaEnd) {
    return 'EMEA'; // 12:30 PM - 8:00 PM IST
  } else {
    return 'AMER'; // 8:00 PM - 5:30 AM IST (next day)
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

  const apacAlertTime = 7 * 60 + 30;  // 7:30 AM IST
  const emeaAlertTime = 14 * 60 + 30; // 2:30 PM IST (14:30)
  const amerAlertTime = 22 * 60 + 30; // 10:30 PM IST (22:30)

  const isAPACTime = Math.abs(currentTime - apacAlertTime) <= 5;
  const isEMEATime = Math.abs(currentTime - emeaAlertTime) <= 5;
  const isAMERTime = Math.abs(currentTime - amerAlertTime) <= 5;

  if (!(isAPACTime || isEMEATime || isAMERTime)) {
    return;
  }

  // Determine which region alert this is
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

  const ghoQuery = `SELECT Id, CreatedDate, Account.Name, Owner.Name, SE_Target_Response__c, Severity_Level__c, CaseNumber, Subject, CaseRoutingTaxonomy__r.Name, SE_Initial_Response_Status__c, Contact.Is_MVP__c, support_available_timezone__c, (SELECT Transfer_Reason__c, CreatedDate, CreatedById, Preferred_Shift_Old_Value__c, Preferred_Shift_New_Value__c FROM Case_Routing_Logs__r ORDER BY CreatedDate DESC LIMIT 10) FROM Case WHERE ((Owner.Name IN ('Skills Queue','Kase Changer', 'Working in Org62','GHO Queue') AND ((Case_Support_level__c='Premier Priority' OR Case_Support_level__c='Signature' OR Case_Support_level__c='Signature Success') OR (Case_Support_level__c='Signature' OR Case_Support_level__c='Premier Priority' OR Case_Support_level__c='Signature Success'))) OR (Contact.Is_MVP__c=true AND Owner.Name='GHO Queue')) AND IsClosed=false AND ${shiftCondition} AND ((CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry%' OR CaseRoutingTaxonomy__r.Name LIKE 'Community-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Scale Center%' OR CaseRoutingTaxonomy__r.Name LIKE 'Customer Success Score%' OR CaseRoutingTaxonomy__r.Name LIKE 'Data Cloud-%') AND (Severity_Level__c='Level 1 - Critical' OR Severity_Level__c='Level 2 - Urgent')) AND CaseRoutingTaxonomy__r.Name NOT IN ('Disability and Product Accessibility','DORA')`;

  conn.query(ghoQuery, function (err, result) {
    if (err) {
      console.error('GHO Alert Query Error:', err);
      return;
    }

    if (!result.records || result.records.length === 0) {
      localStorage.setItem(alertKey, 'true');
      return;
    }

    // Check if any cases need action (not all have #GHOTriage)
    const caseIds = result.records.map(record => record.Id);
    const commentQuery = `SELECT ParentId, Body, CreatedById, LastModifiedDate FROM CaseFeed WHERE Visibility = 'InternalUsers' AND ParentId IN ('${caseIds.join("','")}') AND Type = 'TextPost'`;

    conn.query(commentQuery, function (commentErr, commentResult) {
      if (commentErr) {
        console.error('GHO Comment Query Error:', commentErr);
        return;
      }

      // Get current user ID for tracking
      conn.query(`SELECT Id FROM User WHERE Name = '${currentUserName}' AND IsActive = True AND Username LIKE '%orgcs.com'`, function (userErr, userResult) {
        let currentUserId = null;
        if (!userErr && userResult.records.length > 0) {
          currentUserId = userResult.records[0].Id;
        }

        // Get cases that have #GHOTriage comments - filter using JavaScript
        const actionedCaseIds = new Set();
        if (commentResult.records) {
          commentResult.records.forEach(comment => {
            if (comment.Body && comment.Body.includes('#GHOTriage')) {
              actionedCaseIds.add(comment.ParentId);

              // Track GHO Triage actions if current user made the comment
              if (currentUserId && comment.CreatedById === currentUserId) {
                const caseRecord = result.records.find(c => c.Id === comment.ParentId);
                console.log('GHO alert - Case record for triage comment:', caseRecord);
                if (caseRecord) {
                  const ghoTrackingKey = `gho_tracked_${caseRecord.Id}`;
                  if (!localStorage.getItem(ghoTrackingKey)) {
                    // Track to Google Sheet with "GHO" prefix for cloud type to distinguish from regular cases
                    trackAction(comment.LastModifiedDate, caseRecord.CaseNumber, caseRecord.Severity_Level__c, 'GHO', caseRecord.CaseRoutingTaxonomy__r.Name.split('-')[0], currentMode, currentUserName, '');
                    localStorage.setItem(ghoTrackingKey, 'true');
                    console.log('GHO alert - GHO triage action tracked for case:', caseRecord.CaseNumber);
                  }
                }
              }
            }
          });
        }

        // Filter out cases that have been actioned
        const unactionedCases = result.records.filter(caseRecord => !actionedCaseIds.has(caseRecord.Id));

        if (unactionedCases.length > 0) {
          // Show GHO alert
          showGHOAlert(region, unactionedCases);
        }

        // Mark alert as shown regardless of whether we showed it
        localStorage.setItem(alertKey, 'true');
      });
    });
  });
}

function showGHOAlert(region, ghoRecords) {
  const alertTime = region === 'APAC' ? '7:30 AM' : region === 'EMEA' ? '2:30 PM' : '10:30 PM';

  // Create and show modal
  let existingModal = document.getElementById('gho-alert-modal');
  if (existingModal) {
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

  document.body.insertAdjacentHTML('beforeend', modalHtml);

  // Play GHO alert sound immediately when modal appears
  try {
    const audio = new Audio('../assets/audio/ghoalert.wav');
    audio.volume = 1.0;
    audio.play().catch(e => {
      console.log('Audio autoplay blocked or failed:', e);
    });
  } catch (e) {
    console.log('Could not play GHO alert sound:', e);
  }

  // Add event listeners
  document.getElementById('gho-alert-close').addEventListener('click', () => {
    document.getElementById('gho-alert-modal').remove();
  });

  document.getElementById('gho-alert-dismiss').addEventListener('click', () => {
    document.getElementById('gho-alert-modal').remove();
  });

  document.getElementById('gho-alert-check').addEventListener('click', () => {
    document.getElementById('gho-alert-modal').remove();
    checkGHOStatus(); // Open the GHO status modal
  });
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

// Function to ensure only one extension tab exists
async function ensureSingleTab() {
  try {
    // Send a message to background script to perform cleanup
    chrome.runtime.sendMessage('ensureSingleTab', function (response) {
      console.log('Tab cleanup requested:', response);
    });
  } catch (error) {
    console.error('Error requesting tab cleanup:', error);
  }
}

// Keyboard shortcuts handler
document.addEventListener('keydown', function (e) {
  // Check if Command (Mac) or Ctrl (Windows/Linux) is pressed
  if (e.metaKey || e.ctrlKey) {
    switch (e.key.toLowerCase()) {
      case 'f':
        e.preventDefault();
        // Focus on search input
        const searchInput = document.getElementById("search-input");
        if (searchInput) {
          searchInput.focus();
          searchInput.select(); // Select all text if any exists
        }
        break;

      case 'g':
        e.preventDefault();
        // Open GHO status modal (only in signature mode)
        if (currentMode === 'signature') {
          checkGHOStatus();
        } else {
          showToast('GHO status is only available in Signature mode');
        }
        break;

      case 'r':
        e.preventDefault();
        // Refresh the page
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
        // Switch to signature mode
        const modeSwitch = document.getElementById("mode-switch");
        if (modeSwitch && modeSwitch.checked) {
          modeSwitch.checked = false;
          modeSwitch.dispatchEvent(new Event('change'));
        }
        break;

      case '2':
        e.preventDefault();
        // Switch to premier mode
        const modeSwitchPremier = document.getElementById("mode-switch");
        if (modeSwitchPremier && !modeSwitchPremier.checked) {
          modeSwitchPremier.checked = true;
          modeSwitchPremier.dispatchEvent(new Event('change'));
        }
        break;

      case 'a':
        e.preventDefault();
        // Clear all filters and search
        document.getElementById("search-input").value = "";
        document.getElementById("search-button").disabled = true;
        document.getElementById("action-filter").value = "all";
        localStorage.setItem('caseFilter', 'all');
        clearTimeout(searchTimeout);
        applySearch("");
        showToast('Filters and search cleared');
        break;
    }
  }

  // ESC key to close modals
  if (e.key === 'Escape') {
    // Close GHO modal if open
    const ghoModal = document.getElementById('gho-modal');
    if (ghoModal && ghoModal.style.display === 'flex') {
      ghoModal.style.display = 'none';
    }

    // Close GHO alert modal if open
    const ghoAlertModal = document.getElementById('gho-alert-modal');
    if (ghoAlertModal) {
      ghoAlertModal.remove();
    }
  }
});

// Show keyboard shortcuts help on ? key
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
              <kbd style="background: #f3f4f6; padding: 4px 8px; border-radius: 4px; font-family: monospace; font-size: 12px; color: #374151; border: 1px solid #d1d5db;">⌘ + 1</kbd>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px 0; border-bottom: 1px solid #e5e7eb;">
              <span style="font-weight: 600; color: #111827; font-size: 15px;">Premier Mode</span>
              <kbd style="background: #f3f4f6; padding: 4px 8px; border-radius: 4px; font-family: monospace; font-size: 12px; color: #374151; border: 1px solid #d1d5db;">⌘ + 2</kbd>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px 0; border-bottom: 1px solid #e5e7eb;">
              <span style="font-weight: 600; color: #111827; font-size: 15px;">Clear All Filters</span>
              <kbd style="background: #f3f4f6; padding: 4px 8px; border-radius: 4px; font-family: monospace; font-size: 12px; color: #374151; border: 1px solid #d1d5db;">⌘ + A</kbd>
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

  // Remove existing help modal if any
  const existingModal = document.getElementById('shortcuts-help-modal');
  if (existingModal) {
    existingModal.remove();
  }

  document.body.insertAdjacentHTML('beforeend', helpModal);

  // Add event listeners
  document.getElementById('shortcuts-help-close').addEventListener('click', () => {
    document.getElementById('shortcuts-help-modal').remove();
  });

  // Close modal when clicking outside
  document.getElementById('shortcuts-help-modal').addEventListener('click', function (e) {
    if (e.target === this) {
      this.remove();
    }
  });
}

// Check for duplicate tabs when popup loads
document.addEventListener("DOMContentLoaded", function () {
  ensureSingleTab(); // Ensure only one tab exists
  updateWeekendModeIndicator();
  setInterval(updateWeekendModeIndicator, 60000);
  getSessionIds();

  // Start GHO alert checking every minute
  setInterval(checkGHOAlert, 60000);
  // Check immediately on load
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

    // Clear search and filter when changing modes
    document.getElementById("search-input").value = "";
    document.getElementById("search-button").disabled = true;
    document.getElementById("action-filter").value = "all";
    localStorage.setItem('caseFilter', 'all');
    clearTimeout(searchTimeout);

    // Clear existing content and show loading state
    document.getElementById("parentSigSev2").innerHTML = `
      <div class="loading-message">
        <h4 style="color: #374151; font-size: 20px; margin-bottom: 16px; font-weight: 600;">Switching to ${currentMode.charAt(0).toUpperCase() + currentMode.slice(1)} Mode...</h4>
        <p style="color: #6b7280; font-size: 16px; margin: 0;">Loading cases...</p>
      </div>
    `;

    // Reload data for the new mode
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

  // Help Button Event Listener
  document.getElementById("shortcuts-help-btn").addEventListener("click", function () {
    showKeyboardShortcutsHelp();
  });

  // GHO Button Event Listener
  document.getElementById("check-gho-button").addEventListener("click", function () {
    checkGHOStatus();
  });

  // GHO Modal Event Listeners
  document.getElementById("gho-modal-close").addEventListener("click", function () {
    document.getElementById('gho-modal').style.display = 'none';
  });

  // Close modal when clicking outside
  document.getElementById("gho-modal").addEventListener("click", function (e) {
    if (e.target === this) {
      this.style.display = 'none';
    }
  });

  // GHO Toggle Button Event Listener (using event delegation)
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

document.getElementById("clear-button").addEventListener("click", function () {
  // Clear search input
  document.getElementById("search-input").value = "";
  document.getElementById("search-button").disabled = true;

  // Reset filter to "all"
  document.getElementById("action-filter").value = "all";
  localStorage.setItem('caseFilter', 'all');

  // Clear any pending search timeout
  clearTimeout(searchTimeout);

  // Apply the reset state (show all cases)
  applySearch("");
});

document.getElementById("refresh-button").addEventListener("click", function () {
  const refreshIcon = this.querySelector('.fa-refresh');
  refreshIcon.classList.add('fa-spin');
  setTimeout(() => {
    window.location.reload();
  }, 500);
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

    // Show success message
    showToast(`Case snoozed for ${snoozeMinutes} minute${snoozeMinutes === 1 ? '' : 's'}`);
  }
});

// Handle snooze time select change to show/hide custom input
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

// Function to update status indicator based on case status
function updateStatusIndicator(hasUnactionedCases, totalCases, actionedCases) {
  const statusDot = document.querySelector('.status-dot');
  const statusText = document.querySelector('.status-text');

  if (!statusDot || !statusText) return;

  if (totalCases === 0) {
    statusDot.style.backgroundColor = '#22c55e'; // Green for no cases
    statusText.textContent = 'No Cases - All Clear';
  } else if (hasUnactionedCases) {
    statusDot.style.backgroundColor = '#ef4444'; // Red for unactioned cases
    statusText.textContent = `${totalCases - actionedCases} Cases Need Action`;
  } else {
    statusDot.style.backgroundColor = '#22c55e'; // Green for all actioned
    statusText.textContent = 'All Cases Actioned';
  }
}

// GHO Functions
function checkGHOStatus() {
  const modal = document.getElementById('gho-modal');
  const container = document.getElementById('gho-cases-container');

  // Show modal and loading state
  modal.style.display = 'flex';
  container.innerHTML = `
    <div class="loading-message">
      <h4>Loading GHO cases...</h4>
      <p>Please wait while we fetch the latest GHO cases.</p>
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

  const ghoQuery = `SELECT Id, CreatedDate, Account.Name, Owner.Name, SE_Target_Response__c, Severity_Level__c, CaseNumber, Subject, CaseRoutingTaxonomy__r.Name, SE_Initial_Response_Status__c, Contact.Is_MVP__c, support_available_timezone__c, (SELECT Transfer_Reason__c, CreatedDate, CreatedById, Preferred_Shift_Old_Value__c, Preferred_Shift_New_Value__c FROM Case_Routing_Logs__r ORDER BY CreatedDate DESC LIMIT 10) FROM Case WHERE ((Owner.Name IN ('Skills Queue','Kase Changer', 'Working in Org62','GHO Queue') AND ((Case_Support_level__c='Premier Priority' OR Case_Support_level__c='Signature' OR Case_Support_level__c='Signature Success') OR (Case_Support_level__c='Signature' OR Case_Support_level__c='Premier Priority' OR Case_Support_level__c='Signature Success'))) OR (Contact.Is_MVP__c=true AND Owner.Name='GHO Queue')) AND IsClosed=false AND ${shiftCondition} AND ((CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry%' OR CaseRoutingTaxonomy__r.Name LIKE 'Community-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Scale Center%' OR CaseRoutingTaxonomy__r.Name LIKE 'Customer Success Score%' OR CaseRoutingTaxonomy__r.Name LIKE 'Data Cloud-%') AND (Severity_Level__c='Level 1 - Critical' OR Severity_Level__c='Level 2 - Urgent')) AND CaseRoutingTaxonomy__r.Name NOT IN ('Disability and Product Accessibility','DORA')`;

  ghoConn.query(ghoQuery, function (err, result) {
    if (err) {
      console.error('GHO Query Error:', err);
      container.innerHTML = `
        <div class="no-cases-message" style="background-color: #fef2f2; border: 2px solid #ef4444;">
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
        console.log("GHO Comments:", commentResult);

        if (!commentErr && commentResult.records) {
          // Get current user ID for tracking
          ghoConn.query(`SELECT Id FROM User WHERE Name = '${currentUserName}' AND IsActive = True AND Username LIKE '%orgcs.com'`, function (userErr, userResult) {
            let currentUserId = null;
            if (!userErr && userResult.records.length > 0) {
              currentUserId = userResult.records[0].Id;
            }

            commentResult.records.forEach(comment => {
              if (comment.Body && comment.Body.includes('#GHOTriage') && comment.CreatedById === currentUserId) {
                const caseRecord = result.records.find(c => c.Id === comment.ParentId);
                if (caseRecord) {
                  const ghoTrackingKey = `gho_tracked_${caseRecord.Id}`;
                  if (!localStorage.getItem(ghoTrackingKey)) {
                    trackAction(comment.LastModifiedDate, caseRecord.CaseNumber, caseRecord.Severity_Level__c, 'GHO', caseRecord.CaseRoutingTaxonomy__r.Name.split('-')[0], currentMode, currentUserName, '');
                    localStorage.setItem(ghoTrackingKey, 'true');
                    console.log('GHO status check - GHO triage action tracked for case:', caseRecord.CaseNumber);
                  }
                }
              }
            });
          });
        }
      });
    }

    // Show modal with results
    showGHOStatusModal(result.records, ghoConn, comments);
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

function showGHOStatusModal(ghoRecords, conn, comments) {
  const container = document.getElementById('gho-cases-container');
  const currentShift = getCurrentShift();

  if (!ghoRecords || ghoRecords.length === 0) {
    container.innerHTML = `
      <div class="no-cases-message" style="background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%); border: 2px solid #3b82f6;">
        <h4 class="no-cases-title" style="color: #1d4ed8;">No GHO Cases Found</h4>
        <p class="no-cases-text">No cases matching GHO criteria were found for <strong>${currentShift}</strong> shift. Great work!</p>
        <p class="mode-switch-hint">✅ All GHO cases are up to date for current shift.</p>
      </div>
    `;
    return;
  }

  // Collect all unique user IDs from routing logs
  const allUserIds = new Set();
  ghoRecords.forEach(caseRecord => {
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
    let ghoHtml = `
      <div style="margin-bottom: 16px;">
        <h4 style="color: #374151; font-size: 18px; margin-bottom: 8px;">Found ${ghoRecords.length} GHO Case${ghoRecords.length === 1 ? '' : 's'}</h4>
        <p style="color: #6b7280; font-size: 14px;">Cases matching GHO criteria for <strong>${currentShift}</strong> shift</p>
        <div style="margin-top: 8px; padding: 8px 12px; background-color: #f3f4f6; border-radius: 6px; font-size: 12px; color: #374151;">
          <strong>Current Shift:</strong> ${currentShift} | 
          <strong>Time:</strong> ${new Date().toLocaleTimeString()} |
          <strong>Shifts:</strong> APAC (5:30AM-2:30PM), EMEA (12:30PM-9:30PM), AMER (8:30PM-5:30AM)
        </div>
      </div>
    `;

    ghoRecords.forEach(caseRecord => {
      const caseId = caseRecord.Id;
      const isMVP = caseRecord.Contact && caseRecord.Contact.Is_MVP__c === true;
      const actionTaken = comments.some(comment => comment.ParentId === caseId && comment.Body && comment.Body.includes('#GHOTriage'));
      console.log(actionTaken ? `GHO triage action found for case ${caseRecord.CaseNumber}` : `No GHO triage action found for case ${caseRecord.CaseNumber}`);


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
          routingLogHtml = `
            <div class="case-info-item">
              <span class="checkmark">✓</span>
              <span style="color: #9F2B68;">${lastLog.Transfer_Reason__c} (${timeElapsed(new Date(lastLog.CreatedDate))})</span>
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
                <div style="margin: 4px 0; padding: 8px 12px; background-color: #fef3c7; border-radius: 6px; border-left: 3px solid #f59e0b;">
                  <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
                    <span style="font-weight: 600; color: #92400e; font-size: 13px;">${userName}</span>
                    <span style="font-size: 11px; color: #78350f;">${formatDateWithDayOfWeek(transferTime)}</span>
                  </div>
                  <div style="display: flex; gap: 16px; font-size: 11px; color: #78350f;">
                    <span><strong>GHO FROM:</strong> ${ghoFrom}</span>
                    <span><strong>GHO TO:</strong> ${ghoTo}</span>
                  </div>
                  <div style="font-size: 10px; color: #a3a3a3; margin-top: 2px;">${timeElapsed(transferTime)} ago</div>
                </div>
              `;
            });
          }

          // Add expansion button if there are other transfers
          if (otherTransfers.length > 0) {
            nestedListHtml += `
              <div style="margin: 8px 0;">
                <button class="gho-toggle-btn" data-case-id="${caseUniqueId}"
                        style="background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%); border: 1px solid #cbd5e1; border-radius: 6px; padding: 8px 12px; font-size: 12px; color: #475569; cursor: pointer; width: 100%; display: flex; align-items: center; justify-content: center; gap: 8px; transition: all 0.2s ease; box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);">
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
                <div style="margin: 4px 0; padding: 8px 12px; background-color: #f9fafb; border-radius: 6px; border-left: 3px solid #9ca3af;">
                  <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
                    <span style="font-weight: 600; color: #4b5563; font-size: 13px;">${userName}</span>
                    <span style="font-size: 11px; color: #6b7280;">${formatDateWithDayOfWeek(transferTime)}</span>
                  </div>
                  <div style="display: flex; gap: 16px; font-size: 11px; color: #6b7280;">
                    <span><strong>GHO FROM:</strong> ${ghoFrom}</span>
                    <span><strong>GHO TO:</strong> ${ghoTo}</span>
                  </div>
                  <div style="font-size: 10px; color: #a3a3a3; margin-top: 2px;">${timeElapsed(transferTime)} ago</div>
                </div>
              `;
            });

            nestedListHtml += `</div>`;
          }

          ghoTransferHtml = `
            <div class="case-info-item" style="flex-direction: column; align-items: flex-start;">
              <div style="display: flex; align-items: center; margin-bottom: 8px;">
                <span class="checkmark">✓</span>
                <span style="color: #f59e0b; font-weight: bold;">GHO Transfer History (${allGhoTransfers.length} transfer${allGhoTransfers.length > 1 ? 's' : ''}):</span>
              </div>
              <div style="width: 100%; margin-left: 20px;">
                ${nestedListHtml}
              </div>
            </div>
          `;
        } else {
          ghoTransferHtml = `
            <div class="case-info-item">
              <span class="checkmark">✓</span>
              <span style="color: #6b7280; font-style: italic;">No GHO transfers found</span>
            </div>
          `;
        }
      }

      const caseHtml = `
        <div class="case-card ${isMVP ? 'mvp-case' : ''}" style="margin-bottom: 16px; ${isMVP ? 'border-left: 4px solid #9333ea;' : ''}">
          <div class="case-header">
            <div style="display: flex; align-items: center; gap: 8px;">
              ${isMVP ? '<span style="background-color: #9333ea; color: white; padding: 2px 8px; border-radius: 12px; font-size: 12px; font-weight: bold;">MVP</span>' : ''}
              <span style="background-color: #f59e0b; color: white; padding: 2px 8px; border-radius: 12px; font-size: 12px; font-weight: bold;">GHO</span>
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
              <div class="case-info-item">
                <span class="checkmark">✓</span>
                <span style="color: ${statusColor}; font-weight: bold;">${caseRecord.SE_Initial_Response_Status__c}${isMVP ? ' - MVP CASE' : ''}</span>
              </div>
              ${actionTaken ? `
              <div class="case-info-item">
                <span class="checkmark">✓</span>
                <span style="color: green; font-weight: bold;">Action Taken (#GHOTriage)</span>
              </div>
              ` : ''}
              ${routingLogHtml}
              ${ghoTransferHtml}
            </div>
            
            <div class="case-actions">
              <a target="_blank" href="https://orgcs.my.salesforce.com/lightning/r/Case/${caseId}/view" 
                 class="preview-btn gho-preview-btn" 
                 data-case-id="${caseId}"
                 style="margin-top: 12px;position:absolute;top:50px;right:10px">
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

    if (!expandedDiv || !toggleText) {
      console.error('GHO toggle elements not found for:', caseUniqueId);
      return;
    }

    if (expandedDiv.style.display === 'none') {
      // Expand the section
      expandedDiv.style.display = 'block';
      toggleText.innerHTML = toggleText.innerHTML.replace('Show', 'Hide');
      if (chevronIcon) {
        chevronIcon.style.transform = 'rotate(180deg)';
      }
    } else {
      // Collapse the section
      expandedDiv.style.display = 'none';
      toggleText.innerHTML = toggleText.innerHTML.replace('Hide', 'Show');
      if (chevronIcon) {
        chevronIcon.style.transform = 'rotate(0deg)';
      }
    }
  } catch (error) {
    console.error('Error toggling GHO transfers:', error);
  }
}
