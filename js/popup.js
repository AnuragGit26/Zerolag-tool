import { timeElapsed, addMinutes, isCurrentlyWeekend } from './utils/datetime.js';
import { applyFilter, updateWeekendModeIndicator } from './utils/dom.js';
import { trackAction } from './utils/api.js';

let SESSION_ID;
let currentMode = localStorage.getItem('caseTriageMode') || 'signature';
let currentUserName;
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
      let currentUserId;

      conn.query(`SELECT Id FROM User WHERE Name = '${currentUserName}' AND IsActive = True AND Username LIKE '%orgcs.com'`, function (err, userResult) {
        if (err) {
          return console.error('Error fetching user ID:', err);
        }
        if (userResult.records.length > 0) {
          currentUserId = userResult.records[0].Id;
        }

        let signatureQuery = "SELECT Id, CreatedDate, Account.Name, Owner.Name, SE_Target_Response__c, Severity_Level__c, CaseNumber, Subject, CaseRoutingTaxonomy__r.Name,SE_Initial_Response_Status__c, Contact.Is_MVP__c, (SELECT Transfer_Reason__c, CreatedDate FROM Case_Routing_Logs__r ORDER BY CreatedDate DESC LIMIT 1) FROM Case WHERE (Owner.Name LIKE '%Skills Queue%' OR Owner.Name='Kase Changer' OR Owner.Name='Working in Org62' OR Owner.Name='Data Cloud Queue') AND IsClosed=false AND Account_Support_SBR_Category__c!='JP MCS' AND Account.Name!='BT Test Account - HPA Premier Plus' AND Status='New' AND (((CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Community%' OR CaseRoutingTaxonomy__r.Name LIKE 'Scale Center%' OR CaseRoutingTaxonomy__r.Name LIKE 'Customer Success Score%') AND (Severity_Level__c='Level 1 - Critical' OR Severity_Level__c='Level 2 - Urgent') AND (Case_Support_level__c='Premier Priority' OR Case_Support_level__c='Signature' OR Case_Support_level__c='Signature Success')) OR ((CaseRoutingTaxonomy__r.Name LIKE 'Industry%') AND (Severity_Level__c='Level 1 - Critical' OR Severity_Level__c='Level 2 - Urgent') AND (Case_Support_level__c='Premier Priority' OR Case_Support_level__c='Signature' OR Case_Support_level__c='Signature Success')) OR (Contact.Is_MVP__c=true AND ((Severity_Level__c IN ('Level 1 - Critical', 'Level 2 - Urgent', 'Level 3 - High', 'Level 4 - Medium') AND (CaseRoutingTaxonomy__r.Name LIKE 'Sales%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry%')) OR (Severity_Level__c IN ('Level 3 - High', 'Level 4 - Medium') AND (CaseRoutingTaxonomy__r.Name LIKE 'Data Cloud-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Sales-Agentforce%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-Agentforce%'))))) AND (CaseRoutingTaxonomy__r.Name NOT IN ('Sales-Disability and Product Accessibility', 'Service-Disability and Product Accessibility', 'Industry-Disability and Product Accessibility', 'Sales-Issues Developing for Salesforce Functions (Product)', 'Sales-Quip', 'Sales-Sales Cloud for Slack', 'Industry-Nonprofit Cloud', 'Industry-Education Cloud', 'Industry-Education Data Architecture (EDA)', 'Industry-Education Packages (Other SFDO)', 'Industry-Nonprofit Packages (Other SFDO)', 'Industry-Nonprofit Success Pack (NPSP)', 'Service-Agentforce', 'Service-Agent for setup', 'Service-AgentforEmail', 'Service-Field Service Agentforce', 'Service-Agentforce for Dev', 'Sales-Agentforce', 'Sales-Agentforce for Dev', 'Sales-Agent for Setup', 'Sales-Prompt Builder', 'Data Cloud-Admin', 'Permissions', 'Flows', 'Reports & Dashboards', 'Data Cloud-Model Builder', 'Data Cloud-Connectors & Data Streams', 'Data Cloud-Developer', 'Calculated Insights & Consumption', 'Data Cloud-Segments', 'Activations & Identity Resolution')) ORDER BY CreatedDate DESC";

        let premierQuery = "SELECT Id, CreatedDate, Account.Name, Owner.Name, SE_Target_Response__c, Severity_Level__c, CaseNumber, Subject, CaseRoutingTaxonomy__r.Name, SE_Initial_Response_Status__c, Initial_Case_Severity__c, Contact.Is_MVP__c, (SELECT Transfer_Reason__c, CreatedDate FROM Case_Routing_Logs__r ORDER BY CreatedDate DESC LIMIT 1) FROM Case WHERE (Owner.Name IN ('Kase Changer', 'Working in Org62', 'Service Cloud Skills Queue', 'Sales Cloud Skills Queue', 'Industry Skills Queue', 'EXP Skills Queue', 'Data Cloud Queue')) AND (RecordType.Name IN ('Support', 'Partner Program Support', 'Platform / Application Support')) AND (Reason != 'Sales Request') AND (CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry-%') AND (Account_Support_SBR_Category__c != 'JP') AND (Case_Support_level__c IN ('Partner Premier', 'Premier', 'Premier+', 'Premium')) AND (IsClosed = false) AND (SE_Initial_Response_Status__c NOT IN ('Met', 'Completed After Violation', 'Missed', 'Violated')) AND (Account_Support_SBR_Category__c != 'JP') AND ((Severity_Level__c IN ('Level 1 - Critical', 'Level 2 - Urgent')) OR (Initial_Case_Severity__c IN ('Level 2 - Urgent', 'Level 1 - Critical'))) AND (CaseRoutingTaxonomy__r.Name NOT IN ('Service-Agentforce', 'Service-Agent for setup', 'Service-AgentforEmail', 'Service-Field Service Agentforce', 'Service-Agentforce for Dev', 'Sales-Agentforce', 'Sales-Agentforce for Dev', 'Sales-Agent for Setup', 'Sales-Prompt Builder', 'Data Cloud-Admin', 'Permissions', 'Flows', 'Reports & Dashboards', 'Data Cloud-Model Builder', 'Data Cloud-Connectors & Data Streams', 'Data Cloud-Developer', 'Calculated Insights & Consumption', 'Data Cloud-Segments', 'Activations & Identity Resolution')) AND CreatedDate = TODAY ORDER BY CreatedDate DESC";

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
            const commentQuery = `SELECT ParentId, Body, CreatedById, LastModifiedDate FROM CaseFeed WHERE Visibility = 'InternalUsers' AND ParentId IN ('${caseIds.join("','")}') AND Type = 'TextPost'`;

            conn.query(commentQuery, function (commentErr, commentResult) {
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
                      if (caseRecord) {
                        const trackingKey = `tracked_${caseRecord.Id}`;
                        if (!localStorage.getItem(trackingKey)) {
                          trackAction(caseRecord.LastModifiedDate, caseRecord.CaseNumber, caseRecord.Severity_Level__c, caseRecord.CaseRoutingTaxonomy__r.Name.split('-')[0], currentMode, currentUserName);
                          localStorage.setItem(trackingKey, 'true');
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
                    return false; // Exclude this record from the display list
                  }
                  return true; // Include all other records
                });

                console.log('Filtered records (after removing MVP Met cases):', filteredRecords.length, 'out of', result.records.length);

                // totalCasesCount = filteredRecords.length; // This was the bug, moved it up.

                for (var x in filteredRecords) {
                  const caseRecord = filteredRecords[x]; // Added for clarity
                  if (caseRecord.Contact && caseRecord.Contact.Is_MVP__c === true &&
                    (caseRecord.SE_Initial_Response_Status__c === 'In Warning' || caseRecord.SE_Initial_Response_Status__c === 'Warning')) {

                    const snoozeUntil = localStorage.getItem('snooze_' + caseId);
                    if (snoozeUntil && new Date().getTime() < parseInt(snoozeUntil)) {
                      continue;
                    } else if (snoozeUntil) {
                      localStorage.removeItem('snooze_' + caseId);
                    }

                    if (actionedCaseIds.has(caseId)) {
                      localStorage.setItem(caseId, 'true');
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
                    }

                    const mvpCaseHtml = `
                      <div class="case-card mvp-warning-case" style="border: 3px solid #ef4444; background-color: #fef2f2;">
                        <div class="case-header">
                          <div style="display: flex; align-items: center; gap: 8px;">
                            <span style="background-color: #ef4444; color: white; padding: 2px 8px; border-radius: 12px; font-size: 12px; font-weight: bold;">MVP URGENT</span>
                            <h3 class="case-title">${caseRecord.Subject}</h3>
                          </div>
                          <div class="case-timestamp">${new Date(caseRecord.CreatedDate).toLocaleString()} (${timeElapsed(new Date(caseRecord.CreatedDate))})</div>
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

                  console.log('Processing case:', filteredRecords[x]);

                  // Check if case meets alert criteria (timing + severity)
                  const meetsAlertCriteria = (filteredRecords[x].CaseRoutingTaxonomy__r.Name == 'Sales-Issues Developing for Salesforce Functions (Product)') ||
                    (today >= addMinutes(minSev1, new Date(filteredRecords[x].CreatedDate)) && filteredRecords[x].Severity_Level__c == 'Level 1 - Critical') ||
                    (today >= addMinutes(minSev2, new Date(filteredRecords[x].CreatedDate)));

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
                      isMVP: filteredRecords[x].Contact && filteredRecords[x].Contact.Is_MVP__c === true
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
                    const routingLogs = filteredRecords[x].Case_Routing_Logs__r;
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
                            <h3 class="case-title">${filteredRecords[x].Subject}</h3>
                          </div>
                          <div class="case-timestamp">${new Date(filteredRecords[x].CreatedDate).toLocaleString()} (${timeElapsed(new Date(filteredRecords[x].CreatedDate))})</div>
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
                               data-is-mvp="${caseData.isMVP}">
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
                    const requiredMinutes = filteredRecords[x].Severity_Level__c === 'Level 1 - Critical' ? minSev1 : minSev2;
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
                            Created: ${caseDetail.createdDate.toLocaleString()} (${caseDetail.minutesSinceCreation}m ago)
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
                            Created: ${caseDetail.createdDate.toLocaleString()} (${caseDetail.minutesSinceCreation}m ago)
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
      });
    }
  });
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

// Check for duplicate tabs when popup loads
document.addEventListener("DOMContentLoaded", function () {
  ensureSingleTab(); // Ensure only one tab exists
  updateWeekendModeIndicator();
  setInterval(updateWeekendModeIndicator, 60000);
  getSessionIds();

  const savedFilter = localStorage.getItem('caseFilter');
  if (savedFilter) {
    document.getElementById('action-filter').value = savedFilter;
  }

  document.getElementById("parentSigSev2").addEventListener("change", function () {
    // Removed manual checkbox interaction - checkboxes are now disabled
    // Auto-checking happens through case comment logic in getCaseDetails function
  });

  document.getElementById("parentSigSev2").addEventListener("click", function (e) {
    if (e.target.classList.contains("preview-record-btn")) {
      const button = e.target;
      const caseDiv = button.closest('.case-card');
      const checkbox = caseDiv.querySelector('.action-checkbox');
    }
  });

  document.getElementById("search-button").addEventListener("click", function () {
    let searchValue = document.getElementById("search-input").value.toLowerCase();
    let allCases = document.querySelectorAll("#parentSigSev2 .case-card");
    allCases.forEach(function (caseDiv) {
      let caseNumberElement = caseDiv.querySelector(".case-info-item:nth-child(3) span:nth-child(2)");
      if (caseNumberElement) {
        let caseNumber = caseNumberElement.textContent.toLowerCase();
        if (caseNumber.includes(searchValue)) {
          caseDiv.style.display = "block";
        } else {
          caseDiv.style.display = "none";
        }
      }
    });
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
      headerTitle.textContent = 'Premier Triage Case Alerts Center';
    } else {
      currentMode = 'signature';
      headerTitle.textContent = 'Signature Triage Case Alerts Center';
    }
    localStorage.setItem('caseTriageMode', currentMode);

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
    headerTitle.textContent = 'Premier Triage Case Alerts Center';
  } else {
    headerTitle.textContent = 'Signature Triage Case Alerts Center';
  }
});

document.getElementById("clear-button").addEventListener("click", function () {
  document.getElementById("search-input").value = "";
  window.location.reload();
});

document.getElementById("parentSigSev2").addEventListener("click", function (e) {
  if (e.target && e.target.classList.contains("preview-record-btn")) {
    const button = e.target;
    const caseDiv = button.closest('.case-card');
    const checkbox = caseDiv.querySelector('.action-checkbox');

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
    } else if (isMVPCase) {
      // Use MVP template for MVP cases
      textToCopy = `Hi\nKindly help with the assignment of new SEV${severity} MVP case, as it has not been assigned through OMNI and SLA is in warning status. Thank you!\nFYI: @Susanna Catherine \n#SigQBmention`;
    } else {
      textToCopy = `Hi\nKindly help with the assignment of new SEV${severity} case, as it has not been assigned through OMNI. \nThank you!\nFYI: @Susanna Catherine \n#SigQBmention`;
    }

    navigator.clipboard.writeText(textToCopy).then(function () {
      const toast = document.getElementById('toast');
      toast.style.display = 'block';
      setTimeout(function () {
        toast.style.display = 'none';
      }, 2000);
    });
  }
});

document.getElementById("parentSigSev2").addEventListener("click", function (e) {
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
