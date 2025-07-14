let SESSION_ID;
let currentMode = localStorage.getItem('caseTriageMode') || 'signature';
let currentUserName;
const SPREADSHEET_ID = '1BKxQLGFrczjhcx9rEt-jXGvlcCPQblwBhFJjoiDD7TI';

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

function getCaseDetails(callback) {
  console.log('SESSION_ID>>>', SESSION_ID);
  let conn = new jsforce.Connection({
    serverUrl: 'https://orgcs.my.salesforce.com',
    sessionId: SESSION_ID,
  });
  conn.identity(function (err, res) {
    if (err) {
      var data = 'closeTab';
      chrome.runtime.sendMessage(data, function (response) {
        console.log('response-----' + response);
      });
      return console.error('error---' + err);
    } else {
      currentUserName = res.display_name;
      let displayedCaseCount = 0;
      let actionTakenCount = 0;

      let signatureQuery = "SELECT Id, CreatedDate, Account.Name, Owner.Name, SE_Target_Response__c, Severity_Level__c, CaseNumber, Subject, CaseRoutingTaxonomy__r.Name,SE_Initial_Response_Status__c FROM Case WHERE CreatedDate = LAST_N_DAYS:7 AND Status='New' AND (Owner.Name='Working in Org62' OR Owner.Name='Sales Cloud Skills Queue' OR Owner.Name='Service Cloud Skills Queue' OR Owner.Name='Industry Skills Queue') AND IsClosed=false AND Account_Support_SBR_Category__c!='JP' AND Account_Support_SBR_Category__c!='JP MCS' AND Account_Support_SBR_Category__c!='GOVT' AND Account_Support_SBR_Category__c!='MCS - GOVT' AND (CaseRoutingTaxonomy__r.Name='Sales-Issues Developing for Salesforce Functions (Product)' OR (Contact.Is_MVP__c=true AND (CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry-%') AND (Severity_Level__c='Level 2 - Urgent' OR Severity_Level__c='Level 1 - Critical') AND ( CaseRoutingTaxonomy__r.Name!='Sales-Quip' AND CaseRoutingTaxonomy__r.Name!='Sales-Sales Cloud for Slack' AND CaseRoutingTaxonomy__r.Name!='Sales-Disability and Product Accessibility' AND CaseRoutingTaxonomy__r.Name!='Service-Disability and Product Accessibility' AND CaseRoutingTaxonomy__r.Name!='Industry-Disability and Product Accessibility' AND CaseRoutingTaxonomy__r.Name!='Industry-Nonprofit Cloud' AND CaseRoutingTaxonomy__r.Name!='Industry-Education Cloud' AND CaseRoutingTaxonomy__r.Name!='Industry-Education Data Architecture (EDA)' AND CaseRoutingTaxonomy__r.Name!='Industry-Education Packages (Other SFDO)' AND CaseRoutingTaxonomy__r.Name!='Industry-Nonprofit Packages (Other SFDO)' AND CaseRoutingTaxonomy__r.Name!='Industry-Nonprofit Success Pack (NPSP)')) OR ( CaseRoutingTaxonomy__r.Name!='Sales-Quip' AND CaseRoutingTaxonomy__r.Name!='Sales-Sales Cloud for Slack' AND CaseRoutingTaxonomy__r.Name!='Sales-Disability and Product Accessibility' AND  CaseRoutingTaxonomy__r.Name!='Service-Disability and Product Accessibility' AND CaseRoutingTaxonomy__r.Name!='Industry-Disability and Product Accessibility' AND CaseRoutingTaxonomy__r.Name!='Industry-Nonprofit Cloud' AND CaseRoutingTaxonomy__r.Name!='Industry-Education Cloud' AND CaseRoutingTaxonomy__r.Name!='Industry-Education Data Architecture (EDA)' AND CaseRoutingTaxonomy__r.Name!='Industry-Education Packages (Other SFDO)' AND CaseRoutingTaxonomy__r.Name!='Industry-Nonprofit Packages (Other SFDO)' AND CaseRoutingTaxonomy__r.Name!='Industry-Nonprofit Success Pack (NPSP)' AND (Case_Support_level__c='Signature' OR Case_Support_level__c='Signature Success' OR Case_Support_level__c='Premier Priority') AND (Severity_Level__c='Level 2 - Urgent' OR Severity_Level__c='Level 1 - Critical') AND (CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry-%') ) ) ORDER BY CreatedDate DESC";

      let premierQuery = "SELECT Id, CreatedDate, Account.Name, Owner.Name, SE_Target_Response__c, Severity_Level__c, CaseNumber, Subject, CaseRoutingTaxonomy__r.Name, SE_Initial_Response_Status__c, Initial_Case_Severity__c FROM Case WHERE (Owner.Name IN ('Kase Changer', 'Working in Org62', 'Service Cloud Skills Queue', 'Sales Cloud Skills Queue', 'Industry Skills Queue', 'EXP Skills Queue', 'Data Cloud Queue')) AND (RecordType.Name IN ('Support', 'Partner Program Support', 'Platform / Application Support')) AND (Reason != 'Sales Request') AND (CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry-%') AND (Account_Support_SBR_Category__c != 'JP') AND (Case_Support_level__c IN ('Partner Premier', 'Premier', 'Premier+', 'Premium')) AND (IsClosed = false) AND (SE_Initial_Response_Status__c NOT IN ('Met', 'Completed After Violation', 'Missed', 'Violated')) AND (Account_Support_SBR_Category__c != 'JP') AND ((Severity_Level__c IN ('Level 1 - Critical', 'Level 2 - Urgent')) OR (Initial_Case_Severity__c IN ('Level 2 - Urgent', 'Level 1 - Critical'))) AND (CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry-%' OR CaseRoutingTaxonomy__r.Name IN ('Data Cloud-Admin', 'Permissions', 'Flows', 'Reports & Dashboards')) AND CreatedDate = TODAY ORDER BY CreatedDate DESC";

      let query = currentMode === 'premier' ? premierQuery : signatureQuery;

      return conn.query(query,
        function (err, result) {
          if (err) {
            alert('Your query has failed');
            return console.error(err);
          }
          var myHtml;
          var isData = false;
          var today = new Date();


          // Function to check if current time is weekend (Saturday 5:30 AM IST to Monday 5:30 AM IST)
          function isWeekend() {
            return isCurrentlyWeekend();
          }
          var minSev1 = isWeekend() ? 1 : 5;
          var minSev2 = isWeekend() ? 1 : 20;
          if (result.records.length > 0) {
            isData = true;
            for (x in result.records) {
              conn.query("SELECT Name,Id, LastModifiedById, Transfer_Reason__c FROM Case_Routing_Log__c WHERE Case__c = '" + result.records[x].Id + "'", function (err, logResult) {
                if (err) { return console.error(err); }
                console.log("Case_Routing_Log__c result for case " + result.records[x].Id + ": ", logResult);
              });
              console.log('result.records[x]>>>', result.records[x]);
              if ((result.records[x].CaseRoutingTaxonomy__r.Name == 'Sales-Issues Developing for Salesforce Functions (Product)') || (today >= addMinutes(minSev1, new Date(result.records[x].CreatedDate)) && result.records[x].Severity_Level__c == 'Level 1 - Critical') || (today >= addMinutes(minSev2, new Date(result.records[x].CreatedDate)))) {
                const caseId = result.records[x].Id;
                const snoozeUntil = localStorage.getItem('snooze_' + caseId);
                if (snoozeUntil && new Date().getTime() < parseInt(snoozeUntil)) {
                  continue;
                } else if (snoozeUntil) {
                  localStorage.removeItem('snooze_' + caseId);
                }
                const isActionTaken = localStorage.getItem(caseId) === 'true';
                const caseData = {
                  number: result.records[x].CaseNumber,
                  severity: result.records[x].Severity_Level__c,
                  cloud: result.records[x].CaseRoutingTaxonomy__r.Name.split('-')[0]
                };
                displayedCaseCount++;
                if (isActionTaken) {
                  actionTakenCount++;
                }
                let statusColor = '';
                if (result.records[x].SE_Initial_Response_Status__c === 'Met') {
                  statusColor = 'green';
                } else if (result.records[x].SE_Initial_Response_Status__c === 'In Warning' || result.records[x].SE_Initial_Response_Status__c === 'Warning') {
                  statusColor = 'red';
                }
                const newHtml = '<div style="margin-top:20px;"></div> <div class="d-style btn btn-brc-tp border-2 w-100 my-2 py-3 shadow-sm" style="width: 100%; position: relative; background-color: #FBFBFF; border-color: #657ED4;"> <div style="position: absolute; top: 5px; right: 10px; color: #FB5012;font-weight:bold;">' + new Date(result.records[x].CreatedDate).toLocaleString() + ' (' + timeElapsed(new Date(result.records[x].CreatedDate)) + ')</div> <div class="row align-items-center" style="width: 100%"> <div class="col-12 col-md-4"> <h4 class="pt-3 text-170 text-600 letter-spacing" style="color: #3626A7;">' + result.records[x].Subject + '</h4> </div> <ul class="list-unstyled mb-0 col-12 col-md-4 text-90 text-left my-4 my-md-0" style="color: #0D0106;"> <li><i class="fa fa-check text-success-m2 text-110 mr-2 mt-1"></i><span><span class="text-110">' + result.records[x].Account.Name + '</span></span></li> <li class="mt-25"><i class="fa fa-check text-success-m2 text-110 mr-2 mt-1"></i><span class="text-110">' + result.records[x].CaseRoutingTaxonomy__r.Name + '</span></li> <li class="mt-25"><i class="fa fa-check text-success-m2 text-110 mr-2 mt-1"></i><span class="text-110">' + result.records[x].CaseNumber + '</span></li> <li class="mt-25"><i class="fa fa-check text-success-m2 text-110 mr-2 mt-1"></i><span class="text-110">' + result.records[x].Severity_Level__c + '</span></li> <li class="mt-25"><i class="fa fa-check text-success-m2 text-110 mr-2 mt-1"></i><span class="text-110" style="color:' + statusColor + '">' + result.records[x].SE_Initial_Response_Status__c + '</span></li> </ul> <div class="col-12 col-md-4 text-center"><a target="_blank" href="https://orgcs.my.salesforce.com/lightning/r/Case/' + caseId + '/view" class="f-n-hover btn btn-raised px-4 py-25 w-75 text-600 preview-record-btn" style="background-color: #3626A7; color: #FBFBFF;" data-case-number="' + caseData.number + '" data-severity="' + caseData.severity + '" data-cloud="' + caseData.cloud + '">Preview Record</a></div> </div> <div style="position: absolute; top:30px; right: 10px;"><input type="checkbox" class="action-checkbox" data-case-id="' + caseId + '" ' + (isActionTaken ? 'checked' : '') + ' data-case-number="' + caseData.number + '" data-severity="' + caseData.severity + '" data-cloud="' + caseData.cloud + '"> <span class="action-taken-text" style="display: ' + (isActionTaken ? 'inline' : 'none') + '; color: #214E34;font-weight: bold;">Action taken</span></div> <div class="snooze-controls" style="position: absolute; bottom: 30px; right: 10px; display: flex; align-items: center;"><select class="snooze-time" data-case-id="' + caseId + '" style="margin-right: 5px; border-radius: 4px; border: 1px solid #657ED4; background-color: #FBFBFF; color: #3626A7;"><option value="5">5 mins</option><option value="10">10 mins</option><option value="15">15 mins</option><option value="20">20 mins</option></select><button class="snooze-btn" data-case-id="' + caseId + '" style="background-color: #3626A7; color: #FBFBFF; border: none; border-radius: 4px; padding: 5px 10px; cursor: pointer;">Snooze</button></div> </div> <div style="margin-top:10px;"></div>';
                if (myHtml) {
                  myHtml = myHtml + newHtml;
                } else {
                  myHtml = newHtml;
                }
              }
            }

          }
          if (isData && myHtml != undefined) {
            document.getElementById("parentSigSev2").innerHTML += myHtml;

            const savedFilter = localStorage.getItem('caseFilter');
            if (savedFilter) {
              applyFilter(savedFilter);
            }

            // Only focus tab and play audio if there are cases without action taken
            if (actionTakenCount < displayedCaseCount) {
              var audio = new Audio('noti.wav');
              audio.play();
              var data = 'openTab';
              chrome.runtime.sendMessage(data, function (response) {
                console.log('response-----' + response);
              });
            } else {
              // All cases have action taken, don't focus the tab
              var data = 'closeTab';
              chrome.runtime.sendMessage(data, function (response) {
                console.log('response-----' + response);
              });
            }
            //}, 1000);
          } else {
            // Show "no cases to action" message when there are no records
            const noCasesHtml = '<div style="margin-top:20px;"></div> <div class="d-style btn btn-brc-tp border-2 w-100 my-2 py-3 shadow-sm" style="width: 100%; position: relative; background-color: #F0F8FF; border-color: #28A745;"> <div class="row align-items-center justify-content-center" style="width: 100%"> <div class="col-12 text-center"> <h4 class="pt-3 text-170 text-600 letter-spacing" style="color: #28A745;">No Cases to Action</h4> <p class="text-110" style="color: #0D0106; margin-top: 10px;">All cases are up to date. Great work!</p> </div> </div> </div> <div style="margin-top:10px;"></div>';
            document.getElementById("parentSigSev2").innerHTML += noCasesHtml;
            // All cases have action taken, don't focus the tab
            var data = 'closeTab';
            chrome.runtime.sendMessage(data, function (response) {
              console.log('response-----' + response);
            });
          }
        });
    }
  });
}

function timeElapsed(createdDate) {
  const now = new Date();
  const diff = now - createdDate;
  const seconds = Math.floor(diff / 1000);
  const minutes = Math.floor(seconds / 60);
  const hours = Math.floor(minutes / 60);
  const days = Math.floor(hours / 24);

  if (days > 0) {
    return `${days}d ago`;
  }
  if (hours > 0) {
    return `${hours}h ${minutes % 60}m ago`;
  }
  if (minutes > 0) {
    return `${minutes}m ago`;
  }
  return `${seconds}s ago`;
}

function addMinutes(numOfMinutes, date = new Date()) {
  date.setMinutes(date.getMinutes() + numOfMinutes);
  return date;
}

function getAuthToken(callback) {
  chrome.identity.getAuthToken({ interactive: true }, function (token) {
    if (chrome.runtime.lastError) {
      console.error(chrome.runtime.lastError);
      return;
    }
    callback(token);
  });
}

function trackAction(caseNumber, severity, cloud) {
  getAuthToken(function (token) {
    const sheetName = currentMode === 'premier' ? 'premier' : 'signature';
    const now = new Date();
    const pstDate = now.toLocaleString('en-US', {
      timeZone: 'America/Los_Angeles',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: true
    });

    const values = [
      [pstDate, caseNumber, currentUserName, severity, cloud]
    ];

    const body = {
      values: values
    };

    fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${sheetName}:append?valueInputOption=USER_ENTERED`, {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(body)
    }).then(response => response.json())
      .then(data => {
        console.log('Sheet updated:', data);
      })
      .catch(error => {
        console.error('Error updating sheet:', error);
      });
  });
}


function isCurrentlyWeekend() {
  const now = new Date();
  const currentOffset = now.getTimezoneOffset();
  const istOffset = -330;

  let istTime;
  if (currentOffset === istOffset) {
    istTime = now;
  } else {
    const istOffsetMs = 5.5 * 60 * 60 * 1000;
    istTime = new Date(now.getTime() + istOffsetMs + (currentOffset * 60 * 1000));
  }

  const day = istTime.getDay();
  const hour = istTime.getHours();
  const minute = istTime.getMinutes();
  const timeInMinutes = hour * 60 + minute;
  const cutoffTime = 5 * 60 + 30;

  // Weekend: Saturday 5:30 AM IST to Monday 5:30 AM IST
  if (day === 6 && timeInMinutes >= cutoffTime) return true;
  if (day === 0) return true;
  if (day === 1 && timeInMinutes < cutoffTime) return true;

  return false;
}

function applyFilter(filterValue) {
  let allCases = document.querySelectorAll("#parentSigSev2 > .d-style");
  allCases.forEach(function (caseDiv) {
    let checkbox = caseDiv.querySelector(".action-checkbox");
    if (checkbox) {
      let isActionTaken = checkbox.checked;
      if (filterValue === "all") {
        caseDiv.style.display = "block";
      } else if (filterValue === "action-taken" && isActionTaken) {
        caseDiv.style.display = "block";
      } else if (filterValue === "not-action-taken" && !isActionTaken) {
        caseDiv.style.display = "block";
      } else {
        caseDiv.style.display = "none";
      }
    }
  });
}

function updateWeekendModeIndicator() {
  const weekendIndicator = document.getElementById("weekend-mode-indicator");
  if (isCurrentlyWeekend()) {
    weekendIndicator.style.display = "block";
  } else {
    weekendIndicator.style.display = "none";
  }
}

document.addEventListener("DOMContentLoaded", function () {
  updateWeekendModeIndicator();
  setInterval(updateWeekendModeIndicator, 60000);
  getSessionIds();

  const savedFilter = localStorage.getItem('caseFilter');
  if (savedFilter) {
    document.getElementById('action-filter').value = savedFilter;
  }

  document.getElementById("parentSigSev2").addEventListener("change", function (e) {
    if (e.target.classList.contains("action-checkbox")) {
      const checkbox = e.target;
      const actionText = checkbox.nextElementSibling;
      const caseId = checkbox.dataset.caseId;
      if (checkbox.checked) {
        actionText.style.display = "inline";
        localStorage.setItem(caseId, 'true');
        trackAction(checkbox.dataset.caseNumber, checkbox.dataset.severity, checkbox.dataset.cloud);
      } else {
        actionText.style.display = "none";
        localStorage.removeItem(caseId);
      }
    }
  });

  document.getElementById("parentSigSev2").addEventListener("click", function (e) {
    if (e.target.classList.contains("snooze-btn")) {
      const button = e.target;
      const caseId = button.dataset.caseId;
      const caseDiv = button.closest('.d-style');
      const snoozeTimeSelect = caseDiv.querySelector('.snooze-time');
      const snoozeMinutes = parseInt(snoozeTimeSelect.value);
      const snoozeUntil = new Date().getTime() + snoozeMinutes * 60 * 1000;
      localStorage.setItem('snooze_' + caseId, snoozeUntil);
      caseDiv.style.display = 'none';
    }
  });

  document.getElementById("parentSigSev2").addEventListener("click", function (e) {
    if (e.target.classList.contains("preview-record-btn")) {
      const button = e.target;
      const caseDiv = button.closest('.d-style');
      const checkbox = caseDiv.querySelector('.action-checkbox');
      const actionText = caseDiv.querySelector('.action-taken-text');
      const caseId = checkbox.dataset.caseId;

      if (checkbox && !checkbox.checked) {
        checkbox.checked = true;
        localStorage.setItem(caseId, 'true');
        trackAction(button.dataset.caseNumber, button.dataset.severity, button.dataset.cloud);
        if (actionText) {
          actionText.style.display = "inline";
        }
      } else if (checkbox.checked) {
        // Already actioned, but we can ensure the text is visible
        if (actionText) {
          actionText.style.display = "inline";
        }
      }
    }
  });

  document.getElementById("search-button").addEventListener("click", function () {
    let searchValue = document.getElementById("search-input").value.toLowerCase();
    let allCases = document.querySelectorAll("#parentSigSev2 > div");
    allCases.forEach(function (caseDiv) {
      let caseNumberElement = caseDiv.querySelector("li:nth-child(3) span");
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
    const pageTitle = document.getElementById("page-title");
    if (this.checked) {
      currentMode = 'premier';
      pageTitle.textContent = 'Premier Triage Case Alerts Center';
    } else {
      currentMode = 'signature';
      pageTitle.textContent = 'Signature Triage Case Alerts Center';
    }
    localStorage.setItem('caseTriageMode', currentMode);
    document.getElementById("parentSigSev2").innerHTML = '';
    getSessionIds();
  });

  const modeSwitch = document.getElementById("mode-switch");
  const pageTitle = document.getElementById("page-title");
  if (currentMode === 'premier') {
    modeSwitch.checked = true;
    pageTitle.textContent = 'Premier Triage Case Alerts Center';
  } else {
    pageTitle.textContent = 'Signature Triage Case Alerts Center';
  }
});

document.getElementById("clear-button").addEventListener("click", function () {
  document.getElementById("search-input").value = "";
  window.location.reload();
});

document.getElementById("parentSigSev2").addEventListener("click", function (e) {
  if (e.target && e.target.classList.contains("preview-record-btn")) {
    const button = e.target;
    const caseDiv = button.closest('.d-style');
    const checkbox = caseDiv.querySelector('.action-checkbox');
    const actionText = caseDiv.querySelector('.action-taken-text');
    const caseId = checkbox.dataset.caseId;

    if (checkbox && !checkbox.checked) {
      checkbox.checked = true;
      localStorage.setItem(caseId, 'true');
      trackAction(button.dataset.caseNumber, button.dataset.severity, button.dataset.cloud);
      if (actionText) {
        actionText.style.display = "inline";
      }
    }

    const severityText = e.target.closest('.d-style').querySelector('li:nth-child(4)').textContent;
    const severity = severityText.includes('Level 1') ? '1' : '2';
    let textToCopy = '';

    if (currentMode === 'premier') {
      textToCopy = `Hi\nNew SEV${severity} assigned to you & App is updated...!`;
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
