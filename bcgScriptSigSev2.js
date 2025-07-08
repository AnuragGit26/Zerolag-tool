let SESSION_ID;
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
      let displayedCaseCount = 0;
      let actionTakenCount = 0;
      return conn.query("SELECT Id, CreatedDate, Account.Name, Owner.Name, SE_Target_Response__c, Severity_Level__c, CaseNumber, Subject, CaseRoutingTaxonomy__r.Name,SE_Initial_Response_Status__c FROM Case WHERE CreatedDate = LAST_N_DAYS:7 AND Status='New' AND (Owner.Name='Working in Org62' OR Owner.Name='Sales Cloud Skills Queue' OR Owner.Name='Service Cloud Skills Queue' OR Owner.Name='Industry Skills Queue') AND IsClosed=false AND Account_Support_SBR_Category__c!='JP' AND Account_Support_SBR_Category__c!='JP MCS' AND Account_Support_SBR_Category__c!='GOVT' AND Account_Support_SBR_Category__c!='MCS - GOVT' AND (CaseRoutingTaxonomy__r.Name='Sales-Issues Developing for Salesforce Functions (Product)' OR (Contact.Is_MVP__c=true AND (CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry-%') AND (Severity_Level__c='Level 2 - Urgent' OR Severity_Level__c='Level 1 - Critical') AND ( CaseRoutingTaxonomy__r.Name!='Sales-Quip' AND CaseRoutingTaxonomy__r.Name!='Sales-Sales Cloud for Slack' AND CaseRoutingTaxonomy__r.Name!='Sales-Disability and Product Accessibility' AND CaseRoutingTaxonomy__r.Name!='Service-Disability and Product Accessibility' AND CaseRoutingTaxonomy__r.Name!='Industry-Disability and Product Accessibility' AND CaseRoutingTaxonomy__r.Name!='Industry-Nonprofit Cloud' AND CaseRoutingTaxonomy__r.Name!='Industry-Education Cloud' AND CaseRoutingTaxonomy__r.Name!='Industry-Education Data Architecture (EDA)' AND CaseRoutingTaxonomy__r.Name!='Industry-Education Packages (Other SFDO)' AND CaseRoutingTaxonomy__r.Name!='Industry-Nonprofit Packages (Other SFDO)' AND CaseRoutingTaxonomy__r.Name!='Industry-Nonprofit Success Pack (NPSP)')) OR ( CaseRoutingTaxonomy__r.Name!='Sales-Quip' AND CaseRoutingTaxonomy__r.Name!='Sales-Sales Cloud for Slack' AND CaseRoutingTaxonomy__r.Name!='Sales-Disability and Product Accessibility' AND  CaseRoutingTaxonomy__r.Name!='Service-Disability and Product Accessibility' AND CaseRoutingTaxonomy__r.Name!='Industry-Disability and Product Accessibility' AND CaseRoutingTaxonomy__r.Name!='Industry-Nonprofit Cloud' AND CaseRoutingTaxonomy__r.Name!='Industry-Education Cloud' AND CaseRoutingTaxonomy__r.Name!='Industry-Education Data Architecture (EDA)' AND CaseRoutingTaxonomy__r.Name!='Industry-Education Packages (Other SFDO)' AND CaseRoutingTaxonomy__r.Name!='Industry-Nonprofit Packages (Other SFDO)' AND CaseRoutingTaxonomy__r.Name!='Industry-Nonprofit Success Pack (NPSP)' AND (Case_Support_level__c='Signature' OR Case_Support_level__c='Signature Success' OR Case_Support_level__c='Premier Priority') AND (Severity_Level__c='Level 2 - Urgent' OR Severity_Level__c='Level 1 - Critical') AND (CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry-%') ) ) ORDER BY CreatedDate DESC",
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
                  continue; // Skip displaying this case if it's snoozed
                } else if (snoozeUntil) {
                  // Snooze expired, remove from local storage
                  localStorage.removeItem('snooze_' + caseId);
                }
                const isActionTaken = localStorage.getItem(caseId) === 'true';
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
                const newHtml = `
                <div class="case-card card mb-3 shadow-sm" data-case-id="${caseId}">
                    <div class="card-header bg-light d-flex justify-content-between align-items-center">
                        <h5 class="mb-0 text-primary">
                            ${result.records[x].Subject}
                        </h5>
                        <small class="text-muted">${new Date(result.records[x].CreatedDate).toLocaleString()} (${timeElapsed(new Date(result.records[x].CreatedDate))})</small>
                    </div>
                    <div class="card-body">
                        <div class="row">
                            <div class="col-md-8">
                                <ul class="list-unstyled">
                                    <li><strong>Account:</strong> ${result.records[x].Account.Name}</li>
                                    <li><strong>Taxonomy:</strong> ${result.records[x].CaseRoutingTaxonomy__r.Name}</li>
                                    <li><strong>Case Number:</strong> ${result.records[x].CaseNumber}</li>
                                    <li><strong>Severity:</strong> ${result.records[x].Severity_Level__c}</li>
                                    <li><strong>Status:</strong> <span style="color:${statusColor}">${result.records[x].SE_Initial_Response_Status__c}</span></li>
                                </ul>
                            </div>
                            <div class="col-md-4 d-flex flex-column justify-content-between align-items-end">
                                <a target="_blank" href="https://orgcs.my.salesforce.com/lightning/r/Case/${caseId}/view" class="btn btn-primary btn-sm preview-record-btn mb-2">Preview Record</a>
                                <div class="snooze-controls d-flex align-items-center">
                                    <select class="form-control form-control-sm snooze-time" data-case-id="${caseId}">
                                        <option value="5">5 mins</option>
                                        <option value="10">10 mins</option>
                                        <option value="15">15 mins</option>
                                        <option value="20">20 mins</option>
                                    </select>
                                    <button class="btn btn-secondary btn-sm snooze-btn ml-2" data-case-id="${caseId}">Snooze</button>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="card-footer bg-light d-flex justify-content-end align-items-center">
                        <div class="form-check">
                            <input type="checkbox" class="form-check-input action-checkbox" id="action-${caseId}" data-case-id="${caseId}" ${isActionTaken ? 'checked' : ''}>
                            <label class="form-check-label action-taken-text" for="action-${caseId}" style="display: ${isActionTaken ? 'inline' : 'none'};">Action taken</label>
                        </div>
                    </div>
                </div>`;

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

      if (checkbox) {
        checkbox.checked = true;
        localStorage.setItem(caseId, 'true');
      }
      if (actionText) {
        actionText.style.display = "inline";
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
});

document.getElementById("clear-button").addEventListener("click", function () {
  document.getElementById("search-input").value = "";
  window.location.reload();
});

document.getElementById("parentSigSev2").addEventListener("click", function (e) {
  if (e.target && e.target.classList.contains("preview-record-btn")) {
    const severityText = e.target.closest('.d-style').querySelector('li:nth-child(4)').textContent;
    const severity = severityText.includes('Level 1') ? '1' : '2';
    const textToCopy = `Hi\nKindly help with the assignment of new SEV${severity} case, as it has not been assigned through OMNI. \nThank you!\nFYI: @Susanna Catherine \n#SigQBmention`;

    navigator.clipboard.writeText(textToCopy).then(function () {
      const toast = document.getElementById('toast');
      toast.style.display = 'block';
      setTimeout(function () {
        toast.style.display = 'none';
      }, 2000);
    });
  }
});
