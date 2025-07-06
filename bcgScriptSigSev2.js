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
      return console.error('erororo---' + err);
    } else {
      return conn.query("SELECT Id, CreatedDate, Account.Name, Owner.Name, SE_Target_Response__c, Severity_Level__c, CaseNumber, Subject, CaseRoutingTaxonomy__r.Name FROM Case WHERE CreatedDate = LAST_N_DAYS:7 AND Status='New' AND (Owner.Name='Working in Org62' OR Owner.Name='Sales Cloud Skills Queue' OR Owner.Name='Service Cloud Skills Queue' OR Owner.Name='Industry Skills Queue') AND IsClosed=false AND Account_Support_SBR_Category__c!='JP' AND Account_Support_SBR_Category__c!='JP MCS' AND Account_Support_SBR_Category__c!='GOVT' AND Account_Support_SBR_Category__c!='MCS - GOVT' AND (CaseRoutingTaxonomy__r.Name='Sales-Issues Developing for Salesforce Functions (Product)' OR (Contact.Is_MVP__c=true AND (CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry-%') AND (Severity_Level__c='Level 2 - Urgent' OR Severity_Level__c='Level 1 - Critical') AND ( CaseRoutingTaxonomy__r.Name!='Sales-Quip' AND CaseRoutingTaxonomy__r.Name!='Sales-Sales Cloud for Slack' AND CaseRoutingTaxonomy__r.Name!='Sales-Disability and Product Accessibility' AND CaseRoutingTaxonomy__r.Name!='Service-Disability and Product Accessibility' AND CaseRoutingTaxonomy__r.Name!='Industry-Disability and Product Accessibility' AND CaseRoutingTaxonomy__r.Name!='Industry-Nonprofit Cloud' AND CaseRoutingTaxonomy__r.Name!='Industry-Education Cloud' AND CaseRoutingTaxonomy__r.Name!='Industry-Education Data Architecture (EDA)' AND CaseRoutingTaxonomy__r.Name!='Industry-Education Packages (Other SFDO)' AND CaseRoutingTaxonomy__r.Name!='Industry-Nonprofit Packages (Other SFDO)' AND CaseRoutingTaxonomy__r.Name!='Industry-Nonprofit Success Pack (NPSP)')) OR ( CaseRoutingTaxonomy__r.Name!='Sales-Quip' AND CaseRoutingTaxonomy__r.Name!='Sales-Sales Cloud for Slack' AND CaseRoutingTaxonomy__r.Name!='Sales-Disability and Product Accessibility' AND  CaseRoutingTaxonomy__r.Name!='Service-Disability and Product Accessibility' AND CaseRoutingTaxonomy__r.Name!='Industry-Disability and Product Accessibility' AND CaseRoutingTaxonomy__r.Name!='Industry-Nonprofit Cloud' AND CaseRoutingTaxonomy__r.Name!='Industry-Education Cloud' AND CaseRoutingTaxonomy__r.Name!='Industry-Education Data Architecture (EDA)' AND CaseRoutingTaxonomy__r.Name!='Industry-Education Packages (Other SFDO)' AND CaseRoutingTaxonomy__r.Name!='Industry-Nonprofit Packages (Other SFDO)' AND CaseRoutingTaxonomy__r.Name!='Industry-Nonprofit Success Pack (NPSP)' AND (Case_Support_level__c='Signature' OR Case_Support_level__c='Signature Success' OR Case_Support_level__c='Premier Priority') AND (Severity_Level__c='Level 2 - Urgent' OR Severity_Level__c='Level 1 - Critical') AND (CaseRoutingTaxonomy__r.Name LIKE 'Sales-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Service-%' OR CaseRoutingTaxonomy__r.Name LIKE 'Industry-%') ) ) ORDER BY CreatedDate DESC", function (err, result) {
        if (err) {
          alert('Your query has failed');
          return console.error(err);
        }
        var myHtml;
        var isData = false;
        var today = new Date();
        let displayedCaseCount = 0;
        let actionTakenCount = 0;

        // Function to check if current time is weekend (Saturday 5:30 AM IST to Monday 5:30 AM IST)
        function isWeekend() {
          return isCurrentlyWeekend();
        }
        var minSev1 = isWeekend() ? 1 : 5;
        var minSev2 = isWeekend() ? 1 : 20;
        if (result.records.length > 0) {
          isData = true;
          for (x in result.records) {
            if ((result.records[x].CaseRoutingTaxonomy__r.Name == 'Sales-Issues Developing for Salesforce Functions (Product)') || (today >= addMinutes(minSev1, new Date(result.records[x].CreatedDate)) && result.records[x].Severity_Level__c == 'Level 1 - Critical') || (today >= addMinutes(minSev2, new Date(result.records[x].CreatedDate)))) {
              const caseId = result.records[x].Id;
              const isActionTaken = localStorage.getItem(caseId) === 'true';
              displayedCaseCount++;
              if (isActionTaken) {
                actionTakenCount++;
              }
              const newHtml = '<div style="margin-top:20px;"></div> <div class="d-style btn btn-brc-tp border-2 w-100 my-2 py-3 shadow-sm" style="width: 100%; position: relative; background-color: #FBFBFF; border-color: #657ED4;"> <div style="position: absolute; top: 5px; right: 10px; color: #FB5012;font-weight:bold;">' + new Date(result.records[x].CreatedDate).toLocaleString() + ' (' + timeElapsed(new Date(result.records[x].CreatedDate)) + ')</div> <div class="row align-items-center" style="width: 100%"> <div class="col-12 col-md-4"> <h4 class="pt-3 text-170 text-600 letter-spacing" style="color: #3626A7;">' + result.records[x].Subject + '</h4> </div> <ul class="list-unstyled mb-0 col-12 col-md-4 text-90 text-left my-4 my-md-0" style="color: #0D0106;"> <li><i class="fa fa-check text-success-m2 text-110 mr-2 mt-1"></i><span><span class="text-110">' + result.records[x].Account.Name + '</span></span></li> <li class="mt-25"><i class="fa fa-check text-success-m2 text-110 mr-2 mt-1"></i><span class="text-110">' + result.records[x].CaseRoutingTaxonomy__r.Name + '</span></li> <li class="mt-25"><i class="fa fa-check text-success-m2 text-110 mr-2 mt-1"></i><span class="text-110">' + result.records[x].CaseNumber + '</span></li> <li class="mt-25"><i class="fa fa-check text-success-m2 text-110 mr-2 mt-1"></i><span class="text-110">' + result.records[x].Severity_Level__c + '</span></li> </ul> <div class="col-12 col-md-4 text-center"><a target="_blank" href="https://orgcs.my.salesforce.com/lightning/r/Case/' + caseId + '/view" class="f-n-hover btn btn-raised px-4 py-25 w-75 text-600 preview-record-btn" style="background-color: #3626A7; color: #FBFBFF;">Preview Record</a></div> </div> <div style="position: absolute; bottom: 5px; right: 10px;"><input type="checkbox" class="action-checkbox" data-case-id="' + caseId + '" ' + (isActionTaken ? 'checked' : '') + '> <span class="action-taken-text" style="display: ' + (isActionTaken ? 'inline' : 'none') + '; color: #214E34;font-weight: bold;">Action taken</span></div> </div> <div style="margin-top:10px;"></div>';
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
    let allCases = document.querySelectorAll("#parentSigSev2 > .d-style");
    allCases.forEach(function (caseDiv) {
      let checkbox = caseDiv.querySelector(".action-checkbox");
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
    });
  });
});

document.getElementById("clear-button").addEventListener("click", function () {
  document.getElementById("search-input").value = "";
  window.location.reload();
});
