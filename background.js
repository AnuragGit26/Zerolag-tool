chrome.runtime.onInstalled.addListener(details => {
	if (details.reason === chrome.runtime.OnInstalledReason.INSTALL) {
		chrome.alarms.create('SigSEV2Triaging', { periodInMinutes: 3 });
	}
});

let tabId;
chrome.alarms.onAlarm.addListener(alarm => {
	chrome.tabs.create({ 'active': false, 'url': chrome.runtime.getURL('bcgPageSigSev2.html') }, function (tab) {
		tabId = tab.id;
	});
});

chrome.runtime.onMessage.addListener(
	function (request, sender, sendResponse) {
		if (request == 'closeTab') {
			chrome.tabs.remove(tabId, function () { });
			sendResponse("tab closed");
		}
		if (request == 'openTab') {
			chrome.tabs.update(tabId, { "active": true });
			sendResponse("tab open");
		}
	}
);