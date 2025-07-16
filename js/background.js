chrome.runtime.onInstalled.addListener(details => {
	if (details.reason === chrome.runtime.OnInstalledReason.INSTALL) {
		chrome.alarms.create('SigSEV2Triaging', { periodInMinutes: 1 });
	}
	// Clean up any existing tabs on installation/update
	closeExistingExtensionTabs();
});

// Also clean up on startup
chrome.runtime.onStartup.addListener(() => {
	closeExistingExtensionTabs();
});

let tabId;

// Function to close all existing extension tabs
async function closeExistingExtensionTabs() {
	const extensionUrl = chrome.runtime.getURL('popup/popup.html');
	try {
		const tabs = await chrome.tabs.query({ url: extensionUrl });
		if (tabs.length > 0) {
			const tabIds = tabs.map(tab => tab.id);
			await chrome.tabs.remove(tabIds);
			console.log(`Closed ${tabs.length} existing extension tab(s)`);
		}
	} catch (error) {
		console.error('Error closing existing tabs:', error);
	}
}

// Function to create a single extension tab
async function createSingleExtensionTab() {
	try {
		// First, close any existing extension tabs
		await closeExistingExtensionTabs();

		// Then create a new tab
		const tab = await chrome.tabs.create({
			'active': false,
			'url': chrome.runtime.getURL('popup/popup.html')
		});
		tabId = tab.id;
		console.log(`Created new extension tab with ID: ${tabId}`);
	} catch (error) {
		console.error('Error creating extension tab:', error);
	}
}

chrome.alarms.onAlarm.addListener(async (alarm) => {
	if (alarm.name === 'SigSEV2Triaging') {
		await createSingleExtensionTab();
	}
});

chrome.runtime.onMessage.addListener(
	function (request, sender, sendResponse) {
		if (request == 'closeTab') {
			if (tabId) {
				chrome.tabs.remove(tabId, function () {
					tabId = null;
				});
			}
			sendResponse("tab closed");
		}
		if (request == 'openTab') {
			// Check if the tab still exists before trying to update it
			if (tabId) {
				chrome.tabs.get(tabId, function (tab) {
					if (chrome.runtime.lastError) {
						// Tab doesn't exist anymore, create a new one
						createSingleExtensionTab().then(() => {
							if (tabId) {
								chrome.tabs.update(tabId, { "active": true });
							}
						});
					} else {
						chrome.tabs.update(tabId, { "active": true });
					}
				});
			} else {
				// No tab ID stored, create a new one
				createSingleExtensionTab().then(() => {
					if (tabId) {
						chrome.tabs.update(tabId, { "active": true });
					}
				});
			}
			sendResponse("tab open");
		}
		if (request == 'openTabSilent') {
			// Ensure only one extension tab exists
			if (tabId) {
				chrome.tabs.get(tabId, function (tab) {
					if (chrome.runtime.lastError) {
						// Tab doesn't exist anymore, create a new one
						createSingleExtensionTab();
					}
					// Tab exists, keep it as is
				});
			} else {
				// No tab ID stored, create a new one
				createSingleExtensionTab();
			}
			sendResponse("tab open silently");
		}
		if (request == 'ensureSingleTab') {
			// Perform immediate cleanup to ensure only one tab exists
			performPeriodicCleanup().then(() => {
				sendResponse("single tab ensured");
			});
			return true; // Indicates async response
		}
	}
);

// Listen for tab creation to prevent duplicates
chrome.tabs.onCreated.addListener(async (tab) => {
	const extensionUrl = chrome.runtime.getURL('popup/popup.html');
	if (tab.url === extensionUrl) {
		// A new extension tab was created, ensure we only keep one
		const tabs = await chrome.tabs.query({ url: extensionUrl });
		if (tabs.length > 1) {
			// Keep the newest tab and close others
			const sortedTabs = tabs.sort((a, b) => b.id - a.id);
			const tabsToClose = sortedTabs.slice(1).map(t => t.id);
			if (tabsToClose.length > 0) {
				await chrome.tabs.remove(tabsToClose);
				console.log(`Closed ${tabsToClose.length} duplicate extension tab(s)`);
			}
			// Update our tracked tab ID to the newest one
			tabId = sortedTabs[0].id;
		}
	}
});

// Listen for tab removal to clean up our reference
chrome.tabs.onRemoved.addListener((removedTabId) => {
	if (removedTabId === tabId) {
		tabId = null;
		console.log('Extension tab was closed, cleared tab reference');
	}
});

// Periodic cleanup to ensure only one extension tab exists
async function performPeriodicCleanup() {
	const extensionUrl = chrome.runtime.getURL('popup/popup.html');
	try {
		const tabs = await chrome.tabs.query({ url: extensionUrl });
		if (tabs.length > 1) {
			// Keep the most recently created tab and close others
			const sortedTabs = tabs.sort((a, b) => b.id - a.id);
			const tabsToClose = sortedTabs.slice(1).map(t => t.id);
			await chrome.tabs.remove(tabsToClose);
			tabId = sortedTabs[0].id;
			console.log(`Periodic cleanup: Closed ${tabsToClose.length} duplicate tab(s)`);
		} else if (tabs.length === 1) {
			// Update our reference to the existing tab
			tabId = tabs[0].id;
		} else if (tabs.length === 0) {
			// No extension tabs exist, clear our reference
			tabId = null;
		}
	} catch (error) {
		console.error('Error during periodic cleanup:', error);
	}
}

// Run cleanup every 2 minutes
setInterval(performPeriodicCleanup, 2 * 60 * 1000);