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
let persistentCases = new Map(); // Store persistent cases by case ID

// Function to close all existing extension tabs (except active ones)
async function closeExistingExtensionTabs() {
	const extensionUrl = chrome.runtime.getURL('popup/popup.html');
	try {
		const tabs = await chrome.tabs.query({ url: extensionUrl });
		if (tabs.length > 0) {
			const inactiveTabs = tabs.filter(tab => !tab.active);
			if (inactiveTabs.length > 0) {
				const tabIds = inactiveTabs.map(tab => tab.id);
				await chrome.tabs.remove(tabIds);
				console.log(`Closed ${inactiveTabs.length} inactive extension tab(s)`);
			}
		}
	} catch (error) {
		console.error('Error closing existing tabs:', error);
	}
}

// Function to create a single extension tab
async function createSingleExtensionTab() {
	try {
		// First, check if there's already an active extension tab
		const extensionUrl = chrome.runtime.getURL('popup/popup.html');
		const existingTabs = await chrome.tabs.query({ url: extensionUrl });

		// If there's an active tab, don't create a new one
		const activeTabs = existingTabs.filter(tab => tab.active);
		if (activeTabs.length > 0) {
			tabId = activeTabs[0].id;
			console.log(`Active extension tab already exists with ID: ${tabId}, skipping creation`);
			return;
		}

		// Close any inactive extension tabs before creating a new one
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
		try {
			if (request == 'closeTab') {
				if (tabId) {
					chrome.tabs.remove(tabId, function () {
						tabId = null;
					});
				}
				sendResponse("tab closed");
				return true;
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
								sendResponse("tab open");
							});
						} else {
							chrome.tabs.update(tabId, { "active": true });
							sendResponse("tab open");
						}
					});
				} else {
					// No tab ID stored, create a new one
					createSingleExtensionTab().then(() => {
						if (tabId) {
							chrome.tabs.update(tabId, { "active": true });
						}
						sendResponse("tab open");
					});
				}
				return true; // Indicates async response
			}
			if (request == 'openTabSilent') {
				// Ensure only one extension tab exists
				if (tabId) {
					chrome.tabs.get(tabId, function (tab) {
						if (chrome.runtime.lastError) {
							// Tab doesn't exist anymore, create a new one
							createSingleExtensionTab().then(() => {
								sendResponse("tab open silently");
							});
						} else {
							// Tab exists, keep it as is
							sendResponse("tab open silently");
						}
					});
				} else {
					// No tab ID stored, create a new one
					createSingleExtensionTab().then(() => {
						sendResponse("tab open silently");
					});
				}
				return true; // Indicates async response
			}
			if (request == 'ensureSingleTab') {
				// Perform immediate cleanup to ensure only one tab exists
				performPeriodicCleanup().then(() => {
					sendResponse("single tab ensured");
				});
				return true; // Indicates async response
			}

			// Handle persistent case management
			if (request.action === 'addCasesToPersistentSet') {
				try {
					if (request.cases && Array.isArray(request.cases)) {
						request.cases.forEach(caseObj => {
							persistentCases.set(caseObj.Id, {
								...caseObj,
								mode: request.currentMode,
								userId: request.currentUserId,
								userName: request.currentUserName,
								addedAt: Date.now()
							});
						});
						sendResponse({
							success: true,
							message: `Added ${request.cases.length} cases to persistent set`,
							count: persistentCases.size
						});
					} else {
						sendResponse({ success: false, message: 'Invalid cases data' });
					}
				} catch (error) {
					console.error('Error adding cases to persistent set:', error);
					sendResponse({ success: false, message: error.message });
				}
				return true;
			}

			if (request.action === 'removeCaseFromPersistentSet') {
				try {
					if (request.caseId) {
						const wasRemoved = persistentCases.delete(request.caseId);
						sendResponse({
							success: true,
							removed: wasRemoved,
							message: wasRemoved ? 'Case removed from persistent set' : 'Case not found in persistent set',
							count: persistentCases.size
						});
					} else {
						sendResponse({ success: false, message: 'No case ID provided' });
					}
				} catch (error) {
					console.error('Error removing case from persistent set:', error);
					sendResponse({ success: false, message: error.message });
				}
				return true;
			}

			if (request.action === 'getPersistentCaseCount') {
				try {
					sendResponse({
						success: true,
						count: persistentCases.size,
						cases: Array.from(persistentCases.values())
					});
				} catch (error) {
					console.error('Error getting persistent case count:', error);
					sendResponse({ success: false, message: error.message, count: 0 });
				}
				return true;
			}

			if (request.action === 'getPersistentCaseDetails') {
				try {
					const cases = Array.from(persistentCases.values());
					sendResponse({
						success: true,
						cases: cases,
						count: cases.length
					});
				} catch (error) {
					console.error('Error getting persistent case details:', error);
					sendResponse({ success: false, message: error.message, cases: [] });
				}
				return true;
			}

			// Default response for unhandled messages
			console.warn('Unhandled message type:', request);
			sendResponse({ success: false, message: 'Unknown request type', request: request });
			return true;
		} catch (error) {
			console.error('Error in message handler:', error);
			sendResponse({ success: false, message: 'Internal error', error: error.message });
			return true;
		}
	}
);

// Listen for tab creation to prevent duplicates
chrome.tabs.onCreated.addListener(async (tab) => {
	const extensionUrl = chrome.runtime.getURL('popup/popup.html');
	if (tab.url === extensionUrl) {
		// A new extension tab was created, ensure we only keep one active
		const tabs = await chrome.tabs.query({ url: extensionUrl });
		if (tabs.length > 1) {
			// Check if there's already an active extension tab
			const activeTabs = tabs.filter(t => t.active);
			const inactiveTabs = tabs.filter(t => !t.active);

			if (activeTabs.length > 0) {
				// There's an active tab, close all inactive ones including the new one if it's inactive
				if (inactiveTabs.length > 0) {
					const tabsToClose = inactiveTabs.map(t => t.id);
					await chrome.tabs.remove(tabsToClose);
					console.log(`Closed ${tabsToClose.length} inactive duplicate tab(s)`);
				}
				tabId = activeTabs[0].id;
			} else {
				// No active tabs, keep the newest and close others
				const sortedTabs = tabs.sort((a, b) => b.id - a.id);
				const tabsToClose = sortedTabs.slice(1).map(t => t.id);
				if (tabsToClose.length > 0) {
					await chrome.tabs.remove(tabsToClose);
					console.log(`Closed ${tabsToClose.length} duplicate extension tab(s)`);
				}
				tabId = sortedTabs[0].id;
			}
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
			// Check for active tabs first
			const activeTabs = tabs.filter(tab => tab.active);
			const inactiveTabs = tabs.filter(tab => !tab.active);

			if (activeTabs.length > 0) {
				// Keep the active tab and close all inactive ones
				if (inactiveTabs.length > 0) {
					const tabsToClose = inactiveTabs.map(t => t.id);
					await chrome.tabs.remove(tabsToClose);
					console.log(`Periodic cleanup: Closed ${tabsToClose.length} inactive duplicate tab(s)`);
				}
				tabId = activeTabs[0].id;
			} else {
				// No active tabs, keep the most recently created tab and close others
				const sortedTabs = tabs.sort((a, b) => b.id - a.id);
				const tabsToClose = sortedTabs.slice(1).map(t => t.id);
				await chrome.tabs.remove(tabsToClose);
				tabId = sortedTabs[0].id;
				console.log(`Periodic cleanup: Closed ${tabsToClose.length} duplicate tab(s)`);
			}
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