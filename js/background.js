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

// Helper to upsert a case record in the persistent map with bookkeeping
function upsertPersistentCase(caseObj, meta) {
	const existing = persistentCases.get(caseObj.Id) || {};
	const now = Date.now();
	const updated = {
		...existing,
		...caseObj,
		mode: meta.currentMode || existing.mode,
		userId: meta.currentUserId || existing.userId,
		userName: meta.currentUserName || existing.userName,
		addedAt: existing.addedAt || now,
		lastSeenMs: now,
		missingCount: 0,
		processed: existing.processed || false,
		processedReason: existing.processedReason || null,
	};
	persistentCases.set(caseObj.Id, updated);
}

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
							upsertPersistentCase(caseObj, {
								currentMode: request.currentMode,
								currentUserId: request.currentUserId,
								currentUserName: request.currentUserName
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

			// Sync cases seen in latest query. Do not auto-remove missing; only mark as missing and report back.
			if (request.action === 'syncPersistentCases') {
				try {
					const mode = request.mode;
					const ids = Array.isArray(request.caseIds) ? new Set(request.caseIds) : new Set();
					if (!mode) {
						sendResponse({ success: false, message: 'Mode is required' });
						return true;
					}
					const now = Date.now();
					const MAX_MISSING_DURATION_MS = 10 * 60 * 60 * 1000; // fallback cleanup after 10h
					const missingCases = [];
					const removedCases = [];
					let removed = 0;

					// First, update lastSeen for all ids in current result
					for (const id of ids) {
						const data = persistentCases.get(id);
						if (data && data.mode === mode) {
							data.lastSeenMs = now;
							data.missingCount = 0;
							persistentCases.set(id, data);
						}
					}

					// Mark missing cases, only remove if explicitly processed or very stale
					for (const [caseId, data] of persistentCases.entries()) {
						if (data.mode !== mode) continue;
						if (!ids.has(caseId)) {
							const updated = { ...data };
							updated.missingCount = (updated.missingCount || 0) + 1;
							updated.lastMissingAt = now;
							persistentCases.set(caseId, updated);
							// If already processed, we can safely remove
							if (updated.processed === true) {
								persistentCases.delete(caseId);
								removed++;
								removedCases.push({ id: caseId, data: updated, reason: 'processed' });
							} else if ((updated.lastSeenMs && (now - updated.lastSeenMs) > MAX_MISSING_DURATION_MS)) {
								// Fallback cleanup for very old entries
								persistentCases.delete(caseId);
								removed++;
								removedCases.push({ id: caseId, data: updated, reason: 'stale-timeout' });
							} else {
								missingCases.push({ id: caseId, data: updated });
							}
						}
					}

					sendResponse({ success: true, removed, count: persistentCases.size, removedCases, missingCases });
				} catch (error) {
					console.error('Error syncing persistent cases:', error);
					sendResponse({ success: false, message: error.message });
				}
				return true;
			}

			if (request.action === 'removeCaseFromPersistentSet') {
				try {
					if (request.caseId) {
						// Mark as processed before removal for bookkeeping
						const existing = persistentCases.get(request.caseId);
						if (existing) {
							persistentCases.set(request.caseId, { ...existing, processed: true, processedReason: request.reason || 'explicit-remove' });
						}
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
					let count;
					if (request.mode) {
						count = Array.from(persistentCases.values()).filter(v => v.mode === request.mode).length;
					} else {
						count = persistentCases.size;
					}
					sendResponse({
						success: true,
						count,
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