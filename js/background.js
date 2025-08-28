// Polyfills for Service Worker environment
if (typeof globalThis === 'undefined') {
	globalThis = self;
}

// Ensure Promise is available
if (typeof Promise === 'undefined') {
	// Basic Promise polyfill
	globalThis.Promise = function (executor) {
		let resolve, reject;
		this.then = function (onResolve, onReject) {
			resolve = onResolve;
			reject = onReject;
		};
		executor(resolve, reject);
	};
}

// Ensure fetch is available
if (typeof fetch === 'undefined') {
	globalThis.fetch = function (url, options = {}) {
		return new Promise((resolve, reject) => {
			// Use XMLHttpRequest as fallback
			const xhr = new XMLHttpRequest();
			xhr.open(options.method || 'GET', url);

			if (options.headers) {
				Object.keys(options.headers).forEach(key => {
					xhr.setRequestHeader(key, options.headers[key]);
				});
			}

			xhr.onload = function () {
				resolve({
					status: xhr.status,
					statusText: xhr.statusText,
					text: () => Promise.resolve(xhr.responseText),
					json: () => Promise.resolve(JSON.parse(xhr.responseText))
				});
			};

			xhr.onerror = function () {
				reject(new Error('Network error'));
			};

			xhr.send(options.body);
		});
	};
}

if (typeof XMLHttpRequest === 'undefined') {
	// Polyfill XMLHttpRequest for Service Worker
	globalThis.XMLHttpRequest = class XMLHttpRequest {
		constructor() {
			this.readyState = 0;
			this.status = 0;
			this.statusText = '';
			this.responseText = '';
			this.response = null;
			this.onreadystatechange = null;
			this.onload = null;
			this.onerror = null;
		}

		open(method, url, async = true) {
			this.method = method;
			this.url = url;
			this.async = async;
			this.readyState = 1;
		}

		setRequestHeader(name, value) {
			if (!this.headers) this.headers = {};
			this.headers[name] = value;
		}

		send(data) {
			this.readyState = 2;

			// Use fetch API instead of XMLHttpRequest
			fetch(this.url, {
				method: this.method,
				headers: this.headers,
				body: data
			})
				.then(response => {
					this.status = response.status;
					this.statusText = response.statusText;
					this.readyState = 3;
					if (this.onreadystatechange) this.onreadystatechange();

					return response.text();
				})
				.then(text => {
					this.responseText = text;
					this.response = text;
					this.readyState = 4;
					if (this.onreadystatechange) this.onreadystatechange();
					if (this.onload) this.onload();
				})
				.catch(error => {
					this.readyState = 4;
					if (this.onreadystatechange) this.onreadystatechange();
					if (this.onerror) this.onerror(error);
				});
		}

		abort() {
			this.readyState = 0;
		}

		getResponseHeader(name) {
			return this.headers ? this.headers[name] : null;
		}

		getAllResponseHeaders() {
			if (!this.headers) return '';
			return Object.entries(this.headers)
				.map(([name, value]) => `${name}: ${value}`)
				.join('\r\n');
		}
	};
}

// Polyfill for localStorage in Service Worker
if (typeof localStorage === 'undefined') {
	globalThis.localStorage = {
		getItem: (key) => {
			// Use chrome.storage.local as fallback
			return new Promise((resolve) => {
				chrome.storage.local.get([key], (result) => {
					resolve(result[key] || null);
				});
			});
		},
		setItem: (key, value) => {
			// Use chrome.storage.local as fallback
			chrome.storage.local.set({ [key]: value });
		},
		removeItem: (key) => {
			// Use chrome.storage.local as fallback
			chrome.storage.local.remove([key]);
		},
		clear: () => {
			// Use chrome.storage.local as fallback
			chrome.storage.local.clear();
		}
	};
}



// Polyfill for console in Service Worker
if (typeof console === 'undefined') {
	globalThis.console = {
		log: (...args) => { },
		error: (...args) => { },
		warn: (...args) => { },
		info: (...args) => { },
		debug: (...args) => { }
	};
}

// Minimize background noise in production: silence log/info/debug/trace while keeping warnings/errors
try {
	if (typeof console !== 'undefined') {
		const __warn = console.warn && console.warn.bind(console);
		const __error = console.error && console.error.bind(console);
		console.log = () => { };
		console.info = () => { };
		console.debug = () => { };
		console.trace = () => { };
		// Keep warn/error intact
		if (__warn) console.warn = __warn;
		if (__error) console.error = __error;
	}
} catch { }

// Polyfill for sessionStorage in Service Worker
if (typeof sessionStorage === 'undefined') {
	globalThis.sessionStorage = {
		getItem: (key) => {
			// Use chrome.storage.local as fallback - return synchronously
			let result = null;
			chrome.storage.local.get([key], (storageResult) => {
				result = storageResult[key] || null;
			});
			return result;
		},
		setItem: (key, value) => {
			// Use chrome.storage.local as fallback (session storage not available in Service Worker)
			chrome.storage.local.set({ [key]: value });
		},
		removeItem: (key) => {
			// Use chrome.storage.local as fallback
			chrome.storage.local.remove([key]);
		},
		clear: () => {
			// Use chrome.storage.local as fallback
			chrome.storage.local.clear();
		}
	};
}

// Import jsforce for Salesforce API operations
importScripts('vendor/jsforce.js');

chrome.runtime.onInstalled.addListener(details => {
	if (details.reason === chrome.runtime.OnInstalledReason.INSTALL) {
		chrome.alarms.create('SigSEV2Triaging', { periodInMinutes: 1 });
		// Create daily cleanup alarm at 5:30 AM IST (12:00 AM UTC)
		chrome.alarms.create('DailyCaseCleanup', {
			when: getNextDailyCleanupTime(),
			periodInMinutes: 24 * 60 // 24 hours
		});
	}
	// Clean up any existing tabs on installation/update
	closeExistingExtensionTabs();
});

// Also clean up on startup
chrome.runtime.onStartup.addListener(() => {
	closeExistingExtensionTabs();
});

let tabId;
// Removed: persistentCases (no longer tracking cases in background)

// Helper function to get next daily cleanup time (5:30 AM IST = 12:00 AM UTC)
function getNextDailyCleanupTime() {
	const now = new Date();
	const utcMidnight = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate() + 1, 0, 0, 0));
	return utcMidnight.getTime();
}

// Perform daily cleanup of tracked flags and stale pending data
async function performDailyCaseCleanup() {
	try {
		let cleanedCount = 0;

		// Remove session tracked_* keys stored in chrome.storage.local
		try {
			const allItems = await new Promise(resolve => chrome.storage.local.get(null, resolve));
			const keysToRemove = Object.keys(allItems).filter(key => key && key.startsWith('tracked_'));
			if (keysToRemove.length > 0) {
				await new Promise(resolve => chrome.storage.local.remove(keysToRemove, resolve));
				cleanedCount += keysToRemove.length;
			}
		} catch (e) {
			// ignore storage errors for tracked_* cleanup
		}

		// Cleanup stale pending tracking data (> 2 days old)
		try {
			let pending = [];
			let useChromeStorageForPending = false;
			if (typeof localStorage !== 'undefined') {
				try { pending = JSON.parse(localStorage.getItem('pendingTrackingData') || '[]'); } catch { pending = []; }
			} else {
				const result = await new Promise(resolve => chrome.storage.local.get(['pendingTrackingData'], resolve));
				pending = Array.isArray(result?.pendingTrackingData) ? result.pendingTrackingData : [];
				useChromeStorageForPending = true;
			}

			if (Array.isArray(pending) && pending.length) {
				const cutoff = Date.now() - (2 * 24 * 60 * 60 * 1000);
				const before = pending.length;
				const filtered = pending.filter(item => !item?.timestamp || item.timestamp >= cutoff);
				const removed = before - filtered.length;
				if (removed > 0) {
					cleanedCount += removed;
					if (typeof localStorage !== 'undefined') {
						localStorage.setItem('pendingTrackingData', JSON.stringify(filtered));
					} else if (useChromeStorageForPending) {
						await new Promise(resolve => chrome.storage.local.set({ pendingTrackingData: filtered }, resolve));
					}
				}
			}
		} catch (e) {
			// ignore pending cleanup errors
		}

		// Notify UI (popup) that daily cleanup completed
		try {
			chrome.runtime.sendMessage({ action: 'dailyCleanupCompleted', cleanedCases: cleanedCount });
		} catch (e) {
			// ignore notification errors
		}

		return cleanedCount;
	} catch (error) {
		console.error('Error performing daily case cleanup:', error);
		throw error;
	}
}

// Function to process cases in background for assignment detection using new approach
async function processCasesInBackground(cases, connectionInfo, currentMode, currentUserId, currentUserName) {
	// Disabled: background case processing has been turned off
	console.log('Background case processing is disabled');
	return;
}

// New Queue-based case processing system for TrackedChange approach - FIRST CLASS
class NewCaseProcessingQueue {
	constructor(connection, mode, userId, userName) {
		this.connection = connection;
		this.mode = mode;
		this.userId = userId;
		this.userName = userName;
		this.queue = [];
		this.processing = false;
		this.stats = {
			total: 0,
			processed: 0,
			successful: 0,
			failed: 0,
			assignmentsDetected: 0,
			outsideBusinessWindow: 0, // Track cases filtered out due to time constraints
			startTime: null,
			endTime: null
		};
		this.batchSize = 15; // Optimal batch size for Salesforce API
		this.delayBetweenBatches = 800; // 800ms delay between batches
		this.maxConcurrentBatches = 2; // Process up to 2 batches concurrently
		this.activeBatches = 0;
	}

	// Add cases to the processing queue
	async addCases(cases) {
		this.queue = [...cases];
		this.stats.total = cases.length;
		this.stats.startTime = Date.now();

		// Clean localStorage to remove expired entries
		await this.cleanLocalStorage();
	}

	// Process all cases in the queue
	async processAll() {
		if (this.processing) {
			return;
		}

		this.processing = true;

		try {
			// Process in batches with progress tracking
			while (this.queue.length > 0) {
				// Wait if we've reached max concurrent batches
				while (this.activeBatches >= this.maxConcurrentBatches) {
					await this.delay(100);
				}

				// Take next batch
				const batch = this.queue.splice(0, this.batchSize);
				this.activeBatches++;

				// Process batch asynchronously
				this.processBatch(batch).finally(() => {
					this.activeBatches--;
				});

				// Progress update
				this.updateProgress();

				// Delay before next batch (if not the last batch)
				if (this.queue.length > 0) {
					await this.delay(this.delayBetweenBatches);
				}
			}

			// Wait for all active batches to complete
			while (this.activeBatches > 0) {
				await this.delay(100);
			}

			this.stats.endTime = Date.now();

		} catch (error) {
			// Handle error silently
		} finally {
			this.processing = false;
		}

		return this.stats;
	}

	// Process a single batch of cases
	async processBatch(batch) {
		const batchId = `batch_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
		const batchStartTime = Date.now();

		try {
			// Get case IDs for this batch
			const caseIds = batch.map(c => c.Id);

			// Process each case in the batch using new approach
			for (const caseRecord of batch) {
				await this.processSingleCaseNewApproach(caseRecord);
			}

		} catch (error) {
			this.stats.failed += batch.length;
		}
	}

	// Process a single case using the new TrackedChange approach
	async processSingleCaseNewApproach(caseRecord) {
		try {
			// Check if case has already been tracked to prevent duplicates using sessionStorage logic
			const trackingKey = `tracked_${caseRecord.Id}`;
			const isTracked = await checkIfTracked(trackingKey);
			if (isTracked) {
				this.stats.processed++;
				this.stats.successful++;
				return;
			}

			// Check if current user assigned the case using trackNewCaseFromHistory function logic
			console.log(`Processing case ${caseRecord.Id}, checking if current user assigned it...`);
			const assignmentCheck = await this.checkIfCurrentUserAssignedCase(caseRecord.Id);
			console.log(`Result for case ${caseRecord.Id}: assignmentCheck =`, assignmentCheck);

			if (assignmentCheck.qualifies) {
				// Check if action date is within business day window
				if (!this.isWithinBusinessDayWindow(assignmentCheck.actionDate)) {
					console.log(`Case ${caseRecord.Id} action date outside business day window, skipping`);
					this.stats.outsideBusinessWindow++;
					return;
				}

				console.log(`Case ${caseRecord.Id} action date within business day window, proceeding`);
				// Determine action type from routing logs
				let actionType = 'New Case';
				let routingLogReason = '';

				try {
					const routingLogQuery = `SELECT Transfer_Reason__c FROM Case_Routing_Log__c WHERE Case__c = '${caseRecord.Id}' ORDER BY CreatedDate DESC LIMIT 20`;
					let routingLogResult = await this.connection.query(routingLogQuery);

					if (typeof routingLogResult === 'string') {
						routingLogResult = JSON.parse(routingLogResult);
					}

					if (routingLogResult.records && routingLogResult.records.length > 0) {
						routingLogReason = routingLogResult.records[0].Transfer_Reason__c;
						if (routingLogReason === 'GHO') {
							actionType = 'GHO';
						}
					}
				} catch (routingLogError) {
					console.warn(`Could not query routing logs for case ${caseRecord.Id}:`, routingLogError);
				}

				console.log(`Case ${caseRecord.Id} assigned by current user. Action type determined: ${actionType}`);

				// Get case details for tracking
				const caseDetailsQuery = `SELECT Id, CaseNumber, Severity_Level__c, CaseRoutingTaxonomy__r.Name, Owner.Name FROM Case WHERE Id = '${caseRecord.Id}'`;
				let caseDetailsResult = await this.connection.query(caseDetailsQuery);

				if (typeof caseDetailsResult === 'string') {
					caseDetailsResult = JSON.parse(caseDetailsResult);
				}

				if (caseDetailsResult.records && caseDetailsResult.records.length > 0) {
					const caseDetail = caseDetailsResult.records[0];
					const cloud = caseDetail.CaseRoutingTaxonomy__r && caseDetail.CaseRoutingTaxonomy__r.Name ?
						caseDetail.CaseRoutingTaxonomy__r.Name.split('-')[0] : '';

					// Get assignedTo value from CaseHistory
					const historyQuery = `SELECT CaseId, CreatedById, CreatedDate, Field, NewValue FROM CaseHistory WHERE CaseId = '${caseRecord.Id}' AND CreatedById = '${this.userId}' AND (Field = 'Routing_Status__c' OR Field = 'Owner') ORDER BY CreatedDate ASC LIMIT 20`;
					let historyResult = await this.connection.query(historyQuery);

					if (typeof historyResult === 'string') {
						historyResult = JSON.parse(historyResult);
					}

					let assignedTo = '';
					let actionDate = caseRecord.CreatedDate;
					if (historyResult.records && historyResult.records.length > 0) {
						// Find Owner field change following Manually Assigned status
						let foundManuallyAssigned = false;
						for (const hist of historyResult.records) {
							if (hist.Field === 'Routing_Status__c' && hist.NewValue && hist.NewValue.startsWith('Manually Assigned')) {
								foundManuallyAssigned = true;
								actionDate = hist.CreatedDate;
							} else if (hist.Field === 'Owner' && hist.NewValue && foundManuallyAssigned) {
								if (!/^005[\w]{12,15}$/.test(hist.NewValue)) {
									assignedTo = hist.NewValue;
									break;
								} else if (hist.NewValue && /^005[\w]{12,15}$/.test(hist.NewValue)) {
									// Query User object to get name from ID
									try {
										const userQuery = `SELECT Name, Email, Username FROM User WHERE Id = '${hist.NewValue}'`;
										let userResult = await this.connection.query(userQuery);

										if (typeof userResult === 'string') {
											userResult = JSON.parse(userResult);
										}

										if (userResult.records && userResult.records.length > 0) {
											assignedTo = userResult.records[0].Name;
											break;
										} else {
											assignedTo = hist.NewValue; // Fallback to ID if user not found
											break;
										}
									} catch (userQueryError) {
										console.error('Error querying User object:', userQueryError);
										assignedTo = hist.NewValue; // Fallback to ID if query fails
										break;
									}
								}
							}
						}
					}

					// Fallback to case owner if no history found
					if (!assignedTo && caseDetail.Owner && caseDetail.Owner.Name) {
						assignedTo = caseDetail.Owner.Name;
					}

					console.log(`Final assignedTo value: ${assignedTo}`);

					// Generate unique tracking ID to prevent Google Sheet duplicates
					const trackingId = this.generateTrackingId(caseDetail.CaseNumber, actionDate, actionType, assignedTo);
					console.log(`Generated tracking ID: ${trackingId}`);

					// Check if this exact tracking entry has already been sent to Google Sheets
					if (this.checkIfAlreadySentToSheets(trackingId)) {
						console.log(`Case ${caseDetail.CaseNumber} with tracking ID ${trackingId} has already been sent to Google Sheets. Skipping to prevent duplicate.`);
						this.stats.processed++;
						this.stats.successful++;
						return;
					}

					// Prepare tracking data
					const trackingData = {
						createdDate: actionDate, // Use the actual action date from CaseHistory
						caseNumber: caseDetail.CaseNumber,
						severity: caseDetail.Severity_Level__c,
						actionType: actionType,
						cloud: cloud,
						mode: this.mode,
						userName: this.userName,
						newValue: assignedTo,
						trackingId: trackingId // Include tracking ID for reference
					};

					console.log(`Tracking data prepared for case ${caseDetail.CaseNumber}:`, trackingData);
					console.log(`Mode: ${this.mode} - Will be sent to '${this.mode === 'premier' ? 'premier' : 'signature'}' sheet`);

					// Try to send message to popup first
					let messageSent = false;
					try {
						const extensionUrl = chrome.runtime.getURL('popup/popup.html');
						const tabs = await chrome.tabs.query({ url: extensionUrl });
						for (const tab of tabs) {
							chrome.tabs.sendMessage(tab.id, {
								action: 'trackActionFromBackground',
								data: trackingData
							}).catch(() => {
								// Ignore errors if tab doesn't have message listener
							});
						}
						messageSent = tabs.length > 0;
					} catch (error) {
						console.log('Error sending message to popup:', error);
					}

					// If popup is not available, store the tracking data for later processing
					if (!messageSent) {
						this.storePendingTrackingData(trackingData);
						console.log('Popup not available, stored tracking data for later processing:', trackingData);
						console.log('Total pending tracking data items:', this.getPendingTrackingData().length);
					} else {
						console.log('Message sent to popup successfully for case:', caseDetail.CaseNumber);
					}

					// Don't mark as sent yet - wait for popup to confirm successful Google Sheets update
					console.log(`Tracking ID ${trackingId} prepared, will be marked as sent after successful Google Sheets update`);

					console.log(`Case ${caseDetail.CaseNumber} processed successfully. Assignment detected: true, Action type: ${actionType}`);

					console.log(`About to increment assignmentsDetected. Current value: ${this.stats.assignmentsDetected}`);
					this.stats.assignmentsDetected++;
					console.log(`assignmentsDetected incremented to: ${this.stats.assignmentsDetected}`);

					// Mark case as tracked in sessionStorage to prevent duplicates
					this.markAsTracked(trackingKey);

					// Mark case as processed in persistent set
					const existing = persistentCases.get(caseRecord.Id);
					if (existing) {
						persistentCases.set(caseRecord.Id, {
							...existing,
							processed: true,
							processedReason: 'background-tracked-change-processed',
							lastProcessed: Date.now()
						});
					}

				} else {
					console.warn(`No case details found for case ${caseRecord.Id}. CaseDetailsResult:`, caseDetailsResult);
					console.warn(`This case will not be tracked and assignmentsDetected will not be incremented.`);
				}
			}

			this.stats.processed++;
			this.stats.successful++;

		} catch (error) {
			console.error('Error processing case:', error);
			console.error('Error stack:', error.stack);

			// Increment failed count
			this.stats.failed++;

			// Recovery attempt for Case_Routing_Logs__c error
			if (error.message && error.message.includes('Case_Routing_Logs__c')) {
				console.warn('Case_Routing_Logs__c error detected, attempting recovery...');

				try {
					const caseDetailsQuery = `SELECT Id, CaseNumber, Severity_Level__c, CaseRoutingTaxonomy__r.Name, Owner.Name FROM Case WHERE Id = '${caseRecord.Id}'`;
					let caseDetailsResult = await this.connection.query(caseDetailsQuery);

					if (typeof caseDetailsResult === 'string') {
						caseDetailsResult = JSON.parse(caseDetailsResult);
					}

					if (caseDetailsResult.records && caseDetailsResult.records.length > 0) {
						const caseDetail = caseDetailsResult.records[0];
						const cloud = caseDetail.CaseRoutingTaxonomy__r && caseDetail.CaseRoutingTaxonomy__r.Name ?
							caseDetail.CaseRoutingTaxonomy__r.Name.split('-')[0] : '';

						// Determine action type from routing logs
						let actionType = 'New Case';
						try {
							const routingLogQuery = `SELECT Transfer_Reason__c FROM Case_Routing_Log__c WHERE Case__c = '${caseRecord.Id}' ORDER BY CreatedDate DESC LIMIT 20`;
							let routingLogResult = await this.connection.query(routingLogQuery);

							if (typeof routingLogResult === 'string') {
								routingLogResult = JSON.parse(routingLogResult);
							}

							if (routingLogResult.records && routingLogResult.records.length > 0) {
								const routingLogReason = routingLogResult.records[0].Transfer_Reason__c;
								if (routingLogReason === 'GHO') {
									actionType = 'GHO';
								}
							}
						} catch (routingLogError) {
							console.warn(`Recovery: Could not query routing logs for case ${caseRecord.Id}:`, routingLogError);
						}

						// Get action date from CaseHistory
						let actionDate = caseRecord.CreatedDate;
						try {
							const historyQuery = `SELECT CaseId, CreatedById, CreatedDate, Field, NewValue FROM CaseHistory WHERE CaseId = '${caseRecord.Id}' AND CreatedById = '${this.userId}' AND (Field = 'Routing_Status__c' OR Field = 'Owner') ORDER BY CreatedDate ASC LIMIT 20`;
							let historyResult = await this.connection.query(historyQuery);

							if (typeof historyResult === 'string') {
								historyResult = JSON.parse(historyResult);
							}

							if (historyResult.records && historyResult.records.length > 0) {
								for (const hist of historyResult.records) {
									if (hist.Field === 'Routing_Status__c' && hist.NewValue && hist.NewValue.startsWith('Manually Assigned')) {
										actionDate = hist.CreatedDate;
										break;
									}
								}
							}
						} catch (historyError) {
							console.warn(`Recovery: Could not query CaseHistory for case ${caseRecord.Id}:`, historyError);
						}

						// Check if action date is within business day window
						if (!this.isWithinBusinessDayWindow(actionDate)) {
							console.log(`Recovery: Case ${caseDetail.CaseNumber} action date outside business day window, skipping`);
							this.stats.outsideBusinessWindow++;
							return;
						}

						console.log(`Recovery: Case ${caseDetail.CaseNumber} action date within business day window, proceeding`);

						// Generate tracking ID and check for duplicates
						const trackingId = this.generateTrackingId(caseDetail.CaseNumber, actionDate, actionType, caseDetail.Owner ? caseDetail.Owner.Name : '');

						if (this.checkIfAlreadySentToSheets(trackingId)) {
							this.stats.processed++;
							this.stats.successful++;
							return;
						}

						// Prepare and store tracking data
						const trackingData = {
							createdDate: actionDate,
							caseNumber: caseDetail.CaseNumber,
							severity: caseDetail.Severity_Level__c,
							actionType: actionType,
							cloud: cloud,
							mode: this.mode,
							userName: this.userName,
							newValue: caseDetail.Owner ? caseDetail.Owner.Name : '',
							trackingId: trackingId
						};

						console.log(`Recovery: Tracking data prepared for case ${caseDetail.CaseNumber}:`, trackingData);
						console.log(`Recovery: Mode: ${this.mode} - Will be sent to '${this.mode === 'premier' ? 'premier' : 'signature'}' sheet`);

						this.storePendingTrackingData(trackingData);

						this.stats.assignmentsDetected++;
						this.stats.processed++;
						this.stats.successful++;
						return;
					}
				} catch (recoveryError) {
					console.error('Error during recovery attempt:', recoveryError);
				}
			}

			this.stats.failed++;
		}
	}

	// Store pending tracking data when popup is not available
	storePendingTrackingData(trackingData) {
		try {
			const pendingData = JSON.parse(localStorage.getItem('pendingTrackingData') || '[]');
			pendingData.push({
				...trackingData,
				timestamp: Date.now()
			});
			// Keep only last 100 pending items to prevent memory issues
			if (pendingData.length > 100) {
				pendingData.splice(0, pendingData.length - 100);
			}
			localStorage.setItem('pendingTrackingData', JSON.stringify(pendingData));
			console.log(`Stored pending tracking data. Total pending items: ${pendingData.length}`);
		} catch (error) {
			console.error('Error storing pending tracking data:', error);
		}
	}

	// Get pending tracking data
	getPendingTrackingData() {
		try {
			return JSON.parse(localStorage.getItem('pendingTrackingData') || '[]');
		} catch (error) {
			console.error('Error getting pending tracking data:', error);
			return [];
		}
	}

	// Clear pending tracking data
	clearPendingTrackingData() {
		try {
			localStorage.removeItem('pendingTrackingData');
		} catch (error) {
			console.error('Error clearing pending tracking data:', error);
		}
	}

	// Helper function to resolve Salesforce User ID to User Name
	async resolveUserIdToName(userId) {
		try {
			if (!userId || !/^005[\w]{12,15}$/.test(userId)) {
				return userId; // Return as-is if not a valid User ID
			}

			const userQuery = `SELECT Name, Email, Username FROM User WHERE Id = '${userId}'`;
			const userResult = await this.connection.query(userQuery);

			if (userResult.records && userResult.records.length > 0) {
				const userName = userResult.records[0].Name;
				console.log(`Resolved User ID ${userId} to name: ${userName}`);
				return userName;
			} else {
				console.log(`User not found for ID: ${userId}`);
				return userId; // Fallback to ID if user not found
			}
		} catch (error) {
			console.error('Error resolving User ID to name:', error);
			return userId; // Fallback to ID if query fails
		}
	}

	// Mark case as tracked
	markAsTracked(trackingKey) {
		try {
			sessionStorage.setItem(trackingKey, 'true');
		} catch (error) {
			console.error('Error marking case as tracked:', error);
		}
	}

	// Generate a unique tracking identifier to prevent Google Sheet duplicates
	generateTrackingId(caseNumber, actionDate, actionType, assignedTo) {
		try {
			// Create a unique identifier based on case details
			const dateStr = actionDate instanceof Date ? actionDate.toISOString().split('T')[0] : actionDate.split('T')[0];
			const uniqueId = `${caseNumber}_${dateStr}_${actionType}_${assignedTo}`;
			return uniqueId.replace(/[^a-zA-Z0-9_-]/g, '_'); // Sanitize for storage
		} catch (error) {
			console.error('Error generating tracking ID:', error);
			// Fallback to a simple ID
			return `tracking_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
		}
	}

	// Mark case as tracked in this session
	markAsTracked(trackingId) {
		try {
			if (typeof sessionStorage !== 'undefined') {
				sessionStorage.setItem(trackingId, 'true');
				console.log(`Marked ${trackingId} as tracked in session`);
			}
		} catch (error) {
			console.error('Error marking as tracked:', error);
		}
	}

	// Check if this exact tracking entry has already been sent to Google Sheets
	checkIfAlreadySentToSheets(trackingId) {
		try {
			const rawData = localStorage.getItem('sentToSheets');

			// Debug logging to understand what type of data we're getting
			if (rawData !== null) {
				console.log(`localStorage data type: ${typeof rawData}, length: ${rawData ? rawData.length : 0}`);
				if (typeof rawData === 'string' && rawData.length > 100) {
					console.log(`localStorage data preview: ${rawData.substring(0, 100)}...`);
				}
			}

			// If no data exists, nothing has been sent
			if (!rawData) {
				console.log('No localStorage data found, returning false');
				return false;
			}

			let sentEntries = [];
			try {
				sentEntries = JSON.parse(rawData);
				console.log('Parsed localStorage data successfully');
			} catch (parseError) {
				console.warn('localStorage data is not valid JSON, clearing it');
				try { localStorage.removeItem('sentToSheets'); } catch { }
				return false;
			}

			if (!Array.isArray(sentEntries)) {
				console.warn('Parsed localStorage data is not an array, clearing it');
				try { localStorage.removeItem('sentToSheets'); } catch { }
				return false;
			}

			for (const entry of sentEntries) {
				if (entry && entry.trackingId === trackingId) {
					console.log(`Found existing tracking ID ${trackingId} in localStorage`);
					return true;
				}
			}

			console.log(`Tracking ID ${trackingId} not found in localStorage`);
			return false;
		} catch (error) {
			console.error('Error checking if already sent to Sheets:', error);
			return false;
		}
	}

	// Reset localStorage if it gets corrupted
	resetLocalStorage() {
		try {
			console.log('Resetting localStorage due to corruption...');
			localStorage.removeItem('sentToSheets');
			console.log('localStorage reset successfully');
		} catch (error) {
			console.error('Failed to reset localStorage:', error);
		}
	}

	// Mark tracking entry as sent to Google Sheets
	markAsSentToSheets(trackingId) {
		try {
			// Validate trackingId is a string
			if (typeof trackingId !== 'string') {
				console.error('Invalid trackingId type:', typeof trackingId, trackingId);
				return;
			}

			const rawData = localStorage.getItem('sentToSheets');
			let sentEntries = [];

			try {
				if (rawData && typeof rawData === 'string') {
					// Check if data is corrupted
					if (rawData.includes('[object Promise]') || rawData.includes('[object Promise]')) {
						console.warn('Corrupted localStorage data detected. Clearing and resetting.');
						localStorage.removeItem('sentToSheets');
					} else {
						sentEntries = JSON.parse(rawData);
					}
				}
			} catch (parseError) {
				console.warn('Failed to parse existing localStorage data. Clearing and resetting.');
				localStorage.removeItem('sentToSheets');
			}

			// Ensure sentEntries is an array
			if (!Array.isArray(sentEntries)) {
				sentEntries = [];
			}

			// Filter out any non-string entries
			sentEntries = sentEntries.filter(entry => typeof entry === 'string');

			sentEntries.push(trackingId);

			// Keep only last 1000 entries to prevent localStorage from growing too large
			if (sentEntries.length > 1000) {
				sentEntries.splice(0, sentEntries.length - 1000);
			}

			localStorage.setItem('sentToSheets', JSON.stringify(sentEntries));
			console.log(`Marked tracking ID ${trackingId} as sent to Google Sheets`);
		} catch (error) {
			console.error('Error marking as sent to sheets:', error);
		}
	}

	// Check if date falls within business day window (5:30 AM of day 1 to 5:30 AM of day 2)
	isWithinBusinessDayWindow(actionDate) {
		try {
			const date = actionDate instanceof Date ? actionDate : new Date(actionDate);
			const now = new Date();
			const istTime = new Date(now.toLocaleString("en-US", { timeZone: "Asia/Kolkata" }));
			const actionDateIst = new Date(date.toLocaleString("en-US", { timeZone: "Asia/Kolkata" }));

			const businessDayStart = new Date(istTime);
			businessDayStart.setHours(5, 30, 0, 0); // 5:30 AM of current day

			if (istTime < businessDayStart) {
				businessDayStart.setDate(businessDayStart.getDate() - 1);
			}

			const businessDayEnd = new Date(businessDayStart);
			businessDayEnd.setDate(businessDayEnd.getDate() + 1);
			businessDayEnd.setHours(5, 30, 0, 0); // 5:30 AM next day

			const isWithinWindow = actionDateIst >= businessDayStart && actionDateIst <= businessDayEnd;

			// Log detailed information about the date filtering
			console.log(`=== Date Filtering for Case ===`);
			console.log(`Action Date (IST): ${actionDateIst.toLocaleString("en-US", { timeZone: "Asia/Kolkata" })}`);
			console.log(`Business Day Start (IST): ${businessDayStart.toLocaleString("en-US", { timeZone: "Asia/Kolkata" })}`);
			console.log(`Business Day End (IST): ${businessDayEnd.toLocaleString("en-US", { timeZone: "Asia/Kolkata" })}`);
			console.log(`Is Within Business Window: ${isWithinWindow}`);
			console.log(`Final Result: ${isWithinWindow ? 'PROCESS' : 'SKIP'}`);
			console.log(`================================`);

			// Only process if case is within business hours (5:30 AM to 5:30 AM next day)
			// This automatically handles the case where we're before 5:30 AM and still in yesterday's business day
			return isWithinWindow;
		} catch (error) {
			console.error('Error checking business day window:', error);
			return false;
		}
	}

	// Check if current user assigned the case
	async checkIfCurrentUserAssignedCase(caseId) {
		try {
			const historyQuery = `SELECT CaseId, CreatedById, CreatedDate, Field, NewValue FROM CaseHistory WHERE CaseId = '${caseId}' AND CreatedById = '${this.userId}' AND (Field = 'Routing_Status__c' OR Field = 'Owner') ORDER BY CreatedDate ASC`;
			let historyResult = await this.connection.query(historyQuery);

			if (typeof historyResult === 'string') {
				historyResult = JSON.parse(historyResult);
			}

			if (!historyResult.records || historyResult.records.length === 0) {
				return { qualifies: false, actionDate: null };
			}

			// Check for manual assignment pattern: Routing_Status__c = "Manually Assigned" followed by Owner change
			let foundManuallyAssigned = false;
			let foundOwnerChange = false;
			let ownerName = '';
			let actionDate = null;

			for (const hist of historyResult.records) {
				if (hist.Field === 'Routing_Status__c' && hist.NewValue && hist.NewValue.startsWith('Manually Assigned')) {
					foundManuallyAssigned = true;
					actionDate = hist.CreatedDate;
				} else if (hist.Field === 'Owner' && hist.NewValue && foundManuallyAssigned) {
					if (!/^005[\w]{12,15}$/.test(hist.NewValue)) {
						foundOwnerChange = true;
						ownerName = hist.NewValue;
					}
				}
			}

			if (foundManuallyAssigned && foundOwnerChange) {
				return { qualifies: true, actionDate: actionDate };
			}

			return { qualifies: false, actionDate: null };
		} catch (error) {
			console.error(`Error in checkIfCurrentUserAssignedCase for case ${caseId}:`, error);
			return { qualifies: false, actionDate: null };
		}
	}

	// Update and display progress
	updateProgress() {
		// Progress tracking logic (silent)
	}

	// Utility function for delays
	delay(ms) {
		return new Promise(resolve => setTimeout(resolve, ms));
	}

	// Method to store pending tracking data for later processing
	storePendingTrackingData(trackingData) {
		storePendingTrackingData(trackingData);
	}

	// Method to clean localStorage and remove expired entries
	async cleanLocalStorage() {
		try {
			console.log('Cleaning localStorage to remove expired entries...');
			const rawData = localStorage.getItem('sentToSheets');

			if (!rawData) {
				console.log('No localStorage data to clean');
				return;
			}

			try {
				const sentEntries = JSON.parse(rawData);
				if (!Array.isArray(sentEntries)) {
					console.log('Invalid localStorage format, clearing completely');
					localStorage.removeItem('sentToSheets');
					return;
				}

				// Filter out entries older than current business day to prevent localStorage from growing indefinitely
				const now = new Date();
				const istTime = new Date(now.toLocaleString("en-US", { timeZone: "Asia/Kolkata" }));

				// Calculate start of current business day (5:30 AM of previous day to 5:30 AM of current day)
				// If current time is before 5:30 AM, we're still in yesterday's business day
				// If current time is after 5:30 AM, we're in today's business day
				const currentBusinessDayStart = new Date(istTime);
				currentBusinessDayStart.setHours(5, 30, 0, 0); // 5:30 AM of current day

				// If it's before 5:30 AM today, use yesterday's 5:30 AM as business day start
				if (istTime < currentBusinessDayStart) {
					currentBusinessDayStart.setDate(currentBusinessDayStart.getDate() - 1);
				}

				const filteredEntries = sentEntries.filter(entry => {
					try {
						// Extract date from tracking ID format: CASE_NUMBER_DATE_ACTIONTYPE_ASSIGNEDTO
						const parts = entry.split('_');
						if (parts.length >= 3) {
							const dateStr = parts[1]; // Format: YYYY-MM-DD
							const entryDate = new Date(dateStr);
							// Keep entries from current business day onwards
							return entryDate >= currentBusinessDayStart;
						}
						return false; // Remove malformed entries
					} catch (error) {
						console.log(`Error parsing date from entry ${entry}:`, error);
						return false; // Remove entries with invalid dates
					}
				});

				if (filteredEntries.length !== sentEntries.length) {
					const removedCount = sentEntries.length - filteredEntries.length;
					console.log(`Removed ${removedCount} expired entries from localStorage`);
					localStorage.setItem('sentToSheets', JSON.stringify(filteredEntries));
				} else {
					console.log('No expired entries to remove');
				}

			} catch (parseError) {
				console.log('Error parsing localStorage data, clearing completely:', parseError);
				localStorage.removeItem('sentToSheets');
			}

			// Also clean popup localStorage if available
			try {
				const extensionUrl = chrome.runtime.getURL('popup/popup.html');
				const tabs = await chrome.tabs.query({ url: extensionUrl });
				for (const tab of tabs) {
					chrome.tabs.sendMessage(tab.id, {
						action: 'cleanLocalStorage'
					}).catch(() => {
						// Ignore errors if tab doesn't have message listener
					});
				}
			} catch (error) {
				console.log('Could not clean popup localStorage:', error);
			}
		} catch (error) {
			console.error('Error cleaning localStorage:', error);
		}
	}

	// Method to manually clear localStorage (for debugging/testing)
	async clearLocalStorage() {
		try {
			console.log('Manually clearing localStorage...');
			localStorage.removeItem('sentToSheets');
			console.log('localStorage manually cleared successfully');

			// Also clear popup localStorage if available
			try {
				const extensionUrl = chrome.runtime.getURL('popup/popup.html');
				const tabs = await chrome.tabs.query({ url: extensionUrl });
				for (const tab of tabs) {
					chrome.tabs.sendMessage(tab.id, {
						action: 'clearLocalStorage'
					}).catch(() => {
						// Ignore errors if tab doesn't have message listener
					});
				}
			} catch (error) {
				console.log('Could not clear popup localStorage:', error);
			}
		} catch (error) {
			console.error('Error manually clearing localStorage:', error);
		}
	}
}

// Legacy Queue-based case processing system (kept for backward compatibility)
class CaseProcessingQueue {
	constructor(connection, mode, userId, userName) {
		this.connection = connection;
		this.mode = mode;
		this.userId = userId;
		this.userName = userName;
		this.queue = [];
		this.processing = false;
		this.stats = {
			total: 0,
			processed: 0,
			successful: 0,
			failed: 0,
			assignmentsDetected: 0,
			outsideBusinessWindow: 0, // Track cases filtered out due to time constraints
			startTime: null,
			endTime: null
		};
		this.batchSize = 15; // Optimal batch size for Salesforce API
		this.delayBetweenBatches = 800; // 800ms delay between batches
		this.maxConcurrentBatches = 2; // Process up to 2 batches concurrently
		this.activeBatches = 0;
	}

	// Add cases to the processing queue
	async addCases(cases) {
		this.queue = [...cases];
		this.stats.total = cases.length;
		this.stats.startTime = Date.now();

		// Clean localStorage to remove expired entries
		await this.cleanLocalStorage();
	}

	// Process all cases in the queue
	async processAll() {
		if (this.processing) {
			return;
		}

		this.processing = true;

		try {
			// Process in batches with progress tracking
			while (this.queue.length > 0) {
				// Wait if we've reached max concurrent batches
				while (this.activeBatches >= this.maxConcurrentBatches) {
					await this.delay(100);
				}

				// Take next batch
				const batch = this.queue.splice(0, this.batchSize);
				this.activeBatches++;

				// Process batch asynchronously
				this.processBatch(batch).finally(() => {
					this.activeBatches--;
				});

				// Progress update
				this.updateProgress();

				// Delay before next batch (if not the last batch)
				if (this.queue.length > 0) {
					await this.delay(this.delayBetweenBatches);
				}
			}

			// Wait for all active batches to complete
			while (this.activeBatches > 0) {
				await this.delay(100);
			}

			this.stats.endTime = Date.now();

		} catch (error) {
			// Handle error silently
		} finally {
			this.processing = false;
		}

		return this.stats;
	}

	// Process a single batch of cases
	async processBatch(batch) {
		const batchId = `batch_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
		const batchStartTime = Date.now();

		try {
			// Get case IDs for this batch
			const caseIds = batch.map(c => c.Id);

			// Fetch case history for this batch
			const historyData = await this.fetchCaseHistory(caseIds);

			// Process each case in the batch
			for (const caseRecord of batch) {
				await this.processSingleCase(caseRecord, historyData.get(caseRecord.Id) || []);
			}

		} catch (error) {
			this.stats.failed += batch.length;
		}
	}

	// Fetch case history for a batch of cases
	async fetchCaseHistory(caseIds) {
		try {
			const histQuery = `SELECT CaseId, CreatedById, CreatedDate, Field, NewValue FROM CaseHistory WHERE CaseId IN ('${caseIds.join("','")}') AND (Field='Routing_Status__c' OR Field='Owner') ORDER BY CreatedDate DESC`;

			const histResult = await this.connection.query(histQuery);
			if (!histResult.records) return new Map();

			// Group history by case
			const caseHistory = new Map();
			histResult.records.forEach(record => {
				if (!caseHistory.has(record.CaseId)) {
					caseHistory.set(record.CaseId, []);
				}
				caseHistory.get(record.CaseId).push(record);
			});

			return caseHistory;

		} catch (error) {
			return new Map();
		}
	}

	// Process a single case for assignment detection
	async processSingleCase(caseRecord, history) {
		try {
			// Check if case has been manually assigned or ownership changed
			let assignmentDetected = false;
			let assignmentType = '';
			let assignedTo = '';
			let assignmentDate = null;

			for (const hist of history) {
				if (hist.Field === 'Routing_Status__c' && hist.NewValue && hist.NewValue.startsWith('Manually Assigned')) {
					assignmentDetected = true;
					assignmentType = 'New Case';
					assignedTo = hist.NewValue.replace('Manually Assigned to ', '');
					assignmentDate = hist.CreatedDate;
					break;
				} else if (hist.Field === 'Owner' && hist.NewValue && hist.NewValue !== this.userId) {
					// Check if the new owner is a human-readable name (not a Salesforce ID)
					if (!/^005[\w]{12,15}$/.test(hist.NewValue)) {
						assignmentDetected = true;
						assignmentType = 'New Case';
						assignedTo = hist.NewValue;
						assignmentDate = hist.CreatedDate;
						break;
					}
				}
			}

			if (assignmentDetected) {
				await this.handleAssignmentDetected(caseRecord, assignmentType, assignedTo, assignmentDate);
			}

			this.stats.processed++;
			this.stats.successful++;

		} catch (error) {
			this.stats.failed++;
		}
	}

	// Handle detected assignment
	async handleAssignmentDetected(caseRecord, assignmentType, assignedTo, assignmentDate) {
		try {
			// Log the assignment to Google Sheets
			const cloud = caseRecord.CaseRoutingTaxonomy__r && caseRecord.CaseRoutingTaxonomy__r.Name ?
				caseRecord.CaseRoutingTaxonomy__r.Name.split('-')[0] : '';

			// Send message to popup to trigger trackActionAndCount
			let popupAvailable = false;
			try {
				const extensionUrl = chrome.runtime.getURL('popup/popup.html');
				const tabs = await chrome.tabs.query({ url: extensionUrl });

				if (tabs.length > 0) {
					popupAvailable = true;
					for (const tab of tabs) {
						chrome.tabs.sendMessage(tab.id, {
							action: 'trackActionFromBackground',
							data: {
								createdDate: assignmentDate,
								caseNumber: caseRecord.CaseNumber,
								severity: caseRecord.Severity_Level__c,
								actionType: assignmentType,
								cloud: cloud,
								mode: this.mode,
								userName: this.userName,
								newValue: assignedTo
							}
						}).catch(() => {
							// Ignore errors if tab doesn't have message listener
						});
					}
				}

				this.stats.assignmentsDetected++;

				// Mark case as processed in persistent set
				const existing = persistentCases.get(caseRecord.Id);
				if (existing) {
					persistentCases.set(caseRecord.Id, {
						...existing,
						processed: true,
						processedReason: 'background-assignment-detected',
						lastProcessed: Date.now()
					});
				}

			} catch (error) {
				console.error('Error sending message to popup:', error);
			}

			// Fallback: If popup is not available, log directly to Google Sheets
			if (!popupAvailable) {
				try {
					console.log('Popup not available, logging directly to Google Sheets from background script');
					console.log('About to import trackAction function...');
					const { trackAction } = await import('./utils/api.js');
					console.log('Successfully imported trackAction function');

					console.log('About to call trackAction with params:', {
						assignmentDate,
						caseNumber: caseRecord.CaseNumber,
						severity: caseRecord.Severity_Level__c,
						actionType: assignmentType,
						cloud,
						mode: this.mode,
						userName: this.userName,
						assignedTo
					});

					const result = await trackAction(
						assignmentDate,
						caseRecord.CaseNumber,
						caseRecord.Severity_Level__c,
						assignmentType,
						cloud,
						this.mode,
						this.userName,
						assignedTo,
						() => {
							console.log(`Successfully logged case ${caseRecord.CaseNumber} to Google Sheets from background script`);
							// Mark as sent to prevent duplicates
							this.markAsSentToSheets(`${caseRecord.CaseNumber}_${assignmentDate.split('T')[0]}_${assignmentType}_${assignedTo}`);
						}
					);
					console.log('trackAction completed with result:', result);
				} catch (fallbackError) {
					console.error('Fallback Google Sheets logging failed:', fallbackError);
					console.error('Error stack:', fallbackError.stack);
					// Store in pending tracking data for later processing
					this.storePendingTrackingData({
						createdDate: assignmentDate,
						caseNumber: caseRecord.CaseNumber,
						severity: caseRecord.Severity_Level__c,
						actionType: assignmentType,
						cloud: cloud,
						mode: this.mode,
						userName: this.userName,
						newValue: assignedTo
					});
				}
			} else {
				console.log('Popup is available, message sent to popup for Google Sheets logging');
			}

		} catch (error) {
			// Handle error silently
		}
	}

	// Update and display progress
	updateProgress() {
		// Progress tracking logic (silent)
	}

	// Utility function for delays
	delay(ms) {
		return new Promise(resolve => setTimeout(resolve, ms));
	}

	// Method to store pending tracking data for later processing
	storePendingTrackingData(trackingData) {
		storePendingTrackingData(trackingData);
	}

	// Method to clean localStorage and remove expired entries
	async cleanLocalStorage() {
		try {
			console.log('Cleaning localStorage to remove expired entries...');
			const rawData = localStorage.getItem('sentToSheets');

			if (!rawData) {
				console.log('No localStorage data to clean');
				return;
			}

			try {
				const sentEntries = JSON.parse(rawData);
				if (!Array.isArray(sentEntries)) {
					console.log('Invalid localStorage format, clearing completely');
					localStorage.removeItem('sentToSheets');
					return;
				}

				// Filter out entries older than 7 days to prevent localStorage from growing indefinitely
				const now = new Date();
				const sevenDaysAgo = new Date(now.getTime() - (7 * 24 * 60 * 60 * 1000));

				const filteredEntries = sentEntries.filter(entry => {
					try {
						// Extract date from tracking ID format: CASE_NUMBER_DATE_ACTIONTYPE_ASSIGNEDTO
						const parts = entry.split('_');
						if (parts.length >= 3) {
							const dateStr = parts[1]; // Format: YYYY-MM-DD
							const entryDate = new Date(dateStr);
							return entryDate >= sevenDaysAgo;
						}
						return false; // Remove malformed entries
					} catch (error) {
						console.log(`Error parsing date from entry ${entry}:`, error);
						return false; // Remove entries with invalid dates
					}
				});

				if (filteredEntries.length !== sentEntries.length) {
					const removedCount = sentEntries.length - filteredEntries.length;
					console.log(`Removed ${removedCount} expired entries from localStorage`);
					localStorage.setItem('sentToSheets', JSON.stringify(filteredEntries));
				} else {
					console.log('No expired entries to remove');
				}

			} catch (parseError) {
				console.log('Error parsing localStorage data, clearing completely:', parseError);
				localStorage.removeItem('sentToSheets');
			}

			// Also clean popup localStorage if available
			try {
				const extensionUrl = chrome.runtime.getURL('popup/popup.html');
				const tabs = await chrome.tabs.query({ url: extensionUrl });
				for (const tab of tabs) {
					chrome.tabs.sendMessage(tab.id, {
						action: 'cleanLocalStorage'
					}).catch(() => {
						// Ignore errors if tab doesn't have message listener
					});
				}
			} catch (error) {
				console.log('Could not clean popup localStorage:', error);
			}
		} catch (error) {
			console.error('Error cleaning localStorage:', error);
		}
	}

	// Method to manually clear localStorage (for debugging/testing)
	async clearLocalStorage() {
		try {
			console.log('Manually clearing localStorage...');
			localStorage.removeItem('sentToSheets');
			console.log('localStorage manually cleared successfully');

			// Also clear popup localStorage if available
			try {
				const extensionUrl = chrome.runtime.getURL('popup/popup.html');
				const tabs = await chrome.tabs.query({ url: extensionUrl });
				for (const tab of tabs) {
					chrome.tabs.sendMessage(tab.id, {
						action: 'clearLocalStorage'
					}).catch(() => {
						// Ignore errors if tab doesn't have message listener
					});
				}
			} catch (error) {
				console.log('Could not clear popup localStorage:', error);
			}
		} catch (error) {
			console.error('Error manually clearing localStorage:', error);
		}
	}
}

// Removed: upsertPersistentCase (persistent storage disabled)

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
			}
		}
	} catch (error) {
		// Handle error silently
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
	} catch (error) {
		// Handle error silently
	}
}

chrome.alarms.onAlarm.addListener(async (alarm) => {
	if (alarm.name === 'SigSEV2Triaging') {
		await createSingleExtensionTab();
	}
	if (alarm.name === 'DailyCaseCleanup') {
		await performDailyCaseCleanup();
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
						sendResponse({ success: true, message: 'Persistent storage disabled' });
					} else {
						sendResponse({ success: false, message: 'Invalid cases data' });
					}
				} catch (error) {
					sendResponse({ success: false, message: error.message });
				}
				return true;
			}

			// Sync cases seen in latest query. Do not auto-remove missing; only mark as missing and report back.
			if (request.action === 'syncPersistentCases') {
				try {
					const mode = request.mode;
					if (!mode) {
						sendResponse({ success: false, message: 'Mode is required' });
						return true;
					}
					sendResponse({ success: true, removed: 0, count: 0, removedCases: [], missingCases: [] });
				} catch (error) {
					sendResponse({ success: false, message: error.message });
				}
				return true;
			}

			if (request.action === 'removeCaseFromPersistentSet') {
				try {
					if (request.caseId) {
						sendResponse({ success: true, removed: false, message: 'Persistent storage disabled', count: 0 });
					} else {
						sendResponse({ success: false, message: 'No case ID provided' });
					}
				} catch (error) {
					sendResponse({ success: false, message: error.message });
				}
				return true;
			}

			if (request.action === 'getPersistentCaseCount') {
				try {
					sendResponse({ success: true, count: 0, cases: [] });
				} catch (error) {
					sendResponse({ success: false, message: error.message, count: 0 });
				}
				return true;
			}

			if (request.action === 'getPersistentCaseDetails') {
				try {
					sendResponse({ success: true, cases: [], count: 0 });
				} catch (error) {
					sendResponse({ success: false, message: error.message, cases: [] });
				}
				return true;
			}

			// Handle background case processing for assignment detection
			if (request.action === 'processCasesInBackground') {
				sendResponse({ success: false, message: 'Background case processing is disabled' });
				return true;
			}

			// Handle new background case processing using TrackedChange approach
			if (request.action === 'processCasesInBackgroundNew') {
				sendResponse({ success: false, message: 'Background case processing is disabled' });
				return true;
			}

			// Handle continuous CaseFeed processing
			if (request.action === 'startContinuousCaseFeedProcessing') {
				sendResponse({ success: false, message: 'Background case processing is disabled' });
				return true;
			}

			// Handle manual daily cleanup trigger
			if (request.action === 'triggerDailyCleanup') {
				try {
					performDailyCaseCleanup().then(() => {
						sendResponse({
							success: true,
							message: 'Daily cleanup triggered manually'
						});
					}).catch(error => {
						sendResponse({ success: false, message: error.message });
					});
				} catch (error) {
					sendResponse({ success: false, message: error.message });
				}
				return true;
			}

			// Default response for unhandled messages
			sendResponse({ success: false, message: 'Unknown request type', request: request });
			return true;
		} catch (error) {
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
				}
				tabId = activeTabs[0].id;
			} else {
				// No active tabs, keep the newest and close others
				const sortedTabs = tabs.sort((a, b) => b.id - a.id);
				const tabsToClose = sortedTabs.slice(1).map(t => t.id);
				if (tabsToClose.length > 0) {
					await chrome.tabs.remove(tabsToClose);
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
				}
				tabId = activeTabs[0].id;
			} else {
				// No active tabs, keep the most recently created tab and close others
				const sortedTabs = tabs.sort((a, b) => b.id - a.id);
				const tabsToClose = sortedTabs.slice(1).map(t => t.id);
				await chrome.tabs.remove(tabsToClose);
				tabId = sortedTabs[0].id;
			}
		} else if (tabs.length === 1) {
			// Update our reference to the existing tab
			tabId = tabs[0].id;
		} else if (tabs.length === 0) {
			// No extension tabs exist, clear our reference
			tabId = null;
		}
	} catch (error) {
		// Handle error silently
	}
}

// Run cleanup every 2 minutes
setInterval(performPeriodicCleanup, 2 * 60 * 1000);

// ================================================
// Message Handlers for LocalStorage Cleanup
// ================================================

// Listen for messages from popup and other parts of the extension
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
	try {
		if (request.action === 'cleanupLocalStorage') {
			// Handle localStorage cleanup request (remove items older than 3 days)
			console.log('Background: Cleaning up localStorage items older than 3 days...');
			try {
				const result = cleanupLocalStorageOlderThan3Days();
				if (result.success) {
					console.log(`Background: LocalStorage cleanup completed: ${result.cleanedCount} items removed`);
					sendResponse({
						success: true,
						message: `LocalStorage cleanup completed: ${result.cleanedCount} items removed`,
						cleanedCount: result.cleanedCount,
						totalKeys: result.totalKeys
					});
				} else {
					console.error('Background: LocalStorage cleanup failed:', result.error);
					sendResponse({ success: false, error: result.error });
				}
			} catch (error) {
				console.error('Background: Error during localStorage cleanup:', error);
				sendResponse({ success: false, error: error.message });
			}
			return true;
		} else if (request.action === 'comprehensiveLocalStorageCleanup') {
			// Handle comprehensive localStorage cleanup request
			console.log('Background: Running comprehensive localStorage cleanup...');
			try {
				const result = comprehensiveLocalStorageCleanup();
				if (result.success) {
					console.log(`Background: Comprehensive cleanup completed: ${result.totalCleaned} total items removed`);
					sendResponse({
						success: true,
						message: `Comprehensive cleanup completed: ${result.totalCleaned} total items removed`,
						totalCleaned: result.totalCleaned,
						results: result.results,
						generalCleanup: result.generalCleanup
					});
				} else {
					console.error('Background: Comprehensive cleanup failed:', result.error);
					sendResponse({ success: false, error: result.error });
				}
			} catch (error) {
				console.error('Background: Error during comprehensive localStorage cleanup:', error);
				sendResponse({ success: false, error: error.message });
			}
			return true;
		} else if (request.action === 'getLocalStorageStats') {
			// Handle localStorage stats request
			console.log('Background: Getting localStorage statistics...');
			try {
				const stats = getLocalStorageStats();
				if (stats.error) {
					console.error('Background: Error getting localStorage stats:', stats.error);
					sendResponse({ success: false, error: stats.error });
				} else {
					console.log('Background: LocalStorage stats retrieved successfully');
					sendResponse({
						success: true,
						stats: stats
					});
				}
			} catch (error) {
				console.error('Background: Error getting localStorage stats:', error);
				sendResponse({ success: false, error: error.message });
			}
			return true;
		}
	} catch (error) {
		console.error('Background: Error handling message:', error);
		sendResponse({ success: false, error: error.message });
	}
	return false;
});