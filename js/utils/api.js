const SPREADSHEET_ID = '1BKxQLGFrczjhcx9rEt-jXGvlcCPQblwBhFJjoiDD7TI';

let __usageLogsCache = { rows: null, fetchedAt: 0 };
const USAGE_CACHE_TTL = 5 * 60 * 1000;
let __usageFetchInFlight = false;


function getAuthToken(callback) {
    console.log('getAuthToken called, requesting token...');
    chrome.identity.getAuthToken({ interactive: true }, function (token) {
        console.log('getAuthToken callback executed, token received:', token ? 'Yes' : 'No');
        if (chrome.runtime.lastError) {
            const msg = chrome.runtime.lastError && chrome.runtime.lastError.message || '';
            console.error('Auth error (first attempt):', msg);

            console.log('Clearing cached auth tokens and retrying...');
            chrome.identity.clearAllCachedAuthTokens(() => {
                console.log('Cached tokens cleared, requesting new token...');
                chrome.identity.getAuthToken({ interactive: true }, function (token2) {
                    if (chrome.runtime.lastError) {
                        console.error('Auth error (after cache clear):', chrome.runtime.lastError.message);
                        if (/disabled/i.test(chrome.runtime.lastError.message)) {
                            console.error('Hint: This often means your Google Workspace admin has disabled this service/app or the API is disabled in the OAuth client\'s GCP project. See extension README or ask admin to allow Google Sheets API and this app.');
                        }
                        console.log('No token available after retry');
                        return;
                    }
                    console.log('Token received after retry:', token2 ? 'Yes' : 'No');
                    callback(token2);
                });
            });
            return;
        }
        console.log('Token received on first attempt, calling callback');
        callback(token);
    });
}

export function trackAction(dateofAction, caseNumber, severity, actionType, cloud, currentMode, currentUserName, assignedTo = '', onSuccessCallback = null) {
    console.log(`trackAction called for case ${caseNumber} with callback: ${onSuccessCallback ? 'Yes' : 'No'}`);
    console.log('trackAction params:', { dateofAction, caseNumber, severity, actionType, cloud, currentMode, currentUserName, assignedTo });

    getAuthToken(function (token) {
        console.log('getAuthToken callback executed with token:', token ? 'Yes' : 'No');
        if (!token) {
            console.error('No OAuth token available; skipping Sheets update');
            console.log('No success callback executed due to missing OAuth token');
            return;
        }

        const sheetName = currentMode === 'premier' ? 'premier' : 'signature';

        if (!assignedTo) {
            assignedTo = "--";
        }

        // Handle dateofAction - convert to Date object if it's a string, or use current date as fallback
        let actionDate;
        if (dateofAction) {
            actionDate = dateofAction instanceof Date ? dateofAction : new Date(dateofAction);
        } else {
            actionDate = new Date();
        }

        const pstDate = actionDate.toLocaleDateString('en-US', {
            timeZone: 'America/Los_Angeles',
            year: 'numeric',
            month: '2-digit',
            day: '2-digit'
        });

        const istTime = actionDate.toLocaleTimeString('en-US', {
            timeZone: 'Asia/Kolkata',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit',
            hour12: true
        });

        // Updated column order: Date, Type, Time of Action, Case Number, Engineer Name, Severity, Cloud, Assigned To
        const values = [
            [pstDate, actionType, istTime, caseNumber, currentUserName, severity, cloud, assignedTo]
        ];

        const body = { values };

        const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${sheetName}:append?valueInputOption=USER_ENTERED`;

        const send = (bearer, attempt = 1) => {
            console.log(`Sending request to Google Sheets API for case ${caseNumber}, attempt ${attempt}`);
            return fetch(url, {
                method: 'POST',
                headers: {
                    'Authorization': 'Bearer ' + bearer,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(body)
            }).then(async response => {
                if (!response.ok) {
                    let errText;
                    try { errText = await response.text(); } catch (e) { errText = String(e); }
                    console.error(`Sheets API error (status ${response.status}):`, errText);
                    console.log('No success callback executed due to API error');
                    if ((response.status === 401 || response.status === 403) && attempt === 1) {
                        chrome.identity.clearAllCachedAuthTokens(() => {
                            chrome.identity.getAuthToken({ interactive: true }, function (token3) {
                                if (chrome.runtime.lastError) {
                                    console.error('Re-auth failed after API error:', chrome.runtime.lastError.message);
                                    return;
                                }
                                send(token3, attempt + 1);
                            });
                        });
                    }
                    return;
                }
                console.log(`Google Sheets API response status: ${response.status}, ok: ${response.ok}`);
                return response.json();
            })
                .then(data => {
                    console.log('Sheet updated:', data);
                    if (onSuccessCallback && typeof onSuccessCallback === 'function') {
                        try {
                            console.log('Executing success callback for case tracking...');
                            console.log('Success callback function:', onSuccessCallback);
                            onSuccessCallback();
                            console.log('Success callback executed successfully');
                        } catch (callbackError) {
                            console.error('Error in success callback:', callbackError);
                            console.error('Callback error stack:', callbackError.stack);
                        }
                    } else {
                        console.log('No success callback provided for case tracking');
                    }
                })
                .catch(error => {
                    console.error('Error parsing response or executing callback:', error);
                    console.log('No success callback executed due to response parsing error');
                });
        };

        send(token);
    });
}

// Logs daily usage metrics (total active minutes & last active timestamps) to the 'Logs' sheet
// Params: { datePSTString: 'MM/DD/YYYY', userName: string, totalActiveMinutes: number, lastActiveISO: string }
export function logUsageDaily({ dateLocalString, userName, totalActiveMinutes, lastActiveISO, actionedCases = 0 }) {
    try {
        if (!dateLocalString || !userName) return;
        const normalizedName = String(userName).trim().replace(/\s+/g, ' ');

        if (typeof totalActiveMinutes !== 'number' || totalActiveMinutes < 0) return;

        const guardKey = `usage_logged_${dateLocalString}_${normalizedName}`;
        const lockKey = `usage_lock_${dateLocalString}_${normalizedName}`;
        let lastLogged = null;
        try { lastLogged = JSON.parse(localStorage.getItem(guardKey) || 'null'); } catch { }
        if (lastLogged && typeof lastLogged.lastMinutes === 'number') {
            if (totalActiveMinutes < lastLogged.lastMinutes) return;
            if (totalActiveMinutes - lastLogged.lastMinutes < 2) return;
        }

        getAuthToken(async function (token) {
            if (!token) return;
            const sheetName = 'Logs';
            const lastActive = lastActiveISO ? new Date(lastActiveISO) : new Date();
            const lastActivePST = lastActive.toLocaleTimeString('en-US', { timeZone: 'America/Los_Angeles', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true });
            const lastActiveIST = lastActive.toLocaleTimeString('en-US', { timeZone: 'Asia/Kolkata', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true });
            const rowValues = [dateLocalString, normalizedName, totalActiveMinutes, lastActivePST, lastActiveIST, actionedCases];

            const doUpdate = (rowNumber) => {
                const range = `${sheetName}!A${rowNumber}:F${rowNumber}`;
                return fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`, {
                    method: 'PUT',
                    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
                    body: JSON.stringify({ values: [rowValues] })
                }).then(r => { if (!r.ok) return r.text().then(t => { throw new Error(t); }); return r.json(); })
                    .then(d => console.log('Usage row updated', d))
                    .catch(e => console.error('Usage update error', e));
            };

            const appendNew = () => {
                return fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${sheetName}:append?valueInputOption=USER_ENTERED`, {
                    method: 'POST',
                    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
                    body: JSON.stringify({ values: [rowValues] })
                }).then(r => { if (!r.ok) return r.text().then(t => { throw new Error(t); }); return r.json(); })
                    .then(d => console.log('Usage row appended', d))
                    .catch(e => console.error('Usage append error', e));
            };

            const fetchRowsFresh = async () => {
                try {
                    const resp = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${sheetName}!A:F`, { headers: { 'Authorization': 'Bearer ' + token } });
                    if (!resp.ok) { console.warn('Usage fetch rows failed', resp.status); return []; }
                    const data = await resp.json();
                    __usageLogsCache.rows = data.values || [];
                    __usageLogsCache.fetchedAt = Date.now();
                    return __usageLogsCache.rows;
                } catch (e) {
                    console.warn('Usage rows fetch error', e);
                    return [];
                }
            };

            // Acquire local lock to avoid concurrent writes
            const nowMs = Date.now();
            let existingLock = null;
            try { existingLock = JSON.parse(localStorage.getItem(lockKey) || 'null'); } catch { existingLock = null; }
            if (existingLock && (nowMs - (existingLock.at || 0)) < 15000) {
                return;
            }
            try { localStorage.setItem(lockKey, JSON.stringify({ at: nowMs })); } catch { /* noop */ }

            // Fetch latest rows to decide update vs append
            let rows = await fetchRowsFresh();
            rows = rows && rows.length ? rows : (__usageLogsCache.rows || []);
            let startIndex = 0;
            if (rows.length > 0) {
                const first = rows[0].map(v => v.toLowerCase());
                if (first.includes('date') && first.includes('name')) startIndex = 1;
            }
            let targetRowNumber = null;
            let duplicateRowNumbers = [];
            for (let i = startIndex; i < rows.length; i++) {
                const r = rows[i];
                if (r[0] === dateLocalString && String(r[1]).trim().replace(/\s+/g, ' ') === normalizedName) {
                    if (targetRowNumber === null) targetRowNumber = i + 1; else duplicateRowNumbers.push(i + 1);
                }
            }

            if (targetRowNumber) {
                await doUpdate(targetRowNumber);
                if (__usageLogsCache.rows) __usageLogsCache.rows[targetRowNumber - 1] = rowValues;
            } else {
                // Double-check before append with randomized backoff
                await new Promise(r => setTimeout(r, Math.floor(100 + Math.random() * 300)));
                rows = await fetchRowsFresh();
                startIndex = 0;
                if (rows.length > 0) {
                    const first = rows[0].map(v => v.toLowerCase());
                    if (first.includes('date') && first.includes('name')) startIndex = 1;
                }
                for (let i = startIndex; i < rows.length; i++) {
                    const r = rows[i];
                    if (r[0] === dateLocalString && String(r[1]).trim().replace(/\s+/g, ' ') === normalizedName) {
                        targetRowNumber = i + 1;
                        break;
                    }
                }
                if (targetRowNumber) {
                    await doUpdate(targetRowNumber);
                    if (__usageLogsCache.rows) __usageLogsCache.rows[targetRowNumber - 1] = rowValues;
                } else {
                    await appendNew();
                    __usageLogsCache.fetchedAt = 0;
                }
            }

            localStorage.setItem(guardKey, JSON.stringify({ lastMinutes: totalActiveMinutes, loggedAt: Date.now() }));
            try { localStorage.removeItem(lockKey); } catch { /* noop */ }
        });
    } catch (e) {
        console.error('logUsageDaily upsert error', e);
    }
}
