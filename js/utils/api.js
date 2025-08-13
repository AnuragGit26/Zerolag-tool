const SPREADSHEET_ID = '1BKxQLGFrczjhcx9rEt-jXGvlcCPQblwBhFJjoiDD7TI';

let __usageLogsCache = { rows: null, fetchedAt: 0 };
const USAGE_CACHE_TTL = 5 * 60 * 1000;
let __usageFetchInFlight = false;


function getAuthToken(callback) {
    chrome.identity.getAuthToken({ interactive: true }, function (token) {
        if (chrome.runtime.lastError) {
            const msg = chrome.runtime.lastError && chrome.runtime.lastError.message || '';
            console.error('Auth error (first attempt):', msg);

            chrome.identity.clearAllCachedAuthTokens(() => {
                chrome.identity.getAuthToken({ interactive: true }, function (token2) {
                    if (chrome.runtime.lastError) {
                        console.error('Auth error (after cache clear):', chrome.runtime.lastError.message);
                        if (/disabled/i.test(chrome.runtime.lastError.message)) {
                            console.error('Hint: This often means your Google Workspace admin has disabled this service/app or the API is disabled in the OAuth client\'s GCP project. See extension README or ask admin to allow Google Sheets API and this app.');
                        }
                        return;
                    }
                    callback(token2);
                });
            });
            return;
        }
        callback(token);
    });
}

export function trackAction(dateofAction, caseNumber, severity, actionType, cloud, currentMode, currentUserName, assignedTo = '') {
    getAuthToken(function (token) {
        if (!token) {
            console.error('No OAuth token available; skipping Sheets update');
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

        const send = (bearer, attempt = 1) => fetch(url, {
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
            return response.json();
        })
            .then(data => {
                if (data) {
                    console.log('Sheet updated:', data);
                }
            })
            .catch(error => {
                console.error('Network error updating sheet:', error);
            });

        send(token);
    });
}

// Logs daily usage metrics (total active minutes & last active timestamps) to the 'Logs' sheet
// Params: { datePSTString: 'MM/DD/YYYY', userName: string, totalActiveMinutes: number, lastActiveISO: string }
// Debounce / duplicate control should be handled by caller (we also add a lightweight localStorage guard)
export function logUsageDaily({ dateLocalString, userName, totalActiveMinutes, lastActiveISO, actionedCases = 0 }) {
    try {
        if (!dateLocalString || !userName) return;
        if (typeof totalActiveMinutes !== 'number' || totalActiveMinutes < 0) return;

        const guardKey = `usage_logged_${dateLocalString}_${userName}`;
        let lastLogged = null;
        try { lastLogged = JSON.parse(localStorage.getItem(guardKey) || 'null'); } catch { }
        if (lastLogged && typeof lastLogged.lastMinutes === 'number') {
            if (totalActiveMinutes < lastLogged.lastMinutes) return; // don't regress
            if (totalActiveMinutes - lastLogged.lastMinutes < 2) return; // throttle updates (<2 min delta)
        }

        getAuthToken(async function (token) {
            if (!token) return;
            const sheetName = 'Logs';
            const lastActive = lastActiveISO ? new Date(lastActiveISO) : new Date();
            const lastActivePST = lastActive.toLocaleTimeString('en-US', { timeZone: 'America/Los_Angeles', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true });
            const lastActiveIST = lastActive.toLocaleTimeString('en-US', { timeZone: 'Asia/Kolkata', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: true });
            // Required columns: Date, Name, Active Time (Minutes), Last Active Time(PST), Last Active Time(IST), Actioned Cases
            const rowValues = [dateLocalString, userName, totalActiveMinutes, lastActivePST, lastActiveIST, actionedCases];

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

            const ensureCache = async () => {
                const now = Date.now();
                if (__usageLogsCache.rows && now - __usageLogsCache.fetchedAt < USAGE_CACHE_TTL) return;
                if (__usageFetchInFlight) return;
                __usageFetchInFlight = true;
                try {
                    const resp = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${sheetName}!A:F`, { headers: { 'Authorization': 'Bearer ' + token } });
                    if (!resp.ok) { console.warn('Usage fetch rows failed', resp.status); return; }
                    const data = await resp.json();
                    __usageLogsCache.rows = data.values || [];
                    __usageLogsCache.fetchedAt = Date.now();
                } catch (e) { console.warn('Usage rows fetch error', e); }
                finally { __usageFetchInFlight = false; }
            };

            await ensureCache();
            let rows = __usageLogsCache.rows || [];
            let startIndex = 0;
            if (rows.length > 0) {
                const first = rows[0].map(v => v.toLowerCase());
                if (first.includes('date') && first.includes('name')) startIndex = 1;
            }
            let targetRowNumber = null;
            for (let i = startIndex; i < rows.length; i++) {
                const r = rows[i];
                if (r[0] === dateLocalString && r[1] === userName) { targetRowNumber = i + 1; break; }
            }

            if (targetRowNumber) {
                await doUpdate(targetRowNumber);
                if (__usageLogsCache.rows) __usageLogsCache.rows[targetRowNumber - 1] = rowValues;
            } else {
                await appendNew();
                __usageLogsCache.fetchedAt = 0; // invalidate
            }

            localStorage.setItem(guardKey, JSON.stringify({ lastMinutes: totalActiveMinutes, loggedAt: Date.now() }));
        });
    } catch (e) {
        console.error('logUsageDaily upsert error', e);
    }
}
