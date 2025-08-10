const SPREADSHEET_ID = '1BKxQLGFrczjhcx9rEt-jXGvlcCPQblwBhFJjoiDD7TI';


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
