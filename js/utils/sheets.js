// Utilities for Google Sheets access and Weekend Roster caching

// Google Sheets read helper
export function googleSheetsGET(rangeA1, callback, onError, spreadsheetId) {
    try {
        if (!spreadsheetId) throw new Error('spreadsheetId required');
        chrome.identity.getAuthToken({ interactive: true }, function (token) {
            if (chrome.runtime.lastError || !token) {
                try { console.error('Auth token error:', chrome.runtime.lastError); } catch { }
                onError && onError(new Error('Auth token error'));
                return;
            }
            fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(rangeA1)}?majorDimension=ROWS`, {
                headers: { 'Authorization': 'Bearer ' + token }
            }).then(r => r.json()).then(data => {
                if (data.error) throw new Error(data.error.message || 'Sheets error');
                callback && callback(data);
            }).catch(err => {
                try { console.error('Sheets read failed:', err); } catch { }
                onError && onError(err);
            });
        });
    } catch (e) {
        onError && onError(e);
    }
}

// Read a single cell value excluding text runs that are strikethrough
export async function getCellValueWithoutStrikethrough(sheetName, rowIndexZero, colLetter, spreadsheetId) {
    return new Promise((resolve, reject) => {
        try {
            if (!spreadsheetId) throw new Error('spreadsheetId required');
            const a1 = `${colLetter}${rowIndexZero + 1}`;
            chrome.identity.getAuthToken({ interactive: true }, function (token) {
                if (chrome.runtime.lastError || !token) {
                    return reject(new Error('Auth token error'));
                }
                const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?includeGridData=true&ranges=${encodeURIComponent(`${sheetName}!${a1}`)}`;
                fetch(url, { headers: { 'Authorization': 'Bearer ' + token } })
                    .then(r => r.json())
                    .then(j => {
                        try {
                            const sheets = j.sheets || [];
                            for (const sh of sheets) {
                                const data = (sh.data || []);
                                for (const block of data) {
                                    const rows = block.rowData || [];
                                    for (const row of rows) {
                                        const vals = row.values || [];
                                        for (const v of vals) {
                                            const text = (v.formattedValue) || (v.effectiveValue && v.effectiveValue.stringValue) || '';
                                            const runs = v.textFormatRuns || [];
                                            if (!runs || runs.length === 0) return resolve(String(text || ''));
                                            const ordered = runs.slice().sort((a, b) => (a.startIndex || 0) - (b.startIndex || 0));
                                            let out = '';
                                            for (let i = 0; i < ordered.length; i++) {
                                                const start = ordered[i].startIndex || 0;
                                                const end = (i + 1 < ordered.length) ? (ordered[i + 1].startIndex || text.length) : text.length;
                                                const seg = text.substring(start, end);
                                                const fmt = ordered[i].format || {};
                                                if (fmt.strikethrough) continue;
                                                out += seg;
                                            }
                                            return resolve(out.trim());
                                        }
                                    }
                                }
                            }
                            resolve('');
                        } catch (e) { reject(e); }
                    })
                    .catch(err => reject(err));
            });
        } catch (e) { reject(e); }
    });
}

// Weekend roster cache helpers (sessionStorage)
export function getWeekendRosterCacheKey(prefix, sheetName, shift, weekendDateStr, mode) {
    return `${prefix}${sheetName}_${shift}_${weekendDateStr}_${mode}`;
}

export function getWeekendRosterFromCache(ttlMs, cacheKey) {
    try {
        const cached = sessionStorage.getItem(cacheKey);
        if (!cached) return null;
        const parsed = JSON.parse(cached);
        const now = Date.now();
        if (!parsed.timestamp || (now - parsed.timestamp) > ttlMs) {
            sessionStorage.removeItem(cacheKey);
            return null;
        }
        return parsed.data;
    } catch {
        sessionStorage.removeItem(cacheKey);
        return null;
    }
}

export function setWeekendRosterCache(cacheKey, data, ttlMs) {
    const cacheEntry = { data, timestamp: Date.now(), ttl: ttlMs };
    try {
        sessionStorage.setItem(cacheKey, JSON.stringify(cacheEntry));
    } catch {
        // best-effort fallback handled by callers via cleanup
    }
}

export function cleanupWeekendRosterCache(prefix, ttlMs) {
    try {
        const keys = Object.keys(sessionStorage);
        const now = Date.now();
        for (const key of keys) {
            if (key.startsWith(prefix)) {
                try {
                    const cached = JSON.parse(sessionStorage.getItem(key));
                    if (!cached.timestamp || (now - cached.timestamp) > ttlMs) {
                        sessionStorage.removeItem(key);
                    }
                } catch {
                    sessionStorage.removeItem(key);
                }
            }
        }
    } catch { }
}

export function getWeekendRosterData(sheetName, shift, weekendDateStr, mode, spreadsheetId, prefix, ttlMs, callback, onError) {
    const cacheKey = getWeekendRosterCacheKey(prefix, sheetName, shift, weekendDateStr, mode);
    const cachedData = getWeekendRosterFromCache(ttlMs, cacheKey);
    if (cachedData) {
        callback(cachedData);
        return;
    }
    const range = `'${sheetName}'!A:X`;
    googleSheetsGET(range, (resp) => {
        try {
            setWeekendRosterCache(cacheKey, resp, ttlMs);
            callback(resp);
        } catch (error) {
            try { console.error('Failed to process weekend roster data:', error); } catch { }
            onError && onError(error);
        }
    }, (error) => {
        try { console.error('Failed to fetch weekend roster data:', error); } catch { }
        onError && onError(error);
    }, spreadsheetId);
}

export function googleSheetsGETWithCache(rangeA1, callback, onError, { useCache = false, sheetName, shift, weekendDateStr, mode, spreadsheetId, prefix = 'weekend_roster_cache_', ttlMs = 60 * 60 * 1000 } = {}) {
    if (useCache && sheetName && shift && weekendDateStr && mode && spreadsheetId) {
        return getWeekendRosterData(sheetName, shift, weekendDateStr, mode, spreadsheetId, prefix, ttlMs, callback, onError);
    }
    return googleSheetsGET(rangeA1, callback, onError, spreadsheetId);
}

export function getWeekendRosterCacheStats(prefix, ttlMs) {
    try {
        const keys = Object.keys(sessionStorage);
        const cacheKeys = keys.filter(k => k.startsWith(prefix));
        const now = Date.now();
        let totalEntries = 0, validEntries = 0, expiredEntries = 0, totalSize = 0;
        for (const key of cacheKeys) {
            totalEntries++;
            const raw = sessionStorage.getItem(key);
            totalSize += raw ? raw.length : 0;
            try {
                const parsed = JSON.parse(raw || '{}');
                if (parsed && parsed.timestamp && (now - parsed.timestamp) <= ttlMs) validEntries++; else expiredEntries++;
            } catch { expiredEntries++; }
        }
        return { totalEntries, validEntries, expiredEntries, totalSizeBytes: totalSize, cacheTTLHours: ttlMs / (1000 * 60 * 60) };
    } catch {
        return null;
    }
}

export function clearWeekendRosterCache(prefix) {
    try {
        const keys = Object.keys(sessionStorage);
        const cacheKeys = keys.filter(key => key.startsWith(prefix));
        let clearedCount = 0;
        for (const key of cacheKeys) { sessionStorage.removeItem(key); clearedCount++; }
        return clearedCount;
    } catch {
        return 0;
    }
}

export function refreshWeekendRosterCache(prefix, sheetName, shift, weekendDateStr, mode) {
    try {
        const cacheKey = getWeekendRosterCacheKey(prefix, sheetName, shift, weekendDateStr, mode);
        sessionStorage.removeItem(cacheKey);
        return true;
    } catch {
        return false;
    }
}

export function startWeekendRosterCacheCleanup(prefix, ttlMs) {
    try {
        setInterval(() => { cleanupWeekendRosterCache(prefix, ttlMs); }, 15 * 60 * 1000);
        if (!window._weekendRosterCacheVisHandlerAdded) {
            document.addEventListener('visibilitychange', () => {
                if (document.visibilityState === 'visible') cleanupWeekendRosterCache(prefix, ttlMs);
            });
            window._weekendRosterCacheVisHandlerAdded = true;
        }
    } catch { }
}


