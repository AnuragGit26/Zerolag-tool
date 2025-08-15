import { showToast } from './toast.js';
const STORAGE_KEY = 'premierRosterCounts';
const OVERRIDES_KEY = 'premierRosterOverridesV1';

function todayKeyFromDateStr(dateStr) { return String(dateStr || '').trim(); }
function getNowIST() { return new Date(new Date().toLocaleString('en-US', { timeZone: 'Asia/Kolkata' })); }
function prevDateStr(dateStr) {
    try {
        const [mStr, dStr, yStr] = String(dateStr).split('/');
        const m = parseInt(mStr, 10), d = parseInt(dStr, 10), y = parseInt(yStr, 10);
        const dt = new Date(y, m - 1, d); dt.setDate(dt.getDate() - 1);
        const outM = dt.getMonth() + 1, outD = String(dt.getDate()).padStart(2, '0'), outY = dt.getFullYear();
        return `${outM}/${outD}/${outY}`;
    } catch { return todayKeyFromDateStr(dateStr); }
}
function computeCycleKey(dateStr, shift) {
    const base = todayKeyFromDateStr(dateStr);
    if (String(shift).toUpperCase() === 'AMER') {
        const ist = getNowIST();
        const day = ist.getDay(), h = ist.getHours(), m = ist.getMinutes();
        const todayStr = `${ist.getMonth() + 1}/${String(ist.getDate()).padStart(2, '0')}/${ist.getFullYear()}`;
        if (day === 0 && base === todayStr && (h < 20 || (h === 20 && m < 30))) return prevDateStr(base);
    }
    return base;
}

function readStore() {
    try {
        const parsed = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
        if (parsed && parsed.lastDateKey && !parsed.lastByShift) { parsed.lastByShift = { APAC: parsed.lastDateKey, EMEA: parsed.lastDateKey, AMER: parsed.lastDateKey }; delete parsed.lastDateKey; }
        if (!parsed || typeof parsed !== 'object') return { lastByShift: {}, data: {} };
        parsed.lastByShift = parsed.lastByShift || {}; parsed.data = parsed.data || {};
        return parsed;
    } catch { return { lastByShift: {}, data: {} }; }
}

function writeStore(obj) { try { localStorage.setItem(STORAGE_KEY, JSON.stringify(obj)); } catch { } }
function purgeShiftData(store, shift) {
    const out = {}; const target = String(shift || '').toUpperCase();
    for (const k of Object.keys(store.data || {})) {
        const parts = k.split('|');
        if (parts.length >= 3 && String(parts[1]).toUpperCase() === target) continue;
        out[k] = store.data[k];
    }
    store.data = out;
}

export function initPremierCounters(dateStr, shift) {
    const store = readStore();
    const sh = String(shift || '').toUpperCase();
    if (!sh) {
        const k = todayKeyFromDateStr(dateStr);
        const legacyLast = store.lastByShift && (store.lastByShift.APAC || store.lastByShift.EMEA || store.lastByShift.AMER);
        if (!legacyLast || legacyLast !== k) writeStore({ lastByShift: { APAC: k, EMEA: k, AMER: k }, data: {} });
        return;
    }
    const cycleKey = computeCycleKey(dateStr, sh);
    const lastForShift = store.lastByShift[sh];
    if (!lastForShift || lastForShift !== cycleKey) { purgeShiftData(store, sh); store.lastByShift[sh] = cycleKey; writeStore(store); }
}

export function resetPremierCountersAll(dateStr, shift) {
    const store = readStore();
    const sh = String(shift || '').toUpperCase();
    if (sh) { purgeShiftData(store, sh); store.lastByShift[sh] = computeCycleKey(dateStr, sh); writeStore(store); }
    else { const k = todayKeyFromDateStr(dateStr); writeStore({ lastByShift: { APAC: k, EMEA: k, AMER: k }, data: {} }); }
}

function scopeKey(dateStr, shift, blockId) { const key = computeCycleKey(dateStr, shift); return `${key}|${shift}|${blockId}`; }

function getCounts(dateStr, shift, blockId) { const s = readStore(); const sk = scopeKey(dateStr, shift, blockId); return (s.data && s.data[sk]) ? s.data[sk] : {}; }

function setCount(dateStr, shift, blockId, name, value) { const s = readStore(); const sk = scopeKey(dateStr, shift, blockId); if (!s.data) s.data = {}; if (!s.data[sk]) s.data[sk] = {}; s.data[sk][name] = value; writeStore(s); }

// --- Overrides store (per shift, per block) for edited TE names and emails ---
function readOverridesStore() {
    try {
        const parsed = JSON.parse(localStorage.getItem(OVERRIDES_KEY) || '{}');
        return parsed && typeof parsed === 'object' ? parsed : { data: {} };
    } catch { return { data: {} }; }
}
function writeOverridesStore(obj) { try { localStorage.setItem(OVERRIDES_KEY, JSON.stringify(obj)); } catch { } }
function lcKey(s) { return String(s || '').trim().toLowerCase(); }
function getOverrides(dateStr, shift, blockId) {
    const s = readOverridesStore();
    const sk = scopeKey(dateStr, shift, blockId);
    return (s.data && s.data[sk]) ? s.data[sk] : {};
}
function setOverrides(dateStr, shift, blockId, map) {
    const s = readOverridesStore();
    const sk = scopeKey(dateStr, shift, blockId);
    if (!s.data) s.data = {};
    s.data[sk] = map || {};
    writeOverridesStore(s);
}

// Exported: set a single override; migrates counts from old name to new name for this scope
export function setPremierOverride(dateStr, shift, blockId, originalName, newName, email) {
    try {
        const o = getOverrides(dateStr, shift, blockId);
        const origKey = lcKey(originalName);
        const newRec = { name: String(newName || originalName).trim(), email: String(email || '').trim() || undefined };
        o[origKey] = newRec;
        // Also allow lookup by new name key to be idempotent
        o[lcKey(newRec.name)] = newRec;
        setOverrides(dateStr, shift, blockId, o);

        // Migrate counts if needed
        const counts = getCounts(dateStr, shift, blockId);
        if (counts && Object.prototype.hasOwnProperty.call(counts, originalName)) {
            const val = Number(counts[originalName] || 0);
            if (!Object.prototype.hasOwnProperty.call(counts, newRec.name)) {
                // move
                delete counts[originalName];
                counts[newRec.name] = val;
                const s = readStore();
                const sk = scopeKey(dateStr, shift, blockId);
                if (!s.data) s.data = {}; if (!s.data[sk]) s.data[sk] = {};
                s.data[sk] = counts;
                writeStore(s);
            } else if (originalName !== newRec.name) {
                // drop old key
                delete counts[originalName];
                const s = readStore();
                const sk = scopeKey(dateStr, shift, blockId);
                if (!s.data) s.data = {}; if (!s.data[sk]) s.data[sk] = {};
                s.data[sk] = counts;
                writeStore(s);
            }
        }
        showToast('Saved edit');
        return true;
    } catch { return false; }
}

// Apply overrides to names array and merge email map
function applyOverridesToNames(namesArray, { dateStr, shift, blockId, emailMap }) {
    const overrides = getOverrides(dateStr, shift, blockId) || {};
    const outNames = [];
    const mergedEmail = { ...(emailMap || {}) };
    for (const name of (namesArray || [])) {
        const key = lcKey(name);
        const rec = overrides[key] || null;
        if (rec && rec.name) {
            outNames.push(rec.name);
            if (rec.email) {
                mergedEmail[lcKey(rec.name)] = rec.email;
                mergedEmail[rec.name] = rec.email;
            }
        } else {
            outNames.push(name);
        }
    }
    return { names: outNames, emailMap: mergedEmail };
}

export function parseRosterNames(raw) {
    if (!raw) return [];
    const s = String(raw);
    let outStr = '';
    let depth = 0;
    for (let i = 0; i < s.length; i++) {
        const c = s[i];
        if (c === '\r') continue;
        if (c === '(') depth++;
        if (c === ')' && depth > 0) depth--;
        if (c === '\n' && depth > 0) {
            outStr += ' & ';
            i++;
            while (i < s.length && /\s/.test(s[i])) i++;
            const save = i;
            let j = i;
            while (j < s.length && /[0-9]/.test(s[j])) j++;
            if (j < s.length && s[j] === '.') { j++; while (j < s.length && s[j] === ' ') j++; i = j - 1; } else { i = save - 1; }
            continue;
        }
        outStr += c;
    }
    const tokens = [];
    let buf = '';
    depth = 0;
    for (let i = 0; i < outStr.length; i++) {
        const c = outStr[i];
        if (c === '(') depth++;
        else if (c === ')' && depth > 0) depth--;
        if ((c === '\n' || c === ',' || c === '&') && depth === 0) { if (buf.trim()) tokens.push(buf.trim()); buf = ''; continue; }
        buf += c;
    }
    if (buf.trim()) tokens.push(buf.trim());
    const cleaned = tokens.map(t => t.replace(/^\s*\d+\.\s*/, '').trim()).filter(Boolean);
    const seen = new Set();
    const out = [];
    for (const p of cleaned) { if (!seen.has(p)) { seen.add(p); out.push(p); } }
    return out;
}

export function renderPremierCounters(containerEl, namesArray, { dateStr, shift, blockId, emailMap }) {
    if (!containerEl) return;
    const counts = getCounts(dateStr, shift, blockId);
    // Apply per-shift overrides (renamed entries, email fixes)
    const applied = applyOverridesToNames(namesArray, { dateStr, shift, blockId, emailMap });
    const finalNames = applied.names;
    const finalEmailMap = applied.emailMap;

    const rows = finalNames.map((name, idx) => {
        const val = Number(counts[name] || 0), rowId = `${blockId}-row-${idx}`;
        const lc = name && name.toLowerCase();
        const email = finalEmailMap && (finalEmailMap[lc] || finalEmailMap[name] || finalEmailMap[(lc || '')]);
        return `
            <div class="pc-row" id="${rowId}" data-name="${name.replace(/"/g, '&quot;')}" style="display:flex; align-items:center; justify-content:space-between; gap:12px; padding:10px 12px; border:1px solid #e5e7eb; border-radius:10px; background:#ffffff;">
                <div style="display:flex; align-items:center; gap:10px; min-width:0;">
                    <div style="width:8px; height:8px; border-radius:50%; background:#0ea5e9;"></div>
                    <div title="${name}" style="font-size:13px; color:#0f172a; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${name}</div>
                    ${email ? `<button class=\"pc-mail\" data-email=\"${email.replace(/&/g, '&amp;').replace(/"/g, '&quot;')}\" title=\"Copy email\" style=\"margin-left:6px; display:inline-flex; align-items:center; justify-content:center; width:22px; height:22px; border:1px solid #cbd5e1; background:#f8fafc; color:#0f172a; border-radius:6px; cursor:pointer;\">✉️</button>` : ''}
                </div>
                <div style="display:flex; align-items:center; gap:8px;">
                    <button class="pc-dec" aria-label="Decrease" style="border:1px solid #cbd5e1; background:#f8fafc; color:#0f172a; width:28px; height:28px; border-radius:8px; cursor:pointer;">-</button>
                    <div class="pc-val" style="min-width:22px; text-align:center; font-weight:600; color:#0f172a;">${val}</div>
                    <button class="pc-inc" aria-label="Increase" style="border:1px solid #cbd5e1; background:#f8fafc; color:#0f172a; width:28px; height:28px; border-radius:8px; cursor:pointer;">+</button>
                </div>
            </div>`;
    }).join('');
    containerEl.innerHTML = `<div class="pc-list" style="display:flex; flex-direction:column; gap:8px;">${rows || '<div style="color:#64748b; font-size:13px;">No names to display.</div>'}</div>`;
    containerEl.querySelectorAll('.pc-row').forEach(row => {
        const name = row.getAttribute('data-name');
        const valEl = row.querySelector('.pc-val');
        const dec = row.querySelector('.pc-dec');
        const inc = row.querySelector('.pc-inc');
        const getVal = () => Number(valEl.textContent || '0') || 0;
        const setVal = (n) => { valEl.textContent = String(n); setCount(dateStr, shift, blockId, name, n); };
        if (dec) dec.addEventListener('click', (e) => { e.stopPropagation(); const v = Math.max(0, getVal() - 1); setVal(v); });
        if (inc) inc.addEventListener('click', (e) => { e.stopPropagation(); const v = getVal() + 1; setVal(v); });
        const mailBtn = row.querySelector('.pc-mail');
        if (mailBtn) mailBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            const em = mailBtn.getAttribute('data-email');
            if (em) {
                navigator.clipboard.writeText(em)
                    .then(() => showToast('Email copied'))
                    .catch(() => showToast('Copy failed', 'error'));
            }
        });
    });
}
