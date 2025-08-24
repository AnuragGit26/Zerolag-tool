import { isToday, getShiftForDate, getCurrentShift } from './shift.js';
import { showToast } from './toast.js';

// Manual Track Case modal logic
export async function showManualTrackModal(ctx) {
    try {
        const modal = document.getElementById('manual-track-modal');
        const input = document.getElementById('manual-track-input');
        const fetchBtn = document.getElementById('manual-track-fetch');
        const loading = document.getElementById('manual-track-loading');
        const details = document.getElementById('manual-track-details');
        const confirmBtn = document.getElementById('manual-track-confirm');
        const cancelBtn = document.getElementById('manual-track-cancel');
        const closeX = document.getElementById('manual-track-close');

        if (!modal || !input || !fetchBtn || !loading || !details || !confirmBtn || !cancelBtn || !closeX) return;

        const resetView = () => {
            loading.style.display = 'none';
            details.style.display = 'none';
            details.innerHTML = '';
            confirmBtn.disabled = true;
        };

        resetView();
        input.value = '';
        modal.style.display = 'flex';
        setTimeout(() => modal.classList.add('modal-show'), 0);
        setTimeout(() => input.focus(), 80);

        const closeModal = () => { modal.classList.remove('modal-show'); setTimeout(() => modal.style.display = 'none', 150); };
        cancelBtn.onclick = closeModal; closeX.onclick = closeModal; modal.onclick = (e) => { if (e.target === modal) closeModal(); };

        const doFetch = async () => {
            const num = (input.value || '').trim();
            if (!num) { showToast('Enter a Case Number'); return; }
            const sessionId = typeof ctx.getSessionId === 'function' ? ctx.getSessionId() : ctx.sessionId;
            if (!window.jsforce || !sessionId) { showToast('Session not ready'); return; }
            try {
                loading.style.display = 'block';
                details.style.display = 'none';
                const conn = new jsforce.Connection({ serverUrl: 'https://orgcs.my.salesforce.com', sessionId, version: '64.0' });
                const q = `SELECT Id, CaseNumber, Severity_Level__c, Subject, CaseRoutingTaxonomy__r.Name, Owner.Name, LastModifiedDate FROM Case WHERE CaseNumber='${num}' LIMIT 1`;
                const res = await conn.query(q);
                loading.style.display = 'none';
                if (!res.records || !res.records.length) { details.style.display = 'block'; details.innerHTML = '<div style="color:#b91c1c;">Case not found</div>'; confirmBtn.disabled = true; return; }
                const c = res.records[0];
                const cloud = (c.CaseRoutingTaxonomy__r && c.CaseRoutingTaxonomy__r.Name) ? c.CaseRoutingTaxonomy__r.Name.split('-')[0] : '';

                // Resolve current user id once for all manual queries
                let userId = null;
                try {
                    const name = typeof ctx.getCurrentUserName === 'function' ? ctx.getCurrentUserName() : ctx.currentUserName;
                    const ures = await conn.query(`SELECT Id FROM User WHERE Name = '${name}' AND IsActive = True AND Username LIKE '%orgcs.com'`);
                    if (ures && ures.records && ures.records.length) userId = ures.records[0].Id;
                } catch { }

                let assignedToForConfirm = 'QB';
                try {
                    if (!userId) { throw new Error('Missing user id'); }
                    const hres = await conn.query(`SELECT Field, NewValue, CreatedDate FROM CaseHistory WHERE CaseId='${c.Id}' AND CreatedById='${userId}' ORDER BY CreatedDate ASC LIMIT 50`);
                    const recs = (hres && hres.records) ? hres.records : [];
                    let ownerVal = '';
                    for (let i = 0; i < recs.length; i++) {
                        const h = recs[i];
                        const newValLower = String(h.NewValue || '').toLowerCase();
                        if (h.Field === 'Routing_Status__c' && (newValLower.includes('transferred') || newValLower.includes('manually assigned'))) {
                            for (let j = i + 1; j < recs.length; j++) {
                                const oh = recs[j];
                                if (oh.Field === 'Owner' && oh.NewValue) { ownerVal = String(oh.NewValue); break; }
                            }
                            break;
                        }
                    }
                    if (ownerVal) {
                        const isUserId = /^005[\w]{12,15}$/.test(ownerVal);
                        const isGroupId = /^00G[\w]{12,15}$/.test(ownerVal);
                        if (isUserId) {
                            try {
                                const u = await conn.query(`SELECT Name FROM User WHERE Id='${ownerVal}' LIMIT 1`);
                                assignedToForConfirm = (u && u.records && u.records.length) ? (u.records[0].Name || ownerVal) : ownerVal;
                            } catch { assignedToForConfirm = ownerVal; }
                        } else if (isGroupId) {
                            try {
                                const g = await conn.query(`SELECT Name FROM Group WHERE Id='${ownerVal}' LIMIT 1`);
                                assignedToForConfirm = (g && g.records && g.records.length) ? (g.records[0].Name || ownerVal) : ownerVal;
                            } catch { assignedToForConfirm = ownerVal; }
                        } else {
                            assignedToForConfirm = ownerVal;
                        }
                    } else if (recs.some(h => h.Field === 'Routing_Status__c' && (String(h.NewValue || '').toLowerCase().includes('transferred') || String(h.NewValue || '').toLowerCase().includes('manually assigned')))) {
                        assignedToForConfirm = (c.Owner && c.Owner.Name) ? c.Owner.Name : 'Case Owner';
                    }
                } catch { }

                let actionTypeForConfirm = 'New Case';
                try {
                    if (userId) {
                        const r = await conn.query(`SELECT Transfer_Reason__c FROM Case_Routing_Log__c WHERE Case__c='${c.Id}' AND CreatedById='${userId}' ORDER BY CreatedDate DESC LIMIT 20`);
                        const recs = (r && r.records) ? r.records : [];
                        if (recs.some(rr => String(rr.Transfer_Reason__c || '').toUpperCase() === 'GHO')) actionTypeForConfirm = 'GHO';
                    }
                } catch { }
                details.innerHTML = `
          <div style="display:grid; grid-template-columns: 1fr 1fr; gap:8px;">
            <div><strong>Case</strong>: ${c.CaseNumber}</div>
            <div><strong>Severity</strong>: ${c.Severity_Level__c || '-'}</div>
            <div style="grid-column: 1 / -1;"><strong>Subject</strong>: ${String(c.Subject || '').replace(/</g, '&lt;')}</div>
            <div><strong>Owner</strong>: ${c.Owner && c.Owner.Name ? c.Owner.Name : '-'}</div>
            <div><strong>Cloud</strong>: ${cloud || '-'}</div>
          </div>
          <div id="manual-track-activity" style="margin-top:12px; padding-top:10px; border-top:1px solid #e5e7eb;">
            <div style="font-weight:700; color:#0f172a; margin-bottom:6px;">Your recent actions</div>
            <div id="manual-track-activity-body" style="display:flex; flex-direction:column; gap:6px;">Loading...</div>
          </div>`;
                details.style.display = 'block';
                confirmBtn.disabled = false;

                confirmBtn.onclick = async () => {
                    try {
                        const mode = typeof ctx.getCurrentMode === 'function' ? ctx.getCurrentMode() : ctx.currentMode;
                        const name = typeof ctx.getCurrentUserName === 'function' ? ctx.getCurrentUserName() : ctx.currentUserName;
                        ctx.trackActionAndCount(new Date(), c.CaseNumber, c.Severity_Level__c || '', actionTypeForConfirm, cloud || '', mode, name, assignedToForConfirm);
                        showToast(`Tracked New Case for ${c.CaseNumber}`);
                        closeModal();
                    } catch (err) {
                        console.warn('Manual confirm track failed:', err);
                        showToast('Failed to track');
                    }
                };

                (async () => {
                    try {
                        const bodyEl = document.getElementById('manual-track-activity-body');
                        if (!bodyEl) return;
                        if (!userId) { bodyEl.textContent = 'Could not resolve your user id.'; return; }

                        const [feedRes, histRes] = await Promise.all([
                            conn.query(`SELECT Body, CreatedDate FROM CaseFeed WHERE Visibility='InternalUsers' AND ParentId='${c.Id}' AND Type='TextPost' AND CreatedById='${userId}' ORDER BY CreatedDate DESC LIMIT 20`),
                            conn.query(`SELECT Field, NewValue, CreatedDate FROM CaseHistory WHERE CaseId='${c.Id}' AND CreatedById='${userId}' AND (Field='Routing_Status__c' OR Field='Owner') ORDER BY CreatedDate DESC LIMIT 20`)
                        ]);
                        const items = [];
                        (feedRes.records || []).forEach(cm => {
                            const body = String(cm.Body || ''); const low = body.toLowerCase();
                            if (low.includes('#ghotriage') || low.includes('#sigqbmention')) items.push({ t: 'comment', when: cm.CreatedDate, text: body });
                        });
                        (histRes.records || []).forEach(h => {
                            const f = h.Field; const val = (typeof h.NewValue === 'string') ? h.NewValue : (h.NewValue && h.NewValue.name) ? h.NewValue.name : '';
                            items.push({ t: 'history', when: h.CreatedDate, text: `${f}: ${val}` });
                        });
                        items.sort((a, b) => new Date(b.when) - new Date(a.when));
                        if (!items.length) { bodyEl.textContent = 'No recent comments or history by you.'; return; }
                        bodyEl.innerHTML = items.slice(0, 10).map(it => {
                            const when = new Date(it.when).toLocaleString();
                            const badge = it.t === 'comment' ? '<span class="badge-soft badge-soft--info" style="margin-right:6px;">Comment</span>' : '<span class="badge-soft" style="background:#eef2ff; color:#3730a3; border:1px solid #e0e7ff; margin-right:6px;">History</span>';
                            const safe = String(it.text || '').replace(/</g, '&lt;');
                            return `<div style="font-size:12px; color:#334155;">${badge}<span style="color:#0f172a; font-weight:600;">${when}</span> — ${safe}</div>`;
                        }).join('');
                    } catch (e) {
                        const bodyEl = document.getElementById('manual-track-activity-body');
                        if (bodyEl) bodyEl.textContent = 'Failed to load your recent actions.';
                    }
                })();
            } catch (e) {
                loading.style.display = 'none';
                details.style.display = 'block';
                details.innerHTML = `<div style="color:#b91c1c;">${String(e.message || e).replace(/</g, '&lt;')}</div>`;
                confirmBtn.disabled = true;
            }
        };

        fetchBtn.onclick = doFetch;
        input.onkeypress = (e) => { if (e.key === 'Enter') doFetch(); };
    } catch (e) { console.warn('showManualTrackModal failed', e); }
}

// History helper: track new case from history
export async function trackNewCaseFromHistory(conn, params) {
    const { caseIds, currentUserId, currentUserName, currentMode, strategy, removeFromPersistent } = params || {};
    try {
        if (!Array.isArray(caseIds) || caseIds.length === 0) return { processed: [] };

        const caseRes = await conn.query(`SELECT Id, CaseNumber, Severity_Level__c, CaseRoutingTaxonomy__r.Name, Owner.Name FROM Case WHERE Id IN ('${caseIds.join("','")}')`);
        const caseMap = new Map((caseRes.records || []).map(r => [r.Id, r]));

        const orderDir = strategy === 'firstManualByUser' ? 'ASC' : 'DESC';
        const histRes = await conn.query(`SELECT CaseId, CreatedById, CreatedDate, Field, NewValue FROM CaseHistory WHERE CaseId IN ('${caseIds.join("','")}') AND CreatedById='${currentUserId}' AND (Field='Routing_Status__c' OR Field='Owner') ORDER BY CreatedDate ${orderDir}`);
        const records = histRes.records || [];

        const byCase = new Map();
        for (const h of records) {
            if (!byCase.has(h.CaseId)) byCase.set(h.CaseId, []);
            byCase.get(h.CaseId).push(h);
        }

        const processed = [];
        for (const caseId of caseIds) {
            const hist = byCase.get(caseId) || [];
            let candidate = null;
            for (const h of hist) {
                const isManualAssign = (h.Field === 'Routing_Status__c' && typeof h.NewValue === 'string' && h.NewValue.startsWith('Manually Assigned'));
                const isOwnerChange = (h.Field === 'Owner');
                if (strategy === 'firstManualByUser') {
                    if (isManualAssign && h.CreatedById === currentUserId) { candidate = h; break; }
                } else {
                    if ((isManualAssign || isOwnerChange) && h.CreatedById === currentUserId) { candidate = h; break; }
                }
            }
            if (!candidate) continue;

            const cRec = caseMap.get(caseId);
            if (!cRec) continue;

            const trackingKey = `tracked_${currentMode}_assignment_${caseId}`;
            if (localStorage.getItem(trackingKey)) continue;

            const cloud = (cRec.CaseRoutingTaxonomy__r && cRec.CaseRoutingTaxonomy__r.Name) ? cRec.CaseRoutingTaxonomy__r.Name.split('-')[0] : '';
            const isSfUserId = (v) => typeof v === 'string' && /^005[\w]{12,15}$/.test(v);
            const scanArr = strategy === 'firstManualByUser' ? hist.slice().reverse() : hist;
            let ownerNameFromHist = '';
            for (const oh of scanArr) {
                if (oh.Field === 'Owner' && oh.NewValue && !isSfUserId(oh.NewValue)) { ownerNameFromHist = oh.NewValue; break; }
            }
            const assignedTo = ownerNameFromHist || ((cRec.Owner && cRec.Owner.Name) ? cRec.Owner.Name : (candidate.NewValue || ''));

            // Note: actual logging handled elsewhere via ctx.trackActionAndCount in callers
            localStorage.setItem(trackingKey, 'true');

            processed.push({ caseId, action: 'New Case (history)', assignedTo });
        }
        return { processed };
    } catch (e) {
        console.warn('trackNewCaseFromHistory failed:', e);
        return { processed: [], error: e.message };
    }
}

// Detect assignments for cases (no logging)
export async function detectAssignmentsForCases(conn, params) {
    const { caseIds } = params || {};
    try {
        if (!Array.isArray(caseIds) || caseIds.length === 0) return { processed: [] };

        const caseRes = await conn.query(`SELECT Id, CaseNumber, Owner.Name, Status, IsClosed FROM Case WHERE Id IN ('${caseIds.join("','")}')`);
        const caseMap = new Map((caseRes.records || []).map(r => [r.Id, r]));

        const sinceIso = new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString();
        const histRes = await conn.query(`SELECT CaseId, CreatedById, CreatedDate, Field, NewValue FROM CaseHistory WHERE CaseId IN ('${caseIds.join("','")}') AND CreatedDate >= ${sinceIso} AND (Field='Routing_Status__c' OR Field='Owner') ORDER BY CreatedDate DESC`);
        const records = histRes.records || [];
        const byCase = new Map();
        for (const h of records) {
            if (!byCase.has(h.CaseId)) byCase.set(h.CaseId, []);
            byCase.get(h.CaseId).push(h);
        }

        const processed = [];
        for (const caseId of caseIds) {
            const hist = byCase.get(caseId) || [];
            let assigned = false;
            let reason = '';

            const cRec = caseMap.get(caseId);
            if (cRec) {
                const ownerName = (cRec.Owner && cRec.Owner.Name) ? String(cRec.Owner.Name) : '';
                const isQueueOwner = /Queue/i.test(ownerName) || ['Kase Changer', 'Working in Org62', 'Data Cloud Queue'].some(q => ownerName.includes(q));
                if (cRec.IsClosed === true) { assigned = true; reason = 'closed'; }
                else if (cRec.Status && cRec.Status !== 'New') { assigned = true; reason = 'status-changed'; }
                else if (!isQueueOwner && ownerName) { assigned = true; reason = 'owner-human'; }
            }

            for (const h of hist) {
                const isManualAssign = (h.Field === 'Routing_Status__c' && typeof h.NewValue === 'string' && h.NewValue.startsWith('Manually Assigned'));
                const isOwnerChange = (h.Field === 'Owner');
                if (isManualAssign || isOwnerChange) { assigned = true; reason = isManualAssign ? 'history-manual' : 'history-owner'; break; }
            }
            if (assigned) {
                processed.push({ caseId, action: 'assignment-detected', reason, owner: (cRec && cRec.Owner && cRec.Owner.Name) || '' });
            }
        }
        return { processed };
    } catch (e) {
        console.warn('detectAssignmentsForCases failed:', e);
        return { processed: [], error: e.message };
    }
}

// Force-process by Case Number
export async function forceProcessCase(ctx, caseNumber) {
    try {
        if (!caseNumber || typeof caseNumber !== 'string') {
            showToast('Provide a valid Case Number');
            return { success: false, message: 'Invalid Case Number' };
        }
        const sessionId = typeof ctx.getSessionId === 'function' ? ctx.getSessionId() : ctx.sessionId;
        if (!sessionId) {
            showToast('Session not ready. Try again after data loads.');
            return { success: false, message: 'Missing SESSION_ID' };
        }

        const conn = new jsforce.Connection({ serverUrl: 'https://orgcs.my.salesforce.com', sessionId });

        let userId = null;
        const name = typeof ctx.getCurrentUserName === 'function' ? ctx.getCurrentUserName() : ctx.currentUserName;
        try {
            const ures = await conn.query(`SELECT Id FROM User WHERE Name = '${name}' AND IsActive = True AND Username LIKE '%orgcs.com'`);
            if (ures.records && ures.records.length > 0) userId = ures.records[0].Id;
        } catch (e) { console.warn('Failed fetching user id:', e); }
        if (!userId) { showToast('Could not resolve current user'); return { success: false, message: 'Missing user id' }; }

        const caseRes = await conn.query(`SELECT Id, CaseNumber, Severity_Level__c, Subject, CaseRoutingTaxonomy__r.Name, Owner.Name, LastModifiedDate FROM Case WHERE CaseNumber='${caseNumber}' LIMIT 5`);
        if (!caseRes.records || caseRes.records.length === 0) { showToast(`Case ${caseNumber} not found`); return { success: false, message: 'Case not found' }; }
        const caseRec = caseRes.records[0];
        const caseId = caseRec.Id;
        const cloud = (caseRec.CaseRoutingTaxonomy__r && caseRec.CaseRoutingTaxonomy__r.Name) ? caseRec.CaseRoutingTaxonomy__r.Name.split('-')[0] : '';

        let actions = [];
        try {
            const res = await trackNewCaseFromHistory(conn, {
                caseIds: [caseId],
                currentUserId: userId,
                currentUserName: name,
                currentMode: typeof ctx.getCurrentMode === 'function' ? ctx.getCurrentMode() : ctx.currentMode,
                strategy: 'latestByUser',
                removeFromPersistent: false
            });
            if (res && res.processed && res.processed.length > 0) actions.push('New Case (history)');
        } catch (e) { console.warn('History helper failed in forceProcessCase:', e); }

        try {
            const cRes = await conn.query(`SELECT ParentId, Body, CreatedById, LastModifiedDate, CreatedDate FROM CaseFeed WHERE Visibility='InternalUsers' AND ParentId='${caseId}' AND Type='TextPost' ORDER BY CreatedDate DESC LIMIT 20`);
            const comments = cRes.records || [];
            for (const c of comments) {
                if (c.Body && c.Body.includes('#SigQBmention') && c.CreatedById === userId) {
                    const trackingKey = `tracked_${caseId}`;
                    if (!localStorage.getItem(trackingKey)) {
                        ctx.trackActionAndCount(c.LastModifiedDate || c.CreatedDate, caseRec.CaseNumber, caseRec.Severity_Level__c, 'New Case', cloud, typeof ctx.getCurrentMode === 'function' ? ctx.getCurrentMode() : ctx.currentMode, name, 'QB');
                        localStorage.setItem(trackingKey, 'true');
                        actions.push('New Case (QB mention)');
                    }
                    break;
                }
            }
            for (const c of comments) {
                if (c.Body && c.Body.includes('#GHOTriage') && c.CreatedById === userId) {
                    const commentDate = c.LastModifiedDate || c.CreatedDate;
                    if (isToday(commentDate) && getShiftForDate(commentDate) === getCurrentShift()) {
                        const ghoKey = `gho_tracked_${caseId}`;
                        if (!localStorage.getItem(ghoKey)) {
                            ctx.trackActionAndCount(commentDate, caseRec.CaseNumber, caseRec.Severity_Level__c, 'GHO', cloud, typeof ctx.getCurrentMode === 'function' ? ctx.getCurrentMode() : ctx.currentMode, name, 'QB');
                            localStorage.setItem(ghoKey, 'true');
                            actions.push('GHO (#GHOTriage)');
                        }
                    }
                    break;
                }
            }
        } catch (e) { console.warn('Comments query failed in forceProcessCase:', e); }

        try { chrome.runtime.sendMessage({ action: 'removeCaseFromPersistentSet', caseId }, () => { }); } catch { }

        if (actions.length === 0) {
            showToast(`No actions tracked for ${caseNumber}`);
            return { success: true, message: 'No actions', caseId, actions };
        } else {
            showToast(`Processed ${actions.length} action(s) for ${caseNumber}`);
            return { success: true, message: 'Processed', caseId, actions };
        }
    } catch (e) {
        console.error('forceProcessCase failed:', e);
        showToast('Force process failed (see console)');
        return { success: false, message: e.message };
    }
}

// Force-process by Case Id
export async function forceProcessCaseById(ctx, caseId) {
    try {
        if (!caseId || typeof caseId !== 'string') { showToast('Provide a valid Case Id'); return { success: false, message: 'Invalid Case Id' }; }
        const sessionId = typeof ctx.getSessionId === 'function' ? ctx.getSessionId() : ctx.sessionId;
        if (!sessionId) { showToast('Session not ready. Try again after data loads.'); return { success: false, message: 'Missing SESSION_ID' }; }

        const conn = new jsforce.Connection({ serverUrl: 'https://orgcs.my.salesforce.com', sessionId });

        let userId = null;
        const name = typeof ctx.getCurrentUserName === 'function' ? ctx.getCurrentUserName() : ctx.currentUserName;
        try {
            const ures = await conn.query(`SELECT Id FROM User WHERE Name = '${name}' AND IsActive = True AND Username LIKE '%orgcs.com'`);
            if (ures.records && ures.records.length > 0) userId = ures.records[0].Id;
        } catch (e) { console.warn('Failed fetching user id:', e); }
        if (!userId) { showToast('Could not resolve current user'); return { success: false, message: 'Missing user id' }; }

        const caseRes = await conn.query(`SELECT Id, CaseNumber, Severity_Level__c, Subject, CaseRoutingTaxonomy__r.Name, Owner.Name, LastModifiedDate FROM Case WHERE Id='${caseId}' LIMIT 5`);
        if (!caseRes.records || caseRes.records.length === 0) { showToast(`Case ${caseId} not found`); return { success: false, message: 'Case not found' }; }
        const caseRec = caseRes.records[0];
        const cloud = (caseRec.CaseRoutingTaxonomy__r && caseRec.CaseRoutingTaxonomy__r.Name) ? caseRec.CaseRoutingTaxonomy__r.Name.split('-')[0] : '';

        let actions = [];
        try {
            const res = await trackNewCaseFromHistory(conn, {
                caseIds: [caseRec.Id],
                currentUserId: userId,
                currentUserName: name,
                currentMode: typeof ctx.getCurrentMode === 'function' ? ctx.getCurrentMode() : ctx.currentMode,
                strategy: 'latestByUser',
                removeFromPersistent: false
            });
            if (res && res.processed && res.processed.length > 0) actions.push('New Case (history)');
        } catch (e) { console.warn('History helper failed in forceProcessCaseById:', e); }

        try {
            const cRes = await conn.query(`SELECT ParentId, Body, CreatedById, LastModifiedDate, CreatedDate FROM CaseFeed WHERE Visibility='InternalUsers' AND ParentId='${caseRec.Id}' AND Type='TextPost' ORDER BY CreatedDate DESC LIMIT 20`);
            const comments = cRes.records || [];
            for (const c of comments) {
                if (c.Body && c.Body.includes('#SigQBmention') && c.CreatedById === userId) {
                    const trackingKey = `tracked_${caseRec.Id}`;
                    if (!localStorage.getItem(trackingKey)) {
                        ctx.trackActionAndCount(c.LastModifiedDate || c.CreatedDate, caseRec.CaseNumber, caseRec.Severity_Level__c, 'New Case', cloud, typeof ctx.getCurrentMode === 'function' ? ctx.getCurrentMode() : ctx.currentMode, name, 'QB');
                        localStorage.setItem(trackingKey, 'true');
                        actions.push('New Case (QB mention)');
                    }
                    break;
                }
            }
            for (const c of comments) {
                if (c.Body && c.Body.includes('#GHOTriage') && c.CreatedById === userId) {
                    const commentDate = c.LastModifiedDate || c.CreatedDate;
                    if (isToday(commentDate) && getShiftForDate(commentDate) === getCurrentShift()) {
                        const ghoKey = `gho_tracked_${caseRec.Id}`;
                        if (!localStorage.getItem(ghoKey)) {
                            ctx.trackActionAndCount(commentDate, caseRec.CaseNumber, caseRec.Severity_Level__c, 'GHO', cloud, typeof ctx.getCurrentMode === 'function' ? ctx.getCurrentMode() : ctx.currentMode, name, 'QB');
                            localStorage.setItem(ghoKey, 'true');
                            actions.push('GHO (#GHOTriage)');
                        }
                    }
                    break;
                }
            }
        } catch (e) { console.warn('Comments query failed in forceProcessCaseById:', e); }

        try { chrome.runtime.sendMessage({ action: 'removeCaseFromPersistentSet', caseId: caseRec.Id }, () => { }); } catch { }

        if (actions.length === 0) {
            showToast(`No actions tracked for ${caseRec.CaseNumber}`);
            return { success: true, message: 'No actions', caseId: caseRec.Id, actions };
        } else {
            showToast(`Processed ${actions.length} action(s) for ${caseRec.CaseNumber}`);
            return { success: true, message: 'Processed', caseId: caseRec.Id, actions };
        }
    } catch (e) {
        console.error('forceProcessCaseById failed:', e);
        showToast('Force process by Id failed (see console)');
        return { success: false, message: e.message };
    }
}

// Entry that accepts number, id, or URL
export async function forceProcess(ctx, input) {
    const s = String(input || '').trim();
    if (!s) { showToast('Provide a Case Number, Id, or URL'); return { success: false, message: 'No input' }; }
    let idMatch = null;
    const m1 = s.match(/Case\/([a-zA-Z0-9]{15,18})/);
    const m2 = s.match(/(500[\w]{12,15})/);
    if (m1 && m1[1]) idMatch = m1[1];
    else if (m2 && m2[1]) idMatch = m2[1];
    if (idMatch) return await forceProcessCaseById(ctx, idMatch);
    return await forceProcessCase(ctx, s);
}


