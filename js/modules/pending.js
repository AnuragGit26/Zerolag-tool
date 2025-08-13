import { formatDateWithDayOfWeek } from './shift.js';

export function buildPendingCardsHtml(pendingCasesDetails) {
    let html = '';
    pendingCasesDetails.forEach(caseDetail => {
        const sevClass = caseDetail.severity.includes('Level 1') ? 'sev1' : caseDetail.severity.includes('Level 2') ? 'sev2' : caseDetail.severity.includes('Level 3') ? 'sev3' : 'sev4';
        const sevShort = caseDetail.severity.includes('Level 1') ? 'SEV1' : caseDetail.severity.includes('Level 2') ? 'SEV2' : caseDetail.severity.includes('Level 3') ? 'SEV3' : 'SEV4';
        const total = Math.max(0, Number(caseDetail.totalRequiredMinutes || 0));
        const elapsed = Math.max(0, Number(caseDetail.minutesSinceCreation || 0));
        const remaining = Math.max(0, total - elapsed);
        const totalDisplay = elapsed + remaining;
        const progressPct = totalDisplay > 0 ? Math.max(0, Math.min(100, Math.round((remaining / totalDisplay) * 100))) : 0;
        const dueNow = remaining <= 0;
        const mvpBadgeTop = caseDetail.isMVP ? '<span class="badge-soft badge-soft--purple">MVP</span>' : '';
        html += `
      <div class="pending-card ${sevClass}" data-created="${new Date(caseDetail.createdDate).toISOString()}" data-total="${totalDisplay}">
        <div class="pending-card-top">
          <div class="pending-id">${caseDetail.caseNumber}</div>
          <span class="severity-badge ${sevClass}">${sevShort}</span>
          ${mvpBadgeTop}
        </div>
        <div class="pending-body">
          <div class="pending-account">${caseDetail.account}</div>
          <div class="pending-meta">Created: ${formatDateWithDayOfWeek(caseDetail.createdDate)} (<span class="js-elapsed">${elapsed}m</span> ago)</div>
        </div>
        <div class="pending-progress">
          <div class="pending-progress-bar" style="width: ${progressPct}%;"></div>
          <div class="pending-progress-label"><span class="js-progress-remaining">${remaining}</span>m / <span class="js-progress-total">${totalDisplay}</span>m</div>
        </div>
        <div class="pending-footer">
          <span class="remaining-badge ${dueNow ? 'due' : ''}">${dueNow ? 'Due now' : `<span class=\"js-remaining\">${remaining}</span>m remaining`}</span>
        </div>
      </div>`;
    });
    return html;
}

export function getPendingSectionHtml({ title, pendingCasesCount, subTitleText, pendingGridHtml }) {
    return `
    <section class="pending-section">
      <div class="pending-header">
        <div class="pending-title-wrap">
          <h4 class="pending-title">${title}</h4>
          <span class="pending-count-badge">${pendingCasesCount}</span>
        </div>
        <div class="pending-subtitle">${subTitleText}</div>
      </div>
      <div class="pending-grid">
        ${pendingGridHtml}
      </div>
    </section>`;
}
