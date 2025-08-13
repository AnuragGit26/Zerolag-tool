import { timeElapsed } from '../utils/datetime.js';
import { formatDateWithDayOfWeek, isToday, getShiftForDate, getCurrentShift, getPreferredShiftValues, buildPreferredShiftCondition, getGHOTemplate } from './shift.js';
import { showToast } from './toast.js';



export function attachGhoPreviewTemplateCopy(containerId = 'gho-cases-container') {
    const container = document.getElementById(containerId);
    if (!container) return;
    container.addEventListener('click', function (e) {
        if (e.target && e.target.classList.contains('gho-preview-btn')) {
            const templateText = getGHOTemplate();
            navigator.clipboard.writeText(templateText).then(() => {
                showToast('GHO template copied to clipboard!');
            }).catch(err => {
                console.error('Failed to copy GHO template: ', err);
                showToast('Failed to copy template to clipboard');
            });
        }
    });
}
