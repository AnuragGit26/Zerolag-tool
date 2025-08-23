import { isCurrentlyWeekend } from '../utils/datetime.js';

export function formatDateWithDayOfWeek(date) {
    const dateObj = new Date(date);
    const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const dayOfWeek = dayNames[dateObj.getDay()];
    return `${dayOfWeek}, ${dateObj.toLocaleString()}`;
}

export function getShiftForDate(date) {
    const dateObj = new Date(date);
    const hours = dateObj.getHours();
    const minutes = dateObj.getMinutes();
    const totalMinutes = hours * 60 + minutes;
    const apacStart = 5 * 60 + 30;
    const emeaStart = 12 * 60 + 30;
    const apacEnd = 12 * 60 + 30;
    const emeaEnd = 20 * 60;
    if (totalMinutes >= apacStart && totalMinutes < apacEnd) return 'APAC';
    if (totalMinutes >= emeaStart && totalMinutes < emeaEnd) return 'EMEA';
    return 'AMER';
}

export function isToday(date) {
    const today = new Date();
    const checkDate = new Date(date);
    return today.toDateString() === checkDate.toDateString();
}

export function getCurrentShift() {
    return getShiftForDate(new Date());
}

export function getPreferredShiftValues(currentShift) {
    if (currentShift === 'APAC') return ['APAC', 'IST'];
    if (currentShift === 'EMEA') return ['EMEA', 'IST'];
    if (currentShift === 'AMER') return ['AMER', 'AMER-PST', 'AMER-EST'];
    return [currentShift];
}

export function buildPreferredShiftCondition(shiftValues) {
    return shiftValues.length === 1
        ? `Preferred_Shift__c='${shiftValues[0]}'`
        : `Preferred_Shift__c IN ('${shiftValues.join("','")}')`;
}

// Weekend utilities reused elsewhere
export function getWeekendSignatureTemplate(caseSeverity = '') {
    return `Hi @,\nNew Sev${caseSeverity} case assigned to you & App is updated.\nThank You`;
}

export function getGHOTemplate() {
    if (isCurrentlyWeekend()) {
        return `Hi @,\nGHO (WOC) case assigned to you & App is updated.\nThank You`;
    }
    return `Hi @QB,\nGreetings for the day.!\n\nKindly assist with GHO case assignment on this case.\n\nThank You\n#GHOTriage`;
}
