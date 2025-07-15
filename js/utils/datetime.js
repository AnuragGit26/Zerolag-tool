export function timeElapsed(createdDate) {
    const now = new Date();
    const diff = now - createdDate;
    const seconds = Math.floor(diff / 1000);
    const minutes = Math.floor(seconds / 60);
    const hours = Math.floor(minutes / 60);
    const days = Math.floor(hours / 24);

    if (days > 0) {
        return `${days}d ago`;
    }
    if (hours > 0) {
        return `${hours}h ${minutes % 60}m ago`;
    }
    if (minutes > 0) {
        return `${minutes}m ago`;
    }
    return `${seconds}s ago`;
}

export function addMinutes(numOfMinutes, date = new Date()) {
    date.setMinutes(date.getMinutes() + numOfMinutes);
    return date;
}

export function isCurrentlyWeekend() {
    const now = new Date();
    const currentOffset = now.getTimezoneOffset();
    const istOffset = -330;

    let istTime;
    if (currentOffset === istOffset) {
        istTime = now;
    } else {
        const istOffsetMs = 5.5 * 60 * 60 * 1000;
        istTime = new Date(now.getTime() + istOffsetMs + (currentOffset * 60 * 1000));
    }

    const day = istTime.getDay();
    const hour = istTime.getHours();
    const minute = istTime.getMinutes();
    const timeInMinutes = hour * 60 + minute;
    const cutoffTime = 5 * 60 + 30;

    // Weekend: Saturday 5:30 AM IST to Monday 5:30 AM IST
    if (day === 6 && timeInMinutes >= cutoffTime) return true;
    if (day === 0) return true;
    if (day === 1 && timeInMinutes < cutoffTime) return true;

    return false;
}
