import { isCurrentlyWeekend } from './datetime.js';

export function applyFilter(filterValue) {
    let allCases = document.querySelectorAll("#parentSigSev2 > .d-style");
    allCases.forEach(function (caseDiv) {
        let checkbox = caseDiv.querySelector(".action-checkbox");
        if (checkbox) {
            let isActionTaken = checkbox.checked;
            if (filterValue === "all") {
                caseDiv.style.display = "block";
            } else if (filterValue === "action-taken" && isActionTaken) {
                caseDiv.style.display = "block";
            } else if (filterValue === "not-action-taken" && !isActionTaken) {
                caseDiv.style.display = "block";
            } else {
                caseDiv.style.display = "none";
            }
        }
    });
}

export function updateWeekendModeIndicator() {
    const weekendIndicator = document.getElementById("weekend-mode-indicator");
    if (isCurrentlyWeekend()) {
        weekendIndicator.style.display = "block";
    } else {
        weekendIndicator.style.display = "none";
    }
}
