import { isCurrentlyWeekend } from './datetime.js';

export function applyFilter(filterValue) {
    // Apply both filter and search criteria together
    applyFilterAndSearch(filterValue, getCurrentSearchValue());
}

export function applySearch(searchValue) {
    // Apply both search and filter criteria together
    applyFilterAndSearch(getCurrentFilterValue(), searchValue);
}

export function applyFilterAndSearch(filterValue, searchValue) {
    let allCases = document.querySelectorAll("#parentSigSev2 .case-card");

    allCases.forEach(function (caseDiv) {
        let shouldShow = true;

        // Apply filter criteria
        let checkbox = caseDiv.querySelector(".action-checkbox");
        if (checkbox && filterValue !== "all") {
            let isActionTaken = checkbox.checked;
            if (filterValue === "action-taken" && !isActionTaken) {
                shouldShow = false;
            } else if (filterValue === "not-action-taken" && isActionTaken) {
                shouldShow = false;
            }
        }

        // Apply search criteria
        if (shouldShow && searchValue && searchValue.trim() !== "") {
            let caseNumberElement = caseDiv.querySelector(".case-info-item:nth-child(3) span:nth-child(2)");
            if (caseNumberElement) {
                let caseNumber = caseNumberElement.textContent.toLowerCase();
                if (!caseNumber.includes(searchValue.toLowerCase())) {
                    shouldShow = false;
                }
            } else {
                shouldShow = false;
            }
        }

        // Show or hide the case based on combined criteria
        caseDiv.style.display = shouldShow ? "block" : "none";
    });
}

// Helper functions to get current values
function getCurrentFilterValue() {
    const filterSelect = document.getElementById("action-filter");
    return filterSelect ? filterSelect.value : "all";
}

function getCurrentSearchValue() {
    const searchInput = document.getElementById("search-input");
    return searchInput ? searchInput.value.trim() : "";
}

export function updateWeekendModeIndicator(currentMode = (localStorage.getItem('caseTriageMode') || 'signature')) {
    const weekendIndicator = document.getElementById("weekend-mode-indicator");
    const slaIndicator = document.getElementById("sla-indicator");

    const weekend = isCurrentlyWeekend();

    if (weekendIndicator) {
        weekendIndicator.style.display = weekend ? "block" : "none";
    }

    if (!slaIndicator) return;

    if (currentMode === 'signature') {
        if (weekend) {
            slaIndicator.style.display = 'inline-block';
            slaIndicator.textContent = 'SLA · Weekend relaxed';
        } else {
            slaIndicator.style.display = 'inline-block';
            slaIndicator.textContent = 'SLA · SEV1: 5m · SEV2: 30m';
        }
    } else {
        // Hide in Premier mode
        slaIndicator.style.display = 'none';
    }
}
