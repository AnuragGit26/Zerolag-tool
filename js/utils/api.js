const SPREADSHEET_ID = '1BKxQLGFrczjhcx9rEt-jXGvlcCPQblwBhFJjoiDD7TI';

function getAuthToken(callback) {
    chrome.identity.getAuthToken({ interactive: true }, function (token) {
        if (chrome.runtime.lastError) {
            console.error(chrome.runtime.lastError);
            return;
        }
        callback(token);
    });
}

export function trackAction(dateofAction, caseNumber, severity, actionType, cloud, currentMode, currentUserName) {
    getAuthToken(function (token) {
        const sheetName = currentMode === 'premier' ? 'premier' : 'signature';

        // Handle undefined dateofAction by using current date as fallback
        const actionDate = dateofAction || new Date();

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


        // Updated column order: Date, Type, Time of Action, Case Number, Engineer Name, Severity, Cloud
        const values = [
            [pstDate, actionType, istTime, caseNumber, currentUserName, severity, cloud]
        ];

        const body = {
            values: values
        };

        fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${sheetName}:append?valueInputOption=USER_ENTERED`, {
            method: 'POST',
            headers: {
                'Authorization': 'Bearer ' + token,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(body)
        }).then(response => response.json())
            .then(data => {
                console.log('Sheet updated:', data);
            })
            .catch(error => {
                console.error('Error updating sheet:', error);
            });
    });
}
