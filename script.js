// Global Configuration
const GOOGLE_SHEET_URL = "https://script.google.com/macros/s/AKfycbxvMtq1wd_xOeAyeyNo3DYnGcxlENSgrUEfB8xkuzunz7QBqBVyhKXHRT3LuLBb6jNQTw/exec";

// State Management
let salesData = JSON.parse(localStorage.getItem('salesData')) || [];

// Wait for DOM to load
document.addEventListener('DOMContentLoaded', function () {
    updateDashboardUI();
    renderTable(salesData);

    // Attach File Upload Event Listener
    const uploadBtn = document.getElementById('uploadBtn');
    if (uploadBtn) {
        uploadBtn.addEventListener('click', handleFileUpload);
    }

    // Attach Search Event Listeners (Triggers on click AND as you type)
    const searchBtn = document.getElementById('searchBtn');
    if (searchBtn) {
        searchBtn.addEventListener('click', function(e) {
            e.preventDefault();
            handleSearch();
        });
    }

    // Live search as you type in any search box
    ['searchName', 'searchEmail', 'searchMonth'].forEach(id => {
        const inputElem = document.getElementById(id);
        if (inputElem) {
            inputElem.addEventListener('input', handleSearch);
            inputElem.addEventListener('keypress', function(e) {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    handleSearch();
                }
            });
        }
    });
});

// Handle Search Filtering
function handleSearch() {
    const nameQuery = (document.getElementById('searchName')?.value || '').toLowerCase().trim();
    const emailQuery = (document.getElementById('searchEmail')?.value || '').toLowerCase().trim();
    const monthQuery = (document.getElementById('searchMonth')?.value || '').toLowerCase().trim();

    // If all search inputs are empty, show full data
    if (!nameQuery && !emailQuery && !monthQuery) {
        renderTable(salesData);
        return;
    }

    const filteredData = salesData.filter(item => {
        // Safe string conversions
        const nameVal = String(item.name || item.Name || '').toLowerCase();
        const emailVal = String(item.emailId || item['Email ID'] || '').toLowerCase();
        const monthVal = String(item.month || item.Month || '').toLowerCase();

        const matchesName = !nameQuery || nameVal.includes(nameQuery);
        const matchesEmail = !emailQuery || emailVal.includes(emailQuery);
        const matchesMonth = !monthQuery || monthVal.includes(monthQuery);

        return matchesName && matchesEmail && matchesMonth;
    });

    renderTable(filteredData);
}

// Handle Excel / CSV Upload
function handleFileUpload() {
    const fileInput = document.getElementById('excelFile');
    const statusMsg = document.getElementById('uploadStatus');

    if (!fileInput || !fileInput.files || fileInput.files.length === 0) {
        alert("Please select an Excel or CSV file first!");
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // Convert to JSON
            const rawJson = XLSX.utils.sheet_to_json(worksheet);

            if (!rawJson.length) {
                alert("The selected file contains no data.");
                return;
            }

            // Map incoming data to schema fields
            const formattedData = rawJson.map(row => ({
                month: row['Month'] || row['month'] || '',
                projectName: row['Project Name'] || row['projectName'] || '',
                emailId: row['Email ID'] || row['emailId'] || '',
                name: row['Name'] || row['name'] || '',
                source: row['Source'] || row['source'] || '',
                language: row['Language'] || row['language'] || '',
                rate: parseFloat(row['Rate'] || row['rate'] || 0),
                hours: parseFloat(row['Hours'] || row['hours'] || 0),
                amount: parseFloat(row['Amount'] || row['amount'] || 0),
                projectCode: row['Project Code'] || row['projectCode'] || ''
            }));

            // 1. Update Local UI Data
            salesData = formattedData;
            localStorage.setItem('salesData', JSON.stringify(salesData));

            // 2. Refresh UI
            updateDashboardUI();
            renderTable(salesData);

            // 3. Dispatch to Google Sheets Backend
            sendToGoogleSheet(formattedData);

            if (statusMsg) {
                statusMsg.className = "text-success fw-bold mt-2 d-block";
                statusMsg.innerHTML = `✓ ${formattedData.length} records saved locally and sent to Google Sheets!`;
            }

        } catch (err) {
            console.error("Error parsing file:", err);
            alert("Error reading file. Please upload a valid .xlsx or .csv file.");
        }
    };

    reader.readAsArrayBuffer(file);
}

// Send Data to Google Apps Script
function sendToGoogleSheet(data) {
    if (!GOOGLE_SHEET_URL || GOOGLE_SHEET_URL.includes("YOUR_DEPLOYMENT_ID_HERE")) {
        console.warn("Google Sheet URL is not configured properly in script.js");
        return;
    }

    fetch(GOOGLE_SHEET_URL, {
        method: 'POST',
        mode: 'no-cors',
        headers: {
            'Content-Type': 'text/plain'
        },
        body: JSON.stringify(data)
    })
    .then(() => {
        console.log("Data successfully dispatched to Google Sheet!");
    })
    .catch(error => {
        console.error("Error sending data to Google Sheet:", error);
    });
}

// Refresh Dashboard KPI Cards
function updateDashboardUI() {
    const totalAmountElem = document.getElementById('totalAmount');
    const totalRecordsElem = document.getElementById('totalRecords');
    const totalHoursElem = document.getElementById('totalHours');
    const avgRateElem = document.getElementById('avgRate');

    if (!salesData.length) {
        if (totalAmountElem) totalAmountElem.innerText = "0";
        if (totalRecordsElem) totalRecordsElem.innerText = "0";
        if (totalHoursElem) totalHoursElem.innerText = "0";
        if (avgRateElem) avgRateElem.innerText = "0";
        return;
    }

    let totalAmount = 0;
    let totalHours = 0;
    let totalRate = 0;

    salesData.forEach(item => {
        totalAmount += item.amount || 0;
        totalHours += item.hours || 0;
        totalRate += item.rate || 0;
    });

    const avgRate = salesData.length ? (totalRate / salesData.length).toFixed(0) : 0;

    if (totalAmountElem) totalAmountElem.innerText = totalAmount.toLocaleString();
    if (totalRecordsElem) totalRecordsElem.innerText = salesData.length;
    if (totalHoursElem) totalHoursElem.innerText = totalHours.toLocaleString();
    if (avgRateElem) avgRateElem.innerText = avgRate;
}

// Render Records Table
function renderTable(data) {
    const tbody = document.getElementById('salesTableBody');
    if (!tbody) return;

    if (!data || data.length === 0) {
        tbody.innerHTML = `<tr><td colspan="10" class="text-center py-4 text-muted fw-bold">No Matching Records Found</td></tr>`;
        return;
    }

    let rowsHtml = '';
    data.forEach(item => {
        rowsHtml += `
            <tr>
                <td>${item.month}</td>
                <td>${item.projectName}</td>
                <td>${item.emailId}</td>
                <td>${item.name}</td>
                <td>${item.source}</td>
                <td>${item.language}</td>
                <td>${item.rate}</td>
                <td>${item.hours}</td>
                <td>${item.amount}</td>
                <td>${item.projectCode}</td>
            </tr>
        `;
    });

    tbody.innerHTML = rowsHtml;
}
