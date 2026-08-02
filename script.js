// =========================================================================
// 1. CONFIGURATION
// =========================================================================
// Replace the link below with your Google Apps Script Web App URL
const GOOGLE_SHEET_URL = "https://script.google.com/macros/s/AKfycbzjaw45_8ABlqMgrSUcpbpkLpbO4vREuqJriOk4PabMmLq-K9lnGF7Rc7QS1QCoDnLVFg/exec";


// =========================================================================
// 2. INITIALIZATION
// =========================================================================
document.addEventListener("DOMContentLoaded", () => {
    loadSalesData();
});


// =========================================================================
// 3. CORE EXCEL UPLOAD & DATA PARSING
// =========================================================================
function uploadExcel() {
    const fileInput = document.getElementById("fileInput");
    const file = fileInput ? fileInput.files[0] : null;

    if (!file) {
        alert("Please select an Excel or CSV file first!");
        return;
    }

    const reader = new FileReader();

    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // Convert raw sheet data into JSON objects
            const rawJson = XLSX.utils.sheet_to_json(worksheet);

            if (!rawJson || rawJson.length === 0) {
                alert("Uploaded file contains no data!");
                return;
            }

            // Normalize JSON data keys to match the system schema
            const formattedData = rawJson.map(row => {
                const rate = Number(getProp(row, "rate")) || 0;
                const hours = Number(getProp(row, "hours")) || 0;
                const amount = Number(getProp(row, "amount")) || (rate * hours);

                return {
                    month: String(getProp(row, "month")).trim(),
                    projectName: String(getProp(row, "project name") || getProp(row, "projectname")).trim(),
                    emailId: String(getProp(row, "email id") || getProp(row, "emailid") || getProp(row, "email")).trim(),
                    name: String(getProp(row, "name")).trim(),
                    source: String(getProp(row, "source")).trim(),
                    language: String(getProp(row, "language")).trim(),
                    rate: rate,
                    hours: hours,
                    amount: amount,
                    projectCode: String(getProp(row, "project code") || getProp(row, "projectcode")).trim()
                };
            });

            // 1. Save to LocalStorage for local dashboard & reports.html
            localStorage.setItem("salesData", JSON.stringify(formattedData));
            
            // 2. Refresh the local UI table
            loadSalesData();

            // 3. Push data to Google Sheets in background
            sendToGoogleSheet(formattedData);

            alert(`Success! ${formattedData.length} records processed and sent to Google Sheets.`);

        } catch (error) {
            console.error("Error parsing file:", error);
            alert("Failed to process file. Please upload a valid Excel (.xlsx) or CSV file.");
        }
    };

    reader.readAsArrayBuffer(file);
}


// =========================================================================
// 4. GOOGLE SHEETS API SYNC (FIXED CORS & CONTENT-TYPE)
// =========================================================================
async function sendToGoogleSheet(data) {
    if (!GOOGLE_SHEET_URL || GOOGLE_SHEET_URL.includes("PASTE_YOUR_GOOGLE_WEB_APP_URL_HERE")) {
        console.warn("Google Sheet URL is not configured in script.js!");
        return;
    }

    try {
        // Send as text/plain to avoid CORS pre-flight OPTIONS blocks from Google Apps Script
        await fetch(GOOGLE_SHEET_URL, {
            method: "POST",
            mode: "no-cors", 
            headers: {
                "Content-Type": "text/plain;charset=utf-8"
            },
            body: JSON.stringify(data)
        });

        console.log("Data successfully dispatched to Google Sheet!");
    } catch (err) {
        console.error("Error syncing to Google Sheet:", err);
    }
}


// =========================================================================
// 5. HELPER FUNCTIONS & RENDER LOGIC
// =========================================================================

// Robust case-insensitive and space-insensitive column matcher
function getProp(obj, propName) {
    if (!obj) return "";
    const cleanProp = propName.toLowerCase().replace(/[\s_]/g, '');
    const key = Object.keys(obj).find(k => k.toLowerCase().replace(/[\s_]/g, '') === cleanProp);
    return key ? obj[key] : "";
}

// Render data to index.html dashboard table
function loadSalesData() {
    const raw = localStorage.getItem("salesData");
    const tbody = document.getElementById("salesTableBody");
    if (!tbody) return; // Exit if called from a page without the table element

    if (!raw) {
        tbody.innerHTML = `<tr><td colspan="10" class="text-center text-muted">No sales data loaded. Upload a file above.</td></tr>`;
        return;
    }

    try {
        const salesData = JSON.parse(raw);
        if (!Array.isArray(salesData) || salesData.length === 0) {
            tbody.innerHTML = `<tr><td colspan="10" class="text-center text-muted">No sales data available.</td></tr>`;
            return;
        }

        let html = "";
        salesData.forEach((row) => {
            html += `
            <tr>
                <td>${row.month || '-'}</td>
                <td>${row.projectName || '-'}</td>
                <td>${row.emailId || '-'}</td>
                <td>${row.name || '-'}</td>
                <td>${row.source || '-'}</td>
                <td>${row.language || '-'}</td>
                <td>$${row.rate}</td>
                <td>${row.hours} hrs</td>
                <td class="fw-bold text-success">$${row.amount.toLocaleString()}</td>
                <td><span class="badge bg-secondary">${row.projectCode || '-'}</span></td>
            </tr>`;
        });

        tbody.innerHTML = html;
    } catch(e) {
        console.error("Error rendering table:", e);
    }
}
