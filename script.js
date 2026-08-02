// Paste your copied Google Web App URL inside the quotes below
const GOOGLE_SHEET_URL = "https://script.google.com/macros/s/AKfycbzjaw45_8ABlqMgrSUcpbpkLpbO4vREuqJriOk4PabMmLq-K9lnGF7Rc7QS1QCoDnLVFg/exec";

document.addEventListener("DOMContentLoaded", () => {
    loadSalesData();
});

// Function to upload Excel file & Sync to Google Sheets
function uploadExcel() {
    const fileInput = document.getElementById("fileInput");
    const file = fileInput.files[0];

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

            // Convert worksheet to JSON
            const rawJson = XLSX.utils.sheet_to_json(worksheet);

            if (rawJson.length === 0) {
                alert("Uploaded file is empty!");
                return;
            }

            // Normalize JSON data keys to match schema
            const formattedData = rawJson.map(row => {
                return {
                    month: getProp(row, "month"),
                    projectName: getProp(row, "project name") || getProp(row, "projectname"),
                    emailId: getProp(row, "email id") || getProp(row, "emailid") || getProp(row, "email"),
                    name: getProp(row, "name"),
                    source: getProp(row, "source"),
                    language: getProp(row, "language"),
                    rate: Number(getProp(row, "rate")) || 0,
                    hours: Number(getProp(row, "hours")) || 0,
                    amount: Number(getProp(row, "amount")) || (Number(getProp(row, "rate")) * Number(getProp(row, "hours"))) || 0,
                    projectCode: getProp(row, "project code") || getProp(row, "projectcode")
                };
            });

            // Save to LocalStorage for instant local display
            localStorage.setItem("salesData", JSON.stringify(formattedData));
            
            // Render local table
            loadSalesData();

            // Send data to Google Sheets in background
            sendToGoogleSheet(formattedData);

            alert(`Success! ${formattedData.length} records loaded into website & syncing to Google Sheets.`);

        } catch (error) {
            console.error("Error reading file:", error);
            alert("Failed to process file. Please ensure it's a valid Excel or CSV file.");
        }
    };

    reader.readAsArrayBuffer(file);
}

// Function to POST data to Google Apps Script Web App
async function sendToGoogleSheet(data) {
    if (!GOOGLE_SHEET_URL || GOOGLE_SHEET_URL.includes("PASTE_YOUR_GOOGLE_WEB_APP_URL_HERE")) {
        console.warn("Google Sheet URL is not configured yet.");
        return;
    }

    try {
        await fetch(GOOGLE_SHEET_URL, {
            method: "POST",
            mode: "no-cors", // Bypasses CORS policy for Apps Script Web Apps
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify(data)
        });
        console.log("Data successfully sent to Google Sheet!");
    } catch (err) {
        console.error("Error syncing to Google Sheet:", err);
    }
}

// Helper function to handle case-insensitive Excel column header matching
function getProp(obj, propName) {
    if (!obj) return "";
    const key = Object.keys(obj).find(k => k.toLowerCase().replace(/[\s_]/g, '') === propName.toLowerCase().replace(/[\s_]/g, ''));
    return key ? obj[key] : "";
}

// Function to load and render table on Dashboard
function loadSalesData() {
    const raw = localStorage.getItem("salesData");
    const tbody = document.getElementById("salesTableBody");
    if (!tbody) return; // Not on index.html page

    if (!raw) {
        tbody.innerHTML = `<tr><td colspan="10" class="text-center text-muted">No sales data available. Upload an Excel file to get started.</td></tr>`;
        return;
    }

    const salesData = JSON.parse(raw);
    let html = "";

    salesData.forEach((row, index) => {
        html += `
        <tr>
            <td>${row.month}</td>
            <td>${row.projectName}</td>
            <td>${row.emailId}</td>
            <td>${row.name}</td>
            <td>${row.source}</td>
            <td>${row.language}</td>
            <td>$${row.rate}</td>
            <td>${row.hours} hrs</td>
            <td class="fw-bold text-success">$${row.amount}</td>
            <td><span class="badge bg-secondary">${row.projectCode}</span></td>
        </tr>`;
    });

    tbody.innerHTML = html;
}
