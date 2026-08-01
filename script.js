// ==========================================
// Sales Management System - Position-Based Fix
// ==========================================

let salesData = [];

// ------------------------------------------
// 1. Upload & Parse Excel File
// ------------------------------------------
function uploadExcel() {
    const fileInput = document.getElementById("excelFile");
    const file = fileInput ? fileInput.files[0] : null;

    if (!file) {
        alert("Please select an Excel file.");
        return;
    }

    const reader = new FileReader();

    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];

            // Read sheet strictly as 2D array matrix [[row1], [row2], ...]
            const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

            if (!rawRows || rawRows.length === 0) {
                alert("Excel file appears to be empty.");
                return;
            }

            // Number parser helper
            const parseNum = (val) => {
                if (typeof val === "number") return val;
                if (!val) return 0;
                return Number(String(val).replace(/[^0-9.-]+/g, "")) || 0;
            };

            // Find the exact row index where data starts (Row containing 'January' or data)
            let startRowIndex = -1;
            for (let i = 0; i < rawRows.length; i++) {
                const firstColValue = String(rawRows[i][0] || "").trim().toLowerCase();
                // Skip header row ("month") and find actual month data
                if (firstColValue && firstColValue !== "month") {
                    startRowIndex = i;
                    break;
                }
            }

            if (startRowIndex === -1) {
                // Fallback: If no "Month" header was found, assume row 1 is data (index 1)
                startRowIndex = 1;
            }

            salesData = [];

            // Extract strictly by position (Column A = 0, B = 1, C = 2, etc.)
            for (let i = startRowIndex; i < rawRows.length; i++) {
                const row = rawRows[i];
                if (!row || row.length === 0) continue;

                const monthVal = String(row[0] || "").trim();
                const projVal  = String(row[1] || "").trim();
                const emailVal = String(row[2] || "").trim();

                // Stop or skip if the row is completely empty
                if (!monthVal && !projVal && !emailVal) continue;

                salesData.push({
                    month:       monthVal,                   // Col A (Index 0)
                    projectName: projVal,                    // Col B (Index 1)
                    email:       emailVal,                   // Col C (Index 2)
                    name:        String(row[3] || "").trim(),// Col D (Index 3)
                    source:      String(row[4] || "").trim(),// Col E (Index 4)
                    language:    String(row[5] || "").trim(),// Col F (Index 5)
                    rate:        parseNum(row[6]),           // Col G (Index 6)
                    hours:       parseNum(row[7]),           // Col H (Index 7)
                    amount:      parseNum(row[8]),           // Col I (Index 8)
                    projectCode: String(row[9] || "").trim() // Col J (Index 9)
                });
            }

            loadTable(salesData);
            updateDashboard();

            const statusEl = document.getElementById("uploadStatus");
            if (statusEl) {
                statusEl.innerHTML = "✅ " + salesData.length + " records uploaded successfully.";
            }
        } catch (err) {
            console.error("Error parsing Excel:", err);
            alert("Error reading file. Please check console (F12) for details.");
        }
    };

    reader.readAsArrayBuffer(file);
}

// ------------------------------------------
// 2. Render Table Rows
// ------------------------------------------
function loadTable(data) {
    const tbody = document.getElementById("salesTable");
    if (!tbody) return;

    if (!data || data.length === 0) {
        tbody.innerHTML = "<tr><td colspan='10' class='text-center'>No Data Available</td></tr>";
        return;
    }

    let html = "";
    data.forEach(r => {
        html += `
        <tr>
            <td>${r.month}</td>
            <td>${r.projectName}</td>
            <td>${r.email}</td>
            <td>${r.name}</td>
            <td>${r.source}</td>
            <td>${r.language}</td>
            <td>${r.rate}</td>
            <td>${r.hours}</td>
            <td>${r.amount}</td>
            <td>${r.projectCode}</td>
        </tr>`;
    });

    tbody.innerHTML = html;
}

// ------------------------------------------
// 3. Update Dashboard Cards
// ------------------------------------------
function updateDashboard() {
    const setSafeText = (id, text) => {
        const el = document.getElementById(id);
        if (el) el.innerHTML = text;
    };

    const totalSales = salesData.reduce((a, b) => a + (b.amount || 0), 0);
    const totalHours = salesData.reduce((a, b) => a + (b.hours || 0), 0);
    const avgRate = salesData.length > 0 
        ? Math.round(salesData.reduce((a, b) => a + (b.rate || 0), 0) / salesData.length) 
        : 0;

    setSafeText("totalSales", totalSales.toLocaleString());
    setSafeText("totalCustomers", salesData.length); // Total Records
    setSafeText("totalProducts", totalHours.toLocaleString()); // Total Hours
    setSafeText("totalQty", avgRate); // Avg Rate
}

// ------------------------------------------
// 4. Live Search (Name, Email ID, & Month)
// ------------------------------------------
function searchData(e) {
    if (e && e.preventDefault) {
        e.preventDefault();
    }

    const nameInput = document.getElementById("customerSearch");
    const emailInput = document.getElementById("emailSearch");
    const monthInput = document.getElementById("monthSearch");

    const nameQuery = nameInput ? nameInput.value.trim().toLowerCase() : "";
    const emailQuery = emailInput ? emailInput.value.trim().toLowerCase() : "";
    const monthQuery = monthInput ? monthInput.value.trim().toLowerCase() : "";

    const filtered = salesData.filter(item => {
        const nameVal = item.name.toLowerCase();
        const emailVal = item.email.toLowerCase();
        const monthVal = item.month.toLowerCase();

        const matchesName = !nameQuery || nameVal.includes(nameQuery);
        const matchesEmail = !emailQuery || emailVal.includes(emailQuery);
        const matchesMonth = !monthQuery || monthVal.includes(monthQuery);

        return matchesName && matchesEmail && matchesMonth;
    });

    loadTable(filtered);
}

// ------------------------------------------
// 5. Download CSV
// ------------------------------------------
function downloadCSV() {
    if (salesData.length === 0) {
        alert("No data available to download.");
        return;
    }

    let csv = "Month,Project Name,Email ID,Name,Source,Language,Rate,Hours,Amount,Project Code\n";
    salesData.forEach(r => {
        csv += `"${r.month}","${r.projectName}","${r.email}","${r.name}","${r.source}","${r.language}",${r.rate},${r.hours},${r.amount},"${r.projectCode}"\n`;
    });

    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "SalesData.csv";
    a.click();
    URL.revokeObjectURL(url);
}

function downloadExcel() {
    alert("Excel download feature coming soon!");
}

// ------------------------------------------
// 6. Automatic Initialization & Key Listeners
// ------------------------------------------
document.addEventListener("DOMContentLoaded", () => {
    const nameInput = document.getElementById("customerSearch");
    const emailInput = document.getElementById("emailSearch");
    const monthInput = document.getElementById("monthSearch");

    if (nameInput) nameInput.addEventListener("input", searchData);
    if (emailInput) emailInput.addEventListener("input", searchData);
    if (monthInput) monthInput.addEventListener("input", searchData);

    updateDashboard();
});
