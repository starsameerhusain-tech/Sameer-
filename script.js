// Global dataset stores
let salesData = [];
let currentFilteredData = []; // Keeps track of currently visible/filtered rows

// 1. Upload & Parse Excel File
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

            const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

            if (!rawRows || rawRows.length === 0) {
                alert("Excel file is empty.");
                return;
            }

            const parseNum = (val) => {
                if (typeof val === "number") return val;
                if (!val) return 0;
                return Number(String(val).replace(/[^0-9.-]+/g, "")) || 0;
            };

            let startRowIndex = -1;
            for (let i = 0; i < rawRows.length; i++) {
                const firstColValue = String(rawRows[i][0] || "").trim().toLowerCase();
                if (firstColValue && firstColValue !== "month") {
                    startRowIndex = i;
                    break;
                }
            }

            if (startRowIndex === -1) startRowIndex = 1;

            salesData = [];

            for (let i = startRowIndex; i < rawRows.length; i++) {
                const row = rawRows[i];
                if (!row || row.length === 0) continue;

                const monthVal = String(row[0] || "").trim();
                const projVal  = String(row[1] || "").trim();
                const emailVal = String(row[2] || "").trim();

                if (!monthVal && !projVal && !emailVal) continue;

                salesData.push({
                    month:       monthVal,
                    projectName: projVal,
                    email:       emailVal,
                    name:        String(row[3] || "").trim(),
                    source:      String(row[4] || "").trim(),
                    language:    String(row[5] || "").trim(),
                    rate:        parseNum(row[6]),
                    hours:       parseNum(row[7]),
                    amount:      parseNum(row[8]),
                    projectCode: String(row[9] || "").trim()
                });
            }

            try {
                localStorage.setItem("salesData", JSON.stringify(salesData));
            } catch(e) {}

            filterAll();

            const statusEl = document.getElementById("uploadStatus");
            if (statusEl) {
                statusEl.innerHTML = "✅ " + salesData.length + " records uploaded successfully.";
            }
        } catch (err) {
            console.error("Error reading file:", err);
            alert("Error reading Excel file.");
        }
    };

    reader.readAsArrayBuffer(file);
}

// 2. Combined Filter Engine: Updates Table, Cards & Active Export Data
function filterAll() {
    const nameInput = document.getElementById("customerSearch");
    const emailInput = document.getElementById("emailSearch");
    const monthInput = document.getElementById("monthSearch");

    const nameQuery = nameInput ? nameInput.value.trim().toLowerCase() : "";
    const emailQuery = emailInput ? emailInput.value.trim().toLowerCase() : "";
    const monthQuery = monthInput ? monthInput.value.trim().toLowerCase() : "";

    // Filter master array
    currentFilteredData = salesData.filter(item => {
        const nameVal = (item.name || "").toLowerCase();
        const emailVal = (item.email || "").toLowerCase();
        const monthVal = (item.month || "").toLowerCase();

        const matchesName = !nameQuery || nameVal.includes(nameQuery);
        const matchesEmail = !emailQuery || emailVal.includes(emailQuery);
        const matchesMonth = !monthQuery || monthVal.includes(monthQuery);

        return matchesName && matchesEmail && matchesMonth;
    });

    // Update Table
    renderTable(currentFilteredData);

    // Update Dashboard Cards
    renderDashboard(currentFilteredData);
}

// 3. Render Table Function
function renderTable(data) {
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

// 4. Render Dashboard KPI Cards Function
function renderDashboard(data) {
    const totalSales = data.reduce((sum, item) => sum + (Number(item.amount) || 0), 0);
    const totalHours = data.reduce((sum, item) => sum + (Number(item.hours) || 0), 0);
    const avgRate = data.length > 0 
        ? Math.round(data.reduce((sum, item) => sum + (Number(item.rate) || 0), 0) / data.length) 
        : 0;

    document.getElementById("totalSales").innerText = totalSales.toLocaleString();
    document.getElementById("totalCustomers").innerText = data.length.toLocaleString();
    document.getElementById("totalProducts").innerText = totalHours.toLocaleString();
    document.getElementById("totalQty").innerText = avgRate.toLocaleString();
}

// 5. Download Excel (.xlsx) - Active for Filtered and Unfiltered Data
function downloadExcel() {
    if (!currentFilteredData || currentFilteredData.length === 0) {
        alert("No data available to download.");
        return;
    }

    // Format data with exact header labels for clean Excel output
    const formattedData = currentFilteredData.map(r => ({
        "Month": r.month,
        "Project Name": r.projectName,
        "Email ID": r.email,
        "Name": r.name,
        "Source": r.source,
        "Language": r.language,
        "Rate": r.rate,
        "Hours": r.hours,
        "Amount": r.amount,
        "Project Code": r.projectCode
    }));

    // Create Worksheet & Workbook using xlsx.full.min.js
    const worksheet = XLSX.utils.json_to_sheet(formattedData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sales Records");

    // Download file
    XLSX.writeFile(workbook, "SalesData_Export.xlsx");
}

// 6. Download CSV (.csv) - Active for Filtered and Unfiltered Data
function downloadCSV() {
    if (!currentFilteredData || currentFilteredData.length === 0) {
        alert("No data available to download.");
        return;
    }

    let csv = "Month,Project Name,Email ID,Name,Source,Language,Rate,Hours,Amount,Project Code\n";
    currentFilteredData.forEach(r => {
        csv += `"${r.month}","${r.projectName}","${r.email}","${r.name}","${r.source}","${r.language}",${r.rate},${r.hours},${r.amount},"${r.projectCode}"\n`;
    });

    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "SalesData_Export.csv";
    a.click();
    URL.revokeObjectURL(url);
}

// 7. Page Initialization
document.addEventListener("DOMContentLoaded", () => {
    try {
        const saved = localStorage.getItem("salesData");
        if (saved) {
            salesData = JSON.parse(saved);
            const statusEl = document.getElementById("uploadStatus");
            if (statusEl) {
                statusEl.innerHTML = "🔄 Loaded " + salesData.length + " records from browser memory.";
            }
        }
    } catch(e) {}

    filterAll();
});
