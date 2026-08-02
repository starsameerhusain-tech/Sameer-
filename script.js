// Global dataset stores
let salesData = [];
let currentFilteredData = []; 

// Track index of the row currently being edited (-1 means no row is being edited)
let editingIndex = -1;

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

            saveToStorage();
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

// Helper to save to LocalStorage
function saveToStorage() {
    try {
        localStorage.setItem("salesData", JSON.stringify(salesData));
    } catch(e) {}
}

// ------------------------------------------
// 2. Filter Function (Unified Engine)
// ------------------------------------------
function filterAll(e) {
    if (e && e.preventDefault) e.preventDefault();

    const nameInput = document.getElementById("customerSearch");
    const emailInput = document.getElementById("emailSearch");
    const monthInput = document.getElementById("monthSearch");

    const nameQuery = nameInput ? nameInput.value.trim().toLowerCase() : "";
    const emailQuery = emailInput ? emailInput.value.trim().toLowerCase() : "";
    const monthQuery = monthInput ? monthInput.value.trim().toLowerCase() : "";

    currentFilteredData = salesData.filter(item => {
        const nameVal = (item.name || "").toLowerCase();
        const emailVal = (item.email || "").toLowerCase();
        const monthVal = (item.month || "").toLowerCase();

        const matchesName = !nameQuery || nameVal.includes(nameQuery);
        const matchesEmail = !emailQuery || emailVal.includes(emailQuery);
        const matchesMonth = !monthQuery || monthVal.includes(monthQuery);

        return matchesName && matchesEmail && matchesMonth;
    });

    renderTable(currentFilteredData);
    renderDashboard(currentFilteredData);
}

function searchData(e) {
    filterAll(e);
}

// ------------------------------------------
// 3. Render Table (With Edit / Save Actions)
// ------------------------------------------
function renderTable(data) {
    const tbody = document.getElementById("salesTable");
    if (!tbody) return;

    if (!data || data.length === 0) {
        tbody.innerHTML = "<tr><td colspan='11' class='text-center'>No Data Available</td></tr>";
        return;
    }

    let html = "";
    data.forEach((r, index) => {
        // Find master array index for this row
        const masterIndex = salesData.indexOf(r);

        if (editingIndex === masterIndex) {
            // EDITING ROW VIEW
            html += `
            <tr class="table-warning">
                <td><input type="text" id="edit_month" class="form-control form-control-sm" value="${r.month}"></td>
                <td><input type="text" id="edit_projectName" class="form-control form-control-sm" value="${r.projectName}"></td>
                <td><input type="email" id="edit_email" class="form-control form-control-sm" value="${r.email}"></td>
                <td><input type="text" id="edit_name" class="form-control form-control-sm" value="${r.name}"></td>
                <td><input type="text" id="edit_source" class="form-control form-control-sm" value="${r.source}"></td>
                <td><input type="text" id="edit_language" class="form-control form-control-sm" value="${r.language}"></td>
                <td><input type="number" id="edit_rate" class="form-control form-control-sm" value="${r.rate}" oninput="autoCalcAmount()"></td>
                <td><input type="number" id="edit_hours" class="form-control form-control-sm" value="${r.hours}" oninput="autoCalcAmount()"></td>
                <td><input type="number" id="edit_amount" class="form-control form-control-sm" value="${r.amount}" readonly></td>
                <td><input type="text" id="edit_projectCode" class="form-control form-control-sm" value="${r.projectCode}"></td>
                <td>
                    <button class="btn btn-sm btn-success me-1" onclick="saveRow(${masterIndex})">Save</button>
                    <button class="btn btn-sm btn-secondary" onclick="cancelEdit()">Cancel</button>
                </td>
            </tr>`;
        } else {
            // NORMAL DISPLAY VIEW
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
                <td>
                    <button class="btn btn-sm btn-outline-primary" onclick="editRow(${masterIndex})">Edit</button>
                </td>
            </tr>`;
        }
    });

    tbody.innerHTML = html;
}

// Automatically recalculate Amount = Rate * Hours while typing
function autoCalcAmount() {
    const rate = Number(document.getElementById("edit_rate")?.value || 0);
    const hours = Number(document.getElementById("edit_hours")?.value || 0);
    const amountEl = document.getElementById("edit_amount");
    if (amountEl) {
        amountEl.value = rate * hours;
    }
}

// Trigger Edit Mode for a row
function editRow(masterIndex) {
    editingIndex = masterIndex;
    renderTable(currentFilteredData);
}

// Cancel Editing
function cancelEdit() {
    editingIndex = -1;
    renderTable(currentFilteredData);
}

// Save Changes to Row
function saveRow(masterIndex) {
    if (masterIndex < 0 || masterIndex >= salesData.length) return;

    const rate = Number(document.getElementById("edit_rate").value || 0);
    const hours = Number(document.getElementById("edit_hours").value || 0);

    salesData[masterIndex] = {
        month:       document.getElementById("edit_month").value.trim(),
        projectName: document.getElementById("edit_projectName").value.trim(),
        email:       document.getElementById("edit_email").value.trim(),
        name:        document.getElementById("edit_name").value.trim(),
        source:      document.getElementById("edit_source").value.trim(),
        language:    document.getElementById("edit_language").value.trim(),
        rate:        rate,
        hours:       hours,
        amount:      rate * hours, // Auto-calculated Amount
        projectCode: document.getElementById("edit_projectCode").value.trim()
    };

    saveToStorage();
    editingIndex = -1;
    filterAll(); // Refreshes table, exports, and KPI cards
}

// ------------------------------------------
// 4. Render Dashboard KPI Cards
// ------------------------------------------
function renderDashboard(data) {
    const totalSales = data.reduce((sum, item) => sum + (Number(item.amount) || 0), 0);
    const totalHours = data.reduce((sum, item) => sum + (Number(item.hours) || 0), 0);
    const avgRate = data.length > 0 
        ? Math.round(data.reduce((sum, item) => sum + (Number(item.rate) || 0), 0) / data.length) 
        : 0;

    const setEl = (id, val) => {
        const el = document.getElementById(id);
        if (el) el.innerText = val;
    };

    setEl("totalSales", totalSales.toLocaleString());
    setEl("totalCustomers", data.length.toLocaleString());
    setEl("totalProducts", totalHours.toLocaleString());
    setEl("totalQty", avgRate.toLocaleString());
}

// ------------------------------------------
// 5. Download Excel (.xlsx)
// ------------------------------------------
function downloadExcel() {
    const dataToExport = (currentFilteredData && currentFilteredData.length > 0) 
        ? currentFilteredData 
        : salesData;

    if (!dataToExport || dataToExport.length === 0) {
        alert("No data available to download.");
        return;
    }

    const formattedData = dataToExport.map(r => ({
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

    const worksheet = XLSX.utils.json_to_sheet(formattedData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sales Records");

    XLSX.writeFile(workbook, "SalesData_Export.xlsx");
}

// ------------------------------------------
// 6. Download CSV (.csv)
// ------------------------------------------
function downloadCSV() {
    const dataToExport = (currentFilteredData && currentFilteredData.length > 0) 
        ? currentFilteredData 
        : salesData;

    if (!dataToExport || dataToExport.length === 0) {
        alert("No data available to download.");
        return;
    }

    let csv = "Month,Project Name,Email ID,Name,Source,Language,Rate,Hours,Amount,Project Code\n";
    dataToExport.forEach(r => {
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

// ------------------------------------------
// 7. Initialization
// ------------------------------------------
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
