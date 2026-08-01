// ==========================================
// Sales Management System - Full Script
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
            const jsonData = XLSX.utils.sheet_to_json(sheet);

            // Safe number parser (handles currency strings like "$1,200")
            const parseNum = (val) => {
                if (typeof val === "number") return val;
                if (!val) return 0;
                return Number(String(val).replace(/[^0-9.-]+/g, "")) || 0;
            };

            // Map and normalize rows regardless of header casing or spaces
            salesData = jsonData.map(r => {
                const getVal = (key) => {
                    const k = Object.keys(r).find(k => k.trim().toLowerCase() === key.toLowerCase());
                    return k ? r[k] : "";
                };

                return {
                    date: getVal("Date"),
                    month: getVal("Month"),
                    customer: getVal("Customer"),
                    product: getVal("Product"),
                    qty: parseNum(getVal("Qty")),
                    amount: parseNum(getVal("Amount"))
                };
            }).filter(item => item.customer || item.product || item.amount); // Clean out empty rows

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
        tbody.innerHTML = "<tr><td colspan='6' class='text-center'>No Data Available</td></tr>";
        return;
    }

    let html = "";
    data.forEach(r => {
        html += `
        <tr>
            <td>${r.date}</td>
            <td>${r.month}</td>
            <td>${r.customer}</td>
            <td>${r.product}</td>
            <td>${r.qty}</td>
            <td>${r.amount}</td>
        </tr>`;
    });

    tbody.innerHTML = html;
}

// ------------------------------------------
// 3. Update Summary Cards / Dashboard
// ------------------------------------------
function updateDashboard() {
    const setSafeText = (id, text) => {
        const el = document.getElementById(id);
        if (el) el.innerHTML = text;
    };

    const totalSales = salesData.reduce((a, b) => a + (b.amount || 0), 0);
    const totalQty = salesData.reduce((a, b) => a + (b.qty || 0), 0);
    const uniqueCustomers = new Set(salesData.map(x => x.customer).filter(Boolean)).size;
    const uniqueProducts = new Set(salesData.map(x => x.product).filter(Boolean)).size;

    setSafeText("totalSales", totalSales.toLocaleString());
    setSafeText("totalQty", totalQty.toLocaleString());
    setSafeText("totalCustomers", uniqueCustomers);
    setSafeText("totalProducts", uniqueProducts);
}

// ------------------------------------------
// 4. Dynamic Live Search (Matches Initials & Substrings)
// ------------------------------------------
function searchData() {
    const customerInput = document.getElementById("customerSearch");
    const monthInput = document.getElementById("monthSearch");

    const customerQuery = customerInput ? customerInput.value.trim().toLowerCase() : "";
    const monthQuery = monthInput ? monthInput.value.trim().toLowerCase() : "";

    const filtered = salesData.filter(item => {
        const custVal = String(item.customer || "").toLowerCase();
        const monthVal = String(item.month || "").toLowerCase();

        // Checks if string starts with OR contains query
        const matchesCustomer = !customerQuery || custVal.includes(customerQuery);
        const matchesMonth = !monthQuery || monthVal.includes(monthQuery);

        return matchesCustomer && matchesMonth;
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

    let csv = "Date,Month,Customer,Product,Qty,Amount\n";
    salesData.forEach(r => {
        csv += `"${r.date}","${r.month}","${r.customer}","${r.product}",${r.qty},${r.amount}\n`;
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
// 6. Automatic Event Binding & Initialization
// ------------------------------------------
document.addEventListener("DOMContentLoaded", () => {
    const customerInput = document.getElementById("customerSearch");
    const monthInput = document.getElementById("monthSearch");

    // Live search as you type
    if (customerInput) {
        customerInput.addEventListener("input", searchData);
        customerInput.addEventListener("keyup", searchData);
    }
    if (monthInput) {
        monthInput.addEventListener("input", searchData);
        monthInput.addEventListener("keyup", searchData);
    }

    // Initial blank state setup
    updateDashboard();
});
