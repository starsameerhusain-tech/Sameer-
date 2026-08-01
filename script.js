let salesData = [];

// =========================
// 1. Upload & Parse Excel
// =========================
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

            const parseNum = (val) => {
                if (typeof val === "number") return val;
                if (!val) return 0;
                return Number(String(val).replace(/[^0-9.-]+/g, "")) || 0;
            };

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
            }).filter(item => item.customer || item.product || item.amount);

            loadTable(salesData);
            updateDashboard();

            const statusEl = document.getElementById("uploadStatus");
            if (statusEl) {
                statusEl.innerHTML = "✅ " + salesData.length + " records uploaded successfully.";
            }
        } catch (err) {
            console.error("Error parsing Excel:", err);
            alert("Error reading file. Please check console for details.");
        }
    };

    reader.readAsArrayBuffer(file);
}

// =========================
// 2. Load Table
// =========================
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

// =========================
// 3. Update Dashboard
// =========================
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

// =========================
// 4. Dynamic Search Function
// =========================
function searchData() {
    const customerQuery = (document.getElementById("customerSearch")?.value || "").trim().toLowerCase();
    const monthQuery = (document.getElementById("monthSearch")?.value || "").trim().toLowerCase();

    const filtered = salesData.filter(item => {
        // Checks if customer/month starts with the typed letters
        const customerMatch = !customerQuery || item.customer.toLowerCase().startsWith(customerQuery);
        const monthMatch = !monthQuery || item.month.toLowerCase().startsWith(monthQuery);

        return customerMatch && monthMatch;
    });

    loadTable(filtered);
}

// Live typing event listeners
document.addEventListener("DOMContentLoaded", () => {
    const customerInput = document.getElementById("customerSearch");
    const monthInput = document.getElementById("monthSearch");

    if (customerInput) customerInput.addEventListener("input", searchData);
    if (monthInput) monthInput.addEventListener("input", searchData);
});

// Initial Dashboard Load
updateDashboard();
