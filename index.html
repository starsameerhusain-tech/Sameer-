let salesData = [];

document.addEventListener("DOMContentLoaded", () => {
    loadSavedData();
    document.getElementById("uploadBtn").addEventListener("click", handleUpload);
});

// Fetch saved data from localStorage on initialization
function loadSavedData() {
    const data = localStorage.getItem("salesData");
    if (data) {
        try {
            salesData = JSON.parse(data);
            renderData(salesData);
        } catch (e) {
            console.error("Error parsing stored sales data", e);
        }
    }
}

// Upload & Parse Excel/CSV
function handleUpload() {
    const fileInput = document.getElementById("excelFile");
    const file = fileInput.files[0];
    if (!file) {
        alert("Please select an Excel or CSV file first!");
        return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            salesData = XLSX.utils.sheet_to_json(worksheet);
            localStorage.setItem("salesData", JSON.stringify(salesData));
            renderData(salesData);
            
            document.getElementById("uploadStatus").innerHTML = 
                `<span class="text-success fw-bold mt-2 d-block"><i class="fa-solid fa-check me-1"></i> File uploaded and rendered successfully!</span>`;
        } catch (err) {
            alert("Error reading file: " + err.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

// Case-insensitive key retriever for flexible Excel column headers
function getProp(obj, keyName) {
    if (!obj) return "";
    const norm = (str) => String(str).replace(/[\s_]/g, '').toLowerCase();
    const key = Object.keys(obj).find(k => norm(k) === norm(keyName));
    return key ? obj[key] : "";
}

// Case-insensitive key setter
function setProp(obj, keyName, val) {
    if (!obj) return;
    const norm = (str) => String(str).replace(/[\s_]/g, '').toLowerCase();
    const key = Object.keys(obj).find(k => norm(k) === norm(keyName));
    if (key) {
        obj[key] = val;
    } else {
        obj[keyName] = val;
    }
}

// Render Data Table
function renderData(data) {
    const tbody = document.getElementById("salesTableBody");
    tbody.innerHTML = "";

    if (!data || data.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="11" class="text-center py-5 text-muted">
                    <i class="fa-regular fa-folder-open fs-2 d-block mb-2 text-secondary"></i>
                    No Data Available
                </td>
            </tr>`;
        updateKPIs([]);
        return;
    }

    data.forEach((item, index) => {
        const rate = Number(getProp(item, "rate")) || 0;
        const hours = Number(getProp(item, "hours")) || 0;
        const amount = Number(getProp(item, "amount")) || (rate * hours);

        const tr = document.createElement("tr");
        tr.id = `row-${index}`;
        tr.innerHTML = `
            <td>${getProp(item, "month") || "-"}</td>
            <td>${getProp(item, "project name") || getProp(item, "project") || "-"}</td>
            <td>${getProp(item, "email id") || getProp(item, "email") || "-"}</td>
            <td>${getProp(item, "name") || getProp(item, "client name") || "-"}</td>
            <td>${getProp(item, "source") || "-"}</td>
            <td>${getProp(item, "language") || "-"}</td>
            <td>$${rate}</td>
            <td>${hours} hrs</td>
            <td class="fw-bold text-success">$${amount}</td>
            <td>${getProp(item, "project code") || getProp(item, "code") || "-"}</td>
            <td class="text-center">
                <button class="btn btn-sm btn-outline-primary py-1 px-3 border-0 rounded-2 fw-semibold" onclick="enableEdit(${index})">
                    <i class="fa-solid fa-pen-to-square me-1"></i>Edit
                </button>
            </td>
        `;
        tbody.appendChild(tr);
    });

    updateKPIs(data);
}

// Enable Inline Row Editing
function enableEdit(index) {
    const item = salesData[index];
    const tr = document.getElementById(`row-${index}`);

    const rate = Number(getProp(item, "rate")) || 0;
    const hours = Number(getProp(item, "hours")) || 0;
    const amount = Number(getProp(item, "amount")) || (rate * hours);

    tr.innerHTML = `
        <td><input type="text" class="edit-input" id="edit-month-${index}" value="${getProp(item, "month")}"></td>
        <td><input type="text" class="edit-input" id="edit-project-${index}" value="${getProp(item, "project name") || getProp(item, "project")}"></td>
        <td><input type="text" class="edit-input" id="edit-email-${index}" value="${getProp(item, "email id") || getProp(item, "email")}"></td>
        <td><input type="text" class="edit-input" id="edit-name-${index}" value="${getProp(item, "name") || getProp(item, "client name")}"></td>
        <td><input type="text" class="edit-input" id="edit-source-${index}" value="${getProp(item, "source")}"></td>
        <td><input type="text" class="edit-input" id="edit-language-${index}" value="${getProp(item, "language")}"></td>
        <td><input type="number" class="edit-input" id="edit-rate-${index}" value="${rate}" oninput="calcEditAmount(${index})"></td>
        <td><input type="number" class="edit-input" id="edit-hours-${index}" value="${hours}" oninput="calcEditAmount(${index})"></td>
        <td><input type="number" class="edit-input fw-bold" id="edit-amount-${index}" value="${amount}" readonly></td>
        <td><input type="text" class="edit-input" id="edit-code-${index}" value="${getProp(item, "project code") || getProp(item, "code")}"></td>
        <td class="text-center">
            <button class="btn btn-sm btn-success py-1 px-2 me-1 rounded-2" title="Save" onclick="saveEdit(${index})">
                <i class="fa-solid fa-check"></i>
            </button>
            <button class="btn btn-sm btn-secondary py-1 px-2 rounded-2" title="Cancel" onclick="renderData(salesData)">
                <i class="fa-solid fa-xmark"></i>
            </button>
        </td>
    `;
}

// Automatically recalculate Amount on Rate/Hours input
function calcEditAmount(index) {
    const r = Number(document.getElementById(`edit-rate-${index}`).value) || 0;
    const h = Number(document.getElementById(`edit-hours-${index}`).value) || 0;
    document.getElementById(`edit-amount-${index}`).value = r * h;
}

// Save edited row data back to main array & localStorage
function saveEdit(index) {
    const item = salesData[index];

    const rate = Number(document.getElementById(`edit-rate-${index}`).value) || 0;
    const hours = Number(document.getElementById(`edit-hours-${index}`).value) || 0;

    setProp(item, "month", document.getElementById(`edit-month-${index}`).value);
    setProp(item, "project name", document.getElementById(`edit-project-${index}`).value);
    setProp(item, "email id", document.getElementById(`edit-email-${index}`).value);
    setProp(item, "name", document.getElementById(`edit-name-${index}`).value);
    setProp(item, "source", document.getElementById(`edit-source-${index}`).value);
    setProp(item, "language", document.getElementById(`edit-language-${index}`).value);
    setProp(item, "rate", rate);
    setProp(item, "hours", hours);
    setProp(item, "amount", rate * hours);
    setProp(item, "project code", document.getElementById(`edit-code-${index}`).value);

    salesData[index] = item;
    localStorage.setItem("salesData", JSON.stringify(salesData));
    renderData(salesData);
}

// Update Top Dashboard KPIs
function updateKPIs(data) {
    let totalAmt = 0;
    let totalHrs = 0;
    let rateSum = 0;

    data.forEach(item => {
        const r = Number(getProp(item, "rate")) || 0;
        const h = Number(getProp(item, "hours")) || 0;
        const a = Number(getProp(item, "amount")) || (r * h);

        totalAmt += a;
        totalHrs += h;
        rateSum += r;
    });

    document.getElementById("totalAmount").innerText = `$${totalAmt.toLocaleString()}`;
    document.getElementById("totalRecords").innerText = data.length.toLocaleString();
    document.getElementById("totalHours").innerText = `${totalHrs.toLocaleString()} hrs`;
    document.getElementById("avgRate").innerText = `$${data.length > 0 ? Math.round(rateSum / data.length) : 0}/hr`;
}

// Search and Filter functionality
function handleSearch() {
    const nameQuery = document.getElementById("searchName").value.toLowerCase();
    const emailQuery = document.getElementById("searchEmail").value.toLowerCase();
    const monthQuery = document.getElementById("searchMonth").value.toLowerCase();

    if (!nameQuery && !emailQuery && !monthQuery) {
        renderData(salesData);
        return;
    }

    const filtered = salesData.filter(item => {
        const name = String(getProp(item, "name") || getProp(item, "client name")).toLowerCase();
        const email = String(getProp(item, "email id") || getProp(item, "email")).toLowerCase();
        const month = String(getProp(item, "month")).toLowerCase();

        return name.includes(nameQuery) && email.includes(emailQuery) && month.includes(monthQuery);
    });

    renderData(filtered);
}

// Reset filters
function resetFilters() {
    document.getElementById("searchName").value = "";
    document.getElementById("searchEmail").value = "";
    document.getElementById("searchMonth").value = "";
    renderData(salesData);
}
