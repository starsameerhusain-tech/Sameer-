// =========================
// Sales Management System
// script.js
// =========================

let salesData = [];

// Upload Excel
function uploadExcel() {

    const file = document.getElementById("excelFile").files[0];

    if (!file) {
        alert("Please select an Excel file.");
        return;
    }

    const reader = new FileReader();

    reader.onload = function (e) {

        const data = new Uint8Array(e.target.result);

        const workbook = XLSX.read(data, { type: "array" });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        const jsonData = XLSX.utils.sheet_to_json(sheet);

        salesData = jsonData.map(r => ({
            date: r.Date || "",
            month: r.Month || "",
            customer: r.Customer || "",
            product: r.Product || "",
            qty: Number(r.Qty || 0),
            amount: Number(r.Amount || 0)
        }));

        loadTable(salesData);

        updateDashboard();

        document.getElementById("uploadStatus").innerHTML =
            "✅ " + salesData.length + " records uploaded successfully.";

    };

    reader.readAsArrayBuffer(file);

}

// Load Table
function loadTable(data){

    const tbody = document.getElementById("salesTable");

    tbody.innerHTML="";

    if(data.length==0){

        tbody.innerHTML=
        "<tr><td colspan='6' class='text-center'>No Data Available</td></tr>";

        return;

    }

    data.forEach(r=>{

        tbody.innerHTML+=`
        <tr>
            <td>${r.date}</td>
            <td>${r.month}</td>
            <td>${r.customer}</td>
            <td>${r.product}</td>
            <td>${r.qty}</td>
            <td>${r.amount}</td>
        </tr>`;

    });

}

// Dashboard
function updateDashboard(){

document.getElementById("totalSales").innerHTML =
salesData.reduce((a,b)=>a+b.amount,0).toLocaleString();

document.getElementById("totalQty").innerHTML =
salesData.reduce((a,b)=>a+b.qty,0);

document.getElementById("totalCustomers").innerHTML =
new Set(salesData.map(x=>x.customer)).size;

document.getElementById("totalProducts").innerHTML =
new Set(salesData.map(x=>x.product)).size;

}

// Search
function searchData(){

const customer=document.getElementById("customerSearch").value.toLowerCase();

const month=document.getElementById("monthSearch").value.toLowerCase();

const filtered=salesData.filter(x=>

x.customer.toLowerCase().includes(customer)

&&

x.month.toLowerCase().includes(month)

);

loadTable(filtered);

}

// Download CSV
function downloadCSV(){

if(salesData.length==0){

alert("No data available");

return;

}

let csv="Date,Month,Customer,Product,Qty,Amount\n";

salesData.forEach(r=>{

csv+=`${r.date},${r.month},${r.customer},${r.product},${r.qty},${r.amount}\n`;

});

const blob=new Blob([csv],{type:"text/csv"});

const url=URL.createObjectURL(blob);

const a=document.createElement("a");

a.href=url;

a.download="SalesData.csv";

a.click();

URL.revokeObjectURL(url);

}

// Download Excel
function downloadExcel(){

alert("Excel download will be added in next version.");

}

// Load Empty Dashboard
updateDashboard();
