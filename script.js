function uploadExcel() {

    const file = document.getElementById("excelFile").files[0];

    if (!file) {
        alert("Please select an Excel file.");
        return;
    }

    const reader = new FileReader();

    reader.onload = function (e) {

        const data = new Uint8Array(e.target.result);

        const workbook = XLSX.read(data, {
            type: "array"
        });

        const firstSheet = workbook.SheetNames[0];

        const worksheet = workbook.Sheets[firstSheet];

        const jsonData = XLSX.utils.sheet_to_json(worksheet);

        salesData = [];

        jsonData.forEach(row => {

            salesData.push({

                date: row.Date || "",

                month: row.Month || "",

                customer: row.Customer || "",

                product: row.Product || "",

                qty: row.Qty || 0,

                amount: row.Amount || 0

            });

        });

        loadTable(salesData);

        document.getElementById("uploadStatus").innerHTML =
            "✅ " + salesData.length + " records loaded successfully.";

        updateDashboard();

    };

    reader.readAsArrayBuffer(file);

}
function updateDashboard(){

document.getElementById("totalSales").innerHTML =
salesData.reduce((a,b)=>a+Number(b.amount),0).toLocaleString();

document.getElementById("totalQty").innerHTML =
salesData.reduce((a,b)=>a+Number(b.qty),0);

document.getElementById("totalCustomers").innerHTML =
new Set(salesData.map(x=>x.customer)).size;

document.getElementById("totalProducts").innerHTML =
new Set(salesData.map(x=>x.product)).size;

}
