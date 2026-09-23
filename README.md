<!DOCTYPE html>

<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<title>Laptop & iPad Borrowing System</title>

<!-- Excel Library -->

<script src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"></script>

<style>

body {
    font-family: Arial, sans-serif;
    background: #f4f6f8;
    margin: 0;
    padding: 30px;
}

.container {
    max-width: 1200px;
    margin: auto;
}

h1 {
    margin-bottom: 5px;
}

.subtitle {
    color: #666;
}

.card {
    background: white;
    padding: 20px;
    margin-top: 20px;
    border-radius: 10px;
    box-shadow: 0 3px 10px rgba(0,0,0,.1);
}

input, select {
    padding: 10px;
    margin: 5px;
    border: 1px solid #ccc;
    border-radius: 5px;
}

button {
    padding: 10px 15px;
    border: none;
    border-radius: 5px;
    cursor: pointer;
    margin: 5px;
}

.add {
    background: #2563eb;
    color: white;
}

.export {
    background: #16a34a;
    color: white;
}

.load {
    background: #7c3aed;
    color: white;
}

table {
    width: 100%;
    border-collapse: collapse;
    margin-top: 20px;
}

th {
    background: #1f2937;
    color: white;
    padding: 12px;
}

td {
    padding: 10px;
    border-bottom: 1px solid #ddd;
}

.returned {
    color: green;
    font-weight: bold;
}

.not-returned {
    color: red;
    font-weight: bold;
}

.status {
    margin-top: 10px;
    padding: 10px;
    background: #eef2ff;
    border-radius: 5px;
}

</style>

</head>

<body>

<div class="container">

<h1>💻 Laptop & iPad Borrowing System</h1>
<p class="subtitle">
    Records are stored in Excel
</p>

<div class="card">

```
<h3>Load Existing Excel</h3>

<input
    type="file"
    id="excelFile"
    accept=".xlsx,.xls"
>

<button class="load" onclick="loadExcel()">
    📂 Load Excel
</button>

<div id="fileStatus" class="status">
    No Excel file loaded.
</div>
```

</div>

<div class="card">

```
<h3>Add Borrow Record</h3>

<input
    type="text"
    id="device"
    placeholder="Device Name"
>

<select id="type">
    <option value="Laptop">Laptop</option>
    <option value="iPad">iPad</option>
</select>

<input
    type="text"
    id="serial"
    placeholder="Serial Number"
>

<input
    type="text"
    id="borrower"
    placeholder="Borrower Name"
>

<input
    type="date"
    id="dateBorrowed"
>

<button class="add" onclick="addRecord()">
    ➕ Add Record
</button>
```

</div>

<div class="card">

```
<h3>Records</h3>

<button class="export" onclick="exportExcel()">
    📥 Save / Export to Excel
</button>

<table>

    <thead>

        <tr>
            <th>ID</th>
            <th>Device</th>
            <th>Type</th>
            <th>Serial Number</th>
            <th>Borrower</th>
            <th>Date Borrowed</th>
            <th>Status</th>
            <th>Date Returned</th>
            <th>Action</th>
        </tr>

    </thead>

    <tbody id="recordTable">
    </tbody>

</table>
```

</div>

</div>

<script>

let records = [];


/* =========================
   LOAD EXCEL
========================= */

function loadExcel() {

    const file =
        document.getElementById("excelFile").files[0];

    if (!file) {

        alert("Please select your records.xlsx file.");

        return;
    }

    const reader = new FileReader();

    reader.onload = function(e) {

        try {

            const data = new Uint8Array(e.target.result);

            const workbook =
                XLSX.read(data, { type: "array" });

            const sheetName =
                workbook.SheetNames[0];

            const worksheet =
                workbook.Sheets[sheetName];

            records =
                XLSX.utils.sheet_to_json(worksheet);

            displayRecords();

            document.getElementById("fileStatus").innerHTML =
                "✅ Excel loaded: " + file.name;

        }

        catch(error) {

            alert("Unable to read Excel file.");

            console.error(error);
        }
    };

    reader.readAsArrayBuffer(file);
}


/* =========================
   DISPLAY RECORDS
========================= */

function displayRecords() {

    const table =
        document.getElementById("recordTable");

    table.innerHTML = "";

    records.forEach((record, index) => {

        const row =
            document.createElement("tr");

        const status =
            record.Status || "Not Return";

        const statusClass =
            status === "Return"
            ? "returned"
            : "not-returned";

        row.innerHTML = `

            <td>${record.ID || index + 1}</td>

            <td>${record.Device || ""}</td>

            <td>${record.Type || ""}</td>

            <td>${record["Serial Number"] || ""}</td>

            <td>${record.Borrower || ""}</td>

            <td>${formatDate(record["Date Borrowed"])}</td>

            <td class="${statusClass}">
                ${status}
            </td>

            <td>
                ${formatDate(record["Date Returned"])}
            </td>

            <td>

                ${
                    status !== "Return"
                    ?
                    `<button
                        onclick="returnDevice(${index})"
                        style="background:#16a34a;color:white">
                        Return
                    </button>`
                    :
                    "✅ Returned"
                }

            </td>
        `;

        table.appendChild(row);

    });
}


/* =========================
   ADD RECORD
========================= */

function addRecord() {

    const device =
        document.getElementById("device").value;

    const type =
        document.getElementById("type").value;

    const serial =
        document.getElementById("serial").value;

    const borrower =
        document.getElementById("borrower").value;

    const dateBorrowed =
        document.getElementById("dateBorrowed").value;


    if (
        !device ||
        !serial ||
        !borrower ||
        !dateBorrowed
    ) {

        alert("Please complete all fields.");

        return;
    }


    const newRecord = {

        ID: records.length + 1,

        Device: device,

        Type: type,

        "Serial Number": serial,

        Borrower: borrower,

        "Date Borrowed": dateBorrowed,

        Status: "Not Return",

        "Date Returned": ""

    };


    records.push(newRecord);

    displayRecords();


    // Clear form

    document.getElementById("device").value = "";

    document.getElementById("serial").value = "";

    document.getElementById("borrower").value = "";

    document.getElementById("dateBorrowed").value = "";

}


/* =========================
   RETURN DEVICE
========================= */

function returnDevice(index) {

    const confirmReturn =
        confirm("Mark this device as RETURNED?");

    if (!confirmReturn) {
        return;
    }


    records[index].Status = "Return";

    records[index]["Date Returned"] =
        new Date().toISOString().split("T")[0];


    displayRecords();

}


/* =========================
   EXPORT EXCEL
========================= */

function exportExcel() {

    if (records.length === 0) {

        alert("There are no records to save.");

        return;
    }


    const worksheet =
        XLSX.utils.json_to_sheet(records);


    const workbook =
        XLSX.utils.book_new();


    XLSX.utils.book_append_sheet(
        workbook,
        worksheet,
        "Records"
    );


    XLSX.writeFile(
        workbook,
        "records.xlsx"
    );


    alert(
        "✅ Excel file saved successfully!"
    );
}


/* =========================
   DATE FORMAT
========================= */

function formatDate(value) {

    if (!value) {
        return "";
    }

    return value;

}

</script>

</body>
</html>
