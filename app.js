let invoices = {};
const CLIENT_ID = "353757302738-9n4q0bd4aq4a4se37qusa9312pmoq6oc.apps.googleusercontent.com";

// Put your Google Drive folder ID here.
const FOLDER_ID = "1PvRCWsRP_NeoaScpdPQ94ZUYXZzD_ZvO";

let accessToken = null;

function newInvoice() {
  document.getElementById("client").value = "";
  document.getElementById("invoiceNo").value = "";
  document.getElementById("note").value = "";
  document.querySelector("#invoiceTable tbody").innerHTML = "";
  addRow();
  document.getElementById("message").textContent = "New invoice";
}

function addRow(data = {}) {
  const row = document.createElement("tr");

  row.innerHTML = `
    <td><input class="date" value="${data.date || ""}"></td>
    <td><input class="rowInvoice" value="${data.invoice || document.getElementById("invoiceNo").value}"></td>
    <td><input class="passenger" value="${data.passenger || ""}"></td>
    <td><input class="segment" value="${data.segment || ""}"></td>
    <td><input class="travelDate" value="${data.travelDate || ""}"></td>
    <td><input class="airline" value="${data.airline || ""}"></td>
    <td><input class="amount" type="number" step="0.01" value="${data.amount || ""}"></td>
    <td><button onclick="this.closest('tr').remove(); total()">X</button></td>
  `;

  row.querySelector(".amount").addEventListener("input", total);
  document.querySelector("#invoiceTable tbody").appendChild(row);
  total();
}

function total() {
  let sum = 0;
  document.querySelectorAll(".amount").forEach(function(input) {
    sum += Number(input.value) || 0;
  });
  document.getElementById("grandTotal").textContent = sum.toFixed(2);
}

function getInvoice() {
  const rows = [];

  document.querySelectorAll("#invoiceTable tbody tr").forEach(function(row) {
    rows.push({
      date: row.querySelector(".date").value,
      invoice: row.querySelector(".rowInvoice").value,
      passenger: row.querySelector(".passenger").value,
      segment: row.querySelector(".segment").value,
      travelDate: row.querySelector(".travelDate").value,
      airline: row.querySelector(".airline").value,
      amount: Number(row.querySelector(".amount").value) || 0
    });
  });

  return {
    client: document.getElementById("client").value,
    invoiceNo: document.getElementById("invoiceNo").value,
    note: document.getElementById("note").value,
    rows: rows
  };
}

async function searchInvoice() {
    const no = document.getElementById("searchNo").value.trim();
    if (!no) {
        alert("Enter invoice number.");
        return;
    }

    if (typeof accessToken === 'undefined' || !accessToken) {
        document.getElementById("message").textContent = "Error: Please connect Google Drive first.";
        return;
    }

    document.getElementById("message").textContent = "Searching Google Drive...";

    try {
        const query = `'${FOLDER_ID}' in parents and (name contains '${no}' or name = '${no}.xlsx') and trashed = false`;
        const searchUrl = `https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,name)`;
        
        const searchRes = await fetch(searchUrl, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });

        const searchData = await searchRes.json();
        if (!searchData.files || searchData.files.length === 0) {
            document.getElementById("message").textContent = "Invoice not found";
            return;
        }

        const fileId = searchData.files[0].id;
        const fileRes = await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });

        const arrayBuffer = await fileRes.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        if (data[5]) {
            document.getElementById("client").value = data[5][2] || "";
            document.getElementById("invoiceNo").value = data[5][6] || no;
        }

        for (let i = 0; i < data.length; i++) {
            if (data[i] && (data[i][0] === "NOTE" || data[i][0] === "Note")) {
                document.getElementById("note").value = data[i][1] || "";
            }
        }

        document.querySelector("#invoiceTable tbody").innerHTML = "";

        let foundRows = false;
        for (let i = 7; i < data.length; i++) {
            const row = data[i];
            if (!row || row[0] === "GRAND TOTAL" || row[5] === "GRAND TOTAL" || row[0] === "NOTE" || (!row[0] && !row[1] && !row[2] && !row[3] && !row[4] && !row[5] && !row[6])) {
                break;
            }

            addRow({
                date: row[0] !== undefined ? row[0] : "",
                invoice: row[1] !== undefined ? row[1] : "",
                passenger: row[2] !== undefined ? row[2] : "",
                segment: row[3] !== undefined ? row[3] : "",
                travelDate: row[4] !== undefined ? row[4] : "",
                airline: row[5] !== undefined ? row[5] : "",
                amount: row[6] !== undefined ? row[6] : ""
            });
            foundRows = true;
        }

        if (!foundRows) {
            addRow();
        }

        document.getElementById("message").textContent = "Invoice " + no + " opened successfully";
    } catch (error) {
        console.error("Error loading invoice from Google Drive:", error);
        document.getElementById("message").textContent = "Error: " + (error.message || JSON.stringify(error));
    }
}
async function createExcel(data) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Invoice");

  sheet.mergeCells("A1:G1");
  sheet.getCell("A1").value = "Traveling.co";
  sheet.getCell("A1").font = { size: 24, bold: true, italic: true };
  sheet.getCell("A1").alignment = { horizontal: "center" };

  sheet.mergeCells("A2:G2");
  sheet.getCell("A2").value = "Client Account Statement";
  sheet.getCell("A2").alignment = { horizontal: "center" };

  sheet.mergeCells("A3:G4");
  sheet.getCell("A3").value =
    "Office no 1, 1st Floor Plot No. 33-E Muslim Commercial Phase VI DHA Karachi\n" +
    "PHONE # 0321-2021321";

  sheet.getCell("A3").alignment = {
    horizontal: "center",
    vertical: "middle",
    wrapText: true
  };

  sheet.getCell("A6").value = "CLIENT / ACCOUNT";
  sheet.getCell("C6").value = data.client;
  sheet.getCell("F6").value = "INVOICE NO.";
  sheet.getCell("G6").value = data.invoiceNo;

  const headers = [
    "DATE", "INVOICE", "PASSENGER", "SEGMENT",
    "TRAVELING DATE", "AIRLINE", "TOTAL"
  ];

  sheet.addRow(headers);

  data.rows.forEach(function(row) {
    sheet.addRow([
      row.date,
      row.invoice,
      row.passenger,
      row.segment,
      row.travelDate,
      row.airline,
      row.amount
    ]);
  });

  const grand = data.rows.reduce(function(sum, row) {
    return sum + Number(row.amount || 0);
  }, 0);

  sheet.addRow(["", "", "", "", "", "GRAND TOTAL", grand]);
  sheet.addRow([]);
  sheet.addRow(["NOTE", data.note]);

  const widths = [14, 12, 30, 20, 18, 18, 16];

  widths.forEach(function(width, i) {
    sheet.getColumn(i + 1).width = width;
  });

  sheet.eachRow(function(row) {
    row.eachCell(function(cell) {
      cell.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" }
      };
    });
  });

  return await workbook.xlsx.writeBuffer();
}

async function downloadExcel(data) {
  const buffer = await createExcel(data);

  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = "Invoice-" + data.invoiceNo + ".xlsx";
  link.click();

  URL.revokeObjectURL(url);
}

function googleLogin() {
  if (!window.google) {
    alert("Google login is still loading. Please try again.");
    return;
  }

  const client = google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: "https://www.googleapis.com/auth/drive.file",

    callback: function(response) {
      if (response.error) {
        console.error(response);
        alert("Google authorization failed.");
        return;
      }

      accessToken = response.access_token;

      document.getElementById("driveStatus").textContent =
        "Google Drive Connected";
    }
  });

  client.requestAccessToken();
}

async function saveToGoogleDrive() {
  const data = getInvoice();

  if (!data.invoiceNo) {
    alert("Enter invoice number first.");
    return;
  }

  if (!accessToken) {
    alert("Please connect Google Drive first.");
    return;
  }

  if (FOLDER_ID === "YOUR_GOOGLE_DRIVE_FOLDER_ID") {
    alert("Add your Google Drive Folder ID in app.js first.");
    return;
  }

  try {
    document.getElementById("driveStatus").textContent =
      "Creating Excel...";

    const buffer = await createExcel(data);

    const file = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    const metadata = {
      name: "" + data.invoiceNo + ".xlsx",
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      parents: [FOLDER_ID]
    };

    const form = new FormData();

    form.append(
      "metadata",
      new Blob([JSON.stringify(metadata)], {
        type: "application/json"
      })
    );

    form.append("file", file);

    document.getElementById("driveStatus").textContent =
      "Uploading to Google Drive...";

    const response = await fetch(
      "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
      {
        method: "POST",
        headers: {
          Authorization: "Bearer " + accessToken
        },
        body: form
      }
    );

    const result = await response.json();

    if (!response.ok) {
      throw new Error(
        result.error?.message || "Google Drive upload failed"
      );
    }

    document.getElementById("driveStatus").textContent =
      "Saved to Google Drive";

    document.getElementById("message").textContent =
      "Invoice " + data.invoiceNo + " saved to Google Drive";

    alert("Invoice Excel saved to Google Drive!");
  } catch (error) {
    console.error(error);

    document.getElementById("driveStatus").textContent =
      "Upload failed";

    alert("Google Drive error: " + error.message);
  }
}

function printInvoice() {
  window.print();
}

newInvoice();
