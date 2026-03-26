const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

// Path to your SharePoint-synced Excel file
const excelPath = "C:/Users/Lynden.DeLaCruz/Cloud Direct/Network Site Info Update - Asset Register/JCL-VM-Asset-Register-Live-Ext.xlsx";

// Output CSV path inside your repo
const outputCsv = path.join(
  "C:/Users/Lynden.DeLaCruz/OneDrive - Cloud Direct/Desktop/dev/network-asset-register",
  "asset-register.csv"
);

console.log("Reading Excel:", excelPath);

// Load workbook
const workbook = xlsx.readFile(excelPath);
const sheet = workbook.Sheets[workbook.SheetNames[0]];

// Convert sheet → JSON rows
const rows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

// CSV header (must match extension)
const header = [
  "legacyName",
  "futureName",
  "wanIp",
  "internalIp",
  "deviceModel",
  "serialNumber",
  "firmware",
  "address1",
  "address2",
  "postcode",
  "pstn",
  "service",
  "status",
  "passwordLocation",
  "remoteAccess",
  "prtg",
  "syslog"
];

// Convert rows → CSV lines
const csvLines = [header.join(";")];

for (const row of rows) {
  const line = header.map(h => (row[h] || "").toString().trim()).join(";");
  csvLines.push(line);
}

// Write CSV
fs.writeFileSync(outputCsv, csvLines.join("\n"), "utf8");

console.log("CSV written to:", outputCsv);