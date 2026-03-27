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

// Convert sheet → JSON rows (keys = Excel headers)
const rows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

// CSV header (must match your extension)
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

// Start CSV lines
const csvLines = [header.join(";")];

// Convert Excel rows → CSV rows
for (const row of rows) {
  const line = [
    row["Site Name (Legacy Aug 2024)"],
    row["Site Name (Future Jan 2025)"],
    row["WAN IP"],
    row["Internal IP"],
    row["Device Model"],
    row["Serial Number"],
    row["Firmware version"],
    row["Address 1"],
    row["Address 2"],
    row["Postcode"],
    row["PSTN No."],
    row["Service"],
    row["status"],
    row["Password Location"],
    row["Remote Access FGT VPN"],
    row["PRTG Monitoring"],
    row["Syslog?"]
  ]
    .map(v => (v || "").toString().trim())
    .join(";");

  csvLines.push(line);
}

// Write CSV
fs.writeFileSync(outputCsv, csvLines.join("\n"), "utf8");

console.log("CSV written to:", outputCsv);
