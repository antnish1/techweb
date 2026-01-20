/***************************************************
 * CONFIG
 ***************************************************/
const SHEET_ID = "15S8wwS4cbi8dFEvpoeIw0EDD1WMZceub5IpmtF5SFes"; // 🔴 REQUIRED
const SHEET_NAME = "OCR";

const URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${SHEET_NAME}`;

let filteredRows = [];

/***************************************************
 * HELPER FUNCTIONS (PUT THEM AT TOP)
 ***************************************************/

/**
 * Cleans string values from Google Sheets
 * - trims spaces
 * - converts null/undefined safely
 */
function clean(value) {
  return value ? value.toString().trim() : "";
}

/**
 * Robust date parser for Google Sheets cells
 */
function parseDate(cell) {
  if (!cell) return null;

  // Native Date object
  if (cell.v instanceof Date) return cell.v;

  // "Date(2025,0,15)" case
  if (typeof cell.v === "string" && cell.v.startsWith("Date(")) {
    const parts = cell.v
      .replace("Date(", "")
      .replace(")", "")
      .split(",")
      .map(Number);
    return new Date(parts[0], parts[1], parts[2]);
  }

  // Formatted string (dd/mm/yyyy or similar)
  if (cell.f) {
    const d = new Date(cell.f);
    return isNaN(d) ? null : d;
  }

  return null;
}

/***************************************************
 * FETCH & PROCESS DATA
 ***************************************************/
fetch(URL)
  .then(res => res.text())
  .then(text => {
    const json = JSON.parse(text.substring(47).slice(0, -2));
    const rows = json.table.rows;

    console.log("TOTAL ROWS FOUND:", rows.length);

    const today = new Date();

    const oneMonthAgo = new Date(today);
    oneMonthAgo.setMonth(today.getMonth() - 1);

    const oneYearAgo = new Date(today);
    oneYearAgo.setFullYear(today.getFullYear() - 1);

    rows.forEach((r, index) => {
      const c = r.c;
      if (!c) return;

      const createDate = parseDate(c[1]);    // Create Date
      const installDate = parseDate(c[12]);  // Install Date

      if (!createDate || !installDate) return;

      const isValid =
        createDate >= oneMonthAgo &&
        installDate >= oneYearAgo &&
        clean(c[4]?.v) === "3. Fix" &&
        clean(c[19]?.v) === "OnSite" &&
        clean(c[20]?.v) === "Warranty" &&
        clean(c[21]?.v) === "Failure" &&
        clean(c[27]?.v) !== "Service";

      if (isValid) {
        filteredRows.push(c);
      }
    });

    // Update UI
    document.getElementById("count").innerText = filteredRows.length;
    document.getElementById("viewBtn").disabled = false;

    console.log("MATCHING ROWS:", filteredRows.length);
  })
  .catch(err => {
    console.error("SHEET FETCH ERROR:", err);
  });

/***************************************************
 * RENDER TABLE
 ***************************************************/
document.getElementById("viewBtn").addEventListener("click", () => {
  const table = document.getElementById("dataTable");
  const tbody = table.querySelector("tbody");
  tbody.innerHTML = "";

  filteredRows.forEach(c => {
    const tr = document.createElement("tr");

    // Columns you requested
    const cols = [
      0,  // Call ID
      1,  // Create Date
      3,  // CRM Call No
      4,  // Subject
      5,  // Status
      6,  // Customer
      8,  // Mobile
      9,  // Machine Number
      10, // Machine Model
      11, // HMR
      12, // Install Date
      20, // Call Type
      21, // Call Sub Type
      22, // Branch
      24, // City
      25, // Service Engg.
      27  // Machine Status
    ];

    cols.forEach(i => {
      const td = document.createElement("td");
      td.textContent = c[i]?.f || c[i]?.v || "";
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  table.hidden = false;
});
