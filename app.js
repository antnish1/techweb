/***************************************************
 * CONFIGURATION
 ***************************************************/
const SHEET_ID = "15S8wwS4cbi8dFEvpoeIw0EDD1WMZceub5IpmtF5SFes";
const SHEET_NAME = "OCR";

const URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${SHEET_NAME}`;

let filteredRows = [];

/***************************************************
 * HELPERS
 ***************************************************/
function clean(value) {
  return value ? value.toString().trim() : "";
}

function parseDate(cell) {
  if (!cell) return null;

  if (cell.v instanceof Date) return cell.v;

  if (typeof cell.v === "string" && cell.v.startsWith("Date(")) {
    const p = cell.v.replace("Date(", "").replace(")", "").split(",").map(Number);
    return new Date(p[0], p[1], p[2]);
  }

  if (cell.f) {
    const d = new Date(cell.f);
    return isNaN(d) ? null : d;
  }

  return null;
}

/***************************************************
 * FETCH + FILTER
 ***************************************************/
fetch(URL)
  .then(res => res.text())
  .then(text => {
    const json = JSON.parse(text.substring(47).slice(0, -2));
    const rows = json.table.rows;

    const today = new Date();
    const oneMonthAgo = new Date(today);
    oneMonthAgo.setMonth(today.getMonth() - 1);

    const oneYearAgo = new Date(today);
    oneYearAgo.setFullYear(today.getFullYear() - 1);

    rows.forEach(r => {
      const c = r.c;
      if (!c) return;

      const createDate = parseDate(c[1]);   // Create Date
      const installDate = parseDate(c[12]); // Install Date

      if (!createDate || !installDate) return;

      const isValid =
        createDate >= oneMonthAgo &&
        installDate >= oneYearAgo &&
        clean(c[5]?.v) === "3. Fix" &&
        clean(c[18]?.v) === "OnSite" &&
        clean(c[19]?.v) === "Warranty" &&
        clean(c[20]?.v) === "Failure" &&
        clean(c[27]?.v) !== "Service";

      if (isValid) filteredRows.push(c);
    });

    document.getElementById("count").innerText = filteredRows.length;
    document.getElementById("viewBtn").disabled = false;

    console.log("MATCHING ROWS:", filteredRows.length);
  })
  .catch(err => console.error(err));

/***************************************************
 * TABLE RENDER
 ***************************************************/
document.getElementById("viewBtn").addEventListener("click", () => {
  const tbody = document.querySelector("#dataTable tbody");
  tbody.innerHTML = "";

  filteredRows.forEach(c => {
    const tr = document.createElement("tr");

    [
      0, 1, 3, 4, 5, 6, 8, 9, 10, 11,
      12, 19, 20, 21, 23, 24, 27
    ].forEach(i => {
      const td = document.createElement("td");
      td.textContent = c[i]?.f || c[i]?.v || "";
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  document.getElementById("dataTable").hidden = false;
});
