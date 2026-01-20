// 🔴 CHANGE THIS
const SHEET_ID = "15S8wwS4cbi8dFEvpoeIw0EDD1WMZceub5IpmtF5SFes";
const SHEET_NAME = "OCR";

const URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${SHEET_NAME}`;

let filteredRows = [];

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

      if (!c[1] || !c[12]) return;

      const createDate = new Date(c[1].f);
      const installDate = new Date(c[12].f);

      const isValid =
        createDate >= oneMonthAgo &&
        installDate >= oneYearAgo &&
        c[4]?.v === "3. Fix" &&
        c[19]?.v === "OnSite" &&
        c[20]?.v === "Warranty" &&
        c[21]?.v === "Failure" &&
        c[27]?.v !== "Service";

      if (isValid) filteredRows.push(c);
    });

    document.getElementById("count").innerText = filteredRows.length;
    document.getElementById("viewBtn").disabled = false;
  });

document.getElementById("viewBtn").addEventListener("click", () => {
  const table = document.getElementById("dataTable");
  const tbody = table.querySelector("tbody");
  tbody.innerHTML = "";

  filteredRows.forEach(c => {
    const row = document.createElement("tr");

    const cols = [
      0, 1, 3, 4, 5, 6, 8, 9, 10, 11,
      12, 20, 21, 22, 24, 25, 27
    ];

    cols.forEach(i => {
      const td = document.createElement("td");
      td.innerText = c[i]?.f || c[i]?.v || "";
      row.appendChild(td);
    });

    tbody.appendChild(row);
  });

  table.hidden = false;
});
