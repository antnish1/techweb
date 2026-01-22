/***************************************************
 * CONFIG
 ***************************************************/
const SHEET_ID = "15S8wwS4cbi8dFEvpoeIw0EDD1WMZceub5IpmtF5SFes";
const SHEET_NAME = "OCR";
const PROCESSED_SHEET_NAME = "OCR_PROCESSED";

const URL =
  `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${SHEET_NAME}`;

const PROCESSED_URL =
  `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${PROCESSED_SHEET_NAME}`;

const SAVE_URL =
  "https://script.google.com/macros/s/AKfycby4xnVtuSZek7VigeWZI_41-IXfO99xUrNrbeKm31T2pHjbL8LLtvvoj3qklFHlYq1E/exec";

/***************************************************
 * GLOBAL STATE
 ***************************************************/
let filteredRows = [];
let processedMap = {}; // callId -> completed data

/***************************************************
 * HELPERS
 ***************************************************/
function clean(v) {
  return v ? v.toString().trim() : "";
}

function parseDate(cell) {
  if (!cell) return null;

  if (cell.v instanceof Date) return cell.v;

  if (cell.f) {
    const d = new Date(cell.f);
    return isNaN(d) ? null : d;
  }

  const d = new Date(cell.v);
  return isNaN(d) ? null : d;
}

function formatDate(date) {
  if (!date) return "";
  const d = String(date.getDate()).padStart(2, "0");
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const y = String(date.getFullYear()).slice(-2);
  return `${d}/${m}/${y}`;
}

/***************************************************
 * FETCH OCR DATA + APPLY FULL FILTERING
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

      const createDate = parseDate(c[1]);
      const installDate = parseDate(c[12]);

      if (!createDate || !installDate) return;

      const isMatch =
        createDate >= oneMonthAgo &&
        installDate >= oneYearAgo &&
        clean(c[5]?.v) === "3. Fix" &&
        clean(c[18]?.v) === "OnSite" &&
        clean(c[19]?.v) === "Warranty" &&
        clean(c[20]?.v) === "Failure" &&
        clean(c[27]?.v) !== "Service";

      if (isMatch) {
        filteredRows.push(c);
      }
    });

    document.getElementById("count").innerText = filteredRows.length;
    document.getElementById("viewBtn").disabled = false;
  });

/***************************************************
 * FETCH PROCESSED DATA (SOURCE OF TRUTH)
 ***************************************************/
fetch(PROCESSED_URL)
  .then(res => res.text())
  .then(text => {
    const json = JSON.parse(text.substring(47).slice(0, -2));
    const rows = json.table.rows;

    rows.forEach(r => {
      const c = r.c;
      if (!c || !c[0]?.v) return;

      const callId = clean(c[0].v);
      processedMap[callId] = {
        engineNo: clean(c[17]?.v),
        failedPartName: clean(c[18]?.v),
        failedPartNo: clean(c[19]?.v),
        actionRequired: clean(c[20]?.v)
      };
    });
  });

/***************************************************
 * RENDER TABLE
 ***************************************************/
document.getElementById("viewBtn").addEventListener("click", () => {
  const tbody = document.querySelector("#dataTable tbody");
  tbody.innerHTML = "";

  filteredRows.forEach(c => {
    const tr = document.createElement("tr");

    [
      0, 1, 4, 6, 9, 10, 11, 24
    ].forEach(i => {
      const td = document.createElement("td");
      td.textContent = c[i]?.f || c[i]?.v || "";
      tr.appendChild(td);
    });

    const callId = clean(c[0]?.v);
    const isCompleted = !!processedMap[callId];

    const statusTd = document.createElement("td");
    statusTd.textContent = isCompleted ? "Completed" : "Pending";
    statusTd.className = isCompleted ? "status-done" : "status-pending";
    tr.appendChild(statusTd);

    const actionTd = document.createElement("td");
    const btn = document.createElement("button");
    btn.textContent = isCompleted ? "Copy Completed" : "Open";
    btn.onclick = () =>
      isCompleted ? openCompletedFormat(c) : generateCopyFormat(c);
    actionTd.appendChild(btn);
    tr.appendChild(actionTd);

    tbody.appendChild(tr);
  });

  document.getElementById("dataTable").hidden = false;
});

/***************************************************
 * COPY FORMAT (PENDING)
 ***************************************************/
function generateCopyFormat(c) {
  document.getElementById("copyText").value = `
Is M/C Covered Under JCB Care / Engine Care / Warranty : U/W
Call ID : ${clean(c[0]?.v)}
Customer Name : ${clean(c[6]?.v)}
Machine SL No. : ${clean(c[9]?.v)}
Engine No : __________
M/C Model : ${clean(c[10]?.v)}
HMR : ${clean(c[11]?.v)}
Date of Installation : ${formatDate(parseDate(c[12]))}
Date of Failure : ${formatDate(parseDate(c[1]))}
M/C Location : ${clean(c[21]?.v)}
M/C Application : Material Handling
Dealership & Branch Name : FCV
Engineer Name : ${clean(c[24]?.v)}
M/C Condition : Running with problem
Nature of Complaint : ${clean(c[4]?.v)}
Failed Part Name : __________
Failed Part No. : __________
Action Required : __________
`.trim();

  document.getElementById("copyModal").hidden = false;
}

/***************************************************
 * COPY COMPLETED
 ***************************************************/
function openCompletedFormat(c) {
  const callId = clean(c[0]?.v);
  const d = processedMap[callId];

  document.getElementById("copyText").value = `
Is M/C Covered Under JCB Care / Engine Care / Warranty : U/W
Call ID : ${callId}
Customer Name : ${clean(c[6]?.v)}
Machine SL No. : ${clean(c[9]?.v)}
Engine No : ${d.engineNo}
M/C Model : ${clean(c[10]?.v)}
HMR : ${clean(c[11]?.v)}
Date of Installation : ${formatDate(parseDate(c[12]))}
Date of Failure : ${formatDate(parseDate(c[1]))}
M/C Location : ${clean(c[21]?.v)}
M/C Application : Material Handling
Dealership & Branch Name : FCV
Engineer Name : ${clean(c[24]?.v)}
M/C Condition : Running with problem
Nature of Complaint : ${clean(c[4]?.v)}
Failed Part Name : ${d.failedPartName}
Failed Part No. : ${d.failedPartNo}
Action Required : ${d.actionRequired}
`.trim();

  document.getElementById("copyModal").hidden = false;
}

/***************************************************
 * MODAL CONTROLS
 ***************************************************/
function copyToClipboard() {
  const ta = document.getElementById("copyText");
  ta.select();
  ta.setSelectionRange(0, 99999);
  document.execCommand("copy");
  alert("Copied ✔️");
}

function closeModal() {
  document.getElementById("copyModal").hidden = true;
}
