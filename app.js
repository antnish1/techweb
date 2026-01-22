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
 * STATE
 ***************************************************/
let filteredRows = [];
let processedMap = {};
let activeRow = null;

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
 * FETCH MAIN DATA (FULL FILTER)
 ***************************************************/
fetch(URL)
  .then(r => r.text())
  .then(t => {
    const rows = JSON.parse(t.substring(47).slice(0, -2)).table.rows;

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

      const match =
        createDate >= oneMonthAgo &&
        installDate >= oneYearAgo &&
        clean(c[5]?.v) === "3. Fix" &&
        clean(c[18]?.v) === "OnSite" &&
        clean(c[19]?.v) === "Warranty" &&
        clean(c[20]?.v) === "Failure" &&
        clean(c[27]?.v) !== "Service";

      if (match) filteredRows.push(c);
    });

    document.getElementById("count").innerText = filteredRows.length;
    document.getElementById("viewBtn").disabled = false;
  });

/***************************************************
 * FETCH PROCESSED DATA
 ***************************************************/
fetch(PROCESSED_URL)
  .then(r => r.text())
  .then(t => {
    const rows = JSON.parse(t.substring(47).slice(0, -2)).table.rows;
    rows.forEach(r => {
      const c = r.c;
      if (!c || !c[0]?.v) return;
      processedMap[clean(c[0].v)] = {
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
document.getElementById("viewBtn").onclick = () => {
  const tbody = document.querySelector("#dataTable tbody");
  tbody.innerHTML = "";

  filteredRows.forEach(c => {
    const callId = clean(c[0]?.v);
    const done = processedMap[callId];

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${callId}</td>
      <td>${formatDate(parseDate(c[1]))}</td>
      <td>${clean(c[4]?.v)}</td>
      <td>${clean(c[6]?.v)}</td>
      <td>${clean(c[9]?.v)}</td>
      <td>${clean(c[10]?.v)}</td>
      <td>${clean(c[11]?.v)}</td>
      <td>${clean(c[24]?.v)}</td>
      <td class="${done ? "status-done" : "status-pending"}">
        ${done ? "Completed" : "Pending"}
      </td>
      <td>
        <button onclick="openModal('${callId}')">
          ${done ? "Copy" : "Process"}
        </button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  document.getElementById("dataTable").hidden = false;
};

/***************************************************
 * MODAL LOGIC
 ***************************************************/
function openModal(callId) {
  activeRow = filteredRows.find(r => clean(r[0]?.v) === callId);
  const done = processedMap[callId];

  document.getElementById("engineInput").value = done?.engineNo || "";
  document.getElementById("failedPartNameInput").value = done?.failedPartName || "";
  document.getElementById("failedPartNoInput").value = done?.failedPartNo || "";
  document.getElementById("actionRequiredInput").value = done?.actionRequired || "";

  generateCopyText(done);
  document.getElementById("copyModal").hidden = false;
}

function generateCopyText(done) {
  const c = activeRow;
  document.getElementById("copyText").value = `
Call ID : ${clean(c[0]?.v)}
Customer : ${clean(c[6]?.v)}
Machine No : ${clean(c[9]?.v)}
Engine No : ${done?.engineNo || "__________"}
Model : ${clean(c[10]?.v)}
HMR : ${clean(c[11]?.v)}
Failure Date : ${formatDate(parseDate(c[1]))}
Failed Part : ${done?.failedPartName || "__________"}
Action : ${done?.actionRequired || "__________"}
`.trim();
}

/***************************************************
 * SAVE + COPY
 ***************************************************/
function copyToClipboard() {
  const ta = document.getElementById("copyText");
  ta.select();
  document.execCommand("copy");
  alert("Copied ✔️");
}

function saveCompleted() {
  const callId = clean(activeRow[0]?.v);

  const payload = {
    callId,
    engineNo: engineInput.value,
    failedPartName: failedPartNameInput.value,
    failedPartNo: failedPartNoInput.value,
    actionRequired: actionRequiredInput.value
  };

  fetch(SAVE_URL, {
    method: "POST",
    body: JSON.stringify(payload)
  }).then(() => {
    processedMap[callId] = payload;
    closeModal();
    alert("Saved ✔️");
  });
}

function closeModal() {
  document.getElementById("copyModal").hidden = true;
}
