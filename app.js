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
let completedRows = [];
let currentView = "pending"; // or "completed"
let activeCompletedRow = null;

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

      const callId = clean(c[0].v);

      const rowData = {
        callId,
        createDate: c[1]?.v,
        crmCallNo: c[2]?.v,
        subject: c[3]?.v,
        status: c[4]?.v,
        customer: c[5]?.v,
        mobile: c[6]?.v,
        machineNumber: c[7]?.v,
        machineModel: c[8]?.v,
        hmr: c[9]?.v,
        installDate: c[10]?.v,
        callType: c[11]?.v,
        callSubType: c[12]?.v,
        branch: c[13]?.v,
        city: c[14]?.v,
        serviceEngg: c[15]?.v,
        machineStatus: c[16]?.v,
        engineNo: c[17]?.v,
        failedPartName: c[18]?.v,
        failedPartNo: c[19]?.v,
        actionRequired: c[20]?.v,
        techwebNo // 👈 NEW
      };

      processedMap[callId] = rowData;
      completedRows.push(rowData);
    });

    document.getElementById("completedCount").innerText = completedRows.length;
  });

;

function renderPendingTable() {
  const tbody = document.querySelector("#dataTable tbody");
  tbody.innerHTML = "";

  filteredRows.forEach(c => {
    const callId = clean(c[0]?.v);
    const isCompleted = !!processedMap[callId];

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
      <td class="${isCompleted ? "status-done" : "status-pending"}">
        ${isCompleted ? "Completed" : "Pending"}
      </td>
      <td>
        <button onclick="openModal('${callId}')">
          ${isCompleted ? "Copy" : "Process"}
        </button>
      </td>
    `;

    tbody.appendChild(tr);
  });

  document.getElementById("dataTable").hidden = false;
}


/***************************************************
 * MODAL LOGIC
 ***************************************************/
function openModal(callId) {
  activeRow = filteredRows.find(r => clean(r[0]?.v) === callId);
  if (!activeRow) return;

  const saved = processedMap[callId];

  if (saved) {
    // 🟢 COMPLETED ROW → LOAD SAVED DATA
    engineInput.value = saved.engineNo || "";
    failedPartNameInput.value = saved.failedPartName || "";
    failedPartNoInput.value = saved.failedPartNo || "";
    actionRequiredInput.value = saved.actionRequired || "";
  } else {
    // 🟡 PENDING ROW → RESET INPUTS
    resetModalInputs();
  }

  refreshCopyText();
  document.getElementById("copyModal").hidden = false;
}


function generateCopyText(doneData = {}) {
  const c = activeRow;

  document.getElementById("copyText").value = `
Is M/C Covered Under JCB Care / Engine Care / Warranty : U/W
Call ID : ${clean(c[0]?.v)}
Customer Name : ${clean(c[6]?.v)}
Machine SL No. : ${clean(c[9]?.v)}
Engine No : ${doneData.engineNo || "__________"}
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
Failed Part Name : ${doneData.failedPartName || "__________"}
Failed Part No. : ${doneData.failedPartNo || "__________"}
Action Required : ${doneData.actionRequired || "__________"}
`.trim();
}



function refreshCopyText() {
  if (!activeRow) return;

  const c = activeRow;

  const data = {
    engineNo: engineInput.value.trim(),
    failedPartName: failedPartNameInput.value.trim(),
    failedPartNo: failedPartNoInput.value.trim(),
    actionRequired: actionRequiredInput.value.trim()
  };

  document.getElementById("copyText").value = `
Is M/C Covered Under JCB Care / Engine Care / Warranty : U/W
Call ID : ${clean(c[0]?.v)}
Customer Name : ${clean(c[6]?.v)}
Machine SL No. : ${clean(c[9]?.v)}
Engine No : ${data.engineNo || "__________"}
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
Failed Part Name : ${data.failedPartName || "__________"}
Failed Part No. : ${data.failedPartNo || "__________"}
Action Required : ${data.actionRequired || "__________"}
`.trim();
}


/***************************************************
 * SAVE + COPY
 ***************************************************/
function copyToClipboard() {
  refreshCopyText(); // ✅ ensure latest data
  const ta = document.getElementById("copyText");
  ta.select();
  document.execCommand("copy");
  alert("Copied ✔️");
}


function saveCompleted() {
  if (!activeRow) {
    alert("No record selected");
    return;
  }

  const saveBtn = document.getElementById("saveBtn");
  const loader = document.getElementById("saveLoader");

  const c = activeRow;
  const callId = clean(c[0]?.v);

  const payload = {
    callId,
    createDate: formatDate(parseDate(c[1])),
    crmCallNo: clean(c[3]?.v),
    subject: clean(c[4]?.v),
    status: clean(c[5]?.v),
    customer: clean(c[6]?.v),
    mobile: clean(c[8]?.v),
    machineNumber: clean(c[9]?.v),
    machineModel: clean(c[10]?.v),
    hmr: clean(c[11]?.v),
    installDate: formatDate(parseDate(c[12])),
    callType: clean(c[18]?.v),
    callSubType: clean(c[19]?.v),
    branch: clean(c[21]?.v),
    city: clean(c[23]?.v),
    serviceEngg: clean(c[24]?.v),
    machineStatus: clean(c[27]?.v),

    engineNo: engineInput.value.trim(),
    failedPartName: failedPartNameInput.value.trim(),
    failedPartNo: failedPartNoInput.value.trim(),
    actionRequired: actionRequiredInput.value.trim()
  };

  // 🔒 Validation
  if (
    !payload.engineNo ||
    !payload.failedPartName ||
    !payload.failedPartNo ||
    !payload.actionRequired
  ) {
    alert("Please fill ALL mandatory fields before saving.");
    return;
  }

  // 🔄 Show loader + disable buttons
  loader.hidden = false;
  saveBtn.disabled = true;

  fetch(SAVE_URL, {
    method: "POST",
    body: JSON.stringify(payload)
  })
    .then(res => res.text())
    .then(() => {
      processedMap[callId] = payload;
      document.getElementById("viewBtn").click();
      closeModal();
      alert("Full details saved ✔️");
    })
    .catch(err => {
      console.error(err);
      alert("Failed to save data ❌");
    })
    .finally(() => {
      loader.hidden = true;
      saveBtn.disabled = false;
    });
}

function resetModalInputs() {
  engineInput.value = "";
  failedPartNameInput.value = "";
  failedPartNoInput.value = "";
  actionRequiredInput.value = "";
}


function closeModal() {
  resetModalInputs();
  document.getElementById("copyModal").hidden = true;
}



function renderCompletedTable() {
  const tbody = document.querySelector("#dataTable tbody");
  tbody.innerHTML = "";

  completedRows.forEach(r => {
    const isTWDone = !!r.techwebNo;

    const tr = document.createElement("tr");
    if (isTWDone) tr.classList.add("tw-done-row");

    tr.innerHTML = `
      <td>${r.callId}</td>
      <td>${r.createDate || ""}</td>
      <td>${r.subject || ""}</td>
      <td>${r.customer || ""}</td>
      <td>${r.machineNumber || ""}</td>
      <td>${r.machineModel || ""}</td>
      <td>${r.hmr || ""}</td>
      <td>${r.serviceEngg || ""}</td>
      <td class="status-done">Completed</td>
      <td>
        ${
          isTWDone
            ? `<button disabled>TW DONE ✔</button>`
            : `<button onclick="openTWModal('${r.callId}')">TW DONE</button>`
        }
      </td>
    `;

    tbody.appendChild(tr);
  });

  document.getElementById("dataTable").hidden = false;
}


function openCompletedOnly(callId) {
  const saved = processedMap[callId];
  if (!saved) return;

  activeRow = null; // no OCR row here

  engineInput.value = saved.engineNo;
  failedPartNameInput.value = saved.failedPartName;
  failedPartNoInput.value = saved.failedPartNo;
  actionRequiredInput.value = saved.actionRequired;

  refreshCopyTextCompleted(saved);

  document.getElementById("copyModal").hidden = false;
}


function refreshCopyTextCompleted(d) {
  document.getElementById("copyText").value = `
Is M/C Covered Under JCB Care / Engine Care / Warranty : U/W
Call ID : ${d.callId}
Customer Name : ${d.customer}
Machine SL No. : ${d.machineNumber}
Engine No : ${d.engineNo}
M/C Model : ${d.machineModel}
HMR : ${d.hmr}
Date of Installation : ${d.installDate}
Date of Failure : ${d.createDate}
M/C Location : ${d.branch}
Engineer Name : ${d.serviceEngg}
Failed Part Name : ${d.failedPartName}
Failed Part No. : ${d.failedPartNo}
Action Required : ${d.actionRequired}
`.trim();
}

function openTWModal(callId) {
  activeCompletedRow = completedRows.find(r => r.callId === callId);
  if (!activeCompletedRow) return;

  document.getElementById("techwebInput").value = "";
  document.getElementById("twModal").hidden = false;
}

function closeTWModal() {
  document.getElementById("twModal").hidden = true;
  activeCompletedRow = null;
}


function saveTWDone() {
  const input = document.getElementById("techwebInput");
  const techwebNo = input.value.trim();

  if (!techwebNo) {
    alert("Please enter Techweb Number");
    return;
  }

  const payload = {
    callId: activeCompletedRow.callId,
    techwebNo
  };

  fetch(SAVE_URL, {
    method: "POST",
    body: JSON.stringify({
      type: "TW_DONE",
      ...payload
    })
  })
    .then(r => r.text())
    .then(() => {
      // update local state
      activeCompletedRow.techwebNo = techwebNo;

      closeTWModal();
      renderCompletedTable();

      alert("Marked as TW DONE ✔️");
    })
    .catch(() => alert("Failed to save Techweb number"));
}


document.addEventListener("DOMContentLoaded", () => {
  const viewBtn = document.getElementById("viewBtn");
  const viewCompletedBtn = document.getElementById("viewCompletedBtn");

  if (viewBtn) {
    viewBtn.onclick = () => {
      currentView = "pending";
      renderPendingTable();
    };
  }

  if (viewCompletedBtn) {
    viewCompletedBtn.onclick = () => {
      currentView = "completed";
      renderCompletedTable();
    };
  }
});
