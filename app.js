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
  const done = processedMap[callId] || {};

  engineInput.value = done.engineNo || "";
  failedPartNameInput.value = done.failedPartName || "";
  failedPartNoInput.value = done.failedPartNo || "";
  actionRequiredInput.value = done.actionRequired || "";

  generateCopyText(done);
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
  if (!activeRow) {
    alert("No record selected");
    return;
  }

  const c = activeRow;
  const callId = clean(c[0]?.v);

  if (!callId) {
    alert("Invalid Call ID");
    return;
  }

  const payload = {
    // 🔹 ORIGINAL CALL DATA (FULL DETAILS)
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

    // 🔹 USER INPUT (MANDATORY)
    engineNo: engineInput.value.trim(),
    failedPartName: failedPartNameInput.value.trim(),
    failedPartNo: failedPartNoInput.value.trim(),
    actionRequired: actionRequiredInput.value.trim()
  };

  // 🔒 VALIDATION – ALL INPUTS REQUIRED
  if (
    !payload.engineNo ||
    !payload.failedPartName ||
    !payload.failedPartNo ||
    !payload.actionRequired
  ) {
    alert("Please fill ALL mandatory fields before saving.");
    return;
  }

  fetch(SAVE_URL, {
    method: "POST",
    body: JSON.stringify(payload)
  })
    .then(res => res.text())
    .then(() => {
      // Mark as completed locally
      processedMap[callId] = payload;

      // Refresh table UI
      document.getElementById("viewBtn").click();

      closeModal();
      alert("Full details saved to OCR_PROCESSED ✔️");
    })
    .catch(err => {
      console.error(err);
      alert("Failed to save data ❌");
    });
}


function closeModal() {
  document.getElementById("copyModal").hidden = true;
}
