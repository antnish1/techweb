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
let processedMap = {};

/***************************************************
 * HELPERS
 ***************************************************/
function clean(v) {
  return v ? v.toString().trim() : "";
}

function parseDate(cell) {
  if (!cell) return "";
  if (cell.f) return cell.f;
  if (cell.v instanceof Date) return cell.v.toLocaleDateString("en-GB");
  return "";
}

function renderTable() {
  document.getElementById("viewBtn").click();
}

/***************************************************
 * FETCH OCR DATA
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

      const createDate = new Date(c[1]?.f);
      const installDate = new Date(c[12]?.f);

      if (
        createDate >= oneMonthAgo &&
        installDate >= oneYearAgo &&
        clean(c[5]?.v) === "3. Fix" &&
        clean(c[18]?.v) === "OnSite" &&
        clean(c[19]?.v) === "Warranty" &&
        clean(c[20]?.v) === "Failure" &&
        clean(c[27]?.v) !== "Service"
      ) {
        filteredRows.push(c);
      }
    });

    document.getElementById("count").innerText = filteredRows.length;
    document.getElementById("viewBtn").disabled = false;
  });

/***************************************************
 * FETCH PROCESSED DATA
 ***************************************************/
fetch(PROCESSED_URL)
  .then(res => res.text())
  .then(text => {
    const json = JSON.parse(text.substring(47).slice(0, -2));
    json.table.rows.forEach(r => {
      const c = r.c;
      if (!c || !c[0]?.v) return;

      processedMap[String(c[0].v)] = {
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
    const callId = String(c[0]?.v);
    const isCompleted = !!processedMap[callId];

    const tr = document.createElement("tr");

    [
      0,1,3,4,5,6,8,9,10,11,12,19,20,21,23,24,27
    ].forEach(i => {
      const td = document.createElement("td");
      td.textContent = c[i]?.f || c[i]?.v || "";
      tr.appendChild(td);
    });

    const statusTd = document.createElement("td");
    statusTd.textContent = isCompleted ? "Completed" : "Pending";
    statusTd.style.color = isCompleted ? "green" : "orange";
    tr.appendChild(statusTd);

    const btnTd = document.createElement("td");

    if (!isCompleted) {
      const copyBtn = document.createElement("button");
      copyBtn.textContent = "Copy Format";
      copyBtn.onclick = () => generateCopyFormat(c);
      btnTd.appendChild(copyBtn);

      const doneBtn = document.createElement("button");
      doneBtn.textContent = "Mark Completed";
      doneBtn.style.marginLeft = "6px";
      doneBtn.onclick = () => saveCompleted(c);
      btnTd.appendChild(doneBtn);
    } else {
      const doneCopyBtn = document.createElement("button");
      doneCopyBtn.textContent = "Copy Completed";
      doneCopyBtn.onclick = () => openCompletedFormat(c);
      btnTd.appendChild(doneCopyBtn);
    }

    tr.appendChild(btnTd);
    tbody.appendChild(tr);
  });

  document.getElementById("dataTable").hidden = false;
});

/***************************************************
 * COPY FORMAT
 ***************************************************/
function generateCopyFormat(c) {
  document.getElementById("copyText").value = `
Is M/C Covered Under JCB Care / Engine Care / Warranty : U/W
Call ID : ${clean(c[0]?.v)}
Customer Name : ${clean(c[6]?.v)}
Machine SL No. : ${clean(c[9]?.v)}
Engine No : 
M/C Model : ${clean(c[10]?.v)}
HMR : ${clean(c[11]?.v)}
Date of Installation : ${parseDate(c[12])}
Date of Failure : ${parseDate(c[1])}
M/C Location : ${clean(c[21]?.v)}
M/C Application : Material Handling
Dealership & Branch Name : FCV
Engineer Name : ${clean(c[24]?.v)}
M/C Condition : Running with problem
Nature of Complaint : ${clean(c[4]?.v)}
Failed Part Name : 
Failed Part No. : 
Action Required : 
`.trim();

  document.getElementById("copyBox").hidden = false;
}

/***************************************************
 * SAVE COMPLETED (FULL PAYLOAD)
 ***************************************************/
function saveCompleted(c) {
  const callId = String(c[0]?.v);

  const payload = {
    callId,
    createDate: parseDate(c[1]),
    crmCallNo: clean(c[3]?.v),
    subject: clean(c[4]?.v),
    status: clean(c[5]?.v),
    customer: clean(c[6]?.v),
    mobile: clean(c[8]?.v),
    machineNumber: clean(c[9]?.v),
    machineModel: clean(c[10]?.v),
    hmr: clean(c[11]?.v),
    installDate: parseDate(c[12]),
    callType: clean(c[19]?.v),
    callSubType: clean(c[20]?.v),
    branch: clean(c[21]?.v),
    city: clean(c[23]?.v),
    serviceEngg: clean(c[24]?.v),
    machineStatus: clean(c[27]?.v),

    engineNo: document.getElementById("engineInput").value,
    failedPartName: document.getElementById("failedPartNameInput").value,
    failedPartNo: document.getElementById("failedPartNoInput").value,
    actionRequired: document.getElementById("actionRequiredInput").value
  };

  fetch(SAVE_URL, {
    method: "POST",
    body: JSON.stringify(payload)
  }).then(() => {
    processedMap[callId] = payload;
    document.getElementById("copyBox").hidden = true;
    renderTable();
    alert("Saved to OCR_PROCESSED ✔️");
  });
}

/***************************************************
 * COPY COMPLETED
 ***************************************************/
function openCompletedFormat(c) {
  const callId = String(c[0]?.v);
  const d = processedMap[callId];

  document.getElementById("copyText").value = `
Is M/C Covered Under JCB Care / Engine Care / Warranty : U/W
Call ID : ${callId}
Customer Name : ${clean(c[6]?.v)}
Machine SL No. : ${clean(c[9]?.v)}
Engine No : ${d.engineNo}
M/C Model : ${clean(c[10]?.v)}
HMR : ${clean(c[11]?.v)}
Date of Installation : ${parseDate(c[12])}
Date of Failure : ${parseDate(c[1])}
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

  document.getElementById("copyBox").hidden = false;
}

/***************************************************
 * POPUP BUTTONS
 ***************************************************/
function copyToClipboard() {
  const ta = document.getElementById("copyText");
  ta.select();
  document.execCommand("copy");
  alert("Copied ✔️");
}

function closeCopyBox() {
  document.getElementById("copyBox").hidden = true;
}
