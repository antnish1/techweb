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
  return null;
}

function formatDate(date) {
  if (!date) return "";
  const d = String(date.getDate()).padStart(2, "0");
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const y = String(date.getFullYear()).slice(-2);
  return `${d}/${m}/${y}`;
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

      const createDate = parseDate(c[1]);
      const installDate = parseDate(c[12]);
      if (!createDate || !installDate) return;

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

      const callId = String(c[0].v).trim();
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
      0, 1, 3, 4, 5, 6, 8, 9, 10, 11,
      12, 19, 20, 21, 23, 24, 27
    ].forEach(i => {
      const td = document.createElement("td");
      td.textContent = c[i]?.f || c[i]?.v || "";
      tr.appendChild(td);
    });

    const callId = String(c[0]?.v).trim();
    const hasProcessedData = !!processedMap[callId];

    const statusTd = document.createElement("td");
    statusTd.textContent = hasProcessedData ? "Completed" : "Pending";
    statusTd.style.fontWeight = "bold";
    statusTd.style.color = hasProcessedData ? "green" : "orange";
    tr.appendChild(statusTd);

    const btnTd = document.createElement("td");

    if (!hasProcessedData) {

      const copyBtn = document.createElement("button");
      copyBtn.textContent = "Copy Format";
      copyBtn.onclick = () => generateCopyFormat(c);
      btnTd.appendChild(copyBtn);

      const doneBtn = document.createElement("button");
      doneBtn.textContent = "Mark Completed";
      doneBtn.style.marginLeft = "6px";

      doneBtn.onclick = () => saveCompleted(c, callId, statusTd);
      btnTd.appendChild(doneBtn);

    } else {

      const copyDoneBtn = document.createElement("button");
      copyDoneBtn.textContent = "Copy Completed";
      copyDoneBtn.onclick = () => openCompletedFormat(c);
      btnTd.appendChild(copyDoneBtn);
    }

    tr.appendChild(btnTd);
    tbody.appendChild(tr);
  });

  document.getElementById("dataTable").hidden = false;
});

/***************************************************
 * COPY FORMAT (PENDING)
 ***************************************************/
function generateCopyFormat(c) {
  const machineNo = clean(c[9]?.v);

  if (machineNo) {
    const t = document.createElement("textarea");
    t.value = machineNo;
    document.body.appendChild(t);
    t.select();
    document.execCommand("copy");
    document.body.removeChild(t);
  }

  document.getElementById("copyText").value = `
Is M/C Covered Under JCB Care / Engine Care / Warranty : U/W
Call ID : ${clean(c[0]?.v)}
Customer Name : ${clean(c[6]?.v)}
Machine SL No. : ${machineNo}
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

  document.getElementById("copyBox").hidden = false;
}

/***************************************************
 * SAVE COMPLETED
 ***************************************************/
function saveCompleted(c, callId, statusTd) {
  const payload = {
    callId,
    engineNo: document.getElementById("engineInput")?.value || "",
    failedPartName: document.getElementById("failedPartNameInput")?.value || "",
    failedPartNo: document.getElementById("failedPartNoInput")?.value || "",
    actionRequired: document.getElementById("actionRequiredInput")?.value || ""
  };

  fetch(SAVE_URL, {
    method: "POST",
    body: JSON.stringify(payload)
  }).then(() => {

    processedMap[callId] = payload;

    statusTd.textContent = "Completed";
    statusTd.style.color = "green";

    document.getElementById("copyBox").hidden = true;
    alert("Saved to sheet ✔️");
  });
}

/***************************************************
 * COPY COMPLETED
 ***************************************************/
function openCompletedFormat(c) {
  const callId = String(c[0]?.v).trim();
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

  document.getElementById("copyBox").hidden = false;
}



/***************************************************
 * POPUP BUTTON HANDLERS (REQUIRED)
 ***************************************************/
function copyToClipboard() {
  const ta = document.getElementById("copyText");
  if (!ta) return;

  ta.select();
  ta.setSelectionRange(0, 99999); // mobile support
  document.execCommand("copy");

  alert("Copied ✔️");
}

function closeCopyBox() {
  const box = document.getElementById("copyBox");
  if (box) box.hidden = true;
}

