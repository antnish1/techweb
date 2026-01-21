const SHEET_ID = "15S8wwS4cbi8dFEvpoeIw0EDD1WMZceub5IpmtF5SFes";
const SHEET_NAME = "OCR";
const URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${SHEET_NAME}`;
const SAVE_URL = "https://script.google.com/macros/s/AKfycby4xnVtuSZek7VigeWZI_41-IXfO99xUrNrbeKm31T2pHjbL8LLtvvoj3qklFHlYq1E/exec";
const PROCESSED_SHEET_NAME = "OCR_PROCESSED";
const PROCESSED_URL =
  `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${PROCESSED_SHEET_NAME}`;

let processedMap = {}; // callId -> processed row




let filteredRows = [];


/***************************************************
 * ROW STATUS STORAGE (LOCAL)
 ***************************************************/
function getRowStatus(callId) {
  return localStorage.getItem("row_status_" + callId) || "Pending";
}

function setRowStatus(callId, status) {
  localStorage.setItem("row_status_" + callId, status);
}


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
 * FETCH & FILTER DATA
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
      0, 1, 3, 4, 5, 6, 8, 9, 10, 11,
      12, 19, 20, 21, 23, 24, 27
    ].forEach(i => {
      const td = document.createElement("td");
      td.textContent = c[i]?.f || c[i]?.v || "";
      tr.appendChild(td);
    });

    // Progress column
    const callId = clean(c[0]?.v);
    const currentStatus = getRowStatus(callId);
    
    const statusTd = document.createElement("td");
    statusTd.textContent = currentStatus;
    statusTd.style.fontWeight = "bold";
    
    if (currentStatus === "Completed") {
      statusTd.style.color = "green";
    } else if (currentStatus === "In Progress") {
      statusTd.style.color = "blue";
    } else {
      statusTd.style.color = "orange";
    }
    
    tr.appendChild(statusTd);

    
    // Action button
    // Action buttons
    const btnTd = document.createElement("td");
    
    // Copy Format button
    const btn = document.createElement("button");
    btn.textContent = "Copy Format";
    btn.onclick = () => generateCopyFormat(c, statusTd);
    btnTd.appendChild(btn);
    
    // Mark Completed button
    const doneBtn = document.createElement("button");
    doneBtn.textContent = "Mark Completed";
    doneBtn.style.marginLeft = "6px";


    if (currentStatus === "Completed") {
      const copyDoneBtn = document.createElement("button");
      copyDoneBtn.textContent = "Copy Completed";
      copyDoneBtn.style.marginLeft = "6px";
    
      copyDoneBtn.onclick = () => {
        openCompletedFormat(c);
      };
    
      btnTd.appendChild(copyDoneBtn);
    }


    
    doneBtn.onclick = () => {

      const payload = {
        callId: clean(c[0]?.v),
        createDate: c[1]?.f || "",
        crmCallNo: clean(c[3]?.v),
        subject: clean(c[4]?.v),
        status: clean(c[5]?.v),
        customer: clean(c[6]?.v),
        mobile: clean(c[8]?.v),
        machineNumber: clean(c[9]?.v),
        machineModel: clean(c[10]?.v),
        hmr: clean(c[11]?.v),
        installDate: c[12]?.f || "",
        callType: clean(c[19]?.v),
        callSubType: clean(c[20]?.v),
        branch: clean(c[21]?.v),
        city: clean(c[23]?.v),
        serviceEngg: clean(c[24]?.v),
        machineStatus: clean(c[27]?.v),
    
        engineNo: document.getElementById("engineInput")?.value || "",
        failedPartName: document.getElementById("failedPartNameInput")?.value || "",
        failedPartNo: document.getElementById("failedPartNoInput")?.value || "",
        actionRequired: document.getElementById("actionRequiredInput")?.value || ""
      };
    
      fetch(SAVE_URL, {
        method: "POST",
        body: JSON.stringify(payload)
      }).then(() => {
        setRowStatus(callId, "Completed");
        statusTd.textContent = "Completed";
        statusTd.style.color = "green";
        alert("Saved to sheet ✔️");
      }).catch(() => {
        alert("❌ Failed to save");
      });
    };

    
    btnTd.appendChild(doneBtn);
    tr.appendChild(btnTd);



    tbody.appendChild(tr);
  });

  document.getElementById("dataTable").hidden = false;
});

/***************************************************
 * COPY FORMAT
 ***************************************************/
function generateCopyFormat(c, statusTd) {

  if (statusTd) {
    statusTd.textContent = "In Progress";
    statusTd.style.color = "blue";
    setRowStatus(clean(c[0]?.v), "In Progress");
  }

  
  const machineNo = clean(c[9]?.v);

  // Copy machine number silently
  if (machineNo) {
    const taTemp = document.createElement("textarea");
    taTemp.value = machineNo;
    document.body.appendChild(taTemp);
    taTemp.select();
    document.execCommand("copy");
    document.body.removeChild(taTemp);

    console.log("Machine Number copied:", machineNo);
  }

  const text = `
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

  document.getElementById("copyText").value = text;
  document.getElementById("copyBox").hidden = false;
}






/***************************************************
 * CLIPBOARD
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

/***************************************************
 * ENGINE NO INJECTION (TEXT-BASED, CORRECT)
 ***************************************************/
function pasteEngineNo() {
  navigator.clipboard.readText().then(text => {
    if (!text.startsWith("__ENGINE_NO__=")) {
      alert("❌ Clipboard does not contain Engine No");
      return;
    }

    const engineNo = text.replace("__ENGINE_NO__=", "").trim();
    const ta = document.getElementById("copyText");

    if (!ta.value.includes("Engine No :")) {
      alert("❌ Copy format not open");
      return;
    }

    ta.value = ta.value.replace(
      /Engine No\s*:\s*.*/i,
      `Engine No : ${engineNo}`
    );

    const engineInput = document.getElementById("engineInput");
    if (engineInput) {
      engineInput.value = engineNo;
    }


    console.log("✅ Engine No injected:", engineNo);
  });
}


function openCompletedFormat(c) {
  const callId = clean(c[0]?.v);
  const completed = processedMap[callId];

  if (!completed) {
    alert("❌ Completed data not found in OCR_PROCESSED");
    return;
  }

  const text = `
Is M/C Covered Under JCB Care / Engine Care / Warranty : U/W
Call ID : ${clean(c[0]?.v)}
Customer Name : ${clean(c[6]?.v)}
Machine SL No. : ${clean(c[9]?.v)}
Engine No : ${completed.engineNo}
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
Failed Part Name : ${completed.failedPartName}
Failed Part No. : ${completed.failedPartNo}
Action Required : ${completed.actionRequired}
`.trim();

  document.getElementById("copyText").value = text;
  document.getElementById("copyBox").hidden = false;
}




/***************************************************
 * ENTER KEY = PASTE ENGINE NO
 ***************************************************/
document.addEventListener("keydown", function (e) {
  // Only react to Enter key
  if (e.key !== "Enter") return;

  const copyBox = document.getElementById("copyBox");

  // Only when copy box is visible
  if (!copyBox || copyBox.hidden) return;

  // Prevent accidental new lines or form submits
  e.preventDefault();

  // Trigger paste engine number
  pasteEngineNo();
});

