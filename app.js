const SHEET_ID = "15S8wwS4cbi8dFEvpoeIw0EDD1WMZceub5IpmtF5SFes";
const SHEET_NAME = "OCR";
const URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${SHEET_NAME}`;

let filteredRows = [];

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

    const btnTd = document.createElement("td");
    const btn = document.createElement("button");
    btn.textContent = "Copy Format";
    btn.onclick = () => generateCopyFormat(c);
    btnTd.appendChild(btn);
    tr.appendChild(btnTd);

    tbody.appendChild(tr);
  });

  document.getElementById("dataTable").hidden = false;
});

/***************************************************
 * COPY FORMAT
 ***************************************************/
function generateCopyFormat(c) {
  const machineNo = clean(c[9]?.v);

  // 1️⃣ COPY MACHINE NUMBER TO CLIPBOARD (SAFE METHOD)
  if (machineNo) {
    const taTemp = document.createElement("textarea");
    taTemp.value = machineNo;
    document.body.appendChild(taTemp);
    taTemp.select();
    document.execCommand("copy");
    document.body.removeChild(taTemp);

    console.log("✅ Machine Number copied to clipboard:", machineNo);
  }

  // 2️⃣ GENERATE COPY FORMAT TEXT
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

  // 3️⃣ SHOW COPY BOX
  document.getElementById("copyText").value = text;
  document.getElementById("copyBox").hidden = false;

  alert(
    "✔ Copy format opened\n\n" +
    "Machine Number copied to clipboard:\n" +
    machineNo
  );
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

    console.log("✅ Engine No injected:", engineNo);
  });
}
