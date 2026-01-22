const SHEET_ID = "15S8wwS4cbi8dFEvpoeIw0EDD1WMZceub5IpmtF5SFes";
const SHEET = "OCR";
const PROCESSED = "OCR_PROCESSED";

const URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${SHEET}`;
const PROCESSED_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${PROCESSED}`;
const SAVE_URL = "https://script.google.com/macros/s/AKfycby4xnVtuSZek7VigeWZI_41-IXfO99xUrNrbeKm31T2pHjbL8LLtvvoj3qklFHlYq1E/exec";

let filtered = [];
let processed = {};

const clean = v => v ? v.toString().trim() : "";

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
        filtered.push(c);
      }
    });

    document.getElementById("count").innerText = filtered.length;
    document.getElementById("viewBtn").disabled = false;
  });

fetch(PROCESSED_URL)
  .then(r => r.text())
  .then(t => {
    const rows = JSON.parse(t.substring(47).slice(0, -2)).table.rows;
    rows.forEach(r => processed[clean(r.c[0].v)] = r.c);
  });

document.getElementById("viewBtn").onclick = () => {
  const tbody = document.querySelector("tbody");
  tbody.innerHTML = "";

  filtered.forEach(c => {
    const id = clean(c[0].v);
    const done = processed[id];

    tbody.innerHTML += `
      <tr>
        <td>${id}</td>
        <td>${c[1]?.f || ""}</td>
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
          <button onclick="openModal('${id}')">Open</button>
        </td>
      </tr>`;
  });

  document.getElementById("dataTable").hidden = false;
};

function openModal(id) {
  document.getElementById("copyText").value = `Call ID : ${id}`;
  document.getElementById("copyModal").hidden = false;
}

function closeModal() {
  document.getElementById("copyModal").hidden = true;
}

function copyToClipboard() {
  const ta = document.getElementById("copyText");
  ta.select();
  document.execCommand("copy");
  alert("Copied ✔️");
}
