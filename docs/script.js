// ========================
// 📜 script.js
// ========================

let workbookData = [];
let selectedSheet = null;
let addressColumn = null;
let apiKey = '';
let geocodedResults = [];

const dropArea = document.getElementById("drop-area");
const fileElem = document.getElementById("fileElem");
const clearFileBtn = document.getElementById("clearFileBtn");
const addressSelect = document.getElementById("addressSelect");
const addressColumnContainer = document.getElementById("addressColumnContainer");
const apiKeyInput = document.getElementById("apiKey");
const toggleApiKey = document.getElementById("toggleApiKey");
const geocodeBtn = document.getElementById("geocodeBtn");
const downloadBtn = document.getElementById("downloadBtn");
const actionButtons = document.getElementById("actionButtons");
const tablePreview = document.getElementById("tablePreview");
const fileNameDisplay = document.getElementById("fileNameDisplay");
const progressContainer = document.getElementById("progressContainer");
const progressBar = document.getElementById("progressBar");
const progressText = document.getElementById("progressText");
const resultPreview = document.getElementById("resultPreview");
const resultTable = document.getElementById("resultTable");
const themeToggle = document.getElementById("themeToggle");

// 📂 File Upload
fileElem.addEventListener("change", handleFiles);
dropArea.addEventListener("click", () => fileElem.click());
dropArea.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropArea.classList.add("highlight");
});
dropArea.addEventListener("dragleave", () => dropArea.classList.remove("highlight"));
dropArea.addEventListener("drop", (e) => {
  e.preventDefault();
  dropArea.classList.remove("highlight");
  handleFiles(e);
});

clearFileBtn.addEventListener("click", () => {
  fileElem.value = "";
  workbookData = [];
  selectedSheet = null;
  addressColumn = null;
  tablePreview.innerHTML = "";
  addressSelect.innerHTML = "";
  addressColumnContainer.classList.add("hidden");
  actionButtons.classList.add("hidden");
  clearFileBtn.classList.add("hidden");
  fileNameDisplay.innerHTML = "";
});

function handleFiles(e) {
  const file = e.target.files?.[0] || e.dataTransfer.files[0];
  if (!file) return;
  fileNameDisplay.innerHTML = `✅ File selected: <strong>${file.name}</strong>`;
  clearFileBtn.classList.remove("hidden");

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    selectedSheet = workbook.SheetNames[0];
    workbookData = XLSX.utils.sheet_to_json(workbook.Sheets[selectedSheet]);
    displayTablePreview(workbookData);
  };
  reader.readAsArrayBuffer(file);
}

function displayTablePreview(data) {
  if (!data.length) return;
  const keys = Object.keys(data[0]);

  // Address column options
  addressSelect.innerHTML = keys.map(k => `<option value="${k}">${k}</option>`).join("");
  addressColumnContainer.classList.remove("hidden");
  actionButtons.classList.remove("hidden");

  // Table preview
  let html = '<table><thead><tr>' + keys.map(k => `<th>${k}</th>`).join('') + '</tr></thead><tbody>';
  data.slice(0, 5).forEach(row => {
    html += '<tr>' + keys.map(k => `<td>${row[k] ?? ""}</td>`).join('') + '</tr>';
  });
  html += '</tbody></table>';
  tablePreview.innerHTML = html;
}

// 👁️ Toggle API Key
let showKey = false;
toggleApiKey.addEventListener("click", () => {
  showKey = !showKey;
  apiKeyInput.type = showKey ? "text" : "password";
  toggleApiKey.textContent = showKey ? "🙈" : "👁️";
});

// 🌗 Toggle Dark Mode
let dark = false;
themeToggle.addEventListener("click", () => {
  dark = !dark;
  document.body.classList.toggle("dark-mode", dark);
});

// 📍 Geocode
geocodeBtn.addEventListener("click", async () => {
  apiKey = apiKeyInput.value.trim();
  if (!apiKey) return alert("Please enter your API key");
  addressColumn = addressSelect.value;
  if (!addressColumn) return alert("Select an address column");

  geocodedResults = [];
  progressContainer.classList.remove("hidden");
  resultPreview.classList.add("hidden");

  for (let i = 0; i < workbookData.length; i++) {
    const row = workbookData[i];
    const address = row[addressColumn];
    if (!address) {
      geocodedResults.push({ ...row, Latitude: "", Longitude: "", Status: "No address" });
      updateProgress(i + 1, workbookData.length);
      continue;
    }

    try {
      const result = await fetch(
        `https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(address)}&key=${apiKey}`
      ).then(res => res.json());

      if (result.status === "OK") {
        const loc = result.results[0].geometry.location;
        geocodedResults.push({ ...row, Latitude: loc.lat, Longitude: loc.lng, Status: "OK" });
      } else {
        geocodedResults.push({ ...row, Latitude: "", Longitude: "", Status: result.status });
      }
    } catch (err) {
      geocodedResults.push({ ...row, Latitude: "", Longitude: "", Status: "Error" });
    }
    updateProgress(i + 1, workbookData.length);
    await new Promise(r => setTimeout(r, 300)); // delay to avoid throttling
  }

  displayResultTable();
  downloadBtn.classList.remove("hidden");
});

function updateProgress(done, total) {
  const percent = Math.round((done / total) * 100);
  progressBar.value = percent;
  progressText.textContent = `Progress: ${done} of ${total}`;
}

function displayResultTable() {
  resultPreview.classList.remove("hidden");
  const keys = Object.keys(geocodedResults[0]);
  let html = '<table><thead><tr>' + keys.map(k => `<th>${k}</th>`).join('') + '</tr></thead><tbody>';
  geocodedResults.slice(0, 10).forEach(row => {
    html += '<tr>' + keys.map(k => `<td>${row[k] ?? ""}</td>`).join('') + '</tr>';
  });
  html += '</tbody></table>';
  resultTable.innerHTML = html;
}

// ⬇️ Download CSV
function downloadCSV(rows) {
  const sheet = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet, "Geocoded");
  XLSX.writeFile(wb, "geocoded_addresses.xlsx");
}
downloadBtn.addEventListener("click", () => downloadCSV(geocodedResults));
