import * as XLSX from "https://cdn.sheetjs.com/xlsx-0.20.0/package/xlsx.mjs";

let workbook, sheetData;
const fileElem = document.getElementById("fileElem");
const fileNameDisplay = document.getElementById("fileNameDisplay");
const clearFileBtn = document.getElementById("clearFileBtn");
const tablePreview = document.getElementById("tablePreview");
const dropArea = document.getElementById("drop-area");
const geocodeBtn = document.getElementById("geocodeBtn");
const downloadBtn = document.getElementById("downloadBtn");
const addressSelect = document.getElementById("addressSelect");
const addressColumnContainer = document.getElementById("addressColumnContainer");
const actionButtons = document.getElementById("actionButtons");
const apiKeyInput = document.getElementById("apiKey");
const toggleApiKey = document.getElementById("toggleApiKey");

// 🔐 API Toggle logic
toggleApiKey.addEventListener("click", () => {
  const type = apiKeyInput.getAttribute("type");
  apiKeyInput.setAttribute("type", type === "password" ? "text" : "password");
  toggleApiKey.textContent = type === "password" ? "🔓" : "🔒";
});

// 📂 Drag & Drop File Upload
dropArea.addEventListener("click", () => {
  fileElem.click();
});
dropArea.addEventListener("dragover", e => {
  e.preventDefault();
  dropArea.classList.add("highlight");
});
dropArea.addEventListener("dragleave", () => dropArea.classList.remove("highlight"));
dropArea.addEventListener("drop", e => {
  e.preventDefault();
  dropArea.classList.remove("highlight");
  handleFile(e.dataTransfer.files[0]);
});
fileElem.addEventListener("change", () => {
  if (fileElem.files.length > 0) handleFile(fileElem.files[0]);
});
clearFileBtn.addEventListener("click", (e) => {
  e.stopPropagation();
  resetApp();
});

// 📄 File Handling and Table Preview
function handleFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
    displayTable(sheetData);
    populateAddressOptions(sheetData[0]);
    addressColumnContainer.classList.remove("hidden");
    actionButtons.classList.remove("hidden");
    fileNameDisplay.textContent = file.name;
    clearFileBtn.style.display = "inline-block";
  };
  reader.readAsArrayBuffer(file);
}

function displayTable(data) {
  let html = `<table><thead><tr>${data[0].map(col => `<th>${col}</th>`).join("")}</tr></thead><tbody>`;
  data.slice(1, 6).forEach(row => {
    html += `<tr>${data[0].map((_, i) => `<td>${row[i] || ""}</td>`).join("")}</tr>`;
  });
  html += "</tbody></table>";
  tablePreview.innerHTML = html;
}

function populateAddressOptions(headers) {
  addressSelect.innerHTML = headers.map(h => `<option value="${h}">${h}</option>`).join("");
}

// 📍 Geocode Logic with Delay, Status, and Progress
geocodeBtn.addEventListener("click", async () => {
  const addressColumn = addressSelect.value;
  const headers = sheetData[0];
  const addressIdx = headers.indexOf(addressColumn);
  if (addressIdx === -1) return alert("Address column not found.");

  const apiKey = apiKeyInput.value.trim();
  if (!apiKey) return alert("Please enter your API key.");

  const updatedData = [headers.concat(["Latitude", "Longitude", "Geocode Status"])];
  const progressDiv = document.createElement("div");
  progressDiv.id = "progress";
  tablePreview.before(progressDiv);

  for (let i = 1; i < sheetData.length; i++) {
    const row = sheetData[i];
    const address = row[addressIdx];
    progressDiv.textContent = `Geocoding ${i} of ${sheetData.length - 1}...`;

    if (!address) {
      updatedData.push(row.concat(["", "", "Empty Address"]));
      continue;
    }

    try {
      const res = await fetch(`https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(address)}&key=${apiKey}`);
      const data = await res.json();
      if (data.status === "OK") {
        const location = data.results[0].geometry.location;
        updatedData.push(row.concat([location.lat, location.lng, "OK"]));
      } else {
        updatedData.push(row.concat(["", "", data.status]));
      }
    } catch (error) {
      updatedData.push(row.concat(["", "", "Fetch Error"]));
    }

    await new Promise(r => setTimeout(r, 150)); // ⏱️ Rate limiting
  }

  progressDiv.textContent = "✅ Geocoding Complete";

  const ws = XLSX.utils.aoa_to_sheet(updatedData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Geocoded");
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/octet-stream" });

  downloadBtn.classList.remove("hidden");
  downloadBtn.onclick = () => {
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "geocoded_addresses.xlsx";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };
});

// 🔁 Reset
function resetApp() {
  workbook = null;
  sheetData = null;
  apiKeyInput.value = "";
  fileElem.value = "";
  fileNameDisplay.textContent = "";
  clearFileBtn.style.display = "none";
  tablePreview.innerHTML = "";
  addressSelect.innerHTML = "";
  addressColumnContainer.classList.add("hidden");
  actionButtons.classList.add("hidden");
  downloadBtn.classList.add("hidden");
  toggleApiKey.textContent = "🔒";
  apiKeyInput.type = "password";
  const progressDiv = document.getElementById("progress");
  if (progressDiv) progressDiv.remove();
}
