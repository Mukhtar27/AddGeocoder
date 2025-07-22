let excelData = [];
let selectedAddressColumn = null;

// Handle drag and drop
const dropArea = document.getElementById('drop-area');

['dragenter', 'dragover'].forEach(eventName => {
  dropArea.addEventListener(eventName, (e) => {
    e.preventDefault();
    dropArea.classList.add('highlight');
  }, false);
});

['dragleave', 'drop'].forEach(eventName => {
  dropArea.addEventListener(eventName, (e) => {
    e.preventDefault();
    dropArea.classList.remove('highlight');
  }, false);
});

dropArea.addEventListener('drop', (e) => {
  const dt = e.dataTransfer;
  const file = dt.files[0];
  document.getElementById('fileElem').files = dt.files;
  handleFile(file);
});

document.getElementById('fileElem').addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (file) {
    handleFile(file);
  }
});

function handleFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    excelData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    showPreview(excelData);
    populateColumnSelector(excelData[0]);
  };
  reader.readAsArrayBuffer(file);
}

function showPreview(data) {
  const table = document.getElementById('tablePreview');
  table.innerHTML = generateTableHTML(data);
}

function populateColumnSelector(headers) {
  const select = document.getElementById('columnSelect');
  select.innerHTML = `<option value="">-- Select Column --</option>`;
  headers.forEach((header, index) => {
    const option = document.createElement('option');
    option.value = index;
    option.textContent = header;
    select.appendChild(option);
  });
}

document.getElementById('columnSelect').addEventListener('change', (e) => {
  selectedAddressColumn = parseInt(e.target.value);
});

document.getElementById('startButton').addEventListener('click', async () => {
  const apiKey = document.getElementById('apiKey').value;
  if (!apiKey || selectedAddressColumn === null) {
    alert("Please enter an API key and select address column.");
    return;
  }

  const outputData = [['Address', 'Latitude', 'Longitude']];
  const progressBar = document.getElementById('progressBar');
  progressBar.value = 0;
  progressBar.max = excelData.length - 1;

  for (let i = 1; i < excelData.length; i++) {
    const row = excelData[i];
    const address = row[selectedAddressColumn];
    if (!address) continue;

    try {
      const response = await fetch(`https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(address)}&key=${apiKey}`);
      const result = await response.json();
      const location = result.results[0]?.geometry?.location;
      if (location) {
        outputData.push([address, location.lat, location.lng]);
      } else {
        outputData.push([address, '', '']);
      }
    } catch {
      outputData.push([address, '', '']);
    }

    progressBar.value = i;
  }

  displayResults(outputData);
  createDownload(outputData);
});

function displayResults(data) {
  const table = document.getElementById('resultTable');
  table.innerHTML = generateTableHTML(data);
}

function generateTableHTML(data) {
  return `
    <table>
      <thead>
        <tr>${data[0].map(cell => `<th>${cell}</th>`).join('')}</tr>
      </thead>
      <tbody>
        ${data.slice(1).map(row => `<tr>${row.map(cell => `<td>${cell ?? ''}</td>`).join('')}</tr>`).join('')}
      </tbody>
    </table>
  `;
}

function createDownload(data) {
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Geocoded Results');
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: "application/octet-stream" });

  const downloadLink = document.createElement('a');
  downloadLink.href = URL.createObjectURL(blob);
  downloadLink.download = 'geocoded_output.xlsx';
  downloadLink.textContent = 'Download Geocoded Excel File';
  downloadLink.style.marginTop = '1rem';
  downloadLink.style.display = 'inline-block';

  const container = document.querySelector('.container');
  container.appendChild(downloadLink);
}
