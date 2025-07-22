let uploadedData = [];
let selectedAddressColumn = '';
let apiKey = '';
let darkMode = false;

document.addEventListener('DOMContentLoaded', () => {
  const dropArea = document.getElementById('drop-area');
  const fileInput = document.getElementById('fileElem');
  const tablePreview = document.getElementById('tablePreview');
  const addressSelect = document.getElementById('addressColumn');
  const geocodeBtn = document.getElementById('geocodeBtn');
  const resultTable = document.getElementById('resultTable');
  const apiKeyInput = document.getElementById('apiKey');
  const toggleIcon = document.getElementById('toggleIcon');
  const progressBar = document.getElementById('progressBar');
  const toggleDarkBtn = document.getElementById('toggleDark');

  // Dark mode toggle
  toggleDarkBtn.addEventListener('click', () => {
    document.body.classList.toggle('dark');
  });

  // API key visibility toggle
  toggleIcon.addEventListener('click', () => {
    if (apiKeyInput.type === 'password') {
      apiKeyInput.type = 'text';
      toggleIcon.textContent = '🙈';
    } else {
      apiKeyInput.type = 'password';
      toggleIcon.textContent = '👁️';
    }
  });

  // Drag-and-drop file upload handlers
  ['dragenter', 'dragover'].forEach(eventName => {
    dropArea.addEventListener(eventName, e => {
      e.preventDefault();
      dropArea.classList.add('highlight');
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropArea.addEventListener(eventName, e => {
      e.preventDefault();
      dropArea.classList.remove('highlight');
    });
  });

  dropArea.addEventListener('click', () => fileInput.click());

  dropArea.addEventListener('drop', e => {
    const file = e.dataTransfer.files[0];
    handleFile(file);
  });

  fileInput.addEventListener('change', e => {
    const file = e.target.files[0];
    handleFile(file);
  });

  function handleFile(file) {
    const reader = new FileReader();
    reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      uploadedData = XLSX.utils.sheet_to_json(sheet);
      showPreview(uploadedData);
      populateColumnSelector(uploadedData);
    };
    reader.readAsArrayBuffer(file);
  }

  function showPreview(data) {
    if (data.length === 0) return;
    let html = '<table><thead><tr>';
    Object.keys(data[0]).forEach(key => {
      html += `<th>${key}</th>`;
    });
    html += '</tr></thead><tbody>';
    data.slice(0, 5).forEach(row => {
      html += '<tr>';
      Object.values(row).forEach(cell => {
        html += `<td>${cell || ''}</td>`;
      });
      html += '</tr>';
    });
    html += '</tbody></table>';
    tablePreview.innerHTML = html;
  }

  function populateColumnSelector(data) {
    if (!data || data.length === 0) return;
    const columns = Object.keys(data[0]);
    addressSelect.innerHTML = '<option value="">--Select Column--</option>';
    columns.forEach(col => {
      addressSelect.innerHTML += `<option value="${col}">${col}</option>`;
    });
  }

  geocodeBtn.addEventListener('click', async () => {
    apiKey = apiKeyInput.value.trim();
    selectedAddressColumn = addressSelect.value;

    if (!apiKey || !selectedAddressColumn || uploadedData.length === 0) {
      alert('Please provide API key, select address column and upload file.');
      return;
    }

    progressBar.value = 0;
    progressBar.max = uploadedData.length;

    const results = [];

    for (let i = 0; i < uploadedData.length; i++) {
      const row = uploadedData[i];
      const address = row[selectedAddressColumn];
      const coords = await geocodeAddress(address);
      results.push({
        ...row,
        Latitude: coords.lat,
        Longitude: coords.lng
      });
      progressBar.value = i + 1;
    }

    showResultTable(results);
  });

  async function geocodeAddress(address) {
    try {
      const response = await fetch(`https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(address)}&key=${apiKey}`);
      const data = await response.json();
      if (data.status === 'OK') {
        const loc = data.results[0].geometry.location;
        return { lat: loc.lat, lng: loc.lng };
      } else {
        return { lat: '', lng: '' };
      }
    } catch (e) {
      return { lat: '', lng: '' };
    }
  }

  function showResultTable(data) {
    if (data.length === 0) return;
    let html = '<table><thead><tr>';
    Object.keys(data[0]).forEach(key => {
      html += `<th>${key}</th>`;
    });
    html += '</tr></thead><tbody>';
    data.slice(0, 5).forEach(row => {
      html += '<tr>';
      Object.values(row).forEach(cell => {
        html += `<td>${cell || ''}</td>`;
      });
      html += '</tr>';
    });
    html += '</tbody></table>';
    resultTable.innerHTML = html;
  }
});
