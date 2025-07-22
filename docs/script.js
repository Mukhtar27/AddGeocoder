document.getElementById('fileInput').addEventListener('change', handleFileUpload);
document.getElementById('fetchBtn').addEventListener('click', fetchCoordinates);
document.getElementById('downloadBtn').addEventListener('click', downloadResults);
document.getElementById('togglePassword').addEventListener('change', togglePasswordVisibility);

let resultsArray = [];

function handleFileUpload(event) {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(firstSheet);
        displayPreview(json);
    };
    reader.readAsArrayBuffer(file);
}

function displayPreview(data) {
    let preview = document.getElementById('preview');
    preview.innerHTML = '<strong>Preview of uploaded file:</strong><br><table><tr><th>Name</th><th>Address</th></tr>';
    data.forEach(row => {
        preview.innerHTML += `<tr><td>${row.Name}</td><td>${row.Address}</td></tr>`;
    });
    preview.innerHTML += '</table>';
}

async function fetchCoordinates() {
    const apiKey = document.getElementById('apiKey').value;
    const addresses = Array.from(document.querySelectorAll('#preview table tr td:nth-child(2)'))
                           .map(td => td.textContent);
    resultsArray = [];

    for (const address of addresses) {
        const response = await fetch(`https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(address)}&key=${apiKey}`);
        const result = await response.json();
        if (result.results.length > 0) {
            const location = result.results[0].geometry.location;
            resultsArray.push({ address, latitude: location.lat, longitude: location.lng });
        } else {
            resultsArray.push({ address, latitude: 'Not Found', longitude: 'Not Found' });
        }
    }
    displayResults();
}

function displayResults() {
    let results = document.getElementById('results');
    results.innerHTML = '<strong>Results:</strong><br><table><tr><th>Address</th><th>Latitude</th><th>Longitude</th></tr>';
    resultsArray.forEach(result => {
        results.innerHTML += `<tr><td>${result.address}</td><td>${result.latitude}</td><td>${result.longitude}</td></tr>`;
    });
    results.innerHTML += '</table>';
    document.getElementById('downloadBtn').style.display = 'block';
}

function downloadResults() {
    const ws = XLSX.utils.json_to_sheet(resultsArray);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Results");
    XLSX.writeFile(wb, "GeocodedResults.xlsx");
}

function togglePasswordVisibility() {
    const apiKeyInput = document.getElementById('apiKey');
    if (this.checked) {
        apiKeyInput.type = 'text'; // Show the API key
    } else {
        apiKeyInput.type = 'password'; // Hide the API key
    }
}
