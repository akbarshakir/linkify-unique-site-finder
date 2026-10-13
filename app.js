let masterWorkbook, vendorWorkbook;
let masterDomains = new Set();
let vendorRows = [];
let uniqueRows = [];
let vendorColumns = [];

function cleanText(val) { return val ? String(val).trim() : ""; }

function extractDomain(url) {
    if (!url) return "";
    let clean = cleanText(url).toLowerCase();
    clean = clean.replace(/<[^>]*>/g, " ").replace(/^https?:\/\//, '').replace(/^www\./, '');
    return clean.split('/')[0].split('?')[0].replace(/[.,;:]+$/, "").replace(/\s+/g, "");
}

// ==========================================
// MASTER UI TABS
// ==========================================
document.getElementById('masterTabCloud').addEventListener('click', (e) => {
    e.target.classList.add('active'); 
    document.getElementById('masterTabFile').classList.remove('active');
    document.getElementById('masterCloudSection').style.display = 'block'; 
    document.getElementById('masterFileSection').style.display = 'none';
});

document.getElementById('masterTabFile').addEventListener('click', (e) => {
    e.target.classList.add('active'); 
    document.getElementById('masterTabCloud').classList.remove('active');
    document.getElementById('masterFileSection').style.display = 'block'; 
    document.getElementById('masterCloudSection').style.display = 'none';
});

// ==========================================
// MASTER LOGIC (CLOUD & FILE)
// ==========================================
document.getElementById('masterLoadBtnCloud').addEventListener('click', async () => {
    const statusEl = document.getElementById('masterStatus');
    statusEl.textContent = "Fetching database... Please wait.";
    statusEl.style.borderLeftColor = "#f39c12";

    try {
        const response = await fetch('/api/domains');
        if (!response.ok) throw new Error("Database API error");
        const domainsArray = await response.json();
        
        masterDomains = new Set(domainsArray);
        document.getElementById('masterCount').textContent = masterDomains.size;
        
        statusEl.textContent = `Loaded ${masterDomains.size} unique domains from Cloudflare D1.`;
        statusEl.style.borderLeftColor = "#2ecc71";
        checkCompareReady();
    } catch (e) {
        statusEl.textContent = "Error loading database: " + e.message;
        statusEl.style.borderLeftColor = "#e74c3c";
    }
});

document.getElementById('masterFile').addEventListener('change', e => {
    if(!e.target.files[0]) return;
    const reader = new FileReader();
    reader.onload = e => {
        masterWorkbook = XLSX.read(new Uint8Array(e.target.result), {type: 'array'});
        const select = document.getElementById('masterSheetSelect');
        select.innerHTML = '';
        masterWorkbook.SheetNames.forEach(name => select.add(new Option(name, name)));
        document.getElementById('masterSheetSelectArea').style.display = 'block';
    };
    reader.readAsArrayBuffer(e.target.files[0]);
});

document.getElementById('masterLoadBtnFile').addEventListener('click', () => {
    if(!masterWorkbook) return;
    const sheet = masterWorkbook.Sheets[document.getElementById('masterSheetSelect').value];
    const rows = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ""});
    
    masterDomains.clear();
    rows.forEach(row => {
        row.forEach(cell => {
            const domain = extractDomain(cell);
            if(domain && domain.includes('.')) masterDomains.add(domain);
        });
    });

    document.getElementById('masterCount').textContent = masterDomains.size;
    const statusEl = document.getElementById('masterStatus');
    statusEl.textContent = `Loaded ${masterDomains.size} unique domains from Excel.`;
    statusEl.style.borderLeftColor = "#2ecc71";
    checkCompareReady();
});

// ==========================================
// VENDOR LOGIC
// ==========================================
document.getElementById('vendorFile').addEventListener('change', e => {
    if(!e.target.files[0]) return;
    const reader = new FileReader();
    reader.onload = e => {
        vendorWorkbook = XLSX.read(new Uint8Array(e.target.result), {type: 'array'});
        const select = document.getElementById('vendorSheetSelect');
        select.innerHTML = '';
        vendorWorkbook.SheetNames.forEach(name => select.add(new Option(name, name)));
        document.getElementById('vendorSheetSelectArea').style.display = 'block';
        autoDetectVendorHeader();
    };
    reader.readAsArrayBuffer(e.target.files[0]);
});

function autoDetectVendorHeader() {
    if(!vendorWorkbook) return;
    const sheet = vendorWorkbook.Sheets[document.getElementById('vendorSheetSelect').value];
    const rows = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ""});
    let headerRow = 1;
    for(let i = 0; i < Math.min(rows.length, 20); i++) {
        let str = rows[i].map(c => String(c).toLowerCase()).join(" ");
        if(str.includes("website") || str.includes("domain") || str.includes("url")) {
            headerRow = i + 1; break;
        }
    }
    document.getElementById('vendorHeaderRow').value = headerRow;
}
document.getElementById('vendorSheetSelect').addEventListener('change', autoDetectVendorHeader);

document.getElementById('vendorLoadBtn').addEventListener('click', () => {
    if(!vendorWorkbook) return;
    const sheet = vendorWorkbook.Sheets[document.getElementById('vendorSheetSelect').value];
    let headerIndex = parseInt(document.getElementById('vendorHeaderRow').value) - 1;
    if(headerIndex < 0) headerIndex = 0;

    let rawData = XLSX.utils.sheet_to_json(sheet, {defval: "", range: headerIndex});
    if(rawData.length === 0) return alert("No vendor data found.");

    vendorColumns = Object.keys(rawData[0]);
    vendorRows = [];

    let webCol = vendorColumns.find(c => ["website", "domain", "url", "site"].includes(c.toLowerCase().trim()));
    if(!webCol) webCol = vendorColumns[0]; 

    rawData.forEach(row => {
        let domain = extractDomain(row[webCol]);
        if(domain && domain.includes('.')) {
            row["__normalized_domain"] = domain;
            vendorRows.push(row);
        }
    });

    document.getElementById('vendorCount').textContent = vendorRows.length;
    const statusEl = document.getElementById('vendorStatus');
    statusEl.textContent = `Loaded ${vendorRows.length} valid vendor rows.`;
    statusEl.style.borderLeftColor = "#2ecc71";
    checkCompareReady();
});

// ==========================================
// COMPARE, EXPORT & SAVE TO DB
// ==========================================
function checkCompareReady() {
    const btn = document.getElementById('compareBtn');
    if(btn) btn.disabled = !(masterDomains.size > 0 && vendorRows.length > 0);
}

document.getElementById('compareBtn').addEventListener('click', () => {
    uniqueRows = [];
    let existingCount = 0;
    let seenVendorDomains = new Set();

    vendorRows.forEach(row => {
        let domain = row["__normalized_domain"];
        if(seenVendorDomains.has(domain)) return; 
        seenVendorDomains.add(domain);

        if(masterDomains.has(domain)) {
            existingCount++;
        } else {
            uniqueRows.push(row);
        }
    });

    document.getElementById('existingCount').textContent = existingCount;
    document.getElementById('uniqueCount').textContent = uniqueRows.length;
    
    renderResults();
    document.getElementById('exportBtn').disabled = uniqueRows.length === 0;
    document.getElementById('saveDbBtn').disabled = uniqueRows.length === 0;
});

function renderResults() {
    const thead = document.getElementById('resultHead');
    const tbody = document.getElementById('resultBody');
    if(!thead || !tbody) return;
    
    thead.innerHTML = ''; tbody.innerHTML = '';
    if(uniqueRows.length === 0) return;

    let trHead = document.createElement('tr');
    vendorColumns.forEach(col => {
        let th = document.createElement('th'); th.textContent = col; trHead.appendChild(th);
    });
    thead.appendChild(trHead);

    uniqueRows.slice(0, 50).forEach(row => {
        let tr = document.createElement('tr');
        vendorColumns.forEach(col => {
            let td = document.createElement('td'); td.textContent = row[col]; tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    document.getElementById('resultsSection').style.display = 'block';
}

document.getElementById('exportBtn').addEventListener('click', () => {
    if(uniqueRows.length === 0) return;
    const vendorNameInput = document.getElementById('vendorName');
    const vendorName = (vendorNameInput ? vendorNameInput.value.trim() : "") || "Unknown_Vendor";
    
    let exportData = uniqueRows.map(row => {
        let cleanRow = {};
        vendorColumns.forEach(col => cleanRow[col] = row[col]);
        cleanRow["Normalized Domain"] = row["__normalized_domain"];
        cleanRow["Vendor Name"] = vendorName; 
        return cleanRow;
    });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(exportData);
    XLSX.utils.book_append_sheet(wb, ws, "Unique Sites");
    XLSX.writeFile(wb, `Unique_Sites_${vendorName}.xlsx`);
});

document.getElementById('saveDbBtn').addEventListener('click', async () => {
    if(uniqueRows.length === 0) return;
    
    const btn = document.getElementById('saveDbBtn');
    btn.textContent = "Saving...";
    btn.disabled = true;

    const newDomains = uniqueRows.map(r => r["__normalized_domain"]);

    try {
        const response = await fetch('/api/domains', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ domains: newDomains })
        });

        if (!response.ok) throw new Error("Failed to save to database");
        
        alert("Successfully saved unique domains to Cloudflare D1 Master Database!");
        
        // Agar cloud tab active hai toh auto-refresh karein
        if(document.getElementById('masterTabCloud').classList.contains('active')) {
            document.getElementById('masterLoadBtnCloud').click();
        }
        
        btn.textContent = "☁️ Saved!";
    } catch (e) {
        alert("Error saving: " + e.message);
        btn.textContent = "☁️ Save Unique to Database";
        btn.disabled = false;
    }
});
