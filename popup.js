document.getElementById('processBtn').addEventListener('click', processFiles);
document.getElementById('csvUpload').addEventListener('change', handleFileSelect);

let selectedFiles = [];

function handleFileSelect(event) {
  selectedFiles = Array.from(event.target.files);
  updateButtonState();
}

function updateButtonState() {
  const btn = document.getElementById('processBtn');
  btn.disabled = selectedFiles.length === 0;
}

function updateStatus(message, className) {
  const statusEl = document.getElementById('status');
  statusEl.textContent = message;
  statusEl.className = className;
}

async function processFiles() {
  if (selectedFiles.length === 0) {
    updateStatus("Please select CSV files first", "error");
    return;
  }

  updateStatus("Processing files...", "progress");
  
  try {
    // Classify files
    const dispatchFiles = selectedFiles.filter(f => /^[^()]+\.csv$/i.test(f.name));
    const activityFiles = selectedFiles.filter(f => /\(1\)\.csv$/i.test(f.name));
    const sampleFiles = selectedFiles.filter(f => /\(2\)\.csv$/i.test(f.name));
    
    if (dispatchFiles.length === 0 || activityFiles.length === 0 || sampleFiles.length === 0) {
      throw new Error("Couldn't find all required CSV types");
    }
    
    // Read and process files
    const dispatchData = await readAndProcessFile(dispatchFiles[0], 'dispatch');
    const activityData = await readAndProcessFile(activityFiles[0], 'activity');
    const sampleData = await readAndProcessFile(sampleFiles[0], 'sample');
    
    // Create workbook
    const wb = XLSX.utils.book_new();
    
    XLSX.utils.book_append_sheet(
      wb, 
      XLSX.utils.aoa_to_sheet(dispatchData), 
      "Dispatch"
    );
    
    XLSX.utils.book_append_sheet(
      wb, 
      XLSX.utils.aoa_to_sheet(activityData), 
      "Activities"
    );
    
    XLSX.utils.book_append_sheet(
      wb, 
      XLSX.utils.aoa_to_sheet(sampleData), 
      "Samples"
    );
    
    // Generate filename with date
    const today = new Date();
    const dateString = `${today.getFullYear()}${(today.getMonth()+1).toString().padStart(2,'0')}${today.getDate().toString().padStart(2,'0')}`;
    const fileName = `${dateString}_Report.xlsx`;
    
    // Export to Excel
    XLSX.writeFile(wb, fileName);
    updateStatus(`Success! File saved as ${fileName}`, "success");
    
  } catch (error) {
    updateStatus(`Error: ${error.message}`, "error");
    console.error(error);
  }
}

async function readAndProcessFile(file, type) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        let processedData;
        switch(type) {
          case 'dispatch':
            processedData = processDispatch(rows);
            break;
          case 'activity':
            processedData = processActivity(rows);
            break;
          case 'sample':
            processedData = processSample(rows);
            break;
          default:
            processedData = rows;
        }
        resolve(processedData);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(new Error("Error reading file"));
    reader.readAsArrayBuffer(file);
  });
}

function processDispatch(rows) {
  if (!rows.length) return [];
  const header = rows[0];
  //const data = rows.slice(1).sort((a, b) => (a[1] || '').localeCompare(b[1] || ''));
  var data = rows.slice(1).sort((a, b) => String(a[1] ?? "").localeCompare(String(b[1] ?? "")));
  //const data = rows.slice(1).sort((a, b) => String((a[1] || '').localeCompare((b[1] || ''))));
  var proccessData = data.sort((a, b) => String(a[3] ?? '').localeCompare(String(b[3] ?? '')));
  proccessData = data.sort((a, b) => String(a[0] ?? '').localeCompare(String(b[0] ?? '')));
  return [header, ...data];
}

function processActivity(rows) {
  if (!rows.length) return [];
  const header = rows[0];
  const data = rows.slice(1);
  
  const processedData = data.map(row => {
    const newRow = Array(2).fill('');
    newRow[0] = row[10] || ''; // K -> A
    newRow[1] = row[5] || '';  // F -> B
    
    // Add remaining columns excluding specified indices
    for (let i = 0; i < row.length; i++) {
      if (![5, 9, 10, 11, 12, 15, 17].includes(i)) {
        newRow.push(row[i] || '');
      }
    }
    return newRow;
  }).sort((a, b) => (a[3] || '').localeCompare(b[3] || ''));
  
  return [['K', 'F', ...header.filter((_, i) => ![5, 9, 10, 11, 12, 15, 17].includes(i))], 
          ...processedData];
}

function processSample(rows) {
  if (!rows.length) return [];
  const header = rows[0];
  const data = rows.slice(1);
  
  const processedData = data.map(row => {
    const newRow = [row[4] || '']; // E -> A
    // Add remaining columns except index 4 (E)
    for (let i = 0; i < row.length; i++) {
      if (i !== 4) newRow.push(row[i] || '');
    }
    return newRow;
  }).sort((a, b) => (String(a[0] ?? '')).localeCompare(String(b[0] ?? '')));
  
  return [['E', ...header.filter((_, i) => i !== 4)], ...processedData];
}
