// ====================================================
// Challan Extractor — Main Application
// PDF.js + SheetJS (XLSX) for client-side processing
// ====================================================

// ---- Load PDF.js ----
const pdfjsLib = await import('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.min.mjs');
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.worker.min.mjs';

// ---- Load SheetJS (XLSX) ----
// We load it dynamically via script tag since it's UMD
await new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
    s.onload = resolve;
    s.onerror = reject;
    document.head.appendChild(s);
});

// =====================================================
// STATE
// =====================================================
const state = {
    pdfFiles: [],        // Array of { file: File, name, size }
    templateFile: null,  // File object or null
    extractedData: [],   // Array of { challan, weight, page, fileName }
};

// =====================================================
// DOM REFERENCES
// =====================================================
const dom = {
    dropzone:          document.getElementById('dropzone'),
    fileInput:         document.getElementById('fileInput'),
    fileList:          document.getElementById('fileList'),
    templateDropzone:  document.getElementById('templateDropzone'),
    templateInput:     document.getElementById('templateInput'),
    templateInfo:      document.getElementById('templateInfo'),
    processBtn:        document.getElementById('processBtn'),
    progressContainer: document.getElementById('progressContainer'),
    progressLabel:     document.getElementById('progressLabel'),
    progressPct:       document.getElementById('progressPct'),
    progressBar:       document.getElementById('progressBar'),
    resultsSection:    document.getElementById('resultsSection'),
    resultsCount:      document.getElementById('resultsCount'),
    dataTableBody:     document.getElementById('dataTableBody'),
    downloadBtn:       document.getElementById('downloadBtn'),
    bgParticles:       document.getElementById('bgParticles'),
};

// =====================================================
// BACKGROUND PARTICLES
// =====================================================
function createParticles() {
    for (let i = 0; i < 20; i++) {
        const p = document.createElement('div');
        p.className = 'particle';
        const size = Math.random() * 4 + 1;
        p.style.width = size + 'px';
        p.style.height = size + 'px';
        p.style.left = Math.random() * 100 + '%';
        p.style.animationDuration = (Math.random() * 15 + 10) + 's';
        p.style.animationDelay = (Math.random() * 10) + 's';
        dom.bgParticles.appendChild(p);
    }
}
createParticles();

// =====================================================
// FILE UPLOAD — PDF
// =====================================================
function handlePdfFiles(files) {
    for (const file of files) {
        if (file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')) {
            // Avoid duplicates
            if (!state.pdfFiles.some(f => f.name === file.name && f.size === file.size)) {
                state.pdfFiles.push({ file, name: file.name, size: file.size });
            }
        }
    }
    renderFileList();
    updateProcessBtn();
}

function renderFileList() {
    dom.fileList.innerHTML = '';
    state.pdfFiles.forEach((f, idx) => {
        const el = document.createElement('div');
        el.className = 'file-item';
        const sizeStr = f.size > 1024 * 1024
            ? (f.size / (1024 * 1024)).toFixed(1) + ' MB'
            : (f.size / 1024).toFixed(0) + ' KB';
        el.innerHTML = `
            <div class="file-icon">PDF</div>
            <div class="file-details">
                <div class="file-name">${escHtml(f.name)}</div>
                <div class="file-size">${sizeStr}</div>
            </div>
            <button class="file-remove" data-idx="${idx}" title="Remove">&times;</button>
        `;
        dom.fileList.appendChild(el);
    });
    // Attach remove handlers
    dom.fileList.querySelectorAll('.file-remove').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const idx = parseInt(e.currentTarget.dataset.idx, 10);
            state.pdfFiles.splice(idx, 1);
            renderFileList();
            updateProcessBtn();
        });
    });
}

// Dropzone events
dom.dropzone.addEventListener('click', () => dom.fileInput.click());
dom.fileInput.addEventListener('change', (e) => {
    handlePdfFiles(e.target.files);
    e.target.value = '';
});

['dragenter', 'dragover'].forEach(ev => {
    dom.dropzone.addEventListener(ev, (e) => {
        e.preventDefault();
        dom.dropzone.classList.add('drag-over');
    });
});
['dragleave', 'drop'].forEach(ev => {
    dom.dropzone.addEventListener(ev, (e) => {
        e.preventDefault();
        dom.dropzone.classList.remove('drag-over');
    });
});
dom.dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    handlePdfFiles(e.dataTransfer.files);
});

// =====================================================
// TEMPLATE UPLOAD
// =====================================================
dom.templateDropzone.addEventListener('click', () => dom.templateInput.click());
dom.templateInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        state.templateFile = e.target.files[0];
        dom.templateInfo.innerHTML = `<div class="template-loaded">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6L9 17l-5-5"/></svg>
            ${escHtml(state.templateFile.name)} loaded
        </div>`;
    }
    e.target.value = '';
});

// =====================================================
// PROCESS BUTTON STATE
// =====================================================
function updateProcessBtn() {
    dom.processBtn.disabled = state.pdfFiles.length === 0;
}

// =====================================================
// PDF TEXT EXTRACTION
// =====================================================
async function extractTextFromPdf(file) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const pages = [];
    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        const text = textContent.items.map(item => item.str).join(' ');
        pages.push({ pageNum: i, text });
        
        // Clean up page resources to save memory for large PDFs
        try { page.cleanup(); } catch(e) {}
    }
    try { await pdf.destroy(); } catch(e) {}
    return pages;
}

// =====================================================
// DATA EXTRACTION LOGIC
// =====================================================
function extractChallanData(pages, fileName) {
    const results = [];

    for (const { pageNum, text } of pages) {
        // 1. Challan extraction: Look for any 18-25 digit sequence
        const challanMatch = text.match(/\b\d{18,25}\b/);
        const challanNo = challanMatch ? challanMatch[0] : null;

        // 2. Weight extraction: Look for all decimal/integer numbers followed by M t or Mt
        let weight = null;
        const mtMatches = [...text.matchAll(/(\d+\.?\d*)\s*M\s*t/gi)];
        const smallWeights = mtMatches
            .map(m => parseFloat(m[1]))
            .filter(v => v < 1000 && v > 0);
        
        if (smallWeights.length > 0) {
            weight = smallWeights[0];
        }

        // Only add if we found both values
        if (challanNo && weight !== null) {
            results.push({
                challan: challanNo,
                weight: weight,
                page: pageNum,
                fileName: fileName,
            });
        }
    }

    return results;
}

// =====================================================
// PROGRESS HELPERS
// =====================================================
function showProgress(label, pct) {
    dom.progressContainer.style.display = 'block';
    dom.progressLabel.textContent = label;
    dom.progressPct.textContent = Math.round(pct) + '%';
    dom.progressBar.style.width = pct + '%';
}

function hideProgress() {
    dom.progressContainer.style.display = 'none';
}

// =====================================================
// RESULTS TABLE
// =====================================================
function renderResults() {
    dom.resultsSection.style.display = 'block';
    dom.resultsCount.textContent = `${state.extractedData.length} record${state.extractedData.length !== 1 ? 's' : ''} extracted from ${state.pdfFiles.length} file${state.pdfFiles.length !== 1 ? 's' : ''}`;

    dom.dataTableBody.innerHTML = '';
    state.extractedData.forEach((d, i) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="td-num">${i + 1}</td>
            <td class="td-challan">${escHtml(d.challan)}</td>
            <td class="td-weight">${d.weight.toFixed(2)}</td>
            <td class="td-source">Page ${d.page}</td>
        `;
        dom.dataTableBody.appendChild(tr);
    });

    // Scroll to results
    dom.resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// =====================================================
// EXCEL GENERATION
// =====================================================
async function generateExcel() {
    const formData = new FormData();
    
    // Instead of sending the heavy PDFs, send the already-extracted data as JSON
    const jsonStr = JSON.stringify(state.extractedData);
    const jsonBlob = new Blob([jsonStr], { type: 'application/json' });
    formData.append('data', jsonBlob, 'data.json');
    
    // Add template if present
    if (state.templateFile) {
        formData.append('template', state.templateFile);
    }

    const response = await fetch('/process', {
        method: 'POST',
        body: formData
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(errorText || 'Server failed to process files');
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'challan_data_output.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// =====================================================
// MAIN PROCESS
// =====================================================
dom.processBtn.addEventListener('click', async () => {
    if (state.pdfFiles.length === 0) return;

    dom.processBtn.disabled = true;
    dom.processBtn.classList.add('loading');
    state.extractedData = [];
    dom.resultsSection.style.display = 'none';

    try {
        const totalFiles = state.pdfFiles.length;
        
        for (let fi = 0; fi < totalFiles; fi++) {
            const f = state.pdfFiles[fi];
            showProgress(`Processing: ${f.name}`, ((fi) / totalFiles) * 80);

            // Extract text
            const pages = await extractTextFromPdf(f.file);

            showProgress(`Extracting data: ${f.name}`, ((fi + 0.5) / totalFiles) * 80);

            // Extract challan data
            const data = extractChallanData(pages, f.name);
            state.extractedData.push(...data);

            showProgress(`Done: ${f.name}`, ((fi + 1) / totalFiles) * 80);
        }

        if (state.extractedData.length === 0) {
            showToast('No challan data found in the uploaded PDFs.', 'error');
            hideProgress();
            dom.processBtn.disabled = false;
            dom.processBtn.classList.remove('loading');
            return;
        }

        // Generate Excel
        showProgress('Generating Excel...', 90);
        await generateExcel();

        showProgress('Complete!', 100);
        renderResults();
        showToast(`Extracted ${state.extractedData.length} records successfully!`, 'success');

        // Hide progress after a moment
        setTimeout(() => hideProgress(), 1500);

    } catch (err) {
        console.error('Processing error:', err);
        showToast('Error processing files: ' + err.message, 'error');
        hideProgress();
    } finally {
        dom.processBtn.disabled = false;
        dom.processBtn.classList.remove('loading');
        updateProcessBtn();
    }
});

// Download button re-download
dom.downloadBtn.addEventListener('click', () => {
    if (state.extractedData.length > 0) {
        generateExcel();
    }
});

// =====================================================
// UTILITIES
// =====================================================
function escHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

function showToast(message, type = 'success') {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.textContent = message;
    document.body.appendChild(toast);
    setTimeout(() => {
        if (toast.parentNode) toast.remove();
    }, 3500);
}
