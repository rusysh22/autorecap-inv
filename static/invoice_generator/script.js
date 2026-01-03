document.addEventListener('DOMContentLoaded', () => {
    // --- Elements ---
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');

    const configPanel = document.getElementById('config-panel');
    const fileList = document.getElementById('file-list');
    const fileCount = document.getElementById('file-count');
    const clearBtn = document.getElementById('clear-files');
    const processBtn = document.getElementById('process-btn');
    const btnSpinner = document.getElementById('btn-spinner');
    const btnText = processBtn.querySelector('span:not(.sr-only)');

    const filenameSuffixInput = document.getElementById('filename-suffix');

    // Master Data Elements
    const tabUpload = document.getElementById('tab-upload');
    const tabPaste = document.getElementById('tab-paste');
    const contentUpload = document.getElementById('content-upload');
    const contentPaste = document.getElementById('content-paste');
    const masterFileInput = document.getElementById('master-file-input');
    const masterFileList = document.getElementById('master-file-list');
    const masterPasteArea = document.getElementById('master-paste-area');
    const pasteCount = document.getElementById('paste-count');
    const clearPasteBtn = document.getElementById('clear-paste-btn');
    const previewPasteBtn = document.getElementById('preview-paste-btn');

    // Modal Elements
    const previewModal = document.getElementById('preview-modal');
    const closePreviewBtn = document.getElementById('close-preview-btn');
    const closePreviewFooterBtn = document.getElementById('close-preview-footer-btn');
    const previewBackdrop = document.getElementById('preview-backdrop');
    const previewTableBody = document.getElementById('preview-table-body');

    const loadingSection = document.getElementById('loading');
    const resultsSection = document.getElementById('results-section');

    const summaryFiles = document.getElementById('summary-files');
    const summaryRows = document.getElementById('summary-rows');
    const summaryAmount = document.getElementById('summary-amount');

    const breakdownBody = document.getElementById('breakdown-body');

    const downloadBtn = document.getElementById('download-btn');
    const tableHeader = document.getElementById('table-header');
    const tableBody = document.getElementById('table-body');
    const resetBtn = document.getElementById('reset-btn');
    const tableSearch = document.getElementById('table-search');
    const outputFilenameDisplay = document.getElementById('output-filename');
    const notificationArea = document.getElementById('notification-area');

    // --- State ---
    let selectedFiles = [];


    // ... (drag & drop handlers) ...

    if (resetBtn) {
        resetBtn.addEventListener('click', () => {
            if (!confirm("Are you sure you want to reset all files and data?")) return;

            selectedFiles = [];
            renderFileList(); // Updates UI and hides config panel
            resultsSection.classList.add('hidden');
            if (notificationArea) notificationArea.classList.add('hidden');

            fileInput.value = '';
            if (filenameSuffixInput) filenameSuffixInput.value = '';
            if (tableSearch) tableSearch.value = '';
            if (outputFilenameDisplay) outputFilenameDisplay.textContent = '';

            // Clear Column Filters
            const colFilters = document.querySelectorAll('.column-filter');
            colFilters.forEach(input => input.value = '');

            // Hide Reset Button again
            resetBtn.classList.add('hidden');

            // Scroll to top smoothly
            window.scrollTo({ top: 0, behavior: 'smooth' });

            // Clear Master Data
            selectedMasterFiles = [];
            renderMasterFileList();
            if (masterPasteArea) {
                masterPasteArea.value = '';
                pasteCount.textContent = '0';
                clearPasteBtn.classList.add('hidden');
            }
        });
    }

    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
        document.body.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.add('border-blue-500', 'bg-blue-50', 'scale-[1.02]'));
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.classList.remove('border-blue-500', 'bg-blue-50', 'scale-[1.02]');
    });

    dropZone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        handleFiles(files);
    });

    dropZone.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        handleFiles(e.target.files);
    });

    // --- Filename Validation ---
    if (filenameSuffixInput) {
        // Set dynamic placeholder to today's date
        const today = new Date().toISOString().split('T')[0];
        filenameSuffixInput.placeholder = `${today} (Optional)`;

        filenameSuffixInput.addEventListener('input', () => {
            // Remove invalid filename characters: / \ : * ? " < > |
            filenameSuffixInput.value = filenameSuffixInput.value.replace(/[<>:"/\\|?*]/g, '');
        });
    }

    // --- File Handling ---

    function handleFiles(files) {
        const MAX_FILE_SIZE = 5 * 1024 * 1024; // 5 MB
        const MAX_FILE_COUNT = 15;

        // 1. Validate Total Count
        if (selectedFiles.length + files.length > MAX_FILE_COUNT) {
            alert(`You can only upload a maximum of ${MAX_FILE_COUNT} files.`);
            return;
        }

        const newFiles = Array.from(files).filter(file => {
            // 2. Validate Extension
            if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
                return false;
            }
            // 3. Validate Size
            if (file.size > MAX_FILE_SIZE) {
                alert(`File "${file.name}" exceeds the 5MB limit and was skipped.`);
                return false;
            }
            // 4. Validate Duplicate
            if (selectedFiles.some(f => f.name === file.name)) {
                alert(`File "${file.name}" is already selected.`);
                return false;
            }
            return true;
        });

        if (newFiles.length === 0 && files.length > 0) return; // All skipped or invalid

        selectedFiles = [...selectedFiles, ...newFiles];
        renderFileList();
    }

    function renderFileList() {
        if (selectedFiles.length > 0) {
            configPanel.classList.remove('hidden');
        } else {
            configPanel.classList.add('hidden');
            resultsSection.classList.add('hidden'); // Hide results if no files
        }

        fileCount.textContent = selectedFiles.length;
        fileList.innerHTML = '';
        processBtn.disabled = selectedFiles.length === 0;

        selectedFiles.forEach((file, index) => {
            const li = document.createElement('li');
            li.className = 'flex justify-between items-center bg-slate-50/80 p-3 rounded-lg border border-slate-100 text-sm group transition-colors hover:bg-white hover:shadow-xs';
            li.innerHTML = `
                <div class="flex items-center overflow-hidden">
                    <svg class="w-4 h-4 text-green-600 mr-2 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>
                    <span class="truncate font-medium text-slate-700">${file.name}</span>
                </div>
                <button class="ml-2 text-slate-400 hover:text-red-500 opacity-0 group-hover:opacity-100 transition-opacity p-1" onclick="removeFile(${index})" title="Remove">
                    <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg>
                </button>
            `;
            fileList.appendChild(li);
        });

        window.removeFile = (index) => {
            selectedFiles.splice(index, 1);
            renderFileList();
        };
    }

    clearBtn.addEventListener('click', () => {
        selectedFiles = [];
        renderFileList();
        fileInput.value = '';
    });



    // --- Master Data Handling ---
    let selectedMasterFiles = [];
    let pastedMasterData = [];

    // TABS
    if (tabUpload && tabPaste) {
        tabUpload.addEventListener('click', () => {
            tabUpload.classList.add('bg-white', 'text-blue-600', 'shadow-sm');
            tabUpload.classList.remove('text-slate-500');
            tabPaste.classList.remove('bg-white', 'text-blue-600', 'shadow-sm');
            tabPaste.classList.add('text-slate-500');
            contentUpload.classList.remove('hidden');
            contentPaste.classList.add('hidden');
        });
        tabPaste.addEventListener('click', () => {
            tabPaste.classList.add('bg-white', 'text-blue-600', 'shadow-sm');
            tabPaste.classList.remove('text-slate-500');
            tabUpload.classList.remove('bg-white', 'text-blue-600', 'shadow-sm');
            tabUpload.classList.add('text-slate-500');
            contentPaste.classList.remove('hidden');
            contentUpload.classList.add('hidden');
        });
    }

    // MULTIPLE FILES
    if (masterFileInput) {
        masterFileInput.addEventListener('change', (e) => {
            const MAX_FILE_SIZE = 5 * 1024 * 1024; // 5 MB
            const MAX_FILE_COUNT = 15;

            // 1. Validate Total Count
            if (selectedMasterFiles.length + e.target.files.length > MAX_FILE_COUNT) {
                alert(`You can only upload a maximum of ${MAX_FILE_COUNT} Master Data files.`);
                masterFileInput.value = '';
                return;
            }

            const newFiles = Array.from(e.target.files).filter(file => {
                // 2. Validate Extension
                if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
                    return false;
                }
                // 3. Validate Size
                if (file.size > MAX_FILE_SIZE) {
                    alert(`Master File "${file.name}" exceeds the 5MB limit and was skipped.`);
                    return false;
                }
                // 4. Validate Duplicate
                if (selectedMasterFiles.some(f => f.name === file.name)) {
                    alert(`Master File "${file.name}" is already selected.`);
                    return false;
                }
                return true;
            });

            if (newFiles.length > 0) {
                selectedMasterFiles = [...selectedMasterFiles, ...newFiles];
                renderMasterFileList();
            }
            masterFileInput.value = ''; // reset to allow re-selecting same file
        });
    }

    function renderMasterFileList() {
        if (!masterFileList) return;
        masterFileList.innerHTML = '';
        selectedMasterFiles.forEach((file, index) => {
            const li = document.createElement('li');
            li.className = "flex justify-between items-center bg-green-50 dark:bg-green-900/20 px-2 py-1 rounded border border-green-200 dark:border-green-800";
            li.innerHTML = `
                <div class="flex items-center overflow-hidden">
                    <span class="truncate pr-2 text-green-700 dark:text-green-400 text-xs">${file.name}</span>
                </div>
                <button class="text-green-400 hover:text-red-500 font-bold ml-2" onclick="removeMasterFile(${index})">×</button>
            `;
            masterFileList.appendChild(li);
        });

        window.removeMasterFile = (index) => {
            selectedMasterFiles.splice(index, 1);
            renderMasterFileList();
        };
    }

    // PASTE HANDLING
    if (masterPasteArea) {
        masterPasteArea.addEventListener('input', () => {
            // Validasi "silent" saat mengetik (hanya update jumlah record, tak muncul error)
            parsePastedData(masterPasteArea.value, true);
            const hasData = masterPasteArea.value.trim() !== '';
            clearPasteBtn.classList.toggle('hidden', !hasData);
            if (previewPasteBtn) previewPasteBtn.classList.toggle('hidden', !hasData);
        });

        if (clearPasteBtn) {
            clearPasteBtn.addEventListener('click', () => {
                masterPasteArea.value = '';
                parsePastedData('', true);
                // Clear any lingering error messages about past data
                const notif = document.getElementById('notification-area');
                if (notif && notif.textContent.includes("Pasted Data")) {
                    notif.classList.add('hidden');
                    notif.innerHTML = '';
                }
                clearPasteBtn.classList.add('hidden');
                if (previewPasteBtn) previewPasteBtn.classList.add('hidden');
            });
        }

        // --- Modal Events ---
        if (previewPasteBtn) {
            previewPasteBtn.addEventListener('click', () => {
                // Re-parse to be sure
                // If invalid, parsePastedData returns false and renders error (if suppressError=false)
                const isValid = parsePastedData(masterPasteArea.value, false);

                if (isValid && pastedMasterData.length > 0) {
                    renderPastePreview();
                    previewModal.classList.remove('hidden');
                } else {
                    // Explicitly show error if validation failed
                    if (masterPasteArea.value.trim() === "") {
                        alert("Please paste data first.");
                    } else {
                        // Scroll to notification
                        const notif = document.getElementById('notification-area');
                        if (notif) notif.scrollIntoView({ behavior: 'smooth' });
                    }
                }
            });
        }

        const closeModal = () => previewModal.classList.add('hidden');
        if (closePreviewBtn) closePreviewBtn.addEventListener('click', closeModal);
        if (closePreviewFooterBtn) closePreviewFooterBtn.addEventListener('click', closeModal);
        if (previewBackdrop) previewBackdrop.addEventListener('click', closeModal);
    }

    function renderPastePreview() {
        if (!previewTableBody) return;
        previewTableBody.innerHTML = pastedMasterData.map(item => `
            <tr class="hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors">
                <td class="whitespace-nowrap px-3 py-2 text-xs font-mono text-slate-600 dark:text-slate-300 border-b border-slate-100 dark:border-slate-700/50">${item.kode}</td>
                <td class="whitespace-nowrap px-3 py-2 text-xs text-slate-700 dark:text-slate-200 border-b border-slate-100 dark:border-slate-700/50">${item.nama}</td>
            </tr>
        `).join('');
    }

    /**
     * Parse Pasted Data
     * @param {string} text 
     * @param {boolean} suppressError If true, won't show notification popup (for realtime typing)
     * @returns {boolean} isValid
     */
    function parsePastedData(text, suppressError = false) {
        // Reset state
        const notif = document.getElementById('notification-area');

        // If clearing text, clear errors and state
        if (!text || !text.trim()) {
            pastedMasterData = [];
            pasteCount.textContent = '0';
            if (notif && notif.textContent.includes("Pasted Data")) {
                notif.innerHTML = '';
                notif.classList.add('hidden');
            }
            return true;
        }

        const lines = text.trim().split(/\r?\n/).filter(line => line.trim());
        if (lines.length === 0) return true;

        // 1. Detect Headers in First Row & Delimiters
        // Try Tab first, then Comma, then Semicolon, then Pipe
        let delimiter = '\t';
        let possibleDelimiters = ['\t', ',', ';', '|'];
        let detectedHeaders = [];

        // Simple heuristic: check which delimiter gives us > 1 column in the first line
        for (let d of possibleDelimiters) {
            let cols = lines[0].split(d);
            if (cols.length > 1) {
                delimiter = d;
                detectedHeaders = cols.map(h => h.trim());
                break; // Found a working delimiter
            }
        }

        // If still 1 column, fallback to tab (maybe it's just 1 column data? unlikely for Kode+Nama)
        if (detectedHeaders.length === 0) {
            detectedHeaders = lines[0].split(delimiter).map(h => h.trim());
        }

        const headerRow = detectedHeaders.map(h => h.toLowerCase());

        const aliasesKode = ["kode", "tugas id", "kode tugas", "任务单号", "id", "no.surat jalan di sistem (kode tugas)"];
        const aliasesNama = ["rute", "ritase", "kode ritase", "nama rute", "nama tugas", "线路", "nama tugas"];

        let kodeIdx = -1;
        let namaIdx = -1;

        // Find indices
        headerRow.forEach((col, idx) => {
            if (kodeIdx === -1 && aliasesKode.some(alias => col.includes(alias))) kodeIdx = idx;
            if (namaIdx === -1 && aliasesNama.some(alias => col.includes(alias))) namaIdx = idx;
        });

        // 2. Validate Headers
        if (kodeIdx === -1 || namaIdx === -1) {
            if (!suppressError) {
                // Pass the raw found headers for clearer error message
                renderNotifications(["Master Data Error: Columns not found in Pasted Data. Found: " + JSON.stringify(detectedHeaders)], []);
            }
            pastedMasterData = [];
            pasteCount.textContent = 'Error';
            return false;
        }

        // 3. Parse Rows
        const parsed = [];
        for (let i = 1; i < lines.length; i++) {
            const parts = lines[i].split(delimiter); // Use the detected delimiter
            // Safety check for bounds
            const k = parts[kodeIdx] ? parts[kodeIdx].trim() : '';
            const n = parts[namaIdx] ? parts[namaIdx].trim() : '';

            if (k && n) {
                parsed.push({ kode: k, nama: n });
            }
        }

        pastedMasterData = parsed;
        pasteCount.textContent = parsed.length;

        // Clear error if valid
        if (notif && notif.textContent.includes("Pasted Data")) {
            notif.innerHTML = '';
            notif.classList.add('hidden');
        }

        return true;
    }

    // --- API & Processing ---

    processBtn.addEventListener('click', async () => {
        if (selectedFiles.length === 0) return;

        // Pre-flight Validation for Paste Data
        if (masterPasteArea && masterPasteArea.value.trim().length > 0) {
            const isPasteValid = parsePastedData(masterPasteArea.value, false); // showError = true
            if (!isPasteValid) {
                // Validation failed, error notification is shown relative of the failure
                // We return here to prevent sending invalid data
                const notif = document.getElementById('notification-area');
                if (notif) notif.scrollIntoView({ behavior: 'smooth' });
                return;
            }
        }

        // (Removed duplicate validation)

        // UI Loading State
        loadingSection.classList.remove('hidden');
        resultsSection.classList.add('hidden');
        processBtn.disabled = true;
        btnSpinner.classList.remove('hidden');
        btnText ? btnText.textContent = "Processing..." : null;

        const formData = new FormData();
        selectedFiles.forEach(file => {
            formData.append('files', file);
        });

        // Append Config
        if (filenameSuffixInput) {
            formData.append('filename_suffix', filenameSuffixInput.value);
        }

        // Append Master Files (Multiple)
        selectedMasterFiles.forEach(file => {
            formData.append('master_files', file);
        });

        // Append Pasted Data as JSON
        if (pastedMasterData.length > 0) {
            formData.append('master_data_json', JSON.stringify(pastedMasterData));
        }

        try {
            const response = await fetch(window.API_URL || '/api/process', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            if (result.success) {
                renderResults(result);

                // Collect anomalies from file summaries
                // Collect anomalies from file summaries
                const anomalies = [];
                // Fix: Access file_details from summary object (result.summary.file_details)
                const fileSummaries = result.summary && result.summary.file_details ? result.summary.file_details : [];

                if (fileSummaries.length > 0) {
                    fileSummaries.forEach(f => {
                        if (f.anomalies && f.anomalies.length > 0) {
                            f.anomalies.forEach(a => {
                                anomalies.push(`<strong>${f.filename}</strong>: ${a}`);
                            });
                        }
                    });
                }
                // Combine global warnings with anomalies
                const allWarnings = [...(result.warnings || []), ...anomalies];

                renderNotifications(allWarnings, result.missing_codes);

                // Show Reset Button only after successful processing
                if (resetBtn) resetBtn.classList.remove('hidden');
            } else {
                alert('Error processing files: ' + result.error);
            }
        } catch (error) {
            console.error('Error:', error);
            alert('An error occurred while communicating with the server.');
        } finally {
            loadingSection.classList.add('hidden');
            processBtn.disabled = false;
            btnSpinner.classList.add('hidden');
            btnText ? btnText.textContent = "Process Files" : null;
        }
    });

    function renderNotifications(warnings, missingCodes) {
        if (!notificationArea) return;

        let htmlContent = '';

        // 1. Warnings & Errors
        // 1. Warnings & Errors (Refactored for Grouping)
        if (warnings && warnings.length > 0) {

            // Group anomalies by filename
            const groupedAnomalies = {};
            const standardWarnings = [];

            warnings.forEach(msg => {
                // Try to extract filename from "<strong>Filename</strong>: Error msg" or similar patterns
                // Logic: Check if it starts with <strong>...</strong> (our formatted structure)
                if (msg.startsWith('<strong>')) {
                    const closingTag = '</strong>: ';
                    const idx = msg.indexOf(closingTag);
                    if (idx !== -1) {
                        const filename = msg.substring(8, idx); // removed <strong>
                        const errorContent = msg.substring(idx + closingTag.length);

                        if (!groupedAnomalies[filename]) {
                            groupedAnomalies[filename] = [];
                        }
                        groupedAnomalies[filename].push(errorContent);
                        return;
                    }
                }

                // Handling "Master Data File Error" specially or leaving as generic
                if (msg.includes("Master Data Error") || msg.includes("Columns not found")) {
                    // Keep existing complex logic or simplify? 
                    // Let's keep the existing complex renderer for Master Data separately if possible,
                    // but the current loop structure mixes them.
                    // For safety, let's treat these complex HTML blocks as "Standard Warnings" 
                    // unless we want to parse them too. 
                    // The previous logic generated HTML blocks strings directly into `warnings`? 
                    // No, `warnings` from backend is list of strings. 
                    // Wait, `process_excel_files` returns generic strings or formatted?
                    // Backend returns simpler strings now usually.
                    // The previous `renderNotifications` implementation handled the internal HTML refactoring.
                    // Let's re-implement the Master Data HTML logic for `standardWarnings`
                    standardWarnings.push(msg);
                } else {
                    // Check if it's a generic "<strong>Filename</strong>" pattern we missed?
                    // If backend sends "<strong>File</strong>: ...", we catch it.
                    // If backend sends "Row X: ...", it is generic.
                    standardWarnings.push(msg);
                }
            });

            // A. Render Standard Warnings (Global/Master Data)
            standardWarnings.forEach(msg => {
                // ... (Keep existing Master Data logic here) ...
                if (msg.includes("Master Data Error") || msg.includes("Columns not found")) {
                    // Attempt to parse "Columns not found in FILENAME. Found: [LIST]"
                    let formattedMsg = msg.replace('Master Data Error:', '').trim();
                    let filename = 'Unknown File';
                    let foundCols = [];

                    if (formattedMsg.includes("Columns not found in")) {
                        try {
                            const parts = formattedMsg.split("Found: ");
                            if (parts.length > 1) {
                                const prefix = parts[0].replace("Columns not found in ", "").trim();
                                filename = prefix.endsWith('.') ? prefix.slice(0, -1) : prefix;
                                let colStr = parts[1];
                                try {
                                    foundCols = JSON.parse(colStr);
                                } catch (e) {
                                    colStr = colStr.replace(/[\[\]']/g, "");
                                    foundCols = colStr.split(",").map(c => c.trim()).filter(c => c);
                                }
                            }
                        } catch (e) { console.error(e); }
                    }

                    htmlContent += `
                        <div class="bg-red-50 dark:bg-red-900/40 border-l-4 border-red-500 p-4 rounded-md shadow-sm mb-3">
                            <div class="flex items-start">
                                <svg class="h-6 w-6 text-red-500 dark:text-red-400 mr-3 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/>
                                </svg>
                                <div class="w-full">
                                    <h4 class="font-bold text-base text-red-900 dark:text-red-50 mb-1">Master Data File Error</h4>
                                    ${foundCols.length > 0 ? `
                                        <p class="text-sm text-red-800 dark:text-red-200 mb-2">Required columns <strong>(Kode Tugas, Nama Tugas)</strong> not found in file:</p>
                                        <div class="bg-white/50 dark:bg-black/20 p-2 rounded border border-red-100 dark:border-red-500/30 mb-2">
                                            <p class="font-mono text-xs text-red-700 dark:text-red-100 font-semibold break-all">📄 ${filename}</p>
                                        </div>
                                        <p class="text-xs text-red-700 dark:text-red-300 mb-1 font-medium">Header columns detected in this file:</p>
                                        <div class="bg-red-100/50 dark:bg-black/40 p-2 rounded border border-red-200 dark:border-red-500/30 max-h-32 overflow-y-auto custom-scrollbar">
                                            <div class="flex flex-wrap gap-1">
                                                ${foundCols.map(col => `<span class="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-red-100 dark:bg-red-800/50 text-red-800 dark:text-red-100 border border-red-200 dark:border-red-700">${col}</span>`).join('')}
                                            </div>
                                        </div>
                                    ` : `<p class="text-sm text-red-800 dark:text-red-200 opacity-90 break-words leading-relaxed">${formattedMsg}</p>`}
                                </div>
                            </div>
                        </div>
                    `;
                } else {
                    htmlContent += `
                        <div class="flex items-start bg-amber-50 dark:bg-amber-900/40 border-l-4 border-amber-400 p-4 rounded-md shadow-sm mb-3">
                            <svg class="h-5 w-5 text-amber-500 dark:text-amber-400 mr-3 mt-0.5 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"/>
                            </svg>
                            <div class="text-sm text-amber-800 dark:text-amber-100">
                                <p class="font-medium leading-relaxed">${msg}</p>
                            </div>
                        </div>
                    `;
                }
            });

            // B. Render Grouped Anomalies
            for (const [filename, errors] of Object.entries(groupedAnomalies)) {
                const uniqueID = 'anomaly-group-' + Math.random().toString(36).substr(2, 9);
                const errorCount = errors.length;
                const isCollapsible = errorCount > 3;

                const renderedErrors = errors.map(e => `<li class="py-1 border-b border-amber-200/30 last:border-0">${e}</li>`).join('');

                htmlContent += `
                    <div class="bg-amber-50 dark:bg-amber-900/20 border-l-4 border-amber-500 rounded-md shadow-sm mb-3 overflow-hidden">
                        <div class="p-4 pb-2 flex items-start justify-between cursor-pointer" onclick="document.getElementById('${uniqueID}').classList.toggle('hidden'); this.querySelector('.arrow-icon').classList.toggle('-rotate-90')">
                             <div class="flex items-center gap-3">
                                <div class="bg-amber-100 dark:bg-amber-800 p-2 rounded-lg text-amber-600 dark:text-amber-200">
                                    <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"/></svg>
                                </div>
                                <div>
                                    <h4 class="font-bold text-amber-900 dark:text-amber-100 text-sm">${filename}</h4>
                                    <p class="text-xs text-amber-700 dark:text-amber-300 mt-0.5">${errorCount} Anomalies found</p>
                                </div>
                             </div>
                             <button class="text-amber-500 hover:text-amber-700 transition-transform duration-200 arrow-icon">
                                <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7"></path></svg>
                             </button>
                        </div>
                        
                        <div id="${uniqueID}" class="${isCollapsible ? 'hidden' : ''} bg-amber-100/50 dark:bg-amber-900/40 px-4 py-2 text-xs font-mono text-amber-900 dark:text-amber-200 max-h-60 overflow-y-auto custom-scrollbar border-t border-amber-200 dark:border-amber-700">
                            <ul class="list-none space-y-0">
                                ${renderedErrors}
                            </ul>
                        </div>
                        ${isCollapsible ? `
                             <div class="px-4 py-1.5 bg-amber-100/30 dark:bg-amber-800/20 text-center cursor-pointer hover:bg-amber-100 dark:hover:bg-amber-800/40 transition-colors" onclick="document.getElementById('${uniqueID}').classList.toggle('hidden'); this.previousElementSibling.previousElementSibling.querySelector('.arrow-icon').classList.toggle('-rotate-90')">
                                <span class="text-[10px] font-bold text-amber-600 dark:text-amber-400 uppercase tracking-widest">Warning Details</span>
                            </div>
                        ` : ''}
                    </div>
                 `;
            }

        }

        // 2. Missing Codes Specific Block
        if (missingCodes && missingCodes.length > 0) {
            const count = missingCodes.length;
            const uniqueID = 'missing-codes-' + Date.now();

            htmlContent += `
                <div class="flex flex-col bg-orange-50 dark:bg-orange-900/40 border-l-4 border-orange-400 p-4 rounded-md shadow-sm mb-3">
                    <div class="flex items-start">
                        <svg class="h-5 w-5 text-orange-500 dark:text-orange-400 mr-3 mt-0.5 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/>
                        </svg>
                        <div class="text-sm text-orange-800 dark:text-orange-100 flex-1">
                            <p class="font-medium">⚠️ ${count} Kode Tugas not found in Master Data (using default name).</p>
                             <button type="button" onclick="document.getElementById('${uniqueID}').classList.toggle('hidden')" class="mt-2 text-orange-600 dark:text-orange-300 hover:text-orange-800 dark:hover:text-orange-100 font-semibold underline text-xs focus:outline-none transition-colors">
                                Show/Hide List
                            </button>
                        </div>
                    </div>
                    
                    <div id="${uniqueID}" class="mt-3 pl-8">
                        <div class="bg-white/50 dark:bg-black/30 p-3 rounded border border-orange-100 dark:border-orange-500/30 max-h-40 overflow-y-auto custom-scrollbar">
                            <ul class="list-disc list-inside text-xs font-mono text-orange-900 dark:text-orange-200 space-y-1">
                                ${missingCodes.map(code => `<li>${code}</li>`).join('')}
                            </ul>
                        </div>
                    </div>
                </div>
            `;
        }

        if (htmlContent === '') {
            notificationArea.classList.add('hidden');
            notificationArea.innerHTML = '';
        } else {
            notificationArea.innerHTML = htmlContent;
            notificationArea.classList.remove('hidden');
        }
    }

    function renderResults(result) {
        // Summary Cards
        summaryFiles.textContent = result.summary.total_files;
        summaryRows.textContent = result.summary.total_rows;
        summaryAmount.textContent = formatCurrency(result.summary.total_amount);

        // Download Link
        // downloadBtn.href = result.summary.download_url; // Removed for stateless
        setupDownloadButton(result.summary.output_filename, result.summary.excel_data);

        // Output Filename Display
        if (outputFilenameDisplay && result.summary.output_filename) {
            outputFilenameDisplay.textContent = result.summary.output_filename;
            outputFilenameDisplay.title = result.summary.output_filename;
        }

        // Color Palette for files
        // Color Palette for files (Light & Dark Mode)
        const colors = [
            { bg: 'bg-red-50 dark:bg-red-900/20', border: 'border-red-500 dark:border-red-500', text: 'text-red-700 dark:text-gray-100', badge: 'bg-red-100 dark:bg-red-600/40' },
            { bg: 'bg-orange-50 dark:bg-orange-900/20', border: 'border-orange-500 dark:border-orange-500', text: 'text-orange-700 dark:text-gray-100', badge: 'bg-orange-100 dark:bg-orange-600/40' },
            { bg: 'bg-amber-50 dark:bg-amber-900/20', border: 'border-amber-500 dark:border-amber-500', text: 'text-amber-700 dark:text-gray-100', badge: 'bg-amber-100 dark:bg-amber-600/40' },
            { bg: 'bg-green-50 dark:bg-green-900/20', border: 'border-green-500 dark:border-green-500', text: 'text-green-700 dark:text-gray-100', badge: 'bg-green-100 dark:bg-green-600/40' },
            { bg: 'bg-emerald-50 dark:bg-emerald-900/20', border: 'border-emerald-500 dark:border-emerald-500', text: 'text-emerald-700 dark:text-gray-100', badge: 'bg-emerald-100 dark:bg-emerald-600/40' },
            { bg: 'bg-teal-50 dark:bg-teal-900/20', border: 'border-teal-500 dark:border-teal-500', text: 'text-teal-700 dark:text-gray-100', badge: 'bg-teal-100 dark:bg-teal-600/40' },
            { bg: 'bg-cyan-50 dark:bg-cyan-900/20', border: 'border-cyan-500 dark:border-cyan-500', text: 'text-cyan-700 dark:text-gray-100', badge: 'bg-cyan-100 dark:bg-cyan-600/40' },
            { bg: 'bg-sky-50 dark:bg-sky-900/20', border: 'border-sky-500 dark:border-sky-500', text: 'text-sky-700 dark:text-gray-100', badge: 'bg-sky-100 dark:bg-sky-600/40' },
            { bg: 'bg-blue-50 dark:bg-blue-900/20', border: 'border-blue-500 dark:border-blue-500', text: 'text-blue-700 dark:text-gray-100', badge: 'bg-blue-100 dark:bg-blue-600/40' },
            { bg: 'bg-indigo-50 dark:bg-indigo-900/20', border: 'border-indigo-500 dark:border-indigo-500', text: 'text-indigo-700 dark:text-gray-100', badge: 'bg-indigo-100 dark:bg-indigo-600/40' },
            { bg: 'bg-violet-50 dark:bg-violet-900/20', border: 'border-violet-500 dark:border-violet-500', text: 'text-violet-700 dark:text-gray-100', badge: 'bg-violet-100 dark:bg-violet-600/40' },
            { bg: 'bg-purple-50 dark:bg-purple-900/20', border: 'border-purple-500 dark:border-purple-500', text: 'text-purple-700 dark:text-gray-100', badge: 'bg-purple-100 dark:bg-purple-600/40' },
            { bg: 'bg-fuchsia-50 dark:bg-fuchsia-900/20', border: 'border-fuchsia-500 dark:border-fuchsia-500', text: 'text-fuchsia-700 dark:text-gray-100', badge: 'bg-fuchsia-100 dark:bg-fuchsia-600/40' },
            { bg: 'bg-pink-50 dark:bg-pink-900/20', border: 'border-pink-500 dark:border-pink-500', text: 'text-pink-700 dark:text-gray-100', badge: 'bg-pink-100 dark:bg-pink-600/40' },
            { bg: 'bg-rose-50 dark:bg-rose-900/20', border: 'border-rose-500 dark:border-rose-500', text: 'text-rose-700 dark:text-gray-100', badge: 'bg-rose-100 dark:bg-rose-600/40' }
        ];

        // Map filename to color index
        const fileColorMap = {};
        const fileDetails = result.summary.file_details || [];

        fileDetails.forEach((f, index) => {
            fileColorMap[f.filename] = colors[index % colors.length];
        });

        // --- 1. File Breakdown Table ---
        if (breakdownBody) {
            breakdownBody.innerHTML = fileDetails.map((f, index) => {
                const color = fileColorMap[f.filename];
                const isError = f.status.startsWith('Error');
                const isWarning = f.status === 'Warning';

                let statusClass = 'text-green-600 bg-green-50 border border-green-100 px-2 py-0.5 rounded-full text-xs font-semibold';
                if (isError) {
                    statusClass = 'text-red-600 bg-red-50 border border-red-100 px-2 py-0.5 rounded-full text-xs font-semibold';
                } else if (isWarning) {
                    statusClass = 'text-amber-600 bg-amber-50 border border-amber-100 px-2 py-0.5 rounded-full text-xs font-semibold';
                }

                return `
                    <tr class="hover:bg-slate-50 dark:hover:bg-slate-700/30 transition-colors">
                        <td class="px-6 py-4 font-medium text-slate-700 dark:text-white">
                             <div class="flex items-center gap-3">
                                <div class="w-3 h-3 rounded-full flex-shrink-0 ${color.border.replace('border', 'bg')}" title="Color Tag"></div>
                                <div class="truncate max-w-[250px]" title="${f.filename}">${f.filename}</div>
                            </div>
                        </td>
                        <td class="px-6 py-4 text-center">
                            <span class="${statusClass}">${f.status}</span>
                        </td>
                        <td class="px-6 py-4 text-right text-slate-600 dark:text-gray-200">${f.rows}</td>
                        <td class="px-6 py-4 text-right font-mono text-slate-500 dark:text-gray-200 text-xs">${formatCurrency(f.ppn)}</td>
                        <td class="px-6 py-4 text-right font-mono text-slate-500 dark:text-gray-200 text-xs">${formatCurrency(f.pph)}</td>
                        <td class="px-6 py-4 text-right font-mono text-slate-700 dark:text-white font-bold">${formatCurrency(f.amount)}</td>
                    </tr>
                `;
            }).join('');
        }

        // --- 2. Main Data Preview Table ---
        const data = result.data;
        if (data.length === 0) {
            tableHeader.innerHTML = '<tr><td colspan="100%" class="p-8 text-center text-slate-500">No data found</td></tr>';
            tableBody.innerHTML = '';
            resultsSection.classList.remove('hidden');
            return;
        }

        // Exclude 'source_file' from columns header but use it for row styling
        const displayColumns = result.display_columns || Object.keys(data[0]).filter(col => col !== 'source_file');

        tableHeader.innerHTML = displayColumns.map(col =>
            `<th scope="col" class="px-6 py-3 font-semibold tracking-wide whitespace-nowrap bg-slate-100 dark:bg-slate-700/50 dark:text-slate-200">
                <div class="flex flex-col gap-2">
                    <span>${col}</span>
                    <input type="text" class="column-filter bg-white dark:bg-slate-600 border border-slate-300 dark:border-slate-500 text-slate-900 dark:text-white text-xs rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-1.5 font-normal placeholder-slate-400" placeholder="Filter ${col}..." data-col="${col}">
                </div>
            </th>`
        ).join('');

        tableBody.innerHTML = data.map(row => {
            const filename = row['source_file'];
            const color = fileColorMap[filename] || colors[0];

            return `<tr class="bg-white dark:bg-transparent hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors border-b last:border-0 border-slate-100 dark:border-slate-700/50">
                ${displayColumns.map((col, idx) => {
                const cellValue = row[col] !== null ? row[col] : '';
                // Inject filename badge in the first cell
                if (idx === 0) {
                    return `<td class="px-6 py-4 whitespace-nowrap max-w-[200px]" title="${cellValue}">
                            <div class="flex flex-col">
                                <span class="cell-data truncate block ${isNumeric(row[col]) ? 'font-mono text-right' : ''} text-slate-700 dark:text-slate-200">${cellValue}</span>
                                <span class="text-[10px] ${color.text} mt-1 font-medium px-1.5 py-0.5 rounded ${color.badge} w-fit opacity-90 max-w-full truncate" title="${filename}">${filename}</span>
                            </div>
                         </td>`;
                }
                return `<td class="px-6 py-4 whitespace-nowrap max-w-[200px] truncate ${isNumeric(row[col]) ? 'font-mono text-right' : ''} text-slate-600 dark:text-slate-300" title="${cellValue}">
                        <span class="cell-data">${cellValue}</span>
                    </td>`;
            }).join('')}
            </tr>`;
        }).join('');

        // Attach Event Listeners for Column Filters
        const colFilters = document.querySelectorAll('.column-filter');
        colFilters.forEach(input => {
            input.addEventListener('input', applyFilters);
        });

        resultsSection.classList.remove('hidden');
        resultsSection.scrollIntoView({ behavior: 'smooth' });
    }

    function applyFilters() {
        const globalTerm = tableSearch ? tableSearch.value.toLowerCase() : '';
        const rows = tableBody.querySelectorAll('tr');
        const colFilters = document.querySelectorAll('.column-filter');

        // Build map of active column filters: { 'Nama Tugas': 'bgr', 'Plat': 'b99' }
        const activeColFilters = {};
        colFilters.forEach(input => {
            if (input.value.trim() !== '') {
                activeColFilters[input.dataset.col] = input.value.toLowerCase();
            }
        });

        rows.forEach(row => {
            let matchesGlobal = true;
            let matchesColumns = true;

            // 1. Global Search Check
            if (globalTerm) {
                // We use row.textContent for broad search (includes hidden badge text, which is fine)
                if (!row.textContent.toLowerCase().includes(globalTerm)) {
                    matchesGlobal = false;
                }
            }

            // 2. Column Filters Check
            // We need to match specific cell by column index
            if (matchesGlobal && Object.keys(activeColFilters).length > 0) {
                // Convert row cells to an object keyed by column name or use index mapping
                // Since DOM cells are ordered same as displayColumns, we can use index.
                // But displayColumns is inside renderResults scope. 
                // Easier approach: recreate the mapping or use index if we reconstruct headers.

                // Let's get header names again from DOM to map index
                const headers = Array.from(tableHeader.querySelectorAll('th span:first-child')).map(span => span.textContent);

                for (const [colName, filterVal] of Object.entries(activeColFilters)) {
                    const colIdx = headers.indexOf(colName);
                    if (colIdx !== -1) {
                        const cell = row.cells[colIdx];
                        // Target the specific data span if possible to avoid finding text in badges/tooltips if unnecessary
                        // But simple textContent is robust enough usually.
                        const cellText = cell.textContent.toLowerCase();
                        if (!cellText.includes(filterVal)) {
                            matchesColumns = false;
                            break;
                        }
                    }
                }
            }

            row.style.display = (matchesGlobal && matchesColumns) ? '' : 'none';
        });
    }

    // Update Global Search Listener to use common function
    if (tableSearch) {
        tableSearch.addEventListener('input', applyFilters);
    }

    // Helper
    function formatCurrency(num) {
        return new Intl.NumberFormat('id-ID', {
            style: 'currency',
            currency: 'IDR',
            minimumFractionDigits: 0,
            maximumFractionDigits: 0
        }).format(num);
    }

    function isNumeric(n) {
        return !isNaN(parseFloat(n)) && isFinite(n);
    }

    // --- HELPER: Stateless Download ---
    let currentDownloadHandler = null;

    function setupDownloadButton(filename, base64Data) {
        if (!downloadBtn) return;

        if (currentDownloadHandler) {
            downloadBtn.removeEventListener('click', currentDownloadHandler);
        }

        currentDownloadHandler = (e) => {
            e.preventDefault();
            const b64toBlob = (b64Data, contentType = '', sliceSize = 512) => {
                const byteCharacters = atob(b64Data);
                const byteArrays = [];
                for (let offset = 0; offset < byteCharacters.length; offset += sliceSize) {
                    const slice = byteCharacters.slice(offset, offset + sliceSize);
                    const byteNumbers = new Array(slice.length);
                    for (let i = 0; i < slice.length; i++) {
                        byteNumbers[i] = slice.charCodeAt(i);
                    }
                    const byteArray = new Uint8Array(byteNumbers);
                    byteArrays.push(byteArray);
                }
                return new Blob(byteArrays, { type: contentType });
            }

            const blob = b64toBlob(base64Data, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        };

        downloadBtn.addEventListener('click', currentDownloadHandler);
        downloadBtn.href = "#";
    }

    // --- Theme Toggle ---
    const themeToggleBtn = document.getElementById('themeToggle');
    const sunIcon = document.getElementById('sunIcon');
    const moonIcon = document.getElementById('moonIcon');

    if (themeToggleBtn) {
        // Init: Default to LIGHT (only use dark if explicitly stored)
        const currentTheme = localStorage.getItem('theme');
        if (currentTheme === 'dark') {
            document.documentElement.classList.add('dark');
            sunIcon.classList.remove('hidden');
            moonIcon.classList.add('hidden');
        } else {
            // Default to Light
            document.documentElement.classList.remove('dark');
            sunIcon.classList.add('hidden');
            moonIcon.classList.remove('hidden');
            // Ensure consistency if it was undefined
            if (!currentTheme) localStorage.setItem('theme', 'light');
        }

        themeToggleBtn.addEventListener('click', () => {
            document.documentElement.classList.toggle('dark');
            if (document.documentElement.classList.contains('dark')) {
                localStorage.setItem('theme', 'dark');
                sunIcon.classList.remove('hidden');
                moonIcon.classList.add('hidden');
            } else {
                localStorage.setItem('theme', 'light');
                sunIcon.classList.add('hidden');
                moonIcon.classList.remove('hidden');
            }

            // Hide tooltip immediately on click
            const tooltip = document.getElementById('theme-tooltip');
            if (tooltip) {
                tooltip.classList.add('opacity-0', 'translate-y-4', 'pointer-events-none');
            }
        });

        // Show tooltip after 10 seconds
        setTimeout(() => {
            const tooltip = document.getElementById('theme-tooltip');
            if (tooltip) {
                tooltip.classList.remove('opacity-0', 'translate-y-4', 'pointer-events-none');
                // Auto hide after 5 seconds of showing
                setTimeout(() => {
                    tooltip.classList.add('opacity-0', 'translate-y-4', 'pointer-events-none');
                }, 10000);
            }
        }, 1000);
    }
});
