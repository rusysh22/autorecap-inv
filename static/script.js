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

    // Inputs
    const filenameSuffixInput = document.getElementById('filename-suffix');

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

    // --- State ---
    let selectedFiles = [];


    // ... (drag & drop handlers) ...

    if (resetBtn) {
        resetBtn.addEventListener('click', () => {
            selectedFiles = [];
            renderFileList(); // Updates UI and hides config panel
            resultsSection.classList.add('hidden');
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

    // --- File Handling ---

    function handleFiles(files) {
        const newFiles = Array.from(files).filter(file =>
            file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
        );

        if (newFiles.length === 0) return;

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



    // --- API & Processing ---

    processBtn.addEventListener('click', async () => {
        if (selectedFiles.length === 0) return;

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

        try {
            const response = await fetch('/api/process', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            if (result.success) {
                renderResults(result);
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

    function renderResults(result) {
        // Summary Cards
        summaryFiles.textContent = result.summary.total_files;
        summaryRows.textContent = result.summary.total_rows;
        summaryAmount.textContent = formatCurrency(result.summary.total_amount);

        // Download Link
        downloadBtn.href = result.summary.download_url;

        // Output Filename Display
        if (outputFilenameDisplay && result.summary.output_filename) {
            outputFilenameDisplay.textContent = result.summary.output_filename;
        }

        // Color Palette for files
        const colors = [
            { bg: 'bg-red-50', border: 'border-red-500', text: 'text-red-700', badge: 'bg-red-100' },
            { bg: 'bg-orange-50', border: 'border-orange-500', text: 'text-orange-700', badge: 'bg-orange-100' },
            { bg: 'bg-amber-50', border: 'border-amber-500', text: 'text-amber-700', badge: 'bg-amber-100' },
            { bg: 'bg-green-50', border: 'border-green-500', text: 'text-green-700', badge: 'bg-green-100' },
            { bg: 'bg-emerald-50', border: 'border-emerald-500', text: 'text-emerald-700', badge: 'bg-emerald-100' },
            { bg: 'bg-teal-50', border: 'border-teal-500', text: 'text-teal-700', badge: 'bg-teal-100' },
            { bg: 'bg-cyan-50', border: 'border-cyan-500', text: 'text-cyan-700', badge: 'bg-cyan-100' },
            { bg: 'bg-sky-50', border: 'border-sky-500', text: 'text-sky-700', badge: 'bg-sky-100' },
            { bg: 'bg-blue-50', border: 'border-blue-500', text: 'text-blue-700', badge: 'bg-blue-100' },
            { bg: 'bg-indigo-50', border: 'border-indigo-500', text: 'text-indigo-700', badge: 'bg-indigo-100' },
            { bg: 'bg-violet-50', border: 'border-violet-500', text: 'text-violet-700', badge: 'bg-violet-100' },
            { bg: 'bg-purple-50', border: 'border-purple-500', text: 'text-purple-700', badge: 'bg-purple-100' },
            { bg: 'bg-fuchsia-50', border: 'border-fuchsia-500', text: 'text-fuchsia-700', badge: 'bg-fuchsia-100' },
            { bg: 'bg-pink-50', border: 'border-pink-500', text: 'text-pink-700', badge: 'bg-pink-100' },
            { bg: 'bg-rose-50', border: 'border-rose-500', text: 'text-rose-700', badge: 'bg-rose-100' }
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
                const statusClass = isError
                    ? 'text-red-600 bg-red-50 border border-red-100 px-2 py-0.5 rounded-full text-xs font-semibold'
                    : 'text-green-600 bg-green-50 border border-green-100 px-2 py-0.5 rounded-full text-xs font-semibold';

                return `
                    <tr class="hover:bg-slate-50 transition-colors">
                        <td class="px-6 py-4 font-medium text-slate-700">
                             <div class="flex items-center gap-3">
                                <div class="w-3 h-3 rounded-full flex-shrink-0 ${color.border.replace('border', 'bg')}" title="Color Tag"></div>
                                <div class="truncate max-w-[250px]" title="${f.filename}">${f.filename}</div>
                            </div>
                        </td>
                        <td class="px-6 py-4 text-center">
                            <span class="${statusClass}">${f.status}</span>
                        </td>
                        <td class="px-6 py-4 text-right text-slate-600">${f.rows}</td>
                        <td class="px-6 py-4 text-right font-mono text-slate-500 text-xs">${formatCurrency(f.ppn)}</td>
                        <td class="px-6 py-4 text-right font-mono text-slate-500 text-xs">${formatCurrency(f.pph)}</td>
                        <td class="px-6 py-4 text-right font-mono text-slate-700 font-bold">${formatCurrency(f.amount)}</td>
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
            `<th scope="col" class="px-6 py-3 font-semibold tracking-wide whitespace-nowrap bg-slate-100">
                <div class="flex flex-col gap-2">
                    <span>${col}</span>
                    <input type="text" class="column-filter bg-white border border-slate-300 text-slate-900 text-xs rounded-lg focus:ring-blue-500 focus:border-blue-500 block w-full p-1.5 font-normal" placeholder="Filter ${col}..." data-col="${col}">
                </div>
            </th>`
        ).join('');

        tableBody.innerHTML = data.map(row => {
            const filename = row['source_file'];
            const color = fileColorMap[filename] || colors[0];

            return `<tr class="bg-white hover:bg-slate-50 transition-colors border-b last:border-0 border-slate-100">
                ${displayColumns.map((col, idx) => {
                const cellValue = row[col] !== null ? row[col] : '';
                // Inject filename badge in the first cell
                if (idx === 0) {
                    return `<td class="px-6 py-4 whitespace-nowrap max-w-[200px]" title="${cellValue}">
                            <div class="flex flex-col">
                                <span class="cell-data truncate block ${isNumeric(row[col]) ? 'font-mono text-right' : ''}">${cellValue}</span>
                                <span class="text-[10px] ${color.text} mt-1 font-medium px-1.5 py-0.5 rounded ${color.badge} w-fit opacity-75 max-w-full truncate" title="${filename}">${filename}</span>
                            </div>
                         </td>`;
                }
                return `<td class="px-6 py-4 whitespace-nowrap max-w-[200px] truncate ${isNumeric(row[col]) ? 'font-mono text-right' : ''}" title="${cellValue}">
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
});
