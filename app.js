// All-in-One File Converter Application Logic

let processedData = null;
let pdfResultData = null;
let toPdfResultData = null;

let dragSourceIndex = null;
let dragType = null;

// --- TAB SWITCHER LOGIC ---
function switchTab(tabName) {
    const wordTab = document.getElementById('wordToExcelTab');
    const pdfTab = document.getElementById('pdfToJpgTab');
    const toPdfTab = document.getElementById('fileToPdfTab');
    
    const wordBtn = document.getElementById('tabWordToExcelBtn');
    const pdfBtn = document.getElementById('tabPdfToJpgBtn');
    const toPdfBtn = document.getElementById('tabFileToPdfBtn');

    const activeClass = "flex-1 py-4 px-4 text-center font-bold text-base transition flex items-center justify-center gap-2 border-b-4 border-indigo-600 text-indigo-700 bg-white shadow-sm";
    const inactiveClass = "flex-1 py-4 px-4 text-center font-bold text-base transition flex items-center justify-center gap-2 border-b-4 border-transparent text-gray-500 hover:text-indigo-600 hover:bg-slate-100";

    wordTab.classList.add('hidden');
    pdfTab.classList.add('hidden');
    toPdfTab.classList.add('hidden');

    wordBtn.className = inactiveClass;
    pdfBtn.className = inactiveClass;
    toPdfBtn.className = inactiveClass;

    if (tabName === 'wordToExcel') {
        wordTab.classList.remove('hidden');
        wordBtn.className = activeClass;
    } else if (tabName === 'pdfToJpg') {
        pdfTab.classList.remove('hidden');
        pdfBtn.className = activeClass;
    } else if (tabName === 'fileToPdf') {
        toPdfTab.classList.remove('hidden');
        toPdfBtn.className = activeClass;
    }
}

// --- WORD TO EXCEL CONVERTER LOGIC ---
document.getElementById('uploadForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    
    const formData = new FormData();
    formData.append('word_file', document.getElementById('wordFile').files[0]);
    formData.append('excel_template', document.getElementById('excelTemplate').files[0]);

    try {
        const response = await fetch('/process', {
            method: 'POST',
            body: formData
        });
        
        const result = await response.json();
        if (response.ok) {
            processedData = result;
            renderPreview(result);
        } else {
            alert(result.error || 'কোনো সমস্যা হয়েছে');
        }
    } catch (error) {
        console.error('Error:', error);
        alert('ফাইল প্রসেস করতে ব্যর্থ হয়েছে।');
    }
});

function renderPreview(result) {
    const head = document.getElementById('tableHead');
    const body = document.getElementById('tableBody');
    
    head.innerHTML = `<tr>${result.columns.map(col => `<th class="px-4 py-3 text-left font-bold text-indigo-900 border-b">${col}</th>`).join('')}</tr>`;
    
    body.innerHTML = result.data.map(row => {
        return `<tr class="hover:bg-slate-50">${result.columns.map(col => `<td class="px-4 py-3 whitespace-nowrap">${row[col] || ''}</td>`).join('')}</tr>`;
    }).join('');

    document.getElementById('previewSection').classList.remove('hidden');
}

document.getElementById('downloadBtn').addEventListener('click', async () => {
    if (!processedData) return;

    try {
        const response = await fetch('/download', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(processedData)
        });

        if (response.ok) {
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = (processedData.subject_title || 'Converted_File') + '.xlsx';
            document.body.appendChild(a);
            a.click();
            a.removeChild(a);
        } else {
            alert('ডাউনলোড করতে সমস্যা হয়েছে।');
        }
    } catch (error) {
        console.error('Error:', error);
    }
});

// --- MULTI-FORMAT TO JPG LOGIC ---
document.getElementById('pdfUploadForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    
    const fileInput = document.getElementById('pdfFileInput');
    const dpiSelect = document.getElementById('dpiSelect');
    
    if (!fileInput.files || fileInput.files.length === 0) {
        alert('দয়া করে অন্তত একটি ফাইল আপলোড করুন');
        return;
    }

    const formData = new FormData();
    for (let i = 0; i < fileInput.files.length; i++) {
        formData.append('pdf_file', fileInput.files[i]);
    }
    formData.append('dpi', dpiSelect.value);

    document.getElementById('pdfLoading').classList.remove('hidden');
    document.getElementById('pdfResultSection').classList.add('hidden');

    try {
        const response = await fetch('/convert_file', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();
        document.getElementById('pdfLoading').classList.add('hidden');

        if (response.ok && result.success) {
            pdfResultData = result;
            renderPdfResults(result);
        } else {
            alert(result.error || 'ফাইল প্রসেস করতে ব্যর্থ হয়েছে');
        }
    } catch (error) {
        document.getElementById('pdfLoading').classList.add('hidden');
        console.error('Conversion Error:', error);
        alert('ফাইল কনভার্ট করতে ত্রুটি হয়েছে।');
    }
});

function renderPdfResults(result) {
    const summaryText = document.getElementById('pdfSummaryText');
    const fileCount = result.total_files || 1;
    const fileLabel = fileCount > 1 ? `মোট ${fileCount}টি ফাইল` : `ফাইল: ${result.pdf_name} [${result.ext || 'FILE'}]`;
    summaryText.innerText = `${fileLabel} | মোট পেজ/ছবি: ${result.total_pages} টি | কোয়ালিটি: ${result.dpi} DPI HD`;

    const imagesGrid = document.getElementById('imagesGrid');
    imagesGrid.innerHTML = '';

    result.pages.forEach((pageData, index) => {
        const card = document.createElement('div');
        card.draggable = true;
        card.dataset.index = index;
        card.className = "bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm hover:shadow-md transition flex flex-col cursor-grab active:cursor-grabbing select-none";
        
        const pageNum = pageData.pageNumber || pageData.page || (index + 1);
        const displayLabel = pageData.label || `ছবি/পেজ ${pageNum}`;

        card.innerHTML = `
            <div class="bg-indigo-50 border-b border-indigo-100 px-3 py-2 flex justify-between items-center text-xs font-bold text-indigo-900">
                <span class="flex items-center gap-1.5 cursor-grab text-indigo-600">
                    <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 8h16M4 16h16"></path></svg>
                    ক্রম: #${index + 1}
                </span>
                <span class="text-indigo-700 bg-indigo-100 px-2 py-0.5 rounded font-mono">পেজ ${pageNum}</span>
            </div>
            <div class="relative group cursor-pointer overflow-hidden bg-gray-100 flex items-center justify-center p-2" onclick="openModal('${pageData.data_url}', '${displayLabel} (${pageData.width} x ${pageData.height} px)')">
                <img src="${pageData.data_url}" alt="${displayLabel}" class="max-h-64 object-contain rounded transition transform group-hover:scale-105 pointer-events-none">
                <div class="absolute inset-0 bg-black/40 opacity-0 group-hover:opacity-100 transition flex items-center justify-center text-white font-bold text-sm pointer-events-none">
                    🔍 বড় করে দেখুন
                </div>
            </div>
            <div class="p-4 flex flex-col justify-between flex-grow border-t border-gray-100">
                <div class="flex justify-between items-center mb-3">
                    <span class="font-bold text-gray-800 text-sm truncate" title="${displayLabel}">${displayLabel}</span>
                    <span class="text-xs text-gray-500 font-mono">${pageData.width}×${pageData.height} px</span>
                </div>
                <button onclick="downloadSingleJpg('${pageData.data_url}', '${pageData.filename || result.pdf_name + '_page_' + pageNum + '.jpg'}')" class="w-full bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-bold py-2 px-3 rounded-lg text-sm transition flex items-center justify-center gap-1">
                    📥 ডাউনলোড JPG
                </button>
            </div>
        `;

        attachDragAndDropEvents(card, index, 'jpg');
        imagesGrid.appendChild(card);
    });

    document.getElementById('pdfResultSection').classList.remove('hidden');
}

function downloadSingleJpg(dataUrl, filename) {
    const a = document.createElement('a');
    a.href = dataUrl;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.removeChild(a);
}

document.getElementById('downloadAllZipBtn').addEventListener('click', async () => {
    if (!pdfResultData || !pdfResultData.pages || pdfResultData.pages.length === 0) return;

    try {
        const response = await fetch('/download_pdf_zip', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                pages: pdfResultData.pages,
                zip_b64: pdfResultData.zip_b64,
                pdf_name: pdfResultData.pdf_name
            })
        });

        if (response.ok) {
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${pdfResultData.pdf_name}_images.zip`;
            document.body.appendChild(a);
            a.click();
            setTimeout(() => {
                document.body.removeChild(a);
                window.URL.revokeObjectURL(url);
            }, 200);
        } else {
            alert('জিপ ডাউনলোড করতে সমস্যা হয়েছে।');
        }
    } catch (error) {
        console.error('ZIP Download Error:', error);
    }
});

// --- MULTI-FORMAT TO PDF LOGIC ---
document.getElementById('toPdfForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    
    const fileInput = document.getElementById('toPdfFileInput');
    if (!fileInput.files || fileInput.files.length === 0) {
        alert('দয়া করে অন্তত একটি ফাইল আপলোড করুন');
        return;
    }

    const formData = new FormData();
    for (let i = 0; i < fileInput.files.length; i++) {
        formData.append('file', fileInput.files[i]);
    }

    document.getElementById('toPdfLoading').classList.remove('hidden');
    document.getElementById('toPdfResultSection').classList.add('hidden');

    try {
        const response = await fetch('/convert_to_pdf', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();
        document.getElementById('toPdfLoading').classList.add('hidden');

        if (response.ok && result.success) {
            toPdfResultData = result;
            renderToPdfResults(result);
        } else {
            alert(result.error || 'পিডিএফ রূপান্তর করতে ব্যর্থ হয়েছে');
        }
    } catch (error) {
        document.getElementById('toPdfLoading').classList.add('hidden');
        console.error('To PDF Error:', error);
        alert('পিডিএফ ফাইল তৈরি করতে ত্রুটি হয়েছে।');
    }
});

function renderToPdfResults(result) {
    const summaryText = document.getElementById('toPdfSummaryText');
    const kbSize = Math.round(result.pdf_size / 1024);
    const fileCount = result.total_files || 1;
    const fileLabel = fileCount > 1 ? `মোট ${fileCount}টি ফাইল থেকে তৈরিকৃত Combined_Document.pdf` : `ফাইল: ${result.file_name}.pdf [${result.ext}]`;
    summaryText.innerText = `${fileLabel} | মোট পেজ: ${result.total_pages} টি | সাইজ: ${kbSize} KB`;

    const pdfPagesGrid = document.getElementById('pdfPagesGrid');
    pdfPagesGrid.innerHTML = '';

    result.preview_pages.forEach((pageData, index) => {
        const card = document.createElement('div');
        card.draggable = true;
        card.dataset.index = index;
        card.className = "bg-white border border-gray-200 rounded-xl overflow-hidden shadow-sm hover:shadow-md transition flex flex-col cursor-grab active:cursor-grabbing select-none";
        
        const pageNum = pageData.pageNumber || pageData.page || (index + 1);

        card.innerHTML = `
            <div class="bg-purple-50 border-b border-purple-100 px-3 py-2 flex justify-between items-center text-xs font-bold text-purple-900">
                <span class="flex items-center gap-1.5 cursor-grab text-purple-600">
                    <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 8h16M4 16h16"></path></svg>
                    ক্রম: #${index + 1}
                </span>
                <span class="text-purple-700 bg-purple-100 px-2 py-0.5 rounded font-mono">পেজ ${pageNum}</span>
            </div>
            <div class="relative group cursor-pointer overflow-hidden bg-gray-100 flex items-center justify-center p-2" onclick="openModal('${pageData.data_url}', 'পিডিএফ পেজ ${pageNum} (${pageData.width} x ${pageData.height} px)')">
                <img src="${pageData.data_url}" alt="Page ${pageNum}" class="max-h-64 object-contain rounded transition transform group-hover:scale-105 pointer-events-none">
                <div class="absolute inset-0 bg-black/40 opacity-0 group-hover:opacity-100 transition flex items-center justify-center text-white font-bold text-sm pointer-events-none">
                    🔍 বড় করে দেখুন
                </div>
            </div>
            <div class="p-3 border-t border-gray-100 text-center bg-slate-50">
                <span class="font-bold text-gray-700 text-sm">পেজ ${pageNum}</span>
            </div>
        `;

        attachDragAndDropEvents(card, index, 'pdf');
        pdfPagesGrid.appendChild(card);
    });

    document.getElementById('toPdfResultSection').classList.remove('hidden');
}

// --- DRAG AND DROP REORDERING LOGIC ---
function attachDragAndDropEvents(card, index, type) {
    card.addEventListener('dragstart', (e) => {
        dragSourceIndex = index;
        dragType = type;
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', index);
        card.classList.add('opacity-40', 'scale-95', 'border-indigo-500');
    });

    card.addEventListener('dragend', () => {
        card.classList.remove('opacity-40', 'scale-95', 'border-indigo-500');
        const selector = type === 'jpg' ? '#imagesGrid > div' : '#pdfPagesGrid > div';
        document.querySelectorAll(selector).forEach(c => {
            c.classList.remove('ring-4', 'ring-indigo-400', 'ring-purple-400', 'bg-indigo-50', 'bg-purple-50');
        });
    });

    card.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        const ringColor = type === 'jpg' ? 'ring-indigo-400' : 'ring-purple-400';
        const bgColor = type === 'jpg' ? 'bg-indigo-50' : 'bg-purple-50';
        card.classList.add('ring-4', ringColor, bgColor);
    });

    card.addEventListener('dragleave', () => {
        const ringColor = type === 'jpg' ? 'ring-indigo-400' : 'ring-purple-400';
        const bgColor = type === 'jpg' ? 'bg-indigo-50' : 'bg-purple-50';
        card.classList.remove('ring-4', ringColor, bgColor);
    });

    card.addEventListener('drop', (e) => {
        e.preventDefault();
        const targetIndex = index;
        if (dragSourceIndex === null || dragSourceIndex === targetIndex || dragType !== type) {
            return;
        }
        reorderPages(dragSourceIndex, targetIndex, type);
    });
}

function reorderPages(fromIndex, toIndex, type) {
    if (type === 'jpg' && pdfResultData && pdfResultData.pages) {
        const pages = pdfResultData.pages;
        const [movedPage] = pages.splice(fromIndex, 1);
        pages.splice(toIndex, 0, movedPage);

        const startIndex = Math.min(fromIndex, toIndex);
        for (let i = startIndex; i < pages.length; i++) {
            const newSerial = i + 1;
            pages[i].pageNumber = newSerial;
            pages[i].page = newSerial;
            
            if (pdfResultData.total_files > 1 && pages[i].source_file) {
                const fileBase = pages[i].source_file.split('.')[0];
                pages[i].label = `${fileBase} - পেজ ${newSerial}`;
            } else {
                pages[i].label = `পেজ/ছবি ${newSerial}`;
            }
        }

        renderPdfResults(pdfResultData);

    } else if (type === 'pdf' && toPdfResultData && toPdfResultData.preview_pages) {
        const pages = toPdfResultData.preview_pages;
        const [movedPage] = pages.splice(fromIndex, 1);
        pages.splice(toIndex, 0, movedPage);

        const startIndex = Math.min(fromIndex, toIndex);
        for (let i = startIndex; i < pages.length; i++) {
            const newSerial = i + 1;
            pages[i].pageNumber = newSerial;
            pages[i].page = newSerial;
        }

        renderToPdfResults(toPdfResultData);
    }
}

// Download PDF File (in current Drag & Drop order)
document.getElementById('downloadPdfFileBtn').addEventListener('click', async () => {
    if (!toPdfResultData || !toPdfResultData.preview_pages || toPdfResultData.preview_pages.length === 0) {
        alert('কোনো ডাউনলোডেবল ডাটা পাওয়া যায়নি');
        return;
    }

    try {
        const response = await fetch('/download_pdf_file', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                pages: toPdfResultData.preview_pages,
                pdf_b64: toPdfResultData.pdf_b64,
                file_name: toPdfResultData.file_name
            })
        });

        if (response.ok) {
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);

            const cleanName = (toPdfResultData.file_name || 'converted_document')
                .replace(/\s+/g, ' ')
                .replace(/[\\/*?:"<>|]/g, '')
                .trim() + '.pdf';

            const a = document.createElement('a');
            a.href = url;
            a.download = cleanName;
            document.body.appendChild(a);
            a.click();
            
            setTimeout(() => {
                document.body.removeChild(a);
                window.URL.revokeObjectURL(url);
            }, 200);
        } else {
            alert('পিডিএফ ফাইল ডাউনলোড করতে সমস্যা হয়েছে।');
        }
    } catch (error) {
        console.error('PDF Download Error:', error);
        alert('পিডিএফ ফাইল ডাউনলোড করতে ত্রুটি হয়েছে।');
    }
});

// --- LIGHTBOX MODAL HANDLERS ---
function openModal(imgSrc, caption) {
    document.getElementById('modalImage').src = imgSrc;
    document.getElementById('modalCaption').innerText = caption;
    document.getElementById('imageModal').classList.remove('hidden');
}

function closeModal() {
    document.getElementById('imageModal').classList.add('hidden');
}

document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeModal();
});
