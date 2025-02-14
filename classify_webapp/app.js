pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.11.338/pdf.worker.min.js';

// Add debug logging
console.log('App.js loaded');

let currentPdfPath = null;

async function loadPdfList() {
    try {
        console.log('Fetching PDF list...');
        const response = await fetch('/list-pdfs');
        const files = await response.json();
        console.log('PDF files:', files);
        const fileList = document.getElementById('file-list');
        fileList.innerHTML = '';

        files.forEach(file => {
            const div = document.createElement('div');
            div.className = 'pdf-link';
            div.textContent = file;
            div.onclick = () => {
                // Remove active class from all items
                document.querySelectorAll('.pdf-link').forEach(item => {
                    item.classList.remove('active');
                });
                // Add active class to clicked item
                div.classList.add('active');
                loadPdf(file);
            };
            fileList.appendChild(div);
        });
    } catch (e) {
        console.error('Error loading PDF list:', e);
    }
}

async function loadPdf(filename) {
    currentPdfPath = filename;
    const pdfContainer = document.getElementById('pdf-container');
    pdfContainer.innerHTML = ''; // Clear existing content

    try {
        // Load PDF using PDF.js
        const loadingTask = pdfjsLib.getDocument(`/pdf_downloads/${filename}`);
        const pdf = await loadingTask.promise;
        const jsonname = String(filename).slice(0, -4) + '_gemini_classification.json';

        // Render all pages
        for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
            const page = await pdf.getPage(pageNum);
            const canvas = document.createElement('canvas');
            const context = canvas.getContext('2d');

            // Use a larger scale for better readability
            const viewport = page.getViewport({ scale: 1.5 });

            canvas.height = viewport.height;
            canvas.width = viewport.width;
            canvas.style.marginBottom = '20px'; // Add space between pages

            await page.render({
                canvasContext: context,
                viewport: viewport
            }).promise;

            pdfContainer.appendChild(canvas);
        }

        // Load JSON if exists
        const json = await loadJson(jsonname);
    } catch (error) {
        console.error('Error loading PDF:', error);
        pdfContainer.innerHTML = '<p>Error loading PDF</p>';
    }
}

async function loadJson(jsonFile) {
    try {
        const response = await fetch(`/pdf_downloads/${jsonFile}`);
        const data = await response.json();

        document.getElementById('item-type').value = data.item_type;
        document.getElementById('copyright-status').value = data.copyright_status;

        // Display all JSON data with proper HTML encoding
        const jsonContent = document.getElementById('json-content');
        jsonContent.innerHTML = Object.entries(data)
            .filter(([key]) => !['item_type', 'copyright_status'].includes(key))
            .map(([key, value]) => {
                const formattedValue = Array.isArray(value)
                    ? `<div class="array-content">
                         ${value.map(v => `<div class="array-item">${escapeHtml(v)}</div>`).join('')}
                       </div>`
                    : escapeHtml(value || '(empty)');

                return `
                    <div class="json-field">
                        <strong>${escapeHtml(key)}:</strong>
                        ${formattedValue}
                    </div>
                `;
            }).join('');
        return data;
    } catch (e) {
        console.error('Error loading JSON:', e);
    }
}

// Add HTML escaping function
function escapeHtml(unsafe) {
    if (unsafe === null || unsafe === undefined) return '';
    return String(unsafe)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

document.getElementById('save-changes').onclick = async () => {
    if (!currentPdfPath) return;

    const jsonPath = String(currentPdfPath).slice(0, -4) + '_gemini_classification.json';

    const data = {
        item_type: document.getElementById('item-type').value,
        copyright_status: document.getElementById('copyright-status').value
    };

    try {
        await fetch(`/save-json/${jsonPath}`, {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(data)
        });
        alert('Changes saved successfully!');
    } catch (e) {
        alert('Error saving changes');
        console.error(e);
    }
};

// Initial load
loadPdfList();
