pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.11.338/pdf.worker.min.js';

// Add debug logging
console.log('App.js loaded');

let currentPdfPath = null;
let currentScale = 1.0; // track zoom level

function getColorClassForStatus(status) {
  if (status === 'own material' || status === 'open access') return 'pdf-link-green';
  if (status === 'copyrighted material') return 'pdf-link-red';
  return '';
}

async function loadPdfList() {
    try {
        console.log('Fetching PDF list...');
        const response = await fetch('/list-pdfs');
        const files = await response.json();
        //files is a json object with a list of items
        // each item has 2 keys: pdf, json
        // pdf has the pdf filename
        // json has the json filename
        console.log('PDF list:', files);
        // split the files into pdfs and jsons
        let pdfs = [];
        let jsons = [];
        files.forEach(file => {
            pdfs.push(file.pdf);
            jsons.push(file.json);
        });

        console.log('PDF files:', pdfs);
        const fileList = document.getElementById('file-list');
        fileList.innerHTML = '';

        files.forEach(file => {
            const div = document.createElement('div');
            div.className = `pdf-link ${getColorClassForStatus(file.json_data?.copyright_status)}`;
            div.textContent = file.pdf;
            div.onclick = () => {
                // Remove active class from all items
                document.querySelectorAll('.pdf-link').forEach(item => {
                    item.classList.remove('active');
                });
                // Add active class to clicked item
                div.classList.add('active');
                loadPdf(file.pdf, file.json);
            };
            fileList.appendChild(div);
        });


    } catch (e) {
        console.error('Error loading PDF list:', e);
    }
}

async function loadPdf(filename, jsonfilename) {
    currentPdfPath = filename;
    currentjsonpath = jsonfilename;
    const pdfContainer = document.getElementById('pdf-pages');
    const loadingIndicator = document.getElementById('loading-indicator');
    pdfContainer.innerHTML = '';
    loadingIndicator.style.display = 'flex';

    try {
        const loadingTask = pdfjsLib.getDocument(`/pdf_downloads/${filename}`);
        const pdf = await loadingTask.promise;


        // Create viewport observer for lazy loading
        const viewportObserver = new IntersectionObserver((entries, observer) => {
            entries.forEach(entry => {
                if (entry.isIntersecting) {
                    const pageNum = parseInt(entry.target.dataset.pageNum);
                    renderPage(pdf, pageNum, entry.target);
                    observer.unobserve(entry.target);
                }
            });
        }, { root: null, rootMargin: '100px' });

        // Create placeholder divs for all pages
        for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
            const placeholder = document.createElement('div');
            placeholder.className = 'pdf-page-placeholder';
            placeholder.dataset.pageNum = pageNum;
            placeholder.style.height = '1000px'; // Approximate height
            placeholder.style.width = '100%';
            placeholder.style.display = 'flex';
            placeholder.style.justifyContent = 'center';
            pdfContainer.appendChild(placeholder);
            viewportObserver.observe(placeholder);
        }

        loadingIndicator.style.display = 'none';
        await loadJson();
    } catch (error) {
        console.error('Error loading PDF:', error);
        pdfContainer.innerHTML = '<p>Error loading PDF</p>';
        loadingIndicator.style.display = 'none';
    }
}

async function renderPage(pdf, pageNum, container) {
    try {
        const page = await pdf.getPage(pageNum);
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        const viewport = page.getViewport({ scale: currentScale });
        canvas.height = viewport.height;
        canvas.width = viewport.width;
        canvas.style.maxWidth = '100%';
        canvas.style.height = 'auto';

        await page.render({
            canvasContext: context,
            viewport: viewport
        }).promise;

        container.innerHTML = '';
        container.style.height = 'auto';
        container.appendChild(canvas);
    } catch (error) {
        console.error(`Error rendering page ${pageNum}:`, error);
        container.innerHTML = `<p>Error loading page ${pageNum}</p>`;
    }
}

async function loadJson() {
    try {
        const response = await fetch(`/pdf_downloads/${currentjsonpath}`);
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
                         ${value.map(v => `<div class="array-item">${v}</div>`).join('')}
                       </div>`
                    : value || '(empty)';

                return `
                    <div class="json-field">
                        <strong>${key.replace('_', ' ')}:</strong><br>
                        ${formattedValue}
                    </div>
                `;
            }).join('');
        console.log('JSON data:', data);
        return data;
    } catch (e) {
        console.error('Error loading JSON:', e);
    }
}

document.getElementById('zoom-in').onclick = () => {
    currentScale += 0.2;
    reloadPages();
};

document.getElementById('zoom-out').onclick = () => {
    currentScale = Math.max(0.2, currentScale - 0.2);
    reloadPages();
};

function reloadPages() {
    const pdfContainer = document.getElementById('pdf-pages');
    // Clear existing
    pdfContainer.innerHTML = '';
    // Re-load with updated scale
    if (currentPdfPath) {
        loadPdf(currentPdfPath, currentjsonpath);
    }
}



// Collapse sidebars
document.getElementById('collapse-left').onclick = () => {
    document.body.classList.toggle('collapsed-left');
};
document.getElementById('collapse-right').onclick = () => {
    document.body.classList.toggle('collapsed-right');
};

document.getElementById('save-changes').onclick = async () => {
    if (!currentPdfPath) return;

    const jsonPath = currentjsonpath;

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
        const activeLink = document.querySelector('.pdf-link.active');
        if (activeLink) {
            const newStatus = data.copyright_status;
            activeLink.className = `pdf-link active ${getColorClassForStatus(newStatus)}`;
        }
    } catch (e) {
        alert('Error saving changes');
        console.error(e);
    }

    loadPdfList();
};

// Initial load
loadPdfList();
