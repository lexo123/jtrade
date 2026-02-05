// Item management
let itemCount = 0;

document.getElementById('addItemBtn').addEventListener('click', addItem);
document.getElementById('invoiceForm').addEventListener('submit', handleSubmit);
document.getElementById('newInvoiceBtn').addEventListener('click', resetForm);

function addItem() {
    itemCount++;
    const template = document.getElementById('itemTemplate');
    const clone = template.content.cloneNode(true);
    
    // Set item number
    clone.querySelector('.item-number').textContent = `Item ${itemCount}`;
    
    // Add remove event listener
    clone.querySelector('.btn-remove').addEventListener('click', function(e) {
        e.preventDefault();
        this.parentElement.remove();
    });
    
    document.getElementById('itemsContainer').appendChild(clone);
}

function getItems() {
    const items = [];
    document.querySelectorAll('.item-card').forEach(card => {
        const type = card.querySelector('.item-type').value.trim();
        const quantity = card.querySelector('.item-quantity').value;
        const price = card.querySelector('.item-price').value;
        
        if (type) {
            items.push({
                type: type,
                quantity: quantity ? parseFloat(quantity) : 0,
                price: price ? parseFloat(price) : 0
            });
        }
    });
    return items;
}

function showMessage(message, type) {
    const statusEl = document.getElementById('statusMessage');
    statusEl.textContent = message;
    statusEl.className = `status-message ${type}`;
    
    if (type === 'error') {
        setTimeout(() => {
            statusEl.className = 'status-hidden';
        }, 5000);
    }
}

function handleSubmit(e) {
    e.preventDefault();
    
    // Validate required fields
    const company_name = document.getElementById('company_name').value.trim();
    const sakadastro = document.getElementById('sakadastro').value.trim();
    const address = document.getElementById('address').value.trim();
    const invoice_number = document.getElementById('invoice_number').value.trim();
    const output_filename = document.getElementById('output_filename').value.trim();
    
    if (!company_name || !sakadastro || !address || !invoice_number || !output_filename) {
        showMessage('❌ Please fill in all required fields', 'error');
        return;
    }
    
    const submitBtn = document.querySelector('.btn-primary');
    submitBtn.disabled = true;
    submitBtn.innerHTML = '<span class="spinner"></span> Generating...';
    
    showMessage('⏳ Generating your invoice...', 'loading');
    
    const data = {
        company_name: company_name,
        sakadastro: sakadastro,
        address: address,
        invoice_number: invoice_number,
        output_filename: output_filename,
        generate_pdf: true,  // Always generate PDF
        items: getItems()
    };
    
    fetch('/api/generate', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(data)
    })
    .then(response => {
        if (!response.ok) {
            return response.json().then(data => {
                throw new Error(data.error || 'An error occurred');
            });
        }
        return response.json();
    })
    .then(data => {
        submitBtn.disabled = false;
        submitBtn.textContent = 'Generate Invoice';
        showMessage('✅ ' + data.message, 'success');
        
        // Show results
        displayResults(data);
    })
    .catch(error => {
        submitBtn.disabled = false;
        submitBtn.textContent = 'Generate Invoice';
        showMessage('❌ Error: ' + error.message, 'error');
        console.error('Error:', error);
    });
}

function displayResults(data) {
    const resultsSection = document.getElementById('results');
    const excelFile = document.getElementById('excelFile');
    const pdfFile = document.getElementById('pdfFile');
    const excelFileName = document.getElementById('excelFileName');
    const excelDownload = document.getElementById('excelDownload');
    const pdfFileName = document.getElementById('pdfFileName');
    const pdfDownload = document.getElementById('pdfDownload');
    
    // Set Excel file
    excelFileName.textContent = data.excel_file;
    excelDownload.href = '#';
    excelDownload.onclick = function(e) {
        e.preventDefault();
        downloadFile(data.excel_file);
    };
    
    // Handle PDF file
    if (data.pdf_file) {
        pdfFileName.textContent = data.pdf_file;
        pdfDownload.href = '#';
        pdfDownload.onclick = function(e) {
            e.preventDefault();
            downloadFile(data.pdf_file);
        };
        pdfFile.classList.remove('hidden');
    } else {
        pdfFile.classList.add('hidden');
    }
    
    // Show results section
    resultsSection.classList.add('show');
    document.querySelector('.form-container').scrollIntoView({ behavior: 'smooth' });
}

function downloadFile(filename) {
    console.log('Attempting to download:', filename);

    const url = `/api/download/${encodeURIComponent(filename)}`;
    console.log('Download URL:', url);

    // Detect iOS devices (iPad, iPhone, iPod)
    const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent) && !window.MSStream;

    // For PDF files on iOS devices, we need to handle the download differently
    if (filename.toLowerCase().endsWith('.pdf') && isIOS) {
        // For iOS devices, open the URL directly which should trigger download due to server-side headers
        // Using a temporary iframe to avoid leaving the page
        const iframe = document.createElement('iframe');
        iframe.style.display = 'none';
        iframe.src = url;
        document.body.appendChild(iframe);
        
        // Remove the iframe after a delay
        setTimeout(() => {
            document.body.removeChild(iframe);
        }, 1000);
        
        console.log('iOS PDF download triggered via iframe');
        return;
    }

    // Fetch the file for other cases
    fetch(url)
        .then(response => {
            console.log('Response status:', response.status);
            if (!response.ok) {
                return response.json().then(data => {
                    throw new Error(data.error || `HTTP error! status: ${response.status}`);
                });
            }
            return response.blob();
        })
        .then(blob => {
            console.log('Blob received, size:', blob.size);
            // Create blob URL and trigger download
            const blobUrl = window.URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = blobUrl;
            link.download = filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            window.URL.revokeObjectURL(blobUrl);
            console.log('Download triggered');
        })
        .catch(error => {
            console.error('Download error:', error);
            showMessage('❌ Download failed: ' + error.message, 'error');
        });
}

function resetForm() {
    document.getElementById('invoiceForm').reset();
    document.getElementById('itemsContainer').innerHTML = '';
    itemCount = 0;
    document.getElementById('results').classList.remove('show');
    document.getElementById('statusMessage').className = 'status-hidden';
    document.getElementById('company_name').focus();
}
