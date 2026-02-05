"""
Web App for Excel Template Generator
Responsive Flask app that works on mobile and desktop
"""

from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.middleware.proxy_fix import ProxyFix
import os
import json
from datetime import datetime
from excel_generator import ExcelTemplateGenerator
from werkzeug.utils import secure_filename
import traceback
import re

def safe_filename(filename):
    """Create safe filename while preserving Unicode characters"""
    # Remove leading/trailing whitespace
    filename = filename.strip()
    # Replace spaces with underscores
    filename = filename.replace(' ', '_')
    # Remove only truly problematic characters: / \ : * ? " < > |
    filename = re.sub(r'[/\\:*?"<>|]', '', filename)
    # Remove leading dots
    filename = filename.lstrip('.')

    # Only use secure_filename for ASCII characters, otherwise keep Unicode
    # Check if the filename contains non-ASCII characters
    try:
        filename.encode('ascii')
        # If it's all ASCII, use secure_filename for extra safety
        from werkzeug.utils import secure_filename
        filename = secure_filename(filename)
    except UnicodeEncodeError:
        # If it contains Unicode characters, skip secure_filename to preserve them
        pass

    return filename if filename else 'output'

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size
app.config['UPLOAD_FOLDER'] = 'uploads'

# Apply ProxyFix to handle requests coming through reverse proxies/tunnels
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1)

# Create uploads folder
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Global generator instance
generator = None

def init_generator():
    """Initialize generator with template file"""
    global generator
    if os.path.exists('template.xlsx'):
        generator = ExcelTemplateGenerator('template.xlsx')
    else:
        raise FileNotFoundError("template.xlsx not found. Please ensure template.xlsx is in the app directory.")

@app.route('/')
def index():
    """Main page"""
    return render_template('index.html')

@app.route('/api/generate', methods=['POST'])
def generate():
    """Generate Excel and optionally PDF"""
    try:
        if generator is None:
            init_generator()
        
        # Get form data
        data = request.json
        
        company_name = data.get('company_name', '').strip()
        sakadastro = data.get('sakadastro', '').strip()
        address = data.get('address', '').strip()
        invoice_number = data.get('invoice_number', '').strip()
        output_filename = data.get('output_filename', 'invoice').strip()
        generate_pdf = data.get('generate_pdf', True)  # Default to True to always generate PDF
        
        # Validate required fields
        if not all([company_name, sakadastro, address, invoice_number, output_filename]):
            return jsonify({'error': 'All required fields must be filled'}), 400
        
        # Parse items
        items = []
        items_data = data.get('items', [])
        for item in items_data:
            if item.get('type'):
                try:
                    quantity = float(item.get('quantity', 0)) if item.get('quantity') else ''
                    price = float(item.get('price', 0)) if item.get('price') else ''
                except ValueError:
                    return jsonify({'error': f'Invalid quantity or price for item: {item.get("type")}'}), 400
                
                items.append({
                    'type': item.get('type'),
                    'quantity': quantity,
                    'price': price
                })
        
        # Create safe filename - keep the alphanumeric and basic chars, handle dots properly
        # Use custom safe_filename function that preserves Unicode
        base_name = safe_filename(output_filename)
        # Remove any trailing dots/extensions that secure_filename might have removed
        base_name = base_name.replace('.xlsx', '').replace('.xls', '')
        # If empty, use default
        if not base_name:
            base_name = 'output'
        
        filename_with_ext = base_name + '.xlsx'
        
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], filename_with_ext)
        
        # Generate Excel
        generator.generate(
            excel_path,
            company_name=company_name,
            sakadastro=sakadastro,
            address=address,
            invoice_number=invoice_number,
            items=items
        )
        
        # Generate PDF if requested
        pdf_path = None
        if generate_pdf:
            pdf_filename = filename_with_ext.replace('.xlsx', '.pdf')
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_filename)
            generator.generate_pdf(excel_path, pdf_path)
        
        return jsonify({
            'success': True,
            'excel_file': filename_with_ext,
            'pdf_file': os.path.basename(pdf_path) if pdf_path else None,
            'message': 'Files generated successfully!'
        })
    
    except Exception as e:
        print(f"Error: {traceback.format_exc()}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/download/<path:filename>')
def download(filename):
    """Download generated file"""
    try:
        filename = safe_filename(filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)

        print(f"Download request for: {filename}")
        print(f"Full path: {file_path}")
        print(f"File exists: {os.path.exists(file_path)}")

        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            # Return a simple error message instead of JSON for direct browser access
            return f"File not found: {filename}", 404

        # Determine mimetype
        if filename.endswith('.pdf'):
            mimetype = 'application/pdf'
        elif filename.endswith('.xlsx'):
            mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        else:
            mimetype = 'application/octet-stream'

        print(f"Sending file with mimetype: {mimetype}")

        # Send file with attachment headers to force download
        # This should work for all browsers including iOS Safari
        return send_file(
            file_path,
            mimetype=mimetype,
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        print(f"Download error: {e}")
        print(traceback.format_exc())
        # Return a simple error message instead of JSON
        return f"Download error: {str(e)}", 500

@app.route('/health')
def health():
    """Health check endpoint"""
    return jsonify({'status': 'ok'})

if __name__ == '__main__':
    try:
        init_generator()
        # Run on 0.0.0.0 to allow access from mobile devices on same network
        app.run(debug=True, host='0.0.0.0', port=5000)
    except FileNotFoundError as e:
        print(f"Error: {e}")
        print("Make sure template.xlsx exists in the jtrade directory")
