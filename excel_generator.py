"""
Excel Template Generator
Reads template.xls, applies specified cell changes, and generates new Excel files.
"""

import os
from datetime import datetime
from openpyxl import load_workbook as openpyxl_load_workbook
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from copy import copy
from pathlib import Path
import xlrd
import subprocess


class ExcelTemplateGenerator:
    def __init__(self, template_path):
        """
        Initialize with template file path.
        
        Args:
            template_path (str): Path to the template Excel file (.xls or .xlsx)
        """
        self.template_path = template_path
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template file not found: {template_path}")
    
    def _load_workbook(self, file_path):
        """
        Load workbook from either .xls or .xlsx format.
        Converts .xls to .xlsx preserving all formatting.
        
        Args:
            file_path (str): Path to Excel file
            
        Returns:
            openpyxl Workbook object
        """
        if file_path.lower().endswith('.xls'):
            # Use pandas to read and convert .xls to .xlsx with formatting preserved
            try:
                import pandas as pd
                import tempfile
                
                # Read all sheets from .xls file
                xls_file = pd.ExcelFile(file_path)
                
                # Create a temporary .xlsx file
                temp_xlsx = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
                temp_xlsx_path = temp_xlsx.name
                temp_xlsx.close()
                
                # Write to temporary .xlsx (pandas preserves basic formatting)
                with pd.ExcelWriter(temp_xlsx_path, engine='openpyxl') as writer:
                    for sheet_name in xls_file.sheet_names:
                        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                
                # Load the temporary .xlsx file
                wb = openpyxl_load_workbook(temp_xlsx_path)
                
                # Store temp file path for cleanup later
                self._temp_file = temp_xlsx_path
                
                return wb
            except ImportError:
                # Fallback if pandas not available
                print("Warning: pandas not available, using basic conversion (some formatting may be lost)")
                return self._load_workbook_basic(file_path)
        else:
            # Load as .xlsx directly
            return openpyxl_load_workbook(file_path)
    
    def _load_workbook_basic(self, file_path):
        """
        Basic fallback conversion from .xls to .xlsx (minimal formatting)
        """
        xls_book = xlrd.open_workbook(file_path)
        xls_sheet = xls_book.sheet_by_index(0)
        
        xlsx_wb = Workbook()
        xlsx_ws = xlsx_wb.active
        
        # Copy all cells from xls to xlsx
        for row_idx in range(xls_sheet.nrows):
            for col_idx in range(xls_sheet.ncols):
                cell_value = xls_sheet.cell_value(row_idx, col_idx)
                if cell_value:
                    xlsx_ws.cell(row=row_idx + 1, column=col_idx + 1, value=cell_value)
        
        return xlsx_wb
    
    def generate(self, output_path, company_name, sakadastro, address, invoice_number, changes=None, items=None):
        """
        Generate a new Excel file based on template with specified changes.
        
        Args:
            output_path (str): Path where the new Excel file will be saved
            company_name (str): Company name for cell A12
            sakadastro (str): Sakadastro value for cell A13
            address (str): Address for cell A14
            invoice_number (str/int): Invoice number for cell D5
            changes (dict): Additional cell changes (optional)
                          Example: {'A1': 'New Value', 'B2': 123, 'C3': 45.67}
            items (list): Optional items/types for rows 17-24
                         Each item is a dict with keys: 'type', 'quantity', 'price'
                         Example: [{'type': 'Service A', 'quantity': 2, 'price': 100},
                                  {'type': 'Service B', 'quantity': 1, 'price': 50}]
        
        Returns:
            str: Path to the generated file
        """
        if changes is None:
            changes = {}
        if items is None:
            items = []
        
        # Load the template
        wb = self._load_workbook(self.template_path)
        ws = wb.active
        
        # Always set required fields
        ws['D4'] = datetime.now()  # Current date
        ws['A12'] = company_name
        ws['A13'] = sakadastro
        ws['A14'] = address
        ws['D5'] = invoice_number
        
        # Fill items in rows 17-24
        start_row = 17
        for i, item in enumerate(items):
            if i >= 8:  # Only 8 rows available (17-24)
                break
            
            row = start_row + i
            if isinstance(item, dict):
                ws[f'A{row}'] = item.get('type', '')
                ws[f'B{row}'] = item.get('quantity', '')
                ws[f'C{row}'] = item.get('price', '')
                # Set formula for D column (sum_price = quantity * price)
                ws[f'D{row}'] = f'=B{row}*C{row}'
            else:
                # Support tuple format (type, quantity, price)
                ws[f'A{row}'] = item[0] if len(item) > 0 else ''
                ws[f'B{row}'] = item[1] if len(item) > 1 else ''
                ws[f'C{row}'] = item[2] if len(item) > 2 else ''
                # Set formula for D column
                ws[f'D{row}'] = f'=B{row}*C{row}'
        
        # Set D36 to sum of D17:D24
        ws['D36'] = '=SUM(D17:D24)'
        
        # Apply additional changes to each specified cell
        for cell_ref, value in changes.items():
            try:
                cell = ws[cell_ref]
                cell.value = value
            except Exception as e:
                print(f"Error setting cell {cell_ref}: {e}")
        
        # Create output directory if it doesn't exist
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        # Save the modified workbook
        wb.save(output_path)
        print(f"Excel file generated: {output_path}")
        
        # Clean up temporary file if it was created
        if hasattr(self, '_temp_file') and os.path.exists(self._temp_file):
            try:
                os.remove(self._temp_file)
            except:
                pass
        
        return output_path
    
    def generate_pdf(self, excel_path, pdf_path=None):
        """
        Convert an Excel file to PDF using LibreOffice.
        
        Args:
            excel_path (str): Path to the Excel file to convert
            pdf_path (str): Optional path for the PDF file. If not provided, 
                           will use same name as Excel file with .pdf extension
        
        Returns:
            str: Path to the generated PDF file
        """
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Excel file not found: {excel_path}")
        
        if pdf_path is None:
            # Generate PDF path from Excel path
            pdf_path = os.path.splitext(excel_path)[0] + '.pdf'
        
        try:
            # Use LibreOffice to convert Excel to PDF
            # --headless: Run without GUI
            # --convert-to pdf: Convert to PDF
            # --outdir: Output directory
            
            output_dir = os.path.dirname(pdf_path) or '.'
            excel_abs_path = os.path.abspath(excel_path)
            
            command = [
                '/snap/bin/libreoffice',
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', output_dir,
                excel_abs_path
            ]
            
            # Try standard libreoffice if snap version not available
            try:
                subprocess.run(command, check=True, capture_output=True)
            except FileNotFoundError:
                command[0] = 'libreoffice'
                subprocess.run(command, check=True, capture_output=True)
            
            print(f"PDF file generated: {pdf_path}")
            return pdf_path
        
        except Exception as e:
            print(f"Error generating PDF: {e}")
            print("Make sure LibreOffice is installed: sudo apt install libreoffice")
            raise
    
    def generate_multiple(self, output_dir, changes_list):
        """
        Generate multiple Excel files from the template.
        
        Args:
            output_dir (str): Directory where files will be saved
            changes_list (list): List of tuples (filename, company_name, sakadastro, address, invoice_number, items, additional_changes_dict)
                                Example: [('file1.xlsx', 'Company A', 'Sak001', 'Address 1', 'INV001', 
                                          [{'type': 'Item1', 'quantity': 2, 'price': 100}], {}), 
                                         ('file2.xlsx', 'Company B', 'Sak002', 'Address 2', 'INV002', [], {})]
        
        Returns:
            list: Paths to all generated files
        """
        generated_files = []
        
        for item in changes_list:
            if len(item) == 5:
                filename, company_name, sakadastro, address, invoice_number = item
                items = []
                additional_changes = {}
            elif len(item) == 6:
                filename, company_name, sakadastro, address, invoice_number, items_or_changes = item
                if isinstance(items_or_changes, list):
                    items = items_or_changes
                    additional_changes = {}
                else:
                    items = []
                    additional_changes = items_or_changes
            else:
                filename, company_name, sakadastro, address, invoice_number, items, additional_changes = item
            
            output_path = os.path.join(output_dir, filename)
            self.generate(output_path, company_name, sakadastro, address, invoice_number, additional_changes, items)
            generated_files.append(output_path)
        
        return generated_files


# Example usage
if __name__ == "__main__":
    template_path = "template.xlsx"
    
    # Example 1: Generate single file with required fields only
    generator = ExcelTemplateGenerator(template_path)
    
    generator.generate(
        "output1.xlsx",
        company_name="ACME Corp",
        sakadastro="SAK-2024-001",
        address="123 Main Street, City",
        invoice_number="INV-001"
    )
    
    # Example 2: Generate with items (type, quantity, price)
    items = [
        {'type': 'Service A', 'quantity': 2, 'price': 100},
        {'type': 'Service B', 'quantity': 1, 'price': 50},
    ]
    
    generator.generate(
        "output2.xlsx",
        company_name="Tech Solutions",
        sakadastro="SAK-2024-002",
        address="456 Tech Avenue",
        invoice_number="INV-002",
        items=items
    )
    
    # Example 3: Generate multiple files with items
    changes_list = [
        ('output3.xlsx', 'Company A', 'SAK-001', 'Address A', 'INV-003',
         [{'type': 'Item1', 'quantity': 5, 'price': 75}]),
        ('output4.xlsx', 'Company B', 'SAK-002', 'Address B', 'INV-004',
         [{'type': 'Item2', 'quantity': 3, 'price': 50},
          {'type': 'Item3', 'quantity': 2, 'price': 100}]),
    ]
    
    generator.generate_multiple(".", changes_list)
