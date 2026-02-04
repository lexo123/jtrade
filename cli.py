"""
Simple CLI interface for Excel Template Generator
"""

import os
from excel_generator import ExcelTemplateGenerator


def main():
    """Interactive command-line interface for generating Excel files."""

    template_path = "template.xlsx"
    if not os.path.exists(template_path):
        print(f"Error: Template file '{template_path}' not found!")
        return

    generator = ExcelTemplateGenerator(template_path)
    
    print("=" * 50)
    print("Excel Template Generator")
    print("=" * 50)
    
    # Get filename without extension
    filename = input("\nEnter output filename (without extension, e.g., 'lexo'): ").strip()
    # Remove extension if user included it
    if filename.endswith(('.xlsx', '.xls', '.pdf')):
        filename = filename.rsplit('.', 1)[0]

    # Sanitize filename to handle Unicode characters properly
    from app import safe_filename
    filename = safe_filename(filename)
    output_file = f"{filename}.xlsx"
    
    # Required fields
    company_name = input("Enter company name (A12 - will be prefixed with 'კომპ/სახელი'): ").strip()
    sakadastro = input("Enter sakadastro (A13 - will be prefixed with 'ს/კ'): ").strip()
    address = input("Enter address (A14 - will be prefixed with 'მისამართი'): ").strip()
    invoice_number = input("Enter invoice number (D5): ").strip()
    
    # Optional items (rows 17-24)
    items = []
    for row in range(17, 25):  # A17 to A24
        add_item = input(f"\nAdd item in row {row}? (y/n): ").strip().lower() == 'y'
        if not add_item:
            break
        
        item_type = input(f"  Type (A{row}): ").strip()
        if not item_type:
            break
        
        try:
            quantity = input(f"  Quantity (B{row}): ").strip()
            quantity = float(quantity) if quantity else ''
        except ValueError:
            quantity = ''
        
        try:
            price = input(f"  Price (C{row}): ").strip()
            price = float(price) if price else ''
        except ValueError:
            price = ''
        
        items.append({
            'type': item_type,
            'quantity': quantity,
            'price': price
        })
        print(f"  Added: {item_type}")
    
    print("\nEnter additional cell changes (optional, format: CELL=VALUE)")
    print("Press Enter twice to finish:")
    
    changes = {}
    while True:
        user_input = input().strip()
        if not user_input:
            break
        
        try:
            cell, value = user_input.split('=', 1)
            cell = cell.strip().upper()
            value_str = value.strip()
            
            # Try to convert to number if possible
            try:
                if '.' in value_str:
                    value = float(value_str)
                else:
                    value = int(value_str)
            except ValueError:
                value = value_str
            
            changes[cell] = value
            print(f"  Added: {cell} = {value}")
        except ValueError:
            print("  Invalid format. Use CELL=VALUE (e.g., A1=Hello)")
    
    generator.generate(output_file, company_name, sakadastro, address, invoice_number, changes, items)
    
    # Ask if user wants to generate PDF
    generate_pdf = input("\nGenerate PDF from the Excel file? (y/n): ").strip().lower() == 'y'
    if generate_pdf:
        # Automatically use same filename with .pdf extension
        pdf_file = f"{filename}.pdf"
        print(f"Generating PDF: {pdf_file}")
        generator.generate_pdf(output_file, pdf_file)


if __name__ == "__main__":
    main()
