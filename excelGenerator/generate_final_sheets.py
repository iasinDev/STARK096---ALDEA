#!/usr/bin/env python3
"""
STARK096 — ALDEA: Generate Final Sheets
Creates individual sheets per housing unit from the base template with mejoras catalog
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
import os
import glob
import json
from utils import (
    create_output_dir,
    get_timestamp,
    print_success,
    print_error,
    print_info,
    print_warning
)


def load_mejoras_catalog():
    """Load the mejoras catalog from JSON file"""
    try:
        with open("mejoras_catalog.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print_error(f"Error loading catalog: {e}")
        return None


def get_latest_template():
    """Get the most recent template file from output directory"""
    output_dir = "output"
    pattern = os.path.join(output_dir, "template_viviendas_*.xlsx")
    files = glob.glob(pattern)
    
    if not files:
        print_error("No template files found in output/ directory")
        return None
    
    latest_file = max(files, key=os.path.getmtime)
    print_info(f"Using template: {os.path.basename(latest_file)}")
    return latest_file


def read_housing_data(filepath):
    """Read housing data from the template Excel file"""
    try:
        wb = openpyxl.load_workbook(filepath)
        ws = wb.active
        
        headers = []
        for col in range(1, ws.max_column + 1):
            header = ws.cell(row=1, column=col).value
            headers.append(header)
        
        viviendas = []
        for row in range(2, ws.max_row + 1):
            vivienda = {}
            for col in range(1, ws.max_column + 1):
                header = headers[col - 1]
                value = ws.cell(row=row, column=col).value
                vivienda[header] = value
            viviendas.append(vivienda)
        
        wb.close()
        print_info(f"Read {len(viviendas)} housing units")
        return viviendas
    
    except Exception as e:
        print_error(f"Error reading template: {e}")
        return None


def create_summary_sheet(wb, viviendas):
    """Create the summary sheet with all viviendas - estilo a.png"""
    ws = wb.active
    ws.title = "Resumen"
    
    # Define styles
    header_font = Font(bold=True, size=11, name="Calibri", color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    border_style = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center_align = Alignment(horizontal="center", vertical="center")
    
    # Title row
    ws.merge_cells('A1:J1')
    title_cell = ws['A1']
    title_cell.value = "Importes Viviendas Totales"
    title_cell.font = Font(bold=True, size=14, name="Calibri")
    title_cell.alignment = center_align
    
    # Headers row 2
    headers = [
        "COMPRADOR 1 NOMBRE",
        "COMPRADOR 1 APELLIDO 1",
        "COMPRADOR 1 APELLIDO 2",
        "COMPRADOR 2 NOMBRE",
        "COMPRADOR 2 APELLIDO 1",
        "COMPRADOR 2 APELLIDO 2",
        "PISO",
        "TOTAL GASTADO",
        "MAX MÓDULO",
        "SOBRANTE"
    ]
    
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border_style
        cell.alignment = center_align
    
    # Data rows starting from row 3
    for row_idx, vivienda in enumerate(viviendas, start=3):
        piso_short = vivienda.get("Piso", "").replace("Escalera ", "E").replace(" - Planta ", " ").replace(" - Puerta ", "")
        
        # Comprador 1
        ws.cell(row=row_idx, column=1, value=vivienda.get("Comprador 1 - Nombre")).border = border_style
        ws.cell(row=row_idx, column=2, value=vivienda.get("Comprador 1 - Apellido 1")).border = border_style
        ws.cell(row=row_idx, column=3, value=vivienda.get("Comprador 1 - Apellido 2")).border = border_style
        
        # Comprador 2
        ws.cell(row=row_idx, column=4, value=vivienda.get("Comprador 2 - Nombre")).border = border_style
        ws.cell(row=row_idx, column=5, value=vivienda.get("Comprador 2 - Apellido 1")).border = border_style
        ws.cell(row=row_idx, column=6, value=vivienda.get("Comprador 2 - Apellido 2")).border = border_style
        
        # Piso
        ws.cell(row=row_idx, column=7, value=piso_short).border = border_style
        
        # Total Gastado - formula referencing individual sheet
        cell_total = ws.cell(row=row_idx, column=8)
        cell_total.value = f"='{piso_short}'!D6"
        cell_total.number_format = '#,##0.00'
        cell_total.border = border_style
        
        # Max Módulo - manual entry
        ws.cell(row=row_idx, column=9, value="").border = border_style
        ws.cell(row=row_idx, column=9).number_format = '#,##0.00'
        
        # Sobrante - formula =MAX MÓDULO - TOTAL GASTADO
        cell_sobrante = ws.cell(row=row_idx, column=10)
        cell_sobrante.value = f"=I{row_idx}-H{row_idx}"
        cell_sobrante.number_format = '#,##0.00'
        cell_sobrante.border = border_style
    
    # Column widths
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 20
    ws.column_dimensions['F'].width = 20
    ws.column_dimensions['G'].width = 12
    ws.column_dimensions['H'].width = 15
    ws.column_dimensions['I'].width = 15
    ws.column_dimensions['J'].width = 15


def create_constructor_summary_sheet(wb):
    """Create constructor summary sheet with aggregated items"""
    ws = wb.create_sheet(title="Resumen Constructora", index=1)
    
    # Define styles
    header_font = Font(bold=True, size=11, name="Calibri", color="FFFFFF")
    header_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
    border_style = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    center_align = Alignment(horizontal="center", vertical="center")
    
    # Title
    ws.merge_cells('A1:C1')
    cell_title = ws.cell(row=1, column=1, value="LISTADO DE OBRA - RESUMEN CONSTRUCTORA")
    cell_title.font = Font(bold=True, size=14, name="Calibri")
    cell_title.alignment = center_align
    
    # Headers
    headers = ["CÓDIGO", "CONCEPTO", "CANTIDAD TOTAL"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border_style
        cell.alignment = center_align
    
    # Note: This will be filled manually or via aggregation logic later
    ws.cell(row=4, column=1, value="[Pendiente]")
    ws.cell(row=4, column=2, value="Se rellenará manualmente o mediante agregación")
    ws.cell(row=4, column=3, value="")
    
    # Column widths
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 18


def create_header_box(ws, vivienda):
    """Create the header box - FICHA DE CLIENTE"""
    
    title_fill = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
    title_font = Font(bold=True, size=12, name="Calibri", color="FFFFFF")
    
    label_font = Font(bold=True, size=11, name="Calibri", color="000000")
    value_font = Font(size=11, name="Calibri", color="000000")
    
    total_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    total_font = Font(bold=True, size=11, name="Calibri", color="000000")
    
    border_style = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    
    current_row = 1
    
    # FICHA DE CLIENTE
    ws.merge_cells(f'A{current_row}:D{current_row}')
    cell = ws.cell(row=current_row, column=1, value="FICHA DE CLIENTE")
    cell.font = title_font
    cell.fill = title_fill
    cell.border = border_style
    cell.alignment = center_align
    for col in range(2, 5):
        ws.cell(row=current_row, column=col).border = border_style
    current_row += 1
    
    # Nombres y Apellidos
    comprador1_nombre = vivienda.get("Comprador 1 - Nombre")
    comprador1_ap1 = vivienda.get("Comprador 1 - Apellido 1")
    comprador1_ap2 = vivienda.get("Comprador 1 - Apellido 2")
    
    compradores_text = ""
    if comprador1_nombre and comprador1_ap1 and comprador1_ap2:
        compradores_text = f"{comprador1_nombre} {comprador1_ap1} {comprador1_ap2}"
    
    comprador2_nombre = vivienda.get("Comprador 2 - Nombre")
    comprador2_ap1 = vivienda.get("Comprador 2 - Apellido 1")
    comprador2_ap2 = vivienda.get("Comprador 2 - Apellido 2")
    
    has_second_buyer = comprador2_nombre and comprador2_ap1 and comprador2_ap2
    
    ws.merge_cells(f'A{current_row}:B{current_row}')
    cell_label = ws.cell(row=current_row, column=1, value="Nombre y Apellidos:")
    cell_label.font = label_font
    cell_label.border = border_style
    cell_label.alignment = right_align
    ws.cell(row=current_row, column=2).border = border_style
    
    ws.merge_cells(f'C{current_row}:D{current_row}')
    cell_value = ws.cell(row=current_row, column=3, value=compradores_text)
    cell_value.font = value_font
    cell_value.border = border_style
    cell_value.alignment = left_align
    ws.cell(row=current_row, column=4).border = border_style
    current_row += 1
    
    # Second buyer if exists
    if has_second_buyer:
        comprador2_text = f"{comprador2_nombre} {comprador2_ap1} {comprador2_ap2}"
        
        ws.merge_cells(f'A{current_row}:B{current_row}')
        cell_label = ws.cell(row=current_row, column=1, value="")
        cell_label.border = border_style
        ws.cell(row=current_row, column=2).border = border_style
        
        ws.merge_cells(f'C{current_row}:D{current_row}')
        cell_value = ws.cell(row=current_row, column=3, value=comprador2_text)
        cell_value.font = value_font
        cell_value.border = border_style
        cell_value.alignment = left_align
        ws.cell(row=current_row, column=4).border = border_style
        current_row += 1
        
        # Separator
        ws.merge_cells(f'A{current_row}:D{current_row}')
        cell = ws.cell(row=current_row, column=1, value="")
        cell.border = border_style
        for col in range(2, 5):
            ws.cell(row=current_row, column=col).border = border_style
        current_row += 1
    else:
        # Empty separator
        ws.merge_cells(f'A{current_row}:D{current_row}')
        cell = ws.cell(row=current_row, column=1, value="")
        cell.border = border_style
        for col in range(2, 5):
            ws.cell(row=current_row, column=col).border = border_style
        current_row += 1
    
    # Vivienda
    piso = vivienda.get("Piso", "")
    vivienda_short = piso.replace("Escalera ", "E").replace(" - Planta ", " ").replace(" - Puerta ", "")
    
    ws.merge_cells(f'A{current_row}:B{current_row}')
    cell_label = ws.cell(row=current_row, column=1, value="Vivienda:")
    cell_label.font = label_font
    cell_label.border = border_style
    cell_label.alignment = right_align
    ws.cell(row=current_row, column=2).border = border_style
    
    ws.merge_cells(f'C{current_row}:D{current_row}')
    cell_value = ws.cell(row=current_row, column=3, value=vivienda_short)
    cell_value.font = value_font
    cell_value.border = border_style
    cell_value.alignment = left_align
    ws.cell(row=current_row, column=4).border = border_style
    current_row += 1
    
    # Empty row
    ws.merge_cells(f'A{current_row}:D{current_row}')
    cell = ws.cell(row=current_row, column=1, value="")
    cell.border = border_style
    for col in range(2, 5):
        ws.cell(row=current_row, column=col).border = border_style
    current_row += 1
    
    # TOTAL GASTADO (row 6)
    ws.merge_cells(f'A{current_row}:B{current_row}')
    cell_label = ws.cell(row=current_row, column=1, value="TOTAL GASTADO:")
    cell_label.font = total_font
    cell_label.fill = total_fill
    cell_label.border = border_style
    cell_label.alignment = right_align
    ws.cell(row=current_row, column=2).border = border_style
    ws.cell(row=current_row, column=2).fill = total_fill
    
    cell_value = ws.cell(row=current_row, column=3, value="")  # Formula added later
    cell_value.fill = total_fill
    cell_value.border = border_style
    cell_value.alignment = right_align
    cell_value.number_format = '#,##0.00'
    
    cell_currency = ws.cell(row=current_row, column=4, value="€ + IVA")
    cell_currency.font = value_font
    cell_currency.fill = total_fill
    cell_currency.border = border_style
    cell_currency.alignment = left_align
    
    # Column widths
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 12
    
    return current_row


def get_cocina_options_for_vivienda(piso_short, catalog):
    """Get cocina options for specific vivienda"""
    cocina_paso = None
    for paso in catalog["pasos_basicos"]:
        if paso["codigo"] == "COCINA":
            cocina_paso = paso
            break
    
    if not cocina_paso:
        return []
    
    for grupo in cocina_paso["grupos_por_tipo"]:
        if piso_short in grupo["tipos"]:
            return grupo["opciones"]
    
    return []


def create_mejoras_table(ws, vivienda, catalog, start_row):
    """Create mejoras selection table - estilo b.png"""
    
    header_fill = PatternFill(start_color="C6AC45", end_color="C6AC45", fill_type="solid")
    header_font = Font(bold=True, size=10, name="Calibri", color="000000")
    
    total_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    total_font = Font(bold=True, size=11, name="Calibri")
    
    border_style = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    
    current_row = start_row + 2
    
    # Headers
    headers = ["Concepto", "Selección", "Precio (€)", "Cantidad", "Seleccionado (Sí/No)", "Importe (€)"]
    cols = ['A', 'B', 'C', 'D', 'E', 'F']
    
    for col, header in zip(cols, headers):
        cell = ws[f'{col}{current_row}']
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border_style
        cell.alignment = center_align
    
    # Column widths
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 12
    
    current_row += 1
    data_start_row = current_row
    
    piso_short = vivienda.get("Piso", "").replace("Escalera ", "E").replace(" - Planta ", " ").replace(" - Puerta ", "")
    
    # Add basic pasos
    for paso in catalog["pasos_basicos"]:
        if paso["codigo"] == "COCINA":
            opciones = get_cocina_options_for_vivienda(piso_short, catalog)
            if opciones:
                ws[f'A{current_row}'].value = paso["nombre"]
                ws[f'A{current_row}'].border = border_style
                ws[f'A{current_row}'].alignment = left_align
                
                ws[f'B{current_row}'].border = border_style
                opciones_str = ",".join([opt["nombre"] for opt in opciones])
                dv = DataValidation(type="list", formula1=f'"{opciones_str}"', allow_blank=True)
                ws.add_data_validation(dv)
                dv.add(f'B{current_row}')
                
                ws[f'C{current_row}'].value = "N/A"
                ws[f'C{current_row}'].border = border_style
                ws[f'C{current_row}'].alignment = right_align
                ws[f'C{current_row}'].number_format = '#,##0.00'
                
                ws[f'D{current_row}'].value = 1
                ws[f'D{current_row}'].border = border_style
                ws[f'D{current_row}'].alignment = center_align
                
                ws[f'E{current_row}'].border = border_style
                dv_sino = DataValidation(type="list", formula1='"Sí,No"', allow_blank=True)
                ws.add_data_validation(dv_sino)
                dv_sino.add(f'E{current_row}')
                
                ws[f'F{current_row}'].value = f'=IF(E{current_row}="Sí",C{current_row}*D{current_row},0)'
                ws[f'F{current_row}'].border = border_style
                ws[f'F{current_row}'].alignment = right_align
                ws[f'F{current_row}'].number_format = '#,##0.00'
                
                current_row += 1
        else:
            ws[f'A{current_row}'].value = paso["nombre"]
            ws[f'A{current_row}'].border = border_style
            ws[f'A{current_row}'].alignment = left_align
            
            ws[f'B{current_row}'].border = border_style
            opciones_str = ",".join([opt["nombre"] for opt in paso["opciones"]])
            dv = DataValidation(type="list", formula1=f'"{opciones_str}"', allow_blank=True)
            ws.add_data_validation(dv)
            dv.add(f'B{current_row}')
            
            ws[f'C{current_row}'].value = 0
            ws[f'C{current_row}'].border = border_style
            ws[f'C{current_row}'].alignment = right_align
            ws[f'C{current_row}'].number_format = '#,##0.00'
            
            ws[f'D{current_row}'].value = 1
            ws[f'D{current_row}'].border = border_style
            ws[f'D{current_row}'].alignment = center_align
            
            ws[f'E{current_row}'].value = "Sí"
            ws[f'E{current_row}'].border = border_style
            dv_sino = DataValidation(type="list", formula1='"Sí,No"', allow_blank=True)
            ws.add_data_validation(dv_sino)
            dv_sino.add(f'E{current_row}')
            
            ws[f'F{current_row}'].value = 0
            ws[f'F{current_row}'].border = border_style
            ws[f'F{current_row}'].alignment = right_align
            ws[f'F{current_row}'].number_format = '#,##0.00'
            
            current_row += 1
    
    # Add key mejoras
    mejoras_sample = [
        ("Armario 2 hojas abatibles", 1093.75),
        ("Armario 3 hojas abatibles", 1475.00),
        ("Armario 4 hojas abatibles", 1620.44),
        ("Mueble+espejo baño principal 80cm", 986.00),
        ("Mampara ducha Walk-in 100", 995.25),
        ("Paquete light cocina+baños", 687.50),
        ("Cambio laminado → gres porcelánico", 2313.94),
    ]
    
    for concepto, precio in mejoras_sample:
        ws[f'A{current_row}'].value = concepto
        ws[f'A{current_row}'].border = border_style
        ws[f'A{current_row}'].alignment = left_align
        
        ws[f'B{current_row}'].value = ""
        ws[f'B{current_row}'].border = border_style
        
        ws[f'C{current_row}'].value = precio
        ws[f'C{current_row}'].border = border_style
        ws[f'C{current_row}'].alignment = right_align
        ws[f'C{current_row}'].number_format = '#,##0.00'
        
        ws[f'D{current_row}'].value = 0
        ws[f'D{current_row}'].border = border_style
        ws[f'D{current_row}'].alignment = center_align
        
        ws[f'E{current_row}'].value = "No"
        ws[f'E{current_row}'].border = border_style
        dv_sino = DataValidation(type="list", formula1='"Sí,No"', allow_blank=True)
        ws.add_data_validation(dv_sino)
        dv_sino.add(f'E{current_row}')
        
        ws[f'F{current_row}'].value = f'=IF(E{current_row}="Sí",C{current_row}*D{current_row},0)'
        ws[f'F{current_row}'].border = border_style
        ws[f'F{current_row}'].alignment = right_align
        ws[f'F{current_row}'].number_format = '#,##0.00'
        
        current_row += 1
    
    data_end_row = current_row - 1
    
    # TOTAL row
    current_row += 1
    
    ws.merge_cells(f'A{current_row}:E{current_row}')
    cell_total_label = ws[f'A{current_row}']
    cell_total_label.value = "TOTAL"
    cell_total_label.font = total_font
    cell_total_label.fill = total_fill
    cell_total_label.border = border_style
    cell_total_label.alignment = right_align
    
    for col in ['B', 'C', 'D', 'E']:
        ws[f'{col}{current_row}'].border = border_style
        ws[f'{col}{current_row}'].fill = total_fill
    
    cell_total_value = ws[f'F{current_row}']
    cell_total_value.value = f'=SUM(F{data_start_row}:F{data_end_row})'
    cell_total_value.font = total_font
    cell_total_value.fill = total_fill
    cell_total_value.border = border_style
    cell_total_value.alignment = right_align
    cell_total_value.number_format = '#,##0.00'
    
    # Update TOTAL GASTADO in header (column C of the header row)
    ws[f'C{start_row}'].value = f'=F{current_row}'
    
    return current_row


def generate_final_sheets():
    """Generate the final Excel with individual sheets"""
    
    print_info("Loading mejoras catalog...")
    catalog = load_mejoras_catalog()
    if not catalog:
        return 1
    
    print_info("Generating final Excel...")
    
    template_file = get_latest_template()
    if not template_file:
        return 1
    
    viviendas = read_housing_data(template_file)
    if not viviendas:
        return 1
    
    output_dir = "output_final"
    create_output_dir(output_dir)
    
    wb = openpyxl.Workbook()
    
    print_info("Creating summary sheet...")
    create_summary_sheet(wb, viviendas)
    
    print_info("Creating constructor summary sheet...")
    create_constructor_summary_sheet(wb)
    
    print_info(f"Creating {len(viviendas)} individual sheets...")
    
    for idx, vivienda in enumerate(viviendas, start=1):
        piso = vivienda.get("Piso", f"Vivienda_{idx}")
        piso_short = piso.replace("Escalera ", "E").replace(" - Planta ", " ").replace(" - Puerta ", "")
        
        sheet_name = piso_short[:31]
        
        print_info(f"  Sheet {idx}/{len(viviendas)}: {sheet_name}")
        
        ws = wb.create_sheet(title=sheet_name)
        
        header_end_row = create_header_box(ws, vivienda)
        create_mejoras_table(ws, vivienda, catalog, header_end_row)
    
    timestamp = get_timestamp()
    output_file = os.path.join(output_dir, f"viviendas_final_{timestamp}.xlsx")
    
    try:
        wb.save(output_file)
        print_success(f"Excel saved: {output_file}")
        print_info(f"Total sheets: {len(wb.sheetnames)}")
        return 0
    except Exception as e:
        print_error(f"Error saving Excel: {e}")
        return 1


if __name__ == "__main__":
    exit(generate_final_sheets())
