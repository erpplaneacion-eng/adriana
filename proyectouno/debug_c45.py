import openpyxl
import xlrd
import os
import re

file_path = r"c:\Users\User\OneDrive\Desktop\CHVS\adriana\librouno_ejemplo.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=False)
sheet = wb['GASTOS ADMINISTRATIVOS ']

# Mapeo
link_map = {}
if hasattr(wb, '_external_links'):
    for i, link in enumerate(wb._external_links):
        target = os.path.basename(link.file_link.Target).replace('%20', ' ')
        parts = target.split('_')
        if len(parts) > 2:
            pattern = parts[0] + "_" + parts[1]
        else:
            pattern = target.replace(".xls", "").replace(".xlsx", "")
        link_map[str(i+1)] = pattern

formula = str(sheet['C45'].value)
print("Fórmula C45:", formula)

folder_path = r"c:\Users\User\OneDrive\Desktop\CHVS\adriana\FEBRERO"

def get_xls_value(folder_path, filename_pattern, sheet_name, cell_ref):
    try:
        matching_files = [f for f in os.listdir(folder_path) if f.lower().startswith(filename_pattern.lower())]
        if not matching_files:
            return f"NO FILE: {filename_pattern}"
        
        filename = matching_files[0]
        file_path = os.path.join(folder_path, filename)
        wb = xlrd.open_workbook(file_path)
        try:
            sheet = wb.sheet_by_name(sheet_name)
        except:
            sheet = wb.sheet_by_index(0)
            
        clean_ref = cell_ref.replace('$', '')
        col_letter = re.match(r"([A-Z]+)", clean_ref).group(1)
        row_num = int(re.search(r"(\d+)", clean_ref).group(1)) - 1
        
        col_idx = 0
        for char in col_letter:
            col_idx = col_idx * 26 + (ord(char) - ord('A') + 1)
        col_idx -= 1
        
        val = sheet.cell_value(row_num, col_idx)
        if val == "" or val is None:
            return 0
        return float(val)
    except Exception as e:
        return f"ERROR: {e}"

def evaluate_external_link(match):
    link_id = match.group(1)
    sheet_n = match.group(2)
    cell_ref = match.group(3)
    
    file_pattern = link_map.get(link_id)
    val = get_xls_value(folder_path, file_pattern, sheet_n, cell_ref)
    print(f"Link [{link_id}] -> Patrón: {file_pattern} | Archivo real en carpeta: {[f for f in os.listdir(folder_path) if f.lower().startswith(file_pattern.lower())]} | Celda {cell_ref} -> VALOR ENCONTRADO: {val}")
    return str(val)

final_formula = re.sub(r"'?\[(\d+)\]([^']+)'?!([A-Z0-9\$]+)", evaluate_external_link, formula)
print("Fórmula inyectada:", final_formula)

