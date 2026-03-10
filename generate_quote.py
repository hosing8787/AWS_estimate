import os
import re
import json
import csv
from copy import copy as python_copy
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

from shared_logic import WON_FMT, USD_FMT, copy_style, sc, reconcile_branding_and_xml

def parse_export(file_path):
    inventory = []
    with open(file_path, 'r', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        in_details_section = False
        headers = []
        service_cols = ['서비스', 'Service']
        desc_cols = ['설명', 'Description']
        config_cols = ['구성 요약', 'Configuration summary', 'Configuration Summary']
        monthly_cols = ['월별', '월간', 'Monthly']
        for row in reader:
            if not row or not any(row): continue
            if any(s in row for s in service_cols) and any(d in row for d in desc_cols):
                headers = [h.strip() for h in row]
                in_details_section = True
                continue
            if in_details_section:
                # End of items detection
                if not row[0]: continue  # skip empty
                fst = row[0].strip().lower()
                if fst.startswith('확인') or fst.startswith('* aws') or fst.startswith('acknowledgement') or fst.startswith('acknowledgment'):
                    break
                def get_idx(candidates):
                    for c in candidates:
                        if c in headers: return headers.index(c)
                    return -1
                idx_service = get_idx(service_cols)
                idx_desc = get_idx(desc_cols)
                idx_monthly = get_idx(monthly_cols)
                idx_config = get_idx(config_cols)
                if idx_service == -1 or idx_desc == -1: continue
                if len(row) <= max(idx_service, idx_desc): continue
                if not row[idx_service].strip() and not row[idx_desc].strip(): break
                try:
                    service_name = row[idx_service].strip()
                    desc_text = row[idx_desc].strip()
                    config_summary = row[idx_config].strip() if idx_config != -1 and idx_config < len(row) else ""
                    mrr = 0.0
                    if idx_monthly != -1 and idx_monthly < len(row):
                        mrr_str = row[idx_monthly].strip().replace(',', '')
                        try: mrr = float(mrr_str) if mrr_str else 0.0
                        except ValueError: mrr = 0.0
                    cat_name = "개발계" if any(kw in desc_text.lower() for kw in ['dev', 'test', 'stg', 'stage']) else "운영계"
                    inventory.append({"category": cat_name, "description": desc_text, "service": service_name, "config_summary": config_summary, "mrr_usd": mrr})
                except Exception: pass
    return inventory

# --- Total Reconstruction Strategy ---
def run_generation(output_filename='AWS_Quote_Output.xlsx', orig_template='template.xlsx'):
    # Load Data from sample_data.json (Matching Azure Pattern)
    try:
        with open("sample_data.json", "r", encoding="utf-8") as f:
            data = json.load(f)
            inventory = data.get("inventory_items", [])
    except Exception as e:
        print(f"Failed to load sample_data.json: {e}")
        inventory = []

    # 1. Load Template and Prepare New Workbook
    wb_orig = load_workbook(orig_template)
    ws_orig = wb_orig['견적서']
    
    wb_gen = Workbook()
    ws_gen = wb_gen.active
    ws_gen.title = '견적서'
    
    # 2. Replicate Column Widths
    for c in range(1, 15):
        let = get_column_letter(c)
        ws_gen.column_dimensions[let].width = ws_orig.column_dimensions[let].width if ws_orig.column_dimensions[let].width else 8.43

    # 3. Phase 1: Port Top Section (Rows 1-12)
    for r in range(1, 13):
        ws_gen.row_dimensions[r].height = ws_orig.row_dimensions[r].height
        for c in range(1, 14):
            src_c = ws_orig.cell(row=r, column=c)
            tgt_c = ws_gen.cell(row=r, column=c)
            tgt_c.value = src_c.value
            copy_style(src_c, tgt_c)
    
    # 4. Phase 2: Dynamic Items (Starting from Row 13)
    num_items = len(inventory)
    groups = {}
    for item in inventory:
        cat = item['category']
        if cat not in groups: groups[cat] = []
        groups[cat].append(item)
    
    current_row = 13
    s2_row_start = 9 # Row in Sheet 2 (AWS Resource 내역)
    item_rows = [] 
    
    for cat_name, items in groups.items():
        cat_start_row = current_row
        for item in items:
            r = current_row
            ws_gen.row_dimensions[r].height = ws_orig.row_dimensions[13].height # Use Item Row Height
            for c in range(1, 14):
                src_c = ws_orig.cell(row=13, column=c) # Style from Item Row 13
                tgt_c = ws_gen.cell(row=r, column=c)
                copy_style(src_c, tgt_c)
            
            # Fill Item Data (Sheet 1)
            sc(ws_gen, f'D{r}', f"='AWS Resource 내역'!D{s2_row_start}", al=Alignment(horizontal='left', vertical='center'))
            sc(ws_gen, f'F{r}', 1, al=Alignment(horizontal='center', vertical='center'))
            sc(ws_gen, f'G{r}', 2.0, al=Alignment(horizontal='center', vertical='center')) # Default duration
            sc(ws_gen, f'H{r}', '식', al=Alignment(horizontal='center', vertical='center'))
            sc(ws_gen, f'I{r}', f"='AWS Resource 내역'!G{s2_row_start}", nf=WON_FMT)
            sc(ws_gen, f'J{r}', f"=I{r}*G{r}*F{r}", nf=WON_FMT)
            sc(ws_gen, f'K{r}', '상동', al=Alignment(horizontal='center', vertical='center'))
            
            # Replicate template merges for item description (D:E)
            ws_gen.merge_cells(f'D{r}:E{r}')
            # Propagate style for merged D:E
            copy_style(ws_gen[f'D{r}'], ws_gen[f'E{r}'])
            
            item_rows.append(r)
            current_row += 1
            s2_row_start += 1
            
        # Merge Category Column C
        ws_gen.merge_cells(f'C{cat_start_row}:C{current_row-1}')
        cat_cell_val = f'{cat_name}\nAWS Resource'
        sc(ws_gen, f'C{cat_start_row}', cat_cell_val, 
           fnt=Font(name='맑은 고딕', bold=True), 
           al=Alignment(horizontal='center', vertical='center', wrap_text=True))
        # Ensure all cells in merge have template borders
        for rm in range(cat_start_row, current_row):
            ws_gen.cell(row=rm, column=3).border = python_copy(ws_orig.cell(row=13, column=3).border)
        
        # In Sheet 2, there is a subtotal row after each group
        s2_row_start += 1

    # 5. Phase 3: Bottom Section (Totals, Remarks)
    if not item_rows: item_rows = [13]
    
    # The bottom section always begins with '운영계 소계' & 'Cloud 사용료 소계' at Row 20
    bottom_start_src = 20
    row_offset = current_row - bottom_start_src
    
    for r_src in range(bottom_start_src, bottom_start_src + 20):
        if not ws_orig.row_dimensions[r_src].height and not ws_orig.cell(row=r_src, column=2).value and not ws_orig.cell(row=r_src, column=3).value:
            break # End of template
            
        r_tgt = r_src + row_offset
        ws_gen.row_dimensions[r_tgt].height = ws_orig.row_dimensions[r_src].height
        ws_gen.row_dimensions[r_tgt].hidden = ws_orig.row_dimensions[r_src].hidden
        for c in range(1, 14):
            src_c = ws_orig.cell(row=r_src, column=c)
            tgt_c = ws_gen.cell(row=r_tgt, column=c)
            tgt_c.value = src_c.value
            copy_style(src_c, tgt_c)
            
            # Dynamic Formula Shifting
            if isinstance(tgt_c.value, str) and tgt_c.value.startswith('='):
                val = tgt_c.value
                val = val.replace('J13:J19', f'J13:J{item_rows[-1]}')
                val = val.replace('J13:J29', f'J13:J{item_rows[-1]}')
                
                def shift_m(m):
                    tr = int(m.group(2))
                    if tr >= bottom_start_src: return f"{m.group(1)}{tr + row_offset}"
                    return m.group(0)
                val = re.sub(r'([A-Z])(\d+)', shift_m, val)
                tgt_c.value = val

    # 6. Global Merges Re-Application (The safe way)
    # 6a. Top merges
    for m in ws_orig.merged_cells.ranges:
        if m.max_row <= 12: ws_gen.merge_cells(str(m))
    
    # 6b. Bottom merges (Shifting all, including Col B)
    # "AWS Resource 소계" is exactly at row 30 in the template
    aws_total_row_tgt = 30 + row_offset
    for m in ws_orig.merged_cells.ranges:
        if m.min_row >= bottom_start_src:
            m_s = python_copy(m)
            m_s.shift(row_shift=row_offset)
            # Only apply if it doesn't conflict with our manual B13:B merge
            if m_s.min_col >= 3 or m_s.min_row >= aws_total_row_tgt:
                try: ws_gen.merge_cells(str(m_s))
                except: pass
                
    # 6c. Big B merge (AWS Resource & Category Branding)
    # Ends exactly one row before "AWS Resource 소계"
    last_aws_row = aws_total_row_tgt - 1
    ws_gen.merge_cells(f'B13:B{last_aws_row}')
    sc(ws_gen, f'B13', 'AWS\nResource', fnt=Font(name='맑은 고딕', bold=True), al=Alignment(horizontal='center', vertical='center', wrap_text=True))
    for rb in range(13, last_aws_row + 1):
        ws_gen.cell(row=rb, column=2).border = python_copy(ws_orig.cell(row=13, column=2).border)

    # 7. Create Sheet 2 (AWS Resource 내역)
    if 'AWS Resource 내역' in wb_gen.sheetnames: del wb_gen['AWS Resource 내역']
    ws2 = wb_gen.create_sheet('AWS Resource 내역')
    build_aws_resource_details(ws2, inventory)

    # Clear Title / Customer Name per user request
    ws_gen['B4'].value = '고 객 사 명 : '
    ws_gen['C5'].value = '프로젝트명 : '

    # 8. Save Workbook
    # 8a. Temporary Save to local disk
    wb_gen.save(output_filename)
    
    # 9. Surgical Injection (Replication of Theme, Logo, VML)
    reconcile_branding_and_xml(orig_template, output_filename)

def build_aws_resource_details(ws, inventory):
    # Set widths to match template precisely
    cols = {'B': 19.125, 'C': 32.625, 'D': 28.25, 'E': 94.875, 'F': 13.375, 'G': 13.0}
    for k, v in cols.items(): ws.column_dimensions[k].width = v
    
    sc(ws,'B2','리전', fill=PatternFill('solid', fgColor='D9D9D9'))
    sc(ws,'C2','SEOUL')
    sc(ws,'B3','환율', fill=PatternFill('solid', fgColor='D9D9D9'))
    sc(ws,'C3', 1405.5)
    
    headers = ['구분','Description','Service','Configuration Summary','MRR($)','MRR(\\)']
    for i, h in enumerate(headers):
        sc(ws, f'{get_column_letter(i+2)}8', h, fill=PatternFill('solid', fgColor='ED7D31'), fnt=Font(color='FFFFFF'))

    current_row = 9
    groups = {}
    for item in inventory:
        cat = item['category']
        if cat not in groups: groups[cat] = []
        groups[cat].append(item)
    
    subs = []
    for cat, items in groups.items():
        start = current_row
        for item in items:
            ws.row_dimensions[current_row].height = 60
            sc(ws, f'B{current_row}', cat, al=Alignment(horizontal='center', vertical='center', wrap_text=True))
            sc(ws, f'C{current_row}', item['description'], al=Alignment(horizontal='left', vertical='center', wrap_text=True))
            sc(ws, f'D{current_row}', item['service'], al=Alignment(horizontal='left', vertical='center', wrap_text=True))
            sc(ws, f'E{current_row}', item.get('config_summary', ''), al=Alignment(horizontal='left', vertical='center', wrap_text=True))
            sc(ws, f'F{current_row}', item['mrr_usd'], nf=USD_FMT, al=Alignment(horizontal='right', vertical='center'))
            sc(ws, f'G{current_row}', f'=F{current_row}*$C$3', nf=WON_FMT, al=Alignment(horizontal='right', vertical='center'))
            current_row += 1
        ws.merge_cells(f'B{start}:B{current_row-1}')
        
        ws.merge_cells(f'C{current_row}:E{current_row}')
        sc(ws, f'C{current_row}', f'{cat} 소계', fill=PatternFill('solid', fgColor='FCE4D6'), al=Alignment(horizontal='right', vertical='center'))
        sc(ws, f'F{current_row}', f'=SUM(F{start}:F{current_row-1})', fill=PatternFill('solid', fgColor='FCE4D6'), nf=USD_FMT)
        sc(ws, f'G{current_row}', f'=SUM(G{start}:G{current_row-1})', fill=PatternFill('solid', fgColor='FCE4D6'), nf=WON_FMT)
        subs.append(current_row)
        current_row += 1
        
    ws.merge_cells(f'B{current_row}:E{current_row}')
    sc(ws, f'B{current_row}', 'Cloud 사용료', fill=PatternFill('solid', fgColor='F8CBAD'), al=Alignment(horizontal='center', vertical='center'))
    sc(ws, f'G{current_row}', f'=SUM({",".join(["G"+str(r) for r in subs])})', fill=PatternFill('solid', fgColor='F8CBAD'), nf=WON_FMT)

if __name__ == '__main__':
    # Default test run
    import os
    if os.path.exists('SuperbAI_250915.csv'):
        run_generation('SuperbAI_250915.csv', 'SuperbAI_AWS_Quote_Final.xlsx')
