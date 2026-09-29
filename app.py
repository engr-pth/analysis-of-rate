import os
import io
import traceback
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Labour/Worker Keywords List
LABOUR_KEYWORDS = [
    "worker", "digger", "mason", "carpenter", "maistry", 
    "blacksmith", "steel worker", "welder", "surveyor", 
    "smith", "machine driver", "charges", "labour",
    "hoisting and fixing", "carriage to site", "site clearing", "dressing"
]

@st.cache_data
def parse_excel_rates(file_path):
    if not os.path.exists(file_path):
        return []
    try:
        df_raw = pd.read_excel(file_path, dtype=str)
    except Exception as e:
        st.error(f"ဖိုင်ဖတ်၍ မရပါ ({file_path}): {e}")
        return []

    items = []
    current_item = None

    for idx, row in df_raw.iterrows():
        item_no = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        particular = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
        unit = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ''
        qty = row.iloc[3] if pd.notna(row.iloc[3]) else '0'

        if 'nat' in item_no.lower() or '202' in item_no or item_no in ['', 'nan', 'No.', 'NaN']:
            if current_item and particular not in ['', 'nan', 'NaN', 'Particular']:
                try:
                    qty_val = float(qty)
                except (ValueError, TypeError):
                    qty_val = 0.0
                
                current_item['breakdown'].append({
                    'particular': particular,
                    'unit': unit,
                    'qty': qty_val
                })
            continue

        if particular not in ['', 'nan', 'NaN', 'Particular']:
            if current_item:
                items.append(current_item)
            current_item = {
                'item_no': item_no,
                'title': particular,
                'unit': unit,
                'std_qty': qty,
                'breakdown': []
            }

    if current_item:
        items.append(current_item)

    return items


def parse_and_auto_select_uploaded_excel(uploaded_file, all_item_maps):
    st.session_state['last_excel_error'] = None

    try:
        if uploaded_file is None:
            return

        if hasattr(uploaded_file, 'seek'):
            uploaded_file.seek(0)

        file_name = getattr(uploaded_file, 'name', '').lower()

        try:
            if file_name.endswith('.csv'):
                df_raw = pd.read_csv(uploaded_file, header=None)
            else:
                df_raw = pd.read_excel(uploaded_file, header=None)
        except Exception:
            if hasattr(uploaded_file, 'seek'):
                uploaded_file.seek(0)
            if file_name.endswith('.csv'):
                df_raw = pd.read_csv(uploaded_file, header=None)
            else:
                df_raw = pd.read_excel(uploaded_file, header=None, engine='openpyxl')

        header_row_idx = None
        for idx, row in df_raw.iterrows():
            row_vals = row.dropna().astype(str).str.lower().tolist()
            if any('particular' in v or 'description' in v for v in row_vals) and any('no' in v or 'item' in v for v in row_vals):
                header_row_idx = idx
                break
        
        if hasattr(uploaded_file, 'seek'):
            uploaded_file.seek(0)

        if header_row_idx is not None:
            if file_name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, header=header_row_idx)
            else:
                df = pd.read_excel(uploaded_file, header=header_row_idx)
        else:
            if file_name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)

        df.columns = [str(c).strip().lower() for c in df.columns]
        
        col_item = next((c for c in df.columns if 'item' in c), None)
        col_desc = next((c for c in df.columns if 'particular' in c or 'desc' in c), None)
        col_no = next((c for c in df.columns if 'no' in c and 'item' not in c), None)
        col_l = next((c for c in df.columns if 'l (' in c or 'l(' in c or 'length' in c or c == 'l'), None)
        col_b = next((c for c in df.columns if 'b (' in c or 'b(' in c or 'breadth' in c or 'width' in c or c == 'b'), None)
        col_h = next((c for c in df.columns if 'h (' in c or 'h(' in c or 'height' in c or 'depth' in c or c == 'h'), None)
        col_ded = next((c for c in df.columns if 'deduction' in c or ('ded' in c and 'type' not in c)), None)
        col_type = next((c for c in df.columns if 'type' in c), None)

        if not col_item or not col_desc:
            st.session_state['last_excel_error'] = "❌ တင်သွင်းသော Excel ဖိုင်တွင် 'Item No.' သို့မဟုတ် 'Particular Description' Column ကို ရှာမတွေ့ပါ။"
            return

        def clean_str_item(val):
            if pd.isna(val):
                return ""
            s = str(val).strip()
            if s.endswith('.0'):
                s = s[:-2]
            return s

        df['clean_item_no'] = df[col_item].apply(clean_str_item)
        
        valid_rows = df[
            (df['clean_item_no'] != '') & 
            ~df['clean_item_no'].str.lower().str.contains('item no|total|detail|description') &
            df[col_desc].notna() &
            (df[col_desc].astype(str).str.strip() != '')
        ].copy()

        excel_item_nos = valid_rows['clean_item_no'].unique().tolist()

        if not excel_item_nos:
            st.session_state['last_excel_error'] = "⚠️ Excel ဖိုင်ထဲတွင် Measurement Data များ ရှာမတွေ့ပါ။"
            return

        full_text_str = " ".join([str(x) for x in df_raw.values.flatten() if pd.notna(x)]).lower()

        selected_ew, selected_cc, selected_ir, selected_rc, selected_rc_m, selected_bw = [], [], [], [], [], []

        if 'ew' in all_item_maps:
            for k, v in all_item_maps['ew'].items():
                item_no_str = clean_str_item(v.get('item_no', ''))
                if item_no_str in excel_item_nos:
                    selected_ew.append(k)

        if 'cc' in all_item_maps:
            for k, v in all_item_maps['cc'].items():
                item_no_str = clean_str_item(v.get('item_no', ''))
                if item_no_str in excel_item_nos:
                    if any(kw in full_text_str for kw in ['concrete', 'cement', 'c.c', '1:2:4', '1:3:6']):
                        selected_cc.append(k)

        if 'ir' in all_item_maps:
            for k, v in all_item_maps['ir'].items():
                item_no_str = clean_str_item(v.get('item_no', ''))
                if item_no_str in excel_item_nos:
                    if any(kw in full_text_str for kw in ['iron', 'steel', 'rebar', 'w.i', 'kg', 'ton', 'bar']):
                        selected_ir.append(k)

        if 'rc' in all_item_maps:
            for k, v in all_item_maps['rc'].items():
                item_no_str = clean_str_item(v.get('item_no', ''))
                if item_no_str in excel_item_nos:
                    if any(kw in full_text_str for kw in ['reinforced', 'r.c.c', 'rcc', 'pipe', 'post']):
                        selected_rc.append(k)

        if 'rc_m' in all_item_maps:
            for k, v in all_item_maps['rc_m'].items():
                item_no_str = clean_str_item(v.get('item_no', ''))
                if item_no_str in excel_item_nos:
                    if any(kw in full_text_str for kw in ['reinforced', 'r.c.c', 'rcc', 'machine', 'mixer']):
                        selected_rc_m.append(k)

        if 'bw' in all_item_maps:
            for k, v in all_item_maps['bw'].items():
                item_no_str = clean_str_item(v.get('item_no', ''))
                if item_no_str in excel_item_nos:
                    if any(kw in full_text_str for kw in ['brick', 'brickwork', 'brick work', 'masonry', 'အုတ်']):
                        selected_bw.append(k)

        st.session_state['selected_ew'] = selected_ew
        st.session_state['selected_cc'] = selected_cc
        st.session_state['selected_ir'] = selected_ir
        st.session_state['selected_rc'] = selected_rc
        st.session_state['selected_rc_m'] = selected_rc_m
        st.session_state['selected_bw'] = selected_bw

        all_selected_keys = selected_ew + selected_cc + selected_ir + selected_rc + selected_rc_m + selected_bw
        all_items_flat = []
        for cat in ['ew', 'cc', 'ir', 'rc', 'rc_m', 'bw']:
            if cat in all_item_maps:
                for k, item in all_item_maps[cat].items():
                    if k in all_selected_keys:
                        all_items_flat.append(item)

        imported_rows_count = 0
        for idx, item in enumerate(all_items_flat):
            item_no_str = clean_str_item(item.get('item_no', ''))
            rows_state_key = f"rows_data_{item_no_str}_{idx}"

            item_df = valid_rows[valid_rows['clean_item_no'] == item_no_str]

            if not item_df.empty:
                new_rows = []
                for _, r in item_df.iterrows():
                    desc_val = str(r[col_desc]).strip() if pd.notna(r[col_desc]) else "အကွက် ၁"
                    
                    try:
                        no_val = int(float(r[col_no])) if col_no and pd.notna(r[col_no]) else 1
                    except (ValueError, TypeError):
                        no_val = 1

                    def safe_float(val):
                        try:
                            if pd.isna(val) or str(val).strip() in ['-', '', 'nan', 'NaN']:
                                return 0.0
                            return float(str(val).replace(',', ''))
                        except (ValueError, TypeError):
                            return 0.0

                    l_val = safe_float(r[col_l]) if col_l else 0.0
                    b_val = safe_float(r[col_b]) if col_b else 0.0
                    h_val = safe_float(r[col_h]) if col_h else 0.0
                    ded_val = safe_float(r[col_ded]) if col_ded else 0.0
                    
                    type_str = str(r[col_type]).lower() if col_type and pd.notna(r[col_type]) else ""
                    is_ded_row = "ded" in type_str or "minus" in type_str or "sub" in type_str or ded_val > 0

                    new_rows.append({
                        "desc": desc_val,
                        "no": max(1, no_val),
                        "l": l_val,
                        "b": b_val,
                        "h": h_val,
                        "ded": ded_val,
                        "is_deduction_row": is_ded_row
                    })

                if new_rows:
                    st.session_state[rows_state_key] = new_rows
                    imported_rows_count += len(new_rows)

        st.session_state['excel_import_success'] = f"✅ Excel မှ Item များနှင့် အတိုင်းအတာ စာရင်း ({imported_rows_count}) ခုကို အောင်မြင်စွာ ထည့်သွင်းပြီးပါပြီ။"

    except Exception as e:
        err_msg = traceback.format_exc()
        st.session_state['last_excel_error'] = f"❌ အမှားအယွင်း ရှိနေပါသည်: {e}\n\n{err_msg}"


def export_measurement_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Measurement Input Template"

    headers = ["Item No.", "Particular Description", "No.", "L (ft)", "B (ft)", "H (ft)", "Deduction", "Type"]
    header_fill = PatternFill(start_color="1E3A8A", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font

    sample_data = [
        ["1", "Excavation Grid A-1", 2, 10, 5, 4, 0, "Addition"],
        ["1", "Column Box Hole Deduction", 1, 2, 2, 4, 0, "Deduction"],
        ["2", "Foundation Concrete Footing", 4, 6, 6, 1.5, 0, "Addition"],
    ]

    for row_idx, row_vals in enumerate(sample_data, start=2):
        for col_idx, val in enumerate(row_vals, start=1):
            ws.cell(row=row_idx, column=col_idx, value=val)

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def export_measurement_sheet_excel(selected_items_list, st_session_state):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Measurement Sheet"

    ws['A1'] = "DETAIL MEASUREMENT SHEET"
    ws['A1'].font = Font(name='Calibri', size=14, bold=True, color='1F497D')
    
    headers = ["Item No.", "Particular Description", "No.", "L (ft)", "B (ft)", "H (ft)", "Deduction", "Type", "Sub-total"]
    header_fill = PatternFill(start_color="1E3A8A", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                         top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

    row_idx = 3

    for idx, item in enumerate(selected_items_list):
        item_no_str = str(item['item_no'])
        rows_state_key = f"rows_data_{item_no_str}_{idx}"

        unit_str = str(item['unit']).lower().strip()
        is_lumpsum = 'l-s' in unit_str or 'ls' in unit_str or 'lump' in unit_str or 'job' in unit_str
        is_sft = 'sft' in unit_str or 'sq.ft' in unit_str or 'sqft' in unit_str
        is_rft = 'rft' in unit_str or 'lin.ft' in unit_str

        ws.cell(row=row_idx, column=1, value=f"Item {item_no_str} - {item['title']} ({item['unit']})").font = Font(bold=True, size=11)
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=9)
        row_idx += 1

        for col_idx, text in enumerate(headers, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=text)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
        row_idx += 1

        start_data_row = row_idx

        if is_lumpsum:
            ws.cell(row=row_idx, column=1, value=item_no_str)
            ws.cell(row=row_idx, column=2, value="Lumpsum Job")
            ws.cell(row=row_idx, column=3, value=1)
            ws.cell(row=row_idx, column=8, value="Add")
            ws.cell(row=row_idx, column=9, value=f"=C{row_idx}")
            row_idx += 1
        else:
            current_rows = st_session_state.get(rows_state_key, [])
            for r in current_rows:
                ws.cell(row=row_idx, column=1, value=item_no_str)
                ws.cell(row=row_idx, column=2, value=r['desc'])
                ws.cell(row=row_idx, column=3, value=r['no'])
                ws.cell(row=row_idx, column=4, value=r['l'])
                ws.cell(row=row_idx, column=5, value=r['b'] if not is_rft else "-")
                ws.cell(row=row_idx, column=6, value=r['h'] if (not is_sft and not is_rft) else "-")
                ws.cell(row=row_idx, column=7, value=r['ded'])
                ws.cell(row=row_idx, column=8, value="Deduction" if r.get("is_deduction_row") else "Addition")

                if is_rft:
                    mult_expr = f"C{row_idx}*D{row_idx}"
                elif is_sft:
                    mult_expr = f"C{row_idx}*D{row_idx}*E{row_idx}"
                elif 'ton' in unit_str or 'cwt' in unit_str or 'lb' in unit_str or 'kg' in unit_str:
                    mult_expr = f"C{row_idx}*D{row_idx}"
                else:
                    mult_expr = f"C{row_idx}*D{row_idx}*E{row_idx}*F{row_idx}"

                formula_str = f"=IF(H{row_idx}=\"Deduction\", -1 * MAX(0, ({mult_expr}) - G{row_idx}), MAX(0, ({mult_expr}) - G{row_idx}))"
                ws.cell(row=row_idx, column=9, value=formula_str)

                for c in range(1, 10):
                    ws.cell(row=row_idx, column=c).border = thin_border

                row_idx += 1

        end_data_row = row_idx - 1

        ws.cell(row=row_idx, column=2, value=f"Total Quantity ({item['unit']})").font = Font(bold=True)
        ws.cell(row=row_idx, column=9, value=f"=MAX(0, SUM(I{start_data_row}:I{end_data_row}))").font = Font(bold=True)
        ws.cell(row=row_idx, column=9).fill = PatternFill(start_color="FFF2CC", fill_type="solid")
        row_idx += 2

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def export_boq_summary_excel(material_summary, labour_summary):
    wb = openpyxl.Workbook()
    
    ws_mat = wb.active
    ws_mat.title = "Material Summary"
    ws_mat['A1'] = "MATERIAL COST BREAKDOWN (BOQ)"
    ws_mat['A1'].font = Font(size=14, bold=True, color='1F497D')

    headers = ["No.", "Particular Description", "Unit", "Quantity", "Rate (MMK)", "Amount (MMK)"]
    header_fill = PatternFill(start_color="1E3A8A", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")

    for col_idx, h in enumerate(headers, 1):
        cell = ws_mat.cell(row=3, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font

    r_idx = 4
    for idx, (p_name, data) in enumerate(material_summary.items(), start=1):
        ws_mat.cell(row=r_idx, column=1, value=idx)
        ws_mat.cell(row=r_idx, column=2, value=p_name)
        ws_mat.cell(row=r_idx, column=3, value=data['unit'])
        ws_mat.cell(row=r_idx, column=4, value=data['qty'])
        ws_mat.cell(row=r_idx, column=5, value=data['rate'])
        ws_mat.cell(row=r_idx, column=6, value=f"=D{r_idx}*E{r_idx}")
        ws_mat.cell(row=r_idx, column=6).number_format = '#,##0.00'
        r_idx += 1

    ws_mat.cell(row=r_idx, column=2, value="TOTAL MATERIAL COST").font = Font(bold=True)
    ws_mat.cell(row=r_idx, column=6, value=f"=SUM(F4:F{r_idx-1})").font = Font(bold=True)
    ws_mat.cell(row=r_idx, column=6).fill = PatternFill(start_color="D9E1F2", fill_type="solid")
    ws_mat.cell(row=r_idx, column=6).number_format = '#,##0.00'

    ws_lab = wb.create_sheet(title="Labour Summary")
    ws_lab['A1'] = "LABOUR COST BREAKDOWN (BOQ)"
    ws_lab['A1'].font = Font(size=14, bold=True, color='1F497D')

    for col_idx, h in enumerate(headers, 1):
        cell = ws_lab.cell(row=3, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font

    r_idx = 4
    for idx, (p_name, data) in enumerate(labour_summary.items(), start=1):
        ws_lab.cell(row=r_idx, column=1, value=idx)
        ws_lab.cell(row=r_idx, column=2, value=p_name)
        ws_lab.cell(row=r_idx, column=3, value=data['unit'])
        ws_lab.cell(row=r_idx, column=4, value=data['qty'])
        ws_lab.cell(row=r_idx, column=5, value=data['rate'])
        ws_lab.cell(row=r_idx, column=6, value=f"=D{r_idx}*E{r_idx}")
        ws_lab.cell(row=r_idx, column=6).number_format = '#,##0.00'
        r_idx += 1

    ws_lab.cell(row=r_idx, column=2, value="TOTAL LABOUR COST").font = Font(bold=True)
    ws_lab.cell(row=r_idx, column=6, value=f"=SUM(F4:F{r_idx-1})").font = Font(bold=True)
    ws_lab.cell(row=r_idx, column=6).fill = PatternFill(start_color="D9E1F2", fill_type="solid")
    ws_lab.cell(row=r_idx, column=6).number_format = '#,##0.00'

    for ws in [ws_mat, ws_lab]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def main():
    st.set_page_config(
        page_title="Civil Site Estimator & QS Tool", layout="wide", page_icon="🏗️"
    )

    # CSS Customizations
    st.markdown("""
        <style>
            /* Top Banner Styling */
            .main-header {
                background: linear-gradient(135deg, #0284c7, #2563eb);
                padding: 20px;
                border-radius: 12px;
                color: white;
                margin-bottom: 20px;
                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
            }
            .main-header h1 {
                color: white !important;
                font-size: 26px !important;
                margin-bottom: 5px !important;
            }
            .main-header p {
                color: #e0f2fe !important;
                font-size: 14px !important;
            }
            
            /* Card & Box Style */
            .qs-card {
                background-color: #ffffff;
                border: 1px solid #e5e7eb;
                border-radius: 10px;
                padding: 16px;
                margin-bottom: 15px;
                box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            }
            
            /* Metric Box Style */
            div[data-testid="stMetricValue"] {
                font-size: 22px !important;
                font-weight: bold !important;
                color: #1e3a8a !important;
            }
            
            /* Buttons Customization */
            .stButton > button {
                font-size: 15px !important;
                font-weight: 600 !important;
                border-radius: 8px !important;
                padding: 8px 16px !important;
                transition: all 0.2s ease-in-out;
            }
            .stDownloadButton > button {
                font-size: 15px !important;
                font-weight: bold !important;
                border-radius: 8px !important;
                background-color: #0284c7 !important;
                color: white !important;
                width: 100%;
            }
            
            /* Subheaders */
            .section-title {
                color: #1e293b;
                font-weight: 700;
                font-size: 18px;
                border-left: 4px solid #0284c7;
                padding-left: 10px;
                margin-top: 15px;
                margin-bottom: 15px;
            }
        </style>
    """, unsafe_allow_html=True)

    # Modern Header Banner
    st.markdown("""
        <div class="main-header">
            <h1>🏗️ Civil Calculation & QS Tool</h1>
            <p>ဆိုဒ်တွက် အတိုင်းအတာများ ရိုက်ထည့်၍ ကုန်ကျစရိတ်နှင့် Material/Labour BOQ စာရင်း လွယ်ကူစွာ တွက်ချက်ပါ</p>
        </div>
    """, unsafe_allow_html=True)

    # Load Rate Master Data
    earthwork_path = os.path.join(BASE_DIR, "1 Earth Work.xls")
    concrete_path = os.path.join(BASE_DIR, "2 Concrete ( Hand mixed ).xls")
    iron_path = os.path.join(BASE_DIR, "3 Iron and Steel work.xls")
    rc_path = os.path.join(BASE_DIR, "4.1 Reinforced concrete ( Hand mixed ).xls")
    rc_machine_path = os.path.join(BASE_DIR, "4.2 Reinforced concrete ( mixed by machine ).xls")
    brickwork_path = os.path.join(BASE_DIR, "5 Brick work.xls")

    earthwork_items = parse_excel_rates(earthwork_path)
    concrete_items = parse_excel_rates(concrete_path)
    iron_items = parse_excel_rates(iron_path)
    rc_items = parse_excel_rates(rc_path)
    rc_machine_items = parse_excel_rates(rc_machine_path)
    brickwork_items = parse_excel_rates(brickwork_path)

    ew_options = {f"[မြေကျင်း] Item {i['item_no']} - {i['title']}": i for i in earthwork_items}
    cc_options = {f"[ကွန်ကရစ်] Item {i['item_no']} - {i['title']}": i for i in concrete_items}
    ir_options = {f"[သံချည်သံကွေး] Item {i['item_no']} - {i['title']}": i for i in iron_items}
    rc_options = {f"[သံကွန်ကရစ် (လက်စပ်)] Item {i['item_no']} - {i['title']}": i for i in rc_items}
    rc_m_options = {f"[သံကွန်ကရစ် (စက်စပ်)] Item {i['item_no']} - {i['title']}": i for i in rc_machine_items}
    bw_options = {f"[အုတ်စီအုတ်ကိုင်] Item {i['item_no']} - {i['title']}": i for i in brickwork_items}

    all_item_maps = {
        'ew': ew_options, 
        'cc': cc_options, 
        'ir': ir_options, 
        'rc': rc_options,
        'rc_m': rc_m_options,
        'bw': bw_options
    }

    if 'selected_ew' not in st.session_state:
        st.session_state['selected_ew'] = []
    if 'selected_cc' not in st.session_state:
        st.session_state['selected_cc'] = []
    if 'selected_ir' not in st.session_state:
        st.session_state['selected_ir'] = []
    if 'selected_rc' not in st.session_state:
        st.session_state['selected_rc'] = []
    if 'selected_rc_m' not in st.session_state:
        st.session_state['selected_rc_m'] = []
    if 'selected_bw' not in st.session_state:
        st.session_state['selected_bw'] = []

    # Messages
    if st.session_state.get('last_excel_error'):
        st.error(st.session_state['last_excel_error'])

    if st.session_state.get('excel_import_success'):
        st.success(st.session_state['excel_import_success'])

    # ==========================================
    # Main Page Tools
    # ==========================================
    with st.expander("🛠️ အရန်ကိရိယာများနှင့် ပေါက်ဈေး ပြင်ဆင်ရန် (Tools & Rates)", expanded=True):
        tab_rates, tab_calc, tab_conv, tab_upload = st.tabs([
            "⚙️ ပစ္စည်း/လုပ်အားခ ပေါက်ဈေး", 
            "🧮 ဂဏန်းတွက်စက်", 
            "🔄 ယူနစ်ပြောင်းရန်",
            "📥 Excel ဖိုင်တင်ရန်"
        ])

        with tab_rates:
            st.subheader("ပေါက်ဈေး သတ်မှတ်ရန် (ကျပ်)")
            
            col_r1, col_r2, col_r3 = st.columns(3)
            
            with col_r1:
                st.markdown("**👷 လုပ်အားခ ပေါက်ဈေးများ**")
                rate_worker = st.number_input("အလုပ်သမား (ကျပ်)", value=25000.0, step=1000.0)
                rate_digger = st.number_input("မြေကျင်းတူး (ကျပ်)", value=25000.0, step=1000.0)
                rate_mason = st.number_input("ပန်းရံဆရာ (ကျပ်)", value=25000.0, step=1000.0)
                rate_carpenter = st.number_input("လက်သမားဆရာ (ကျပ်)", value=35000.0, step=1000.0)
                rate_maistry = st.number_input("ခေါင်းဆောင် / မေစတရီ (ကျပ်)", value=30000.0, step=1000.0)
                rate_blacksmith = st.number_input("သံချည်သံကွေးဆရာ (ကျပ်)", value=28000.0, step=1000.0)
                rate_welder = st.number_input("ဝိန်းဆရာ (ကျပ်)", value=30000.0, step=1000.0)
                rate_surveyor = st.number_input("တိုင်းတာရေး / Surveyor (ကျပ်)", value=35000.0, step=1000.0)

            with col_r2:
                st.markdown("**🧱 ကွန်ကရစ်၊ အုတ်နှင့် မြေလုပ်ငန်း ပစ္စည်းဈေး**")
                rate_brick = st.number_input("အုတ် (၁ လုံး)", value=200.0, step=10.0)
                rate_cement = st.number_input("ဘိလပ်မြေ (၁ အိတ်)", value=12000.0, step=500.0)
                rate_sand = st.number_input("သဲ (ကျင်း)", value=45000.0, step=1000.0)
                rate_shingle = st.number_input("မြစ်ကျောက် (ကျင်း)", value=85000.0, step=1000.0)
                rate_gravel = st.number_input("ကျောက်စုန်း/ကျောက်စိစစ် (ကျင်း)", value=60000.0, step=1000.0)
                rate_granite = st.number_input("ဂရက်နိုက် ကျောက်စိစစ် (ကျင်း)", value=95000.0, step=1000.0)
                rate_impermo = st.number_input("Impermo ရေကာဆေး", value=3500.0)
                rate_ironite = st.number_input("Ironite ဆေး", value=4000.0)
                rate_timber_scantling = st.number_input("သစ် (Timber scantling)", value=35000.0)
                rate_timber_planks = st.number_input("သစ်ပျဉ် (Timber planks)", value=1200.0)
                rate_nails = st.number_input("သံမှို (Nails)", value=3360.0)

            with col_r3:
                st.markdown("**⚙️ သံနှင့် တည်ဆောက်ရေး ပစ္စည်းဈေး**")
                rate_steel_bar = st.number_input("သံချောင်း/ဘား M.S Bar (၁ တန်)", value=2800000.0, step=10000.0)
                rate_binding_wire = st.number_input("ဘိုင်းဒင်းဝါယာ/သံဇကာနန်း (ပိဿာ/ပေါင်)", value=6500.0)
                rate_structural_steel = st.number_input("Structural Steel (၁ တန်)", value=3000000.0)
                rate_rs_girder = st.number_input("R.S Girder (1 cwt)", value=150000.0)
                rate_carriage = st.number_input("ဆိုဒ်အရောက် သယ်ယူခ (1 cwt)", value=5000.0)
                rate_hoisting = st.number_input("အထက်သို့ တင်/ဆင်ခ (1 cwt)", value=15000.0)

        with tab_calc:
            st.subheader("🧮 အလွယ်တွက်စက်")
            calc_expr = st.text_input("တွက်လိုသည်များကို ရိုက်ထည့်ပါ (ဥပမာ- 10*12.5 + 5):", value="")
            if calc_expr:
                try:
                    allowed_chars = "0123456789+-*/(). "
                    if all(char in allowed_chars for char in calc_expr):
                        res = eval(calc_expr)
                        st.success(f"**အဖြေ = {res:,.4f}**")
                    else:
                        st.error("ဂဏန်းနှင့် သင်္ကေတများသာ ရိုက်ထည့်ပါ။")
                except Exception:
                    st.error("တွက်ချက်မှု အမှားရှိနေပါသည် structure ကို ပြန်စစ်ပါ။")

        with tab_conv:
            st.subheader("🔄 ယူနစ် အပြောင်းအလဲ")
            conv_type = st.selectbox("ပြောင်းလိုသည်ကို ရွေးပါ:", [
                "လက်မ -> ပေ (Inches -> Feet)",
                "စတုရန်းပေ <-> စတုရန်းမီတာ (Sft <-> Sq.m)",
                "ကုဗပေ <-> ကုဗမီတာ (Cft <-> Cu.m)",
                "သံအလေးချိန် (အတန်းအစား -> ကီလို/တန်)"
            ])

            if conv_type == "လက်မ -> ပေ (Inches -> Feet)":
                inch_val = st.number_input("လက်မ (Inches):", min_value=0.0, value=6.0)
                st.info(f"👉 **{inch_val} လက်မ = {inch_val / 12.0:.3f} ပေ**")

            elif conv_type == "စတုရန်းပေ <-> စတုရန်းမီတာ (Sft <-> Sq.m)":
                sft_val = st.number_input("စတုရန်းပေ (Sft):", min_value=0.0, value=100.0)
                st.info(f"👉 **{sft_val:,.2f} Sft = {sft_val / 10.764:.2f} Sq.m**")

            elif conv_type == "ကုဗပေ <-> ကုဗမီတာ (Cft <-> Cu.m)":
                cft_val = st.number_input("ကုဗပေ (Cft):", min_value=0.0, value=100.0)
                st.info(f"👉 **{cft_val:,.2f} Cft = {cft_val / 35.315:.2f} Cu.m**")

            elif conv_type == "သံအလေးချိန် (အတန်းအစား -> ကီလို/တန်)":
                bar_dia = st.selectbox("သံဆိုဒ် အရွယ်အစား:", [
                    "10 mm (3/8\")", "12 mm (1/2\")", "16 mm (5/8\")", "20 mm (3/4\")", "25 mm (1\")"
                ])
                length_ft = st.number_input("စုစုပေါင်း အရှည် (ပေ):", min_value=0.0, value=100.0)

                dia_mm_map = {
                    "10 mm (3/8\")": 10,
                    "12 mm (1/2\")": 12,
                    "16 mm (5/8\")": 16,
                    "20 mm (3/4\")": 20,
                    "25 mm (1\")": 25
                }
                d_mm = dia_mm_map[bar_dia]
                wt_kg_per_ft = (d_mm * d_mm) / 533.0
                total_kg = length_ft * wt_kg_per_ft
                total_ton = total_kg / 1000.0

                st.info(f"👉 **အလေးချိန် = {total_kg:,.2f} ကီလိုဂရမ် ({total_ton:.4f} တန်)**")

        with tab_upload:
            st.subheader("📥 တိုင်းတာပြီး Excel ဖိုင်တင်ရန်")
            template_buffer = export_measurement_template()
            st.download_button(
                label="📄 နမူနာ ပုံစံ (Template) ရယူရန်",
                data=template_buffer,
                file_name="Measurement_Input_Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.divider()
            uploaded_meas_file = st.file_uploader("Excel / CSV ဖိုင် ရွေးပါ:", type=["xlsx", "xls", "csv"])
            if uploaded_meas_file is not None:
                if st.button("🚀 ဖိုင်ထဲမှ စာရင်းများ ဖတ်ယူမည်", type="primary", use_container_width=True):
                    parse_and_auto_select_uploaded_excel(uploaded_meas_file, all_item_maps)
                    st.rerun()

    # Fallback Values for Rates
    rate_worker = locals().get('rate_worker', 25000.0)
    rate_digger = locals().get('rate_digger', 25000.0)
    rate_mason = locals().get('rate_mason', 25000.0)
    rate_carpenter = locals().get('rate_carpenter', 35000.0)
    rate_maistry = locals().get('rate_maistry', 30000.0)
    rate_blacksmith = locals().get('rate_blacksmith', 28000.0)
    rate_welder = locals().get('rate_welder', 30000.0)
    rate_surveyor = locals().get('rate_surveyor', 35000.0)

    rate_brick = locals().get('rate_brick', 200.0)
    rate_cement = locals().get('rate_cement', 12000.0)
    rate_sand = locals().get('rate_sand', 45000.0)
    rate_shingle = locals().get('rate_shingle', 85000.0)
    rate_gravel = locals().get('rate_gravel', 60000.0)
    rate_granite = locals().get('rate_granite', 95000.0)
    rate_impermo = locals().get('rate_impermo', 3500.0)
    rate_ironite = locals().get('rate_ironite', 4000.0)
    rate_timber_scantling = locals().get('rate_timber_scantling', 35000.0)
    rate_timber_planks = locals().get('rate_timber_planks', 1200.0)
    rate_nails = locals().get('rate_nails', 3360.0)

    rate_steel_bar = locals().get('rate_steel_bar', 2800000.0)
    rate_binding_wire = locals().get('rate_binding_wire', 6500.0)
    rate_structural_steel = locals().get('rate_structural_steel', 3000000.0)
    rate_rs_girder = locals().get('rate_rs_girder', 150000.0)
    rate_carriage = locals().get('rate_carriage', 5000.0)
    rate_hoisting = locals().get('rate_hoisting', 15000.0)

    rate_map = {
        "Worker": rate_worker,
        "Worker for carrying and ramming": rate_worker,
        "Worker for watering": rate_worker,
        "Worker for carrying": rate_worker,
        "Digger": rate_digger,
        "Mason": rate_mason,
        "Carpenter": rate_carpenter,
        "Maistry": rate_maistry,
        "Blacksmith": rate_blacksmith,
        "Steel worker": rate_blacksmith,
        "Welder": rate_welder,
        "Surveyor": rate_surveyor,
        "Bricks": rate_brick,
        "Brick": rate_brick,
        "1st class bricks": rate_brick,
        "First class bricks": rate_brick,
        "Best burnt bricks": rate_brick,
        "Cement": rate_cement,
        "Sand": rate_sand,
        "River Shingle (1-1/2\" gauge)": rate_shingle,
        "River Shingle (1/4\" to 3/4\" gauge)": rate_shingle,
        "River Shingle (3/4\" gauge)": rate_shingle,
        "River Shingle (3/4\" to 1-1/2\" gauge)": rate_shingle,
        "Gravel": rate_gravel,
        "1/4\" Granite chipping": rate_granite,
        "Impermo": rate_impermo,
        "Ironite": rate_ironite,
        "Timber scantling": rate_timber_scantling,
        "Tinber planks 1\"": rate_timber_planks,
        "Nails and spikes": rate_nails,
        "Wire Nails": rate_nails,
        "M.S. Bar": rate_steel_bar,
        "Reinforcement Steel": rate_steel_bar,
        "10mm Ø M-S rods": rate_steel_bar,
        "Binding Wire": rate_binding_wire,
        "Binding wire": rate_binding_wire,
        "No.6 G.I plain wire": rate_binding_wire,
        "Coal tar for filling in letters": 3000.0,
        "Shuttering lump sum allowing same form to be use several times": 0.0,
        "Triangular meah RIF style no.245": 1500.0,
        "Triangular meah rif style no.245": 1500.0,
        "Special shuttering for pipe": 0.0,
        "B.R.C no.10 fabric fixed": 2500.0,
        "Shuttering formwork (rate reduce to 1/6 for reason of repeated use)": 500.0,
        "Curing Work for 14 days": 0.0,
        "Structural Steel": rate_structural_steel,
        "R.S. girder": rate_rs_girder,
        "Carriage to site": rate_carriage,
        "Hoisting and fixing": rate_hoisting,
    }

    # ==========================================
    # အဆင့် (၁) - လုပ်ငန်းအမျိုးအစား ရွေးချယ်ခြင်း
    # ==========================================
    st.markdown('<div class="section-title">၁။ တွက်ချက်လိုသော လုပ်ငန်းအမျိုးအစားများ ရွေးပါ</div>', unsafe_allow_html=True)
    
    col_e, col_c, col_i, col_rc, col_rc_m, col_bw = st.columns(6)
    with col_e:
        show_earthwork = st.checkbox("🚜 မြေကျင်းလုပ်ငန်း", value=False)
    with col_c:
        show_concrete = st.checkbox("🧱 ကွန်ကရစ်လုပ်ငန်း", value=False)
    with col_i:
        show_iron = st.checkbox("⚙️ သံချည်သံကွေး", value=False)
    with col_rc:
        show_rc = st.checkbox("🏗️ သံကွန်ကရစ် (လက်စပ်)", value=False)
    with col_rc_m:
        show_rc_m = st.checkbox("⚙️🏗️ သံကွန်ကရစ် (စက်စပ်)", value=False)
    with col_bw:
        show_brickwork = st.checkbox("🧱 အုတ်စီအုတ်ကိုင်", value=False)

    selected_items_list = []

    if show_earthwork and earthwork_items:
        ew_selected = st.multiselect("🚜 မြေကျင်းလုပ်ငန်းမှ တွက်မည်များကို ရွေးပါ:", list(ew_options.keys()), key="selected_ew")
        for key in ew_selected:
            selected_items_list.append(ew_options[key])

    if show_concrete and concrete_items:
        cc_selected = st.multiselect("🧱 ကွန်ကရစ်လုပ်ငန်းမှ တွက်မည်များကို ရွေးပါ:", list(cc_options.keys()), key="selected_cc")
        for key in cc_selected:
            selected_items_list.append(cc_options[key])

    if show_iron and iron_items:
        ir_selected = st.multiselect("⚙️ သံချည်သံကွေးလုပ်ငန်းမှ တွက်မည်များကို ရွေးပါ:", list(ir_options.keys()), key="selected_ir")
        for key in ir_selected:
            selected_items_list.append(ir_options[key])

    if show_rc and rc_items:
        rc_selected = st.multiselect("🏗️ သံကွန်ကရစ် (လက်စပ်) လုပ်ငန်းမှ တွက်မည်များကို ရွေးပါ:", list(rc_options.keys()), key="selected_rc")
        for key in rc_selected:
            selected_items_list.append(rc_options[key])

    if show_rc_m and rc_machine_items:
        rc_m_selected = st.multiselect("⚙️🏗️ သံကွန်ကရစ် (စက်စပ်) လုပ်ငန်းမှ တွက်မည်များကို ရွေးပါ:", list(rc_m_options.keys()), key="selected_rc_m")
        for key in rc_m_selected:
            selected_items_list.append(rc_m_options[key])

    if show_brickwork and brickwork_items:
        bw_selected = st.multiselect("🧱 အုတ်စီအုတ်ကိုင်လုပ်ငန်းမှ တွက်မည်များကို ရွေးပါ:", list(bw_options.keys()), key="selected_bw")
        for key in bw_selected:
            selected_items_list.append(bw_options[key])

    if not selected_items_list:
        st.info("💡 **အကြံပြုချက်**: တွက်ချက်လိုသော အကြောင်းအရာများကို အထက်တွင် ရွေးပေးပါ။ သို့မဟုတ် အထက်ပါ Excel Upload အကွက်မှ ဖိုင် တင်သွင်းပါ။")
        return

    st.divider()

    # ==========================================
    # အဆင့် (၂) - အတိုင်းအတာများ ရိုက်ထည့်ခြင်း
    # ==========================================
    st.markdown('<div class="section-title">📐 ၂။ အတိုင်းအတာများ ရိုက်ထည့်ပါ (Detail Measurement)</div>', unsafe_allow_html=True)

    item_quantities = {}
    ls_custom_rates = {}

    for idx, item in enumerate(selected_items_list):
        item_no_str = str(item['item_no'])
        rows_state_key = f"rows_data_{item_no_str}_{idx}"

        unit_str = str(item['unit']).lower().strip()
        is_lumpsum = 'l-s' in unit_str or 'ls' in unit_str or 'lump' in unit_str or 'job' in unit_str
        is_sft = 'sft' in unit_str or 'sq.ft' in unit_str or 'sqft' in unit_str
        is_rft = 'rft' in unit_str or 'lin.ft' in unit_str

        if rows_state_key not in st.session_state:
            st.session_state[rows_state_key] = [
                {
                    "desc": "အကွက် ၁",
                    "no": 1,
                    "l": 50.0 if is_sft else 10.0,
                    "b": 50.0 if is_sft else 10.0,
                    "h": 5.0,
                    "ded": 0.0,
                    "is_deduction_row": False
                }
            ]

        with st.expander(f"📌 Item {item['item_no']} - {item['title']} [{item['unit']}]", expanded=True):
            meas_rows = []
            item_total_qty = 0.0

            if is_lumpsum:
                c_desc, c_no, c_rate = st.columns([3, 1, 2])
                p_desc = c_desc.text_input("လုပ်ငန်းနေရာ / အမျိုးအစား", value="Lumpsum Job", key=f"desc_{idx}_{item_no_str}")
                no_val = c_no.number_input("အရေအတွက် (ခု)", min_value=1, value=1, key=f"no_{idx}_{item_no_str}")
                ls_rate = c_rate.number_input("တစ်စုတစ်ဝေးတည်း ဈေးနှုန်း (ကျပ်)", min_value=0.0, value=50000.0, step=10000.0, key=f"ls_rate_{idx}_{item_no_str}")
                
                item_total_qty = float(no_val)
                ls_custom_rates[item_no_str] = ls_rate

                meas_rows.append({
                    "လုပ်ငန်းနေရာ": p_desc,
                    "အရေအတွက်": no_val,
                    "အရှည် (L)": "-",
                    "အနံ (B)": "-",
                    "အမြင့် (H)": "-",
                    "အမျိုးအစား": "အပေါင်း",
                    "ရလဒ်": no_val
                })
            else:
                current_rows = st.session_state[rows_state_key]
                row_to_copy = None

                for r_idx, r_data in enumerate(current_rows):
                    st.markdown(f"**🔹 စာကြောင်း ({r_idx+1})**")
                    
                    if is_sft:
                        c_desc, c_no, c_l, c_b = st.columns([2.5, 1, 1, 1])
                        c_ded, c_is_ded, c_cp = st.columns([1.5, 1.5, 1])
                    elif is_rft:
                        c_desc, c_no, c_l = st.columns([3, 1, 1])
                        c_ded, c_is_ded, c_cp = st.columns([1.5, 1.5, 1])
                    else:
                        c_desc, c_no, c_l, c_b, c_h = st.columns([2.5, 1, 1, 1, 1])
                        c_ded, c_is_ded, c_cp = st.columns([1.5, 1.5, 1])

                    p_desc = c_desc.text_input(
                        "နေရာ / အကွက်အမည်",
                        value=r_data["desc"],
                        key=f"desc_{idx}_{r_idx}_{item_no_str}"
                    )
                    no_val = c_no.number_input("အရေအတွက်", min_value=1, value=int(r_data["no"]), key=f"no_{idx}_{r_idx}_{item_no_str}")
                    l_val = c_l.number_input("အရှည် L (ပေ)", min_value=0.0, value=float(r_data["l"]), key=f"l_{idx}_{r_idx}_{item_no_str}")
                    
                    b_val = 0.0
                    if not is_rft:
                        b_val = c_b.number_input("အနံ B (ပေ)", min_value=0.0, value=float(r_data["b"]), key=f"b_{idx}_{r_idx}_{item_no_str}")
                    
                    h_val = 0.0
                    if not is_sft and not is_rft:
                        h_val = c_h.number_input("အမြင့်/အထူ H (ပေ)", min_value=0.0, value=float(r_data["h"]), key=f"h_{idx}_{r_idx}_{item_no_str}")
                    
                    ded_val = c_ded.number_input("အနှုတ်ကျင်း (Deduction)", min_value=0.0, value=float(r_data.get("ded", 0.0)), key=f"ded_{idx}_{r_idx}_{item_no_str}")
                    is_ded_row = c_is_ded.checkbox("➖ အနှုတ်လိုင်း ဖြစ်သည်", value=r_data.get("is_deduction_row", False), key=f"is_ded_{idx}_{r_idx}_{item_no_str}")

                    r_data["desc"] = p_desc
                    r_data["no"] = no_val
                    r_data["l"] = l_val
                    r_data["b"] = b_val
                    r_data["h"] = h_val
                    r_data["ded"] = ded_val
                    r_data["is_deduction_row"] = is_ded_row

                    if c_cp.button("📋 ပွားမည် (Copy)", key=f"copy_{idx}_{r_idx}_{item_no_str}"):
                        row_to_copy = dict(r_data)

                    if is_rft:
                        gross_qty = no_val * l_val
                    elif is_sft:
                        gross_qty = no_val * l_val * b_val
                    elif 'ton' in unit_str or 'cwt' in unit_str or 'lb' in unit_str or 'kg' in unit_str:
                        gross_qty = no_val * l_val
                    else:
                        gross_qty = no_val * l_val * b_val * h_val

                    row_qty = max(0.0, gross_qty - ded_val)

                    if is_ded_row:
                        item_total_qty -= row_qty
                        sub_total_display = -round(row_qty, 2)
                    else:
                        item_total_qty += row_qty
                        sub_total_display = round(row_qty, 2)

                    meas_rows.append({
                        "လုပ်ငန်းနေရာ": p_desc,
                        "အရေအတွက်": no_val,
                        "အရှည် (L)": l_val,
                        "အနံ (B)": b_val if not is_rft else "-",
                        "အမြင့် (H)": h_val if (not is_sft and not is_rft) else "-",
                        "အနှုတ်": ded_val,
                        "အမျိုးအစား": "➖ အနှုတ်" if is_ded_row else "➕ အပေါင်း",
                        "ရလဒ်": sub_total_display
                    })
                    st.markdown("---")

                if row_to_copy is not None:
                    copied_row = dict(row_to_copy)
                    copied_row["desc"] = f"{copied_row['desc']} (Copy)"
                    st.session_state[rows_state_key].append(copied_row)
                    st.rerun()

                col_add, col_rem, _ = st.columns([1.5, 1.5, 3])
                if col_add.button("➕ အကွက်အသစ်ထည့်ရန်", key=f"add_{idx}_{item_no_str}", type="primary"):
                    st.session_state[rows_state_key].append({
                        "desc": f"အကွက် {len(st.session_state[rows_state_key]) + 1}",
                        "no": 1,
                        "l": 50.0 if is_sft else 10.0,
                        "b": 50.0 if is_sft else 10.0,
                        "h": 5.0,
                        "ded": 0.0,
                        "is_deduction_row": False
                    })
                    st.rerun()

                if len(st.session_state[rows_state_key]) > 1 and col_rem.button("➖ အကွက်ပြန်ဖြုတ်ရန်", key=f"rem_{idx}_{item_no_str}"):
                    st.session_state[rows_state_key].pop()
                    st.rerun()

            item_total_qty = max(0.0, item_total_qty)
            item_quantities[item_no_str] = item_total_qty
            
            st.markdown("**📊 တိုင်းတာချက် စာရင်းချုပ်**")
            st.dataframe(pd.DataFrame(meas_rows), use_container_width=True)
            
            st.info(f"💡 **Item {item_no_str} စုစုပေါင်း = `{item_total_qty:,.2f} {item['unit']}`**")

    # Excel Download Button
    meas_excel_buffer = export_measurement_sheet_excel(selected_items_list, st.session_state)
    st.download_button(
        label="📥 Detail Measurement Sheet ကို Excel ဖြင့် ဒေါင်းလုဒ်ရယူရန်",
        data=meas_excel_buffer,
        file_name="Detail_Measurement_Sheet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.divider()

    # ==========================================
    # အဆင့် (၃) - စုစုပေါင်း ကုန်ကျစရိတ် တွက်ချက်ခြင်း
    # ==========================================
    st.markdown('<div class="section-title">📊 ၃။ စုစုပေါင်း ကုန်ကျစရိတ် တွက်ချက်မှု (Rate Analysis)</div>', unsafe_allow_html=True)

    grand_total = 0.0
    material_summary = {}
    labour_summary = {}

    for idx, item in enumerate(selected_items_list):
        item_no = str(item['item_no'])
        measured_qty = item_quantities.get(item_no, 0.0)

        st.markdown(f"#### Item {item_no} - {item['title']}")

        unit_str = str(item['unit']).lower().strip()
        is_lumpsum = 'l-s' in unit_str or 'ls' in unit_str or 'lump' in unit_str or 'job' in unit_str

        display_rows = []
        item_total_cost = 0.0

        display_rows.append({
            "Item No": item_no,
            "အကြောင်းအရာ": item['title'],
            "ယူနစ်": item['unit'],
            "ပမာဏ": f"{measured_qty:,.2f}",
            "နှုန်းထား (ကျပ်)": "",
            "ကျသင့်ငွေ (ကျပ်)": ""
        })

        if is_lumpsum and not item['breakdown']:
            ls_rate = ls_custom_rates.get(item_no, 0.0)
            amount = measured_qty * ls_rate
            item_total_cost = amount

            display_rows.append({
                "Item No": "",
                "အကြောင်းအရာ": f"  └ {item['title']}",
                "ယူနစ်": item['unit'],
                "ပမာဏ": f"{measured_qty:,.2f}",
                "နှုန်းထား (ကျပ်)": f"{ls_rate:,.2f}",
                "ကျသင့်ငွေ (ကျပ်)": f"{amount:,.2f}"
            })

            labour_summary[item['title']] = {
                "unit": item['unit'],
                "qty": measured_qty,
                "rate": ls_rate,
                "amount": amount
            }

        elif item['breakdown']:
            try:
                std_base_qty = float(item['std_qty'])
            except (ValueError, TypeError):
                std_base_qty = 100.0

            # std_base_qty 0 (သို့မဟုတ်) 0 ထက် ငယ်/ဗလာ ဖြစ်နေပါက 0 နှင့် စားမိသည့် Error မတက်စေရန် စစ်ဆေးခြင်း
            if std_base_qty <= 0:
                req_qty = std_qty * measured_qty
            else:
                req_qty = (std_qty / std_base_qty) * measured_qty

            mat_breakdown = []
            lab_breakdown = []

            for row in item['breakdown']:
                part = row['particular']
                std_qty = row['qty']
                u = row['unit']

                req_qty = (std_qty / std_base_qty) * measured_qty
                unit_rate = rate_map.get(part, 0.0)
                amount = req_qty * unit_rate
                item_total_cost += amount

                part_lower = part.lower()
                is_labour = any(k in part_lower for k in LABOUR_KEYWORDS)

                row_data = {
                    "part": part,
                    "unit": u,
                    "qty": req_qty,
                    "rate": unit_rate,
                    "amount": amount
                }

                if is_labour:
                    lab_breakdown.append(row_data)
                else:
                    mat_breakdown.append(row_data)

                target_dict = labour_summary if is_labour else material_summary
                if part not in target_dict:
                    target_dict[part] = {"unit": u, "qty": req_qty, "rate": unit_rate, "amount": amount}
                else:
                    target_dict[part]["qty"] += req_qty
                    target_dict[part]["amount"] += amount

            if mat_breakdown:
                display_rows.append({
                    "Item No": "", "အကြောင်းအရာ": "  📦 ပစ္စည်းစရိတ် (Material)", "ယူနစ်": "", "ပမာဏ": "", "နှုန်းထား (ကျပ်)": "", "ကျသင့်ငွေ (ကျပ်)": ""
                })
                for m in mat_breakdown:
                    display_rows.append({
                        "Item No": "",
                        "အကြောင်းအရာ": f"      {m['part']}",
                        "ယူနစ်": m['unit'],
                        "ပမာဏ": f"{m['qty']:,.2f}",
                        "နှုန်းထား (ကျပ်)": f"{m['rate']:,.2f}" if m['rate'] > 0 else "-",
                        "ကျသင့်ငွေ (ကျပ်)": f"{m['amount']:,.2f}" if m['amount'] > 0 else "-"
                    })

            if lab_breakdown:
                display_rows.append({
                    "Item No": "", "အကြောင်းအရာ": "  👷 လုပ်အားခ (Labour)", "ယူနစ်": "", "ပမာဏ": "", "နှုန်းထား (ကျပ်)": "", "ကျသင့်ငွေ (ကျပ်)": ""
                })
                for l in lab_breakdown:
                    display_rows.append({
                        "Item No": "",
                        "အကြောင်းအရာ": f"      {l['part']}",
                        "ယူနစ်": l['unit'],
                        "ပမာဏ": f"{l['qty']:,.2f}",
                        "နှုန်းထား (ကျပ်)": f"{l['rate']:,.2f}",
                        "ကျသင့်ငွေ (ကျပ်)": f"{l['amount']:,.2f}"
                    })

        display_rows.append({
            "Item No": "",
            "အကြောင်းအရာ": "  💰 စုစုပေါင်း ကုန်ကျစရိတ်",
            "ယူနစ်": "",
            "ပမာဏ": "",
            "နှုန်းထား (ကျပ်)": "",
            "ကျသင့်ငွေ (ကျပ်)": f"**{item_total_cost:,.2f}**"
        })

        st.dataframe(pd.DataFrame(display_rows), use_container_width=True, hide_index=True)
        grand_total += item_total_cost

    st.divider()

    # ==========================================
    # အဆင့် (၄) - BOQ ပစ္စည်းနှင့် လုပ်အားခ စာရင်းချုပ်
    # ==========================================
    st.markdown('<div class="section-title">📜 ၄။ ပစ္စည်းနှင့် လုပ်အားခ စာရင်းချုပ် (BOQ Summary)</div>', unsafe_allow_html=True)

    # Material Section
    st.markdown("### 📦 ၁။ ပစ္စည်းကုန်ကျစရိတ် စာရင်း (Material Summary)")
    mat_rows = []
    total_mat_cost = 0.0
    for idx, (p_name, data) in enumerate(material_summary.items(), start=1):
        mat_rows.append({
            "စဉ်": idx,
            "ပစ္စည်းအမည်": p_name,
            "ယူနစ်": data["unit"],
            "လိုအပ်သော ပမာဏ": f"{data['qty']:,.2f}",
            "နှုန်းထား (ကျပ်)": f"{data['rate']:,.2f}",
            "စုစုပေါင်း (ကျပ်)": f"{data['amount']:,.2f}"
        })
        total_mat_cost += data["amount"]

    if mat_rows:
        st.table(pd.DataFrame(mat_rows))
    else:
        st.info("ပစ္စည်းစရိတ် မရှိပါ။")

    st.divider()

    # Labour Section
    st.markdown("### 👷 ၂။ လုပ်အားခ စာရင်း (Labour Summary)")
    lab_rows = []
    total_lab_cost = 0.0
    for idx, (p_name, data) in enumerate(labour_summary.items(), start=1):
        lab_rows.append({
            "စဉ်": idx,
            "အမျိုးအစား": p_name,
            "ယူနစ်": data["unit"],
            "ပမာဏ": f"{data['qty']:,.2f}",
            "နှုန်းထား (ကျပ်)": f"{data['rate']:,.2f}",
            "စုစုပေါင်း (ကျပ်)": f"{data['amount']:,.2f}"
        })
        total_lab_cost += data["amount"]

    if lab_rows:
        st.table(pd.DataFrame(lab_rows))
    else:
        st.info("လုပ်အားခ စရိတ် မရှိပါ။")

    # BOQ Download Button
    boq_excel_buffer = export_boq_summary_excel(material_summary, labour_summary)
    st.download_button(
        label="📥 BOQ ပစ္စည်းနှင့် လုပ်အားခ စာရင်းချုပ်ကို Excel ဖြင့် ရယူရန်",
        data=boq_excel_buffer,
        file_name="BOQ_Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.divider()

    # Metric Dashboard Summary
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("📦 စုစုပေါင်း ပစ္စည်းဖိုး", f"{total_mat_cost:,.2f} ကျပ်")
    col_m2.metric("👷 စုစုပေါင်း လုပ်အားခ", f"{total_lab_cost:,.2f} ကျပ်")
    col_m3.metric("💰 ပရောဂျက် စုစုပေါင်းစရိတ်", f"{grand_total:,.2f} ကျပ်")


if __name__ == "__main__":
    main()
