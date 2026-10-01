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

# Language Translations Dictionary
TRANSLATIONS = {
    "MM": {
        "title": "🚜 Earthwork QS & Calculation Tool",
        "subtitle": "မြေကျင်းတူး/မြေဖို့ လုပ်ငန်းများအတွက် အတိုင်းအတာများ ရိုက်ထည့်၍ ကုန်ကျစရိတ်နှင့် လုပ်အားခ/ပစ္စည်း BOQ စာရင်း တွက်ချက်ပါ",
        "tools_header": "🛠️ အရန်ကိရိယာများနှင့် ပေါက်ဈေး ပြင်ဆင်ရန် (Tools & Rates)",
        "tab_rates": "⚙️️ ပစ္စည်း/လုပ်အားခ ပေါက်ဈေး",
        "tab_calc": "🧮 ဂဏန်းတွက်စက်",
        "tab_conv": "🔄 ယူနစ်ပြောင်းရန်",
        "tab_upload": "📥 Excel ဖိုင်တင်ရန်",
        "rate_title": "မြေကျင်းလုပ်ငန်း ပေါက်ဈေး သတ်မှတ်ရန် (ကျပ်)",
        "labour_rates": "👷 လုပ်အားခ ပေါက်ဈေးများ",
        "other_rates": "📦 အခြား ကုန်ကျစရိတ်များ",
        "worker": "အလုပ်သမား (ကျပ်)",
        "digger": "မြေကျင်းတူး (ကျပ်)",
        "maistry": "ခေါင်းဆောင် / မေစတရီ (ကျပ်)",
        "sand": "သဲဖို့ (ကျင်း)",
        "carriage": "မြေ/သဲ သယ်ယူခ (ကျင်း)",
        "calc_title": "🧮 အလွယ်တွက်စက်",
        "calc_input": "တွက်လိုသည်များကို ရိုက်ထည့်ပါ (ဥပမာ- 10*12.5 + 5):",
        "calc_result": "အဖြေ = ",
        "calc_err_char": "ဂဏန်းနှင့် သင်္ကေတများသာ ရိုက်ထည့်ပါ။",
        "calc_err_struct": "တွက်ချက်မှု အမှားရှိနေပါသည် structure ကို ပြန်စစ်ပါ။",
        "conv_title": "🔄 ယူနစ် အပြောင်းအလဲ",
        "conv_select": "ပြောင်းလိုသည်ကို ရွေးပါ:",
        "upload_title": "📥 တိုင်းတာပြီး Earthwork Excel ဖိုင်တင်ရန်",
        "btn_download_template": "📄 Earthwork နမူနာ ပုံစံ (Template) ရယူရန်",
        "upload_file_label": "Excel / CSV ဖိုင် ရွေးပါ:",
        "btn_parse_file": "🚀 ဖိုင်ထဲမှ စာရင်းများ ဖတ်ယူမည်",
        "step1": "၁။ တွက်ချက်လိုသော Earthwork Item များ ရွေးပါ",
        "select_items": "🚜 မြေကျင်းလုပ်ငန်းမှ တွက်လိုသည့် Item များကို ရွေးပါ:",
        "step2": "📐 ၂။ အတိုင်းအတာများ ရိုက်ထည့်ပါ (Detail Measurement)",
        "lumpsum_loc": "လုပ်ငန်းနေရာ / အမျိုးအစား",
        "lumpsum_rate": "တစ်စုတစ်ဝေးတည်း ဈေးနှုန်း (ကျပ်)",
        "location": "နေရာ / အကွက်အမည်",
        "qty": "အရေအတွက်",
        "length": "အရှည် L (ပေ)",
        "width": "အနံ B (ပေ)",
        "height": "အမြင့်/အထူ H (ပေ)",
        "deduction": "အနှုတ်ကျင်း (Deduction)",
        "is_ded_check": "➖ အနှုတ်လိုင်း ဖြစ်သည်",
        "btn_copy": "📋 ပွားမည် (Copy)",
        "btn_add_row": "➕ အကွက်အသစ်ထည့်ရန်",
        "btn_rem_row": "➖ အကွက်ပြန်ဖြုတ်ရန်",
        "meas_summary": "📊 တိုင်းတာချက် စာရင်းချုပ်",
        "btn_download_meas": "📥 Detail Earthwork Measurement Sheet ကို Excel ဖြင့် ဒေါင်းလုဒ်ရယူရန်",
        "step3": "📊 ၃။ Earthwork ကုန်ကျစရိတ် တွက်ချက်မှု (Rate Analysis)",
        "mat_cost_title": "  📦 ပစ္စည်းစရိတ် (Material)",
        "lab_cost_title": "  👷 လုပ်အားခ (Labour)",
        "total_item_cost": "  💰 စုစုပေါင်း ကုန်ကျစရိတ်",
        "step4": "📜 ၄။ Earthwork BOQ စာရင်းချုပ်",
        "mat_boq_title": "📦 ၁။ ပစ္စည်းကုန်ကျစရိတ် စာရင်း (Material Summary)",
        "lab_boq_title": "👷 ၂။ လုပ်အားခ စာရင်း (Labour Summary)",
        "no_mat": "ပစ္စည်းစရိတ် မရှိပါ။",
        "no_lab": "လုပ်အားခ စရိတ် မရှိပါ။",
        "btn_download_boq": "📥 Earthwork BOQ စာရင်းချုပ်ကို Excel ဖြင့် ရယူရန်",
        "total_mat_metric": "📦 စုစုပေါင်း ပစ္စည်းဖိုး",
        "total_lab_metric": "👷 စုစုပေါင်း လုပ်အားခ",
        "grand_total_metric": "💰 မြေကျင်းလုပ်ငန်း စုစုပေါင်းစရိတ်",
        "col_no": "စဉ်",
        "col_desc": "အကြောင်းအရာ",
        "col_unit": "ယူနစ်",
        "col_qty": "ပမာဏ",
        "col_rate": "နှုန်းထား (ကျပ်)",
        "col_amount": "ကျသင့်ငွေ (ကျပ်)",
        "add_type": "➕ အပေါင်း",
        "ded_type": "➖ အနှုတ်",
        "unit_kyats": "ကျပ်"
    },
    "EN": {
        "title": "🚜 Earthwork QS & Calculation Tool",
        "subtitle": "Enter dimensions for excavation/backfilling works to calculate costs and BOQ summaries.",
        "tools_header": "🛠️ Tools & Market Rates Configuration",
        "tab_rates": "⚙️ Material/Labour Rates",
        "tab_calc": "🧮 Calculator",
        "tab_conv": "🔄 Unit Converter",
        "tab_upload": "📥 Upload Excel",
        "rate_title": "Set Earthwork Unit Rates (MMK)",
        "labour_rates": "👷 Labour Rates",
        "other_rates": "📦 Other Rates & Materials",
        "worker": "General Worker (MMK)",
        "digger": "Digger (MMK)",
        "maistry": "Maistry / Foreman (MMK)",
        "sand": "Sand Filling (Sud)",
        "carriage": "Carriage to Site (Sud)",
        "calc_title": "🧮 Quick Calculator",
        "calc_input": "Enter expression (e.g., 10*12.5 + 5):",
        "calc_result": "Result = ",
        "calc_err_char": "Please enter numbers and valid mathematical operators only.",
        "calc_err_struct": "Calculation error. Please check the expression structure.",
        "conv_title": "🔄 Unit Converter",
        "conv_select": "Select Conversion Type:",
        "upload_title": "📥 Upload Measured Earthwork Excel File",
        "btn_download_template": "📄 Download Earthwork Measurement Template",
        "upload_file_label": "Select Excel / CSV File:",
        "btn_parse_file": "🚀 Import Data From File",
        "step1": "1. Select Earthwork Items",
        "select_items": "🚜 Select Earthwork Items to Calculate:",
        "step2": "📐 2. Detail Measurement Input",
        "lumpsum_loc": "Work Location / Description",
        "lumpsum_rate": "Lumpsum Rate (MMK)",
        "location": "Location / Item Name",
        "qty": "Quantity (No)",
        "length": "Length L (ft)",
        "width": "Width B (ft)",
        "height": "Height/Depth H (ft)",
        "deduction": "Deduction",
        "is_ded_check": "➖ Deduction Row",
        "btn_copy": "📋 Copy",
        "btn_add_row": "➕ Add Row",
        "btn_rem_row": "➖ Remove Row",
        "meas_summary": "📊 Measurement Summary",
        "btn_download_meas": "📥 Download Measurement Sheet (Excel)",
        "step3": "📊 3. Earthwork Rate Analysis",
        "mat_cost_title": "  📦 Material Cost",
        "lab_cost_title": "  👷 Labour Cost",
        "total_item_cost": "  💰 Total Item Cost",
        "step4": "📜 4. Earthwork BOQ Summary",
        "mat_boq_title": "📦 1. Material Summary",
        "lab_boq_title": "👷 2. Labour Summary",
        "no_mat": "No material cost available.",
        "no_lab": "No labour cost available.",
        "btn_download_boq": "📥 Download BOQ Summary (Excel)",
        "total_mat_metric": "📦 Total Material Cost",
        "total_lab_metric": "👷 Total Labour Cost",
        "grand_total_metric": "💰 Grand Total Earthwork Cost",
        "col_no": "No.",
        "col_desc": "Description",
        "col_unit": "Unit",
        "col_qty": "Quantity",
        "col_rate": "Rate (MMK)",
        "col_amount": "Amount (MMK)",
        "add_type": "➕ Addition",
        "ded_type": "➖ Deduction",
        "unit_kyats": "MMK"
    }
}

@st.cache_data
def parse_excel_rates(file_path):
    if not os.path.exists(file_path):
        return []
    try:
        df_raw = pd.read_excel(file_path, dtype=str)
    except Exception as e:
        st.error(f"File read error ({file_path}): {e}")
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


def parse_and_auto_select_uploaded_excel(uploaded_file, ew_options):
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
            st.session_state['last_excel_error'] = "❌ Column 'Item No.' or 'Particular Description' not found in uploaded file."
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
            st.session_state['last_excel_error'] = "⚠️ No Measurement Data found in Excel file."
            return

        selected_ew = []
        for k, v in ew_options.items():
            item_no_str = clean_str_item(v.get('item_no', ''))
            if item_no_str in excel_item_nos:
                selected_ew.append(k)

        st.session_state['selected_ew'] = selected_ew

        all_items_flat = [ew_options[k] for k in selected_ew if k in ew_options]

        imported_rows_count = 0
        for idx, item in enumerate(all_items_flat):
            item_no_str = clean_str_item(item.get('item_no', ''))
            rows_state_key = f"rows_data_{item_no_str}_{idx}"

            item_df = valid_rows[valid_rows['clean_item_no'] == item_no_str]

            if not item_df.empty:
                new_rows = []
                for _, r in item_df.iterrows():
                    desc_val = str(r[col_desc]).strip() if pd.notna(r[col_desc]) else "Grid 1"
                    
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

        st.session_state['excel_import_success'] = f"✅ Successfully imported ({imported_rows_count}) measurement records from Excel."

    except Exception as e:
        err_msg = traceback.format_exc()
        st.session_state['last_excel_error'] = f"❌ Error occurred: {e}\n\n{err_msg}"


def export_measurement_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Earthwork Measurement Template"

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
        ["2", "Backfilling Work", 1, 50, 20, 2, 0, "Addition"],
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
    ws.title = "Earthwork Measurement"

    ws['A1'] = "EARTHWORK DETAIL MEASUREMENT SHEET"
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
    ws_mat['A1'] = "EARTHWORK MATERIAL COST BREAKDOWN"
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
    ws_lab['A1'] = "EARTHWORK LABOUR COST BREAKDOWN"
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
        page_title="Earthwork QS & Estimator", layout="wide", page_icon="🚜"
    )

    # Language Switcher Sidebar
    st.sidebar.title("🌐 Language / ဘာသာစကား")
    lang = st.sidebar.radio(
        "Select Language / ဘာသာစကား ရွေးချယ်ပါ:",
        options=["🇲🇲 မြန်မာ", "🇬🇧 English"],
        index=0
    )
    
    lang_code = "MM" if "မြန်မာ" in lang else "EN"
    t = TRANSLATIONS[lang_code]

    # CSS Customizations
    st.markdown("""
        <style>
            .main-header {
                background: linear-gradient(135deg, #15803d, #16a34a);
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
                color: #f0fdf4 !important;
                font-size: 14px !important;
            }
            div[data-testid="stMetricValue"] {
                font-size: 22px !important;
                font-weight: bold !important;
                color: #15803d !important;
            }
            .stDownloadButton > button {
                font-size: 15px !important;
                font-weight: bold !important;
                border-radius: 8px !important;
                background-color: #16a34a !important;
                color: white !important;
                width: 100%;
            }
            .section-title {
                color: #14532d;
                font-weight: 700;
                font-size: 18px;
                border-left: 4px solid #16a34a;
                padding-left: 10px;
                margin-top: 15px;
                margin-bottom: 15px;
            }
        </style>
    """, unsafe_allow_html=True)

    # Banner
    st.markdown(f"""
        <div class="main-header">
            <h1>{t['title']}</h1>
            <p>{t['subtitle']}</p>
        </div>
    """, unsafe_allow_html=True)

    # Load Earthwork Data
    earthwork_path = os.path.join(BASE_DIR, "1 Earth Work.xls")
    earthwork_items = parse_excel_rates(earthwork_path)

    prefix_tag = "[မြေကျင်း]" if lang_code == "MM" else "[Earthwork]"
    ew_options = {f"{prefix_tag} Item {i['item_no']} - {i['title']}": i for i in earthwork_items}

    if 'selected_ew' not in st.session_state:
        st.session_state['selected_ew'] = []

    if st.session_state.get('last_excel_error'):
        st.error(st.session_state['last_excel_error'])

    if st.session_state.get('excel_import_success'):
        st.success(st.session_state['excel_import_success'])

    # Tools
    with st.expander(t['tools_header'], expanded=True):
        tab_rates, tab_calc, tab_conv, tab_upload = st.tabs([
            t['tab_rates'], 
            t['tab_calc'], 
            t['tab_conv'],
            t['tab_upload']
        ])

        with tab_rates:
            st.subheader(t['rate_title'])
            col_r1, col_r2 = st.columns(2)
            
            with col_r1:
                st.markdown(f"**{t['labour_rates']}**")
                rate_worker = st.number_input(t['worker'], value=25000.0, step=1000.0)
                rate_digger = st.number_input(t['digger'], value=25000.0, step=1000.0)
                rate_maistry = st.number_input(t['maistry'], value=30000.0, step=1000.0)

            with col_r2:
                st.markdown(f"**{t['other_rates']}**")
                rate_sand = st.number_input(t['sand'], value=45000.0, step=1000.0)
                rate_carriage = st.number_input(t['carriage'], value=15000.0, step=1000.0)

        with tab_calc:
            st.subheader(t['calc_title'])
            calc_expr = st.text_input(t['calc_input'], value="")
            if calc_expr:
                try:
                    allowed_chars = "0123456789+-*/(). "
                    if all(char in allowed_chars for char in calc_expr):
                        res = eval(calc_expr)
                        st.success(f"**{t['calc_result']}{res:,.4f}**")
                    else:
                        st.error(t['calc_err_char'])
                except Exception:
                    st.error(t['calc_err_struct'])

        with tab_conv:
            st.subheader(t['conv_title'])
            conv_options = [
                "Inches -> Feet" if lang_code == "EN" else "လက်မ -> ပေ (Inches -> Feet)",
                "Sft <-> Sq.m" if lang_code == "EN" else "စတုရန်းပေ <-> စတုရန်းမီတာ (Sft <-> Sq.m)",
                "Cft <-> Cu.m" if lang_code == "EN" else "ကုဗပေ <-> ကုဗမီတာ (Cft <-> Cu.m)",
                "Cft -> Sud (%Cft)" if lang_code == "EN" else "ကုဗပေ -> ကျင်း (Cft -> Cu.ft/100)"
            ]
            conv_type = st.selectbox(t['conv_select'], conv_options)

            if "Inches" in conv_type or "လက်မ" in conv_type:
                inch_val = st.number_input("Inches / လက်မ:", min_value=0.0, value=6.0)
                st.info(f"👉 **{inch_val} in = {inch_val / 12.0:.3f} ft**")

            elif "Sft" in conv_type or "စတုရန်းပေ" in conv_type:
                sft_val = st.number_input("Sft / စတုရန်းပေ:", min_value=0.0, value=100.0)
                st.info(f"👉 **{sft_val:,.2f} Sft = {sft_val / 10.764:.2f} Sq.m**")

            elif "Cft <->" in conv_type or "ကုဗပေ <->" in conv_type:
                cft_val = st.number_input("Cft / ကုဗပေ:", min_value=0.0, value=100.0)
                st.info(f"👉 **{cft_val:,.2f} Cft = {cft_val / 35.315:.2f} Cu.m**")

            elif "Sud" in conv_type or "ကျင်း" in conv_type:
                cft_val = st.number_input("Cft / ကုဗပေ:", min_value=0.0, value=500.0)
                st.info(f"👉 **{cft_val:,.2f} Cft = {cft_val / 100.0:.2f} Sud (%Cft)**")

        with tab_upload:
            st.subheader(t['upload_title'])
            template_buffer = export_measurement_template()
            st.download_button(
                label=t['btn_download_template'],
                data=template_buffer,
                file_name="Earthwork_Measurement_Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.divider()
            uploaded_meas_file = st.file_uploader(t['upload_file_label'], type=["xlsx", "xls", "csv"])
            if uploaded_meas_file is not None:
                if st.button(t['btn_parse_file'], type="primary", use_container_width=True):
                    parse_and_auto_select_uploaded_excel(uploaded_meas_file, ew_options)
                    st.rerun()

    # Rate Fallbacks
    rate_worker = locals().get('rate_worker', 25000.0)
    rate_digger = locals().get('rate_digger', 25000.0)
    rate_maistry = locals().get('rate_maistry', 30000.0)
    rate_sand = locals().get('rate_sand', 45000.0)
    rate_carriage = locals().get('rate_carriage', 15000.0)

    rate_map = {
        "Worker": rate_worker,
        "Worker for carrying and ramming": rate_worker,
        "Worker for watering": rate_worker,
        "Worker for carrying": rate_worker,
        "Digger": rate_digger,
        "Maistry": rate_maistry,
        "Sand": rate_sand,
        "Carriage to site": rate_carriage,
    }

    # Step 1: Selection
    st.markdown(f'<div class="section-title">{t["step1"]}</div>', unsafe_allow_html=True)
    
    if not earthwork_items:
        st.warning("⚠️ '1 Earth Work.xls' file not found or contains no data.")
        return

    ew_selected = st.multiselect(t['select_items'], list(ew_options.keys()), key="selected_ew")
    selected_items_list = [ew_options[k] for k in ew_selected]

    if not selected_items_list:
        st.info("💡 **Hint**: Please select Earthwork Items from the Multiselect Box above.")
        return

    st.divider()

    # Step 2: Detail Measurement
    st.markdown(f'<div class="section-title">{t["step2"]}</div>', unsafe_allow_html=True)

    item_quantities = {}
    ls_custom_rates = {}

    for idx, item in enumerate(selected_items_list):
        item_no_str = str(item['item_no'])
        rows_state_key = f"rows_data_{item_no_str}_{idx}"

        unit_str = str(item['unit']).lower().strip()
        is_lumpsum = 'l-s' in unit_str or 'ls' in unit_str or 'lump' in unit_str or 'job' in unit_str
        is_sft = 'sft' in unit_str or 'sq.ft' in unit_str or 'sqft' in unit_str
        is_rft = 'rft' in unit_str or 'lin.ft' in unit_str

        default_desc = "အကွက် ၁" if lang_code == "MM" else "Grid 1"

        if rows_state_key not in st.session_state:
            st.session_state[rows_state_key] = [
                {
                    "desc": default_desc,
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
                p_desc = c_desc.text_input(t['lumpsum_loc'], value="Lumpsum Job", key=f"desc_{idx}_{item_no_str}")
                no_val = c_no.number_input(t['qty'], min_value=1, value=1, key=f"no_{idx}_{item_no_str}")
                ls_rate = c_rate.number_input(t['lumpsum_rate'], min_value=0.0, value=50000.0, step=10000.0, key=f"ls_rate_{idx}_{item_no_str}")
                
                item_total_qty = float(no_val)
                ls_custom_rates[item_no_str] = ls_rate

                meas_rows.append({
                    t['col_desc']: p_desc,
                    t['col_qty']: no_val,
                    "L": "-",
                    "B": "-",
                    "H": "-",
                    "Type": t['add_type'],
                    "Result": no_val
                })
            else:
                current_rows = st.session_state[rows_state_key]
                row_to_copy = None

                for r_idx, r_data in enumerate(current_rows):
                    row_label = f"🔹 Row ({r_idx+1})" if lang_code == "EN" else f"🔹 စာကြောင်း ({r_idx+1})"
                    st.markdown(f"**{row_label}**")
                    
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
                        t['location'],
                        value=r_data["desc"],
                        key=f"desc_{idx}_{r_idx}_{item_no_str}"
                    )
                    no_val = c_no.number_input(t['qty'], min_value=1, value=int(r_data["no"]), key=f"no_{idx}_{r_idx}_{item_no_str}")
                    l_val = c_l.number_input(t['length'], min_value=0.0, value=float(r_data["l"]), key=f"l_{idx}_{r_idx}_{item_no_str}")
                    
                    b_val = 0.0
                    if not is_rft:
                        b_val = c_b.number_input(t['width'], min_value=0.0, value=float(r_data["b"]), key=f"b_{idx}_{r_idx}_{item_no_str}")
                    
                    h_val = 0.0
                    if not is_sft and not is_rft:
                        h_val = c_h.number_input(t['height'], min_value=0.0, value=float(r_data["h"]), key=f"h_{idx}_{r_idx}_{item_no_str}")
                    
                    ded_val = c_ded.number_input(t['deduction'], min_value=0.0, value=float(r_data.get("ded", 0.0)), key=f"ded_{idx}_{r_idx}_{item_no_str}")
                    is_ded_row = c_is_ded.checkbox(t['is_ded_check'], value=r_data.get("is_deduction_row", False), key=f"is_ded_{idx}_{r_idx}_{item_no_str}")

                    r_data["desc"] = p_desc
                    r_data["no"] = no_val
                    r_data["l"] = l_val
                    r_data["b"] = b_val
                    r_data["h"] = h_val
                    r_data["ded"] = ded_val
                    r_data["is_deduction_row"] = is_ded_row

                    if c_cp.button(t['btn_copy'], key=f"copy_{idx}_{r_idx}_{item_no_str}"):
                        row_to_copy = dict(r_data)

                    if is_rft:
                        gross_qty = no_val * l_val
                    elif is_sft:
                        gross_qty = no_val * l_val * b_val
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
                        t['col_desc']: p_desc,
                        t['col_qty']: no_val,
                        "L (ft)": l_val,
                        "B (ft)": b_val if not is_rft else "-",
                        "H (ft)": h_val if (not is_sft and not is_rft) else "-",
                        "Ded": ded_val,
                        "Type": t['ded_type'] if is_ded_row else t['add_type'],
                        "Subtotal": sub_total_display
                    })
                    st.markdown("---")

                if row_to_copy is not None:
                    copied_row = dict(row_to_copy)
                    copied_row["desc"] = f"{copied_row['desc']} (Copy)"
                    st.session_state[rows_state_key].append(copied_row)
                    st.rerun()

                col_add, col_rem, _ = st.columns([1.5, 1.5, 3])
                if col_add.button(t['btn_add_row'], key=f"add_{idx}_{item_no_str}", type="primary"):
                    next_num = len(st.session_state[rows_state_key]) + 1
                    next_desc = f"အကွက် {next_num}" if lang_code == "MM" else f"Grid {next_num}"
                    st.session_state[rows_state_key].append({
                        "desc": next_desc,
                        "no": 1,
                        "l": 50.0 if is_sft else 10.0,
                        "b": 50.0 if is_sft else 10.0,
                        "h": 5.0,
                        "ded": 0.0,
                        "is_deduction_row": False
                    })
                    st.rerun()

                if len(st.session_state[rows_state_key]) > 1 and col_rem.button(t['btn_rem_row'], key=f"rem_{idx}_{item_no_str}"):
                    st.session_state[rows_state_key].pop()
                    st.rerun()

            item_total_qty = max(0.0, item_total_qty)
            item_quantities[item_no_str] = item_total_qty
            
            st.markdown(f"**{t['meas_summary']}**")
            st.dataframe(pd.DataFrame(meas_rows), use_container_width=True)
            
            st.info(f"💡 **Item {item_no_str} Total = `{item_total_qty:,.2f} {item['unit']}`**")

    # Excel Download
    meas_excel_buffer = export_measurement_sheet_excel(selected_items_list, st.session_state)
    st.download_button(
        label=t['btn_download_meas'],
        data=meas_excel_buffer,
        file_name="Earthwork_Measurement_Sheet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.divider()

    # Step 3: Cost Analysis
    st.markdown(f'<div class="section-title">{t["step3"]}</div>', unsafe_allow_html=True)

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
            t['col_desc']: item['title'],
            t['col_unit']: item['unit'],
            t['col_qty']: f"{measured_qty:,.2f}",
            t['col_rate']: "",
            t['col_amount']: ""
        })

        if is_lumpsum and not item['breakdown']:
            ls_rate = ls_custom_rates.get(item_no, 0.0)
            amount = measured_qty * ls_rate
            item_total_cost = amount

            display_rows.append({
                "Item No": "",
                t['col_desc']: f"  └ {item['title']}",
                t['col_unit']: item['unit'],
                t['col_qty']: f"{measured_qty:,.2f}",
                t['col_rate']: f"{ls_rate:,.2f}",
                t['col_amount']: f"{amount:,.2f}"
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

            mat_breakdown = []
            lab_breakdown = []

            for row in item['breakdown']:
                part = row['particular']
                std_qty = row['qty']
                u = row['unit']

                if std_base_qty <= 0:
                    req_qty = std_qty * measured_qty
                else:
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
                    "Item No": "", t['col_desc']: t['mat_cost_title'], t['col_unit']: "", t['col_qty']: "", t['col_rate']: "", t['col_amount']: ""
                })
                for m in mat_breakdown:
                    display_rows.append({
                        "Item No": "",
                        t['col_desc']: f"      {m['part']}",
                        t['col_unit']: m['unit'],
                        t['col_qty']: f"{m['qty']:,.2f}",
                        t['col_rate']: f"{m['rate']:,.2f}" if m['rate'] > 0 else "-",
                        t['col_amount']: f"{m['amount']:,.2f}" if m['amount'] > 0 else "-"
                    })

            if lab_breakdown:
                display_rows.append({
                    "Item No": "", t['col_desc']: t['lab_cost_title'], t['col_unit']: "", t['col_qty']: "", t['col_rate']: "", t['col_amount']: ""
                })
                for l in lab_breakdown:
                    display_rows.append({
                        "Item No": "",
                        t['col_desc']: f"      {l['part']}",
                        t['col_unit']: l['unit'],
                        t['col_qty']: f"{l['qty']:,.2f}",
                        t['col_rate']: f"{l['rate']:,.2f}",
                        t['col_amount']: f"{l['amount']:,.2f}"
                    })

        display_rows.append({
            "Item No": "",
            t['col_desc']: t['total_item_cost'],
            t['col_unit']: "",
            t['col_qty']: "",
            t['col_rate']: "",
            t['col_amount']: f"**{item_total_cost:,.2f}**"
        })

        st.dataframe(pd.DataFrame(display_rows), use_container_width=True, hide_index=True)
        grand_total += item_total_cost

    st.divider()

    # Step 4: BOQ Summary
    st.markdown(f'<div class="section-title">{t["step4"]}</div>', unsafe_allow_html=True)

    st.markdown(f"### {t['mat_boq_title']}")
    mat_rows = []
    total_mat_cost = 0.0
    for idx, (p_name, data) in enumerate(material_summary.items(), start=1):
        mat_rows.append({
            t['col_no']: idx,
            t['col_desc']: p_name,
            t['col_unit']: data["unit"],
            t['col_qty']: f"{data['qty']:,.2f}",
            t['col_rate']: f"{data['rate']:,.2f}",
            t['col_amount']: f"{data['amount']:,.2f}"
        })
        total_mat_cost += data["amount"]

    if mat_rows:
        st.table(pd.DataFrame(mat_rows))
    else:
        st.info(t['no_mat'])

    st.divider()

    st.markdown(f"### {t['lab_boq_title']}")
    lab_rows = []
    total_lab_cost = 0.0
    for idx, (p_name, data) in enumerate(labour_summary.items(), start=1):
        lab_rows.append({
            t['col_no']: idx,
            t['col_desc']: p_name,
            t['col_unit']: data["unit"],
            t['col_qty']: f"{data['qty']:,.2f}",
            t['col_rate']: f"{data['rate']:,.2f}",
            t['col_amount']: f"{data['amount']:,.2f}"
        })
        total_lab_cost += data["amount"]

    if lab_rows:
        st.table(pd.DataFrame(lab_rows))
    else:
        st.info(t['no_lab'])

    boq_excel_buffer = export_boq_summary_excel(material_summary, labour_summary)
    st.download_button(
        label=t['btn_download_boq'],
        data=boq_excel_buffer,
        file_name="Earthwork_BOQ_Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.divider()

    # Dashboard Metrics
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric(t['total_mat_metric'], f"{total_mat_cost:,.2f} {t['unit_kyats']}")
    col_m2.metric(t['total_lab_metric'], f"{total_lab_cost:,.2f} {t['unit_kyats']}")
    col_m3.metric(t['grand_total_metric'], f"{grand_total:,.2f} {t['unit_kyats']}")


if __name__ == "__main__":
    main()
