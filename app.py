import os
import io
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


# =========================================================
# Excel Helper Functions
# =========================================================

def export_measurement_template():
    """အသုံးပြုသူများ Measurement Upload လုပ်ရန် နမူနာ Excel Template ထုတ်ပေးခြင်း"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Measurement Input Template"

    headers = ["Item No.", "Particular Description", "No.", "L (ft)", "B (ft)", "H (ft)", "Deduction", "Type"]
    
    header_fill = PatternFill(start_color="366092", fill_type="solid")
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


import os
import io
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
    """
    Excel/CSV ဖိုင် တင်လိုက်သည်နှင့် Normal Template ရော Export ထုတ်ထားသော Measurement Sheet ကိုပါ
    အလိုအလျောက် ရှာဖွေဖတ်ရှုပြီး Auto-Select & Import လုပ်ပေးသည့် Function
    """
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            # First, check raw dataframe without header
            df_raw = pd.read_excel(uploaded_file, header=None)
            
            # Find row index containing 'item' in column values
            header_row_idx = None
            for idx, row in df_raw.iterrows():
                row_vals = row.astype(str).str.lower().tolist()
                if any('item' in v for v in row_vals if v):
                    header_row_idx = idx
                    break
            
            if header_row_idx is not None:
                df = pd.read_excel(uploaded_file, header=header_row_idx)
            else:
                df = pd.read_excel(uploaded_file)

        # Clean column names
        df.columns = [str(c).strip().lower() for c in df.columns]
        
        col_item = next((c for c in df.columns if 'item' in c), None)
        col_desc = next((c for c in df.columns if 'desc' in c or 'particular' in c), None)
        col_no = next((c for c in df.columns if 'no' in c), None)
        col_l = next((c for c in df.columns if 'l' in c or 'length' in c), None)
        col_b = next((c for c in df.columns if 'b' in c or 'breadth' in c or 'width' in c), None)
        col_h = next((c for c in df.columns if 'h' in c or 'height' in c or 'depth' in c), None)
        col_ded = next((c for c in df.columns if 'ded' in c), None)
        col_type = next((c for c in df.columns if 'type' in c or 'deduction' in c), None)

        if not col_item:
            st.error("❌ တင်သွင်းသော Excel ဖိုင်တွင် 'Item No.' Column မပါဝင်ပါ။ အခြား Template ဖိုင် သုံးကြည့်ပါ။")
            return

        # Clean values & filter out empty or total summary rows
        valid_rows = df[df[col_item].notna() & (df[col_item].astype(str).str.strip() != '')].copy()
        valid_rows = valid_rows[~valid_rows[col_item].astype(str).str.lower().str.contains('item no|total|detail')]

        # Excel ဖိုင်ထဲတွင် ပါဝင်သော Item No များ
        excel_item_nos = valid_rows[col_item].astype(str).str.strip().unique().tolist()

        if not excel_item_nos:
            st.warning("⚠️ Excel ဖိုင်ထဲတွင် Item No. များ ရှာမတွေ့ပါ။")
            return

        # Multiselect State များကို Auto-Select လုပ်ပေးရန် ပြင်ဆင်ခြင်း
        st.session_state['selected_ew'] = [k for k, v in all_item_maps['ew'].items() if str(v['item_no']).strip() in excel_item_nos]
        st.session_state['selected_cc'] = [k for k, v in all_item_maps['cc'].items() if str(v['item_no']).strip() in excel_item_nos]
        st.session_state['selected_ir'] = [k for k, v in all_item_maps['ir'].items() if str(v['item_no']).strip() in excel_item_nos]

        # တောက်လျှောက် Session State Measurement Row Data များဖြည့်သွင်းခြင်း
        all_items_flat = []
        for category in ['ew', 'cc', 'ir']:
            for key, item in all_item_maps[category].items():
                item_no_str = str(item['item_no']).strip()
                if item_no_str in excel_item_nos:
                    all_items_flat.append(item)

        imported_rows_count = 0
        for idx, item in enumerate(all_items_flat):
            item_no_str = str(item['item_no']).strip()
            rows_state_key = f"rows_data_{item_no_str}_{idx}"

            item_df = valid_rows[valid_rows[col_item].astype(str).str.strip() == item_no_str]

            if not item_df.empty:
                new_rows = []
                for _, r in item_df.iterrows():
                    desc_val = str(r[col_desc]).strip() if col_desc and pd.notna(r[col_desc]) else "Section"
                    
                    try:
                        no_val = int(float(r[col_no])) if col_no and pd.notna(r[col_no]) else 1
                    except (ValueError, TypeError):
                        no_val = 1

                    def safe_float(val):
                        try:
                            return float(val)
                        except (ValueError, TypeError):
                            return 0.0

                    l_val = safe_float(r[col_l]) if col_l else 0.0
                    b_val = safe_float(r[col_b]) if col_b else 0.0
                    h_val = safe_float(r[col_h]) if col_h else 0.0
                    ded_val = safe_float(r[col_ded]) if col_ded else 0.0
                    
                    type_str = str(r[col_type]).lower() if col_type and pd.notna(r[col_type]) else ""
                    is_ded_row = "ded" in type_str or "minus" in type_str or "sub" in type_str

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

        st.success(f"✅ Excel မှ Item များနှင့် Measurement Row ({imported_rows_count}) ခုကို Auto-Select ပြုလုပ်ပြီးပါပြီ။")

    except Exception as e:
        st.error(f"❌ Excel ဖတ်ရှုရာတွင် အမှားအယွင်းရှိပါသည်: {e}")


def export_measurement_sheet_excel(selected_items_list, st_session_state):
    """Detail Measurement Sheet ကို Excel Formula များဖြင့် Export ထုတ်ပေးသည့် Function"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Measurement Sheet"

    ws['A1'] = "DETAIL MEASUREMENT SHEET"
    ws['A1'].font = Font(name='Calibri', size=14, bold=True, color='1F497D')
    
    headers = ["Item No.", "Particular Description", "No.", "L (ft)", "B (ft)", "H (ft)", "Deduction", "Type", "Sub-total"]
    header_fill = PatternFill(start_color="366092", fill_type="solid")
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
    """BOQ Material & Labour Summary ကို Excel Formula ပါဝင်အောင် Export ထုတ်ပေးသည့် Function"""
    wb = openpyxl.Workbook()
    
    ws_mat = wb.active
    ws_mat.title = "Material Summary"
    ws_mat['A1'] = "MATERIAL COST BREAKDOWN (BOQ)"
    ws_mat['A1'].font = Font(size=14, bold=True, color='1F497D')

    headers = ["No.", "Particular Description", "Unit", "Quantity", "Rate (MMK)", "Amount (MMK)"]
    header_fill = PatternFill(start_color="366092", fill_type="solid")
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
        page_title="QS & Rate Analysis System", layout="wide", page_icon="🏗️"
    )

    st.title("🏗️ Estimation, QS & Rate Analysis System")

    # Load Rate Master Data
    earthwork_path = os.path.join(BASE_DIR, "1 Earth Work.xls")
    concrete_path = os.path.join(BASE_DIR, "2 Concrete ( Hand mixed ).xls")
    iron_path = os.path.join(BASE_DIR, "3 Iron and Steel work.xls")

    earthwork_items = parse_excel_rates(earthwork_path)
    concrete_items = parse_excel_rates(concrete_path)
    iron_items = parse_excel_rates(iron_path)

    ew_options = {f"[Earth] Item {i['item_no']} - {i['title']}": i for i in earthwork_items}
    cc_options = {f"[Concrete] Item {i['item_no']} - {i['title']}": i for i in concrete_items}
    ir_options = {f"[Iron] Item {i['item_no']} - {i['title']}": i for i in iron_items}

    all_item_maps = {'ew': ew_options, 'cc': cc_options, 'ir': ir_options}

    # Session State များ Initialise ပြုလုပ်ခြင်း
    if 'selected_ew' not in st.session_state:
        st.session_state['selected_ew'] = []
    if 'selected_cc' not in st.session_state:
        st.session_state['selected_cc'] = []
    if 'selected_ir' not in st.session_state:
        st.session_state['selected_ir'] = []

    # ==========================================
    # Sidebar Tools: Upload Data, Master Rates & Calc
    # ==========================================
    st.sidebar.title("🛠️ Tools & Settings")
    tab_upload, tab_rates, tab_calc, tab_conv = st.sidebar.tabs(["📥 Upload Excel", "⚙️ Rates", "🧮 Calc", "🔄 Converter"])

    with tab_upload:
        st.header("📥 Upload Measurement Excel")
        st.write("အသင့်ပြင်ထားသော Excel တင်လိုက်ပါက Item များနှင့် Measurement များ Auto-Select ဖြစ်သွားပါမည်။")

        template_buffer = export_measurement_template()
        st.download_button(
            label="📄 Download Input Template Excel",
            data=template_buffer,
            file_name="Measurement_Input_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider()

        uploaded_meas_file = st.file_uploader("Excel သို့မဟုတ် CSV ဖိုင် တင်ရန်:", type=["xlsx", "xls", "csv"])

        # Excel ဖိုင်တင်လိုက်ပါက Auto Select & Process လုပ်ဆောင်ရန် Button
        if uploaded_meas_file is not None:
            if st.button("🚀 Auto-Select & Import Items", type="primary"):
                parse_and_auto_select_uploaded_excel(uploaded_meas_file, all_item_maps)
                st.rerun()

    with tab_rates:
        st.header("Master Unit Rates (MMK)")
        
        st.subheader("👷 Labour Rates")
        rate_worker = st.number_input("Worker", value=25000.0)
        rate_digger = st.number_input("Digger", value=25000.0)
        rate_mason = st.number_input("Mason", value=25000.0)
        rate_carpenter = st.number_input("Carpenter", value=35000.0)
        rate_maistry = st.number_input("Maistry", value=30000.0)
        rate_blacksmith = st.number_input("Blacksmith / Steel Worker", value=28000.0)
        rate_welder = st.number_input("Welder", value=30000.0)
        rate_surveyor = st.number_input("Surveyor", value=35000.0)

        st.subheader("🧱 Concrete & Earth Materials")
        rate_cement = st.number_input("Cement", value=12000.0)
        rate_sand = st.number_input("Sand", value=45000.0)
        rate_shingle = st.number_input("River Shingle", value=85000.0)
        rate_gravel = st.number_input("Gravel", value=60000.0)
        rate_granite = st.number_input("Granite chipping", value=95000.0)
        rate_impermo = st.number_input("Impermo", value=3500.0)
        rate_ironite = st.number_input("Ironite", value=4000.0)
        rate_timber_scantling = st.number_input("Timber scantling", value=35000.0)
        rate_timber_planks = st.number_input("Timber planks", value=1200.0)
        rate_nails = st.number_input("Nails and spikes / Wire Nails", value=3360.0)

        st.subheader("⚙️ Iron & Structural Rates")
        rate_steel_bar = st.number_input("M.S. Bar / Reinforcement (Ton)", value=2800000.0)
        rate_binding_wire = st.number_input("Binding Wire (Viss/Ib)", value=6500.0)
        rate_structural_steel = st.number_input("Structural Steel (Ton)", value=3000000.0)
        rate_rs_girder = st.number_input("R.S. girder (per cwt)", value=150000.0)
        rate_carriage = st.number_input("Carriage to site (per cwt)", value=5000.0)
        rate_hoisting = st.number_input("Hoisting and fixing (per cwt)", value=15000.0)

    with tab_calc:
        st.subheader("🧮 Quick Calculator")
        calc_expr = st.text_input("Expression ရိုက်ပါ (e.g. 10*12.5 + 5):", value="")
        if calc_expr:
            try:
                allowed_chars = "0123456789+-*/(). "
                if all(char in allowed_chars for char in calc_expr):
                    res = eval(calc_expr)
                    st.success(f"**Result = {res:,.4f}**")
                else:
                    st.error("သင်္ချာ ကိန်းဂဏန်းများသာ ရိုက်ထည့်ပါ။")
            except Exception as e:
                st.error("တွက်ချက်မှု မမှန်ကန်ပါ။")

    with tab_conv:
        st.subheader("🔄 Unit Converter")
        conv_type = st.selectbox("Convert Type:", [
            "Inches -> Feet",
            "Sft <-> Sq.m",
            "Cft <-> Cu.m",
            "Steel Weight (Dia -> Kg/Ton)"
        ])

        if conv_type == "Inches -> Feet":
            inch_val = st.number_input("Inches (လက်မ):", min_value=0.0, value=6.0)
            st.info(f"👉 **{inch_val} inches = {inch_val / 12.0:.3f} ft**")

        elif conv_type == "Sft <-> Sq.m":
            sft_val = st.number_input("Sft:", min_value=0.0, value=100.0)
            st.info(f"👉 **{sft_val:,.2f} Sft = {sft_val / 10.764:.2f} Sq.m**")

        elif conv_type == "Cft <-> Cu.m":
            cft_val = st.number_input("Cft:", min_value=0.0, value=100.0)
            st.info(f"👉 **{cft_val:,.2f} Cft = {cft_val / 35.315:.2f} Cu.m**")

        elif conv_type == "Steel Weight (Dia -> Kg/Ton)":
            bar_dia = st.selectbox("Bar Size:", [
                "10 mm (3/8\")", "12 mm (1/2\")", "16 mm (5/8\")", "20 mm (3/4\")", "25 mm (1\")"
            ])
            length_ft = st.number_input("Total Length (ft):", min_value=0.0, value=100.0)

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

            st.info(f"👉 **Weight = {total_kg:,.2f} Kg ({total_ton:.4f} Ton)**")

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
        "Binding Wire": rate_binding_wire,
        "Structural Steel": rate_structural_steel,
        "R.S. girder": rate_rs_girder,
        "Carriage to site": rate_carriage,
        "Hoisting and fixing": rate_hoisting,
    }

    # ==========================================
    # 1. လုပ်ငန်းအမျိုးအစား Checkbox များ
    # ==========================================
    st.subheader("📋 ၁။ တွက်ချက်လိုသော လုပ်ငန်းအမျိုးအစားများ ရွေးချယ်ပါ")
    col_e, col_c, col_i = st.columns(3)
    show_earthwork = col_e.checkbox("🚜 Earth Work", value=True)
    show_concrete = col_c.checkbox("🧱 Concrete Work", value=True)
    show_iron = col_i.checkbox("⚙️ Iron & Steel Work", value=True)

    selected_items_list = []

    if show_earthwork and earthwork_items:
        ew_selected = st.multiselect("🚜 Earth Work မှ တွက်မည်များ ရွေးရန်:", list(ew_options.keys()), key="selected_ew")
        for key in ew_selected:
            selected_items_list.append(ew_options[key])

    if show_concrete and concrete_items:
        cc_selected = st.multiselect("🧱 Concrete Work မှ တွက်မည်များ ရွေးရန်:", list(cc_options.keys()), key="selected_cc")
        for key in cc_selected:
            selected_items_list.append(cc_options[key])

    if show_iron and iron_items:
        ir_selected = st.multiselect("⚙️ Iron & Steel Work မှ တွက်မည်များ ရွေးရန်:", list(ir_options.keys()), key="selected_ir")
        for key in ir_selected:
            selected_items_list.append(ir_options[key])

    if not selected_items_list:
        st.info("👉 တွက်ချက်လိုသော Item များကို အထက်တွင် ရွေးချယ်ပေးပါ သို့မဟုတ် Excel ဖိုင် တင်ပြီး '🚀 Auto-Select & Import Items' ကို နှိပ်ပါ။")
        return

    st.divider()

    # ==========================================
    # 2. Detail Measurement Sheet Input
    # ==========================================
    st.subheader("📐 ၂။ Detail Measurement Sheet (အတိုင်းအတာများ ရိုက်ထည့်ပါ)")

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
                    "desc": "Section 1",
                    "no": 1,
                    "l": 50.0 if is_sft else 10.0,
                    "b": 50.0 if is_sft else 10.0,
                    "h": 5.0,
                    "ded": 0.0,
                    "is_deduction_row": False
                }
            ]

        with st.expander(f"📌 {idx+1}. Item {item['item_no']} - {item['title']} ({item['unit']})", expanded=True):
            meas_rows = []
            item_total_qty = 0.0

            if is_lumpsum:
                c_desc, c_no, c_rate = st.columns([3, 1, 2])
                p_desc = c_desc.text_input("Particular Name", value="Lumpsum Job", key=f"desc_{idx}_{item_no_str}")
                no_val = c_no.number_input("Job / Qty", min_value=1, value=1, key=f"no_{idx}_{item_no_str}")
                ls_rate = c_rate.number_input("Lump Sum Rate (MMK)", min_value=0.0, value=50000.0, step=10000.0, key=f"ls_rate_{idx}_{item_no_str}")
                
                item_total_qty = float(no_val)
                ls_custom_rates[item_no_str] = ls_rate

                meas_rows.append({
                    "Particular": p_desc,
                    "No": no_val,
                    "L (ft)": "-",
                    "B (ft)": "-",
                    "H (ft)": "-",
                    "Type": "Add",
                    "Sub-total": no_val
                })
            else:
                current_rows = st.session_state[rows_state_key]
                row_to_copy = None

                for r_idx, r_data in enumerate(current_rows):
                    if is_sft:
                        c_desc, c_no, c_l, c_b, c_ded, c_is_ded, c_cp = st.columns([2, 0.8, 0.8, 0.8, 0.8, 1, 0.8])
                    elif is_rft:
                        c_desc, c_no, c_l, c_ded, c_is_ded, c_cp = st.columns([2.5, 0.8, 0.8, 0.8, 1, 0.8])
                    else:
                        c_desc, c_no, c_l, c_b, c_h, c_ded, c_is_ded, c_cp = st.columns([2, 0.8, 0.8, 0.8, 0.8, 0.8, 1, 0.8])

                    p_desc = c_desc.text_input(
                        "Particular Name",
                        value=r_data["desc"],
                        key=f"desc_{idx}_{r_idx}_{item_no_str}"
                    )
                    no_val = c_no.number_input("No", min_value=1, value=int(r_data["no"]), key=f"no_{idx}_{r_idx}_{item_no_str}")
                    l_val = c_l.number_input("L (ft)", min_value=0.0, value=float(r_data["l"]), key=f"l_{idx}_{r_idx}_{item_no_str}")
                    
                    b_val = 0.0
                    if not is_rft:
                        b_val = c_b.number_input("B (ft)", min_value=0.0, value=float(r_data["b"]), key=f"b_{idx}_{r_idx}_{item_no_str}")
                    
                    h_val = 0.0
                    if not is_sft and not is_rft:
                        h_val = c_h.number_input("H (ft)", min_value=0.0, value=float(r_data["h"]), key=f"h_{idx}_{r_idx}_{item_no_str}")
                    
                    ded_val = c_ded.number_input("Deduction", min_value=0.0, value=float(r_data.get("ded", 0.0)), key=f"ded_{idx}_{r_idx}_{item_no_str}")

                    st.markdown("<div style='padding-top: 28px;'></div>", unsafe_allow_html=True)
                    is_ded_row = c_is_ded.checkbox("➖ Deduction Row", value=r_data.get("is_deduction_row", False), key=f"is_ded_{idx}_{r_idx}_{item_no_str}")

                    r_data["desc"] = p_desc
                    r_data["no"] = no_val
                    r_data["l"] = l_val
                    r_data["b"] = b_val
                    r_data["h"] = h_val
                    r_data["ded"] = ded_val
                    r_data["is_deduction_row"] = is_ded_row

                    if c_cp.button("📋 Copy", key=f"copy_{idx}_{r_idx}_{item_no_str}"):
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
                        "Particular": p_desc,
                        "No": no_val,
                        "L (ft)": l_val,
                        "B (ft)": b_val if not is_rft else "-",
                        "H (ft)": h_val if (not is_sft and not is_rft) else "-",
                        "Deduction": ded_val,
                        "Row Type": "➖ Deduction" if is_ded_row else "➕ Addition",
                        "Sub-total": sub_total_display
                    })

                if row_to_copy is not None:
                    copied_row = dict(row_to_copy)
                    copied_row["desc"] = f"{copied_row['desc']} (Copy)"
                    st.session_state[rows_state_key].append(copied_row)
                    st.rerun()

                col_add, col_rem, col_blank = st.columns([1, 1, 4])
                if col_add.button("➕ Add Particular Row", key=f"add_{idx}_{item_no_str}"):
                    st.session_state[rows_state_key].append({
                        "desc": f"Section {len(st.session_state[rows_state_key]) + 1}",
                        "no": 1,
                        "l": 50.0 if is_sft else 10.0,
                        "b": 50.0 if is_sft else 10.0,
                        "h": 5.0,
                        "ded": 0.0,
                        "is_deduction_row": False
                    })
                    st.rerun()

                if len(st.session_state[rows_state_key]) > 1 and col_rem.button("➖ Remove Row", key=f"rem_{idx}_{item_no_str}"):
                    st.session_state[rows_state_key].pop()
                    st.rerun()

            item_total_qty = max(0.0, item_total_qty)
            item_quantities[item_no_str] = item_total_qty
            st.dataframe(pd.DataFrame(meas_rows), use_container_width=True)
            st.markdown(f"**Total Quantity for Item {item_no_str} = `{item_total_qty:,.2f} {item['unit']}`**")

    # 📥 Download Measurement Sheet Excel (With Formula)
    meas_excel_buffer = export_measurement_sheet_excel(selected_items_list, st.session_state)
    st.download_button(
        label="📥 Download Measurement Sheet (Excel With Formulas)",
        data=meas_excel_buffer,
        file_name="Detail_Measurement_Sheet_Formulas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.divider()

    # ==========================================
    # 3. အစဉ်လိုက် Cost Breakdown & Total Estimate
    # ==========================================
    st.subheader("📊 ၃။ စုစုပေါင်း Rate Analysis & Cost Estimate")

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
            "Item": item_no,
            "Particular": item['title'],
            "Unit": item['unit'],
            "Quantity": f"{measured_qty:,.2f}",
            "Rate (MMK)": "",
            "Per": "",
            "Amount (MMK)": ""
        })

        if is_lumpsum and not item['breakdown']:
            ls_rate = ls_custom_rates.get(item_no, 0.0)
            amount = measured_qty * ls_rate
            item_total_cost = amount

            display_rows.append({
                "Item": "",
                "Particular": f"  └ {item['title']}",
                "Unit": item['unit'],
                "Quantity": f"{measured_qty:,.2f}",
                "Rate (MMK)": f"{ls_rate:,.2f}",
                "Per": item['unit'],
                "Amount (MMK)": f"{amount:,.2f}"
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
                    "Item": "", "Particular": "  📦 Material", "Unit": "", "Quantity": "", "Rate (MMK)": "", "Per": "", "Amount (MMK)": ""
                })
                for m in mat_breakdown:
                    display_rows.append({
                        "Item": "",
                        "Particular": f"      {m['part']}",
                        "Unit": m['unit'],
                        "Quantity": f"{m['qty']:,.2f}",
                        "Rate (MMK)": f"{m['rate']:,.2f}" if m['rate'] > 0 else "-",
                        "Per": m['unit'],
                        "Amount (MMK)": f"{m['amount']:,.2f}" if m['amount'] > 0 else "-"
                    })

            if lab_breakdown:
                display_rows.append({
                    "Item": "", "Particular": "  👷 Labour", "Unit": "", "Quantity": "", "Rate (MMK)": "", "Per": "", "Amount (MMK)": ""
                })
                for l in lab_breakdown:
                    display_rows.append({
                        "Item": "",
                        "Particular": f"      {l['part']}",
                        "Unit": l['unit'],
                        "Quantity": f"{l['qty']:,.2f}",
                        "Rate (MMK)": f"{l['rate']:,.2f}",
                        "Per": l['unit'],
                        "Amount (MMK)": f"{l['amount']:,.2f}"
                    })

        display_rows.append({
            "Item": "",
            "Particular": "  💰 Total Cost",
            "Unit": "",
            "Quantity": "",
            "Rate (MMK)": "",
            "Per": "",
            "Amount (MMK)": f"**{item_total_cost:,.2f}**"
        })

        st.dataframe(pd.DataFrame(display_rows), use_container_width=True, hide_index=True)
        grand_total += item_total_cost

    st.divider()

    # ==========================================
    # 4. Detailed Bill of Quantity (BOQ Breakdown)
    # ==========================================
    st.subheader("📜 ၄။ Detailed Bill Of Quantity (BOQ Summary)")

    # Material Section
    st.markdown("### 📦 Material Cost Breakdown")
    mat_rows = []
    total_mat_cost = 0.0
    for idx, (p_name, data) in enumerate(material_summary.items(), start=1):
        mat_rows.append({
            "No.": idx,
            "Particular": p_name,
            "Unit": data["unit"],
            "Quantity": f"{data['qty']:,.2f}",
            "Rate (MMK)": f"{data['rate']:,.2f}",
            "Amount (MMK)": f"{data['amount']:,.2f}"
        })
        total_mat_cost += data["amount"]

    if mat_rows:
        st.table(pd.DataFrame(mat_rows))
    else:
        st.info("Material စရိတ် မရှိပါ။")
    st.markdown(f"**Total Material Cost = `{total_mat_cost:,.2f} MMK`**")

    st.divider()

    # Labour Section
    st.markdown("### 👷 Labour Cost Breakdown")
    lab_rows = []
    total_lab_cost = 0.0
    for idx, (p_name, data) in enumerate(labour_summary.items(), start=1):
        lab_rows.append({
            "No.": idx,
            "Particular": p_name,
            "Unit": data["unit"],
            "Quantity": f"{data['qty']:,.2f}",
            "Rate (MMK)": f"{data['rate']:,.2f}",
            "Amount (MMK)": f"{data['amount']:,.2f}"
        })
        total_lab_cost += data["amount"]

    if lab_rows:
        st.table(pd.DataFrame(lab_rows))
    else:
        st.info("Labour စရိတ် မရှိပါ။")
    st.markdown(f"**Total Labour Cost = `{total_lab_cost:,.2f} MMK`**")

    # 📥 Download BOQ Summary Excel Button (With Formulas)
    boq_excel_buffer = export_boq_summary_excel(material_summary, labour_summary)
    st.download_button(
        label="📥 Download BOQ Material & Labour Summary (Excel With Formulas)",
        data=boq_excel_buffer,
        file_name="BOQ_Summary_Formulas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.divider()

    # Grand Totals Summary Box
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("📦 Total Material Cost", f"{total_mat_cost:,.2f} MMK")
    col_m2.metric("👷 Total Labour Cost", f"{total_lab_cost:,.2f} MMK")
    col_m3.metric("💰 Grand Total Cost", f"{grand_total:,.2f} MMK")


if __name__ == "__main__":
    main()
