import os
import io
import math
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
    "smith", "machine driver", "labour",
    "hoisting and fixing", "carriage to site", "site clearing", "dressing"
]

# Item ခေါင်းစဉ်များအတွက် ဘာသာပြန် Dictionary (Item 22 ဖြုတ်ထားပါသည်)
MM_ITEM_TITLES = {
    "1": "မြေပြင်ရှင်းလင်းခြင်းနှင့် သစ်ငုတ်တူးခြင်း",
    "2": "အနက် ၅ ပေ မပိုသော သာမန်မြေကျင်း တူးဖော်ခြင်း",
    "3": "အနက် ၅ ပေ မပိုသော မြေမာ/ကျောက်စရစ်ပါသော မြေကျင်း တူးဖော်ခြင်း",
    "4": "အနက် ၅ ပေ မပိုသော ကျောက်ခဲ/ကျောက်မာ မြေကျင်း တူးဖော်ခြင်း",
    "5": "တူးဖော်ပြီး မြေကျင်းများအတွင်း မြေပြန်ဖို့ခြင်းနှင့် သိုသိပ်ခြင်း",
    "6": "သဲ ဖို့ခြင်းနှင့် ရေလောင်းသိုသိပ်ခြင်း",
    "7": "သဲနှင့် မြေနီစပ် ဖို့ခြင်းနှင့် သိုသိပ်ခြင်း",
    "8": "အဆောက်အအုံပတ်လည် မြေဖို့ခြင်းနှင့် ကျစ်လျစ်အောင် ပြုလုပ်ခြင်း",
    "9": "မြေကျင်းဘေးနံရံ ကမ်းပါးပြိုကျမှု ကာကွယ်ရန် သစ်သားကာရံခြင်း (Timber Shoring)",
    "12": "ကျောက်စရစ်/မြေကျင်း ရေထုတ်ယူခြင်း (Dewatering Work)",
    "13": "၂ ပေပတ်လည်၊ ၂ ပေအနက် သာမန်မြေကျင်းတူးပြီးမြေဖို့ခြင်း",
    "14": "၂ ပေပတ်လည်၊ ၃ ပေအနက် ကျောက်ခဲ/ကျောက်မာ မြေကျင်းတူးပြီးမြေဖို့ခြင်း",
    "15": "၁၂ ကုဗပေ ထက်မပိုသော မြေမာ/ကျောက်စရစ်ပါသော မြေကျင်းတူးပြီးမြေဖို့ခြင်း",
    "16": "၁၀ ကုဗပေ ထက်မပိုသော သာမန်မြေကျင်းတူးပြီးမြေဖို့ခြင်း",
    "17": "၁၀ ကုဗပေ ထက်မပိုသော ကျောက်ခဲ/ကျောက်မာ မြေကျင်းတူးပြီးမြေဖို့ခြင်း",
    "18": "အောက်ခြေ ၉ လက်မ၊ အပေါ်ထိပ် ၁  ပေခွဲ၊ အနက် ၁ ပေ ရှိသော ရေမြောင်းတူးဖော်ခြင်း (သာမန်မြေ)",
    "19": "အောက်ခြေ ၉ လက်မ၊ အပေါ်ထိပ် ၁  ပေခွဲ၊ အနက် ၁ ပေ ရှိသော ရေမြောင်းတူးဖော်ခြင်း ( မြေမာ/ကျောက်စရစ်ပါသော မြေ)",
    "20": "၁၀ ပေ အနက် မိလ္လာကျင်းတူးခြင်း  (မြေမာ/ကျောက်စရစ်ပါသော မြေ)",
    "21": "၁၀ ပေ အနက် မိလ္လာကျင်းတူးခြင်း  (သဲမြေ၊ ရွှံ့မြေ)",
    "23": "ဖောင်ဒေးရှင်းချရန် ငုတ်ရိုက်ခြင်း",
}

# Material / Labour Description ဘာသာပြန် Dictionary
PARTICULAR_TRANSLATIONS = {
    "MM": {
        "Worker": "အလုပ်သမား",
        "Worker for carrying and ramming": "မြေသယ်/မြေရိုက် အလုပ်သမား",
        "Worker for watering": "ရေလောင်း အလုပ်သမား",
        "Worker for carrying": "သယ်ယူပို့ဆောင်ရေး အလုပ်သမား",
        "Digger": "မြေကျင်းတူး အလုပ်သမား",
        "Maistry": "ခေါင်းဆောင် / မေစတရီ",
        "Surveyor": "တိုင်းတာရေးမှူး / Surveyor",
        "Carpenter": "လက်သမား",
        "Timber": "သစ်သား",
        "Wire Nails": "သံရိုက် / သံမှို",
        "Water Charges": "ရေဖိုးရေခ",
        "Sand": "သဲ",
        "Carriage to site": "လုပ်ငန်းခွင်သို့ သယ်ယူခ"
    },
    "EN": {
        "Worker": "Worker",
        "Worker for carrying and ramming": "Worker for carrying and ramming",
        "Worker for watering": "Worker for watering",
        "Worker for carrying": "Worker for carrying",
        "Digger": "Digger",
        "Maistry": "Maistry / Supervisor",
        "Surveyor": "Surveyor",
        "Carpenter": "Carpenter",
        "Timber": "Timber",
        "Wire Nails": "Wire Nails",
        "Water Charges": "Water Charges",
        "Sand": "Sand",
        "Carriage to site": "Carriage to site"
    }
}

# Unit ဘာသာပြန် Dictionary
UNIT_TRANSLATIONS = {
    "MM": {
        "100 Cft": "၁၀၀ ကုဗပေ",
        "Cft": "ကုဗပေ",
        "Sft": "စတုရန်းပေ",
        "Rft": "ပေအရှည်",
        "Ton": "တန်",
        "Viss": "ပိဿာ",
        "L-S": "တစ်စုတစ်ဝေး",
        "LS": "တစ်စုတစ်ဝေး",
        "Job": "လုပ်ငန်း",
        "Hole": "တွင်း",
        "Sud": "ကျင်း"
    },
    "EN": {
        "100 Cft": "100 Cft",
        "Cft": "Cft",
        "Sft": "Sft",
        "Rft": "Rft",
        "Ton": "Ton",
        "Viss": "Viss",
        "L-S": "L-S",
        "LS": "L-S",
        "Job": "Job",
        "Hole": "Hole",
        "Sud": "Sud (%Cft)"
    }
}

def translate_unit(unit_str, lang):
    if not unit_str:
        return ""
    u_clean = str(unit_str).strip()
    return UNIT_TRANSLATIONS.get(lang, {}).get(u_clean, u_clean)

def translate_particular(part_str, lang):
    if not part_str:
        return ""
    p_clean = str(part_str).strip()
    return PARTICULAR_TRANSLATIONS.get(lang, {}).get(p_clean, p_clean)

# Multi-language dictionary
TRANSLATIONS = {
    "MM": {
        "title": "🚜 Earthwork QS & Calculation Tool",
        "subtitle": "မြေကျင်းတူး/မြေဖို့ လုပ်ငန်းများအတွက် အတိုင်းအတာများ ရိုက်ထည့်၍ ကုန်ကျစရိတ်နှင့် လုပ်အားခ/ပစ္စည်း BOQ စာရင်း တွက်ချက်ပါ",
        "tools_title": "🛠️ အရန်ကိရိယာများနှင့် ပေါက်ဈေး ပြင်ဆင်ရန် (Tools & Rates)",
        "tab_rates": "⚙ ပစ္စည်း/လုပ်အားခ ပေါက်ဈေး",
        "tab_calc": "🧮 ဂဏန်းတွက်စက်",
        "tab_conv": "🔄 ယူနစ်ပြောင်းရန်",
        "tab_upload": "📥 Excel ဖိုင်တင်ရန်",
        "rates_subheader": "မြေကျင်းလုပ်ငန်း ပေါက်ဈေး သတ်မှတ်ရန် (ကျပ်)",
        "labour_rates": "**👷 လုပ်အားခ ပေါက်ဈေးများ**",
        "material_rates": "**📦 ပစ္စည်းနှင့် အခြား ကုန်ကျစရိတ်များ**",
        "rate_worker": "အလုပ်သမား (ကျပ်)",
        "rate_digger": "မြေကျင်းတူး (ကျပ်)",
        "rate_maistry": "ခေါင်းဆောင် / မေစတရီ (ကျပ်)",
        "rate_surveyor": "တိုင်းတာရေးမှူး / Surveyor (ကျပ်)",
        "rate_carpenter": "လက်သမား / Carpenter (ကျပ်)",
        "rate_timber": "သစ် / Timber (၁ တန် ကျပ်)",
        "rate_nails": "သံရိုက် / Wire Nails (၁ ပိဿာ ကျပ်)",
        "rate_water": "ရေဖိုးရေခ / Water Charges (L-s ကျပ်)",
        "rate_sand": "သဲဖို့ (ကျင်း)",
        "rate_carriage": "မြေ/သဲ သယ်ယူခ (ကျင်း)",
        "calc_subheader": "🧮 အလွယ်တွက်စက်",
        "calc_input": "တွက်လိုသည်များကို ရိုက်ထည့်ပါ (ဥပမာ- 10*12.5 + 5):",
        "calc_ans": "အဖြေ",
        "calc_err_char": "ဂဏန်းနှင့် သင်္ကေတများသာ ရိုက်ထည့်ပါ။",
        "calc_err_struct": "တွက်ချက်မှု အမှားရှိနေပါသည် structure ကို ပြန်စစ်ပါ။",
        "conv_subheader": "🔄 ယူနစ် အပြောင်းအလဲ",
        "conv_select": "ပြောင်းလိုသည်ကို ရွေးပါ:",
        "conv_inch_ft": "Inches -> Feet (လက်မ -> ပေ)",
        "conv_sft_sqm": "Sft <-> Sq.m (စတုရန်းပေ <-> စတုရန်းမီတာ)",
        "conv_cft_cum": "Cft <-> Cu.m (ကုဗပေ <-> ကုဗမီတာ)",
        "conv_cft_sud": "Cft -> Sud / %Cft (ကုဗပေ -> ကျင်း)",
        "conv_inch_label": "Inches / လက်မ:",
        "conv_sft_label": "Sft / စတုရန်းပေ:",
        "conv_cft_label": "Cft / ကုဗပေ:",
        "conv_sud_label": "Cft / ကုဗပေ ပမာဏ:",
        "upload_subheader": "📥 တိုင်းတာပြီး Earthwork Excel ဖိုင်တင်ရန်",
        "download_template": "📄 Earthwork နမူနာ ပုံစံ (Template) ရယူရန်",
        "upload_file_label": "Excel / CSV ဖိုင် ရွေးပါ:",
        "btn_parse_excel": "🚀 ဖိုင်ထဲမှ စာရင်းများ ဖတ်ယူမည်",
        "sec1_title": "၁။ တွက်ချက်လိုသော Earthwork Item များ ရွေးပါ",
        "select_items": "🚜 မြေကျင်းလုပ်ငန်းမှ တွက်လိုသည့် Item များကို ရွေးပါ:",
        "no_excel_err": "⚠️ '1 Earth Work.xls' ဖိုင်ကို ရှာမတွေ့ပါ သို့မဟုတ် ဖိုင်ထဲတွင် ဒေတာ မရှိပါ။",
        "item_select_hint": "💡 **အကြံပြုချက်**: တွက်ချက်လိုသော Earthwork Item များကို အထက်ပါ Multiselect Box တွင် ရွေးပေးပါ။",
        "sec2_title": "📐 ၂။ အတိုင်းအတာများ ရိုက်ထည့်ပါ (Detail Measurement)",
        "work_location": "လုပ်ငန်းနေရာ / အမျိုးအစား",
        "location_grid": "နေရာ / အကွက်အမည်",
        "grid_default": "အကွက်",
        "qty_count": "No.",
        "hole_count": "တွင်းအရေအတွက် (Holes Count)",
        "lumpsum_rate": "တစ်စုတစ်ဝေးတည်း ဈေးနှုန်း (ကျပ်)",
        "len_ft": "အရှည် L (ပေ)",
        "wid_ft": "အနံ B (ပေ)",
        "hei_ft": "အမြင့်/အထူ H (ပေ)",
        "deduction": "အနှုတ်ကျင်း (Deduction)",
        "is_deduction_row": "➖ အနှုတ်လိုင်း ဖြစ်သည်",
        "btn_copy": "📋 ပွားမည် (Copy)",
        "btn_add_row": "➕ အကွက်အသစ်ထည့်ရန်",
        "btn_rem_row": "➖ အကွက်ပြန်ဖြုတ်ရန်",
        "total_summary": "📊 တိုင်းတာချက် စာရင်းချုပ်",
        "dl_meas_excel": "📥 Detail Earthwork Measurement Sheet ကို Excel ဖြင့် ဒေါင်းလုဒ်ရယူရန်",
        "sec3_title": "📊 ၃။ Earthwork ကုန်ကျစရိတ် တွက်ချက်မှု (Rate Analysis)",
        "particular": "အကြောင်းအရာ",
        "unit": "ယူနစ်",
        "quantity": "ပမာဏ",
        "rate_mmk": "နှုန်းထား (ကျပ်)",
        "amount_mmk": "ကျသင့်ငွေ (ကျပ်)",
        "mat_cost_title": "  📦 ပစ္စည်းစရိတ် (Material)",
        "lab_cost_title": "  👷 လုပ်အားခ (Labour)",
        "total_item_cost": "  💰 စုစုပေါင်း ကုန်ကျစရိတ်",
        "sec4_title": "📜 ၄။ Earthwork BOQ စာရင်းချုပ်",
        "mat_boq_title": "📦 ၁။ ပစ္စည်းကုန်ကျစရိတ် စာရင်း (Material Summary)",
        "lab_boq_title": "👷 ၂။ လုပ်အားခ စာရင်း (Labour Summary)",
        "sr_no": "စဉ်",
        "mat_name": "ပစ္စည်းအမည်",
        "req_qty": "လိုအပ်သော ပမာဏ",
        "total_mmk": "စုစုပေါင်း (ကျပ်)",
        "no_mat_cost": "ပစ္စည်းစရိတ် မရှိပါ။",
        "no_lab_cost": "လုပ်အားခ စရိတ် မရှိပါ။",
        "dl_boq_excel": "📥 Earthwork BOQ စာရင်းချုပ်ကို Excel ဖြင့် ရယူရန်",
        "total_mat_val": "📦 စုစုပေါင်း ပစ္စည်းဖိုး",
        "total_lab_val": "👷 စုစုပေါင်း လုပ်အားခ",
        "grand_total_val": "💰 မြေကျင်းလုပ်ငန်း စုစုပေါင်းစရိတ်",
        "type_add": "➕ အပေါင်း",
        "type_ded": "➖ အနှုတ်",
        "item_no_col": "Item No"
    },
    "EN": {
        "title": "🚜 Earthwork QS & Calculation Tool",
        "subtitle": "Calculate quantities, material/labour costs, and BOQ summaries for earthwork excavation and backfilling.",
        "tools_title": "🛠️ Tools & Unit Rates Configuration",
        "tab_rates": "⚙ Material/Labour Rates",
        "tab_calc": "🧮 Calculator",
        "tab_conv": "🔄 Unit Converter",
        "tab_upload": "📥 Import Excel",
        "rates_subheader": "Set Earthwork Unit Rates (MMK)",
        "labour_rates": "**👷 Labour Rates**",
        "material_rates": "**📦 Material & Other Costs**",
        "rate_worker": "Worker (MMK)",
        "rate_digger": "Digger (MMK)",
        "rate_maistry": "Maistry / Supervisor (MMK)",
        "rate_surveyor": "Surveyor (MMK)",
        "rate_carpenter": "Carpenter (MMK)",
        "rate_timber": "Timber (per Ton MMK)",
        "rate_nails": "Wire Nails (per Viss MMK)",
        "rate_water": "Water Charges (L-s MMK)",
        "rate_sand": "Sand Filling (per Sud MMK)",
        "rate_carriage": "Earth/Sand Carriage (per Sud MMK)",
        "calc_subheader": "🧮 Quick Calculator",
        "calc_input": "Enter expression (e.g. 10*12.5 + 5):",
        "calc_ans": "Result",
        "calc_err_char": "Please enter numbers and mathematical symbols only.",
        "calc_err_struct": "Calculation error. Please check the expression structure.",
        "conv_subheader": "🔄 Unit Converter",
        "conv_select": "Select conversion type:",
        "conv_inch_ft": "Inches -> Feet",
        "conv_sft_sqm": "Sft <-> Sq.m",
        "conv_cft_cum": "Cft <-> Cu.m",
        "conv_cft_sud": "Cft -> Sud / %Cft",
        "conv_inch_label": "Inches:",
        "conv_sft_label": "Sft:",
        "conv_cft_label": "Cft:",
        "conv_sud_label": "Cft Quantity:",
        "upload_subheader": "📥 Import Measured Earthwork Excel File",
        "download_template": "📄 Download Earthwork Template",
        "upload_file_label": "Select Excel / CSV File:",
        "btn_parse_excel": "🚀 Import Data from File",
        "sec1_title": "1. Select Earthwork Items",
        "select_items": "🚜 Choose items to calculate from Earthwork catalog:",
        "no_excel_err": "⚠️ '1 Earth Work.xls' file not found or contains no data.",
        "item_select_hint": "💡 **Tip**: Select items from the multiselect box above to begin.",
        "sec2_title": "📐 2. Detail Measurement Input",
        "work_location": "Work Location / Type",
        "location_grid": "Location / Grid Name",
        "grid_default": "Grid",
        "qty_count": "No.",
        "hole_count": "Holes Count",
        "lumpsum_rate": "Lumpsum Rate (MMK)",
        "len_ft": "Length L (ft)",
        "wid_ft": "Breadth B (ft)",
        "hei_ft": "Height/Thickness H (ft)",
        "deduction": "Deduction",
        "is_deduction_row": "➖ Deduction Row",
        "btn_copy": "📋 Copy",
        "btn_add_row": "➕ Add Row",
        "btn_rem_row": "➖ Remove Row",
        "total_summary": "📊 Measurement Breakdown Summary",
        "dl_meas_excel": "📥 Download Detail Measurement Sheet (Excel)",
        "sec3_title": "📊 3. Rate Analysis & Cost Calculation",
        "particular": "Description",
        "unit": "Unit",
        "quantity": "Quantity",
        "rate_mmk": "Rate (MMK)",
        "amount_mmk": "Amount (MMK)",
        "mat_cost_title": "  📦 Material Cost",
        "lab_cost_title": "  👷 Labour Cost",
        "total_item_cost": "  💰 Total Item Cost",
        "sec4_title": "📜 4. Earthwork BOQ Summary",
        "mat_boq_title": "📦 1. Material Cost Summary",
        "lab_boq_title": "👷 2. Labour Cost Summary",
        "sr_no": "No.",
        "mat_name": "Description",
        "req_qty": "Required Qty",
        "total_mmk": "Total (MMK)",
        "no_mat_cost": "No material costs.",
        "no_lab_cost": "No labour costs.",
        "dl_boq_excel": "📥 Download Earthwork BOQ Summary (Excel)",
        "total_mat_val": "📦 Total Material Cost",
        "total_lab_val": "👷 Total Labour Cost",
        "grand_total_val": "💰 Total Earthwork Cost",
        "type_add": "➕ Addition",
        "type_ded": "➖ Deduction",
        "item_no_col": "Item No"
    }
}

@st.cache_data
def parse_excel_rates(file_path):
    if not os.path.exists(file_path):
        return []
    try:
        df_raw = pd.read_excel(file_path, dtype=str)
    except Exception as e:
        st.error(f"Error reading file ({file_path}): {e}")
        return []

    items = []
    current_item = None

    for idx, row in df_raw.iterrows():
        item_no = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        particular = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
        unit = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ''
        qty = row.iloc[3] if pd.notna(row.iloc[3]) else '0'

        # Item 22 အပါအဝင် မလိုလိုသော Item များကို ကျော်မည်
        if item_no in ['10', '10.0', '11', '11.0', '22', '22.0']:
            continue

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
            st.session_state['last_excel_error'] = "❌ Excel file missing 'Item No.' or 'Particular Description' column."
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
            st.session_state['last_excel_error'] = "⚠ No valid measurement data found in Excel file."
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

        st.session_state['excel_import_success'] = f"✅ Successfully imported {imported_rows_count} measurement items from Excel."

    except Exception as e:
        err_msg = traceback.format_exc()
        st.session_state['last_excel_error'] = f"❌ Import Error: {e}\n\n{err_msg}"


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
        ["2", "Excavation Grid A-1", 2, 10, 5, 8, 0, "Addition"],
        ["3", "Foundation Hole Grid B", 1, 20, 150, 6, 0, "Addition"],
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


def export_measurement_sheet_excel(selected_items_list, st_session_state, lang):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Earthwork Measurement"

    ws['A1'] = "EARTHWORK DETAIL MEASUREMENT SHEET"
    ws['A1'].font = Font(name='Calibri', size=14, bold=True, color='1F497D')
    
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
        is_hole = 'hole' in unit_str
        is_sft = 'sft' in unit_str or 'sq.ft' in unit_str or 'sqft' in unit_str
        is_rft = 'rft' in unit_str or 'lin.ft' in unit_str

        count_header = "Holes Count" if is_hole else "No."
        headers = ["Item No.", "Particular Description", count_header, "L (ft)", "B (ft)", "H (ft)", "Deduction", "Type", "Sub-total"]

        display_unit = translate_unit(item['unit'], lang)
        ws.cell(row=row_idx, column=1, value=f"Item {item_no_str} - {item['display_title']} ({display_unit})").font = Font(bold=True, size=11)
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
                ws.cell(row=row_idx, column=4, value=r['l'] if not is_hole else "-")
                ws.cell(row=row_idx, column=5, value=r['b'] if (not is_rft and not is_hole) else "-")
                ws.cell(row=row_idx, column=6, value=r['h'] if (not is_sft and not is_rft and not is_hole) else "-")
                ws.cell(row=row_idx, column=7, value=r['ded'] if not is_hole else 0)
                ws.cell(row=row_idx, column=8, value="Deduction" if r.get("is_deduction_row") else "Addition")

                if is_hole:
                    mult_expr = f"C{row_idx}"
                elif is_rft:
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

        ws.cell(row=row_idx, column=2, value=f"Total Quantity ({display_unit})").font = Font(bold=True)
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

    col_title_space, col_lang = st.columns([4, 1])
    with col_lang:
        lang_choice = st.radio(
            "🌐 Language / ဘာသာစကား",
            options=["မြန်မာ", "English"],
            horizontal=True,
            key="app_language"
        )
    
    lang = "MM" if lang_choice == "မြန်မာ" else "EN"
    t = TRANSLATIONS[lang]

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

    st.markdown(f"""
        <div class="main-header">
            <h1>{t['title']}</h1>
            <p>{t['subtitle']}</p>
        </div>
    """, unsafe_allow_html=True)

    earthwork_path = os.path.join(BASE_DIR, "1 Earth Work.xls")
    earthwork_items = parse_excel_rates(earthwork_path)

    # ဘာသာစကားပေါ်မူတည်၍ Title နှင့် Unit များကို Dynamic ပြောင်းလဲခြင်း
    prefix = "[မြေကျင်း]" if lang == "MM" else "[Earthwork]"
    ew_options = {}
    for item in earthwork_items:
        item_no_clean = str(item['item_no']).replace('.0', '').strip()
        if lang == "MM" and item_no_clean in MM_ITEM_TITLES:
            display_title = MM_ITEM_TITLES[item_no_clean]
        else:
            display_title = item['title']
        
        display_unit = translate_unit(item['unit'], lang)

        item_copy = dict(item)
        item_copy['display_title'] = display_title
        item_copy['display_unit'] = display_unit
        key_str = f"{prefix} Item {item['item_no']} - {display_title}"
        ew_options[key_str] = item_copy

    if 'selected_ew' not in st.session_state:
        st.session_state['selected_ew'] = []

    if st.session_state.get('last_excel_error'):
        st.error(st.session_state['last_excel_error'])

    if st.session_state.get('excel_import_success'):
        st.success(st.session_state['excel_import_success'])

    with st.expander(t['tools_title'], expanded=True):
        tab_rates, tab_calc, tab_conv, tab_upload = st.tabs([
            t['tab_rates'], 
            t['tab_calc'], 
            t['tab_conv'],
            t['tab_upload']
        ])

        with tab_rates:
            st.subheader(t['rates_subheader'])
            col_r1, col_r2 = st.columns(2)
            
            with col_r1:
                st.markdown(t['labour_rates'])
                rate_worker = st.number_input(t['rate_worker'], value=25000.0, step=1000.0)
                rate_digger = st.number_input(t['rate_digger'], value=25000.0, step=1000.0)
                rate_maistry = st.number_input(t['rate_maistry'], value=30000.0, step=1000.0)
                rate_surveyor = st.number_input(t['rate_surveyor'], value=40000.0, step=1000.0)
                rate_carpenter = st.number_input(t['rate_carpenter'], value=35000.0, step=1000.0)

            with col_r2:
                st.markdown(t['material_rates'])
                rate_timber = st.number_input(t['rate_timber'], value=1800000.0, step=50000.0)
                rate_nails = st.number_input(t['rate_nails'], value=12000.0, step=500.0)
                rate_water = st.number_input(t['rate_water'], value=20000.0, step=1000.0)
                rate_sand = st.number_input(t['rate_sand'], value=45000.0, step=1000.0)
                rate_carriage = st.number_input(t['rate_carriage'], value=15000.0, step=1000.0)

        with tab_calc:
            st.subheader(t['calc_subheader'])
            calc_expr = st.text_input(t['calc_input'], value="")
            if calc_expr:
                try:
                    allowed_chars = "0123456789+-*/(). "
                    if all(char in allowed_chars for char in calc_expr):
                        res = eval(calc_expr)
                        st.success(f"**{t['calc_ans']} = {res:,.4f}**")
                    else:
                        st.error(t['calc_err_char'])
                except Exception:
                    st.error(t['calc_err_struct'])

        with tab_conv:
            st.subheader(t['conv_subheader'])
            conv_type = st.selectbox(t['conv_select'], [
                t['conv_inch_ft'],
                t['conv_sft_sqm'],
                t['conv_cft_cum'],
                t['conv_cft_sud']
            ])

            if conv_type == t['conv_inch_ft']:
                inch_val = st.number_input(t['conv_inch_label'], min_value=0.0, value=6.0)
                st.info(f"👉 **{inch_val} inches = {inch_val / 12.0:.3f} ft**")

            elif conv_type == t['conv_sft_sqm']:
                sft_val = st.number_input(t['conv_sft_label'], min_value=0.0, value=100.0)
                st.info(f"👉 **{sft_val:,.2f} Sft = {sft_val / 10.764:.2f} Sq.m**")

            elif conv_type == t['conv_cft_cum']:
                cft_val = st.number_input(t['conv_cft_label'], min_value=0.0, value=100.0)
                st.info(f"👉 **{cft_val:,.2f} Cft = {cft_val / 35.315:.2f} Cu.m**")

            elif conv_type == t['conv_cft_sud']:
                cft_val = st.number_input(t['conv_sud_label'], min_value=0.0, value=500.0)
                st.info(f"👉 **{cft_val:,.2f} Cft = {cft_val / 100.0:.2f} Sud (%Cft)**")

        with tab_upload:
            st.subheader(t['upload_subheader'])
            template_buffer = export_measurement_template()
            st.download_button(
                label=t['download_template'],
                data=template_buffer,
                file_name="Earthwork_Measurement_Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.divider()
            uploaded_meas_file = st.file_uploader(t['upload_file_label'], type=["xlsx", "xls", "csv"])
            if uploaded_meas_file is not None:
                if st.button(t['btn_parse_excel'], type="primary", use_container_width=True):
                    parse_and_auto_select_uploaded_excel(uploaded_meas_file, ew_options)
                    st.rerun()

    rate_worker = locals().get('rate_worker', 25000.0)
    rate_digger = locals().get('rate_digger', 25000.0)
    rate_maistry = locals().get('rate_maistry', 30000.0)
    rate_surveyor = locals().get('rate_surveyor', 40000.0)
    rate_carpenter = locals().get('rate_carpenter', 35000.0)

    rate_timber = locals().get('rate_timber', 1800000.0)
    rate_nails = locals().get('rate_nails', 12000.0)
    rate_water = locals().get('rate_water', 20000.0)
    rate_sand = locals().get('rate_sand', 45000.0)
    rate_carriage = locals().get('rate_carriage', 15000.0)

    rate_map = {
        "Worker": rate_worker,
        "Worker for carrying and ramming": rate_worker,
        "Worker for watering": rate_worker,
        "Worker for carrying": rate_worker,
        "Digger": rate_digger,
        "Maistry": rate_maistry,
        "Surveyor": rate_surveyor,
        "Carpenter": rate_carpenter,
        "Timber": rate_timber,
        "Wire Nails": rate_nails,
        "Water Charges": rate_water,
        "Sand": rate_sand,
        "Carriage to site": rate_carriage,
    }

    # Step 1: Selection
    st.markdown(f'<div class="section-title">{t["sec1_title"]}</div>', unsafe_allow_html=True)
    
    if not earthwork_items:
        st.warning(t['no_excel_err'])
        return

    ew_selected = st.multiselect(t['select_items'], list(ew_options.keys()), key="selected_ew")
    selected_items_list = [ew_options[k] for k in ew_selected if k in ew_options]

    if not selected_items_list:
        st.info(t['item_select_hint'])
        return

    st.divider()

    # Step 2: Detail Measurement
    st.markdown(f'<div class="section-title">{t["sec2_title"]}</div>', unsafe_allow_html=True)

    item_quantities = {}
    ls_custom_rates = {}

    for idx, item in enumerate(selected_items_list):
        item_no_str = str(item['item_no'])
        rows_state_key = f"rows_data_{item_no_str}_{idx}"

        unit_str = str(item['unit']).lower().strip()
        is_lumpsum = 'l-s' in unit_str or 'ls' in unit_str or 'lump' in unit_str or 'job' in unit_str
        is_hole = 'hole' in unit_str
        is_sft = 'sft' in unit_str or 'sq.ft' in unit_str or 'sqft' in unit_str
        is_rft = 'rft' in unit_str or 'lin.ft' in unit_str

        label_no = t['hole_count'] if is_hole else t['qty_count']
        display_unit = item['display_unit']

        if rows_state_key not in st.session_state:
            st.session_state[rows_state_key] = [
                {
                    "desc": f"{t['grid_default']} 1",
                    "no": 1,
                    "l": 50.0 if is_sft else (10.0 if not is_hole else 0.0),
                    "b": 50.0 if is_sft else (10.0 if not is_hole else 0.0),
                    "h": 5.0 if not is_hole else 0.0,
                    "ded": 0.0,
                    "is_deduction_row": False
                }
            ]

        with st.expander(f"📌 Item {item['item_no']} - {item['display_title']} [{display_unit}]", expanded=True):
            meas_rows = []
            item_total_qty = 0.0

            if is_lumpsum:
                c_desc, c_no, c_rate = st.columns([3, 1, 2])
                p_desc = c_desc.text_input(t['work_location'], value="Lumpsum Job", key=f"desc_{idx}_{item_no_str}")
                no_val = c_no.number_input(label_no, min_value=1, value=1, key=f"no_{idx}_{item_no_str}")
                ls_rate = c_rate.number_input(t['lumpsum_rate'], min_value=0.0, value=50000.0, step=10000.0, key=f"ls_rate_{idx}_{item_no_str}")
                
                item_total_qty = float(no_val)
                ls_custom_rates[item_no_str] = ls_rate

                meas_rows.append({
                    t['work_location']: p_desc,
                    label_no: no_val,
                    t['len_ft']: "-",
                    t['wid_ft']: "-",
                    t['hei_ft']: "-",
                    "Type": t['type_add'],
                    "Result": no_val
                })
            else:
                current_rows = st.session_state[rows_state_key]
                row_to_copy = None

                for r_idx, r_data in enumerate(current_rows):
                    st.markdown(f"**🔹 Row ({r_idx+1})**")
                    
                    if is_hole:
                        c_desc, c_no, c_is_ded, c_cp = st.columns([4, 2, 2, 1])
                    elif is_sft:
                        c_desc, c_no, c_l, c_b = st.columns([2.5, 1, 1, 1])
                        c_ded, c_is_ded, c_cp = st.columns([1.5, 1.5, 1])
                    elif is_rft:
                        c_desc, c_no, c_l = st.columns([3, 1, 1])
                        c_ded, c_is_ded, c_cp = st.columns([1.5, 1.5, 1])
                    else:
                        c_desc, c_no, c_l, c_b, c_h = st.columns([2.5, 1, 1, 1, 1])
                        c_ded, c_is_ded, c_cp = st.columns([1.5, 1.5, 1])

                    p_desc = c_desc.text_input(
                        t['location_grid'],
                        value=r_data["desc"],
                        key=f"desc_{idx}_{r_idx}_{item_no_str}"
                    )
                    no_val = c_no.number_input(label_no, min_value=1, value=int(r_data["no"]), key=f"no_{idx}_{r_idx}_{item_no_str}")
                    
                    l_val = 0.0
                    b_val = 0.0
                    h_val = 0.0
                    ded_val = 0.0

                    if not is_hole:
                        l_val = c_l.number_input(t['len_ft'], min_value=0.0, value=float(r_data["l"]), key=f"l_{idx}_{r_idx}_{item_no_str}")
                        if not is_rft:
                            b_val = c_b.number_input(t['wid_ft'], min_value=0.0, value=float(r_data["b"]), key=f"b_{idx}_{r_idx}_{item_no_str}")
                        if not is_sft and not is_rft:
                            h_val = c_h.number_input(t['hei_ft'], min_value=0.0, value=float(r_data["h"]), key=f"h_{idx}_{r_idx}_{item_no_str}")
                        ded_val = c_ded.number_input(t['deduction'], min_value=0.0, value=float(r_data.get("ded", 0.0)), key=f"ded_{idx}_{r_idx}_{item_no_str}")

                    is_ded_row = c_is_ded.checkbox(t['is_deduction_row'], value=r_data.get("is_deduction_row", False), key=f"is_ded_{idx}_{r_idx}_{item_no_str}")

                    r_data["desc"] = p_desc
                    r_data["no"] = no_val
                    r_data["l"] = l_val
                    r_data["b"] = b_val
                    r_data["h"] = h_val
                    r_data["ded"] = ded_val
                    r_data["is_deduction_row"] = is_ded_row

                    if c_cp.button(t['btn_copy'], key=f"copy_{idx}_{r_idx}_{item_no_str}"):
                        row_to_copy = dict(r_data)

                    if is_hole:
                        gross_qty = float(no_val)
                    elif is_rft:
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
                        t['location_grid']: p_desc,
                        label_no: no_val,
                        t['len_ft']: l_val if not is_hole else "-",
                        t['wid_ft']: b_val if (not is_rft and not is_hole) else "-",
                        t['hei_ft']: h_val if (not is_sft and not is_rft and not is_hole) else "-",
                        t['deduction']: ded_val if not is_hole else "-",
                        "Type": t['type_ded'] if is_ded_row else t['type_add'],
                        "Result": sub_total_display
                    })
                    st.markdown("---")

                if row_to_copy is not None:
                    copied_row = dict(row_to_copy)
                    copied_row["desc"] = f"{copied_row['desc']} (Copy)"
                    st.session_state[rows_state_key].append(copied_row)
                    st.rerun()

                col_add, col_rem, _ = st.columns([1.5, 1.5, 3])
                if col_add.button(t['btn_add_row'], key=f"add_{idx}_{item_no_str}", type="primary"):
                    st.session_state[rows_state_key].append({
                        "desc": f"{t['grid_default']} {len(st.session_state[rows_state_key]) + 1}",
                        "no": 1,
                        "l": 50.0 if is_sft else (10.0 if not is_hole else 0.0),
                        "b": 50.0 if is_sft else (10.0 if not is_hole else 0.0),
                        "h": 5.0 if not is_hole else 0.0,
                        "ded": 0.0,
                        "is_deduction_row": False
                    })
                    st.rerun()

                if len(st.session_state[rows_state_key]) > 1 and col_rem.button(t['btn_rem_row'], key=f"rem_{idx}_{item_no_str}"):
                    st.session_state[rows_state_key].pop()
                    st.rerun()

            item_total_qty = max(0.0, item_total_qty)
            item_quantities[item_no_str] = item_total_qty
            
            st.markdown(f"**{t['total_summary']}**")
            st.dataframe(pd.DataFrame(meas_rows), use_container_width=True)
            
            st.info(f"💡 **Item {item_no_str} Total = `{item_total_qty:,.2f} {display_unit}`**")

    meas_excel_buffer = export_measurement_sheet_excel(selected_items_list, st.session_state, lang)
    st.download_button(
        label=t['dl_meas_excel'],
        data=meas_excel_buffer,
        file_name="Earthwork_Measurement_Sheet.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.divider()

    # Step 3: Cost Analysis
    st.markdown(f'<div class="section-title">{t["sec3_title"]}</div>', unsafe_allow_html=True)

    grand_total = 0.0
    material_summary = {}
    labour_summary = {}

    for idx, item in enumerate(selected_items_list):
        item_no = str(item['item_no'])
        measured_qty = item_quantities.get(item_no, 0.0)
        display_unit = item['display_unit']

        st.markdown(f"#### Item {item_no} - {item['display_title']}")

        unit_str = str(item['unit']).lower().strip()
        is_lumpsum = 'l-s' in unit_str or 'ls' in unit_str or 'lump' in unit_str or 'job' in unit_str

        display_rows = []
        item_total_cost = 0.0

        display_rows.append({
            t['item_no_col']: item_no,
            t['particular']: item['display_title'],
            t['unit']: display_unit,
            t['quantity']: f"{measured_qty:,.2f}",
            t['rate_mmk']: "",
            t['amount_mmk']: ""
        })

        if is_lumpsum and not item['breakdown']:
            ls_rate = ls_custom_rates.get(item_no, 0.0)
            amount = measured_qty * ls_rate
            item_total_cost = amount

            display_rows.append({
                t['item_no_col']: "",
                t['particular']: f"  └ {item['display_title']}",
                t['unit']: display_unit,
                t['quantity']: f"{measured_qty:,.2f}",
                t['rate_mmk']: f"{ls_rate:,.2f}",
                t['amount_mmk']: f"{amount:,.2f}"
            })

            labour_summary[item['display_title']] = {
                "unit": display_unit,
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

            extra_worker_depth = 0.0
            extra_worker_lead = 0.0

            rows_state_key = f"rows_data_{item_no}_{idx}"
            current_rows = st.session_state.get(rows_state_key, [])

            # Dynamic Additional Lift/Lead Calculation Logic
            for r in current_rows:
                if r.get('is_deduction_row'):
                    continue
                
                no_val = float(r.get('no', 1))
                l_val = float(r.get('l', 0.0))
                b_val = float(r.get('b', 0.0))
                h_val = float(r.get('h', 0.0))
                ded_val = float(r.get('ded', 0.0))

                row_cft = max(0.0, (no_val * l_val * b_val * h_val) - ded_val)

                # Item 2, 3, 4 (Initial 5 ft depth)
                if item_no in ['2', '2.0', '3', '3.0', '4', '4.0']:
                    if h_val > 5.0:
                        extra_h = h_val - 5.0
                        depth_steps = math.ceil(extra_h / 5.0)
                        extra_worker_depth += depth_steps * 0.5 * (row_cft / 100.0)

                    max_dist = max(l_val, b_val)
                    if max_dist > 100.0:
                        extra_dist = max_dist - 100.0
                        lead_steps = math.ceil(extra_dist / 100.0)
                        extra_worker_lead += lead_steps * 0.5 * (row_cft / 100.0)

                # Item 20 (Initial 10 ft depth, Add 3.38 workers per additional 5 ft lift)
                elif item_no in ['20', '20.0']:
                    if h_val > 10.0:
                        extra_h = h_val - 10.0
                        depth_steps = math.ceil(extra_h / 5.0)
                        extra_worker_depth += depth_steps * 3.38 * (row_cft / 100.0)

                # Item 21 (Initial 10 ft depth, Add 3.00 workers per additional 5 ft lift)
                elif item_no in ['21', '21.0']:
                    if h_val > 10.0:
                        extra_h = h_val - 10.0
                        depth_steps = math.ceil(extra_h / 5.0)
                        extra_worker_depth += depth_steps * 3.0 * (row_cft / 100.0)

            for row in item['breakdown']:
                part = row['particular']
                std_qty = row['qty']
                u = row['unit']

                if std_base_qty <= 0:
                    req_qty = std_qty * measured_qty
                else:
                    req_qty = (std_qty / std_base_qty) * measured_qty

                part_lower = part.lower()
                if ('worker' in part_lower or 'digger' in part_lower) and item_no in ['2', '2.0', '3', '3.0', '4', '4.0', '20', '20.0', '21', '21.0']:
                    req_qty += (extra_worker_depth + extra_worker_lead)

                u_str = str(u).lower().strip()
                if 'l-s' in u_str or 'ls' in u_str or 'lump' in u_str:
                    req_qty = 1.0

                unit_rate = rate_map.get(part, 0.0)
                amount = req_qty * unit_rate
                item_total_cost += amount

                is_labour = any(k in part_lower for k in LABOUR_KEYWORDS)

                disp_part = translate_particular(part, lang)
                disp_unit = translate_unit(u, lang)

                row_data = {
                    "part": disp_part,
                    "unit": disp_unit,
                    "qty": req_qty,
                    "rate": unit_rate,
                    "amount": amount
                }

                if is_labour:
                    lab_breakdown.append(row_data)
                else:
                    mat_breakdown.append(row_data)

                target_dict = labour_summary if is_labour else material_summary
                if disp_part not in target_dict:
                    target_dict[disp_part] = {"unit": disp_unit, "qty": req_qty, "rate": unit_rate, "amount": amount}
                else:
                    target_dict[disp_part]["qty"] += req_qty
                    target_dict[disp_part]["amount"] += amount

            if mat_breakdown:
                display_rows.append({
                    t['item_no_col']: "", t['particular']: t['mat_cost_title'], t['unit']: "", t['quantity']: "", t['rate_mmk']: "", t['amount_mmk']: ""
                })
                for m in mat_breakdown:
                    display_rows.append({
                        t['item_no_col']: "",
                        t['particular']: f"      {m['part']}",
                        t['unit']: m['unit'],
                        t['quantity']: f"{m['qty']:,.2f}",
                        t['rate_mmk']: f"{m['rate']:,.2f}" if m['rate'] > 0 else "-",
                        t['amount_mmk']: f"{m['amount']:,.2f}" if m['amount'] > 0 else "-"
                    })

            if lab_breakdown:
                display_rows.append({
                    t['item_no_col']: "", t['particular']: t['lab_cost_title'], t['unit']: "", t['quantity']: "", t['rate_mmk']: "", t['amount_mmk']: ""
                })
                for l in lab_breakdown:
                    display_rows.append({
                        t['item_no_col']: "",
                        t['particular']: f"      {l['part']}",
                        t['unit']: l['unit'],
                        t['quantity']: f"{l['qty']:,.2f}",
                        t['rate_mmk']: f"{l['rate']:,.2f}",
                        t['amount_mmk']: f"{l['amount']:,.2f}"
                    })

        display_rows.append({
            t['item_no_col']: "",
            t['particular']: t['total_item_cost'],
            t['unit']: "",
            t['quantity']: "",
            t['rate_mmk']: "",
            t['amount_mmk']: f"**{item_total_cost:,.2f}**"
        })

        st.dataframe(pd.DataFrame(display_rows), use_container_width=True, hide_index=True)
        grand_total += item_total_cost

    st.divider()

    # Step 4: BOQ Summary
    st.markdown(f'<div class="section-title">{t["sec4_title"]}</div>', unsafe_allow_html=True)

    st.markdown(f"### {t['mat_boq_title']}")
    mat_rows = []
    total_mat_cost = 0.0
    for idx, (p_name, data) in enumerate(material_summary.items(), start=1):
        mat_rows.append({
            t['sr_no']: idx,
            t['mat_name']: p_name,
            t['unit']: data["unit"],
            t['req_qty']: f"{data['qty']:,.2f}",
            t['rate_mmk']: f"{data['rate']:,.2f}",
            t['total_mmk']: f"{data['amount']:,.2f}"
        })
        total_mat_cost += data["amount"]

    if mat_rows:
        st.table(pd.DataFrame(mat_rows))
    else:
        st.info(t['no_mat_cost'])

    st.divider()

    st.markdown(f"### {t['lab_boq_title']}")
    lab_rows = []
    total_lab_cost = 0.0
    for idx, (p_name, data) in enumerate(labour_summary.items(), start=1):
        lab_rows.append({
            t['sr_no']: idx,
            t['particular']: p_name,
            t['unit']: data["unit"],
            t['quantity']: f"{data['qty']:,.2f}",
            t['rate_mmk']: f"{data['rate']:,.2f}",
            t['total_mmk']: f"{data['amount']:,.2f}"
        })
        total_lab_cost += data["amount"]

    if lab_rows:
        st.table(pd.DataFrame(lab_rows))
    else:
        st.info(t['no_lab_cost'])

    boq_excel_buffer = export_boq_summary_excel(material_summary, labour_summary)
    st.download_button(
        label=t['dl_boq_excel'],
        data=boq_excel_buffer,
        file_name="Earthwork_BOQ_Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.divider()

    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric(t['total_mat_val'], f"{total_mat_cost:,.2f} MMK")
    col_m2.metric(t['total_lab_val'], f"{total_lab_cost:,.2f} MMK")
    col_m3.metric(t['grand_total_val'], f"{grand_total:,.2f} MMK")


if __name__ == "__main__":
    main()
