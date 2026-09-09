import os
import io
import pandas as pd
import streamlit as st
import openpyxl

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


def export_to_excel_with_formulas(selected_items_list, item_quantities, rate_map, ls_custom_rates, session_state):
    """
    Excel တင်ပို့ရာတွင် cell များ၌ Formula များ အဆင်သင့် ပါဝင်အောင် ဖန်တီးပေးသည့် Function
    """
    wb = openpyxl.Workbook()
    
    # -------------------------------------------------------------
    # Sheet 1: Measurement Sheet with Excel Formulas
    # -------------------------------------------------------------
    ws_meas = wb.active
    ws_meas.title = "Measurement Sheet"
    
    headers_meas = ["Item No", "Particular", "No", "L (ft)", "B (ft)", "H (ft)", "Deduction", "Row Type", "Sub-total"]
    ws_meas.append(headers_meas)
    
    current_row_idx = 2
    item_total_rows = {}
    
    for idx, item in enumerate(selected_items_list):
        item_no_str = str(item['item_no'])
        rows_state_key = f"rows_data_{item_no_str}_{idx}"
        unit_str = str(item['unit']).lower().strip()
        
        is_sft = 'sft' in unit_str or 'sq.ft' in unit_str or 'sqft' in unit_str
        is_rft = 'rft' in unit_str or 'lin.ft' in unit_str
        is_lumpsum = 'l-s' in unit_str or 'ls' in unit_str or 'lump' in unit_str or 'job' in unit_str

        start_row = current_row_idx
        
        if is_lumpsum:
            ws_meas.append([item_no_str, f"Item {item_no_str} - {item['title']}", 1, 0, 0, 0, 0, "Addition", 1])
            current_row_idx += 1
        else:
            rows_data = session_state.get(rows_state_key, [])
            for r in rows_data:
                p_desc = r["desc"]
                no = r["no"]
                l = r["l"]
                b = 0.0 if is_rft else r["b"]
                h = 0.0 if (is_sft or is_rft) else r["h"]
                ded = r.get("ded", 0.0)
                is_ded = r.get("is_deduction_row", False)
                row_type = "Deduction" if is_ded else "Addition"

                # Sub-total အတွက် Dynamic Formula ရေးဆွဲခြင်း
                if is_rft:
                    mult_part = f"C{current_row_idx}*D{current_row_idx}"
                elif is_sft:
                    mult_part = f"C{current_row_idx}*D{current_row_idx}*E{current_row_idx}"
                else:
                    mult_part = f"C{current_row_idx}*D{current_row_idx}*E{current_row_idx}*F{current_row_idx}"

                if is_ded:
                    formula = f"=IF({mult_part}-G{current_row_idx} > 0, -({mult_part}-G{current_row_idx}), 0)"
                else:
                    formula = f"=MAX(0, {mult_part}-G{current_row_idx})"

                ws_meas.append([item_no_str, p_desc, no, l, b, h, ded, row_type, formula])
                current_row_idx += 1
                
        end_row = current_row_idx - 1
        
        # Item တစ်ခုချင်းစီ၏ Total Row တွင် SUM Formula ထည့်ခြင်း
        sum_formula = f"=SUM(I{start_row}:I{end_row})"
        ws_meas.append([item_no_str, f"Total Quantity for Item {item_no_str}", "", "", "", "", "", "TOTAL", sum_formula])
        item_total_rows[item_no_str] = current_row_idx
        current_row_idx += 1

    # -------------------------------------------------------------
    # Sheet 2: Cost Estimate Breakdown
    # -------------------------------------------------------------
    ws_est = wb.create_sheet(title="Rate Analysis")
    headers_est = ["Item No", "Particular", "Unit", "Quantity", "Rate (MMK)", "Amount (MMK)"]
    ws_est.append(headers_est)
    
    est_row_idx = 2
    for item in selected_items_list:
        item_no = str(item['item_no'])
        unit_str = str(item['unit']).lower().strip()
        is_lumpsum = 'l-s' in unit_str or 'ls' in unit_str or 'lump' in unit_str or 'job' in unit_str

        # Measurement Sheet မှ Qty ကို Formula ဖြင့် လှမ်းယူခြင်း
        meas_total_cell = f"'Measurement Sheet'!I{item_total_rows[item_no]}"
        
        ws_est.append([item_no, item['title'], item['unit'], f"={meas_total_cell}", "", ""])
        est_row_idx += 1

        if is_lumpsum and not item['breakdown']:
            ls_rate = ls_custom_rates.get(item_no, 0.0)
            amount_formula = f"=D{est_row_idx-1}*E{est_row_idx}"
            ws_est.append(["", f"  └ {item['title']}", item['unit'], f"=D{est_row_idx-1}", ls_rate, amount_formula])
            est_row_idx += 1
        elif item['breakdown']:
            try:
                std_base_qty = float(item['std_qty'])
            except (ValueError, TypeError):
                std_base_qty = 100.0

            for row in item['breakdown']:
                part = row['particular']
                std_qty = float(row['qty'])
                u = row['unit']
                unit_rate = rate_map.get(part, 0.0)

                # Formula: Quantity = (std_qty / std_base_qty) * Total_Measured_Qty
                qty_formula = f"=({std_qty}/{std_base_qty})*D{est_row_idx-1}"
                amount_formula = f"=D{est_row_idx}*E{est_row_idx}"

                ws_est.append(["", f"      {part}", u, qty_formula, unit_rate, amount_formula])
                est_row_idx += 1

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def main():
    st.set_page_config(
        page_title="QS & Rate Analysis System", layout="wide", page_icon="🏗️"
    )

    st.title("🏗️ Estimation, QS & Rate Analysis System")

    # Load Data
    earthwork_path = os.path.join(BASE_DIR, "1 Earth Work.xls")
    concrete_path = os.path.join(BASE_DIR, "2 Concrete ( Hand mixed ).xls")
    iron_path = os.path.join(BASE_DIR, "3 Iron and Steel work.xls")

    earthwork_items = parse_excel_rates(earthwork_path)
    concrete_items = parse_excel_rates(concrete_path)
    iron_items = parse_excel_rates(iron_path)

    # Sidebar Settings
    st.sidebar.title("🛠️ Tools & Settings")
    tab_rates, tab_calc, tab_conv = st.sidebar.tabs(["⚙️ Master Rates", "🧮 Calc", "🔄 Converter"])

    with tab_rates:
        st.header("Master Unit Rates (MMK)")
        rate_worker = st.number_input("Worker", value=25000.0)
        rate_digger = st.number_input("Digger", value=25000.0)
        rate_mason = st.number_input("Mason", value=25000.0)
        rate_carpenter = st.number_input("Carpenter", value=35000.0)
        rate_maistry = st.number_input("Maistry", value=30000.0)
        rate_blacksmith = st.number_input("Blacksmith / Steel Worker", value=28000.0)
        rate_welder = st.number_input("Welder", value=30000.0)
        rate_surveyor = st.number_input("Surveyor", value=35000.0)

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
            except Exception:
                st.error("တွက်ချက်မှု မမှန်ကန်ပါ။")

    with tab_conv:
        st.subheader("🔄 Unit Converter")
        conv_type = st.selectbox("Convert Type:", [
            "Inches -> Feet", "Sft <-> Sq.m", "Cft <-> Cu.m", "Steel Weight (Dia -> Kg/Ton)"
        ])
        if conv_type == "Inches -> Feet":
            inch_val = st.number_input("Inches (လက်မ):", min_value=0.0, value=6.0)
            st.info(f"👉 **{inch_val} inches = {inch_val / 12.0:.3f} ft**")

    rate_map = {
        "Worker": rate_worker, "Worker for carrying and ramming": rate_worker,
        "Worker for watering": rate_worker, "Worker for carrying": rate_worker,
        "Digger": rate_digger, "Mason": rate_mason, "Carpenter": rate_carpenter,
        "Maistry": rate_maistry, "Blacksmith": rate_blacksmith, "Steel worker": rate_blacksmith,
        "Welder": rate_welder, "Surveyor": rate_surveyor, "Cement": rate_cement,
        "Sand": rate_sand, "River Shingle (1-1/2\" gauge)": rate_shingle,
        "River Shingle (1/4\" to 3/4\" gauge)": rate_shingle, "River Shingle (3/4\" gauge)": rate_shingle,
        "River Shingle (3/4\" to 1-1/2\" gauge)": rate_shingle, "Gravel": rate_gravel,
        "1/4\" Granite chipping": rate_granite, "Impermo": rate_impermo, "Ironite": rate_ironite,
        "Timber scantling": rate_timber_scantling, "Tinber planks 1\"": rate_timber_planks,
        "Nails and spikes": rate_nails, "Wire Nails": rate_nails, "M.S. Bar": rate_steel_bar,
        "Reinforcement Steel": rate_steel_bar, "Binding Wire": rate_binding_wire,
        "Structural Steel": rate_structural_steel, "R.S. girder": rate_rs_girder,
        "Carriage to site": rate_carriage, "Hoisting and fixing": rate_hoisting,
    }

    # Item Selection
    st.subheader("📋 ၁။ တွက်ချက်လိုသော လုပ်ငန်းအမျိုးအစားများ ရွေးချယ်ပါ")
    col_e, col_c, col_i = st.columns(3)
    show_earthwork = col_e.checkbox("🚜 Earth Work", value=True)
    show_concrete = col_c.checkbox("🧱 Concrete Work", value=True)
    show_iron = col_i.checkbox("⚙️ Iron & Steel Work", value=True)

    selected_items_list = []
    if show_earthwork and earthwork_items:
        ew_options = {f"[Earth] Item {i['item_no']} - {i['title']}": i for i in earthwork_items}
        ew_selected = st.multiselect("🚜 Earth Work မှ တွက်မည်များ ရွေးရန်:", list(ew_options.keys()))
        for key in ew_selected:
            selected_items_list.append(ew_options[key])

    if show_concrete and concrete_items:
        cc_options = {f"[Concrete] Item {i['item_no']} - {i['title']}": i for i in concrete_items}
        cc_selected = st.multiselect("🧱 Concrete Work မှ တွက်မည်များ ရွေးရန်:", list(cc_options.keys()))
        for key in cc_selected:
            selected_items_list.append(cc_options[key])

    if show_iron and iron_items:
        ir_options = {f"[Iron] Item {i['item_no']} - {i['title']}": i for i in iron_items}
        ir_selected = st.multiselect("⚙️ Iron & Steel Work မှ တွက်မည်များ ရွေးရန်:", list(ir_options.keys()))
        for key in ir_selected:
            selected_items_list.append(ir_options[key])

    if not selected_items_list:
        st.info("👉 တွက်ချက်လိုသော Item များကို အထက်တွင် ရွေးချယ်ပေးပါ။")
        return

    st.divider()

    # Measurement Sheet Inputs
    st.subheader("📐 ၂။ Detail Measurement Sheet")
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
                {"desc": "Section 1", "no": 1, "l": 50.0 if is_sft else 10.0, "b": 50.0 if is_sft else 10.0, "h": 5.0, "ded": 0.0, "is_deduction_row": False}
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
                meas_rows.append({"Particular": p_desc, "No": no_val, "L (ft)": "-", "B (ft)": "-", "H (ft)": "-", "Type": "Add", "Sub-total": no_val})
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

                    p_desc = c_desc.text_input("Particular Name", value=r_data["desc"], key=f"desc_{idx}_{r_idx}_{item_no_str}")
                    no_val = c_no.number_input("No", min_value=1, value=int(r_data["no"]), key=f"no_{idx}_{r_idx}_{item_no_str}")
                    l_val = c_l.number_input("L (ft)", min_value=0.0, value=float(r_data["l"]), key=f"l_{idx}_{r_idx}_{item_no_str}")
                    
                    b_val = 0.0 if is_rft else c_b.number_input("B (ft)", min_value=0.0, value=float(r_data["b"]), key=f"b_{idx}_{r_idx}_{item_no_str}")
                    h_val = 0.0 if (is_sft or is_rft) else c_h.number_input("H (ft)", min_value=0.0, value=float(r_data["h"]), key=f"h_{idx}_{r_idx}_{item_no_str}")
                    ded_val = c_ded.number_input("Deduction", min_value=0.0, value=float(r_data.get("ded", 0.0)), key=f"ded_{idx}_{r_idx}_{item_no_str}")

                    is_ded_row = c_is_ded.checkbox("➖ Deduction Row", value=r_data.get("is_deduction_row", False), key=f"is_ded_{idx}_{r_idx}_{item_no_str}")

                    r_data.update({"desc": p_desc, "no": no_val, "l": l_val, "b": b_val, "h": h_val, "ded": ded_val, "is_deduction_row": is_ded_row})

                    if c_cp.button("📋 Copy", key=f"copy_{idx}_{r_idx}_{item_no_str}"):
                        row_to_copy = dict(r_data)

                    gross_qty = no_val * l_val * (b_val if not is_rft else 1.0) * (h_val if not (is_sft or is_rft) else 1.0)
                    row_qty = max(0.0, gross_qty - ded_val)

                    if is_ded_row:
                        item_total_qty -= row_qty
                        sub_total_display = -round(row_qty, 2)
                    else:
                        item_total_qty += row_qty
                        sub_total_display = round(row_qty, 2)

                    meas_rows.append({
                        "Particular": p_desc, "No": no_val, "L (ft)": l_val,
                        "B (ft)": b_val if not is_rft else "-", "H (ft)": h_val if not (is_sft or is_rft) else "-",
                        "Deduction": ded_val, "Row Type": "➖ Deduction" if is_ded_row else "➕ Addition",
                        "Sub-total": sub_total_display
                    })

                if row_to_copy:
                    copied = dict(row_to_copy)
                    copied["desc"] += " (Copy)"
                    st.session_state[rows_state_key].append(copied)
                    st.rerun()

                col_add, col_rem, _ = st.columns([1, 1, 4])
                if col_add.button("➕ Add Row", key=f"add_{idx}_{item_no_str}"):
                    st.session_state[rows_state_key].append({"desc": f"Section {len(st.session_state[rows_state_key])+1}", "no": 1, "l": 10.0, "b": 10.0, "h": 5.0, "ded": 0.0, "is_deduction_row": False})
                    st.rerun()

                if len(st.session_state[rows_state_key]) > 1 and col_rem.button("➖ Remove Row", key=f"rem_{idx}_{item_no_str}"):
                    st.session_state[rows_state_key].pop()
                    st.rerun()

            item_total_qty = max(0.0, item_total_qty)
            item_quantities[item_no_str] = item_total_qty
            st.dataframe(pd.DataFrame(meas_rows), use_container_width=True)
            st.markdown(f"**Total Quantity for Item {item_no_str} = `{item_total_qty:,.2f} {item['unit']}`**")

    # Excel Download Export Button
    st.divider()
    excel_file = export_to_excel_with_formulas(selected_items_list, item_quantities, rate_map, ls_custom_rates, st.session_state)
    
    st.sidebar.markdown("---")
    st.sidebar.download_button(
        label="📥 Excel Export (Formulas ပါဝင်သည်)",
        data=excel_file,
        file_name="Estimation_QS_Formulas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.download_button(
        label="📥 Excel ဖိုင်အား Formula များပါရှိစွာ ဒေါင်းလုဒ်ဆွဲရန်",
        data=excel_file,
        file_name="Estimation_QS_Formulas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    main()
