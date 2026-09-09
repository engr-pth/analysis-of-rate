import os
import pandas as pd
import streamlit as st

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

    # ==========================================
    # Sidebar Tools: Master Rates, Calculator & Converter
    # ==========================================
    st.sidebar.title("🛠️ Tools & Settings")
    tab_rates, tab_calc, tab_conv = st.sidebar.tabs(["⚙️ Master Rates", "🧮 Calc", "🔄 Converter"])

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

    # 🧮 Sidebar Quick Calculator
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

    # 🔄 Sidebar Unit Converter
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

        # Session State List အား Default အစပြုခြင်း
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

                    # Deduction Row Checkbox
                    st.markdown("<div style='padding-top: 28px;'></div>", unsafe_allow_html=True)
                    is_ded_row = c_is_ded.checkbox("➖ Deduction Row", value=r_data.get("is_deduction_row", False), key=f"is_ded_{idx}_{r_idx}_{item_no_str}")

                    # State Updates
                    r_data["desc"] = p_desc
                    r_data["no"] = no_val
                    r_data["l"] = l_val
                    r_data["b"] = b_val
                    r_data["h"] = h_val
                    r_data["ded"] = ded_val
                    r_data["is_deduction_row"] = is_ded_row

                    # Copy Button
                    if c_cp.button("📋 Copy", key=f"copy_{idx}_{r_idx}_{item_no_str}"):
                        row_to_copy = dict(r_data)

                    # Calculation Logic
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

                # Copy Trigger
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

    st.divider()

    # Grand Totals Summary Box
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("📦 Total Material Cost", f"{total_mat_cost:,.2f} MMK")
    col_m2.metric("👷 Total Labour Cost", f"{total_lab_cost:,.2f} MMK")
    col_m3.metric("💰 Grand Total Cost", f"{grand_total:,.2f} MMK")


if __name__ == "__main__":
    main()
