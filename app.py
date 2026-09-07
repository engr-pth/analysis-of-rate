import os
import pandas as pd
import streamlit as st

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


@st.cache_data
def parse_excel_rates(file_path):
    if not os.path.exists(file_path):
        return []
    try:
        # Excel ဖိုင်ထဲမှ Data များကို Date မဟုတ်ဘဲ String (Text) အဖြစ် အတင်းပြောင်းဖတ်ပါ
        df_raw = pd.read_excel(file_path, dtype=str)
    except Exception as e:
        st.error(f"ဖိုင်ဖတ်၍ မရပါ ({file_path}): {e}")
        return []

    items = []
    current_item = None

    for idx, row in df_raw.iterrows():
        # Clean text
        item_no = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
        particular = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
        unit = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ''
        qty = row.iloc[3] if pd.notna(row.iloc[3]) else '0'

        # Date သို့မဟုတ် NaT ပေါ်စေသည့် စာသားများကို ကျော်သွားမည်
        if 'nat' in item_no.lower() or '202' in item_no or item_no in ['', 'nan', 'No.', 'NaN']:
            # အကယ်၍ Item No မဟုတ်ဘဲ Particulars သာရှိလျှင် Breakdown အောက်သို့ ထည့်မည်
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

        # Item Header သစ်တွေ့လျှင်
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
        page_title='QS & Rate Analysis System', layout='wide', page_icon='🏗️'
    )

    st.title('🏗️ Estimation, QS & Rate Analysis System')

    # Load Data
    earthwork_path = os.path.join(BASE_DIR, '1 Earth Work.xls')
    concrete_path = os.path.join(BASE_DIR, '2 Concrete ( Hand mixed ).xls')
    iron_path = os.path.join(BASE_DIR, '3 Iron and Steel work.xls')

    earthwork_items = parse_excel_rates(earthwork_path)
    concrete_items = parse_excel_rates(concrete_path)
    iron_items = parse_excel_rates(iron_path)

    # Master Rates Sidebar
    st.sidebar.header('⚙️ Master Unit Rates (MMK)')
    rate_worker = st.sidebar.number_input('Worker', value=15000.0)
    rate_digger = st.sidebar.number_input('Digger', value=18000.0)
    rate_mason = st.sidebar.number_input('Mason', value=25000.0)
    rate_carpenter = st.sidebar.number_input('Carpenter', value=25000.0)
    rate_maistry = st.sidebar.number_input('Maistry', value=30000.0)
    rate_blacksmith = st.sidebar.number_input('Blacksmith / Steel Worker', value=28000.0)
    rate_welder = st.sidebar.number_input('Welder', value=30000.0)

    rate_cement = st.sidebar.number_input('Cement', value=12000.0)
    rate_sand = st.sidebar.number_input('Sand', value=45000.0)
    rate_shingle = st.sidebar.number_input('River Shingle', value=85000.0)
    rate_gravel = st.sidebar.number_input('Gravel', value=60000.0)
    rate_granite = st.sidebar.number_input('Granite chipping', value=95000.0)
    rate_impermo = st.sidebar.number_input('Impermo', value=3500.0)
    rate_ironite = st.sidebar.number_input('Ironite', value=4000.0)
    rate_timber_scantling = st.sidebar.number_input('Timber scantling', value=35000.0)
    rate_timber_planks = st.sidebar.number_input('Timber planks', value=1200.0)
    rate_nails = st.sidebar.number_input('Nails and spikes', value=4500.0)

    # Steel / Iron Material Rates
    rate_steel_bar = st.sidebar.number_input('M.S. Bar / Reinforcement (Ton)', value=2800000.0)
    rate_binding_wire = st.sidebar.number_input('Binding Wire (Viss/Ib)', value=6500.0)
    rate_structural_steel = st.sidebar.number_input('Structural Steel (Ton)', value=3000000.0)

    rate_map = {
        'Worker': rate_worker,
        'Worker for carrying and ramming': rate_worker,
        'Worker for watering': rate_worker,
        'Worker for carrying': rate_worker,
        'Digger': rate_digger,
        'Mason': rate_mason,
        'Carpenter': rate_carpenter,
        'Maistry': rate_maistry,
        'Blacksmith': rate_blacksmith,
        'Steel worker': rate_blacksmith,
        'Welder': rate_welder,
        'Cement': rate_cement,
        'Sand': rate_sand,
        'River Shingle (1-1/2" gauge)': rate_shingle,
        'River Shingle (1/4" to 3/4" gauge)': rate_shingle,
        'River Shingle (3/4" gauge)': rate_shingle,
        'River Shingle (3/4" to 1-1/2" gauge)': rate_shingle,
        'Gravel': rate_gravel,
        '1/4" Granite chipping': rate_granite,
        'Impermo': rate_impermo,
        'Ironite': rate_ironite,
        'Timber scantling': rate_timber_scantling,
        'Tinber planks 1"': rate_timber_planks,
        'Nails and spikes': rate_nails,
        'M.S. Bar': rate_steel_bar,
        'Reinforcement Steel': rate_steel_bar,
        'Binding Wire': rate_binding_wire,
        'Structural Steel': rate_structural_steel,
    }

    # ==========================================
    # 1. လုပ်ငန်းအမျိုးအစား Checkbox များ
    # ==========================================
    st.subheader('📋 ၁။ တွက်ချက်လိုသော လုပ်ငန်းအမျိုးအစားများ ရွေးချယ်ပါ')
    col_e, col_c, col_i = st.columns(3)
    show_earthwork = col_e.checkbox('🚜 Earth Work', value=True)
    show_concrete = col_c.checkbox('🧱 Concrete Work', value=True)
    show_iron = col_i.checkbox('⚙️ Iron & Steel Work', value=True)

    selected_items_list = []

    if show_earthwork and earthwork_items:
        ew_options = {f"[Earth] Item {i['item_no']} - {i['title']}": i for i in earthwork_items}
        ew_selected = st.multiselect('🚜 Earth Work မှ တွက်မည်များ ရွေးရန်:', list(ew_options.keys()))
        for key in ew_selected:
            selected_items_list.append(ew_options[key])

    if show_concrete and concrete_items:
        cc_options = {f"[Concrete] Item {i['item_no']} - {i['title']}": i for i in concrete_items}
        cc_selected = st.multiselect('🧱 Concrete Work မှ တွက်မည်များ ရွေးရန်:', list(cc_options.keys()))
        for key in cc_selected:
            selected_items_list.append(cc_options[key])

    if show_iron and iron_items:
        ir_options = {f"[Iron] Item {i['item_no']} - {i['title']}": i for i in iron_items}
        ir_selected = st.multiselect('⚙️ Iron & Steel Work မှ တွက်မည်များ ရွေးရန်:', list(ir_options.keys()))
        for key in ir_selected:
            selected_items_list.append(ir_options[key])

    if not selected_items_list:
        st.info('👉 တွက်ချက်လိုသော Item များကို အထက်တွင် ရွေးချယ်ပေးပါ။')
        return

    st.divider()

    # ==========================================
    # 2. Detail Measurement Sheet Input
    # ==========================================
    st.subheader('📐 ၂။ Detail Measurement Sheet (အတိုင်းအတာများ ရိုက်ထည့်ပါ)')

    item_quantities = {}

    for idx, item in enumerate(selected_items_list):
        with st.expander(f"📌 {idx+1}. Item {item['item_no']} - {item['title']} ({item['unit']})", expanded=True):
            unit_str = str(item['unit']).lower().strip()

            particular_desc = st.text_input(
                'Particular (အကြောင်းအရာ/နေရာ):',
                value=f"Location/Section for {item['title']}",
                key=f'part_{idx}_{item["item_no"]}'
            )

            c_no, c_l, c_b, c_h, c_ded = st.columns(5)

            no_val = c_no.number_input('No', min_value=1, value=1, key=f'no_{idx}_{item["item_no"]}')
            l_val = c_l.number_input('L (ft)', min_value=0.0, value=10.0, key=f'l_{idx}_{item["item_no"]}')
            b_val = c_b.number_input('B (ft)', min_value=0.0, value=10.0, key=f'b_{idx}_{item["item_no"]}')
            h_val = c_h.number_input('H (ft)', min_value=0.0, value=5.0, key=f'h_{idx}_{item["item_no"]}')
            ded_val = c_ded.number_input('Deduction', min_value=0.0, value=0.0, key=f'ded_{idx}_{item["item_no"]}')

            # Calculation Logic
            if 'rft' in unit_str:
                gross_qty = no_val * l_val
            elif 'sft' in unit_str:
                if l_val > 0 and b_val > 0:
                    gross_qty = no_val * l_val * b_val
                elif l_val > 0 and h_val > 0:
                    gross_qty = no_val * l_val * h_val
                elif b_val > 0 and h_val > 0:
                    gross_qty = no_val * b_val * h_val
                else:
                    gross_qty = 0.0
            elif 'ton' in unit_str or 'cwt' in unit_str or 'lb' in unit_str or 'kg' in unit_str:
                # Steel/Iron Weight based items (If L is used as Weight or Quantity)
                gross_qty = no_val * l_val
            else:
                # Cft and default
                gross_qty = no_val * l_val * b_val * h_val

            sub_total = max(0.0, gross_qty - ded_val)
            item_quantities[item['item_no']] = sub_total

            meas_df = pd.DataFrame([{
                'Particular': particular_desc,
                'No': no_val,
                'L (ft)': l_val,
                'B (ft)': b_val,
                'H (ft)': h_val,
                'Deduction': ded_val,
                'Sub-total': round(sub_total, 2),
                'Total Qty': f"{round(sub_total, 2)} {item['unit']}"
            }])

            st.dataframe(meas_df, use_container_width=True)

    st.divider()

    # ==========================================
    # 3. အစဉ်လိုက် Cost Breakdown & Total Estimate
    # ==========================================
    st.subheader('📊 ၃။ စုစုပေါင်း Rate Analysis & Cost Estimate')

    grand_total = 0.0
    summary_rows = []

    for idx, item in enumerate(selected_items_list):
        item_no = item['item_no']
        measured_qty = item_quantities.get(item_no, 0.0)

        st.markdown(f"#### 🔹 {idx+1}. Item {item_no} - {item['title']} ({measured_qty:,.2f} {item['unit']})")

        if item['breakdown']:
            try:
                std_base_qty = float(item['std_qty'])
            except (ValueError, TypeError):
                std_base_qty = 100.0

            calc_rows = []
            for row in item['breakdown']:
                part = row['particular']
                std_qty = row['qty']
                u = row['unit']

                req_qty = (std_qty / std_base_qty) * measured_qty
                unit_rate = rate_map.get(part, 0.0)
                amount = req_qty * unit_rate

                calc_rows.append({
                    'Particular': part,
                    'Unit': u,
                    'Req Qty': round(req_qty, 3),
                    'Rate (MMK)': unit_rate,
                    'Amount (MMK)': round(amount, 2),
                })

            df_item = pd.DataFrame(calc_rows)
            st.dataframe(df_item, use_container_width=True)

            item_total = df_item['Amount (MMK)'].sum()
            grand_total += item_total

            summary_rows.append({
                'Item No': item_no,
                'Title': item['title'],
                'Total Qty': f"{measured_qty:,.2f} {item['unit']}",
                'Total Amount (MMK)': f"{item_total:,.2f}",
            })
        else:
            st.warning('Breakdown ဒေတာ မရှိပါ။')

    st.divider()

    # Summary Table
    st.subheader('📜 ရွေးချယ်ခဲ့သော လုပ်ငန်းများ၏ စုစုပေါင်း အနှစ်ချုပ် (BOQ Summary)')
    st.table(pd.DataFrame(summary_rows))

    st.metric(label='💰 စုစုပေါင်း ကုန်ကျစရိတ် (Grand Total Estimate)', value=f'{grand_total:,.2f} MMK')


if __name__ == '__main__':
    main()
