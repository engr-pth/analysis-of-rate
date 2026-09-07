import os
import pandas as pd
import streamlit as st

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


@st.cache_data
def parse_excel_rates(file_path):
    if not os.path.exists(file_path):
        return []
    try:
        df_raw = pd.read_excel(file_path)
    except Exception as e:
        st.error(f'ဖိုင်ဖတ်၍ မရပါ ({file_path}): {e}')
        return []

    items = []
    current_item = None

    for idx, row in df_raw.iterrows():
        item_no = str(row.iloc[0]).strip()
        particular = str(row.iloc[1]).strip()
        unit = str(row.iloc[2]).strip()
        qty = row.iloc[3]

        if (
            item_no in ['nan', 'No.', 'NaN']
            and particular in ['nan', 'Particular', 'NaN']
        ):
            continue

        if item_no not in ['nan', 'NaN'] and particular not in ['nan', 'NaN']:
            if current_item:
                items.append(current_item)
            current_item = {
                'item_no': item_no,
                'title': particular,
                'unit': unit,
                'std_qty': qty,
                'breakdown': [],
            }
        elif current_item and particular not in ['nan', 'NaN']:
            try:
                qty_val = float(qty)
            except (ValueError, TypeError):
                qty_val = 0.0

            current_item['breakdown'].append({
                'particular': particular,
                'unit': unit,
                'qty': qty_val,
            })

    if current_item:
        items.append(current_item)

    return items


def main():
    st.set_page_config(
        page_title='QS Multi-Item Estimation System', layout='wide', page_icon='🏗️'
    )

    st.title('🏗️ Estimation, QS & Rate Analysis System')

    # Load Data
    earthwork_path = os.path.join(BASE_DIR, '1 Earth Work.xls')
    concrete_path = os.path.join(BASE_DIR, '2 Concrete ( Hand mixed ).xls')

    earthwork_items = parse_excel_rates(earthwork_path)
    concrete_items = parse_excel_rates(concrete_path)

    # Master Rates Sidebar
    st.sidebar.header('⚙️ Master Unit Rates (MMK)')
    st.sidebar.subheader('👷 Labor Rates')
    rate_worker = st.sidebar.number_input('Worker', value=15000.0)
    rate_digger = st.sidebar.number_input('Digger', value=18000.0)
    rate_mason = st.sidebar.number_input('Mason', value=25000.0)
    rate_carpenter = st.sidebar.number_input('Carpenter', value=25000.0)
    rate_maistry = st.sidebar.number_input('Maistry', value=30000.0)

    st.sidebar.subheader('🧱 Material Rates')
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

    rate_map = {
        'Worker': rate_worker,
        'Worker for carrying and ramming': rate_worker,
        'Worker for watering': rate_worker,
        'Worker for carrying': rate_worker,
        'Digger': rate_digger,
        'Mason': rate_mason,
        'Carpenter': rate_carpenter,
        'Maistry': rate_maistry,
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
    }

    # All Available Items Combination
    all_items = {}
    for item in earthwork_items:
        key = f"[Earth Work] Item {item['item_no']} - {item['title']}"
        all_items[key] = item
    for item in concrete_items:
        key = f"[Concrete Work] Item {item['item_no']} - {item['title']}"
        all_items[key] = item

    st.subheader('📋 ၁။ တွက်ချက်လိုသော Work Items များအားလုံးကို ရွေးချယ်ပါ')
    selected_keys = st.multiselect(
        'Earthwork နှင့် Concrete Work နှစ်သက်ရာ Items အများအပြားကို ရွေးပါ:',
        options=list(all_items.keys()),
        default=list(all_items.keys())[:2] if all_items else []
    )

    if not selected_keys:
        st.info('တွက်ချက်ရန် Item အနည်းဆုံး တစ်ခု ရွေးချယ်ပါ။')
        return

    # Tabs for Measurement & Estimate
    tab_measurement, tab_estimate = st.tabs([
        '📐 1. Detail Measurement ( take-off )',
        '📊 2. Rate Analysis & Summary Estimate',
    ])

    # ---------------------------------------------------------
    # TAB 1: Detail Measurement (Multi-Item Table)
    # ---------------------------------------------------------
    with tab_measurement:
        st.subheader('📐 Measurements ရိုက်ထည့်ရန် ဇယား')
        st.caption('ဇယားထဲတွင် Length, Width, Depth/Thickness, Nos များကို စိတ်ကြိုက် ပြင်ဆင်နိုင်ပါသည်။')

        # Construct Measurement Data Table
        meas_rows = []
        for key in selected_keys:
            item = all_items[key]
            meas_rows.append({
                'Item Name': key,
                'Unit': item['unit'],
                'Length (ft)': 10.0,
                'Width (ft)': 10.0,
                'Depth/Thk (ft/in)': 0.0 if item['unit'] == 'sft.' else 5.0,
                'Nos': 1,
            })

        df_meas_input = pd.DataFrame(meas_rows)

        # Interactive Table
        edited_df = st.data_editor(
            df_meas_input,
            num_rows='fixed',
            use_container_width=True,
            key='measurement_editor'
        )

        # Quantity Calculation
        calculated_quantities = {}
        for idx, row in edited_df.iterrows():
            item_key = row['Item Name']
            u = row['Unit']
            l = float(row['Length (ft)'])
            w = float(row['Width (ft)'])
            d = float(row['Depth/Thk (ft/in)'])
            n = int(row['Nos'])

            if u == 'sft.':
                total_q = l * w * n
            else:
                total_q = l * w * d * n

            calculated_quantities[item_key] = total_q

        st.session_state.calculated_quantities = calculated_quantities

    # ---------------------------------------------------------
    # TAB 2: Cost Analysis & Summary
    # ---------------------------------------------------------
    with tab_estimate:
        st.subheader('📊 စုစုပေါင်း ကုန်ကျစရိတ်နှင့် Detailed Analysis')

        total_grand_cost = 0.0
        summary_list = []
        breakdown_all_list = []

        for key in selected_keys:
            item = all_items[key]
            qty_to_use = st.session_state.calculated_quantities.get(key, 0.0)

            try:
                std_base_qty = float(item['std_qty'])
            except (ValueError, TypeError):
                std_base_qty = 100.0

            item_total_cost = 0.0

            for row in item['breakdown']:
                part = row['particular']
                std_qty = row['qty']
                u = row['unit']

                req_qty = (std_qty / std_base_qty) * qty_to_use if std_base_qty > 0 else 0.0
                unit_rate = rate_map.get(part, 0.0)
                amount = req_qty * unit_rate
                item_total_cost += amount

                breakdown_all_list.append({
                    'Item Category': key,
                    'Particular': part,
                    'Unit': u,
                    'Required Qty': round(req_qty, 3),
                    'Rate (MMK)': unit_rate,
                    'Amount (MMK)': round(amount, 2)
                })

            unit_rate_final = item_total_cost / qty_to_use if qty_to_use > 0 else 0.0
            total_grand_cost += item_total_cost

            summary_list.append({
                'Item Description': key,
                'Quantity': round(qty_to_use, 2),
                'Unit': item['unit'],
                'Unit Rate (MMK)': round(unit_rate_final, 2),
                'Total Amount (MMK)': round(item_total_cost, 2)
            })

        # Display Summary Table
        st.markdown('### 📌 BOQ Summary Table')
        df_summary = pd.DataFrame(summary_list)
        st.dataframe(df_summary, use_container_width=True)

        st.metric('💰 စုစုပေါင်း ကုန်ကျစရိတ် (Grand Total Amount)', f'{total_grand_cost:,.2f} MMK')

        st.divider()

        # Detailed Material/Labor Breakdown Expansion
        st.markdown('### 🔍 Detailed Material & Labor Breakdown (အသေးစိတ် ပစ္စည်း/လုပ်အားခ)')
        df_breakdown = pd.DataFrame(breakdown_all_list)
        st.dataframe(df_breakdown, use_container_width=True)


if __name__ == '__main__':
    main()
