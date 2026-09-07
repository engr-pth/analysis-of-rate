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
        page_title='QS & Rate Analysis System', layout='wide', page_icon='🏗️'
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

    # Main Tabs: Detail Measurement & Estimate
    tab_measurement, tab_estimate = st.tabs([
        '📐 1. Detail Measurement ( take-off )',
        '📊 2. Rate Analysis & Cost Estimate',
    ])

    # Category Selection
    category = st.radio('လုပ်ငန်းအမျိုးအစား ရွေးပါ:', ['Earth Work', 'Concrete Work'], horizontal=True)
    items_data = earthwork_items if category == 'Earth Work' else concrete_items

    if not items_data:
        st.warning(f'{category} အတွက် Excel ဒေတာ ဖတ်၍မရပါ သို့မဟုတ် ဖိုင်လမ်းကြောင်း မမှန်ပါ။')
        return

    item_options = {f"Item {item['item_no']} - {item['title']}": item for item in items_data}
    selected_title = st.selectbox('Item ရွေးချယ်ပါ:', list(item_options.keys()))
    selected_item = item_options[selected_title]

    # Session state တွင် Quantity သိမ်းဆည်းရန်
    if 'final_qty' not in st.session_state:
        st.session_state.final_qty = 100.0

    # ---------------------------------------------------------
    # TAB 1: Detail Measurement Sheet
    # ---------------------------------------------------------
    with tab_measurement:
        st.subheader('📐 Detail Measurement Calculation')
        unit_str = selected_item['unit']

        col_l, col_w, col_d, col_n = st.columns(4)

        length = col_l.number_input('Length (ft)', min_value=0.0, value=10.0, key='m_l')
        width = col_w.number_input('Width (ft)', min_value=0.0, value=10.0, key='m_w')

        if unit_str == 'sft.':
            depth = col_d.number_input('Thickness (in)', min_value=0.0, value=0.0, key='m_d')
            nos = col_n.number_input('Nos', min_value=1, value=1, key='m_n')
            calc_qty = length * width * nos
        else:
            depth = col_d.number_input('Depth/Height (ft)', min_value=0.0, value=5.0, key='m_d')
            nos = col_n.number_input('Nos', min_value=1, value=1, key='m_n')
            calc_qty = length * width * depth * nos

        st.session_state.final_qty = calc_qty

        st.success(f'✅ တွက်ချက်ရရှိသော Total Quantity = **{calc_qty:,.2f} {unit_str}**')
        st.info('👉 အထက်ပါ Quantity ဖြင့် ကုန်ကျစရိတ် တွက်ချက်ရန် **"2. Rate Analysis & Cost Estimate"** Tab သို့ သွားပါ။')

    # ---------------------------------------------------------
    # TAB 2: Rate Analysis & Cost Estimate
    # ---------------------------------------------------------
    with tab_estimate:
        st.subheader('📊 Rate Analysis & Cost Estimate')

        # Measurement မှ ရရှိလာသော Qty သို့မဟုတ် ကိုယ်တိုင် တိုက်ရိုက် ရိုက်ထည့်လိုပါက ပြင်နိုင်ရန်
        qty_to_use = st.number_input(
            f'အသုံးပြုမည့် Quantity ({selected_item["unit"]}):',
            value=float(st.session_state.final_qty),
            key='est_qty_input'
        )

        if selected_item['breakdown']:
            try:
                std_base_qty = float(selected_item['std_qty'])
            except (ValueError, TypeError):
                std_base_qty = 100.0

            calc_rows = []
            for row in selected_item['breakdown']:
                part = row['particular']
                std_qty = row['qty']
                u = row['unit']

                required_qty = (std_qty / std_base_qty) * qty_to_use
                unit_rate = rate_map.get(part, 0.0)
                total_amount = required_qty * unit_rate

                calc_rows.append({
                    'Particular (အကြောင်းအရာ)': part,
                    'Unit': u,
                    'Required Quantity': round(required_qty, 3),
                    'Rate (MMK)': unit_rate,
                    'Total Amount (MMK)': round(total_amount, 2),
                })

            df_res = pd.DataFrame(calc_rows)
            st.dataframe(df_res, use_container_width=True)

            total_cost = df_res['Total Amount (MMK)'].sum()
            final_unit_rate = total_cost / qty_to_use if qty_to_use > 0 else 0

            c1, c2 = st.columns(2)
            c1.metric(
                label=f'စုစုပေါင်း ကုန်ကျစရိတ် ({qty_to_use:,.2f} {selected_item["unit"]} အတွက်)',
                value=f'{total_cost:,.2f} MMK'
            )
            c2.metric(
                label=f'တစ်ယူနစ် နှုန်းထား (Rate per 1 {selected_item["unit"]})',
                value=f'{final_unit_rate:,.2f} MMK'
            )
        else:
            st.warning('ဤ Item အတွက် Breakdown ဒေတာ မရှိပါ။')


if __name__ == '__main__':
    main()
