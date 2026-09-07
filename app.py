import pandas as pd
import streamlit as st


# ==========================================
# 1. Helper Function: Excel Data Parsing
# ==========================================
@st.cache_data
def parse_excel_rates(file_path):
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

        # Header စာကြောင်းများ ကျော်ခြင်း
        if (
            item_no in ['nan', 'No.', 'NaN']
            and particular in ['nan', 'Particular', 'NaN']
        ):
            continue
        if str(qty).strip().lower() in ['quantity', 'nan']:
            if item_no in ['nan', 'NaN']:
                # Breakdown row အလွတ်များ
                pass

        # Item ခေါင်းစဉ်အသစ် တွေ့ရှိပါက (Item No. ပါသော စာကြောင်း)
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
        # Item အောက်ရှိ Breakdown (လုပ်အားခ/ပစ္စည်း) စာကြောင်းများ
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


# ==========================================
# 2. MAIN FUNCTION
# ==========================================
def main():
    st.set_page_config(
        page_title='QS & Rate Analysis System', layout='wide', page_icon='🏗️'
    )

    st.title('🏗️ Estimation, QS & Rate Analysis System')
    st.caption('Earth Work & Concrete Work Analysis System')

    # Load Data
    earthwork_items = parse_excel_rates('1 Earth Work.xls')
    concrete_items = parse_excel_rates('2_Concrete(Hand mixed).xls')

    # Sidebar: Master Rates Input
    st.sidebar.header('⚙️ Master Unit Rates (MMK)')

    st.sidebar.subheader('👷 Labor Rates (နေ့တွက်ခ)')
    rate_worker = st.sidebar.number_input('Worker / အလုပ်သမား (per day)', value=15000.0)
    rate_digger = st.sidebar.number_input('Digger / ကျင်းတူး (per day)', value=18000.0)
    rate_mason = st.sidebar.number_input('Mason / ပန်းရန် (per day)', value=25000.0)
    rate_carpenter = st.sidebar.number_input('Carpenter / လက်သမား (per day)', value=25000.0)
    rate_maistry = st.sidebar.number_input('Maistry / ခေါင်းဆောင် (per day)', value=30000.0)

    st.sidebar.subheader('🧱 Material Rates (ပစ္စည်းဈေးနှုန်းများ)')
    rate_cement = st.sidebar.number_input('Cement (per bag)', value=12000.0)
    rate_sand = st.sidebar.number_input('Sand (per sud)', value=45000.0)
    rate_shingle = st.sidebar.number_input('River Shingle / ကျောက် (per sud)', value=85000.0)
    rate_gravel = st.sidebar.number_input('Gravel / ဂဝံကျောက် (per sud)', value=60000.0)
    rate_granite = st.sidebar.number_input('1/4" Granite chipping (per sud)', value=95000.0)
    rate_impermo = st.sidebar.number_input('Impermo / ရေကာဆေး (per lb)', value=3500.0)
    rate_ironite = st.sidebar.number_input('Ironite (per lb)', value=4000.0)
    rate_timber_scantling = st.sidebar.number_input('Timber scantling (per cft)', value=35000.0)
    rate_timber_planks = st.sidebar.number_input('Timber planks 1" (per sft)', value=1200.0)
    rate_nails = st.sidebar.number_input('Nails and spikes (per lb)', value=4500.0)

    # Master Rates Map
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

    # Tab ခွဲခြားခြင်း (Earth Work & Concrete Work)
    tab1, tab2 = st.tabs(['🚜 Earth Work Analysis', '🧱 Concrete Work Analysis'])

    def process_module(items_data, module_name):
        if not items_data:
            st.error(f'{module_name} ဒေတာ မရှိပါ သို့မဟုတ် ဖတ်၍မရပါ။')
            return

        st.subheader(f'📌 ၁။ {module_name} Item ရွေးချယ်ပါ')
        item_options = {
            f"Item {item['item_no']} - {item['title']}": item
            for item in items_data
        }
        selected_title = st.selectbox(
            'လုပ်ငန်းအမျိုးအစား:',
            list(item_options.keys()),
            key=f'select_{module_name}',
        )
        selected_item = item_options[selected_title]

        st.divider()

        # Detail Measurement Section
        st.subheader('📐 ၂။ Detail Measurement (အတိုင်းအတာ ရိုက်ထည့်ပါ)')
        unit_str = selected_item['unit']

        col_l, col_w, col_d, col_n = st.columns(4)

        # Default Label
        d_label = (
            'Thickness / အထူ (in)'
            if unit_str == 'sft.'
            else 'Depth/Height / အနက်/အမြင့် (ft)'
        )

        length = col_l.number_input(
            'Length / အလျား (ft)',
            min_value=0.0,
            value=10.0,
            key=f'l_{selected_title}',
        )
        width = col_w.number_input(
            'Width / အနံ (ft)',
            min_value=0.0,
            value=10.0,
            key=f'w_{selected_title}',
        )

        if unit_str == 'sft.':
            depth = col_d.number_input(
                d_label, min_value=0.0, value=0.0, key=f'd_{selected_title}'
            )
            measured_qty = length * width * (col_n.number_input('Nos', value=1, key=f'n_{selected_title}'))
        else:
            depth = col_d.number_input(
                d_label, min_value=0.0, value=5.0, key=f'd_{selected_title}'
            )
            nos = col_n.number_input(
                'Nos / အရေအတွက်',
                min_value=1,
                value=1,
                key=f'n_{selected_title}',
            )
            measured_qty = length * width * depth * nos

        st.info(f'📐 **Measured Total Quantity:** **{measured_qty:,.2f} {unit_str}**')

        st.divider()

        # Breakdown Calculation
        st.subheader('📊 ၃။ တွက်ချက်ရရှိသော Cost & Quantity Breakdown')

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

                required_qty = (std_qty / std_base_qty) * measured_qty
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
            final_unit_rate = (
                total_cost / measured_qty if measured_qty > 0 else 0
            )

            c1, c2 = st.columns(2)
            c1.metric(
                label=f'စုစုပေါင်း ကုန်ကျစရိတ် ({measured_qty:,.2f} {unit_str} အတွက်)',
                value=f'{total_cost:,.2f} MMK',
            )
            c2.metric(
                label=f'တစ်ယူနစ် နှုန်းထား (Rate per 1 {unit_str})',
                value=f'{final_unit_rate:,.2f} MMK',
            )
        else:
            st.warning('ဤ Item အတွက် Breakdown ဒေတာ မရှိပါ။')

    with tab1:
        process_module(earthwork_items, 'Earth Work')

    with tab2:
        process_module(concrete_items, 'Concrete Work')


if __name__ == '__main__':
    main()
