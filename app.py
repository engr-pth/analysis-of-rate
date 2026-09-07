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
    rate_worker = st.sidebar.number_input('Worker', value=15000.0)
    rate_digger = st.sidebar.number_input('Digger', value=18000.0)
    rate_mason = st.sidebar.number_input('Mason', value=25000.0)
    rate_carpenter = st.sidebar.number_input('Carpenter', value=25000.0)
    rate_maistry = st.sidebar.number_input('Maistry', value=30000.0)

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

    # Main Category Selection
    category = st.radio('လုပ်ငန်းအမျိုးအစား ရွေးပါ:', ['Earth Work', 'Concrete Work'], horizontal=True)
    items_data = earthwork_items if category == 'Earth Work' else concrete_items

    if not items_data:
        st.warning(f'{category} အတွက် Excel ဒေတာ ဖတ်၍မရပါ သို့မဟုတ် ဖိုင်လမ်းကြောင်း မမှန်ပါ။')
        return

    st.subheader(f'📋 {category} - All Items Batch Estimation')
    st.write('အောက်ပါ ဇယားတွင် Item အသီးသီးအတွက် တိုင်းတာရရှိသော Dimensions များကို ရိုက်ထည့်ပါ -')

    # Data Editor အတွက် DataFrame ပြင်ဆင်ခြင်း
    init_rows = []
    for item in items_data:
        init_rows.append({
            'Item No': item['item_no'],
            'Description': item['title'],
            'Unit': item['unit'],
            'Nos': 1.0,
            'Length (ft)': 10.0,
            'Width (ft)': 10.0,
            'Depth/Height (ft/in)': 1.0 if item['unit'] != 'sft.' else 0.0,
        })

    df_input = pd.DataFrame(init_rows)

    # Editable Data Table
    edited_df = st.data_editor(
        df_input,
        column_config={
            'Item No': st.column_config.TextColumn(disabled=True),
            'Description': st.column_config.TextColumn(disabled=True),
            'Unit': st.column_config.TextColumn(disabled=True),
            'Nos': st.column_config.NumberColumn(min_value=0.0, default=1.0),
            'Length (ft)': st.column_config.NumberColumn(min_value=0.0, default=0.0),
            'Width (ft)': st.column_config.NumberColumn(min_value=0.0, default=0.0),
            'Depth/Height (ft/in)': st.column_config.NumberColumn(min_value=0.0, default=0.0),
        },
        use_container_width=True,
        hide_index=True,
        key=f'editor_{category}'
    )

    if st.button('🚀 အားလုံး အစဉ်လိုက် တွက်ချက်မည်', type='primary'):
        results = []
        grand_total = 0.0

        for idx, row in edited_df.iterrows():
            item_ref = items_data[idx]
            u_str = row['Unit']
            nos = row['Nos']
            l = row['Length (ft)']
            w = row['Width (ft)']
            d = row['Depth/Height (ft/in)']

            # Calculated Quantity
            if u_str == 'sft.':
                calc_qty = l * w * nos
            else:
                calc_qty = l * w * d * nos

            # Rate Analysis Calculation
            item_cost = 0.0
            if item_ref['breakdown'] and calc_qty > 0:
                try:
                    std_base = float(item_ref['std_qty'])
                except (ValueError, TypeError):
                    std_base = 100.0

                for b in item_ref['breakdown']:
                    p = b['particular']
                    q = b['qty']
                    req_q = (q / std_base) * calc_qty
                    r = rate_map.get(p, 0.0)
                    item_cost += req_q * r

            unit_rate = item_cost / calc_qty if calc_qty > 0 else 0.0
            grand_total += item_cost

            results.append({
                'Item No': row['Item No'],
                'Description': row['Description'],
                'Total Qty': round(calc_qty, 2),
                'Unit': u_str,
                'Unit Rate (MMK)': round(unit_rate, 2),
                'Total Cost (MMK)': round(item_cost, 2)
            })

        st.divider()
        st.subheader('📊 ရရှိလာသော အစဉ်လိုက် တွက်ချက်မှု ရလဒ်များ (BoQ Summary)')
        
        df_results = pd.DataFrame(results)
        st.dataframe(df_results, use_container_width=True, hide_index=True)

        st.success(f'💰 **စုစုပေါင်း ခန့်မှန်းခြေ ကုန်ကျစရိတ် (Grand Total): {grand_total:,.2f} MMK**')


if __name__ == '__main__':
    main()
