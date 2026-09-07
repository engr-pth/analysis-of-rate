import pandas as pd
import streamlit as st


# ==========================================
# 1. Helper Function: Excel Data Parsing
# ==========================================
@st.cache_data
def parse_earthwork_excel(file_path):
    df_raw = pd.read_excel(file_path)
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
        if str(qty).strip().lower() == 'quantity':
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


# ==========================================
# 2. MAIN FUNCTION
# ==========================================
def main():
    st.set_page_config(
        page_title='QS & Rate Analysis System', layout='wide', page_icon='📐'
    )

    st.title('📐 Quantity Surveying & Rate Analysis System')
    st.caption('Detail Measurement + Analysis of Rates')

    # Load Data
    try:
        items_data = parse_earthwork_excel('1 Earth Work.xls')
    except Exception as e:
        st.error(f'Excel ဖိုင် ဖတ်၍မရပါ: {e}')
        return

    # Sidebar: Master Rates
    st.sidebar.header('⚙️ Master Unit Rates (MMK)')
    rate_worker = st.sidebar.number_input('Worker (per day)', value=15000.0)
    rate_digger = st.sidebar.number_input('Digger (per day)', value=18000.0)
    rate_maistry = st.sidebar.number_input('Maistry (per day)', value=25000.0)
    rate_sand = st.sidebar.number_input('Sand (per sud)', value=45000.0)

    rate_map = {
        'Worker': rate_worker,
        'Worker for carrying and ramming': rate_worker,
        'Worker for watering': rate_worker,
        'Worker for carrying': rate_worker,
        'Digger': rate_digger,
        'Maistry': rate_maistry,
        'Sand': rate_sand,
    }

    # 1. Select Item
    st.subheader('📌 ၁။ လုပ်ငန်းအမျိုးအစား ရွေးချယ်ပါ')
    item_options = {
        f"Item {item['item_no']} - {item['title']}": item for item in items_data
    }
    selected_title = st.selectbox(
        'လုပ်ငန်းခေါင်းစဉ်:', list(item_options.keys())
    )
    selected_item = item_options[selected_title]

    st.divider()

    # 2. Detail Measurement Sheet (QS Module)
    st.subheader('📐 ၂။ Detail Measurement (အတိုင်းအတာများ ထည့်သွင်းပါ)')
    col_l, col_w, col_d, col_n = st.columns(4)

    length = col_l.number_input(
        'Length / အလျား (ft)', min_value=0.0, value=10.0
    )
    width = col_w.number_input('Width / အနံ (ft)', min_value=0.0, value=10.0)
    depth = col_d.number_input(
        'Depth / အနက် (ft)', min_value=0.0, value=5.0
    )
    nos = col_n.number_input(
        'Nos / အရေအတွက်', min_value=1, value=1, step=1
    )

    # Calculate Total Measured Quantity
    measured_qty = length * width * depth * nos
    unit_str = selected_item['unit']

    st.info(
        f'📐 **Measured Quantity:** {length}\' × {width}\' × {depth}\' × {nos} ='
        f' **{measured_qty:,.2f} {unit_str}**'
    )

    st.divider()

    # 3. Cost Estimation (Rate Analysis Module)
    st.subheader('📊 ၃။ စုစုပေါင်း ကုန်ကျစရိတ် တွက်ချက်မှု (Estimation)')

    if selected_item['breakdown']:
        # Excel ထဲက Standard Base Qty (ဥပမာ- 100 cft)
        try:
            std_base_qty = float(selected_item['std_qty'])
        except (ValueError, TypeError):
            std_base_qty = 100.0 if unit_str == 'cft' else 1.0

        calc_rows = []
        for row in selected_item['breakdown']:
            part = row['particular']
            std_qty = row['qty']
            unit = row['unit']

            # QS Formula: Actual Required Qty = (Std Qty / Base Qty) * Measured Qty
            required_qty = (std_qty / std_base_qty) * measured_qty
            unit_rate = rate_map.get(part, 0.0)
            total_amount = required_qty * unit_rate

            calc_rows.append({
                'Particular (အကြောင်းအရာ)': part,
                'Unit': unit,
                'Required Quantity': round(required_qty, 3),
                'Rate (MMK)': unit_rate,
                'Total Amount (MMK)': round(total_amount, 2),
            })

        df_qs = pd.DataFrame(calc_rows)
        st.dataframe(df_qs, use_container_width=True)

        # Total Estimated Cost
        total_project_cost = df_qs['Total Amount (MMK)'].sum()
        final_unit_rate = (
            total_project_cost / measured_qty if measured_qty > 0 else 0
        )

        col1, col2 = st.columns(2)
        col1.metric(
            label=f'စုစုပေါင်း ကုန်ကျစရိတ် ({measured_qty:,.2f} {unit_str} အတွက်)',
            value=f'{total_project_cost:,.2f} MMK',
        )
        col2.metric(
            label=f'တစ်ယူနစ် ကုန်ကျစရိတ် (Per 1 {unit_str})',
            value=f'{final_unit_rate:,.2f} MMK',
        )

    else:
        st.warning('ဤ Item အတွက် Breakdown မရှိပါ။')


if __name__ == '__main__':
    main()
