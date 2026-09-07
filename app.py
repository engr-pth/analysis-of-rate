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

        # ၁။ Header row များနှင့် 'Quantity' စာသားပါသော ခေါင်းစဉ်တန်းများကို ကျော်ပါ
        if (
            item_no in ['nan', 'No.', 'NaN']
            and particular in ['nan', 'Particular', 'NaN']
        ):
            continue
        if str(qty).strip().lower() == 'quantity':
            continue

        # ၂။ Item ခေါင်းစဉ်အသစ် တွေ့ရှိပါက (Item No. ရှိသော စာကြောင်းများ)
        if item_no not in ['nan', 'NaN'] and particular not in ['nan', 'NaN']:
            if current_item:
                items.append(current_item)
            current_item = {
                'item_no': item_no,
                'title': particular,
                'unit': unit,
                'qty': qty,
                'breakdown': [],
            }
        # ၃။ Item အောက်ရှိ Breakdown (လုပ်အားခ/ပစ္စည်း) စာကြောင်းများ
        elif current_item and particular not in ['nan', 'NaN']:
            # စာသားများပါဝင်နေပါက float သို့ ပြောင်းစဉ် အမှားမတက်စေရန် Safe Conversion ပြုလုပ်ခြင်း
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
# 2. MAIN FUNCTION (အဓိက အက်ပ် လုပ်ဆောင်ချက်)
# ==========================================
def main():
    # ---------- A. Page Setup & Title ----------
    st.set_page_config(
        page_title='Rate Analysis - Earth Work',
        layout='wide',
        page_icon='🚜',
    )

    st.title('🚜 Earth Work Rate Analysis System')
    st.caption('1 Earth Work.xls ဒေတာအပေါ် အခြေခံထားသော နှုန်းထားတွက်ချက်စနစ်')

    # ---------- B. Load Data ----------
    try:
        items_data = parse_earthwork_excel('1 Earth Work.xls')
    except Exception as e:
        st.error(f'Excel ဖိုင် ဖတ်၍မရပါ: {e}')
        return

    # ---------- C. Sidebar: Master Labor/Material Rates ----------
    st.sidebar.header('⚙️ 1. Master Unit Rates (MMK)')
    st.sidebar.info('နေ့တွက်ခ/ပစ္စည်းဈေးနှုန်းများ ပြင်ဆင်ပါ')

    rate_worker = st.sidebar.number_input('Worker (per day)', value=15000.0)
    rate_digger = st.sidebar.number_input('Digger (per day)', value=18000.0)
    rate_maistry = st.sidebar.number_input('Maistry (per day)', value=25000.0)
    rate_carpenter = st.sidebar.number_input(
        'Carpenter (per day)', value=22000.0
    )
    rate_surveyor = st.sidebar.number_input('Surveyor (per day)', value=35000.0)
    rate_sand = st.sidebar.number_input('Sand (per sud)', value=45000.0)
    rate_timber = st.sidebar.number_input('Timber (per ton)', value=1200000.0)
    rate_nail = st.sidebar.number_input('Wire Nail (per lb)', value=4000.0)

    # Master Rates Lookup Dict
    rate_map = {
        'Worker': rate_worker,
        'Worker for carrying and ramming': rate_worker,
        'Worker for watering': rate_worker,
        'Worker for carrying': rate_worker,
        'Digger': rate_digger,
        'Maistry': rate_maistry,
        'Carpenter': rate_carpenter,
        'Surveyor': rate_surveyor,
        'Sand': rate_sand,
        'Timber': rate_timber,
        'Wire Nail': rate_nail,
    }

    # ---------- D. Main Section: Item Selection ----------
    st.subheader('📌 2. Select Earth Work Item')

    item_options = {
        f"Item {item['item_no']} - {item['title']} ({item['qty']} {item['unit']})": item
        for item in items_data
    }

    selected_title = st.selectbox('လုပ်ငန်းအမျိုးအစား ရွေးပါ:', list(item_options.keys()))
    selected_item = item_options[selected_title]

    st.divider()

    # ---------- E. Calculation & Results Display ----------
    st.subheader(
        f"📊 Analysis Breakdown for: Item {selected_item['item_no']}"
    )

    if selected_item['breakdown']:
        calc_rows = []
        for row in selected_item['breakdown']:
            part = row['particular']
            qty = row['qty']
            unit = row['unit']

            # Lookup rate in Master Rates
            unit_rate = rate_map.get(part, 0.0)
            amount = qty * unit_rate

            calc_rows.append({
                'Particular (အကြောင်းအရာ)': part,
                'Unit': unit,
                'Quantity': qty,
                'Rate (MMK)': unit_rate,
                'Amount (MMK)': amount,
            })

        df_result = pd.DataFrame(calc_rows)

        # ဇယား ဖော်ပြခြင်း
        st.dataframe(df_result, use_container_width=True)

        # Metrics (စုစုပေါင်း ကုန်ကျစရိတ်နှင့် တစ်ယူနစ်နှုန်း)
        total_amount = df_result['Amount (MMK)'].sum()
        base_qty = (
            float(selected_item['qty'])
            if str(selected_item['qty']).replace('.', '', 1).isdigit()
            else 1.0
        )
        unit_rate_final = total_amount / base_qty if base_qty > 0 else 0

        col1, col2 = st.columns(2)
        col1.metric(
            label=f"Total Cost for {selected_item['qty']} {selected_item['unit']}",
            value=f'{total_amount:,.2f} MMK',
        )
        col2.metric(
            label=f"Final Rate per 1 {selected_item['unit']}",
            value=f'{unit_rate_final:,.2f} MMK',
        )

    else:
        st.warning('ဤ Item အတွက် သီးခြား Breakdown ဒေတာ မရှိပါ။')


# ==========================================
# 3. App Execution Entry Point
# ==========================================
if __name__ == '__main__':
    main()
