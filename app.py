import os
import pandas as pd
import streamlit as st

# ==========================================
# 2. ရွေးချယ်ထားသော Item များအတွက် Detail Measurement
# ==========================================
st.subheader('📐 ၂။ Detail Measurement Sheet (အတိုင်းအတာများ ရိုက်ထည့်ပါ)')

item_quantities = {}

for idx, item in enumerate(selected_items_list):
    item_no = item['item_no']
    unit_str = item['unit']

    with st.expander(
        f"📌 {idx+1}. Item {item_no} - {item['title']} ({unit_str})",
        expanded=True,
    ):
        st.write('👇 Measurement တန်ဖိုးများကို ဇယားထဲတွင် ရိုက်ထည့်ပါ (Row များ အသစ်ထည့်နိုင်ပါသည်)')

        # Default Table Structure
        default_data = pd.DataFrame([{
            'Particular': 'Main Section',
            'Nos': 1.0,
            'L (ft)': 10.0,
            'B (ft)': 10.0,
            'H (ft)': 5.0 if unit_str != 'sft.' else 0.0,
            'Deduction': 0.0,
        }])

        # Interactive Table Editor
        edited_df = st.data_editor(
            default_data,
            num_rows='dynamic',  # Row များကို စိတ်ကြိုက် တိုး/လျှော့ လုပ္နိုၚ်သည်
            key=f'editor_{idx}_{item_no}',
            use_container_width=True,
            column_config={
                'Particular': st.column_config.TextColumn(
                    'Particular (အကြောင်းအရာ)', default='-'
                ),
                'Nos': st.column_config.NumberColumn(
                    'Nos', min_value=0.0, default=1.0
                ),
                'L (ft)': st.column_config.NumberColumn(
                    'L (ft)', min_value=0.0, default=0.0
                ),
                'B (ft)': st.column_config.NumberColumn(
                    'B (ft)', min_value=0.0, default=0.0
                ),
                'H (ft)': st.column_config.NumberColumn(
                    'H/D (ft or in)'
                    if unit_str == 'sft.'
                    else 'H/D (ft)',
                    min_value=0.0,
                    default=0.0,
                ),
                'Deduction': st.column_config.NumberColumn(
                    'Deduction', min_value=0.0, default=0.0
                ),
            },
        )

        # Calculate Sub-total and Total
        if not edited_df.empty:
            # sft ဖြစ်ပါက H တန်ဖိုးကို မမြှောက်ပါ (သို့မဟုတ် အထူမလိုပါက)
            if unit_str == 'sft.':
                edited_df['Sub-total'] = (
                    edited_df['Nos'] * edited_df['L (ft)'] * edited_df['B (ft)']
                ) - edited_df['Deduction']
            else:
                edited_df['Sub-total'] = (
                    edited_df['Nos']
                    * edited_df['L (ft)']
                    * edited_df['B (ft)']
                    * edited_df['H (ft)']
                ) - edited_df['Deduction']

            # Sub-total အနုတ်ပြသော်လည်း 0 အောက် မငယ်စေရန်
            edited_df['Sub-total'] = edited_df['Sub-total'].apply(
                lambda x: max(0.0, x)
            )

            # Calculation Results Table ကို ပြသခြင်း
            st.markdown('**📊 Calculation Breakdown Sheet:**')
            st.dataframe(edited_df, use_container_width=True)

            total_item_qty = edited_df['Sub-total'].sum()
        else:
            total_item_qty = 0.0

        item_quantities[item_no] = total_item_qty
        st.success(
            f'✅ **Total Quantity for Item {item_no} = {total_item_qty:,.2f} {unit_str}**'
        )
