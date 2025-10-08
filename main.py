import streamlit as st
import pandas as pd
import re
import random
import string
import io

st.set_page_config(page_title="🚀 Monetix Sheet Update", layout="wide")
st.title("🚀 Monetix Sheet Update")

st.markdown("Upload multiple merchant Excel sheets (`.xlsx`) to generate a clean customer payout file.")

merchant_files = st.file_uploader("📁 Upload Merchant Excel Files", type=["xlsx"], accept_multiple_files=True)

if st.button("🚀 Generate Output Sheet"):
    if not merchant_files:
        st.error("Please upload at least one Excel file.")
    else:
        # --- Helpers ---
        def extract_valid_account(value):
            if isinstance(value, str):
                cleaned = re.sub(r'[^A-Z0-9]', '', value.upper())
                if cleaned.startswith("PK") and len(cleaned) >= 24:
                    return cleaned
            return None

        def extract_amount(value):
            if pd.isna(value): return None
            if isinstance(value, str):
                digits = re.sub(r'[^\d]', '', value)
                return digits if digits else None
            elif isinstance(value, (int, float)):
                return str(int(value))
            return None

        def generate_random_id(prefix='', length=12):
            return prefix + ''.join(random.choices(string.ascii_uppercase + string.digits, k=length))

        def generate_random_name():
            first_names = ['Ali', 'Sara', 'Ahmed', 'Zara', 'Usman', 'Ayesha']
            last_names = ['Khan', 'Malik', 'Sheikh', 'Yousaf', 'Rehman']
            return f"{random.choice(first_names)} {random.choice(last_names)}"

        def get_valid_name(name):
            if isinstance(name, str) and name.strip():
                return name.strip()
            return generate_random_name()

        def map_bank_name(name):
            if not name or not isinstance(name, str):
                return 'missing'
            name = name.lower()
            bank_mapping = {
                'bank islami': 'BIPL', 'islami': 'BIPL',
                'alfalah': 'BAFL', 'alfla': 'BAFL',
                'meezan': 'MEEZAN', 'meeza': 'MEEZAN',
                'ubl': 'UBL','united bank limited': 'UBL',
                'sadapay': 'SADAPAY', 'sada': 'SADAPAY',
                'silk': 'SILK',
                'hbl': 'HBL',
                'mcb': 'MCB',
                'allied': 'ABL',
                'habibmetro': 'HMB'
            }
            for key, code in bank_mapping.items():
                if key in name:
                    return code
            return 'missing'

        all_dfs = []
        seen_files = set()

        for file in merchant_files:
            if file.name in seen_files:
                st.warning(f"⚠️ Skipping duplicate file: {file.name}")
                continue
            seen_files.add(file.name)
            try:
                df = pd.read_excel(file)
                df['source_file'] = file.name
                all_dfs.append(df)
            except Exception as e:
                st.warning(f"❌ Could not read file {file.name}: {e}")

        if not all_dfs:
            st.error("No valid data found in uploaded files.")
            st.stop()

        combined_df = pd.concat(all_dfs, ignore_index=True)

        # Clean and prepare data
        combined_df['clean_iban'] = combined_df['IBAN'].apply(extract_valid_account)
        combined_df['clean_account'] = combined_df['customerAccount'].apply(extract_valid_account)
        combined_df['clean_amount'] = combined_df['amount'].apply(extract_amount)

        def get_final_account(row):
            return row['clean_iban'] if row['clean_iban'] else (
                row['clean_account'] if row['clean_account'] else (
                    str(row['customerAccount']).strip() if pd.notna(row['customerAccount']) else None
                )
            )

        combined_df['final_account'] = combined_df.apply(get_final_account, axis=1)
        cleaned = combined_df.dropna(subset=['final_account', 'clean_amount'])

        # Build Output
        summary_rows = []
        for _, row in cleaned.iterrows():
            account = row['final_account']
            amount = row['clean_amount']
            bank_raw = row.get('destinationBank', '')
            destination_bank = map_bank_name(bank_raw)
            customer_name = get_valid_name(row.get('customerName', ''))

            summary_rows.append({
                'reference': generate_random_id('REF'),
                'customerReference': generate_random_id('CREF'),
                'merchantId': '2000061',
                'customerName': customer_name,
                'customerContact': account,
                'customerEmail': 'test@gmail.com',
                'customerDob': 'Unknown',
                'customerGender': 'Male',
                'customerAccount': account,
                'accountType': 'BA',
                'destinationBank': destination_bank,
                'amount': amount
            })

        final_df = pd.DataFrame(summary_rows)

        # Display & Download
        st.success(f"✅ Processed {len(final_df)} rows from {len(merchant_files)} file(s).")
        st.dataframe(final_df.head(10))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Updated')
        output.seek(0)

        st.download_button(
            label="📥 Download Updated Excel File",
            data=output,
            file_name="monetix_updated_sheet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
