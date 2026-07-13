import streamlit as st
import pandas as pd
import io

def parse_taiwanese_date(date_str):
    """Converts 'YYY年MM月' to standard 'YYYY/MM/DD' datetime format."""
    if pd.isna(date_str):
        return pd.NaT
    
    date_str = str(date_str).strip()
    try:
        if '年' in date_str and '月' in date_str:
            year_str, month_str = date_str.split('年')
            greg_year = int(year_str) + 1911 
            month = month_str.replace('月', '').strip().zfill(2)
            return pd.to_datetime(f"{greg_year}-{month}-01")
    except Exception:
        pass
    return pd.NaT

def process_data(files):
    df_list = []
    
    for file in files:
        try:
            df = pd.read_csv(file, encoding='utf-8-sig')
        except UnicodeDecodeError:
            file.seek(0)
            df = pd.read_csv(file, encoding='big5')
            
        df_list.append(df)

    combined_df = pd.concat(df_list, ignore_index=True)

    # 1. Delete unnecessary columns
    cols_to_drop = ['中文貨名', '英文貨名']
    combined_df = combined_df.drop(columns=[col for col in cols_to_drop if col in combined_df.columns])

    # 2. Reorder columns
    final_order = ['進出口別', '日期', '貨品號列', '國家', '重量(公斤)', '新臺幣(千元)']
    existing_cols = [col for col in final_order if col in combined_df.columns]
    combined_df = combined_df[existing_cols]

    # 3. Replace '進口' and '出口'
    if '進出口別' in combined_df.columns:
        combined_df['進出口別'] = combined_df['進出口別'].replace({'進口': 'IMPORT', '出口': 'EXPORT'})

    # 4. Process Dates
    if '日期' in combined_df.columns:
        combined_df['_sort_date'] = combined_df['日期'].apply(parse_taiwanese_date)
        combined_df['日期'] = combined_df['_sort_date'].dt.strftime('%Y/%m/%d')
    else:
        combined_df['_sort_date'] = pd.NaT

    # 5. Force '貨品號列' to be recognized as a Number 
    if '貨品號列' in combined_df.columns:
        combined_df['貨品號列'] = pd.to_numeric(combined_df['貨品號列'], errors='coerce').astype('Int64')

    # 6. Sort Hierarchy
    if '進出口別' in combined_df.columns:
        combined_df['進出口別'] = pd.Categorical(
            combined_df['進出口別'], 
            categories=['IMPORT', 'EXPORT'], 
            ordered=True
        )

    combined_df = combined_df.sort_values(by=['_sort_date', '進出口別'])
    combined_df = combined_df.drop(columns=['_sort_date'])

    return combined_df

# --- STREAMLIT UI ---
st.set_page_config(page_title="CSV Data Formatter", layout="centered")

st.title("📊 Trade Data Formatter")
st.markdown("Upload your CSV files to automatically merge, reorganize, format dates, and output a clean Excel file.")

uploaded_files = st.file_uploader("Upload CSV files", type=["csv"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Process Files"):
        with st.spinner("Processing data..."):
            try:
                final_df = process_data(uploaded_files)
                
                st.success("Data processed successfully!")
                st.write("### Data Preview")
                st.dataframe(final_df.head(10))
                
                # Convert to Excel in memory using openpyxl
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='Data')
                
                processed_data = output.getvalue()
                
                # Download Button for Excel (.xlsx)
                st.download_button(
                    label="📥 Download Formatted Excel File",
                    data=processed_data,
                    file_name="Formatted_Trade_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"An error occurred while processing the files: {e}")