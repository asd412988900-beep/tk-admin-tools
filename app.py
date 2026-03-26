import streamlit as st
import pandas as pd
import io

st.title("淡江大學-期末院系參與統計工具")
st.write("請先上傳對照表，再上傳活動檔案，系統會自動幫您統計並產出 Excel。")

# 1. 上傳對照表
lookup_file = st.file_uploader("1. 請上傳『教師名冊對照表』", type=['xlsx'])

# 2. 上傳活動檔案 (支援多選)
activity_files = st.file_uploader("2. 請選取並上傳那 20 幾個『活動 Excel 檔案』", type=['xlsx', 'xls'], accept_multiple_files=True)

if lookup_file and activity_files:
    df_lookup = pd.read_excel(lookup_file)
    df_lookup.columns = df_lookup.columns.str.strip()
    
    all_data = []
    for uploaded_file in activity_files:
        # 自動偵測標題列
        df_sample = pd.read_excel(uploaded_file, nrows=5, header=None)
        header_row = 0
        for i, row in df_sample.iterrows():
            if row.astype(str).str.contains('姓名').any():
                header_row = i
                break
        
        df_activity = pd.read_excel(uploaded_file, header=header_row)
        df_activity.columns = df_activity.columns.str.strip()
        
        if '姓名' in df_activity.columns:
            all_data.append(df_activity[['姓名']].dropna())

    if all_data:
        df_combined = pd.concat(all_data, ignore_index=True)
        df_final = pd.merge(df_combined, df_lookup, on='姓名', how='left')
        
        # 統計
        target_col = '學院' if '學院' in df_final.columns else df_final.columns[1]
        result = df_final[target_col].value_counts().reset_index()
        result.columns = [target_col, '參與人次']
        
        st.success("統計完成！")
        st.dataframe(result) # 在網頁上預覽結果
        
        # 匯出下載按鈕
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result.to_excel(writer, index=False)
        st.download_button(label="📥 下載期末統計總表", data=output.getvalue(), file_name="期末院系參與統計總表.xlsx")
