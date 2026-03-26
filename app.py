import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="淡江行政助手", layout="wide")
st.title("🎓 淡江大學-期末院系參與統計工具")

# 側邊欄說明
with st.sidebar:
    st.info("使用說明：\n1. 上傳對照表\n2. 上傳多個活動檔\n3. 點擊下載")

# 1. 上傳對照表
lookup_file = st.file_uploader("1. 請上傳『教師名冊對照表』(xlsx)", type=['xlsx'])

# 2. 上傳活動檔案
activity_files = st.file_uploader("2. 請選取並上傳所有『活動 Excel 檔案』", type=['xlsx', 'xls'], accept_multiple_files=True)

if lookup_file and activity_files:
    try:
        df_lookup = pd.read_excel(lookup_file)
        df_lookup.columns = df_lookup.columns.str.strip()
        
        all_data = []
        for uploaded_file in activity_files:
            # 讀取前 10 列找標題
            df_temp = pd.read_excel(uploaded_file, header=None, nrows=10)
            header_row = 0
            for i, row in df_temp.iterrows():
                if row.astype(str).str.contains('姓名').any():
                    header_row = i
                    break
            
            df_act = pd.read_excel(uploaded_file, header=header_row)
            df_act.columns = df_act.columns.str.strip()
            
            if '姓名' in df_act.columns:
                all_data.append(df_act[['姓名']].dropna())
        
        if all_data:
            df_combined = pd.concat(all_data, ignore_index=True)
            df_final = pd.merge(df_combined, df_lookup, on='姓名', how='left')
            
            # 自動抓取『學院』或對照表第二欄進行統計
            count_col = '學院' if '學院' in df_final.columns else df_final.columns[1]
            result = df_final[count_col].value_counts().reset_index()
            result.columns = [count_col, '參與人次']
            
            st.success("✅ 統計成功！")
            st.table(result)
            
            # 下載按鈕
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result.to_excel(writer, index=False)
            st.download_button(label="📥 下載統計報表", data=output.getvalue(), file_name="期末統計成果.xlsx", mime="application/vnd.ms-excel")
    except Exception as e:
        st.error(f"發生錯誤：{e}")
