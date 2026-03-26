import streamlit as st
import pandas as pd
import io

st.title("🎓 淡江行政期末統計工具")

# 1. 上傳對照表 (要有 姓名、學院 兩欄)
lookup_file = st.file_uploader("步驟 1：上傳教師名冊對照表", type=['xlsx'])

# 2. 上傳活動檔 (可以一次選 20 個)
activity_files = st.file_uploader("步驟 2：上傳所有活動 Excel 檔", type=['xlsx', 'xls'], accept_multiple_files=True)

if lookup_file and activity_files:
    try:
        df_lookup = pd.read_excel(lookup_file)
        # 清除空格
        df_lookup.columns = [str(c).strip() for c in df_lookup.columns]
        
        all_names = []
        for f in activity_files:
            # 關鍵：因為你的姓名在第 2 列，我們直接讀取並自動找標題
            df = pd.read_excel(f, header=1) # header=1 代表從第 2 列開始讀
            df.columns = [str(c).strip() for c in df.columns]
            
            if '姓名' in df.columns:
                all_names.append(df[['姓名']].dropna())
        
        if all_names:
            combined = pd.concat(all_names)
            # 串接對照表
            final = pd.merge(combined, df_lookup, on='姓名', how='left')
            
            # 統計「學院」欄位的人次
            col_name = '學院' if '學院' in final.columns else final.columns[1]
            result = final[col_name].value_counts().reset_index()
            result.columns = ['院系', '參與人次']
            
            st.success("統計完成！")
            st.table(result)
            
            # 匯出按鈕
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result.to_excel(writer, index=False)
            st.download_button("📥 下載統計結果", output.getvalue(), "期末統計.xlsx")
    except Exception as e:
        st.error(f"偵測到檔案格式不符，錯誤原因：{e}")
