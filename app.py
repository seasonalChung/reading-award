import streamlit as st
import pandas as pd
import io

# 網頁標題與介紹
st.set_page_config(page_title="閱讀領獎名單自動整理工具", layout="wide")
st.title("📚 閱讀領獎名單自動整理系統")
st.markdown("""
這個工具會自動比對 Excel 中的所有分頁，找出**共同出現**的學生，並計算**達標次數**。
- **規則**：任 3 個區間達標（區間本數 $\ge 6$）即具備領獎資格。
""")

# 1. 上傳檔案
uploaded_file = st.file_uploader("請上傳 114.xlsx 檔案", type=["xlsx"])

if uploaded_file:
    # 讀取所有分頁
    all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
    sheet_names = list(all_sheets.keys())
    
    st.info(f"偵測到分頁：{', '.join(sheet_names)}")
    
    # 讓使用者確認姓名欄位名稱
    target_column = st.text_input("請輸入『姓名』欄位的正確名稱", value="姓名")
    
    if st.button("開始執行整理"):
        result_df = None
        
        # 處理資料
        for sheet_name, df in all_sheets.items():
            df = df.dropna(how='all')
            if target_column in df.columns:
                df[target_column] = df[target_column].astype(str).str.strip()
                new_col_name = f"{sheet_name}區間本數"
                if "區間本數" in df.columns:
                    df = df.rename(columns={"區間本數": new_col_name})
                
                keep_cols = [target_column, "班級", "座號"] + ([new_col_name] if new_col_name in df.columns else [])
                current_df = df[[c for c in keep_cols if c in df.columns]].copy()
                
                if result_df is None:
                    result_df = current_df
                else:
                    result_df = pd.merge(result_df, current_df, on=target_column, how='inner')

        if result_df is not None and not result_df.empty:
            # 整理欄位
            class_cols = [c for c in result_df.columns if "班級" in c]
            no_cols = [c for c in result_df.columns if "座號" in c]
            vol_cols = [c for c in result_df.columns if "區間本數" in c]
            
            final_df = pd.DataFrame()
            final_df["班級"] = result_df[class_cols[0]]
            final_df["座號"] = result_df[no_cols[0]]
            final_df["姓名"] = result_df[target_column]
            for col in vol_cols:
                final_df[col] = result_df[col].fillna(0)

            # 領獎邏輯
            def calculate_award(row):
                meet_count = 0
                third_period = "未達標"
                for col in vol_cols:
                    if row[col] >= 6:
                        meet_count += 1
                        if meet_count == 3:
                            third_period = col.replace("區間本數", "")
                return pd.Series([meet_count, "是" if meet_count >= 3 else "否", third_period])

            final_df[["達標區間數", "具備領獎資格", "首度領獎批次"]] = final_df.apply(calculate_award, axis=1)
            final_df = final_df.sort_values(by=["班級", "座號"])

            # 顯示結果預覽
            st.success(f"整理完成！共有 {len(final_df)} 名學生資料。")
            st.dataframe(final_df)

            # 2. 下載按鈕
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 下載整理好的領獎名單",
                data=output.getvalue(),
                file_name="共同姓名整理名單_網頁版.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("找不到符合條件的姓名，請檢查欄位名稱是否正確。")