import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="閱讀領獎名單整理", layout="wide")
st.title("📚 閱讀領獎自動化系統 (修正合併錯誤版)")

uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx"])

if uploaded_file:
    try:
        # 讀取檔案內容
        file_container = io.BytesIO(uploaded_file.read())
        all_sheets = pd.read_excel(file_container, sheet_name=None, engine='openpyxl')
        
        st.success(f"✅ 成功讀取檔案")
        
        target_col = st.text_input("比對欄位 (姓名)", value="姓名")
        class_col = st.text_input("班級欄位", value="班級")
        no_col = st.text_input("座號欄位", value="座號")

        if st.button("開始執行整理"):
            result_df = None
            
            for i, (sheet_name, df) in enumerate(all_sheets.items()):
                df = df.dropna(how='all')
                
                if target_col in df.columns:
                    # 統一處理姓名
                    df[target_col] = df[target_col].astype(str).str.strip()
                    
                    # 重新命名區間本數
                    new_vol_name = f"{sheet_name}區間本數"
                    if "區間本數" in df.columns:
                        df = df.rename(columns={"區間本數": new_vol_name})
                    
                    # --- 核心修正：避免 MergeError ---
                    if i == 0:
                        # 第一個分頁：保留所有必要欄位
                        keep = [target_col, class_col, no_col]
                        if new_vol_name in df.columns: keep.append(new_vol_name)
                        result_df = df[[c for c in keep if c in df.columns]].copy()
                    else:
                        # 後續分頁：只保留「姓名」和「區間本數」，排除「班級」和「座號」
                        # 這樣 merge 的時候就不會產生 座號_x, 座號_y
                        keep = [target_col]
                        if new_vol_name in df.columns: keep.append(new_vol_name)
                        
                        current_df = df[[c for c in keep if c in df.columns]].copy()
                        # 執行合併
                        result_df = pd.merge(result_df, current_df, on=target_col, how='inner')
            
            if result_df is not None and not result_df.empty:
                # 找出所有本數欄位
                vol_cols = [c for c in result_df.columns if "區間本數" in c]
                for v in vol_cols:
                    result_df[v] = pd.to_numeric(result_df[v], errors='coerce').fillna(0)

                # 領獎邏輯
                def calc_logic(row):
                    m_count = sum(1 for v in vol_cols if row[v] >= 6)
                    first_win = "未達標"
                    count = 0
                    for v in vol_cols:
                        if row[v] >= 6:
                            count += 1
                            if count == 3:
                                first_win = v.replace("區間本數", "")
                                break
                    return pd.Series([m_count, "是" if m_count >= 3 else "否", first_win])

                result_df[["達標次數", "可領獎", "首度領獎批次"]] = result_df.apply(calc_logic, axis=1)
                
                # 重新排序欄位：班級, 座號, 姓名, ...本數, ...結果
                cols = [class_col, no_col, target_col] + vol_cols + ["達標次數", "可領獎", "首度領獎批次"]
                # 只選取確實存在的欄位避免報錯
                final_df = result_df[[c for c in cols if c in result_df.columns]].sort_values(by=[class_col, no_col])
                
                st.dataframe(final_df)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_df.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 下載 Excel 結果",
                    data=output.getvalue(),
                    file_name="領獎名單整理結果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("查無共同姓名，請檢查各分頁是否都有該學生。")

    except Exception as e:
        st.error(f"發生錯誤：{e}")
        st.info("提示：請確保每個分頁的姓名、班級、座號欄位名稱一致。")
