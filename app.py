import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="閱讀領獎名單整理", layout="wide")
st.title("📚 閱讀領獎自動化系統 (獲獎名單專用)")

uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx"])

if uploaded_file:
    try:
        file_container = io.BytesIO(uploaded_file.read())
        all_sheets = pd.read_excel(file_container, sheet_name=None, engine='openpyxl')
        
        st.success(f"✅ 成功讀取檔案")
        
        target_col = st.text_input("比對欄位 (姓名)", value="姓名")
        class_col = st.text_input("班級欄位", value="班級")
        no_col = st.text_input("座號欄位", value="座號")

        if st.button("生成獲獎名單"):
            result_df = None
            
            for sheet_name, df in all_sheets.items():
                df = df.dropna(how='all')
                
                if target_col in df.columns:
                    # 統一處理姓名
                    df[target_col] = df[target_col].astype(str).str.strip()
                    
                    # 重新命名區間本數
                    new_vol_name = f"{sheet_name}區間本數"
                    if "區間本數" in df.columns:
                        df = df.rename(columns={"區間本數": new_vol_name})
                    
                    # 挑選該分頁現有的必要欄位
                    keep = [target_col, class_col, no_col]
                    if new_vol_name in df.columns: keep.append(new_vol_name)
                    current_df = df[[c for c in keep if c in df.columns]].copy()
                    
                    if result_df is None:
                        result_df = current_df
                    else:
                        # 使用 outer join 合併
                        result_df = pd.merge(result_df, current_df, on=target_col, how='outer', suffixes=('', '_drop'))
                        
                        # 補齊班級/座號
                        if f"{class_col}_drop" in result_df.columns:
                            result_df[class_col] = result_df[class_col].fillna(result_df[f"{class_col}_drop"])
                            result_df[no_col] = result_df[no_col].fillna(result_df[f"{no_col}_drop"])
                            result_df = result_df.drop(columns=[f"{class_col}_drop", f"{no_col}_drop"])

            if result_df is not None and not result_df.empty:
                # 找出所有本數欄位
                vol_cols = [c for c in result_df.columns if "區間本數" in c]
                for v in vol_cols:
                    result_df[v] = pd.to_numeric(result_df[v], errors='coerce').fillna(0)

                # 確保班級與座號有基礎值
                result_df[class_col] = result_df[class_col].fillna("未知")
                result_df[no_col] = result_df[no_col].fillna(0)

                # 領獎邏輯計算
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
                    return pd.Series([m_count, first_win])

                result_df[["達標次數", "首度領獎批次"]] = result_df.apply(calc_logic, axis=1)
                
                # --- 關鍵修正：只保留達標次數 >= 3 的同學 ---
                winner_df = result_df[result_df["達標次數"] >= 3].copy()
                
                if not winner_df.empty:
                    # 排序與最終輸出欄位 (移除「可領獎」)
                    final_cols = [class_col, no_col, target_col] + vol_cols + ["達標次數", "首度領獎批次"]
                    final_cols = [c for c in final_cols if c in winner_df.columns]
                    winner_df = winner_df[final_cols].sort_values(by=[class_col, no_col])
                    
                    st.write(f"### 🎉 獲獎學生名單 (共 {len(winner_df)} 人)")
                    st.dataframe(winner_df)

                    # 下載按鈕
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        winner_df.to_excel(writer, index=False)
                    
                    st.download_button(
                        label="📥 下載獲獎學生清單",
                        data=output.getvalue(),
                        file_name="114年度閱讀獲獎名單.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("目前沒有學生符合領獎資格 (未有學生達成 3 個區間本數 >= 6)。")
            else:
                st.warning("查無資料。")

    except Exception as e:
        st.error(f"發生錯誤：{e}")
