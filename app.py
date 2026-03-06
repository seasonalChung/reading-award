import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="閱讀領獎名單整理", layout="wide")
st.title("📚 閱讀領獎自動化系統 (全名單合併版)")

uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx"])

if uploaded_file:
    try:
        file_container = io.BytesIO(uploaded_file.read())
        all_sheets = pd.read_excel(file_container, sheet_name=None, engine='openpyxl')
        
        st.success(f"✅ 成功讀取檔案")
        
        target_col = st.text_input("比對欄位 (姓名)", value="姓名")
        class_col = st.text_input("班級欄位", value="班級")
        no_col = st.text_input("座號欄位", value="座號")

        if st.button("開始執行整理"):
            result_df = None
            
            for sheet_name, df in all_sheets.items():
                df = df.dropna(how='all')
                
                if target_col in df.columns:
                    # 統一處理姓名去空格
                    df[target_col] = df[target_col].astype(str).str.strip()
                    
                    # 重新命名區間本數
                    new_vol_name = f"{sheet_name}區間本數"
                    if "區間本數" in df.columns:
                        df = df.rename(columns={"區間本數": new_vol_name})
                    
                    # 挑選必要欄位
                    keep = [target_col, class_col, no_col]
                    if new_vol_name in df.columns: keep.append(new_vol_name)
                    current_df = df[[c for c in keep if c in df.columns]].copy()
                    
                    if result_df is None:
                        result_df = current_df
                    else:
                        # --- 核心修改：改用 outer join (外部合併) ---
                        # 這樣只要在任一分頁出現過的姓名都會被留下來
                        result_df = pd.merge(result_df, current_df, on=target_col, how='outer')
                        
                        # 合併後，若原本沒班級/座號的人，從新分頁補齊
                        if f"{class_col}_y" in result_df.columns:
                            result_df[class_col] = result_df[class_col].fillna(result_df[f"{class_col}_y"])
                            result_df[no_col] = result_df[no_col].fillna(result_df[f"{no_col}_y"])
                            # 移除多餘的暫存欄位
                            result_df = result_df.drop(columns=[f"{class_col}_y", f"{no_col}_y", f"{class_col}_x", f"{no_col}_x"], errors='ignore')
                        elif f"{class_col}_x" in result_df.columns:
                             # 處理只有 x 的情況
                             result_df = result_df.rename(columns={f"{class_col}_x": class_col, f"{no_col}_x": no_col})

            if result_df is not None and not result_df.empty:
                # 找出所有本數欄位並處理數值
                vol_cols = [c for c in result_df.columns if "區間本數" in c]
                for v in vol_cols:
                    result_df[v] = pd.to_numeric(result_df[v], errors='coerce').fillna(0)

                # 確保班級座號沒有空值（避免排序出錯）
                result_df[class_col] = result_df[class_col].fillna("未註記")
                result_df[no_col] = result_df[no_col].fillna(0)

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
                
                # 排序與選取
                cols = [class_col, no_col, target_col] + vol_cols + ["達標次數", "可領獎", "首度領獎批次"]
                final_df = result_df[[c for c in cols if c in result_df.columns]].sort_values(by=[class_col, no_col])
                
                st.write(f"### 總名單預覽 (共 {len(final_df)} 人)")
                st.dataframe(final_df)

                # 提供下載
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_df.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 下載完整領獎名單",
                    data=output.getvalue(),
                    file_name="全名單領獎整理結果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("檔案中沒有資料。")

    except Exception as e:
        st.error(f"發生錯誤：{e}")
