import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="平安建议书对比工具", layout="wide")
st.title("🛡️ 平安人寿建议书 vs 银行转存一元对比表")

def force_numeric(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return 0
    clean_val = re.sub(r'[^-0-9.]', '', str(val))
    try:
        return float(clean_val) if '.' in clean_val else int(clean_val)
    except:
        return 0

with st.sidebar:
    st.header("📊 测算参数")
    principal = st.number_input("初始投入总本金 (元)", value=400000, step=10000)
    bank_rate = st.number_input("银行定期假定利率 (%)", value=2.5, step=0.1) / 100
    st.info("注：程序会自动提取建议书中的实际保费进行计算。")

uploaded_file = st.file_uploader("请上传平安建议书 PDF 文件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('正在执行智能表头重构与数据清洗...'):
            with pdfplumber.open(uploaded_file) as pdf:
                all_pages_data = []
                # 遍历利益演示核心页
                for page in pdf.pages[10:18]:
                    table = page.extract_table()
                    if table:
                        df_raw = pd.DataFrame(table)
                        # --- 核心逻辑：拉平多层表头 ---
                        header_rows = 3 # 平安通常前3行是表头
                        flat_headers = []
                        for col_idx in range(df_raw.shape[1]):
                            # 提取每一列前3行的文字，去掉 None，去重合并
                            parts = [str(df_raw.iloc[row, col_idx]).replace('\n', '') 
                                     for row in range(header_rows) 
                                     if df_raw.iloc[row, col_idx] and str(df_raw.iloc[row, col_idx]) != 'None']
                            # 拼接成：父标题_子标题
                            header_name = "_".join(dict.fromkeys(parts))
                            flat_headers.append(header_name)
                        
                        # 重新设置表头并删除原始表头行
                        df_raw.columns = flat_headers
                        df_clean = df_raw.iloc[header_rows:].copy()
                        
                        # --- 脏数据过滤：只保留“保单年度”是数字的行 ---
                        # 假设第一列是保单年度
                        df_clean = df_clean[df_clean.iloc[:, 0].apply(lambda x: str(x).isdigit())]
                        all_pages_data.append(df_clean)
                
                if not all_pages_data:
                    st.error("未能识别到有效数据。")
                else:
                    final_df = pd.concat(all_pages_data, ignore_index=True)
                    
                    # 转换所有列为数字格式
                    for col in final_df.columns:
                        final_df[col] = final_df[col].apply(force_numeric)
                    
                    # 动态识别保费列（查找包含“期交”或“保费”字眼的列）
                    premium_col = [c for c in final_df.columns if '期交' in c or '保费' in c]
                    
                    # 银行动态计算
                    bank_balances = []
                    curr_bank = principal
                    for i, row in final_df.iterrows():
                        # 自动获取该年份的实际保费，如果没找到则默认5万
                        actual_premium = row[premium_col[0]] if premium_col else 50000
                        curr_bank = (curr_bank - actual_premium) * (1 + bank_rate)
                        bank_balances.append(round(curr_bank, 2))
                    
                    final_df['银行账户余额'] = bank_balances
                    
                    st.success("重构完成！表头已拉平，数据已清洗。")
                    st.dataframe(final_df)
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        final_df.to_excel(writer, index=False, sheet_name='对比分析')
                    st.download_button("📥 下载完美版 Excel", data=output.getvalue(), file_name="平安对比分析_正式版.xlsx")
    except Exception as e:
        st.error(f"处理出错：{str(e)}")
