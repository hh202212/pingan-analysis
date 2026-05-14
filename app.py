import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="平安建议书对比分析", layout="wide")
st.title("🛡️ 平安人寿建议书 vs 银行转存对比分析")

# 核心：这个函数负责干掉所有非数字杂质，解决“绿三角”
def to_pure_number(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return 0
    # 只提取数字、负号和小数点
    s = re.sub(r'[^-0-9.]', '', str(val))
    try:
        return float(s) if '.' in s else int(s)
    except:
        return 0

with st.sidebar:
    st.header("📊 测算参数")
    principal = st.number_input("初始总本金 (元)", value=400000, step=10000)
    bank_rate = st.number_input("银行定期利率 (%)", value=2.5, step=0.1) / 100

uploaded_file = st.file_uploader("上传平安建议书 PDF", type="pdf")

if uploaded_file:
    try:
        with st.spinner('正在提取数据并强制转换数字格式...'):
            all_data = []
            with pdfplumber.open(uploaded_file) as pdf:
                # 范围稍微大一点，确保抓全
                for page in pdf.pages[10:20]:
                    tables = page.extract_tables()
                    for table in tables:
                        if table:
                            df_tmp = pd.DataFrame(table)
                            all_data.append(df_tmp)
            
            if not all_data:
                st.error("未发现表格。")
            else:
                # 拼接所有表格
                full_df = pd.concat(all_data, ignore_index=True)
                
                # --- 恢复第一版逻辑：只保留第一列是数字的行 ---
                # 先把第一列转成字符串处理
                full_df = full_df[full_df.iloc[:, 0].astype(str).str.contains(r'^\d+$', na=False)]
                
                # --- 核心修复：强制全表数字转换 ---
                for col in full_df.columns:
                    full_df[col] = full_df[col].apply(to_pure_number)
                
                # 给列起个简单的名字，防止 None 报错
                full_df.columns = [f"列_{i}" for i in range(full_df.shape[1])]
                
                # 银行计算逻辑（假设第3列是保费，您可以根据实际调整列号）
                bank_balances = []
                curr_bank = principal
                # 尝试定位保费列：通常是累计保费前的那一列，这里默认拿第3列(索引2)
                p_idx = 2 
                
                for i, row in full_df.iterrows():
                    p_val = row.iloc[p_idx]
                    curr_bank = (curr_bank - p_val) * (1 + bank_rate)
                    bank_balances.append(round(curr_bank, 2))
                
                full_df.insert(0, '银行账户余额(测算)', bank_balances)

                st.success("✅ 数据已提取，已转为纯数字格式。")
                st.dataframe(full_df)

                # --- 导出真正的数字格式 Excel ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    full_df.to_excel(writer, index=False, sheet_name='对比表')
                    workbook = writer.book
                    worksheet = writer.sheets['对比表']
                    # 关键：从底层强制设置 Excel 单元格格式为数字
                    num_format = workbook.add_format({'num_format': '#,##0'})
                    worksheet.set_column('A:Z', 15, num_format)

                st.download_button("📥 下载纯数字 Excel (无绿三角)", data=output.getvalue(), file_name="平安对比分析.xlsx")

    except Exception as e:
        st.error(f"处理出错：{str(e)}。建议刷新页面重新上传。")
