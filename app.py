import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# 页面基础配置
st.set_page_config(page_title="平安建议书对比工具", layout="wide")
st.title("🛡️ 平安人寿建议书 vs 银行转存一元对比表")

# 强制数值转换函数
def force_num(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return 0
    # 彻底去掉逗号、空格、人民币符号等文字干扰
    clean_val = re.sub(r'[^-0-9.]', '', str(val).replace('\n', ''))
    try:
        return float(clean_val) if '.' in clean_val else int(clean_val)
    except:
        return 0

with st.sidebar:
    st.header("📊 测算参数")
    principal = st.number_input("初始投入总本金 (元)", value=400000, step=10000)
    bank_rate = st.number_input("银行定期假定利率 (%)", value=2.5, step=0.1) / 100

uploaded_file = st.file_uploader("请上传平安建议书 PDF 文件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('正在识别表头并清洗数字格式...'):
            all_rows = []
            with pdfplumber.open(uploaded_file) as pdf:
                # 遍历建议书利益页
                for page in pdf.pages[10:18]:
                    table = page.extract_table()
                    if table:
                        all_rows.extend(table)
            
            if not all_rows:
                st.error("未能识别到表格数据。")
            else:
                # --- 1. 识别表头：寻找包含“年度”或“年龄”的一行作为标题 ---
                header_row = None
                header_idx = 0
                for i, r in enumerate(all_rows):
                    if '保单年度' in str(r) or '年龄' in str(r):
                        header_row = [str(item).replace('\n', '') if item else f"列_{idx}" for idx, item in enumerate(r)]
                        header_idx = i
                        break
                
                # --- 2. 提取数据行：只保留表头下方且第一列是数字的行 ---
                data_rows = [r for r in all_rows[header_idx+1:] if str(r[0]).strip().isdigit()]
                
                if not data_rows:
                    st.error("识别到了表头，但没找到具体的利益演示数据。")
                else:
                    # 组装表格
                    df = pd.DataFrame(data_rows, columns=header_row if header_row else None)
                    
                    # --- 3. 核心需求：对数据行执行强制数值转换 ---
                    for col in df.columns:
                        df[col] = df[col].apply(force_num)
                    
                    # 4. 银行动态计算逻辑（保持原始逻辑）
                    bank_flows = []
                    curr_bank = principal
                    for i, row in df.iterrows():
                        # 这里默认拿第3列作为保费支出，如果您的PDF保费在其他列，可以微调
                        out_flow = row.iloc[2] if len(row) > 2 else 0 
                        curr_bank = (curr_bank - out_flow) * (1 + bank_rate)
                        bank_balances.append(round(curr_bank, 2))
                    
                    df.insert(0, '银行账户余额(测算)', bank_balances)

                    st.success("🎉 数据识别成功，已保留表头并转为纯数字！")
                    st.dataframe(df)

                    # 5. 导出 Excel（含数字格式强制修复）
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False, sheet_name='对比分析')
                        workbook = writer.book
                        worksheet = writer.sheets['对比分析']
                        # 强制 Excel 单元格为数字格式，消除绿三角
                        num_format = workbook.add_format({'num_format': '#,##0'})
                        worksheet.set_column('A:Z', 15, num_format)

                    st.download_button("📥 下载带表头的纯数字 Excel", data=output.getvalue(), file_name="平安对比分析结果.xlsx")

    except Exception as e:
        st.error(f"处理出错：{str(e)}")
