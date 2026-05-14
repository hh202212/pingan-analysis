import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="平安建议书对比分析", layout="wide")
st.title("🛡️ 平安人寿建议书 vs 银行转存一元对比表")

# 辅助函数：纯净地提取数字
def clean_to_numeric(val):
    if val is None or str(val).strip() == "":
        return 0
    # 只保留负号、数字和小数点
    s = re.sub(r'[^-0-9.]', '', str(val))
    try:
        return float(s) if '.' in s else int(s)
    except:
        return 0

with st.sidebar:
    st.header("📊 测算参数")
    principal = st.number_input("初始投入总本金 (元)", value=400000, step=10000)
    bank_rate = st.number_input("银行定期假定利率 (%)", value=2.5, step=0.1) / 100
    st.info("💡 提示：建议书通常在第10-18页包含利益表。")

uploaded_file = st.file_uploader("请上传平安建议书 PDF 文件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('正在重新校准数据并清洗 None 值...'):
            all_rows = []
            with pdfplumber.open(uploaded_file) as pdf:
                # 遍历可能的页面
                for page in pdf.pages[10:20]: 
                    table = page.extract_table()
                    if table:
                        all_rows.extend(table)
            
            if not all_rows:
                st.error("未能识别到表格数据。")
            else:
                # 1. 过滤：只保留第一列是纯数字的行（真正的利益演示行）
                data_rows = [r for r in all_rows if str(r[0]).isdigit()]
                
                # 2. 找表头：在前面的行里找包含“保单年度”的那一行作为标题
                header_row = ["列_" + str(i) for i in range(len(all_rows[0]))] # 默认标题
                for r in all_rows[:15]:
                    if '保单年度' in str(r[0]):
                        header_row = [str(i).replace('\n', '') for i in r]
                        break
                
                # 3. 组装表格
                df = pd.DataFrame(data_rows, columns=header_row)
                
                # 4. 强制数字转换（解决 Excel 绿三角和 0 值的关键）
                for col in df.columns:
                    df[col] = df[col].apply(clean_to_numeric)
                
                # 5. 自动定位“期交保费”和“年度末现金价值”
                # 我们通过关键词匹配，防止因为 PDF 版本不同导致的列号变动
                premium_col = next((c for c in df.columns if '期交' in c or '保费' in c), None)
                cash_val_col = next((c for c in df.columns if '现金价值' in c), None)

                # 6. 银行动态复利计算
                bank_balances = []
                curr_bank = principal
                # 确保按年份排序，从第1年开始算
                df = df.sort_values(by=df.columns[0])
                
                for i, row in df.iterrows():
                    p = row[premium_col] if premium_col else 0
                    # 银行余额 = (剩余钱 - 当年保费) * (1 + 利率)
                    curr_bank = (curr_bank - p) * (1 + bank_rate)
                    bank_balances.append(round(curr_bank, 2))
                
                df.insert(0, '银行账户余额(测算)', bank_balances)

                st.success("🎉 数据校准成功！已剔除所有杂质行。")
                st.dataframe(df)

                # 7. 导出纯数字 Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='对比分析')
                    # 这里可以增加一步 Excel 格式设置，确保是数字格式
                    workbook = writer.book
                    worksheet = writer.sheets['对比分析']
                    num_format = workbook.add_format({'num_format': '#,##0.00'})
                    worksheet.set_column('A:Z', 12, num_format)

                st.download_button("📥 下载校准版纯数字 Excel", data=output.getvalue(), file_name="平安对比分析_校准版.xlsx")

    except Exception as e:
        st.error(f"处理出错：{str(e)}")
