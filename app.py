import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# 1. 页面配置
st.set_page_config(page_title="平安建议书对比分析", layout="wide")
st.title("🛡️ 平安人寿建议书 vs 银行转存一元对比表")

# 2. 辅助函数：强力清洗数字格式，确保 Excel 不出绿三角
def clean_to_number(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return 0
    # 去掉逗号、空格、人民币符号，只留数字、负号和小数点
    s = re.sub(r'[^-0-9.]', '', str(val).replace('\n', ''))
    try:
        if '.' in s:
            return float(s)
        return int(s)
    except:
        return 0

with st.sidebar:
    st.header("📊 测算参数")
    principal = st.number_input("初始总本金 (元)", value=400000, step=10000)
    bank_rate = st.number_input("银行假定利率 (%)", value=2.5, step=0.1) / 100

uploaded_file = st.file_uploader("上传平安建议书 PDF", type="pdf")

if uploaded_file:
    try:
        with st.spinner('数据校准中，请稍候...'):
            all_raw_rows = []
            with pdfplumber.open(uploaded_file) as pdf:
                # 遍历利益演示核心页
                for page in pdf.pages[10:18]:
                    table = page.extract_table()
                    if table:
                        all_raw_rows.extend(table)
            
            if not all_raw_rows:
                st.error("未能识别到表格。")
            else:
                # --- 核心改进：智能处理表头（解决您说的“标题没了”问题） ---
                # 我们取前4行来分析表头
                header_candidates = all_raw_rows[:4]
                final_headers = []
                for col_idx in range(len(header_rows[0]) if 'header_rows' in locals() else len(all_raw_rows[0])):
                    # 拼接前4行文字，去掉None和重复词
                    parts = []
                    for r_idx in range(min(4, len(all_raw_rows))):
                        cell = str(all_raw_rows[r_idx][col_idx]).replace('\n', '').strip()
                        if cell and cell != 'None' and cell not in parts:
                            parts.append(cell)
                    h_name = "_".join(parts)
                    final_headers.append(h_name if h_name else f"列_{col_idx}")

                # --- 恢复第一版最稳的逻辑：只保留第一列是“保单年度”数字的行 ---
                data_rows = [r for r in all_raw_rows if str(r[0]).strip().isdigit()]
                
                df = pd.DataFrame(data_rows, columns=final_headers)
                
                # --- 强制转换数字格式 ---
                for col in df.columns:
                    df[col] = df[col].apply(clean_to_number)
                
                # --- 动态银行计算 ---
                # 自动寻找保费列（查找标题里带“期交”或“保费”的列）
                p_col = next((c for c in df.columns if '期交' in c or '保费' in c), None)
                
                bank_balances = []
                curr_bank = principal
                # 确保按年度排序
                df = df.sort_values(by=df.columns[0])
                
                for _, row in df.iterrows():
                    # 如果找到了保费列就用实际值，没找到就默认0
                    annual_premium = row[p_col] if p_col else 0
                    # 银行余额计算公式
                    curr_bank = (curr_bank - annual_premium) * (1 + bank_rate)
                    bank_balances.append(round(curr_bank, 2))
                
                # 把银行余额插在第一列
                df.insert(0, '银行账户余额(测算)', bank_balances)

                st.success("✅ 数据重构完成，表头已校准！")
                st.dataframe(df)

                # --- 导出真正的数字格式 Excel (使用 xlsxwriter 强力修复) ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='对比分析')
                    workbook = writer.book
                    worksheet = writer.sheets['对比分析']
                    # 强制设置全表为数字格式，宽度自适应
                    num_format = workbook.add_format({'num_format': '#,##0', 'font_name': '微软雅黑'})
                    worksheet.set_column('A:ZZ', 15, num_format)
                
                st.download_button("📥 下载校准版 Excel (纯数字格式)", data=output.getvalue(), file_name="平安对比分析表.xlsx")

    except Exception as e:
        st.error(f"运行出错：{str(e)}")
