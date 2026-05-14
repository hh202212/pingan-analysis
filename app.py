import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="平安建议书对比工具 V2.0", layout="wide")
st.title("🛡️ 平安建议书一元对比表（精准校准版）")

def clean_val(val):
    """彻底清除None和杂质，返回纯数字"""
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return 0
    # 提取数字、负号和小数点
    s = re.sub(r'[^-0-9.]', '', str(val).replace('\n', ''))
    try:
        return float(s) if '.' in s else int(s)
    except:
        return 0

with st.sidebar:
    st.header("📊 测算参数")
    principal = st.number_input("初始投入总本金 (元)", value=400000, step=10000)
    bank_rate = st.number_input("银行定期利率 (%)", value=2.5, step=0.1) / 100

uploaded_file = st.file_uploader("请上传平安建议书 PDF", type="pdf")

if uploaded_file:
    try:
        with st.spinner('正在执行深度行列对齐，请稍候...'):
            all_data = []
            with pdfplumber.open(uploaded_file) as pdf:
                # 平安利益表通常在10-15页
                for page in pdf.pages[10:16]:
                    table = page.extract_table({
                        "vertical_strategy": "lines", 
                        "horizontal_strategy": "lines",
                        "intersection_y_tolerance": 5
                    })
                    if table:
                        all_data.extend(table)
            
            if not all_data:
                st.error("未能识别表格，请确认PDF是否完整。")
            else:
                # 1. 提取原始数据：只要第一列是纯数字的行
                raw_df = pd.DataFrame(all_data)
                
                # 2. 智能合并前4行作为表头（解决您说的“没有标题”问题）
                header_rows = all_data[:4]
                final_headers = []
                for col_idx in range(len(header_rows[0])):
                    # 拼接前4行非空文字
                    h_parts = [str(header_rows[row][col_idx]).replace('\n', '') 
                               for row in range(4) 
                               if header_rows[row][col_idx] and str(header_rows[row][col_idx]) != 'None']
                    # 去重合并
                    h_name = "_".join(dict.fromkeys(h_parts))
                    final_headers.append(h_name if h_name else f"列_{col_idx}")
                
                # 3. 过滤数据行
                data_rows = [r for r in all_rows if str(r[0]).strip().isdigit()] if 'all_rows' in locals() else []
                # 修正：直接从 raw_df 过滤
                data_df = raw_df[raw_df.iloc[:, 0].astype(str).str.strip().str.isdigit()].copy()
                data_df.columns = final_headers
                
                # 4. 强制清理全表数字（解决 Excel 绿三角）
                for col in data_df.columns:
                    data_df[col] = data_df[col].apply(clean_val)
                
                # 5. 关键：银行复利计算逻辑
                # 自动根据标题定位“期交保费”和“现金价值”
                p_col = next((c for c in data_df.columns if '期交' in c), data_df.columns[2])
                
                bank_balances = []
                curr_bank = principal
                # 按照保单年度排序
                data_df = data_df.sort_values(by=data_df.columns[0])
                
                for _, row in data_df.iterrows():
                    p_val = row[p_col]
                    # 计算逻辑：(上一年余额 - 当年扣费) * 复利
                    curr_bank = (curr_bank - p_val) * (1 + bank_rate)
                    bank_balances.append(round(curr_bank, 2))
                
                data_df.insert(0, '银行账户余额(测算)', bank_balances)

                st.success("🎉 数据校准成功！表头已重构，数字已转码。")
                st.dataframe(data_df)

                # 6. 导出带格式的 Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    data_df.to_excel(writer, index=False, sheet_name='一元对比表')
                    workbook = writer.book
                    worksheet = writer.sheets['一元对比表']
                    # 强制设置数字格式
                    num_fmt = workbook.add_format({'num_format': '#,##0', 'font_name': '微软雅黑'})
                    worksheet.set_column('A:Z', 15, num_fmt)
                
                st.download_button("📥 下载正式版纯数字 Excel", data=output.getvalue(), file_name="平安对比测算表.xlsx")

    except Exception as e:
        st.error(f"处理出错：{str(e)}。请确保上传的是原始电子版PDF。")
