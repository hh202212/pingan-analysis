import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="平安建议书对比分析", layout="wide")
st.title("🛡️ 平安人寿建议书 vs 银行转存一元对比表")

def clean_num(val):
    """强力清洗数字，处理None、空格、逗号"""
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return 0
    s = re.sub(r'[^-0-9.]', '', str(val))
    try:
        return float(s) if '.' in s else int(s)
    except:
        return 0

with st.sidebar:
    st.header("📊 测算参数")
    principal = st.number_input("初始投入总本金 (元)", value=400000, step=10000)
    bank_rate = st.number_input("银行定期假定利率 (%)", value=2.5, step=0.1) / 100
    st.info("💡 提示：本工具会自动匹配盛世金越建议书中的利益表。")

uploaded_file = st.file_uploader("请上传平安建议书 PDF 文件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('正在进行深度数据校准...'):
            all_rows = []
            with pdfplumber.open(uploaded_file) as pdf:
                # 扩大扫描范围到 10-25 页，确保不漏掉利益表
                for page in pdf.pages[10:25]: 
                    table = page.extract_table()
                    if table:
                        # 核心修正：过滤掉完全为空的行，防止报“0”错误
                        valid_rows = [r for r in table if r and len(r) > 0 and r[0] is not None]
                        all_rows.extend(valid_rows)
            
            if not all_rows:
                st.error("未能在 PDF 中找到表格，请确认文件是否完整。")
            else:
                # 1. 识别表头：寻找包含“年度”或“年龄”的行
                header_row = None
                for r in all_rows:
                    if '保单年度' in str(r) or '年龄' in str(r):
                        header_row = [str(i).replace('\n', '') for i in r]
                        break
                
                # 2. 识别数据行：第一列必须是纯数字（如 1, 2, 3...）
                data_rows = [r for r in all_rows if str(r[0]).strip().isdigit()]
                
                if not data_rows:
                    st.error("识别到了表头，但没找到具体的利益演示数据。")
                else:
                    # 如果没找到表头，就用默认列名
                    if not header_row:
                        header_row = [f"列_{i}" for i in range(len(data_rows[0]))]
                    
                    # 3. 组装数据，确保列数匹配
                    df = pd.DataFrame(data_rows)
                    if df.shape[1] > len(header_row):
                        df = df.iloc[:, :len(header_row)]
                    df.columns = header_row[:df.shape[1]]
                    
                    # 4. 强制数字格式化
                    for col in df.columns:
                        df[col] = df[col].apply(clean_num)
                    
                    # 5. 自动定位关键列：保费和现金价值
                    # 适配盛世金越可能的各种列名
                    premium_keywords = ['期交', '保费', '转入', '支付']
                    p_col = next((c for c in df.columns if any(k in str(c) for k in premium_keywords)), None)
                    
                    # 6. 银行动态复利计算
                    bank_balances = []
                    curr_bank = principal
                    df = df.sort_values(by=df.columns[0]) # 按年度排序
                    
                    for i, row in df.iterrows():
                        # 获取当年保费，如果找不到列则按0计算（盛世金越后期无保费）
                        p_val = row[p_col] if p_col else 0
                        # 核心公式：(银行余额 - 当年保费) * (1 + 利率)
                        curr_bank = (curr_bank - p_val) * (1 + bank_rate)
                        bank_balances.append(round(curr_bank, 2))
                    
                    df.insert(0, '银行账户余额(测算)', bank_balances)

                    st.success("✅ 数据重构完成！Excel 已准备好。")
                    st.dataframe(df)

                    # 7. 导出 Excel（含数字格式强制修复）
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False, sheet_name='对比分析')
                        workbook = writer.book
                        worksheet = writer.sheets['对比分析']
                        # 强制 Excel 单元格为货币/数字格式，消除绿三角
                        num_format = workbook.add_format({'num_format': '#,##0'})
                        worksheet.set_column('A:Z', 12, num_format)

                    st.download_button("📥 点击下载纯数字 Excel", data=output.getvalue(), file_name="平安对比分析结果.xlsx")

    except Exception as e:
        # 增加更详细的错误提示，方便咱们排查
        st.error(f"处理出错：{str(e)}。请检查 PDF 页面是否包含标准的利益演示表。")
