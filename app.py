import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# 1. 页面配置
st.set_page_config(page_title="平安建议书表格提取工具", layout="wide")
st.title("📋 平安人寿建议书表格提取专家")
st.markdown("---")

# 核心：强制数值转换函数（解决 Excel 绿三角和计算报错）
def force_num(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return 0
    # 彻底去掉逗号、空格、人民币符号、换行符
    clean_val = re.sub(r'[^-0-9.]', '', str(val).replace('\n', ''))
    try:
        if '.' in clean_val:
            return float(clean_val)
        return int(clean_val)
    except:
        return 0

# 2. 文件上传
uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 文件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('🔍 正在深度扫描建议书利益演示表...'):
            all_rows = []
            with pdfplumber.open(uploaded_file) as pdf:
                # 扫描 10-20 页，覆盖绝大多数平安建议书的利益演示范围
                for page in pdf.pages[10:20]:
                    table = page.extract_table()
                    if table:
                        all_rows.extend(table)
            
            if not all_rows:
                st.error("未能识别到表格数据，请确认 PDF 是否为原件。")
            else:
                # --- A. 智能表头识别 ---
                header_row = None
                header_idx = 0
                for i, r in enumerate(all_rows[:15]): # 在前15行里寻找表头
                    if '保单年度' in str(r) or '年龄' in str(r):
                        # 清洗表头文字中的换行符
                        header_row = [str(item).replace('\n', '') if item else f"列_{idx}" for idx, item in enumerate(r)]
                        header_idx = i
                        break
                
                # --- B. 数据行提取 ---
                # 只保留表头下方，且第一列是纯数字（1, 2, 3...）的行
                data_rows = [r for r in all_rows[header_idx+1:] if str(r[0]).strip().isdigit()]
                
                if not data_rows:
                    st.error("识别到了表头，但未抓取到具体的数字行。")
                else:
                    # 组装 DataFrame
                    df = pd.DataFrame(data_rows, columns=header_row if header_row else None)
                    
                    # --- C. 强制全表数值化 (解决绿三角的关键) ---
                    for col in df.columns:
                        df[col] = df[col].apply(force_num)
                    
                    st.success(f"🎉 成功提取 {len(df)} 行数据！")
                    
                    # 3. 页面预览
                    st.dataframe(df, use_container_width=True)

                    # 4. 导出纯数字 Excel
                    output = io.BytesIO()
                    # 使用 xlsxwriter 引擎确保从底层锁定数字格式
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False, sheet_name='利益演示')
                        workbook = writer.book
                        worksheet = writer.sheets['利益演示']
                        
                        # 设置 Excel 格式：数字格式 + 字体对齐
                        num_format = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'font_name': '微软雅黑'})
                        header_format = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1, 'align': 'center'})
                        
                        # 应用格式
                        for col_num, value in enumerate(df.columns.values):
                            worksheet.write(0, col_num, value, header_format)
                            worksheet.set_column(col_num, col_num, 15, num_format)

                    st.download_button(
                        label="📥 点击下载 Excel 文件",
                        data=output.getvalue(),
                        file_name="平安建议书数据提取.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    except Exception as e:
        st.error(f"⚠️ 处理出错：{str(e)}")
