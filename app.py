import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="平安建议书复刻神器 V3.5", layout="wide")
st.title("🖨️ 平安建议书“复印级”提取 (V3.5 安全加固版)")

def clean_numeric_val(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return ""
    s = str(val).replace('\n', '').replace(' ', '').strip()
    if re.search(r'\d', s) and not re.search(r'[\u4e00-\u9fa5]', s):
        num_s = re.sub(r'[^-0-9.]', '', s)
        try:
            return float(num_s) if '.' in num_s else int(num_s)
        except: return s
    return s

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 原件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在进行安全扫描与结构对齐...'):
            output = io.BytesIO()
            found_any_table = False
            
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
                
                with pdfplumber.open(uploaded_file) as pdf:
                    for page_idx, page in enumerate(pdf.pages):
                        page_text = page.extract_text() or ""
                        # 只要页面包含关键词，就开始深度尝试
                        if "利益演示表" in page_text or "演示" in page_text:
                            tables = page.find_tables()
                            if not tables: continue
                            
                            for t_idx, table in enumerate(tables):
                                table_data = table.extract()
                                # --- 核心加固：检查表格数据是否有效 ---
                                if not table_data or len(table_data) == 0 or len(table_data[0]) == 0:
                                    continue
                                
                                found_any_table = True
                                sheet_name = f"P{page_idx+1}_T{t_idx+1}"
                                worksheet = workbook.add_worksheet(sheet_name[:31])
                                
                                written_cells = set()
                                for cell in table.cells:
                                    # 强制索引安全化
                                    r0, c0, r1, c1 = int(cell[0]), int(cell[1]), int(cell[2]), int(cell[3])
                                    
                                    # 安全读取，防止越界
                                    if r0 < len(table_data) and c0 < len(table_data[0]):
                                        raw_text = table_data[r0][c0]
                                    else:
                                        continue
                                    
                                    # 识别数据行
                                    first_col_val = str(table_data[r0][0]).strip() if table_data[r0] else ""
                                    is_data = first_col_val.isdigit()
                                    
                                    val = clean_numeric_val(raw_text) if is_data else str(raw_text).replace('\n', ' ')
                                    fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                                    
                                    try:
                                        if r1 - r0 > 1 or c1 - c0 > 1:
                                            worksheet.merge_range(r0, c0, r1-1, c1-1, val, fmt)
                                            for r in range(r0, r1):
                                                for c in range(c0, c1): written_cells.add((r, c))
                                        elif (r0, c0) not in written_cells:
                                            worksheet.write(r0, c0, val, fmt)
                                            written_cells.add((r0, c0))
                                    except:
                                        continue
                                
                                worksheet.set_column(0, 30, 12)

            if not found_any_table:
                st.warning("⚠️ 无法从该 PDF 中提取表格对象。这通常是因为：\n1. 文件是【扫描件】或图片转换的。\n2. 文件被设置了【内容提取限制】。")
            else:
                st.success("🎉 复刻完成！")
                st.download_button(
                    label="📥 下载 Excel 文件",
                    data=output.getvalue(),
                    file_name="建议书提取_加固版.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ 运行中出现未预料的错误: {str(e)}")
