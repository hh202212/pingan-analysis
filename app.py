import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="平安建议书复刻工具 V3.0", layout="wide")
st.title("🖨️ 平安建议书“复印级”提取 (按年度分组)")
st.info("核心逻辑：按‘保单年度1’自动分表 + 完美还原合并单元格 + 纯数字无损转换")

def clean_val(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return ""
    s = str(val).replace('\n', '').strip()
    if re.fullmatch(r'^-?[0-9,.]+$', s):
        num_s = re.sub(r'[^-0-9.]', '', s)
        try:
            return float(num_s) if '.' in num_s else int(num_s)
        except: return s
    return s

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 原件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在进行全量深度扫描与结构对齐...'):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                
                with pdfplumber.open(uploaded_file) as pdf:
                    all_sections = []
                    current_section_cells = []
                    current_row_offset = 0
                    
                    # 1. 深度遍历所有页面，提取所有单元格并处理“分表”逻辑
                    for page_idx in range(10, min(25, len(pdf.pages))):
                        page = pdf.pages[page_idx]
                        tables = page.find_tables()
                        if not tables: continue
                        
                        for table in tables:
                            table_data = table.extract()
                            if not table_data: continue
                            
                            # 检查该表是否开启了新的“保单年度 1”
                            for r_idx, row_content in enumerate(table_data):
                                first_col = str(row_content[0]).strip()
                                # 如果在第一列发现数字 1，且当前已经有数据，则存入上一节
                                if first_col == "1" and current_section_cells:
                                    all_sections.append(current_section_cells)
                                    current_section_cells = []
                                    current_row_offset = 0
                                
                                # 将该行所有的单元格信息转换并存入当前节
                                # 找到属于这一行的所有 cell
                                cells_in_row = [c for c in table.cells if int(c[0]) == r_idx]
                                for cell in cells_in_row:
                                    r0, c0, r1, c1 = int(cell[0]), int(cell[1]), int(cell[2]), int(cell[3])
                                    raw_text = table_data[r0][c0]
                                    # 如果是数据行（年度1-105），执行数字清洗
                                    is_data = first_col.isdigit()
                                    val = clean_val(raw_text) if is_data else str(raw_text).replace('\n', '')
                                    
                                    # 存储相对于当前 Sheet 起始位置的坐标
                                    current_section_cells.append({
                                        'r0': r0 + current_row_offset,
                                        'c0': c0,
                                        'r1': r1 + current_row_offset,
                                        'c1': c1,
                                        'val': val
                                    })
                            
                            # 更新下一张表的行偏移量（保持在同一个 Sheet 里的连续性）
                            current_row_offset += len(table_data)

                    # 存入最后一节
                    if current_section_cells:
                        all_sections.append(current_section_cells)

                    # 2. 将分组后的数据写入 Excel
                    if not all_sections:
                        st.error("未能在指定范围内识别到有效的利益演示表。")
                    else:
                        for idx, section in enumerate(all_sections):
                            sheet_name = f"产品利益方案_{idx + 1}"
                            worksheet = workbook.add_worksheet(sheet_name)
                            written_mark = set()
                            
                            for item in section:
                                r0, c0, r1, c1, val = item['r0'], item['c0'], item['r1'], item['c1'], item['val']
                                fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                                
                                if r1 - r0 > 1 or c1 - c0 > 1:
                                    # 完美还原合并
                                    worksheet.merge_range(r0, c0, r1-1, c1-1, val, fmt)
                                    for r in range(r0, r1):
                                        for c in range(c0, c1): written_mark.add((r, c))
                                else:
                                    if (r0, c0) not in written_mark:
                                        worksheet.write(r0, c0, val, fmt)
                                        written_mark.add((r0, c0))
                            
                            worksheet.set_column(0, 30, 12)

            st.success(f"🎉 复刻完成！共识别到 {len(all_sections)} 组产品利益演示。")
            st.download_button(
                label="📥 下载“年度分组”复印级 Excel",
                data=output.getvalue(),
                file_name="平安建议书数据复刻_分组版.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"⚠️ 处理出错：{str(e)}")
