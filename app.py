import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# 1. 页面基础配置
st.set_page_config(page_title="平安建议书复刻神器", layout="wide")
st.title("🖨️ 平安建议书“复印级”提取 (V3.1 识别增强版)")
st.info("核心改进：增强了对‘保单年度/1’合并格子的识别逻辑，确保不漏掉任何产品。")

# 强制数值转换
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

# 判断是否是新的产品起始点（保单年度 1）
def is_new_section_start(cell_content):
    if cell_content is None:
        return False
    # 去掉空格和换行符，处理合并格子里的“保单年度1”
    s = str(cell_content).replace('\n', '').replace(' ', '')
    # 逻辑：只要格子里包含“保单年度”且以“1”结尾，或者是纯数字“1”
    if (("保单年度" in s) and (s.endswith("1"))) or (s == "1"):
        return True
    return False

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 原件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在深度扫描红框区域，请稍候...'):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                # 定义样式
                num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                
                with pdfplumber.open(uploaded_file) as pdf:
                    all_sections = []
                    current_section_cells = []
                    current_row_offset = 0
                    
                    # 扩大扫描页面（通常建议书利益表在10页之后）
                    for page_idx in range(9, min(30, len(pdf.pages))):
                        page = pdf.pages[page_idx]
                        tables = page.find_tables()
                        if not tables: continue
                        
                        for table in tables:
                            table_data = table.extract()
                            if not table_data: continue
                            
                            for r_idx, row_content in enumerate(table_data):
                                # --- 核心逻辑改进：识别合并后的“保单年度1” ---
                                if is_new_section_start(row_content[0]) and current_section_cells:
                                    all_sections.append(current_section_cells)
                                    current_section_cells = []
                                    current_row_offset = 0
                                
                                # 提取当前行的所有 cell 定义
                                cells_in_row = [c for c in table.cells if int(c[0]) == r_idx]
                                # 判断是否是数据行（除了第一年外，后续年份通常是纯数字）
                                is_data = str(row_content[0]).strip().isdigit() or is_new_section_start(row_content[0])
                                
                                for cell in cells_in_row:
                                    r0, c0, r1, c1 = int(cell[0]), int(cell[1]), int(cell[2]), int(cell[3])
                                    raw_text = table_data[r0][c0]
                                    val = clean_val(raw_text) if is_data else str(raw_text)
                                    
                                    current_section_cells.append({
                                        'r0': r0 + current_row_offset, 'c0': c0,
                                        'r1': r1 + current_row_offset, 'c1': c1,
                                        'val': val
                                    })
                            
                            current_row_offset += len(table_data)

                    if current_section_cells:
                        all_sections.append(current_section_cells)

                    # 写入 Excel
                    if not all_sections:
                        st.error("未能识别到‘保单年度1’。请确认PDF页面是否包含您截图中的红框内容。")
                    else:
                        for idx, section in enumerate(all_sections):
                            sheet_name = f"产品利益_{idx + 1}"
                            worksheet = workbook.add_worksheet(sheet_name)
                            written_mark = set()
                            
                            for item in section:
                                r0, c0, r1, c1, val = item['r0'], item['c0'], item['r1'], item['c1'], item['val']
                                fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                                
                                if r1 - r0 > 1 or c1 - c0 > 1:
                                    worksheet.merge_range(r0, c0, r1-1, c1-1, val, fmt)
                                    for r in range(r0, r1):
                                        for c in range(c0, c1): written_mark.add((r, c))
                                else:
                                    if (r0, c0) not in written_mark:
                                        worksheet.write(r0, c0, val, fmt)
                                        written_mark.add((r0, c0))
                            
                            worksheet.set_column(0, 30, 12)

            st.success(f"🎉 识别成功！共为您提取了 {len(all_sections)} 份产品利益表。")
            st.download_button(
                label="📥 下载‘复印级’Excel 文件",
                data=output.getvalue(),
                file_name="平安建议书复刻_识别增强版.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"⚠️ 处理出错：{str(e)}")
