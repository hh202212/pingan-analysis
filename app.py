import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="平安建议书复刻神器", layout="wide")
st.title("🖨️ 平安建议书“复印级”提取 (V3.9 长表合并版)")
st.info("核心改进：自动识别跨页续接表格 | 剔除中间重复表头 | 保持一方案一长表")

def clean_val(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return ""
    s = str(val).replace('\n', '').replace(' ', '').strip()
    if re.fullmatch(r'^-?[0-9,.]+$', s):
        num_s = re.sub(r'[^-0-9.]', '', s)
        try:
            return float(num_s) if '.' in num_s else int(num_s)
        except: return s
    return s

uploaded_file = st.file_uploader("👉 请上传包含多页利益表的 PDF", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在智能拼接跨页表格，请稍候...'):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                
                with pdfplumber.open(uploaded_file) as pdf:
                    schemes = [] # 存储所有的完整方案
                    current_scheme = []
                    
                    # 1. 深度扫描并执行逻辑拼接
                    for page_idx, page in enumerate(pdf.pages):
                        tables = page.find_tables()
                        if not tables: continue
                        
                        for table_obj in tables:
                            table_data = table_obj.extract()
                            if not table_data: continue
                            
                            # 识别该表的起始行：找到第一个第一列是数字的行
                            data_start_row_idx = -1
                            first_year_val = ""
                            for r_idx, row in enumerate(table_data):
                                cell_0 = str(row[0]).strip()
                                if cell_0.isdigit():
                                    data_start_row_idx = r_idx
                                    first_year_val = cell_0
                                    break
                            
                            # 逻辑判定：
                            # 如果第一年是 "1"，说明是全新方案
                            # 如果第一年 > "1"，且当前已有方案，说明是“续表”
                            is_continuation = False
                            if first_year_val == "1":
                                if current_scheme: schemes.append(current_scheme)
                                current_scheme = {"cells": [], "row_offset": 0, "has_header": False}
                            elif first_year_val != "" and current_scheme:
                                is_continuation = True

                            if current_scheme is not None:
                                # 写入逻辑
                                for cell in table_obj.cells:
                                    r0, c0, r1, c1 = int(cell[0]), int(cell[1]), int(cell[2]), int(cell[3])
                                    
                                    # 如果是续表，且当前单元格属于表头区域（在数据起始行之前），则跳过
                                    if is_continuation and r0 < data_start_row_idx:
                                        continue
                                    
                                    # 计算在长表中的实际行号
                                    # 如果是续表，要把数据往下接，且减去续表自带的表头高度
                                    if is_continuation:
                                        actual_r0 = r0 - data_start_row_idx + current_scheme["row_offset"]
                                        actual_r1 = r1 - data_start_row_idx + current_scheme["row_offset"]
                                    else:
                                        actual_r0 = r0 + current_scheme["row_offset"]
                                        actual_r1 = r1 + current_scheme["row_offset"]
                                    
                                    raw_text = table_data[r0][c0]
                                    # 数据行清洗数字，表头行保留文字
                                    is_data_row = (r0 >= data_start_row_idx)
                                    val = clean_val(raw_text) if is_data_row else str(raw_text).replace('\n', ' ')
                                    
                                    current_scheme["cells"].append({
                                        'r0': actual_r0, 'c0': c0, 'r1': actual_r1, 'c1': c1, 'val': val
                                    })
                                
                                # 更新长表的行偏移量
                                if is_continuation:
                                    current_scheme["row_offset"] += (len(table_data) - data_start_row_idx)
                                else:
                                    current_scheme["row_offset"] += len(table_data)

                    if current_scheme: schemes.append(current_scheme)

                    # 2. 统一写入 Excel
                    for idx, scheme in enumerate(schemes):
                        sheet_name = f"方案_{idx+1}"
                        worksheet = workbook.add_worksheet(sheet_name)
                        written = set()
                        
                        for c in scheme["cells"]:
                            r0, c0, r1, c1, val = c['r0'], c['c0'], c['r1'], c['c1'], c['val']
                            fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                            
                            if r1 - r0 > 1 or c1 - c0 > 1:
                                try: worksheet.merge_range(r0, c0, r1-1, c1-1, val, fmt)
                                except: pass
                                for r in range(r0, r1):
                                    for c_idx in range(c0, c1): written.add((r, c_idx))
                            else:
                                if (r0, c0) not in written:
                                    worksheet.write(r0, c0, val, fmt)
                                    written.add((r0, c0))
                        
                        worksheet.set_column(0, 30, 12)

            st.success(f"🎉 跨页合并完成！已将连续页面整合为 {len(schemes)} 张长表。")
            st.download_button(
                label="📥 下载“长表拼接”Excel 文件",
                data=output.getvalue(),
                file_name="建议书完整长表提取.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"⚠️ 处理出错：{str(e)}")
