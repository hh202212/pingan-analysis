import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="平安建议书复刻神器", layout="wide")
st.title("🖨️ 平安建议书“复印级”提取 (V3.7 逻辑对齐版)")
st.info("核心逻辑：识别左上角‘保单年度’作为整表起始 | 自动关联表头与内容 | 纯数字无损转换")

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

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 原件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在识别表格逻辑结构，请稍候...'):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                
                with pdfplumber.open(uploaded_file) as pdf:
                    current_worksheet = None
                    current_row_offset = 0
                    table_count = 0
                    
                    # 遍历全书，寻找利益演示表
                    for page_idx, page in enumerate(pdf.pages):
                        tables = page.find_tables()
                        if not tables: continue
                        
                        for table_obj in tables:
                            table_data = table_obj.extract()
                            if not table_data: continue
                            
                            # --- 关键判定：左上角第一个单元格是否包含“保单年度” ---
                            top_left_cell = str(table_data[0][0]).replace('\n', '').strip()
                            
                            if "保单年度" in top_left_cell:
                                # 发现新表起始：新建一个子表 (Sheet)
                                table_count += 1
                                sheet_name = f"利益演示方案_{table_count}"
                                current_worksheet = workbook.add_worksheet(sheet_name[:31])
                                current_row_offset = 0 # 重置行偏移量
                            
                            # 如果已经确定了当前的操作子表，则开始写入内容
                            if current_worksheet:
                                written_mark = set()
                                # 物理还原合并单元格结构
                                for cell in table_obj.cells:
                                    r0, c0, r1, c1 = int(cell[0]), int(cell[1]), int(cell[2]), int(cell[3])
                                    
                                    try:
                                        raw_text = table_data[r0][c0]
                                    except IndexError: continue
                                    
                                    # 判定是否为数据行（第一列是数字，或者是合并格子里的数字1）
                                    first_col_val = str(table_data[r0][0]).strip()
                                    is_data_row = any(char.isdigit() for char in first_col_val) and "保单年度" not in first_col_val
                                    
                                    val = clean_val(raw_text) if is_data_row else str(raw_text).replace('\n', ' ')
                                    fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                                    
                                    # 写入 Excel，注意要加上 current_row_offset 实现跨页连表
                                    ex_r0, ex_r1 = r0 + current_row_offset, r1 + current_row_offset
                                    
                                    if ex_r1 - ex_r0 > 1 or c1 - c0 > 1:
                                        # 合并写入
                                        try:
                                            current_worksheet.merge_range(ex_r0, c0, ex_r1 - 1, c1 - 1, val, fmt)
                                        except: pass # 防止重叠报错
                                        for r in range(ex_r0, ex_r1):
                                            for c in range(c0, c1): written_mark.add((r, c))
                                    else:
                                        # 普通写入
                                        if (ex_r0, c0) not in written_mark:
                                            current_worksheet.write(ex_r0, c0, val, fmt)
                                            written_mark.add((ex_r0, c0))
                                
                                # 更新下一张续表的起始位置
                                current_row_offset += len(table_data)
                                current_worksheet.set_column(0, 30, 12)

            if table_count == 0:
                st.warning("⚠️ 未能识别到以‘保单年度’开头的利益表，请检查PDF内容。")
            else:
                st.success(f"🎉 提取成功！已为您整合了 {table_count} 张完整的利益演示表。")
                st.download_button(
                    label="📥 下载“逻辑对齐”Excel 文件",
                    data=output.getvalue(),
                    file_name="平安建议书完整复刻.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"⚠️ 处理出错：{str(e)}")
