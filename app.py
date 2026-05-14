import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="平安建议书复刻 V4.0", layout="wide")
st.title("🖨️ 平安建议书“复印级”长表提取 V4.0")
st.info("已修复崩溃错误 | 增强跨页对齐 | 强制数字转换")

# 强制数值转换
def clean_numeric(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return 0
    s = str(val).replace('\n', '').replace(' ', '').strip()
    if re.search(r'\d', s):
        num_s = re.sub(r'[^-0-9.]', '', s)
        try:
            return float(num_s) if '.' in num_s else int(num_s)
        except: return 0
    return 0

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 原件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在扫描全书并自动拼接长表...'):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                
                with pdfplumber.open(uploaded_file) as pdf:
                    schemes = [] 
                    active_scheme = None # 修复：初始为 None
                    
                    # 遍历利益演示核心页（10-30页）
                    for page_idx in range(9, min(35, len(pdf.pages))):
                        page = pdf.pages[page_idx]
                        tables = page.find_tables()
                        if not tables: continue
                        
                        for table_obj in tables:
                            table_data = table_obj.extract()
                            if not table_data: continue
                            
                            # 1. 寻找数字行起始位置
                            data_start_idx = -1
                            first_val = ""
                            for r_idx, row in enumerate(table_data):
                                val_0 = str(row[0]).strip()
                                if val_0.isdigit():
                                    data_start_idx = r_idx
                                    first_val = val_0
                                    break
                            
                            # 2. 判断逻辑：是新产品开始(1)，还是老产品续表(>1)
                            is_new = (first_val == "1")
                            is_cont = (first_val.isdigit() and int(first_val) > 1)
                            
                            if is_new:
                                if active_scheme: schemes.append(active_scheme)
                                active_scheme = {"cells": [], "offset": 0}
                            
                            # 3. 如果是有效内容，执行像素级坐标平移
                            if active_scheme is not None:
                                for cell in table_obj.cells:
                                    r0, c0, r1, c1 = int(cell[0]), int(cell[1]), int(cell[2]), int(cell[3])
                                    
                                    # 如果是续表，跳过重复的表头行
                                    if is_cont and r0 < data_start_idx:
                                        continue
                                    
                                    # 计算长表中的垂直位置
                                    shift = active_scheme["offset"]
                                    if is_cont:
                                        actual_r0 = r0 - data_start_idx + shift
                                        actual_r1 = r1 - data_start_idx + shift
                                    else:
                                        actual_r0 = r0 + shift
                                        actual_r1 = r1 + shift
                                    
                                    raw_text = table_data[r0][c0]
                                    is_data = (r0 >= data_start_idx)
                                    val = clean_numeric(raw_text) if is_data else str(raw_text).replace('\n', ' ')
                                    
                                    active_scheme["cells"].append({
                                        'r0': actual_r0, 'c0': c0, 'r1': actual_r1, 'c1': c1, 'val': val
                                    })
                                
                                # 更新偏移量：新表按全高算，续表只按数据高度算
                                if is_cont:
                                    active_scheme["offset"] += (len(table_data) - data_start_idx)
                                else:
                                    active_scheme["offset"] += len(table_data)

                    if active_scheme: schemes.append(active_scheme)

                    # 4. 写入 Excel
                    if not schemes:
                        st.warning("⚠️ 未能在 PDF 中识别到有效的保单利益表格。")
                    else:
                        for idx, s in enumerate(schemes):
                            sheet_name = f"方案对比_{idx + 1}"
                            worksheet = workbook.add_worksheet(sheet_name)
                            written = set()
                            
                            for c in s["cells"]:
                                r0, c0, r1, c1, val = c['r0'], c['c0'], c['r1'], c['c1'], c['val']
                                fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                                
                                if r1 - r0 > 1 or c1 - c0 > 1:
                                    try: worksheet.merge_range(r0, c0, r1-1, c1-1, val, fmt)
                                    except: pass # 忽略合并冲突
                                    for r in range(r0, r1):
                                        for col in range(c0, c1): written.add((r, col))
                                else:
                                    if (r0, c0) not in written:
                                        worksheet.write(r0, c0, val, fmt)
                                        written.add((r0, c0))
                            
                            worksheet.set_column(0, 30, 15)

            st.success(f"🎉 处理完成！已成功拼接 {len(schemes)} 份长表。")
            st.download_button(
                label="📥 点击下载“复印级”长表 Excel",
                data=output.getvalue(),
                file_name="平安建议书原样提取_V4.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ 程序运行出错: {str(e)}")
