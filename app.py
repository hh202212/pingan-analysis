import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="平安建议书提取 V5.6", layout="wide")
st.title("🖨️ 平安建议书表格“复印级”提取 V5.6")
st.info("改进：修复浮点索引与越界错误 | 网页预览与下载完全同步 | 自动跨页长表缝合")

# 强制数字转换，彻底解决绿三角
def clean_to_number(val):
    if val is None or str(val).strip() == "": return ""
    # 移除换行、空格、逗号、人民币符号
    s = str(val).replace('\n', '').replace(' ', '').replace(',', '').replace('¥', '').strip()
    if re.fullmatch(r'^-?[0-9.]+$', s):
        try:
            return float(s) if '.' in s else int(s)
        except: return s
    return s

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 原件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在深度扫描并缝合长表，请稍候...'):
            all_schemes = [] # 存储方案
            current_scheme = {"data": [], "cells": [], "row_offset": 0}
            
            with pdfplumber.open(uploaded_file) as pdf:
                for page_idx, page in enumerate(pdf.pages):
                    tables = page.find_tables()
                    if not tables: continue
                    
                    for t_obj in tables:
                        table_data = t_obj.extract()
                        if not table_data or len(table_data) == 0: continue
                        
                        # 判定是否为新表起始
                        is_new = False
                        first_row_str = "".join([str(c) for c in table_data[0] if c])
                        if "保单年度" in first_row_str or (table_data[0][0] and str(table_data[0][0]).strip() == "1"):
                            is_new = True
                        
                        # 如果是新表，结算上一个方案
                        if is_new and current_scheme["data"]:
                            all_schemes.append(current_scheme)
                            current_scheme = {"data": [], "cells": [], "row_offset": 0}
                        
                        # 定位数据行
                        data_start = 0
                        for r_idx, row in enumerate(table_data):
                            if row and str(row[0]).strip().isdigit():
                                data_start = r_idx
                                break
                        
                        # 记录数据（用于预览）
                        start_writing = 0 if (is_new or not current_scheme["data"]) else data_start
                        for r_idx in range(start_writing, len(table_data)):
                            current_scheme["data"].append(table_data[r_idx])
                        
                        # 记录单元格结构（用于 Excel 合并）
                        for cell in t_obj.cells:
                            # 强制转为整数，防止 float 索引报错
                            try:
                                r0, c0, r1, c1 = [int(round(x)) for x in cell[:4]]
                                
                                # 续表逻辑：跳过表头
                                if not is_new and current_scheme["data"] and r0 < data_start:
                                    continue
                                
                                # 坐标平移
                                shift = current_scheme["row_offset"]
                                act_r0 = r0 + shift if (is_new or shift == 0) else (r0 - data_start + shift)
                                act_r1 = r1 + shift if (is_new or shift == 0) else (r1 - data_start + shift)
                                
                                # 越界保护
                                if r0 >= len(table_data) or c0 >= len(table_data[0]): continue
                                
                                current_scheme["cells"].append({
                                    'r0': act_r0, 'c0': c0, 'r1': act_r1, 'c1': c1,
                                    'val': table_data[r0][c0], 
                                    'is_num': (r0 >= data_start)
                                })
                            except: continue
                        
                        # 更新偏移量
                        current_scheme["row_offset"] += (len(table_data) - start_writing)

                if current_scheme["data"]:
                    all_schemes.append(current_scheme)

            # --- 渲染区 ---
            if not all_schemes:
                st.warning("⚠️ 未能识别到有效的利益演示表。")
            else:
                st.success(f"🎉 成功缝合 {len(all_schemes)} 组方案！")
                
                # 1. 网页预览 (显示缝合后的结果)
                for idx, scheme in enumerate(all_schemes):
                    with st.expander(f"👁️ 方案 {idx+1} 预览 (已自动合并跨页数据)"):
                        st.dataframe(pd.DataFrame(scheme['data']), use_container_width=True)
                
                # 2. 生成 Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                    text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                    
                    for idx, scheme in enumerate(all_schemes):
                        ws = workbook.add_worksheet(f"方案_{idx+1}")
                        written_cells = set()
                        
                        # 排序：先写合并单元格，再写普通单元格
                        for c in sorted(scheme['cells'], key=lambda x: (x['r1']-x['r0']), reverse=True):
                            r0, c0, r1, c1, raw_val, is_num = c['r0'], c['c0'], c['r1'], c['c1'], c['val'], c['is_num']
                            val = clean_to_number(raw_val) if is_num else str(raw_val).replace('\n', ' ')
                            fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                            
                            try:
                                if r1 - r0 > 1 or c1 - c0 > 1:
                                    ws.merge_range(r0, c0, r1-1, c1-1, val, fmt)
                                    for r in range(r0, r1):
                                        for col in range(c0, c1): written_cells.add((r, col))
                                elif (r0, c0) not in written_cells:
                                    ws.write(r0, c0, val, fmt)
                                    written_cells.add((r0, c0))
                            except: pass
                        ws.set_column(0, 50, 15)

                st.download_button(
                    label="📥 点击下载“复印级”长表 Excel",
                    data=output.getvalue(),
                    file_name="平安建议书提取_V5.6.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ 程序运行崩溃: {str(e)}")
        st.info("提示：这通常是因为 PDF 内部结构极度复杂。请尝试重新下载 PDF 电子版后再试。")
