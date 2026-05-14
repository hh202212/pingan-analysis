import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="平安建议书提取 V5.5", layout="wide")
st.title("🖨️ 平安建议书“复印级”提取 V5.5 (所见即所得版)")
st.info("改进：网页直接显示拼接后的长表 | 修复 Excel 下载为空 | 强力数字清洗")

def clean_to_number(val):
    if val is None or str(val).strip() == "": return ""
    s = str(val).replace('\n', '').replace(' ', '').replace(',', '').strip()
    if re.fullmatch(r'^-?[0-9.]+$', s):
        try:
            return float(s) if '.' in s else int(s)
        except: return s
    return s

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 原件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在深度扫描并缝合长表，请稍候...'):
            all_schemes = [] # 存储拼接后的方案数据
            current_scheme_data = [] # 当前正在拼接的行数据
            current_scheme_cells = [] # 当前正在拼接的单元格结构
            row_offset = 0
            
            with pdfplumber.open(uploaded_file) as pdf:
                for page_idx, page in enumerate(pdf.pages):
                    tables = page.find_tables()
                    if not tables: continue
                    
                    for t_obj in tables:
                        data = t_obj.extract()
                        if not data or len(data) == 0: continue
                        
                        # 判断是否为新表起始 (扫描前两行)
                        is_new = False
                        for r_check in range(min(2, len(data))):
                            row_str = "".join([str(c) for c in data[r_check] if c])
                            if "保单年度" in row_str or (data[r_check][0] and str(data[r_check][0]).strip() == "1"):
                                is_new = True
                                break
                        
                        # 如果是新表，保存上一个方案
                        if is_new and current_scheme_data:
                            all_schemes.append({"data": current_scheme_data, "cells": current_scheme_cells})
                            current_scheme_data = []
                            current_scheme_cells = []
                            row_offset = 0
                        
                        # 定位数据起始行
                        data_start = 0
                        for r_idx, row in enumerate(data):
                            if row and str(row[0]).strip().isdigit():
                                data_start = r_idx
                                break
                        
                        # 拼接逻辑：如果是续表，跳过表头
                        start_from = 0 if (is_new or not current_scheme_data) else data_start
                        
                        # 记录数据用于预览
                        for r_idx in range(start_from, len(data)):
                            current_scheme_data.append(data[r_idx])
                        
                        # 记录单元格结构用于 Excel
                        for cell in t_obj.cells:
                            r0, c0, r1, c1 = [int(x) for x in cell[:4]]
                            if not is_new and current_scheme_data and r0 < data_start:
                                continue
                            
                            act_r0 = r0 + row_offset if (is_new or not current_scheme_cells) else (r0 - data_start + row_offset)
                            act_r1 = r1 + row_offset if (is_new or not current_scheme_cells) else (r1 - data_start + row_offset)
                            
                            current_scheme_cells.append({
                                'r0': act_r0, 'c0': c0, 'r1': act_r1, 'c1': c1, 
                                'val': data[r0][c0], 'is_num': (r0 >= data_start)
                            })
                        
                        row_offset += (len(data) - start_from)

                if current_scheme_data:
                    all_schemes.append({"data": current_scheme_data, "cells": current_scheme_cells})

            # --- 结果展示区 ---
            if not all_schemes:
                st.warning("⚠️ 未能识别到有效的利益演示表。")
            else:
                st.success(f"🎉 成功缝合 {len(all_schemes)} 组长表方案！")
                
                # 1. 网页预览 (展示缝合后的长表)
                for idx, scheme in enumerate(all_schemes):
                    with st.expander(f"👁️ 方案 {idx+1} 缝合长表预览 (共 {len(scheme['data'])} 行)"):
                        st.dataframe(pd.DataFrame(scheme['data']), use_container_width=True)
                
                # 2. 生成 Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                    text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
                    
                    for idx, scheme in enumerate(all_schemes):
                        ws = workbook.add_worksheet(f"方案_{idx+1}")
                        written = set()
                        for c in scheme['cells']:
                            r0, c0, r1, c1, raw_val, is_num = c['r0'], c['c0'], c['r1'], c['c1'], c['val'], c['is_num']
                            val = clean_to_number(raw_val) if is_num else str(raw_val).replace('\n', ' ')
                            fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                            
                            try:
                                if r1 - r0 > 1 or c1 - c0 > 1:
                                    ws.merge_range(r0, c0, r1-1, c1-1, val, fmt)
                                    for r in range(r0, r1):
                                        for col in range(c0, c1): written.add((r, col))
                                elif (r0, c0) not in written:
                                    ws.write(r0, c0, val, fmt)
                                    written.add((r0, c0))
                            except: pass
                        ws.set_column(0, 50, 12)

                st.download_button(
                    label="📥 点击下载“缝合版”复印级 Excel",
                    data=output.getvalue(),
                    file_name="平安建议书长表拼接版.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ 运行异常: {str(e)}")
