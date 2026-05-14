import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="平安建议书提取神器 V5.4", layout="wide")
st.title("🖨️ 平安建议书表格“复印级”提取 V5.4")
st.info("核心改进：彻底修复 Index 报错 | 模糊匹配表头起始 | 跨页自动无缝拼接 | 纯数字转换")

# 强力数字转换，解决绿三角
def clean_to_number(val):
    if val is None or str(val).strip() == "": return ""
    # 移除换行、空格、逗号
    s = str(val).replace('\n', '').replace(' ', '').replace(',', '').strip()
    # 匹配纯数字/小数点
    if re.fullmatch(r'^-?[0-9.]+$', s):
        try:
            return float(s) if '.' in s else int(s)
        except: return s
    return s

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 原件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在像素级扫描全书并执行逻辑对齐...'):
            output = io.BytesIO()
            all_previews = [] 
            
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                
                with pdfplumber.open(uploaded_file) as pdf:
                    current_ws = None
                    current_row_offset = 0
                    sheet_idx = 0
                    
                    for page_idx, page in enumerate(pdf.pages):
                        tables = page.find_tables()
                        if not tables: continue
                        
                        # 每一页的原始数据先存预览
                        page_raw = page.extract_table()
                        if page_raw:
                            all_previews.append((f"第{page_idx+1}页", pd.DataFrame(page_raw)))

                        for t_obj in tables:
                            data = t_obj.extract()
                            # 防崩检查 1：如果表格没内容，直接跳过
                            if not data or len(data) == 0: continue
                            
                            # --- 逻辑判定：是否为新表起始 ---
                            # 扫描前两行，寻找关键字
                            is_new_sheet = False
                            for r_check in range(min(2, len(data))):
                                row_str = "".join([str(c) for c in data[r_check] if c])
                                if "保单年度" in row_str or "年度" in row_str:
                                    is_new_sheet = True
                                    break
                            
                            if is_new_sheet:
                                sheet_idx += 1
                                current_ws = workbook.add_worksheet(f"方案_{sheet_idx}")
                                current_row_offset = 0
                            
                            # 只要有正在操作的 Sheet，就开始写入
                            if current_ws:
                                # 定位该表的数据起始行（第一列是数字的行）
                                data_start_row = 0
                                for r_idx, row in enumerate(data):
                                    if row and str(row[0]).strip().isdigit():
                                        data_start_row = r_idx
                                        break
                                
                                written_mark = set()
                                # 复刻合并单元格结构
                                for cell in t_obj.cells:
                                    # 防崩检查 2：确保索引为整数且在范围内
                                    try:
                                        r0, c0, r1, c1 = [int(x) for x in cell[:4]]
                                        if r0 >= len(data) or c0 >= len(data[0]): continue
                                        
                                        # 如果是续表（当前页没搜到“年度”），跳过重复的表头
                                        if not is_new_sheet and r0 < data_start_row:
                                            continue
                                        
                                        # 计算长表中的垂直位置
                                        actual_r0 = (r0 - data_start_row + current_row_offset) if not is_new_sheet else (r0 + current_row_offset)
                                        actual_r1 = (r1 - data_start_row + current_row_offset) if not is_new_sheet else (r1 + current_row_offset)
                                        
                                        raw_text = data[r0][c0]
                                        is_num_row = (r0 >= data_start_row)
                                        val = clean_to_number(raw_text) if is_num_row else str(raw_text).replace('\n', ' ')
                                        fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                                        
                                        # 执行合并或写入
                                        if actual_r1 - actual_r0 > 1 or c1 - c0 > 1:
                                            current_ws.merge_range(actual_r0, c0, actual_r1 - 1, c1 - 1, val, fmt)
                                            for r_m in range(actual_r0, actual_r1):
                                                for c_m in range(c0, c1): written_mark.add((r_m, c_m))
                                        elif (actual_r0, c0) not in written_mark:
                                            current_ws.write(actual_r0, c0, val, fmt)
                                            written_mark.add((actual_r0, c0))
                                    except:
                                        continue
                                
                                # 更新偏移量
                                current_row_offset += (len(data) - data_start_row) if not is_new_sheet else len(data)
                                current_ws.set_column(0, 50, 12)

            # --- 结果展示 ---
            st.success(f"🎉 扫描完成！共整合出 {sheet_idx} 份完整长表方案。")
            
            # 网页预览预览区
            for title, df in all_previews:
                with st.expander(f"👁️ {title} 原始数据预览"):
                    st.dataframe(df, use_container_width=True)
            
            if sheet_idx > 0:
                st.download_button(
                    label="📥 下载“长表拼接”纯数字 Excel",
                    data=output.getvalue(),
                    file_name="建议书提取_V5.4.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("并未识别到有效的利益演示表起始标记。")

    except Exception as e:
        st.error(f"❌ 运行遇到异常: {str(e)}")
