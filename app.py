import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# 1. 页面配置
st.set_page_config(page_title="平安建议书提取 V5.9", layout="wide")
st.title("🖨️ 平安建议书表格“复印级”提取 V5.9")
st.info("核心改进：逻辑防崩溃机制 | 真正的全量长表缝合 | 仅保留首个复印级表头 | 强力数字清洗")

# 强制数字清洗函数
def clean_num(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return 0
    s = str(val).replace('\n', '').replace(' ', '').replace(',', '').strip()
    if re.fullmatch(r'^-?[0-9.]+$', s) and not re.search(r'[\u4e00-\u9fa5]', s):
        try:
            return float(s) if '.' in s else int(s)
        except: return s
    return str(val).replace('\n', ' ')

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在为您执行“像素级”缝合与结构还原...'):
            schemes = [] # 存储最终结果：{"header_merges": [], "header_rows": [], "data_rows": []}
            current_scheme = None
            
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    tables = page.find_tables()
                    if not tables: continue
                    
                    for t_obj in tables:
                        data = t_obj.extract()
                        # 防崩溃：数据为空或行数不足直接跳过
                        if not data or len(data) < 1: continue
                        
                        # 识别数据行起始 (第一列是数字的行)
                        data_start_idx = -1
                        first_year = -1
                        for r_idx, row in enumerate(data):
                            if row and str(row[0]).strip().isdigit():
                                data_start_idx = r_idx
                                first_year = int(str(row[0]).strip())
                                break
                        
                        # 判断逻辑：新方案 vs 续表
                        is_new = (first_year == 1)
                        
                        if is_new:
                            # 结算上一个方案
                            if current_scheme: schemes.append(current_scheme)
                            # 初始化新方案
                            current_scheme = {
                                "header_rows": data[:data_start_idx] if data_start_idx != -1 else [],
                                "header_merges": [],
                                "data_rows": [],
                                "header_height": data_start_idx if data_start_idx != -1 else 0
                            }
                            # 记录“复印级”合并表头结构
                            for cell in t_obj.cells:
                                r0, c0, r1, c1 = [int(round(x)) for x in cell[:4]]
                                if r0 < current_scheme["header_height"]:
                                    current_scheme["header_merges"].append({
                                        'r0': r0, 'c0': c0, 'r1': r1, 'c1': c1, 'val': data[r0][c0]
                                    })
                        
                        # 只要有正在操作的方案，就往里塞数据行
                        if current_scheme is not None:
                            # 如果没找到数据起始，可能整页都是标题或垃圾，跳过
                            if data_start_idx != -1:
                                for r_idx in range(data_start_idx, len(data)):
                                    cleaned_row = [clean_num(c) for c in data[r_idx]]
                                    current_scheme["data_rows"].append(cleaned_row)

                if current_scheme: schemes.append(current_scheme)

            # --- 渲染区 ---
            if not schemes:
                st.warning("⚠️ 未能识别到有效的利益演示表起始标记（保单年度 1）。")
            else:
                st.success(f"🎉 缝合成功！已为您处理 {len(schemes)} 份完整利益方案。")
                
                # 1. 网页预览 (将表头和数据拼接展示)
                for idx, sc in enumerate(schemes):
                    full_table = sc["header_rows"] + sc["data_rows"]
                    with st.expander(f"👁️ 方案 {idx+1} 预览 (已无缝缝合 {len(full_table)} 行数据)"):
                        st.dataframe(pd.DataFrame(full_table), use_container_width=True)
                
                # 2. 生成 Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                    text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                    
                    for idx, sc in enumerate(schemes):
                        ws = workbook.add_worksheet(f"方案_{idx+1}")
                        written_cells = set()
                        
                        # A. 写入复印级表头 (带合并单元格)
                        for m in sc["header_merges"]:
                            r0, c0, r1, c1, val = m['r0'], m['c0'], m['r1'], m['c1'], m['val']
                            fmt = text_fmt
                            try:
                                if r1 - r0 > 1 or c1 - c0 > 1:
                                    ws.merge_range(r0, c0, r1-1, c1-1, str(val).replace('\n',' '), fmt)
                                    for r_m in range(r0, r1):
                                        for c_m in range(c0, c1): written_cells.add((r_m, c_m))
                                else:
                                    ws.write(r0, c0, str(val).replace('\n',' '), fmt)
                                    written_cells.add((r0, c0))
                            except: pass
                        
                        # B. 写入无缝拼接的数据行 (年度 1 - 105)
                        h_height = sc["header_height"]
                        for dr_idx, d_row in enumerate(sc["data_rows"]):
                            actual_row = h_height + dr_idx
                            for dc_idx, d_val in enumerate(d_row):
                                fmt = num_fmt if isinstance(d_val, (int, float)) else text_fmt
                                ws.write(actual_row, dc_idx, d_val, fmt)
                        
                        ws.set_column(0, 50, 15)

                st.download_button(
                    label="📥 下载“全量缝合”复印级 Excel",
                    data=output.getvalue(),
                    file_name="平安建议书提取_V5.9.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ 程序发生预料外的崩溃: {str(e)}")
        st.info("提示：请确认 PDF 文件没有被加密，且是直接从平安系统导出的电子版。")
