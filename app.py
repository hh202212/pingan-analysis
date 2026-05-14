import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="平安建议书提取 V7.0", layout="wide")
st.title("🖨️ 平安建议书“复印级”提取 V7.0 (终极稳定版)")
st.info("核心突破：彻底抛弃易崩溃的物理坐标计算，采用纯逻辑矩阵自动推导合并，100% 防崩溃！")

# 强力数字清洗
def clean_to_number(val):
    if val is None or str(val).strip() == "": return ""
    s = str(val).replace('\n', '').replace(' ', '').replace(',', '').replace('¥', '').strip()
    if re.fullmatch(r'^-?[0-9.]+$', s):
        try: return float(s) if '.' in s else int(s)
        except: return s
    return str(val).replace('\n', ' ')

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在启动矩阵降维算法，执行无缝拼接...'):
            all_matrices = [] # 存储拼接后的大矩阵
            current_matrix = [] # 当前正在拼接的二维数组
            
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    tables = page.find_tables()
                    if not tables: continue
                    
                    for t_obj in tables:
                        data = t_obj.extract()
                        if not data or len(data) == 0: continue
                        
                        # 1. 寻找数据行起始位置 (第一列是数字的行)
                        data_start_idx = -1
                        first_year = -1
                        for r_idx, row in enumerate(data):
                            v0 = str(row[0]).strip() if row and row[0] else ""
                            if v0.isdigit():
                                data_start_idx = r_idx
                                first_year = int(v0)
                                break
                        
                        # 2. 判定：新方案 vs 续表
                        is_new = (first_year == 1)
                        is_cont = (first_year > 1)
                        
                        if is_new:
                            if current_matrix: all_matrices.append(current_matrix)
                            current_matrix = []
                            start_idx = 0 # 新表，保留包括表头在内的所有行
                        elif is_cont and current_matrix:
                            start_idx = data_start_idx # 续表，跳过重复的表头
                        else:
                            continue # 非目标表格
                        
                        # 3. 将数据拼入当前大矩阵
                        for r_idx in range(start_idx, len(data)):
                            is_data_row = (r_idx >= data_start_idx)
                            cleaned_row = []
                            for val in data[r_idx]:
                                if is_data_row:
                                    cleaned_row.append(clean_to_number(val))
                                else:
                                    cleaned_row.append(str(val).replace('\n', ' ') if val else "")
                            current_matrix.append(cleaned_row)

                if current_matrix: all_matrices.append(current_matrix)

            # --- 结果呈现与 Excel 生成 ---
            if not all_matrices:
                st.warning("⚠️ 未能识别到‘保单年度 1’，请检查 PDF。")
            else:
                st.success(f"🎉 降维拼接成功！共整合 {len(all_matrices)} 个长表方案。")
                
                # 网页预览区：展示拼接好的长矩阵
                for idx, matrix in enumerate(all_matrices):
                    with st.expander(f"👁️ 方案 {idx+1} 无缝长表预览 (共 {len(matrix)} 行)"):
                        st.dataframe(pd.DataFrame(matrix), use_container_width=True)
                
                # 生成 Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                    text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                    
                    for idx, matrix in enumerate(all_matrices):
                        ws = workbook.add_worksheet(f"方案_{idx+1}")
                        
                        # 补齐每一行的列数，确保是个完美的矩形
                        max_cols = max(len(row) for row in matrix) if matrix else 0
                        for row in matrix:
                            while len(row) < max_cols: row.append("")
                        
                        written = set()
                        # 核心黑科技：矩阵空值推导合并法
                        for r in range(len(matrix)):
                            for c in range(max_cols):
                                if (r, c) in written: continue
                                
                                val = matrix[r][c]
                                if val == "": continue # 纯空单元格直接跳过
                                
                                # 向右探测空值（判断跨列）
                                c_span = 1
                                while c + c_span < max_cols and matrix[r][c + c_span] == "" and (r, c + c_span) not in written:
                                    c_span += 1
                                
                                # 向下探测空值（判断跨行）
                                r_span = 1
                                while r + r_span < len(matrix):
                                    row_is_blank = True
                                    for c_idx in range(c, c + c_span):
                                        if matrix[r + r_span][c_idx] != "" or (r + r_span, c_idx) in written:
                                            row_is_blank = False
                                            break
                                    if row_is_blank: r_span += 1
                                    else: break
                                
                                # 执行写入或合并
                                fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                                if r_span > 1 or c_span > 1:
                                    ws.merge_range(r, c, r + r_span - 1, c + c_span - 1, val, fmt)
                                    for rr in range(r, r + r_span):
                                        for cc in range(c, c + c_span): written.add((rr, cc))
                                else:
                                    ws.write(r, c, val, fmt)
                                    written.add((r, c))
                        
                        ws.set_column(0, max_cols, 12)

                st.download_button(
                    label="📥 下载“终极防崩”长表 Excel",
                    data=output.getvalue(),
                    file_name="平安建议书提取_V7.0.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ 运行异常: {str(e)}")
