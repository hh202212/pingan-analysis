import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="平安建议书复刻 V6.0", layout="wide")
st.title("🖨️ 平安建议书“复印级”缝合工具 V6.0")
st.info("核心：只要第一列不是‘1’就自动向下缝合 | 强制保留原始表头结构 | 纯数字无绿三角")

# 强力数字清洗
def clean_to_number(val):
    if val is None or str(val).strip() == "": return ""
    s = str(val).replace('\n', '').replace(' ', '').replace(',', '').strip()
    if re.fullmatch(r'^-?[0-9.]+$', s):
        try:
            return float(s) if '.' in s else int(s)
        except: return s
    return str(val).replace('\n', ' ')

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在像素级还原并执行长表缝合...'):
            schemes = [] # 存储最终方案
            active_scheme = None # 当前操作的方案：{"rows": [], "merges": [], "y_offset": 0}
            
            with pdfplumber.open(uploaded_file) as pdf:
                # 遍历全书（通常从10页开始是重点）
                for page_idx in range(len(pdf.pages)):
                    page = pdf.pages[page_idx]
                    tables = page.find_tables()
                    if not tables: continue
                    
                    for t_obj in tables:
                        data = t_obj.extract()
                        if not data or len(data) < 2: continue
                        
                        # 判定是否为新产品起始 (第一列出现数字 1)
                        is_new_product = False
                        first_row_data = -1
                        for r_idx, row in enumerate(data):
                            if row[0] and str(row[0]).strip() == "1":
                                is_new_product = True
                                first_row_data = r_idx
                                break
                        
                        # 如果是新产品，结算前一个
                        if is_new_product:
                            if active_scheme: schemes.append(active_scheme)
                            active_scheme = {"rows": data, "merges": [], "y_offset": 0}
                            # 记录这一页的合并逻辑
                            for cell in t_obj.cells:
                                r0, c0, r1, c1 = [int(round(x)) for x in cell[:4]]
                                active_scheme["merges"].append({'r0': r0, 'c0': c0, 'r1': r1, 'c1': c1, 'val': data[r0][c0]})
                            active_scheme["y_offset"] = len(data)
                        
                        # 如果是续表（第一列是大于1的数字），直接拼在 active_scheme 后面
                        elif active_scheme is not None:
                            current_y = active_scheme["y_offset"]
                            # 写入这一页的数据
                            active_scheme["rows"].extend(data)
                            # 写入这一页的合并逻辑（带坐标偏移）
                            for cell in t_obj.cells:
                                r0, c0, r1, c1 = [int(round(x)) for x in cell[:4]]
                                active_scheme["merges"].append({
                                    'r0': r0 + current_y, 
                                    'c0': c0, 
                                    'r1': r1 + current_y, 
                                    'c1': c1, 
                                    'val': data[r0][c0]
                                })
                            active_scheme["y_offset"] += len(data)

                if active_scheme: schemes.append(active_scheme)

            # --- 结果呈现 ---
            if not schemes:
                st.warning("⚠️ 未能识别到‘保单年度 1’，请确认 PDF 是否包含利益演示表。")
            else:
                st.success(f"🎉 缝合成功！已将跨页数据整合为 {len(schemes)} 个方案。")
                
                # 网页预览（胡老师要求的模式）
                for i, sc in enumerate(schemes):
                    with st.expander(f"👁️ 方案 {i+1} 缝合长表预览"):
                        st.dataframe(pd.DataFrame(sc["rows"]), use_container_width=True)
                
                # 生成 Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                    text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                    
                    for idx, sc in enumerate(schemes):
                        ws = workbook.add_worksheet(f"方案_{idx+1}")
                        written_cells = set()
                        
                        # 核心：按照记录的合并逻辑进行“复印级”写入
                        # 排序：先写合并的大框，再写普通的小格子
                        sorted_merges = sorted(sc["merges"], key=lambda x: (x['r1']-x['r0']) * (x['c1']-x['c0']), reverse=True)
                        
                        for m in sorted_merges:
                            r0, c0, r1, c1, raw_val = m['r0'], m['c0'], m['r1'], m['c1'], m['val']
                            # 判断是否为数据行，执行数字转换
                            is_num = any(char.isdigit() for char in str(raw_val)) and "年度" not in str(raw_val)
                            val = clean_to_number(raw_val) if is_num else str(raw_val).replace('\n', ' ')
                            fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                            
                            try:
                                if r1 - r0 > 1 or c1 - c0 > 1:
                                    ws.merge_range(r0, c0, r1-1, c1-1, val, fmt)
                                    for r in range(r0, r1):
                                        for c in range(c0, c1): written_cells.add((r, c))
                                elif (r0, c0) not in written_cells:
                                    ws.write(r0, c0, val, fmt)
                                    written_cells.add((r0, c0))
                            except:
                                pass
                        
                        ws.set_column(0, 50, 15)

                st.download_button(
                    label="📥 点击下载“无损缝合”长表 Excel",
                    data=output.getvalue(),
                    file_name="平安建议书提取_V6.0.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ 程序崩溃：{str(e)}")
