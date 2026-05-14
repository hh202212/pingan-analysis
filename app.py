import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="平安建议书提取 V6.1", layout="wide")
st.title("🖨️ 平安建议书“复印级”缝合工具 V6.1")
st.info("核心：识别‘1’开新表 | 其余自动缝合 | 保留首页复杂表头 | 纯数字转换")

def clean_num(val):
    if val is None or str(val).strip() == "": return ""
    s = str(val).replace('\n', '').replace(' ', '').replace(',', '').strip()
    if re.fullmatch(r'^-?[0-9.]+$', s):
        try: return float(s) if '.' in s else int(s)
        except: return s
    return str(val).replace('\n', ' ')

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在像素级缝合长表并复刻表头...'):
            all_schemes = [] 
            active_scheme = None # {"rows": [], "merges": [], "y_offset": 0}
            
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    tables = page.find_tables()
                    if not tables: continue
                    
                    for t_obj in tables:
                        data = t_obj.extract()
                        if not data or len(data) < 2: continue
                        
                        # 1. 寻找这一页的数据起始行
                        data_start_idx = -1
                        first_year = -1
                        for r_idx, row in enumerate(data):
                            v0 = str(row[0]).strip()
                            if v0.isdigit():
                                data_start_idx = r_idx
                                first_year = int(v0)
                                break
                        
                        # 2. 判定：是新产品(1) 还是 续表(>1)
                        is_new = (first_year == 1)
                        is_cont = (first_year > 1)
                        
                        if is_new:
                            if active_scheme: all_schemes.append(active_scheme)
                            active_scheme = {"rows": [], "merges": [], "y_offset": 0}
                            # 新表：保留所有行（含表头）
                            start_r = 0
                        elif is_cont and active_scheme:
                            # 续表：跳过表头行
                            start_r = data_start_idx
                        else:
                            # 既不是1也不是续表（可能是纯说明页），跳过
                            continue

                        # 3. 记录行数据
                        for r_idx in range(start_r, len(data)):
                            # 如果是数据行执行数字清洗，否则保留原样
                            is_data = (r_idx >= data_start_idx)
                            cleaned_row = [clean_num(c) if is_data else str(c).replace('\n',' ') for c in data[r_idx]]
                            active_scheme["rows"].append(cleaned_row)
                        
                        # 4. 记录合并单元格（复印级关键）
                        current_y = active_scheme["y_offset"]
                        for cell in t_obj.cells:
                            r0, c0, r1, c1 = [int(round(x)) for x in cell[:4]]
                            # 续表跳过重复表头合并
                            if not is_new and r0 < data_start_idx: continue
                            
                            act_r0 = r0 + current_y if is_new else (r0 - data_start_idx + current_y)
                            act_r1 = r1 + current_y if is_new else (r1 - data_start_idx + current_y)
                            
                            active_scheme["merges"].append({
                                'r0': act_r0, 'c0': c0, 'r1': act_r1, 'c1': c1, 'val': data[r0][c0]
                            })
                        
                        # 更新偏移量
                        active_scheme["y_offset"] += (len(data) - start_r)

                if active_scheme: all_schemes.append(active_scheme)

            # --- 渲染区 ---
            if not all_schemes:
                st.warning("⚠️ 未能识别到利益表，请确保 PDF 第一列有数字 1。")
            else:
                st.success(f"🎉 缝合成功！已整合为 {len(all_schemes)} 个长表。")
                
                for idx, sc in enumerate(all_schemes):
                    with st.expander(f"👁️ 方案 {idx+1} 完整长表预览 (共 {len(sc['rows'])} 行)"):
                        st.dataframe(pd.DataFrame(sc["rows"]), use_container_width=True)
                
                # 生成 Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                    text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                    
                    for idx, sc in enumerate(all_schemes):
                        ws = workbook.add_worksheet(f"方案_{idx+1}")
                        written = set()
                        # 先写合并大框
                        sorted_merges = sorted(sc["merges"], key=lambda x: (x['r1']-x['r0']), reverse=True)
                        for m in sorted_merges:
                            r0, c0, r1, c1, raw_val = m['r0'], m['c0'], m['r1'], m['c1'], m['val']
                            val = clean_num(raw_val) if any(char.isdigit() for char in str(raw_val)) else str(raw_val).replace('\n',' ')
                            fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                            try:
                                if r1 - r0 > 1 or c1 - c0 > 1:
                                    ws.merge_range(r0, c0, r1-1, c1-1, val, fmt)
                                    for r in range(r0, r1):
                                        for c in range(c0, c1): written.add((r, c))
                                elif (r0, c0) not in written:
                                    ws.write(r0, c0, val, fmt)
                                    written.add((r0, c0))
                            except: pass
                        ws.set_column(0, 50, 15)

                st.download_button("📥 下载无损缝合长表 Excel", output.getvalue(), "建议书完整提取_V6.1.xlsx")

    except Exception as e:
        st.error(f"❌ 程序崩溃：{str(e)}")
