import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="平安建议书提取 V5.8", layout="wide")
st.title("🖨️ 平安建议书“复印级”提取 V5.8 (全量缝合版)")
st.info("核心：跨页逻辑无缝缝合 | 还原嵌套合并表头 | 网页预览=Excel下载 | 纯数字转换")

def clean_num(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return ""
    s = str(val).replace('\n', '').replace(' ', '').replace(',', '').strip()
    if re.fullmatch(r'^-?[0-9.]+$', s) and not re.search(r'[\u4e00-\u9fa5]', s):
        try:
            return float(s) if '.' in s else int(s)
        except: return s
    return str(val).replace('\n', ' ')

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在像素级缝合长表数据，请稍候...'):
            all_schemes = [] # 存储结构：{"rows": [], "merges": [], "row_offset": 0}
            active_scheme = None
            
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    tables = page.find_tables()
                    if not tables: continue
                    
                    for t_obj in tables:
                        data = t_obj.extract()
                        if not data or len(data) == 0: continue
                        
                        # 1. 寻找数据起始行和起始年度
                        data_start_idx = -1
                        first_year = -1
                        for r_idx, row in enumerate(data):
                            val0 = str(row[0]).strip()
                            if val0.isdigit():
                                data_start_idx = r_idx
                                first_year = int(val0)
                                break
                        
                        # 2. 判定：是新方案起始(1)，还是老方案续接(>1)
                        is_new = (first_year == 1)
                        is_cont = (first_year > 1)
                        
                        if is_new:
                            if active_scheme: all_schemes.append(active_scheme)
                            active_scheme = {"rows": [], "merges": [], "row_offset": 0}
                        
                        if active_scheme is not None:
                            # 3. 记录行数据用于预览和Excel基础写入
                            start_from = 0 if is_new else data_start_idx
                            if data_start_idx == -1: start_from = 0 # 兜底
                            
                            for r_idx in range(start_from, len(data)):
                                # 清洗这一行的数据
                                cleaned_row = [clean_num(c) for c in data[r_idx]]
                                active_scheme["rows"].append(cleaned_row)
                            
                            # 4. 记录合并单元格结构用于Excel“复印”
                            for cell in t_obj.cells:
                                r0, c0, r1, c1 = [int(round(x)) for x in cell[:4]]
                                # 续表逻辑：跳过重复表头
                                if is_cont and r0 < data_start_idx: continue
                                
                                # 计算相对于长表的全局行坐标
                                shift = active_scheme["row_offset"]
                                act_r0 = r0 + shift if is_new else (r0 - data_start_idx + shift)
                                act_r1 = r1 + shift if is_new else (r1 - data_start_idx + shift)
                                
                                active_scheme["merges"].append({
                                    'r0': act_r0, 'c0': c0, 'r1': act_r1, 'c1': c1,
                                    'val': clean_num(data[r0][c0])
                                })
                            
                            # 更新方案的总行数偏移量
                            active_scheme["row_offset"] += (len(data) - start_from)

                if active_scheme: all_schemes.append(active_scheme)

            # --- 结果呈现 ---
            if not all_schemes:
                st.warning("⚠️ 未识别到有效利益表。")
            else:
                st.success(f"🎉 缝合成功！已整合 {len(all_schemes)} 份长表方案。")
                
                # 网页预览预览区
                for idx, scheme in enumerate(all_schemes):
                    with st.expander(f"👁️ 方案 {idx+1} 预览 (已无缝拼接 {len(scheme['rows'])} 行)"):
                        st.dataframe(pd.DataFrame(scheme['rows']), use_container_width=True)
                
                # 生成 Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                    text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                    
                    for idx, scheme in enumerate(all_schemes):
                        ws = workbook.add_worksheet(f"方案_{idx+1}")
                        written_set = set()
                        
                        # 先根据 merges 记录执行“复印级”合并
                        for m in scheme["merges"]:
                            r0, c0, r1, c1, val = m['r0'], m['c0'], m['r1'], m['c1'], m['val']
                            fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                            try:
                                if r1 - r0 > 1 or c1 - c0 > 1:
                                    ws.merge_range(r0, c0, r1-1, c1-1, val, fmt)
                                    for r in range(r0, r1):
                                        for c in range(c0, c1): written_set.add((r, c))
                                elif (r0, c0) not in written_set:
                                    ws.write(r0, c0, val, fmt)
                                    written_set.add((r0, c0))
                            except: pass
                        
                        # 补全可能遗漏的普通单元格
                        for r_idx, row in enumerate(scheme["rows"]):
                            for c_idx, val in enumerate(row):
                                if (r_idx, c_idx) not in written_set:
                                    fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                                    ws.write(r_idx, c_idx, val, fmt)
                        
                        ws.set_column(0, 50, 15)

                st.download_button(
                    label="📥 下载“原样复刻”缝合长表",
                    data=output.getvalue(),
                    file_name="平安建议书长表复刻版.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ 运行异常: {str(e)}")
