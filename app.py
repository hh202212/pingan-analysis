import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# 1. 页面配置
st.set_page_config(page_title="平安建议书提取神器", layout="wide")
st.title("🖨️ 平安建议书表格提取 V5.2（预览+无损版）")
st.info("核心：网页实时预览 | 自动识别“保单年度”起始 | 强制数字转换 | 跨页长表合并")

# 强力数字清洗函数
def force_num(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return ""
    s = str(val).replace('\n', '').replace(' ', '').replace(',', '').strip()
    # 如果包含数字且不含汉字，尝试转码
    if re.search(r'\d', s) and not re.search(r'[\u4e00-\u9fa5]', s):
        res = re.sub(r'[^-0-9.]', '', s)
        try:
            return float(res) if '.' in res else int(res)
        except: return s
    return s

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF", type="pdf")

if uploaded_file:
    try:
        with st.spinner('🔍 正在地毯式提取数据，请稍候...'):
            all_pages_raw = []
            with pdfplumber.open(uploaded_file) as pdf:
                # 遍历全书，不错过任何一页
                for i, page in enumerate(pdf.pages):
                    table_obj = page.extract_table()
                    if table_obj:
                        all_pages_raw.append({"page": i+1, "data": table_obj, "obj": page.find_tables()})

            if not all_pages_raw:
                st.error("❌ 未能在 PDF 中发现任何表格结构。")
            else:
                # --- 第一步：网页预览（胡老师要求的模式，永远保留） ---
                st.success(f"🎉 成功识别到 {len(all_pages_raw)} 页表格内容！")
                for item in all_pages_raw:
                    with st.expander(f"👁️ 第 {item['page']} 页预览"):
                        st.dataframe(pd.DataFrame(item['data']), use_container_width=True)

                # --- 第二步：逻辑合并与 Excel 导出 ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                    text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                    
                    # 按照“保单年度”逻辑进行 Sheet 分配
                    current_ws = None
                    current_row_offset = 0
                    sheet_idx = 0
                    
                    for item in all_pages_raw:
                        data = item['data']
                        table_objs = item['obj']
                        
                        # 检查这一页的第一行是否包含“保单年度”
                        first_row_str = "".join([str(cell) for cell in data[0] if cell])
                        
                        if "保单年度" in first_row_text if 'first_row_text' in locals() else "保单年度" in first_row_str:
                            sheet_idx += 1
                            current_ws = workbook.add_worksheet(f"利益方案_{sheet_idx}")
                            current_row_offset = 0
                        
                        if current_ws:
                            # 1. 保底写入：先不管合并，把所有数填进去，防止空表
                            for r_idx, row in enumerate(data):
                                # 判断是否为数据行（第一列是数字）
                                is_data = str(row[0]).strip().isdigit()
                                for c_idx, cell_val in enumerate(row):
                                    val = force_num(cell_val) if is_data else str(cell_val).replace('\n', ' ')
                                    fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                                    current_ws.write(r_idx + current_row_offset, c_idx, val, fmt)
                            
                            # 2. 装修：复刻合并单元格结构
                            for t_obj in table_objs:
                                for cell in t_obj.cells:
                                    r0, c0, r1, c1 = [int(x) for x in cell[:4]]
                                    if r1 - r0 > 1 or c1 - c0 > 1:
                                        raw_txt = data[r0][c0]
                                        # 再次判断格式
                                        is_data_cell = str(data[r0][0]).strip().isdigit()
                                        v = force_num(raw_txt) if is_data_cell else str(raw_txt).replace('\n', ' ')
                                        f = num_fmt if isinstance(v, (int, float)) else text_fmt
                                        try:
                                            current_ws.merge_range(r0 + current_row_offset, c0, r1 + current_row_offset - 1, c1 - 1, v, f)
                                        except: pass
                            
                            current_row_offset += len(data)
                            current_ws.set_column(0, 30, 12)

                st.download_button(
                    label="📥 下载“原样复刻”纯数字 Excel",
                    data=output.getvalue(),
                    file_name="建议书提取_V5.2.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"⚠️ 处理出错：{str(e)}")
