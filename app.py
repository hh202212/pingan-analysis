import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# 1. 页面基础配置
st.set_page_config(page_title="平安建议书复刻神器 V3.8", layout="wide")
st.title("🖨️ 平安建议书表格“复印级”提取 (V3.8 全行对齐版)")
st.info("改进：全行扫描“保单年度”关键字 | 网页实时预览 | 完美还原合并单元格")

# 强制数值转换
def clean_val(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return ""
    s = str(val).replace('\n', '').replace(' ', '').strip()
    if re.fullmatch(r'^-?[0-9,.]+$', s):
        num_s = re.sub(r'[^-0-9.]', '', s)
        try:
            return float(num_s) if '.' in num_s else int(num_s)
        except: return s
    return s

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 原件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在地毯式搜寻利益演示表...'):
            output = io.BytesIO()
            all_previews = [] # 存储所有抓到的数据，无论是否匹配关键词
            
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                
                with pdfplumber.open(uploaded_file) as pdf:
                    current_worksheet = None
                    current_row_offset = 0
                    table_count = 0
                    
                    for page_idx, page in enumerate(pdf.pages):
                        tables = page.find_tables()
                        if not tables: continue
                        
                        for table_obj in tables:
                            table_data = table_obj.extract()
                            if not table_data: continue
                            
                            # 网页预览：先存起来，保证网页不留白
                            all_previews.append((f"第{page_idx+1}页", pd.DataFrame(table_data)))
                            
                            # --- 逻辑改进：扫描第一整行，寻找“保单年度” ---
                            first_row_text = "".join([str(cell) for cell in table_data[0] if cell])
                            
                            if "保单年度" in first_row_text:
                                # 确定新表起始
                                table_count += 1
                                sheet_name = f"方案_{table_count}"
                                current_worksheet = workbook.add_worksheet(sheet_name[:31])
                                current_row_offset = 0
                            
                            # 只要已经有了 Worksheet（代表已经进过“保单年度”页），就持续写入
                            if current_worksheet:
                                written_mark = set()
                                for cell in table_obj.cells:
                                    r0, c0, r1, c1 = int(cell[0]), int(cell[1]), int(cell[2]), int(cell[3])
                                    try:
                                        raw_text = table_data[r0][c0]
                                    except IndexError: continue
                                    
                                    # 判定是否为数字数据行
                                    first_cell_val = str(table_data[r0][0]).strip()
                                    is_data_row = any(char.isdigit() for char in first_cell_val) and "保单年度" not in first_cell_val
                                    
                                    val = clean_val(raw_text) if is_data_row else str(raw_text).replace('\n', ' ')
                                    fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                                    
                                    ex_r0, ex_r1 = r0 + current_row_offset, r1 + current_row_offset
                                    
                                    # 合并单元格逻辑还原
                                    if ex_r1 - ex_r0 > 1 or c1 - c0 > 1:
                                        try:
                                            current_worksheet.merge_range(ex_r0, c0, ex_r1 - 1, c1 - 1, val, fmt)
                                        except: pass
                                        for r in range(ex_r0, ex_r1):
                                            for c in range(c0, c1): written_mark.add((r, c))
                                    else:
                                        if (ex_r0, c0) not in written_mark:
                                            current_worksheet.write(ex_r0, c0, val, fmt)
                                            written_mark.add((ex_r0, c0))
                                
                                # 更新下一张续表的起始位置
                                current_row_offset += len(table_data)
                                current_worksheet.set_column(0, 30, 12)

            # --- 网页展示部分（放在逻辑判断外，确保有预览） ---
            if all_previews:
                st.success(f"🎉 扫描完成！共在网页上预览到 {len(all_previews)} 处表格。")
                for title, df in all_previews:
                    with st.expander(f"👁️ {title} 数据预览"):
                        st.dataframe(df, use_container_width=True)
                
                if table_count > 0:
                    st.download_button(
                        label="📥 下载“逻辑对齐”Excel 文件",
                        data=output.getvalue(),
                        file_name="平安建议书提取结果.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("⚠️ 注意：网页上有预览，但未识别到‘保单年度’标题，Excel 导出可能不完整。")
            else:
                st.warning("⚠️ 全书扫描结束，未发现任何表格。")

    except Exception as e:
        st.error(f"⚠️ 运行出错：{str(e)}")
