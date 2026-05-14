import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# 页面基础配置
st.set_page_config(page_title="平安建议书复刻神器", layout="wide")
st.title("🖨️ 平安建议书表格“复印级”提取 V3.2")
st.info("核心逻辑：整页物理还原 + 强力数字清洗 + 完美保留合并结构")

# 强制数值转换（处理 PDF 中的文本数字）
def clean_to_num(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return ""
    s = str(val).replace('\n', '').replace(' ', '').strip()
    # 匹配纯数字、千分位数字、带小数点的数字
    if re.fullmatch(r'^-?[0-9,.]+$', s):
        num_s = re.sub(r'[^-0-9.]', '', s)
        try:
            return float(num_s) if '.' in num_s else int(num_s)
        except: return s
    return s

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 原稿", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在进行像素级扫描与结构还原...'):
            output = io.BytesIO()
            # 使用 xlsxwriter 引擎确保合并单元格效果
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                # 定义样式：数字格式（解决绿三角）、文本格式、边框
                num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                
                with pdfplumber.open(uploaded_file) as pdf:
                    # 遍历建议书核心利益页（通常 10-25 页）
                    for page_idx in range(9, min(25, len(pdf.pages))):
                        page = pdf.pages[page_idx]
                        tables = page.find_tables()
                        if not tables: continue
                        
                        for t_idx, table in enumerate(tables):
                            # 每页的表格独立一个 Sheet，避免干扰
                            sheet_name = f"第{page_idx+1}页_表{t_idx+1}"
                            worksheet = workbook.add_worksheet(sheet_name[:31])
                            
                            # 提取该表格的文字矩阵
                            table_data = table.extract()
                            if not table_data: continue
                            
                            # 1. 识别合并单元格结构并写入
                            # pdfplumber 的 table.cells 提供格子的起止行列坐标
                            for cell in table.cells:
                                r0, c0, r1, c1 = int(cell[0]), int(cell[1]), int(cell[2]), int(cell[3])
                                
                                # 获取内容并清洗
                                try:
                                    raw_text = table_data[r0][c0]
                                except IndexError: continue
                                
                                # 识别数据行逻辑：第一列如果是纯数字，则全行按数字转换
                                # 特别处理：如果格子里包含“1”，则认为是数据
                                cell_first_col = str(table_data[r0][0]).replace('\n', '').strip()
                                is_data_row = cell_first_col.isdigit() or cell_first_col == "1"
                                
                                val = clean_to_num(raw_text) if is_data_row else str(raw_text).replace('\n', ' ')
                                
                                # 根据单元格是否合并采取不同操作
                                if r1 - r0 > 1 or c1 - c0 > 1:
                                    # 合并写入
                                    fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                                    worksheet.merge_range(r0, c0, r1-1, c1-1, val, fmt)
                                else:
                                    # 普通写入
                                    fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                                    worksheet.write(r0, c0, val, fmt)
                            
                            # 自动设置列宽
                            worksheet.set_column(0, 30, 12)

            st.success("🎉 数据复刻成功！已精准还原 PDF 中的所有框线与合并结构。")
            st.download_button(
                label="📥 下载“复印级”Excel 文件",
                data=output.getvalue(),
                file_name="平安建议书原样提取_V3.2.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"⚠️ 处理出错：{str(e)}。建议刷新网页后再上传测试。")
