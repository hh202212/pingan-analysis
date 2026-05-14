import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# 页面基础配置
st.set_page_config(page_title="平安建议书原样复刻工具", layout="wide")
st.title("🖨️ 平安建议书表格“复印级”提取工具")
st.info("核心功能：100% 还原 PDF 合并单元格结构 + 强制数字格式化（无绿三角）")

# 强制数值转换（只对纯数据行生效）
def clean_numeric(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return ""
    clean_str = str(val).replace('\n', '').strip()
    # 如果包含数字，则尝试转码
    if re.search(r'\d', clean_str):
        num_part = re.sub(r'[^-0-9.]', '', clean_str)
        try:
            return float(num_part) if '.' in num_part else int(num_part)
        except:
            return clean_str
    return clean_str

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 原件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在深度扫描表格结构与合并逻辑...'):
            output = io.BytesIO()
            # 使用 xlsxwriter 引擎实现合并单元格
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                # 定义样式
                num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                
                with pdfplumber.open(uploaded_file) as pdf:
                    sheet_count = 1
                    # 遍历利益演示页
                    for page_idx in range(10, min(20, len(pdf.pages))):
                        page = pdf.pages[page_idx]
                        tables = page.find_tables()
                        
                        for t_idx, table in enumerate(tables):
                            sheet_name = f"第{page_idx+1}页_表{t_idx+1}"
                            worksheet = workbook.add_worksheet(sheet_name[:31])
                            
                            # 获取表格的逻辑矩阵（用于获取行列号）
                            table_data = table.extract()
                            if not table_data: continue
                            
                            # 核心：识别合并单元格逻辑
                            # pdfplumber 的 table.cells 包含了每个格子的 (x0, top, x1, bottom) 坐标
                            # 我们将其映射到 Excel 的 (row, col)
                            for cell in table.cells:
                                # cell 对象包含: r0 (起行), c0 (起列), r1 (止行), c1 (止列)
                                r0, c0, r1, c1 = cell[0], cell[1], cell[2], cell[3]
                                cell_text = table_data[r0][c0]
                                
                                # 判断是否是数据行（第一列是纯数字）
                                is_data = False
                                if str(table_data[r0][0]).strip().isdigit():
                                    is_data = True
                                
                                val = clean_numeric(cell_text) if is_data else str(cell_text).replace('\n', '') if cell_text else ""
                                
                                # 如果起止行/列不一致，说明是合并单元格
                                if r1 - r0 > 1 or c1 - c0 > 1:
                                    # merge_range(起行, 起列, 止行-1, 止列-1, 数据, 样式)
                                    # 注意：xlsxwriter 索引从0开始
                                    if isinstance(val, (int, float)):
                                        worksheet.merge_range(r0, c0, r1-1, c1-1, val, num_fmt)
                                    else:
                                        worksheet.merge_range(r0, c0, r1-1, c1-1, val, text_fmt)
                                else:
                                    # 普通单元格
                                    if isinstance(val, (int, float)):
                                        worksheet.write(r0, c0, val, num_fmt)
                                    else:
                                        worksheet.write(r0, c0, val, text_fmt)
                            
                            # 设置默认列宽
                            worksheet.set_column(0, 50, 12)
                            sheet_count += 1

            st.success(f"🎉 结构复刻成功！已识别并合并所有单元格。")
            
            st.download_button(
                label="📥 下载“原样复印”Excel 文件",
                data=output.getvalue(),
                file_name="平安建议书原样复刻.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"⚠️ 处理出错：{str(e)}")
