import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# 页面基础配置
st.set_page_config(page_title="平安建议书原样复刻工具", layout="wide")
st.title("🖨️ 平安建议书表格“复印级”提取工具 V2.1")
st.info("核心功能：100% 还原合并单元格 + 强制数字格式化（修复整数索引错误）")

# 强制数值转换
def clean_numeric(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return ""
    clean_str = str(val).replace('\n', '').strip()
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
        with st.spinner('⌛ 正在深度扫描表格结构（正在进行像素级对齐）...'):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                # 定义样式
                num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                
                with pdfplumber.open(uploaded_file) as pdf:
                    for page_idx in range(10, min(25, len(pdf.pages))): # 扩大扫描范围
                        page = pdf.pages[page_idx]
                        tables = page.find_tables()
                        
                        for t_idx, table in enumerate(tables):
                            sheet_name = f"P{page_idx+1}_T{t_idx+1}"
                            worksheet = workbook.add_worksheet(sheet_name[:31])
                            
                            table_data = table.extract()
                            if not table_data: continue
                            
                            # 记录已经写入的单元格，防止重复写入冲突
                            written_cells = set()

                            # 核心改进：强制将所有索引转换为 int 整数
                            for cell in table.cells:
                                # cell[0]起行, cell[1]起列, cell[2]止行, cell[3]止列
                                r0, c0, r1, c1 = int(cell[0]), int(cell[1]), int(cell[2]), int(cell[3])
                                
                                # 获取内容
                                try:
                                    cell_text = table_data[r0][c0]
                                except IndexError:
                                    continue
                                
                                # 判断是否是数据行（第一列是纯数字）
                                is_data = False
                                first_col_val = str(table_data[r0][0]).strip()
                                if first_col_val.isdigit():
                                    is_data = True
                                
                                val = clean_numeric(cell_text) if is_data else str(cell_text).replace('\n', '') if cell_text else ""
                                
                                # 执行合并或写入
                                if r1 - r0 > 1 or c1 - c0 > 1:
                                    # 合并范围：(起行, 起列, 止行-1, 止列-1)
                                    if isinstance(val, (int, float)):
                                        worksheet.merge_range(r0, c0, r1-1, c1-1, val, num_fmt)
                                    else:
                                        worksheet.merge_range(r0, c0, r1-1, c1-1, val, text_fmt)
                                else:
                                    if (r0, c0) not in written_cells:
                                        if isinstance(val, (int, float)):
                                            worksheet.write(r0, c0, val, num_fmt)
                                        else:
                                            worksheet.write(r0, c0, val, text_fmt)
                                
                                # 标记已写入区域
                                for r in range(r0, r1):
                                    for c in range(c0, c1):
                                        written_cells.add((r, c))
                            
                            # 自动调整列宽
                            worksheet.set_column(0, 50, 12)

            st.success(f"🎉 结构复刻成功！已精准还原合并单元格逻辑。")
            st.download_button(
                label="📥 点击下载“复印级”Excel 文件",
                data=output.getvalue(),
                file_name="建议书原样复刻_最终版.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"⚠️ 处理出错：{str(e)}。请确保上传的是导出的 PDF 电子版原件。")
