import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# 1. 页面配置
st.set_page_config(page_title="平安建议书复刻神器", layout="wide")
st.title("🖨️ 平安建议书表格“复印级”提取 V2.3")
st.info("已修复索引越界错误 | 完美还原合并单元格 | 强制数字格式化")

# 核心：纯净数字转换
def clean_val(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return ""
    s = str(val).replace('\n', '').strip()
    # 匹配数字、负号、逗号、小数点
    if re.fullmatch(r'^-?[0-9,.]+$', s):
        num_s = re.sub(r'[^-0-9.]', '', s)
        try:
            return float(num_s) if '.' in num_s else int(num_s)
        except: return s
    return s

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 原件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在深度校准表格坐标与内容...'):
            output = io.BytesIO()
            all_previews = [] # 用于网页预览
            
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                # 定义 Excel 样式
                num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                
                with pdfplumber.open(uploaded_file) as pdf:
                    # 扫描利益演示页（通常 10-25 页）
                    for page_idx in range(10, min(25, len(pdf.pages))):
                        page = pdf.pages[page_idx]
                        # 查找页面上的所有表格
                        tables = page.find_tables()
                        if not tables: continue
                        
                        for t_idx, table in enumerate(tables):
                            # 每个表格一个独立 Sheet，或合并到一页（此处采用 Sheet 区分防止冲突）
                            sheet_name = f"P{page_idx+1}_T{t_idx+1}"
                            worksheet = workbook.add_worksheet(sheet_name[:31])
                            
                            # 直接获取该表格的数据矩阵
                            table_data = table.extract()
                            if not table_data: continue
                            
                            # 存一份用于预览
                            all_previews.append((sheet_name, pd.DataFrame(table_data)))
                            
                            # 记录已写入的合并区域，避免重复写入报错
                            merged_cells = set()

                            # 1. 先识别合并逻辑
                            for cell in table.cells:
                                # r0起行, c0起列, r1止行, c1止列
                                r0, c0, r1, c1 = int(cell[0]), int(cell[1]), int(cell[2]), int(cell[3])
                                
                                # 安全获取文本
                                try:
                                    raw_text = table_data[r0][c0]
                                except IndexError: continue
                                
                                # 判断是否为数据行（第一列为数字）
                                is_data = str(table_data[r0][0]).strip().isdigit()
                                val = clean_val(raw_text) if is_data else str(raw_text).replace('\n', '')
                                
                                if r1 - r0 > 1 or c1 - c0 > 1:
                                    # 执行合并
                                    fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                                    worksheet.merge_range(r0, c0, r1-1, c1-1, val, fmt)
                                    # 标记已合并格子
                                    for r in range(r0, r1):
                                        for c in range(c0, c1):
                                            merged_cells.add((r, c))
                                else:
                                    # 普通单元格（且未被合并覆盖）
                                    if (r0, c0) not in merged_cells:
                                        fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                                        worksheet.write(r0, c0, val, fmt)
                            
                            worksheet.set_column(0, 30, 12)

            st.success("🎉 数据重构完成！")
            
            # --- 网页展示预览 ---
            for title, df in all_previews:
                with st.expander(f"👁️ {title} 预览"):
                    st.dataframe(df, use_container_width=True)
            
            st.download_button(
                label="📥 下载“原样复印”Excel 文件",
                data=output.getvalue(),
                file_name="建议书表格提取_V2.3.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"⚠️ 处理出错：{str(e)}")
