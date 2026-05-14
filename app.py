import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# 1. 页面配置
st.set_page_config(page_title="平安建议书复刻神器", layout="wide")
st.title("🖨️ 平安建议书表格“复印级”提取 V2.2")
st.markdown("---")

# 强制数值转换（只对纯数字行生效，且保留文字）
def clean_val(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return ""
    s = str(val).replace('\n', '').strip()
    # 只有当它是纯数字或带千分位的数字时才转码
    if re.fullmatch(r'^-?[0-9,.]+$', s):
        num_s = re.sub(r'[^-0-9.]', '', s)
        try:
            return float(num_s) if '.' in num_s else int(num_s)
        except: return s
    return s

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 原件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在扫描建议书（正在解决“白卷”问题）...'):
            output = io.BytesIO()
            all_dfs_for_preview = [] # 用于网页预览
            
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                # 定义 Excel 样式
                num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                
                with pdfplumber.open(uploaded_file) as pdf:
                    # 遍历利益演示页（通常 10-20 页）
                    for page_idx in range(10, min(25, len(pdf.pages))):
                        page = pdf.pages[page_idx]
                        # 核心：使用更稳定的 extract_table 获取完整矩阵
                        table_data = page.extract_table()
                        if not table_data: continue
                        
                        # 为了网页显示，存一份 DataFrame
                        preview_df = pd.DataFrame(table_data)
                        all_dfs_for_preview.append((f"第{page_idx+1}页", preview_df))
                        
                        # 写入 Excel：每个 Page 一个 Sheet，不再分 T1/T2
                        sheet_name = f"第{page_idx+1}页"
                        worksheet = workbook.add_worksheet(sheet_name)
                        
                        # 获取这一页的所有表格结构（用于合并单元格）
                        table_objs = page.find_tables()
                        
                        # 1. 先把所有基础数据填进去（填满每一个格子，防止空表）
                        for r_idx, row in enumerate(table_data):
                            for c_idx, cell_val in enumerate(row):
                                val = clean_val(cell_val)
                                if isinstance(val, (int, float)):
                                    worksheet.write(r_idx, c_idx, val, num_fmt)
                                else:
                                    worksheet.write(r_idx, c_idx, val, text_fmt)
                        
                        # 2. 再执行合并操作（覆盖上面的格子，实现“复刻”效果）
                        for t_obj in table_objs:
                            for cell in t_obj.cells:
                                r0, c0, r1, c1 = int(cell[0]), int(cell[1]), int(cell[2]), int(cell[3])
                                # 如果是跨行或跨列的格子
                                if r1 - r0 > 1 or c1 - c0 > 1:
                                    cell_text = table_data[r0][c0]
                                    val = clean_val(cell_text)
                                    if isinstance(val, (int, float)):
                                        worksheet.merge_range(r0, c0, r1-1, c1-1, val, num_fmt)
                                    else:
                                        worksheet.merge_range(r0, c0, r1-1, c1-1, val, text_fmt)
                        
                        # 设置这一页的列宽
                        worksheet.set_column(0, 30, 12)

            st.success("🎉 解析成功！请查看下方数据预览并下载。")
            
            # --- 网页预览区域 ---
            for title, df in all_dfs_for_preview:
                with st.expander(f"👁️ {title} 数据预览"):
                    st.dataframe(df, use_container_width=True)
            
            st.download_button(
                label="📥 点击下载“满分复刻”Excel 文件",
                data=output.getvalue(),
                file_name="建议书表格提取_修复版.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"⚠️ 处理出错：{str(e)}")
