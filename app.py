import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# 页面基础配置
st.set_page_config(page_title="平安建议书提取 V3.3", layout="wide")
st.title("🛡️ 平安建议书数据提取 (V3.3 稳定版)")
st.info("核心改进：放弃不稳定的合并算法，采用“全量提取”逻辑，确保下载不再是空表。")

# 强制数值转换（彻底解决绿三角）
def force_num(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return ""
    s = str(val).replace('\n', '').replace(' ', '').strip()
    # 匹配数字、负号、逗号、小数点
    if re.fullmatch(r'^-?[0-9,.]+$', s):
        num_s = re.sub(r'[^-0-9.]', '', s)
        try:
            return float(num_s) if '.' in num_s else int(num_s)
        except: return s
    return s

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF", type="pdf")

if uploaded_file:
    try:
        with st.spinner('🔍 正在全量提取数据，请稍候...'):
            output = io.BytesIO()
            preview_list = []
            
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                # 定义 Excel 样式
                num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
                
                with pdfplumber.open(uploaded_file) as pdf:
                    # 扫描全书，防止漏页
                    for i, page in enumerate(pdf.pages):
                        # 尝试多种提取策略
                        table = page.extract_table()
                        if not table:
                            continue
                        
                        # 转换为 DataFrame 方便处理
                        df = pd.DataFrame(table)
                        # 数据预览存储
                        preview_list.append((f"第{i+1}页", df))
                        
                        # 写入 Excel
                        sheet_name = f"第{i+1}页"
                        worksheet = workbook.add_worksheet(sheet_name)
                        
                        for r_idx, row in df.iterrows():
                            # 判断是否为数据行（第一列包含数字）
                            is_data = False
                            first_cell = str(row[0]) if row[0] else ""
                            if re.search(r'\d', first_cell):
                                is_data = True
                                
                            for c_idx, cell_val in enumerate(row):
                                val = force_num(cell_val) if is_data else str(cell_val).replace('\n', ' ')
                                
                                # 写入单元格
                                if isinstance(val, (int, float)):
                                    worksheet.write(r_idx, c_idx, val, num_fmt)
                                else:
                                    worksheet.write(r_idx, c_idx, val, text_fmt)
                        
                        # 设置自适应列宽
                        worksheet.set_column(0, 30, 12)

            if not preview_list:
                st.warning("⚠️ 在 PDF 中未发现标准表格，请确认是否为直接导出的电子版。")
            else:
                st.success(f"🎉 提取成功！共发现 {len(preview_list)} 页包含表格。")
                
                # 网页预览
                for title, df in preview_list:
                    with st.expander(f"👁️ {title} 内容预览（只要这里有数，下载就有数）"):
                        st.dataframe(df, use_container_width=True)
                
                st.download_button(
                    label="📥 下载纯数字版 Excel (无绿三角)",
                    data=output.getvalue(),
                    file_name="平安建议书提取结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"⚠️ 运行出错：{str(e)}")
