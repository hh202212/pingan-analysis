import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# 1. 页面基础配置
st.set_page_config(page_title="平安建议书复刻神器 V3.4", layout="wide")
st.title("🖨️ 平安建议书“复印级”提取 (V3.4 关键字锚点版)")
st.info("核心功能：自动定位“利益演示表”页面 + 100% 还原合并结构 + 纯数字无损转换")

# 强制数值转换（只对纯数字行生效，且保留文字标题）
def clean_numeric_val(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return ""
    s = str(val).replace('\n', '').replace(' ', '').strip()
    # 如果包含数字，则尝试转码
    if re.search(r'\d', s) and not re.search(r'[\u4e00-\u9fa5]', s):
        num_s = re.sub(r'[^-0-9.]', '', s)
        try:
            return float(num_s) if '.' in num_s else int(num_s)
        except: return s
    return s

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 原件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在全书搜索“利益演示表”并执行像素级复刻...'):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                # 定义样式
                num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                
                found_target = False
                with pdfplumber.open(uploaded_file) as pdf:
                    for page_idx, page in enumerate(pdf.pages):
                        # 检查页面文字，定位“利益演示表”
                        page_text = page.extract_text() or ""
                        if "利益演示表" in page_text:
                            found_target = True
                            
                        # 一旦定位到关键字，处理当前页及后续页面
                        if found_target:
                            tables = page.find_tables()
                            if not tables: continue
                            
                            for t_idx, table in enumerate(tables):
                                sheet_name = f"第{page_idx+1}页_表{t_idx+1}"
                                worksheet = workbook.add_worksheet(sheet_name[:31])
                                table_data = table.extract()
                                
                                written_cells = set()
                                for cell in table.cells:
                                    # 强制坐标转为整数
                                    r0, c0, r1, c1 = int(cell[0]), int(cell[1]), int(cell[2]), int(cell[3])
                                    raw_text = table_data[r0][c0]
                                    
                                    # 判断是否为数据行（第一列为数字）
                                    is_data = str(table_data[r0][0]).strip().isdigit()
                                    val = clean_numeric_val(raw_text) if is_data else str(raw_text).replace('\n', ' ')
                                    
                                    fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                                    
                                    # 处理合并或写入
                                    if r1 - r0 > 1 or c1 - c0 > 1:
                                        worksheet.merge_range(r0, c0, r1-1, c1-1, val, fmt)
                                        for r in range(r0, r1):
                                            for c in range(c0, c1): written_cells.add((r, c))
                                    else:
                                        if (r0, c0) not in written_cells:
                                            worksheet.write(r0, c0, val, fmt)
                                            written_cells.add((r0, c0))
                                
                                worksheet.set_column(0, 30, 12)
                            
                            # 如果页面包含“声明”或“特别提醒”且已识别过表格，通常意味着利益演示结束
                            if "特别提醒" in page_text or "声明" in page_text:
                                # 这里可以选择是否停止，目前设定为继续扫描后续页以防万能险表格分散
                                pass

            if not found_target:
                st.warning("⚠️ 未能在 PDF 中搜索到“利益演示表”关键字，请确认文件是否为平安官方建议书。")
            else:
                st.success("🎉 定位成功！利益演示表已按原始结构复刻完成。")
                st.download_button(
                    label="📥 下载“利益演示表”复印级 Excel",
                    data=output.getvalue(),
                    file_name="平安建议书利益提取_V3.4.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"⚠️ 处理出错：{str(e)}。建议检查 PDF 是否包含非加密的表格。")
