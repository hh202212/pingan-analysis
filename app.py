import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="平安建议书提取 V5.0", layout="wide")
st.title("🛡️ 平安建议书表格提取专家 (纯净提取版)")
st.info("功能：100% 还原 PDF 表头及合并结构 | 剔除所有银行测算逻辑 | 强制数字格式转换")

# 强制转换数字，解决绿三角
def clean_num(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return ""
    # 去掉逗号、换行、空格
    s = str(val).replace('\n', '').replace(',', '').strip()
    # 匹配纯数字、负号、小数点
    res = re.sub(r'[^-0-9.]', '', s)
    try:
        return float(res) if '.' in res else int(res)
    except:
        return s # 如果含有文字则返回文字（表头）

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 原件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('🔍 正在原样复刻表格结构，请稍候...'):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                # 定义样式
                num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                
                with pdfplumber.open(uploaded_file) as pdf:
                    sheet_count = 0
                    for page_idx, page in enumerate(pdf.pages):
                        page_text = page.extract_text() or ""
                        # 只要页面有“利益”字样或有表格，就尝试提取
                        tables = page.find_tables()
                        if not tables: continue
                        
                        sheet_count += 1
                        ws = workbook.add_worksheet(f"第{page_idx+1}页")
                        
                        for table_obj in tables:
                            table_data = table_obj.extract()
                            if not table_data: continue
                            
                            written_cells = set() # 记录已写的格子，防止重复写入
                            
                            # 1. 遍历单元格，执行复刻
                            for cell in table_obj.cells:
                                r0, c0, r1, c1 = [int(x) for x in cell[:4]]
                                raw_val = table_data[r0][c0]
                                
                                # 智能清洗：如果是年度/年龄后的数据，全转数字
                                is_data = False
                                if str(table_data[r0][0]).strip().isdigit():
                                    is_data = True
                                
                                val = clean_num(raw_val) if is_data else str(raw_val).replace('\n', ' ')
                                fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                                
                                # 处理合并逻辑
                                try:
                                    if r1 - r0 > 1 or c1 - c0 > 1:
                                        ws.merge_range(r0, c0, r1-1, c1-1, val, fmt)
                                        for r in range(r0, r1):
                                            for c in range(c0, c1): written_cells.add((r, c))
                                    elif (r0, c0) not in written_cells:
                                        ws.write(r0, c0, val, fmt)
                                        written_cells.add((r0, c0))
                                except:
                                    pass # 忽略合并坐标冲突
                        
                        ws.set_column(0, 30, 15) # 设置列宽

            if sheet_count == 0:
                st.warning("⚠️ 未能从 PDF 中提取出表格，请确认文件是否为电子版原件。")
            else:
                st.success(f"🎉 提取完成！已复刻 {sheet_count} 页表格内容。")
                st.download_button("📥 下载原样复刻纯数字 Excel", output.getvalue(), "建议书原样提取_V5.xlsx")

    except Exception as e:
        st.error(f"❌ 运行异常: {str(e)}")
