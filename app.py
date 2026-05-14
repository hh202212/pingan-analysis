import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# 1. 页面配置
st.set_page_config(page_title="平安建议书原样提取", layout="wide")
st.title("🖨️ 平安建议书“复印级”表格提取 (V5.1)")
st.info("功能：像素级还原合并单元格结构 | 全量数据强制转纯数字 | 移除银行测算")

# 辅助函数：将PDF里的乱七八糟的文本转成纯数字（用于Excel直接计算）
def force_numeric(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return ""
    # 移除换行、空格、逗号、人民币符号
    s = str(val).replace('\n', '').replace(' ', '').replace(',', '').replace('¥', '')
    # 匹配数字、负号、小数点
    res = re.sub(r'[^-0-9.]', '', s)
    try:
        # 如果是纯数字则转换，否则保留原样（防止破坏表头文字）
        if res == "" or (not any(char.isdigit() for char in s)):
            return str(val).replace('\n', ' ')
        return float(res) if '.' in res else int(res)
    except:
        return str(val).replace('\n', ' ')

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 原件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在扫描全书并同步坐标对齐，请稍候...'):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                # 定义样式：数字格式（无绿三角）、文本居中格式
                num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                
                with pdfplumber.open(uploaded_file) as pdf:
                    table_count = 0
                    # 扫描全书有表格的页面
                    for page_idx, page in enumerate(pdf.pages):
                        # 查找页面表格对象（用于处理合并单元格）
                        table_objs = page.find_tables()
                        if not table_objs:
                            continue
                            
                        for t_idx, table_obj in enumerate(table_objs):
                            table_data = table_obj.extract()
                            if not table_data: continue
                            
                            table_count += 1
                            sheet_name = f"P{page_idx+1}_T{t_idx+1}"
                            ws = workbook.add_worksheet(sheet_name[:31])
                            
                            # 记录已写入的坐标，防止合并单元格重复写入冲突
                            written_mark = set()

                            # 核心提取逻辑：遍历每一个定义的单元格
                            for cell in table_obj.cells:
                                # 物理坐标转为整数索引
                                r0, c0, r1, c1 = [int(x) for x in cell[:4]]
                                
                                # 安全读取原始文本
                                try:
                                    raw_text = table_data[r0][c0]
                                except IndexError:
                                    continue
                                
                                # 数据处理：如果是数据行（判断依据：第一列是数字）
                                # 或者该单元格看起来就是数字，则强转
                                val = force_numeric(raw_text)
                                fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                                
                                # 执行 Excel 写入
                                try:
                                    if r1 - r0 > 1 or c1 - c0 > 1:
                                        # 处理合并单元格 (merge_range)
                                        ws.merge_range(r0, c0, r1-1, c1-1, val, fmt)
                                        for r_m in range(r0, r1):
                                            for c_m in range(c0, c1):
                                                written_mark.add((r_m, c_m))
                                    elif (r0, c0) not in written_mark:
                                        # 普通单元格写入
                                        ws.write(r0, c0, val, fmt)
                                        written_mark.add((r0, c0))
                                except:
                                    pass # 忽略合并逻辑冲突
                            
                            ws.set_column(0, 30, 15) # 设置列宽

            if table_count == 0:
                st.warning("⚠️ 未能在 PDF 中提取到有效表格，请确认文件是否为电子版原件。")
            else:
                st.success(f"🎉 提取成功！已按照 PDF 原样复刻了 {table_count} 处表格内容。")
                st.download_button(
                    label="📥 下载“原样复刻”纯数字 Excel",
                    data=output.getvalue(),
                    file_name="平安建议书纯数字提取.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ 运行异常: {str(e)}")
        st.info("提示：请确保上传的是平安官网导出的 PDF 原件，而非扫描件或图片转 PDF。")
