import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# 页面基础配置
st.set_page_config(page_title="平安建议书原样提取工具", layout="wide")
st.title("📄 平安建议书表格“原样”提取工具")
st.info("功能：保持 PDF 原始嵌套表头结构，仅将数据部分转换为纯数字格式。")

# 核心：精准数字转换（只针对数据，不破坏表头文字）
def try_to_num(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return "" # 保持原样或为空
    
    # 清理掉换行符，方便判断
    clean_str = str(val).replace('\n', '').strip()
    
    # 核心逻辑：如果看起来像数字（包含数字、逗号、百分号等），则尝试转换
    # 我们只针对包含数字的内容进行清理，纯文字（如表头）会跳过
    if re.search(r'\d', clean_str):
        # 去掉逗号、人民币符号、空格等，只留数字、负号和小数点
        numeric_part = re.sub(r'[^-0-9.]', '', clean_str)
        try:
            if '.' in numeric_part:
                return float(numeric_part)
            return int(numeric_part)
        except:
            return clean_str # 转不动就回退到原始文字
    return clean_str

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 文件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('🔍 正在原样提取表格结构...'):
            all_table_data = []
            with pdfplumber.open(uploaded_file) as pdf:
                # 扫描 10-20 页（平安利益表核心区间）
                for page in pdf.pages[10:20]:
                    # 使用最保守的提取策略，保证结构不乱
                    table = page.extract_table()
                    if table:
                        all_table_data.extend(table)
            
            if not all_table_data:
                st.error("未能识别到表格，请确认 PDF 是否为电子版原件。")
            else:
                # --- 核心处理逻辑 ---
                processed_rows = []
                for row in all_table_data:
                    # 对每一行里的每一个格子进行“洗数”
                    # 如果这行第一列是数字（保单年度），说明是数据行，全行洗数
                    # 如果不是数字，说明可能是表头，保持文字原样
                    is_data_row = False
                    if row[0] and str(row[0]).strip().isdigit():
                        is_data_row = True
                    
                    new_row = []
                    for cell in row:
                        if is_data_row:
                            new_row.append(try_to_num(cell))
                        else:
                            # 表头行：仅去掉换行符，不转数字
                            new_row.append(str(cell).replace('\n', '') if cell else "")
                    processed_rows.append(new_row)

                # 转换为 DataFrame（不设置表头，把表头也当成普通行处理，保持原貌）
                df = pd.DataFrame(processed_rows)

                st.success(f"🎉 成功提取！共计 {len(df)} 行。")
                st.dataframe(df, use_container_width=True)

                # --- 导出 Excel ---
                output = io.BytesIO()
                # 强制使用 xlsxwriter 以支持复杂的单元格写入
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    # header=False 表示不使用 Pandas 默认的 0, 1, 2 表头
                    df.to_excel(writer, index=False, header=False, sheet_name='建议书原样提取')
                    
                    workbook = writer.book
                    worksheet = writer.sheets['建议书原样提取']
                    
                    # 定义数字格式（无绿三角）
                    num_fmt = workbook.add_format({'num_format': '#,##0', 'font_name': '微软雅黑'})
                    text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': '微软雅黑'})
                    
                    # 遍历数据，对数字和文字应用不同的格式
                    for r_idx, row in enumerate(processed_rows):
                        for c_idx, cell_val in enumerate(row):
                            if isinstance(cell_val, (int, float)):
                                worksheet.write(r_idx, c_idx, cell_val, num_fmt)
                            else:
                                worksheet.write(r_idx, c_idx, cell_val, text_fmt)
                    
                    # 设置列宽
                    worksheet.set_column(0, len(df.columns)-1, 12)

                st.download_button(
                    label="📥 下载“原样表头”纯数字 Excel",
                    data=output.getvalue(),
                    file_name="建议书原样提取.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"⚠️ 运行出错：{str(e)}")
