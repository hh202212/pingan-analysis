import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# 1. 页面配置
st.set_page_config(page_title="平安建议书提取 V5.7", layout="wide")
st.title("🖨️ 平安建议书表格提取 V5.7 (稳定不留白版)")
st.info("核心改进：预览与下载数据强对齐 | 彻底修复下载空表 | 自动跨页缝合长表")

# 强制数字转换函数
def force_num(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return ""
    s = str(val).replace('\n', '').replace(' ', '').replace(',', '').strip()
    # 匹配纯数字/小数点，且不含汉字
    if re.fullmatch(r'^-?[0-9.]+$', s) and not re.search(r'[\u4e00-\u9fa5]', s):
        try:
            return float(s) if '.' in s else int(s)
        except: return s
    return str(val).replace('\n', ' ')

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 原件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在深度缝合长表数据，请稍候...'):
            all_schemes = [] # 存储最终的方案数据 (List of List)
            current_scheme_rows = []
            
            with pdfplumber.open(uploaded_file) as pdf:
                for page in pdf.pages:
                    # 获取这一页最完整的表格矩阵
                    table = page.extract_table()
                    if not table: continue
                    
                    # --- 拼接逻辑判定 ---
                    # 检查这一页的第一行或第二行是否包含“保单年度”
                    is_new_scheme = False
                    for r_idx in range(min(3, len(table))):
                        row_str = "".join([str(c) for c in table[r_idx] if c])
                        if "保单年度" in row_str:
                            is_new_scheme = True
                            break
                    
                    # 如果是新方案，保存上一个
                    if is_new_scheme and current_scheme_rows:
                        all_schemes.append(current_scheme_rows)
                        current_scheme_rows = []
                    
                    # 定位数据行起始点
                    data_start_idx = 0
                    for r_idx, row in enumerate(table):
                        if row and str(row[0]).strip().isdigit():
                            data_start_idx = r_idx
                            break
                    
                    # 如果是续表（不是新方案），则跳过表头行
                    rows_to_add = table if (is_new_scheme or not current_scheme_rows) else table[data_start_idx:]
                    
                    for r in rows_to_add:
                        # 每一格都做基础清洗
                        clean_row = [force_num(cell) for cell in r]
                        current_scheme_rows.append(clean_row)

                # 存入最后一个方案
                if current_scheme_rows:
                    all_schemes.append(current_scheme_rows)

            # --- 结果呈现区 ---
            if not all_schemes:
                st.warning("⚠️ 未能识别到有效的利益演示表数据。")
            else:
                st.success(f"🎉 成功缝合 {len(all_schemes)} 组长表方案！")
                
                # 1. 网页预览 (胡老师要求的预览模式，必须显示拼接后的长表)
                for idx, scheme_data in enumerate(all_schemes):
                    df_preview = pd.DataFrame(scheme_data)
                    with st.expander(f"👁️ 方案 {idx+1} 预览 (已自动无缝拼接，共 {len(df_preview)} 行)"):
                        st.dataframe(df_preview, use_container_width=True)
                
                # 2. 生成 Excel (采用最稳的逻辑，确保不留白)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                    text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
                    
                    for idx, scheme_data in enumerate(all_schemes):
                        sheet_name = f"方案_{idx+1}"
                        # 写入 Excel
                        df_final = pd.DataFrame(scheme_data)
                        df_final.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                        
                        # 获取 worksheet 对象进行样式微调
                        ws = writer.sheets[sheet_name]
                        for r_idx, row in enumerate(scheme_data):
                            for c_idx, val in enumerate(row):
                                fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                                ws.write(r_idx, c_idx, val, fmt)
                        
                        ws.set_column(0, 50, 15)

                st.download_button(
                    label="📥 点击下载“缝合版”纯数字 Excel (确保有数)",
                    data=output.getvalue(),
                    file_name="平安建议书提取_V5.7.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ 运行异常: {str(e)}")
