import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# 1. 页面配置
st.set_page_config(page_title="平安建议书提取神器", layout="wide")
st.title("🛡️ 平安建议书数据提取 (稳健回归版 V3.6)")
st.info("说明：此版本优先保证“数据不漏”和“数字格式正确”。遇到新的‘保单年度 1’会自动分表。")

# 辅助函数：强力转数字
def to_num(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return ""
    s = str(val).replace('\n', '').replace(' ', '').strip()
    # 匹配数字格式
    if re.fullmatch(r'^-?[0-9,.]+$', s):
        num_s = re.sub(r'[^-0-9.]', '', s)
        try:
            return float(num_s) if '.' in num_s else int(num_s)
        except: return s
    return s

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 原件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('🔍 正在全书扫描“利益演示表”并校准数据...'):
            all_pages_data = []
            with pdfplumber.open(uploaded_file) as pdf:
                found_target = False
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text() or ""
                    # 锚点识别：发现“利益演示表”才开始抓取
                    if "利益演示表" in text:
                        found_target = True
                    
                    if found_target:
                        table = page.extract_table()
                        if table:
                            all_pages_data.append(table)
                        
                        # 遇到声明或提醒，通常意味着这一段利益表结束了
                        if "特别提醒" in text or "重要提示" in text:
                            # 这里不停止，继续往后看有没有附加险的利益表
                            pass
            
            if not all_pages_data:
                st.warning("⚠️ 未能定位到利益演示表，请确认 PDF 是否包含该关键字。")
            else:
                # --- 核心逻辑：按“保单年度 1”进行分组 ---
                sections = []
                current_section = []
                
                for table in all_pages_data:
                    for row in table:
                        # 检查第一列是否是新的“保单年度 1”
                        first_col = str(row[0]).strip() if row[0] else ""
                        # 兼容处理：有些合并格子里是 "保单年度\n1"
                        if (first_col == "1" or "年度1" in first_col.replace('\n','')) and current_section:
                            # 只有当当前 section 里已经有数据行时，才切分（避免把表头切断）
                            # 我们简单判断：如果当前 section 最后几行里有数字，就切分
                            sections.append(current_section)
                            current_section = []
                        current_section.append(row)
                
                if current_section:
                    sections.append(current_section)

                # --- 写入 Excel ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                    text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})

                    for idx, section in enumerate(sections):
                        df = pd.DataFrame(section)
                        sheet_name = f"方案_{idx+1}"
                        worksheet = workbook.add_worksheet(sheet_name)
                        
                        for r_idx, row in df.iterrows():
                            # 判断是否为数据行（第一列包含数字）
                            is_data = False
                            if re.search(r'\d', str(row[0])):
                                is_data = True
                            
                            for c_idx, cell_val in enumerate(row):
                                val = to_num(cell_val) if is_data else str(cell_val).replace('\n', ' ')
                                
                                # 写入
                                if isinstance(val, (int, float)):
                                    worksheet.write(r_idx, c_idx, val, num_fmt)
                                else:
                                    worksheet.write(r_idx, c_idx, val, text_fmt)
                        
                        worksheet.set_column(0, 30, 12)

                st.success(f"🎉 提取成功！共识别到 {len(sections)} 组产品利益。")
                
                # 网页预览预览
                with st.expander("👁️ 点击预览提取到的数据"):
                    for i, s in enumerate(sections):
                        st.write(f"方案 {i+1}")
                        st.table(pd.DataFrame(s).head(10)) # 只显示前10行预览

                st.download_button(
                    label="📥 下载纯数字版 Excel",
                    data=output.getvalue(),
                    file_name="平安提取结果_稳健版.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"⚠️ 运行出错：{str(e)}")
