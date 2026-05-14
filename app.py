import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# 1. 页面配置
st.set_page_config(page_title="平安建议书提取神器", layout="wide")
st.title("🖨️ 平安建议书表格“复印级”提取 V5.3")
st.info("核心改进：逻辑合并跨页长表 | 保持嵌套表头合并 | 网页实时预览 | 纯数字转换")

# 强力数字清洗
def force_num(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return 0
    s = str(val).replace('\n', '').replace(' ', '').replace(',', '').strip()
    # 只要包含数字且不含大量汉字，就尝试转码
    if re.search(r'\d', s) and len(re.findall(r'[\u4e00-\u9fa5]', s)) < 2:
        res = re.sub(r'[^-0-9.]', '', s)
        try:
            return float(res) if '.' in res else int(res)
        except: return s
    return s

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 原件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在深度扫描并执行逻辑拼接...'):
            output = io.BytesIO()
            all_page_previews = [] # 用于网页预览
            
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                
                with pdfplumber.open(uploaded_file) as pdf:
                    current_ws = None
                    current_row_offset = 0
                    scheme_count = 0
                    
                    # 遍历全书
                    for page_idx, page in enumerate(pdf.pages):
                        tables = page.find_tables()
                        if not tables: continue
                        
                        # 网页预览存一份原始数据
                        raw_data = page.extract_table()
                        if raw_data:
                            all_page_previews.append((f"第{page_idx+1}页", pd.DataFrame(raw_data)))

                        for t_obj in tables:
                            data = t_obj.extract()
                            if not data: continue
                            
                            # --- 判定新表起始：检查第一行是否包含“保单年度” ---
                            first_row_str = "".join([str(c) for c in data[0] if c])
                            
                            if "保单年度" in first_row_str:
                                # 发现新方案：新建 Sheet
                                scheme_count += 1
                                current_ws = workbook.add_worksheet(f"方案_{scheme_count}")
                                current_row_offset = 0
                                # 记录这一张大表的起始列定义，用于后续续表对齐（可选）
                            
                            if current_ws:
                                # 确定数据起始行（用于区分表头和数字）
                                data_start_idx = 0
                                for r_idx, row in enumerate(data):
                                    if str(row[0]).strip().isdigit():
                                        data_start_idx = r_idx
                                        break
                                
                                # 识别并写入单元格逻辑
                                written_mark = set()
                                for cell in t_obj.cells:
                                    r0, c0, r1, c1 = [int(x) for x in cell[:4]]
                                    
                                    # 如果是续表（当前页不包含“保单年度”），跳过重复的表头行
                                    if "保单年度" not in first_row_str and r0 < data_start_idx:
                                        continue
                                    
                                    # 计算在 Excel 中的实际行位置
                                    if "保单年度" not in first_row_str:
                                        # 续表需要减去表头偏移
                                        actual_r0 = r0 - data_start_idx + current_row_offset
                                        actual_r1 = r1 - data_start_idx + current_row_offset
                                    else:
                                        actual_r0 = r0 + current_row_offset
                                        actual_r1 = r1 + current_row_offset
                                    
                                    # 获取内容
                                    raw_text = data[r0][c0]
                                    is_data_cell = (r0 >= data_start_idx)
                                    val = force_num(raw_text) if is_data_cell else str(raw_text).replace('\n', ' ')
                                    fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                                    
                                    # 执行合并或写入
                                    try:
                                        if actual_r1 - actual_r0 > 1 or c1 - c0 > 1:
                                            current_ws.merge_range(actual_r0, c0, actual_r1 - 1, c1 - 1, val, fmt)
                                            for r_m in range(actual_r0, actual_r1):
                                                for c_m in range(c0, c1): written_mark.add((r_m, c_m))
                                        elif (actual_r0, c0) not in written_mark:
                                            current_ws.write(actual_r0, c0, val, fmt)
                                            written_mark.add((actual_r0, c0))
                                    except:
                                        pass
                                
                                # 更新下一页续表的行偏移
                                if "保单年度" not in first_row_str:
                                    current_row_offset += (len(data) - data_start_idx)
                                else:
                                    current_row_offset += len(data)
                                
                                current_ws.set_column(0, 50, 12)

            # --- 结果呈现 ---
            st.success(f"🎉 处理完成！共识别到 {scheme_count} 组完整利益演示方案。")
            
            # 网页预览（永不放弃模式）
            for title, df in all_page_previews:
                with st.expander(f"👁️ {title} 实时数据预览"):
                    st.dataframe(df, use_container_width=True)
            
            if scheme_count > 0:
                st.download_button(
                    label="📥 下载“原样复刻”长表 Excel",
                    data=output.getvalue(),
                    file_name="平安建议书完整提取.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("并未识别到以‘保单年度’开头的表格，请检查PDF页面。")

    except Exception as e:
        st.error(f"❌ 程序运行出错: {str(e)}")
