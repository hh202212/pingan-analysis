import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="平安建议书复刻 V4.1", layout="wide")
st.title("🖨️ 平安建议书“复印级”长表提取 V4.1")
st.info("已解决：Index out of range 报错 | 期交保费为0问题 | 跨页数据错位")

# 辅助函数：将PDF里的乱七八糟的文本转成纯数字
def to_pure_num(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return 0
    # 移除换行、空格、逗号
    s = str(val).replace('\n', '').replace(' ', '').replace(',', '')
    # 提取数字、负号、小数点
    res = re.sub(r'[^-0-9.]', '', s)
    try:
        return float(res) if '.' in res else int(res)
    except:
        return 0

with st.sidebar:
    st.header("📊 测算参数")
    principal = st.number_input("初始总投入 (元)", value=400000, step=10000)
    bank_rate = st.number_input("银行假定利率 (%)", value=2.5, step=0.1) / 100

uploaded_file = st.file_uploader("👉 请上传平安建议书 PDF 原件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('⌛ 正在执行动态列校准与长表拼接...'):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                # 样式定义
                num_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': '微软雅黑'})
                text_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': '微软雅黑'})
                
                with pdfplumber.open(uploaded_file) as pdf:
                    schemes = []
                    active_scheme = None
                    
                    # 扩大扫描范围：从第 8 页到最后
                    for page_idx in range(7, len(pdf.pages)):
                        page = pdf.pages[page_idx]
                        tables = page.find_tables()
                        if not tables: continue
                        
                        for table_obj in tables:
                            table_data = table_obj.extract()
                            if not table_data or len(table_data) < 2: continue
                            
                            # --- 步骤 1：寻找这一页表格的“保单年度”起始行 ---
                            data_start_idx = -1
                            premium_col_idx = -1
                            
                            # 全表扫描查找关键字
                            for r_idx, row in enumerate(table_data):
                                row_str = "".join([str(i) for i in row if i])
                                if "保单年度" in row_str or "年龄" in row_str:
                                    # 确定保费所在的列（动态寻找）
                                    for c_idx, cell in enumerate(row):
                                        if cell and ("期交" in cell or "保费" in cell):
                                            premium_col_idx = c_idx
                                    
                                    # 寻找下方第一个数字行
                                    for search_r in range(r_idx + 1, len(table_data)):
                                        if str(table_data[search_r][0]).strip().isdigit():
                                            data_start_idx = search_r
                                            break
                                    if data_start_idx != -1: break
                            
                            if data_start_idx == -1: continue # 没找到年度数据，跳过
                            
                            first_year = int(str(table_data[data_start_idx][0]).strip())
                            
                            # --- 步骤 2：判断是新方案还是续表 ---
                            if first_year == 1:
                                if active_scheme: schemes.append(active_scheme)
                                active_scheme = {"cells": [], "offset": 0, "p_col": premium_col_idx}
                            
                            if active_scheme is not None:
                                # 动态更新保费列索引（防止跨页时列位置变动）
                                if premium_col_idx != -1:
                                    active_scheme["p_col"] = premium_col_idx
                                
                                for cell in table_obj.cells:
                                    r0, c0, r1, c1 = [int(x) for x in cell[:4]]
                                    
                                    # 续表逻辑：跳过重复表头
                                    if first_year > 1 and r0 < data_start_idx:
                                        continue
                                    
                                    # 坐标平移
                                    shift = active_scheme["offset"]
                                    actual_r0 = (r0 - data_start_idx + shift) if first_year > 1 else (r0 + shift)
                                    actual_r1 = (r1 - data_start_idx + shift) if first_year > 1 else (r1 + shift)
                                    
                                    # 确保索引不越界
                                    if r0 >= len(table_data) or c0 >= len(table_data[0]): continue
                                    
                                    raw_text = table_data[r0][c0]
                                    is_data_row = (r0 >= data_start_idx)
                                    
                                    # 数据清洗
                                    val = to_pure_num(raw_text) if is_data_row else str(raw_text).replace('\n', ' ')
                                    
                                    active_scheme["cells"].append({
                                        'r0': actual_r0, 'c0': c0, 'r1': actual_r1, 'c1': c1, 'val': val
                                    })
                                
                                # 更新偏移量
                                active_scheme["offset"] += (len(table_data) - data_start_idx) if first_year > 1 else len(table_data)

                    if active_scheme: schemes.append(active_scheme)

                    # --- 步骤 3：计算银行余额并写入 Excel ---
                    if not schemes:
                        st.error("未能在 PDF 中定位到利益演示数据，请确认是否为导出的电子版 PDF。")
                    else:
                        for idx, s in enumerate(schemes):
                            ws = workbook.add_worksheet(f"方案_{idx+1}")
                            written = set()
                            
                            # 按行组织数据以进行银行利息计算
                            # 我们先将 cell 列表还原成逻辑行，方便计算
                            row_map = {}
                            for c in s["cells"]:
                                r = c['r0']
                                if r not in row_map: row_map[r] = {}
                                row_map[r][c['c0']] = c['val']
                            
                            # 动态计算银行余额
                            bank_val = principal
                            bank_results = {}
                            sorted_rows = sorted(row_map.keys())
                            
                            p_idx = s["p_col"] if s["p_col"] != -1 else 2 # 兜底用第3列
                            
                            for r_num in sorted_rows:
                                first_cell = str(row_map[r_num].get(0, ""))
                                if first_cell.isdigit():
                                    premium = row_map[r_num].get(p_idx, 0)
                                    if not isinstance(premium, (int, float)): premium = 0
                                    bank_val = (bank_val - premium) * (1 + bank_rate)
                                    bank_results[r_num] = round(bank_val, 2)
                            
                            # 正式写入单元格
                            for c in s["cells"]:
                                r0, c0, r1, c1, val = c['r0'], c['c0'], c['r1'], c['c1'], c['val']
                                fmt = num_fmt if isinstance(val, (int, float)) else text_fmt
                                
                                # 插入银行余额列（插在第一列之后）
                                if c0 == 0 and r0 in bank_results:
                                    ws.write(r0, 0, val, num_fmt) # 原保单年度
                                    ws.write(r0, 1, bank_results[r0], num_fmt) # 银行余额
                                    written.add((r0, 0))
                                
                                # 写入原表数据（整体右移一列给银行余额腾位置）
                                target_c0, target_c1 = c0 + 1, c1 + 1
                                if r1 - r0 > 1 or target_c1 - target_c0 > 1:
                                    try: ws.merge_range(r0, target_c0, r1-1, target_c1-1, val, fmt)
                                    except: pass
                                    for r_m in range(r0, r1):
                                        for c_m in range(target_c0, target_c1): written.add((r_m, c_m))
                                elif (r0, target_c0) not in written:
                                    ws.write(r0, target_c0, val, fmt)
                                    written.add((r0, target_c0))
                            
                            # 设置标题
                            ws.write(0, 1, "银行余额(测算)", text_fmt)
                            ws.set_column(0, 30, 15)

            st.success(f"🎉 V4.1 校准成功！已整合 {len(schemes)} 份长表。")
            st.download_button("📥 下载校准版长表 Excel", output.getvalue(), "平安对比测算_V4.1.xlsx")

    except Exception as e:
        st.error(f"❌ 运行异常: {str(e)}")
        st.info("建议检查：PDF 是否包含多层复杂的嵌套表格。")
