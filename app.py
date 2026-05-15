import streamlit as st
import pdfplumber
import pandas as pd
import io
import os
import re
from openpyxl import load_workbook

# 1. 页面配置：沉浸式橙色主题
st.set_page_config(page_title="钱坤大挪移-资产迁移系统", layout="wide")

# CSS 注入：平安特色配色方案（橙色调，思源黑体）
st.markdown("""
    <style>
    :root { --pa-orange: #FF6600; }
    .main { background-color: #f5f5f5; }
    .stButton>button { background-color: var(--pa-orange); color: white; border-radius: 5px; border: none; }
    .stDownloadButton>button { background-color: #2E7D32; color: white; }
    h1, h2, h3 { color: var(--pa-orange); font-family: 'Source Han Sans SC', 'Microsoft YaHei'; }
    .disclaimer { font-size: 12px; color: #666; margin-top: 20px; border-top: 1px solid #ddd; padding-top: 10px; }
    </style>
    """, unsafe_allow_html=True)

st.title("🛡️ 资产迁移自动测算系统 V8.1")

# --- 核心辅助：V7.0 矩阵缝合提取逻辑 ---
def extract_pdf_to_matrix(uploaded_file):
    final_matrix = []
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            tables = page.find_tables()
            if not tables: continue
            for t_obj in tables:
                data = t_obj.extract()
                if not data or len(data) < 1: continue
                data_start = -1
                for r_idx, row in enumerate(data):
                    if row[0] and str(row[0]).strip() == "1":
                        data_start = r_idx; break
                if data_start != -1:
                    start_idx = 0 if not final_matrix else data_start
                    final_matrix.extend(data[start_idx:])
    return final_matrix

# --- 2. 界面布局 ---
with st.sidebar:
    st.header("📋 业务参数录入")
    cust_name = st.text_input("客户姓名", value="客户")
    cust_age = st.number_input("客户年龄", value=45)
    principal = st.number_input("拟迁移总资产 (元)", value=400000)
    bank_rate = st.number_input("假定定存利率 (%)", value=0.95) / 100
    
    st.write("---")
    pdf_file = st.file_uploader("📤 上传平安建议书 PDF", type="pdf")
    
    # 逻辑优化：自动检查仓库中是否存在模板
    template_path = "template.xlsx"
    has_template = os.path.exists(template_path)
    
    if has_template:
        st.success("✅ 测算模板已就绪 (V6.0内核)")
    else:
        st.error("❌ 仓库未检测到 template.xlsx，请上传。")
        template_file = st.file_uploader("请上传模板 Excel", type="xlsx")
        if template_file:
            with open(template_path, "wb") as f:
                f.write(template_file.getbuffer())
            st.rerun()

    start_calc = st.button("🚀 一键生成迁移建议书")

# --- 3. 计算与拼接逻辑 ---
if start_calc and pdf_file and has_template:
    try:
        with st.spinner('⏳ 正在执行钱坤大挪移，请稍候...'):
            # A. 提取 PDF 数据
            matrix_data = extract_pdf_to_matrix(pdf_file)
            if not matrix_data:
                st.error("未能识别到 PDF 里的利益表，请检查原件。")
                st.stop()
            
            # B. 注入 Excel 模板
            wb = load_workbook(template_path)
            
            # 填充原表参数
            if "原表" in wb.sheetnames:
                ws_raw = wb["原表"]
                ws_raw["D1"] = cust_age
                ws_raw["H1"] = principal
                ws_raw["D2"] = bank_rate
            
            # 填充主险建议书数据
            if "贴主险建议书" in wb.sheetnames:
                ws_paste = wb["贴主险建议书"]
                start_row = 5 
                for r_idx, row_data in enumerate(matrix_data):
                    for c_idx, val in enumerate(row_data):
                        try:
                            clean_val = float(str(val).replace(',', ''))
                        except:
                            clean_val = str(val).replace('\n', ' ')
                        ws_paste.cell(row=start_row + r_idx, column=c_idx + 1, value=clean_val)
            
            # C. 准备导出文件
            output_excel = io.BytesIO()
            wb.save(output_excel)
            
            # --- 4. 预览与展示 ---
            st.success(f"🎉 {cust_name} 的测算模型已生成！")
            
            # 模拟“建议展示”页面的预览
            st.subheader("👁️ 建议书预览 (已按行业配色校准)")
            
            # 使用 Pandas 样式模拟平安配色
            preview_df = pd.DataFrame(matrix_data).head(20)
            styled_df = preview_df.style.set_properties(**{
                'background-color': 'white',
                'color': '#333',
                'border-color': '#eee'
            }).set_table_styles([
                {'selector': 'th', 'props': [('background-color', '#FF6600'), ('color', 'white')]}
            ])
            
            st.table(styled_df)
            
            # 强制插入免责声明
            st.markdown("""
                <div class="disclaimer">
                <strong>温馨提示与免责声明：</strong><br>
                1. 本演示表仅供参考，不构成任何保险合同的组成部分。具体的保险责任、给付条件及犹豫期等权利义务请以正式保险合同条款为准。<br>
                2. 万能账户利率演示部分，高于保证利率的部分是不确定的。实际结算利率以官方每月公布为准。<br>
                3. 银行存款对比数据基于当前利率环境假定，不代表未来实际银行利息走向。<br>
                4. 测算数据仅用于辅助理解资产管理逻辑，请根据自身风险承受能力进行资产配置。
                </div>
            """, unsafe_allow_html=True)
            
            st.write("---")
            st.download_button(
                label="📥 下载完整版 Excel (含自动计算公式)",
                data=output_excel.getvalue(),
                file_name=f"{cust_name}_资产迁移建议书.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"系统运行异常: {str(e)}")

elif not pdf_file and start_calc:
    st.warning("👈 请先在侧边栏上传 PDF 建议书。")
