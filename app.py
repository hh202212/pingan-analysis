import streamlit as st
import pdfplumber
import pandas as pd
import io
import os
import re
from openpyxl import load_workbook

# 1. 页面配置：锁定视觉标准
st.set_page_config(page_title="资产迁移自动测算系统", layout="wide")

# 全局样式注入：锁定思源黑体与平安橙
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@400;700&display=swap');
    :root { --pa-orange: #FF6600; }
    
    html, body, [class*="css"], h1, h2, h3, div, span, p, td, th { 
        font-family: 'Noto Sans SC', '思源黑体', sans-serif !important; 
    }
    
    .stButton>button { 
        background-color: var(--pa-orange); 
        color: white; 
        border-radius: 8px; 
        height: 3.5em;
        font-weight: bold;
        border: none;
    }
    
    .disclaimer-box { 
        font-size: 13px; 
        color: #555; 
        margin-top: 30px; 
        border-left: 5px solid var(--pa-orange); 
        background-color: #fff8f5;
        padding: 20px; 
        border-radius: 0 10px 10px 0;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("🛡️ 钱坤大挪移-资产迁移系统 (V8.7)")

# --- 核心逻辑：绝对路径锁定 (物理寻址) ---
def get_template_path():
    # 获取 app.py 所在的绝对目录
    base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, "template.xlsx")

def extract_scheme_2(uploaded_file):
    all_schemes = []
    current_data = []
    pdf_bytes = io.BytesIO(uploaded_file.getvalue())
    try:
        with pdfplumber.open(pdf_bytes) as pdf:
            for page in pdf.pages:
                tables = page.find_tables()
                if not tables: continue
                for t_obj in tables:
                    data = t_obj.extract()
                    if not data: continue
                    data_start = -1
                    for r_idx, row in enumerate(data):
                        if row[0] and str(row[0]).strip().isdigit():
                            data_start = r_idx; break
                    if data_start != -1:
                        first_year = int(str(data[data_start][0]).strip())
                        if first_year == 1:
                            if current_data: all_schemes.append(current_data)
                            current_data = data[data_start:]
                        elif current_data:
                            current_data.extend(data[data_start:])
            if current_data: all_schemes.append(current_data)
        if len(all_schemes) >= 2: return all_schemes[1]
        return all_schemes[0] if all_schemes else None
    except: return None

# --- 2. 侧边栏与录入 ---
with st.sidebar:
    st.header("📋 方案核心参数")
    cust_name = st.text_input("客户姓名", value="苏女士")
    cust_age = st.number_input("客户年龄", value=45)
    principal = st.number_input("拟迁移总资产 (元)", value=400000)
    bank_rate = st.number_input("定存利率 (%)", value=0.95) / 100
    
    st.write("---")
    pdf_file = st.file_uploader("📤 上传 PDF 建议书", type="pdf")
    
    # 路径锁定逻辑检查
    target_path = get_template_path()
    has_temp = os.path.exists(target_path)
    
    if has_temp:
        st.success("✅ 物理路径已锁定，模板就绪")
    else:
        st.error(f"❌ 物理路径未找到模板")
        # 调试信息：列出当前目录下所有文件，帮助排除万一
        st.info(f"当前目录内容: {os.listdir(os.path.dirname(os.path.abspath(__file__)))}")

    go = st.button("🚀 执行钱坤大挪移")

# --- 3. 处理逻辑 ---
if go and pdf_file and has_temp:
    try:
        with st.spinner('正在同步数据层...'):
            matrix = extract_scheme_2(pdf_file)
            if not matrix:
                st.error("无法解析数据。")
                st.stop()
            
            wb = load_workbook(target_path)
            
            # 原表注入
            if "原表" in wb.sheetnames:
                ws = wb["原表"]
                ws["D1"], ws["H1"], ws["D2"] = cust_age, principal, bank_rate
            
            # 数据贴入 (第5行开始)
            if "贴主险建议书" in wb.sheetnames:
                ws_p = wb["贴主险建议书"]
                for r_idx, row in enumerate(matrix):
                    for c_idx, val in enumerate(row):
                        if val is None or str(val).lower() == "none": cv = ""
                        else:
                            try: cv = float(str(val).replace(',', '').replace(' ', ''))
                            except: cv = str(val).replace('\n', ' ')
                        ws_p.cell(row=5 + r_idx, column=c_idx + 1, value=cv)
            
            out = io.BytesIO()
            wb.save(out)
            
            # 结果预览
            st.success(f"🎉 {cust_name} 的测算模型已完美生成！")
            if "建议展示" in wb.sheetnames:
                preview = []
                for row in wb["建议展示"].iter_rows(min_row=1, max_row=35, values_only=True):
                    preview.append(row)
                st.dataframe(pd.DataFrame(preview).fillna(""), use_container_width=True, height=600)

            st.download_button("📥 下载完整测算 Excel (内含公式)", out.getvalue(), f"{cust_name}_测算建议书.xlsx")

    except Exception as e:
        st.error(f"发生错误: {e}")
