import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# 页面基础配置
st.set_page_config(page_title="平安建议书对比工具", layout="wide")
st.title("🛡️ 平安人寿建议书 vs 银行转存一元对比表")

# --- 新增：强制数值转换函数 ---
def to_numeric(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return 0
    # 去掉逗号、空格等杂质，只留数字、负号和小数点
    clean_val = re.sub(r'[^-0-9.]', '', str(val).replace('\n', ''))
    try:
        return float(clean_val) if '.' in clean_val else int(clean_val)
    except:
        return 0

# 侧边栏：输入参数
with st.sidebar:
    st.header("📊 测算参数")
    principal = st.number_input("初始投入总本金 (元)", value=400000, step=10000)
    bank_rate = st.number_input("银行定期假定利率 (%)", value=2.5, step=0.1) / 100
    st.info("注：程序默认前8年每年从银行转出5万交保费。")

uploaded_file = st.file_uploader("请上传平安建议书 PDF 文件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('正在解析 PDF 并转换数字格式...'):
            with pdfplumber.open(uploaded_file) as pdf:
                all_data = []
                # 遍历建议书核心利益页（通常在10-18页之间）
                for page in pdf.pages[10:18]:
                    table = page.extract_table()
                    if table:
                        df = pd.DataFrame(table)
                        all_data.append(df)
                
                if not all_data:
                    st.error("未能识别到利益演示表格。")
                else:
                    st.success("解析成功！")
                    full_df = pd.concat(all_data, ignore_index=True)
                    
                    # --- 核心新增：将全表所有单元格强制转为数字 ---
                    for col in full_df.columns:
                        full_df[col] = full_df[col].apply(to_numeric)
                    
                    # 银行动态计算逻辑（保持您的原始逻辑不变）
                    bank_flows = []
                    curr_bank = principal
                    for year in range(len(full_df)):
                        out_flow = 50000 if year < 8 else 0
                        curr_bank = (curr_bank - out_flow) * (1 + bank_rate)
                        bank_flows.append(round(curr_bank, 2))
                    
                    full_df['银行账户余额'] = bank_flows
                    st.dataframe(full_df)
                    
                    # 导出 Excel
                    output = io.BytesIO()
                    # 使用您代码里的 openpyxl 引擎
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        full_df.to_excel(writer, index=False)
                    st.download_button("📥 下载纯数字版 Excel", data=output.getvalue(), file_name="对比分析结果.xlsx")
    except Exception as e:
        st.error(f"运行出错：{str(e)}")
