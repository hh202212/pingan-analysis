import streamlit as st
import pdfplumber
import pandas as pd
import io

# 页面基础配置
st.set_page_config(page_title="平安建议书对比工具", layout="wide")
st.title("🛡️ 平安人寿建议书 vs 银行转存一元对比表")

# 侧边栏：输入参数
with st.sidebar:
    st.header("📊 测算参数")
    principal = st.number_input("初始投入总本金 (元)", value=400000, step=10000)
    bank_rate = st.number_input("银行定期假定利率 (%)", value=2.5, step=0.1) / 100
    st.info("注：程序默认前8年每年从银行转出5万交保费。")

uploaded_file = st.file_uploader("请上传平安建议书 PDF 文件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('正在深度解析 PDF 表格，请稍候...'):
            with pdfplumber.open(uploaded_file) as pdf:
                all_data = []
                # 遍历建议书核心利益页（通常在10-18页之间）
                for page in pdf.pages[10:18]:
                    table = page.extract_table()
                    if table:
                        df = pd.DataFrame(table)
                        all_data.append(df)
                
                if not all_data:
                    st.error("未能识别到利益演示表格，请确认 PDF 是否为官方导出的原件。")
                else:
                    # 简单展示原始数据，证明“通了”
                    st.success("解析成功！正在重构对比表格...")
                    full_df = pd.concat(all_data, ignore_index=True)
                    
                    # 银行动态计算逻辑
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
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        full_df.to_excel(writer, index=False)
                    st.download_button("📥 点击下载重构后的 Excel", data=output.getvalue(), file_name="对比分析结果.xlsx")
    except Exception as e:
        st.error(f"运行出错：{str(e)}")