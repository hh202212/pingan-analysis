import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="平安建议书对比工具", layout="wide")
st.title("🛡️ 平安人寿建议书 vs 银行转存一元对比表")

# 辅助函数：清理数字中的杂质
def force_numeric(val):
    if val is None or str(val).strip() == "" or str(val).lower() == "none":
        return 0
    # 提取数字和小数点
    clean_val = re.sub(r'[^-0-9.]', '', str(val))
    try:
        if '.' in clean_val:
            return float(clean_val)
        return int(clean_val)
    except:
        return 0

# 辅助函数：确保列名唯一
def make_unique(labels):
    new_labels = []
    for i, label in enumerate(labels):
        if not label or label == 'None':
            label = f"未知列_{i}"
        count = labels[:i].count(label)
        new_labels.append(f"{label}_{count}" if count > 0 else label)
    return new_labels

with st.sidebar:
    st.header("📊 测算参数")
    principal = st.number_input("初始投入总本金 (元)", value=400000, step=10000)
    bank_rate = st.number_input("银行定期假定利率 (%)", value=2.5, step=0.1) / 100
    st.info("说明：程序会解析建议书利益表并计算银行转存对比。")

uploaded_file = st.file_uploader("请上传平安建议书 PDF 文件", type="pdf")

if uploaded_file:
    try:
        with st.spinner('正在深度清洗复杂表头并重构数据...'):
            with pdfplumber.open(uploaded_file) as pdf:
                all_pages_data = []
                # 遍历利益演示核心页
                for page in pdf.pages[10:20]: 
                    table = page.extract_table()
                    if table:
                        df_raw = pd.DataFrame(table)
                        
                        # 1. 自动合并前3-4行复杂的嵌套表头
                        header_rows = 4 
                        flat_headers = []
                        for col_idx in range(df_raw.shape[1]):
                            parts = []
                            for row in range(header_rows):
                                cell = str(df_raw.iloc[row, col_idx]).replace('\n', '').strip()
                                if cell and cell != 'None':
                                    parts.append(cell)
                            # 使用下划线连接，并去重
                            header_name = "_".join(dict.fromkeys(parts))
                            flat_headers.append(header_name)
                        
                        # 2. 强制处理重复列名，防止报错
                        df_raw.columns = make_unique(flat_headers)
                        
                        # 3. 剔除表头行，只留数据
                        df_clean = df_raw.iloc[header_rows:].copy()
                        
                        # 4. 精准过滤：只保留第一列是保单年度（纯数字）的行
                        df_clean = df_clean[df_clean.iloc[:, 0].apply(lambda x: str(x).isdigit())]
                        
                        if not df_clean.empty:
                            all_pages_data.append(df_clean)
                
                if not all_pages_data:
                    st.error("未能识别到有效数字行。请确认 PDF 是否为官方原始电子版。")
                else:
                    final_df = pd.concat(all_pages_data, ignore_index=True)
                    
                    # 5. 转换全表为纯数字
                    for col in final_df.columns:
                        final_df[col] = final_df[col].apply(force_numeric)
                    
                    # 6. 自动寻找保费列（查找包含关键词的列）
                    premium_col = [c for c in final_df.columns if '期交' in c or '保费' in c]
                    
                    # 7. 银行动态复利计算
                    bank_balances = []
                    curr_bank = principal
                    for i, row in final_df.iterrows():
                        # 如果找不着保费列，默认用 0（因为盛世金越后期没保费）
                        p = row[premium_col[0]] if premium_col else 0
                        curr_bank = (curr_bank - p) * (1 + bank_rate)
                        bank_balances.append(round(curr_bank, 2))
                    
                    # 插入计算结果
                    final_df.insert(0, '银行账户余额(测算)', bank_balances)
                    
                    st.success("🎉 数据重构成功！")
                    st.dataframe(final_df)
                    
                    # 8. 导出纯数字 Excel
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        final_df.to_excel(writer, index=False, sheet_name='对比分析')
                    st.download_button("📥 下载完美版纯数字 Excel", data=output.getvalue(), file_name="对比分析结果.xlsx")

    except Exception as e:
        st.error(f"处理出错：{str(e)}")
