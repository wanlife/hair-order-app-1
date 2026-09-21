# web_app.py
import streamlit as st
import os
import io
import config
import process_data

st.set_page_config(page_title="假发订单自动化处理系统", layout="wide")
st.title("✂️ 假发订单数据自动化处理与话术配置系统")

DB_FILE = "seen_database.txt"

def load_local_database(db_path):
    seen_phones, seen_emails = set(), set()
    if os.path.exists(db_path):
        with open(db_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line.startswith("PHONE:"):
                    seen_phones.add(line.replace("PHONE:", ""))
                elif line.startswith("EMAIL:"):
                    seen_emails.add(line.replace("EMAIL:", ""))
    return seen_phones, seen_emails

def save_local_database(db_path, seen_phones, seen_emails):
    with open(db_path, "w", encoding="utf-8") as f:
        for p in sorted(seen_phones):
            f.write(f"PHONE:{p}\n")
        for e in sorted(seen_emails):
            f.write(f"EMAIL:{e}\n")

global_seen_phones, global_seen_emails = load_local_database(DB_FILE)

# 侧边栏显示
st.sidebar.header("💾 历史记忆库状态")
st.sidebar.metric("已记录手机号", len(global_seen_phones))
st.sidebar.metric("已记录邮箱", len(global_seen_emails))

# 话术配置模块（默认自动读取 config.py）
st.header("💬 话术修改模块（默认已有经典版本，可在线修改）")
col_h, col_k = st.columns(2)

with col_h:
    msg_h_text = st.text_area("1. H列 奖励邀请话术 (Hello 姓名,)", value=config.DEFAULT_MSG_H, height=120)

with col_k:
    msg_k_text = st.text_area("2. K列 Ambassador 体验话术", value=config.DEFAULT_MSG_K, height=120)

st.markdown("---")

# 数据导出模块
st.header("📂 数据的导出与处理")
uploaded_file = st.file_uploader("请上传待处理的原始订单 Excel 文件 (.xlsx)", type=["xlsx"])

if uploaded_file and st.button("🚀 开始处理数据并生成结果", type="primary"):
    wb_tar = process_data.run_excel_processing(
        uploaded_file, msg_h_text, msg_k_text, global_seen_phones, global_seen_emails
    )
    
    if wb_tar:
        save_local_database(DB_FILE, global_seen_phones, global_seen_emails)
        output = io.BytesIO()
        wb_tar.save(output)
        
        st.success("✅ 数据处理完成！")
        st.download_button(
            label="📥 点击下载处理后的结果文件 (.xlsx)",
            data=output.getvalue(),
            file_name=f"已处理+{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
