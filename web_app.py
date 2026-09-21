# web_app.py
import streamlit as st
import os
import io
import pandas as pd
import openpyxl
import config
import process_data

st.set_page_config(
    page_title="假发订单数据自动化处理与自定义导出系统",
    layout="wide",
    initial_sidebar_state="expanded"
)

DB_FILE = "seen_database.txt"

# 1. 记忆库加载与保存
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

# 2. 解决 Streamlit 重复列名报错的工具函数
def make_unique_columns(cols):
    seen = {}
    new_cols = []
    for c in cols:
        col_str = str(c) if c is not None else "Unnamed"
        if col_str in seen:
            seen[col_str] += 1
            new_cols.append(f"{col_str}_{seen[col_str]}")
        else:
            seen[col_str] = 0
            new_cols.append(col_str)
    return new_cols

# 3. 生成范例数据（用于未上传文件时的实时预览）
def get_sample_dataframe(msg_h_template, msg_k_template):
    sample_data = [
        {
            "订单号": "113-1234567-8901234",
            "买家姓名": "Jessica Smith",
            "尺寸": "Body Wave 20 inch",
            "电话": "+1 202-555-0143",
            "邮箱": "jessica.example@gmail.com",
            "H列 奖励邀请话术": msg_h_template.replace("[NAME]", "Jessica Smith") if "[NAME]" in msg_h_template else f"Hello Jessica Smith, {msg_h_template}",
            "K列 Ambassador 体验话术": msg_k_template.replace("[NAME]", "Jessica Smith") if "[NAME]" in msg_k_template else f"Dear Jessica Smith, {msg_k_template}"
        },
        {
            "订单号": "113-7654321-8904321",
            "买家姓名": "Ashley Brown",
            "尺寸": "Straight 24 inch",
            "电话": "+1 202-555-0188",
            "邮箱": "ashley.test@yahoo.com",
            "H列 奖励邀请话术": msg_h_template.replace("[NAME]", "Ashley Brown") if "[NAME]" in msg_h_template else f"Hello Ashley Brown, {msg_h_template}",
            "K列 Ambassador 体验话术": msg_k_template.replace("[NAME]", "Ashley Brown") if "[NAME]" in msg_k_template else f"Dear Ashley Brown, {msg_k_template}"
        }
    ]
    return pd.DataFrame(sample_data)

# ---------------- 侧边栏：历史记忆库管理 ----------------
st.sidebar.title("⚙️ 控制台")
st.sidebar.subheader("💾 历史数据去重库")
st.sidebar.metric("已记录手机号", f"{len(global_seen_phones)} 个")
st.sidebar.metric("已记录邮箱", f"{len(global_seen_emails)} 个")

if st.sidebar.button("🧹 清空历史去重记忆", type="secondary"):
    if os.path.exists(DB_FILE):
        os.remove(DB_FILE)
    global_seen_phones.clear()
    global_seen_emails.clear()
    st.sidebar.success("历史去重记忆已成功重置！")
    st.rerun()

# ---------------- 主界面 ----------------
st.title("✂️ 假发订单数据自动化处理与自定义导出系统")

# 模块 1：话术配置（固定显示）
st.header("1. 💬 话术模板在线配置")
col_h, col_k = st.columns(2)
with col_h:
    msg_h_text = st.text_area("H列 奖励邀请话术 (Hello 姓名,)", value=config.DEFAULT_MSG_H, height=100)
with col_k:
    msg_k_text = st.text_area("K列 Ambassador 体验话术", value=config.DEFAULT_MSG_K, height=100)

st.markdown("---")

# 模块 2：文件上传（固定显示）
st.header("2. 📂 上传订单文件")
uploaded_file = st.file_uploader("请上传待处理的原始订单 Excel 文件 (.xlsx)", type=["xlsx"])

# 提取/加载当前处理的数据集（若未上传则使用模拟示例）
is_real_data = False
if uploaded_file:
    uploaded_file.seek(0)
    with st.spinner("正在智能提取数据、格式化话术并过滤历史重复记录..."):
        wb_processed = process_data.run_excel_processing(
            uploaded_file, msg_h_text, msg_k_text, global_seen_phones, global_seen_emails
        )

    if wb_processed:
        ws = wb_processed.active
        raw_data = list(ws.values)
        if len(raw_data) > 0:
            raw_headers = list(raw_data[0])
            unique_headers = make_unique_columns(raw_headers)
            df_current = pd.DataFrame(raw_data[1:], columns=unique_headers)
            is_real_data = True
            st.success("✅ 真实数据解析成功！")
else:
    st.info("💡 当前未上传文件，下方模块正显示**系统示例范例数据**。您可以先配置筛选与导出列，上传文件后将自动应用该配置。")
    df_current = get_sample_dataframe(msg_h_text, msg_k_text)

st.markdown("---")

# 模块 3：在线查看与搜索（固定显示）
st.header("3. 🔍 数据的在线交互查看与筛选")

c1, c2 = st.columns(2)
with c1:
    search_kw = st.text_input("🔍 关键字搜索（姓名/电话/邮箱/订单号）：", "")
with c2:
    # 查找尺寸/SKU列
    sku_col = [c for c in df_current.columns if "尺寸" in c or "SKU" in c or "Size" in c]
    if sku_col:
        all_skus = [str(x) for x in df_current[sku_col[0]].dropna().unique() if str(x) != "None"]
        selected_skus = st.multiselect("按尺寸/SKU 筛选：", options=all_skus, default=all_skus)
    else:
        selected_skus = []

df_display = df_current.copy()
if search_kw:
    mask = df_display.astype(str).apply(lambda row: row.str.contains(search_kw, case=False).any(), axis=1)
    df_display = df_display[mask]
if sku_col and selected_skus:
    df_display = df_display[df_display[sku_col[0]].astype(str).isin(selected_skus)]

status_label = "真实数据" if is_real_data else "范例数据"
st.write(f"📊 当前筛选呈现 **{len(df_display)}** 条记录（【{status_label}】共 {len(df_current)} 条）:")
st.dataframe(df_display, use_container_width=True, height=280)

st.markdown("---")

# 模块 4：自定义导出列与范例浏览（固定显示）
st.header("4. 🎯 自定义导出列配置与范例预览")

all_cols = list(df_current.columns)
selected_cols = st.multiselect(
    "请勾选你需要导出的 Excel 列（勾选后可在下方的范例效果中实时查看导出样式）：",
    options=all_cols,
    default=all_cols
)

if not selected_cols:
    st.warning("⚠️ 请至少选择一列导出！")
else:
    # 渲染当前勾选列的范例/真实预览
    st.subheader("👁️ 导出结果范例效果预览")
    df_export_preview = df_display[selected_cols]
    st.dataframe(df_export_preview.head(5), use_container_width=True)

    # 下载/导出区域
    if is_real_data:
        # 实际导出逻辑
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            df_export_preview.to_excel(writer, index=False, sheet_name="处理后的订单数据")
        
        save_local_database(DB_FILE, global_seen_phones, global_seen_emails)
        
        st.download_button(
            label="📥 导出并下载定制版 Excel 文件",
            data=output_buffer.getvalue(),
            file_name=f"自定义导出_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    else:
        st.button("📥 导出并下载定制版 Excel 文件（请先上传真实文件）", disabled=True, type="primary")
