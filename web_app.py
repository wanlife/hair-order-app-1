# web_app.py
import streamlit as st
import os
import io
import glob
import datetime
import traceback
import pandas as pd
import openpyxl
import config
import process_data

st.set_page_config(
    page_title="假发跨境电商订单智能处理系统 - Enterprise SaaS",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== 现代商业 SaaS UI 样式注入 ====================
st.markdown("""
<style>
    .stApp {
        background-color: #f8fafc;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e2e8f0;
    }
    .stTextInput input, .stTextArea textarea, .stSelectbox select {
        border-radius: 8px !important;
        border: 1px solid #cbd5e1 !important;
    }
    .stButton button, .stDownloadButton button {
        border-radius: 8px !important;
        font-weight: 500 !important;
        transition: all 0.2s ease-in-out;
    }
    [data-testid="stDataFrame"] {
        border-radius: 8px;
        overflow: hidden;
        border: 1px solid #e2e8f0;
        background-color: white;
    }
</style>
""", unsafe_allow_html=True)

UPLOAD_DIR = "历史输入文件"
OUTPUT_DIR = "处理完成"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 修复：直接定义去重数据库文件名，不依赖 process_data 内部属性
DB_FILE = "seen_database.txt"

# 加载本地去重记忆库
global_seen_phones, global_seen_emails = process_data.load_local_database(DB_FILE)

def get_default_sample_df():
    return pd.DataFrame([
        {
            "订单号": "113-1234567-8901234",
            "尺寸": "13X6 ST-Body Wave 20",
            "订单日期": "2026-09-15",
            "姓名": "Jessica Smith",
            "电话": "+1 (202) 555-0143",
            "回复情况": "Sep 2026",
            "邮箱": "jessica@icloud.com",
            "Hello 姓名,": '="Hello "&D2&","&CHAR(10)&CHAR(10)&"Thanks for purchasing..."',
            "": "",
            "": "",
            "Ambassador 话术": '="Hello "&D2&","&CHAR(10)&CHAR(10)&"Thank you for choosing..."'
        }
    ])

df_current = get_default_sample_df()
is_real_data = False
wb_processed = None
last_output_filename = ""

# ---------------- 侧边栏控制台 ----------------
with st.sidebar:
    st.markdown("### ⚡ 商务控制台")
    st.caption("引擎状态: 🟢 运行中 (去重记忆同步)")
    st.markdown("---")
    
    st.subheader("💾 客户全局去重库")
    col_m1, col_m2 = st.columns(2)
    col_m1.metric("已锁手机", f"{len(global_seen_phones)}")
    col_m2.metric("已锁邮箱", f"{len(global_seen_emails)}")

    if st.button("🧹 清空去重记忆库", use_container_width=True, type="secondary"):
        if os.path.exists(DB_FILE):
            os.remove(DB_FILE)
        global_seen_phones.clear()
        global_seen_emails.clear()
        st.success("去重记忆已重置！")
        st.rerun()

    st.markdown("---")
    st.subheader("📂 历史档案管理")

    input_history_files = sorted(glob.glob(os.path.join(UPLOAD_DIR, "*.xlsx")), key=os.path.getmtime, reverse=True)
    with st.expander(f"📥 原始订单上传记录 ({len(input_history_files)})"):
        if input_history_files:
            for f in input_history_files[:8]:
                fname = os.path.basename(f)
                mtime = datetime.datetime.fromtimestamp(os.path.getmtime(f)).strftime('%m-%d %H:%M')
                st.text(f"• {fname} ({mtime})")
        else:
            st.caption("暂无历史上传")

    output_history_files = sorted(glob.glob(os.path.join(OUTPUT_DIR, "*.xlsx")), key=os.path.getmtime, reverse=True)
    with st.expander(f"📤 导出成品云端归档 ({len(output_history_files)})"):
        if output_history_files:
            for f in output_history_files[:8]:
                fname = os.path.basename(f)
                mtime = datetime.datetime.fromtimestamp(os.path.getmtime(f)).strftime('%m-%d %H:%M')
                with open(f, "rb") as file_data:
                    st.download_button(
                        label=f"⬇️ {fname[:18]}...",
                        data=file_data.read(),
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"hist_dl_{fname}",
                        use_container_width=True
                    )
        else:
            st.caption("暂无导出历史")

# ---------------- 主界面 ----------------
st.title("✂️ 假发跨境电商订单智能处理系统")
st.markdown("集成 **Apple ID / iMessage 绿色高亮标注**、**跨表全局去重**与**动态话术公式注入**的私域转化引擎。")
st.markdown("---")

# 模块 1：话术模板配置
with st.container():
    st.subheader("1. 💬 自动化话术配置模板")
    col_h, col_k = st.columns(2)
    with col_h:
        msg_h_text = st.text_area("H列：奖励邀请话术模板提示", value=getattr(config, 'DEFAULT_MSG_H', "Thanks for purchasing..."), height=90)
    with col_k:
        msg_k_text = st.text_area("K列：Ambassador 体验话术模板提示", value=getattr(config, 'DEFAULT_MSG_K', "Thank you for choosing..."), height=90)

st.markdown("---")

# 模块 2：订单文件上传
with st.container():
    st.subheader("2. 📂 订单数据接入")
    uploaded_file = st.file_uploader("支持拖拽原始 Amazon 订单表格 (.xlsx)", type=["xlsx"])

if uploaded_file:
    saved_input_path = os.path.join(UPLOAD_DIR, uploaded_file.name)
    with open(saved_input_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    with st.spinner("🚀 商业引擎正在全速解析，执行去重并生成格式化表格..."):
        process_error = None
        try:
            process_data.process_single_file(saved_input_path, global_seen_phones, global_seen_emails)
            process_data.save_local_database(DB_FILE, global_seen_phones, global_seen_emails)
            
            last_output_filename = f"已处理+{uploaded_file.name}"
            out_file_path = os.path.join(OUTPUT_DIR, last_output_filename)
            
            if os.path.exists(out_file_path):
                wb_processed = openpyxl.load_workbook(out_file_path, data_only=False)
                ws = wb_processed.active
                raw_data = list(ws.values)
                if len(raw_data) > 0:
                    headers = [str(h) if h is not None else "" for h in raw_data[0]]
                    df_current = pd.DataFrame(raw_data[1:], columns=headers)
                    is_real_data = True
                    st.success("✨ 订单清洗成功！已完成跨表全局去重、Apple ID/iMessage 亮绿标识。")
        except Exception as ex:
            process_error = traceback.format_exc()

        if process_error:
            st.error("❌ 引擎运行遇到异常：")
            st.code(process_error)
else:
    st.info("💡 提示：当前未上传文件，下方展示系统**标准内测范例数据**。上传真实文件后将自动加载处理结果。")
    df_current = get_default_sample_df()

st.markdown("---")

# 模块 3：数据在线交互与筛选
with st.container():
    st.subheader("3. 🔍 实时数据交互与多维筛选")
    c1, c2 = st.columns(2)
    with c1:
        search_kw = st.text_input("全局检索（支持姓名、电话、邮箱、订单号）：", placeholder="输入关键词快速搜索...")
    with c2:
        sku_col = [c for c in df_current.columns if "尺寸" in c or "SKU" in c]
        if sku_col:
            all_skus = [str(x) for x in df_current[sku_col[0]].dropna().unique() if str(x) != "None" and str(x) != ""]
            selected_skus = st.multiselect("按尺寸/SKU 过滤：", options=all_skus, default=all_skus)
        else:
            selected_skus = []

df_display = df_current.copy()
if search_kw:
    mask = df_display.astype(str).apply(lambda row: row.str.contains(search_kw, case=False).any(), axis=1)
    df_display = df_display[mask]
if sku_col and selected_skus:
    df_display = df_display[df_display[sku_col[0]].astype(str).isin(selected_skus)]

status_tag = "已加载真实业务数据" if is_real_data else "系统预设范例"
st.markdown(f"**数据预览清单** （当前筛选出 **{len(df_display)}** 条记录 / 总计 {len(df_current)} 条，当前状态：{status_tag}）：")
st.dataframe(df_display, use_container_width=True, height=280)

st.markdown("---")

# 模块 4：自定义导出列与下载中心
with st.container():
    st.subheader("4. 🎯 定制化导出与下载中心")
    all_available_cols = [c for c in df_current.columns if c != ""]

    selected_cols = st.multiselect(
        "勾选你需要包含在最终 Excel 中的字段列：",
        options=all_available_cols,
        default=all_available_cols
    )

    if not selected_cols:
        st.warning("⚠️ 请至少选择一列导出！")
    else:
        with st.expander("👁️ 查看最终导出结果前 5 行预览"):
            st.dataframe(df_display[selected_cols].head(5), use_container_width=True)

        if is_real_data and wb_processed:
            ws_tar = wb_processed.active
            header_row = [cell.value for cell in ws_tar[1]]
            
            cols_to_delete = [idx for idx, h in enumerate(header_row, start=1) if h not in selected_cols and h is not None and h != ""]
            for col_idx in sorted(cols_to_delete, reverse=True):
                ws_tar.delete_cols(col_idx)

            output_buffer = io.BytesIO()
            wb_processed.save(output_buffer)

            col_down1, col_down2 = st.columns(2)
            with col_down1:
                st.download_button(
                    label="📥 立即下载主订单商业报表 (.xlsx)",
                    data=output_buffer.getvalue(),
                    file_name=last_output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            with col_down2:
                imessage_df = df_current[df_current.astype(str).apply(lambda row: row.str.contains("iMessage|Apple", case=False).any(), axis=1)]
                imessage_buffer = io.BytesIO()
                imessage_df.to_excel(imessage_buffer, index=False)
                st.download_button(
                    label="🍏 一键导出 iMessage / 苹果生态精准触达表",
                    data=imessage_buffer.getvalue(),
                    file_name=f"iMessage专用触达+{uploaded_file.name if uploaded_file else 'sample.xlsx'}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.button("📥 立即下载定制处理后的 Excel 商业报表（请先在上方上传真实订单）", disabled=True, use_container_width=True)
