# web_app.py
import streamlit as st
import os
import io
import glob
import time
import datetime
import traceback
import pandas as pd
import openpyxl
import config
import process_data  # 引入核心处理引擎

st.set_page_config(
    page_title="假发订单自动化处理系统",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 创建历史记录存储文件夹
UPLOAD_DIR = "历史输入文件"
OUTPUT_DIR = "处理完成"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

DB_FILE = "seen_database.txt"

# 1. 数据库加载与保存
def local_load_db(db_path):
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

def local_save_db(db_path, seen_phones, seen_emails):
    with open(db_path, "w", encoding="utf-8") as f:
        for p in sorted(seen_phones):
            f.write(f"PHONE:{p}\n")
        for e in sorted(seen_emails):
            f.write(f"EMAIL:{e}\n")

global_seen_phones, global_seen_emails = local_load_db(DB_FILE)

# 2. 默认预览范例数据
def get_default_sample_df():
    return pd.DataFrame([
        {
            "订单号": "113-1234567-8901234",
            "尺寸": "13X6 ST-Body Wave 20",
            "订单日期": "2026-09-15",
            "姓名": "Jessica Smith",
            "电话": "+1 (202) 555-0143",
            "回复情况": "Sep 2026",
            "邮箱": "jessica@gmail.com",
            "Hello 姓名,": '="Hello "&D2&","&CHAR(10)&CHAR(10)&"Thanks for purchasing..."',
            "Ambassador 话术": '="Hello "&D2&","&CHAR(10)&CHAR(10)&"Thank you for choosing..."'
        }
    ])

df_current = get_default_sample_df()

# ---------------- 侧边栏：控制台与历史文件归档 ----------------
st.sidebar.title("⚙️ 控制台")
st.sidebar.subheader("💾 历史数据去重库")
st.sidebar.metric("已记录手机号", f"{len(global_seen_phones)} 个")
st.sidebar.metric("已记录邮箱", f"{len(global_seen_emails)} 个")

if st.sidebar.button("🧹 清空历史去重记忆", type="secondary"):
    if os.path.exists(DB_FILE):
        os.remove(DB_FILE)
    global_seen_phones.clear()
    global_seen_emails.clear()
    st.sidebar.success("历史去重记忆已重置！")
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("📂 文件历史档案馆")

# 查看输入历史
input_history_files = sorted(glob.glob(os.path.join(UPLOAD_DIR, "*.xlsx")), key=os.path.getmtime, reverse=True)
with st.sidebar.expander(f"📥 原始上传历史 ({len(input_history_files)} 个)"):
    if input_history_files:
        for f in input_history_files[:10]:
            fname = os.path.basename(f)
            mtime = datetime.datetime.fromtimestamp(os.path.getmtime(f)).strftime('%m-%d %H:%M')
            st.caption(f"📄 {fname} ({mtime})")
    else:
        st.write("暂无历史上传")

# 查看输出历史并支持随时重新下载
output_history_files = sorted(glob.glob(os.path.join(OUTPUT_DIR, "*.xlsx")), key=os.path.getmtime, reverse=True)
with st.sidebar.expander(f"📤 已导出文件历史 ({len(output_history_files)} 个)"):
    if output_history_files:
        for f in output_history_files[:10]:
            fname = os.path.basename(f)
            mtime = datetime.datetime.fromtimestamp(os.path.getmtime(f)).strftime('%m-%d %H:%M')
            with open(f, "rb") as file_data:
                st.download_button(
                    label=f"⬇️ {fname}",
                    data=file_data.read(),
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"hist_dl_{fname}"
                )
    else:
        st.write("暂无导出历史")

# ---------------- 主界面 ----------------
st.title("✂️ 假发订单数据自动化处理与高级定制导出系统")

# 模块 1：话术模板配置
st.header("1. 💬 话术模板配置")
col_h, col_k = st.columns(2)
with col_h:
    msg_h_text = st.text_area("H列 奖励邀请话术 (Hello 姓名,)", value=getattr(config, 'DEFAULT_MSG_H', "Thanks for purchasing..."), height=100)
with col_k:
    msg_k_text = st.text_area("K列 Ambassador 体验话术", value=getattr(config, 'DEFAULT_MSG_K', "Thank you for choosing..."), height=100)

st.markdown("---")

# 模块 2：订单文件上传
st.header("2. 📂 上传订单 Excel 文件")
uploaded_file = st.file_uploader("请拖入或选择需要处理的原始订单文件 (.xlsx)", type=["xlsx"])

is_real_data = False
wb_processed = None

if uploaded_file:
    config.DEFAULT_MSG_H = msg_h_text
    config.DEFAULT_MSG_K = msg_k_text

    # 1. 保存到输入历史档案库
    saved_input_path = os.path.join(UPLOAD_DIR, uploaded_file.name)
    with open(saved_input_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    with st.spinner("正在生成话术公式、计算去重并处理格式..."):
        process_error = None
        out_file_path = None
        try:
            out_file_path = process_data.run_excel_processing(
                saved_input_path, 
                msg_h_text, 
                msg_k_text, 
                global_seen_phones, 
                global_seen_emails
            )
        except Exception as ex:
            process_error = traceback.format_exc()

        if process_error:
            st.error("❌ 引擎处理失败，详细错误：")
            st.code(process_error)
        elif out_file_path and os.path.exists(out_file_path):
            try:
                wb_processed = openpyxl.load_workbook(out_file_path, data_only=False)
                ws = wb_processed.active
                raw_data = list(ws.values)
                if len(raw_data) > 0:
                    headers = [str(h) if h is not None else "" for h in raw_data[0]]
                    df_current = pd.DataFrame(raw_data[1:], columns=headers)
                    is_real_data = True
                    st.success(f"✅ 处理完成！已自动加入 H/K 列话术公式，且已保存至【输出历史档案】！")
            except Exception as read_ex:
                st.error(f"❌ 读取结果文件失败：{read_ex}")

else:
    st.info("💡 当前未上传文件，下方展示系统**标准默认输出范例**。")
    df_current = get_default_sample_df()

st.markdown("---")

# 模块 3：数据在线交互查看与筛选
st.header("3. 🔍 数据的在线交互查看与筛选")
c1, c2 = st.columns(2)
with c1:
    search_kw = st.text_input("🔍 全局关键字搜索（姓名/电话/邮箱/订单号）：", "")
with c2:
    sku_col = [c for c in df_current.columns if "尺寸" in c or "SKU" in c]
    if sku_col:
        all_skus = [str(x) for x in df_current[sku_col[0]].dropna().unique() if str(x) != "None" and str(x) != ""]
        selected_skus = st.multiselect("按尺寸 SKU 筛选：", options=all_skus, default=all_skus)
    else:
        selected_skus = []

df_display = df_current.copy()
if search_kw:
    mask = df_display.astype(str).apply(lambda row: row.str.contains(search_kw, case=False).any(), axis=1)
    df_display = df_display[mask]
if sku_col and selected_skus:
    df_display = df_display[df_display[sku_col[0]].astype(str).isin(selected_skus)]

status_tag = "真实处理结果" if is_real_data else "标准范例数据"
st.write(f"📊 当前呈现 **{len(df_display)}** 行数据（【{status_tag}】共 {len(df_current)} 行）:")
st.dataframe(df_display, use_container_width=True, height=260)

st.markdown("---")

# 模块 4：自定义导出列与范例预览
st.header("4. 🎯 自定义导出列配置与范例预览")
all_available_cols = [c for c in df_current.columns if c != ""]

selected_cols = st.multiselect(
    "请勾选你本次需要导出的 Excel 列：",
    options=all_available_cols,
    default=all_available_cols
)

if not selected_cols:
    st.warning("⚠️ 请至少选择一列导出！")
else:
    st.subheader("👁️ 导出表格范例预览")
    st.dataframe(df_display[selected_cols].head(5), use_container_width=True)

    if is_real_data and wb_processed:
        ws_tar = wb_processed.active
        header_row = [cell.value for cell in ws_tar[1]]
        
        # 倒序删除未勾选的列
        cols_to_delete = [idx for idx, h in enumerate(header_row, start=1) if h not in selected_cols and h is not None and h != ""]
        for col_idx in sorted(cols_to_delete, reverse=True):
            ws_tar.delete_cols(col_idx)

        # 保存更新去重数据库
        local_save_db(DB_FILE, global_seen_phones, global_seen_emails)

        output_buffer = io.BytesIO()
        wb_processed.save(output_buffer)

        st.download_button(
            label="📥 下载处理后的 Excel 最终文件",
            data=output_buffer.getvalue(),
            file_name=f"已处理+{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    else:
        st.button("📥 下载处理后的 Excel 最终文件（请先上传真实文件）", disabled=True, type="primary")
