# web_app.py
import streamlit as st
import os
import io
import pandas as pd
import openpyxl
import config
import process_data  # 引入核心处理逻辑

st.set_page_config(
    page_title="假发订单自动化处理系统",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 历史去重数据库文件名
DB_FILE = "seen_database.txt"

# 1. 独立数据库加载与保存逻辑（不依赖 process_data 内部是否有该函数）
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
        },
        {
            "订单号": "113-7654321-8904321",
            "尺寸": "Straight 24 inch",
            "订单日期": "2026-09-16",
            "姓名": "Ashley Brown",
            "电话": "+1 (202) 555-0188",
            "回复情况": "",
            "邮箱": "ashley.test@yahoo.com",
            "Hello 姓名,": '="Hello "&D3&","&CHAR(10)&CHAR(10)&"Thanks for purchasing..."',
            "Ambassador 话术": '="Hello "&D3&","&CHAR(10)&CHAR(10)&"Thank you for choosing..."'
        }
    ])

# 默认初始化预览数据
df_current = get_default_sample_df()

# ---------------- 侧边栏：数据库控制 ----------------
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

# ---------------- 主界面（四大模块固定排列） ----------------
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
    # 动态把界面输入更新进 config
    config.DEFAULT_MSG_H = msg_h_text
    config.DEFAULT_MSG_K = msg_k_text

    temp_input_path = f"temp_in_{uploaded_file.name}"
    with open(temp_input_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    with st.spinner("正在调用引擎进行数据处理（保留单元格合并/公式/高亮/去重）..."):
        # 搜索 process_data 中的主处理函数
        target_fn = None
        for fn_name in dir(process_data):
            if callable(getattr(process_data, fn_name)) and not fn_name.startswith("__"):
                if fn_name in ["process_single_file", "process_excel", "process_file", "main", "run"]:
                    target_fn = getattr(process_data, fn_name)
                    break

        # 如果没有固定名字，自动选取带 process 的函数
        if not target_fn:
            for fn_name in dir(process_data):
                if "process" in fn_name.lower() and callable(getattr(process_data, fn_name)):
                    target_fn = getattr(process_data, fn_name)
                    break

        if target_fn:
            try:
                # 优先按 3 个参数（含去重集合）调用
                target_fn(temp_input_path, global_seen_phones, global_seen_emails)
            except TypeError:
                try:
                    # 尝试单参数调用
                    target_fn(temp_input_path)
                except Exception as ex:
                    st.error(f"调用处理函数时发生错误: {ex}")
        else:
            st.error("⚠️ 未能匹配到 process_data.py 中的主入口函数！")

        # 读取处理完成的文件
        out_path = os.path.join("处理完成", f"已处理+{uploaded_file.name}")
        if not os.path.exists(out_path) and os.path.exists("处理完成"):
            files = [f for f in os.listdir("处理完成") if f.endswith(".xlsx")]
            if files:
                out_path = os.path.join("处理完成", files[0])

        if os.path.exists(out_path):
            wb_processed = openpyxl.load_workbook(out_path)
            ws = wb_processed.active
            raw_data = list(ws.values)
            if len(raw_data) > 0:
                headers = [str(h) if h is not None else "" for h in raw_data[0]]
                df_current = pd.DataFrame(raw_data[1:], columns=headers)
                is_real_data = True
                st.success("✅ 引擎处理完成！合并单元格与颜色高亮已成功保留。")
            
            if os.path.exists(temp_input_path):
                os.remove(temp_input_path)
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
        
        # 倒序删除未选中的列
        cols_to_delete = [idx for idx, h in enumerate(header_row, start=1) if h not in selected_cols and h is not None and h != ""]
        for col_idx in sorted(cols_to_delete, reverse=True):
            ws_tar.delete_cols(col_idx)

        # 保存更新去重库
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
