# web_app.py
import streamlit as st
import os
import io
import pandas as pd
import openpyxl
import config
import process_data  # 引入拆分后的核心数据处理引擎

st.set_page_config(
    page_title="假发订单自动化处理系统",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 动态获取后端配置的数据库文件名
DB_FILE = getattr(process_data, 'DB_FILE_NAME', "seen_database.txt")

# 1. 加载历史去重数据库
global_seen_phones, global_seen_emails = process_data.load_local_database(DB_FILE)

# 2. 生成默认范例数据（用于未上传 Excel 时在模块 3 和模块 4 进行排版展示与列勾选预览）
def get_default_sample_df():
    sample_data = [
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
    ]
    return pd.DataFrame(sample_data)

# ---------------- 侧边栏：数据库状态与控制台 ----------------
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

# ---------------- 主界面（四大模块绝对固定显示） ----------------
st.title("✂️ 假发订单数据自动化处理与高级定制导出系统")

# 模块 1：话术模板配置（固定显示）
st.header("1. 💬 话术模板配置")
col_h, col_k = st.columns(2)
with col_h:
    msg_h_text = st.text_area(
        "H列 奖励邀请话术 (Hello 姓名,)", 
        value=getattr(config, 'DEFAULT_MSG_H', "Thanks for purchasing..."), 
        height=100
    )
with col_k:
    msg_k_text = st.text_area(
        "K列 Ambassador 体验话术", 
        value=getattr(config, 'DEFAULT_MSG_K', "Thank you for choosing..."), 
        height=100
    )

st.markdown("---")

# 模块 2：文件上传与引擎处理（固定显示）
st.header("2. 📂 上传订单 Excel 文件")
uploaded_file = st.file_uploader("请拖入或选择需要处理的原始订单文件 (.xlsx)", type=["xlsx"])

is_real_data = False
wb_processed = None

if uploaded_file:
    # 动态将用户选中的话术更新到 config 中，保证 process_data 调用时能获取最新话术
    config.DEFAULT_MSG_H = msg_h_text
    config.DEFAULT_MSG_K = msg_k_text

    # 创建临时文件供后端引擎读取
    temp_input_path = f"temp_in_{uploaded_file.name}"
    with open(temp_input_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    with st.spinner("正在调用后端默认引擎执行多行绑定、合并单元格、高亮标注及跨表去重..."):
        # 100% 调用 process_data.py 的核心函数
        process_data.process_single_file(temp_input_path, global_seen_phones, global_seen_emails)
        
        # 读取 process_data 生成的输出文件
        out_filename = f"已处理+{uploaded_file.name}"
        out_path = os.path.join("处理完成", out_filename)
        
        if os.path.exists(out_path):
            wb_processed = openpyxl.load_workbook(out_path)
            ws = wb_processed.active
            raw_data = list(ws.values)
            if len(raw_data) > 0:
                headers = [str(h) if h is not None else "" for h in raw_data[0]]
                df_current = pd.DataFrame(raw_data[1:], columns=headers)
                is_real_data = True
                st.success("✅ 核心引擎处理完成！已应用合并单元格与颜色标注。")
            
            # 及时清除临时输入文件
            if os.path.exists(temp_input_path):
                os.remove(temp_input_path)
else:
    st.info("💡 当前未上传文件，下方展示系统**标准默认输出范例**。您可以先配置模板并进行导出列预览。")
    df_current = get_default_sample_df()

st.markdown("---")

# 模块 3：在线搜索与查看（固定显示）
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

# 模块 4：自定义导出列与范例预览（固定显示）
st.header("4. 🎯 自定义导出列配置与范例预览")

all_available_cols = [c for c in df_current.columns if c != ""]

selected_cols = st.multiselect(
    "请勾选你本次需要导出的 Excel 列（未勾选的列将在导出时自动剔除）：",
    options=all_available_cols,
    default=all_available_cols
)

if not selected_cols:
    st.warning("⚠️ 请至少选择一列导出！")
else:
    st.subheader("👁️ 导出表格范例预览")
    st.dataframe(df_display[selected_cols].head(5), use_container_width=True)

    if is_real_data and wb_processed:
        # 在保留 OpenPyXL 格式（合并单元格/公式/高亮）前提下，删除未勾选的列
        ws_tar = wb_processed.active
        header_row = [cell.value for cell in ws_tar[1]]
        
        cols_to_delete = []
        for idx, h in enumerate(header_row, start=1):
            if h not in selected_cols and h is not None and h != "":
                cols_to_delete.append(idx)
        
        # 从右往左删除列，防止索引偏移
        for col_idx in sorted(cols_to_delete, reverse=True):
            ws_tar.delete_cols(col_idx)

        # 持久化保存最新的去重数据库文本
        process_data.save_local_database(DB_FILE, global_seen_phones, global_seen_emails)

        # 写入内存流，供浏览器下载
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
