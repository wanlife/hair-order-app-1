# web_app.py
import streamlit as st
import os
import io
import pandas as pd
import plotly.express as px
import config
import process_data

st.set_page_config(
    page_title="假发订单智能数据引擎",
    layout="wide",
    initial_sidebar_state="expanded"
)

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

# 侧边栏：数据库与全局开关
st.sidebar.title("⚙️ 智能控制台")
st.sidebar.subheader("💾 历史数据库状态")
st.sidebar.metric("已去重手机号", f"{len(global_seen_phones)} 个")
st.sidebar.metric("已去重邮箱", f"{len(global_seen_emails)} 个")

if st.sidebar.button("🧹 清空历史去重记忆", type="secondary"):
    if os.path.exists(DB_FILE):
        os.remove(DB_FILE)
    global_seen_phones.clear()
    global_seen_emails.clear()
    st.sidebar.success("记忆库已成功重置！")

st.title("✂️ 假发订单智能数据处理与分析系统")

# 标签页导航，划分清晰的功能板块
tab1, tab2, tab3 = st.tabs(["📂 数据处理与自定义导出", "📊 数据分析看板", "💬 话术模板管理"])

with tab3:
    st.header("💬 话术模板在线配置")
    col_h, col_k = st.columns(2)
    with col_h:
        msg_h_text = st.text_area("H列 奖励邀请话术", value=config.DEFAULT_MSG_H, height=120)
    with col_k:
        msg_k_text = st.text_area("K列 Ambassador 体验话术", value=config.DEFAULT_MSG_K, height=120)

with tab1:
    st.header("1. 上传订单文件")
    uploaded_file = st.file_uploader("上传原始 Excel 文件 (.xlsx)", type=["xlsx"])

    if uploaded_file:
        # 重置指针
        uploaded_file.seek(0)
        df_raw = pd.read_excel(uploaded_file)

        # 核心清洗计算
        uploaded_file.seek(0)
        wb_processed = process_data.run_excel_processing(
            uploaded_file, msg_h_text, msg_k_text, global_seen_phones, global_seen_emails
        )

        st.success("✅ 数据解析与智能去重完成！")

        # 转换为 DataFrame 用于网页端的高级交互预览
        ws = wb_processed.active
        data = list(ws.values)
        headers = data[0]
        df_processed = pd.DataFrame(data[1:], columns=headers)

        st.markdown("---")
        st.header("2. 🔍 智能高级筛选与交互查看")

        # 动态多条件筛选器
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            all_skus = [str(x) for x in df_processed["尺寸"].dropna().unique() if str(x) != "None"]
            selected_skus = st.multiselect("按 SKU/尺寸 筛选：", options=all_skus, default=all_skus)
        with col_f2:
            search_keyword = st.text_input("搜索买家姓名/订单号/电话：", "")

        # 过滤逻辑
        df_filtered = df_processed[df_processed["尺寸"].astype(str).isin(selected_skus)]
        if search_keyword:
            df_filtered = df_filtered[
                df_filtered["订单号"].astype(str).str.contains(search_keyword, case=False, na=False) |
                df_filtered["姓名"].astype(str).str.contains(search_keyword, case=False, na=False) |
                df_filtered["电话"].astype(str).str.contains(search_keyword, case=False, na=False)
            ]

        # 交互式数据表格（支持排序、全屏放大、单元格内容复制）
        st.dataframe(df_filtered, use_container_width=True, height=400)
        st.caption(f"当前筛选条件下共显示 **{len(df_filtered)}** 条记录")

        st.markdown("---")
        st.header("3. 🎯 自由列选择与结果导出")
        
        all_cols = list(df_processed.columns)
        selected_cols = st.multiselect("勾选需要保留导出的列：", options=all_cols, default=all_cols)

        if st.button("🚀 导出所选筛选结果为 Excel", type="primary"):
            if not selected_cols:
                st.error("请至少选择一列！")
            else:
                # 重新裁切 Excel
                ws_tar = wb_processed.active
                cols_to_delete = [idx for idx, h in enumerate(all_cols, start=1) if h not in selected_cols]
                for col_idx in sorted(cols_to_delete, reverse=True):
                    ws_tar.delete_cols(col_idx)

                save_local_database(DB_FILE, global_seen_phones, global_seen_emails)
                output = io.BytesIO()
                wb_processed.save(output)

                st.download_button(
                    label="📥 下载定制版的 Excel 结果",
                    data=output.getvalue(),
                    file_name=f"智能精简版_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # 标签页 2：图表与数据看板
        with tab2:
            st.header("📈 本批次订单数据分析看板")
            
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("总处理数据行", len(df_processed))
            m2.metric("有效客户数", df_processed["订单号"].nunique())
            
            valid_phones = df_processed["电话"].replace("", None).dropna().count()
            valid_emails = df_processed["邮箱"].replace("", None).dropna().count()
            
            m3.metric("成功提取电话数", valid_phones)
            m4.metric("成功提取邮箱数", valid_emails)

            st.markdown("---")
            c1, c2 = st.columns(2)

            with c1:
                st.subheader("📦 SKU 尺寸分布图")
                sku_counts = df_processed["尺寸"].value_counts().reset_index()
                sku_counts.columns = ["尺寸 SKU", "数量"]
                fig_sku = px.bar(sku_counts, x="尺寸 SKU", y="数量", color="数量", text_auto=True)
                st.plotly_chart(fig_sku, use_container_width=True)

            with c2:
                st.subheader("📧 联系方式获取占比")
                contact_df = pd.DataFrame({
                    "类别": ["含电话订单", "含邮箱订单"],
                    "数量": [valid_phones, valid_emails]
                })
                fig_pie = px.pie(contact_df, names="类别", values="数量", hole=0.4)
                st.plotly_chart(fig_pie, use_container_width=True)
