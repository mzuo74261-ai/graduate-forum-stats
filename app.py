import streamlit as st
import pandas as pd
import io

# 设置网页配置
st.set_page_config(page_title="研究生论坛名单统计", layout="centered")

st.title("📊 集成电路研究生论坛名单统计")

# ==========================================
# 0. UI 输入区域
# ==========================================
period = st.text_input("请输入这是第几期？(用于生成文件名)", value="1")

st.info("👇 请在下方依次上传三个文件")
col1, col2, col3 = st.columns(3)

with col1:
    file_reg_upload = st.file_uploader("1. 上传报名表", type=['xls', 'xlsx'], key="reg")
with col2:
    file_in_upload = st.file_uploader("2. 上传签到表", type=['xls', 'xlsx'], key="in")
with col3:
    file_out_upload = st.file_uploader("3. 上传签退表", type=['xls', 'xlsx'], key="out")


# ==========================================
# 1. 数据清洗函数
# ==========================================
def clean_data(df, tag="表"):
    df.columns = df.columns.str.strip()
    try:
        name_col = [c for c in df.columns if "姓名" in c][0]
        id_col = [c for c in df.columns if "学号" in c or "学工号" in c][0]
    except IndexError:
        st.error(f"❌ 在【{tag}】中没找到'姓名'或'学号'列，请检查文件内容。")
        st.stop()
        # 现在的写法 (学号在前，姓名在后)
    df_new = df[[id_col, name_col]].copy()
    df_new.columns = ['学号', '姓名']

    # 强制转换为字符串并清洗
    df_new['学号'] = df_new['学号'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df_new['姓名'] = df_new['姓名'].astype(str).str.strip()
    return df_new


# ==========================================
# 2. 核心处理逻辑
# ==========================================
if file_reg_upload and file_in_upload and file_out_upload:

    # 添加一个分割线，让按钮更明显
    st.divider()

    if st.button("🚀 开始统计并生成名单", type="primary", use_container_width=True):
        try:
            with st.spinner('正在分析数据...'):
                # 读取
                df_reg = pd.read_excel(file_reg_upload)
                df_in = pd.read_excel(file_in_upload)
                df_out = pd.read_excel(file_out_upload)

                # 清洗
                df_reg_clean = clean_data(df_reg, "报名表")
                df_in_clean = clean_data(df_in, "签到表")
                df_out_clean = clean_data(df_out, "签退表")

                # 逻辑比对
                set_reg = set(df_reg_clean['姓名'])
                set_in = set(df_in_clean['姓名'])
                set_out = set(df_out_clean['姓名'])

                success_names = set_reg & set_in & set_out
                anomaly_names = set_out - set_reg

                # 结果表
                result_success = df_reg_clean[df_reg_clean['姓名'].isin(success_names)].drop_duplicates()
                result_anomaly = df_out_clean[df_out_clean['姓名'].isin(anomaly_names)].drop_duplicates()

            # ---------------------------------------------------------
            # 3. 结果展示区
            # ---------------------------------------------------------
            st.success("✅ 统计完成！")

            # 指标卡片
            m1, m2 = st.columns(2)
            m1.metric("最终成功参会人数", f"{len(result_success)} 人")
            m2.metric("异常人数 (未报名却签退)", f"{len(result_anomaly)} 人", delta_color="inverse")

            # >>>>> 关键修改：显示异常名单表格 <<<<<
            st.write("---")  # 分割线
            if not result_anomaly.empty:
                st.error(f"⚠️ 发现 {len(result_anomaly)} 名未报名却签退的人员：")
                # 使用 st.table 展示，比 dataframe 更直观，且一定会展开显示
                st.table(result_anomaly)
            else:
                st.info("👍 完美！没有发现异常人员。")
            st.write("---")  # 分割线

            # ---------------------------------------------------------
            # 4. 下载按钮
            # ---------------------------------------------------------
            output_buffer = io.BytesIO()
            with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
                result_success.to_excel(writer, sheet_name='参加名单(成功)', index=False)
                result_anomaly.to_excel(writer, sheet_name='异常名单(未报名)', index=False)
            output_buffer.seek(0)

            st.download_button(
                label="📥 下载 Excel 结果文件",
                data=output_buffer,
                file_name=f"第{period}期集成电路研究生论坛参加名单.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        except Exception as e:

            st.error(f"发生错误: {e}")
