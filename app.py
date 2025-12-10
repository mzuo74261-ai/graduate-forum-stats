import streamlit as st
import pandas as pd
import io

# 设置网页配置
st.set_page_config(page_title="研究生论坛名单统计", layout="centered")

st.title("📊 集成电路研究生论坛名单统计")

# ==========================================
# 0. UI 输入区域
# ==========================================
# 默认值改为空，方便你输入 "十一" 或 "11"
period = st.text_input("请输入这是第几期？(用于生成文件名)", value="一")

st.info("👇 请在下方依次上传三个文件")
col1, col2, col3 = st.columns(3)

with col1:
    file_reg_upload = st.file_uploader("1. 上传报名表", type=['xls', 'xlsx'], key="reg")
with col2:
    file_in_upload = st.file_uploader("2. 上传签到表", type=['xls', 'xlsx'], key="in")
with col3:
    file_out_upload = st.file_uploader("3. 上传签退表", type=['xls', 'xlsx'], key="out")

# ==========================================
# A. 新增：智能读取函数 (自动找表头)
# ==========================================
# ==========================================
# A. 新增：智能读取函数 (带强力纠错模式)
# ==========================================
def smart_read_excel(file):
    """
    1. 自动跳过大标题，寻找含'姓名'的表头。
    2. 如果遇到 'Workbook corruption' 错误，启用 xlrd 强力模式读取。
    """
    import xlrd # 确保引入 xlrd
    
    # 辅助函数：定位表头并清理数据
    def find_header_and_clean(df_raw):
        target_row_index = -1
        # 在前 10 行里找，哪一行含有 "姓名" 两个字
        for i, row in df_raw.head(10).iterrows():
            row_str = " ".join([str(x) for x in row.values])
            if "姓名" in row_str:
                target_row_index = i
                break
        
        if target_row_index != -1:
            df_raw.columns = df_raw.iloc[target_row_index] # 设置新表头
            df_raw = df_raw.iloc[target_row_index + 1:].reset_index(drop=True) # 截取数据
        return df_raw

    try:
        # --- 尝试 1: 标准读取 ---
        file.seek(0) # 确保指针在开头
        df = pd.read_excel(file, header=None)
        
    except Exception as e:
        # 如果报错包含 "corruption"，说明是老式 xls 文件损坏
        if "corruption" in str(e) or "xlrd" in str(e):
            try:
                # --- 尝试 2: 强力模式 (忽略损坏) ---
                file.seek(0)
                file_content = file.read()
                # 使用 ignore_workbook_corruption=True 强行读取
                wb = xlrd.open_workbook(file_contents=file_content, ignore_workbook_corruption=True)
                sheet = wb.sheet_by_index(0)
                
                # 手动将数据转为 DataFrame
                data = []
                for row_idx in range(sheet.nrows):
                    data.append(sheet.row_values(row_idx))
                df = pd.DataFrame(data)
            except Exception as e2:
                st.error(f"❌ 文件严重损坏无法读取，请尝试用 Excel 打开并另存为 .xlsx 格式再上传。\n错误详情: {e2}")
                st.stop()
        else:
            st.error(f"❌ 读取文件出错: {e}")
            st.stop()

    # 统一进行表头查找清洗
    return find_header_and_clean(df)

# ==========================================
# 1. 数据清洗函数
# ==========================================
def clean_data(df, tag="表"):
    df.columns = df.columns.astype(str).str.strip()
    try:
        # 找姓名列
        name_col = [c for c in df.columns if "姓名" in c][0]
        
        # 找学号列 (支持 "学号" 或 "学工号")
        id_col = [c for c in df.columns if "学号" in c or "学工号" in c][0]
        
    except IndexError:
        st.error(f"❌ 在【{tag}】中没找到 '姓名' 列，或者没找到 '学号'/'学工号' 列。\n请检查文件是否包含这些列名，或者是否有大标题挡住了。")
        st.stop() 

    # 提取数据 (学号在前，姓名在后)
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
    
    st.divider()
    
    if st.button("🚀 开始统计并生成名单", type="primary", use_container_width=True):
        try:
            with st.spinner('正在智能分析数据结构...'):
                # >>> 修改点：使用 smart_read_excel 替代 pd.read_excel <<<
                df_reg = smart_read_excel(file_reg_upload)
                df_in = smart_read_excel(file_in_upload)
                df_out = smart_read_excel(file_out_upload)

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

            m1, m2 = st.columns(2)
            m1.metric("最终成功参会人数", f"{len(result_success)} 人")
            m2.metric("异常人数 (未报名却签退)", f"{len(result_anomaly)} 人", delta_color="inverse")

            st.write("---") 
            if not result_anomaly.empty:
                st.error(f"⚠️ 发现 {len(result_anomaly)} 名未报名却签退的人员：")
                st.table(result_anomaly)
            else:
                st.info("👍 完美！没有发现异常人员。")
            st.write("---") 

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

