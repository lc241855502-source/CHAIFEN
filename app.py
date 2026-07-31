import streamlit as st
import pandas as pd
import io
import zipfile
import re

# ========== 工具函数 ==========
def col_letter_to_index(letter):
    """Excel列字母转0-based索引"""
    letter = letter.upper()
    result = 0
    for char in letter:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1

def sanitize_filename(name):
    """清理文件名非法字符"""
    name = str(name)
    name = re.sub(r'[\\/:*?"<>|]', '_', name)
    name = name.strip()
    if not name:
        name = "未命名"
    return name

def split_excel(df):
    """核心拆分逻辑，返回 {文件名: DataFrame} 字典"""
    all_columns = df.columns.tolist()
    total_cols = len(all_columns)

    # 关键列校验
    bk_idx = col_letter_to_index('BK')
    aw_idx = col_letter_to_index('AW')
    a_idx = col_letter_to_index('A')

    for col_letter, idx in [('BK', bk_idx), ('AW', aw_idx), ('A', a_idx)]:
        if idx >= total_cols:
            raise ValueError(f"原表缺少关键列：{col_letter}列（原表仅{total_cols}列）")

    bk_col = all_columns[bk_idx]
    aw_col = all_columns[aw_idx]
    a_col = all_columns[a_idx]

    # 严格按指定顺序定义前后段列
    front_cols_letters = ['AA', 'AB', 'A', 'C', 'L', 'K', 'E', 'AC']
    back_cols_letters = [
        'R', 'S', 'T', 'U', 'V', 'W', 'X',
        'AI', 'AG', 'AH', 'AN', 'AP', 'AO', 'AT',
        'CL', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP'
    ]

    def letters_to_col_names(letter_list):
        col_names = []
        for letter in letter_list:
            idx = col_letter_to_index(letter)
            if idx >= total_cols:
                raise ValueError(f"原表缺少指定列：{letter}列（原表仅{total_cols}列）")
            col_names.append(all_columns[idx])
        return col_names

    front_col_names = letters_to_col_names(front_cols_letters)
    back_col_names = letters_to_col_names(back_cols_letters)

    # 季度对应列映射：季度值 -> 原表列名
    quarter_map = {
        'Q1': all_columns[col_letter_to_index('N')] if col_letter_to_index('N') < total_cols else None,
        'Q2': all_columns[col_letter_to_index('O')] if col_letter_to_index('O') < total_cols else None,
        'Q3': all_columns[col_letter_to_index('P')] if col_letter_to_index('P') < total_cols else None,
        'Q4': all_columns[col_letter_to_index('Q')] if col_letter_to_index('Q') < total_cols else None,
    }

    # 获取BK列非空唯一值
    bk_values = df[bk_col].dropna().unique()
    bk_values = [v for v in bk_values if str(v).strip() != '']

    if not bk_values:
        raise ValueError("BK列中未找到有效数据")

    result_dict = {}

    for bk_val in bk_values:
        # 双重筛选：BK列匹配 或 AW列匹配
        mask_bk = df[bk_col].astype(str).str.strip() == str(bk_val).strip()
        mask_aw = df[aw_col].astype(str).str.strip() == str(bk_val).strip()
        subset = df[mask_bk | mask_aw].copy()

        if subset.empty:
            continue

        # 判断当前子集包含几个季度，决定动态列表头
        q_unique = subset[a_col].astype(str).str.strip().str.upper().unique()
        valid_quarters = [q for q in q_unique if q in quarter_map and quarter_map[q] is not None]

        if len(valid_quarters) == 1:
            # 只有一个季度：直接使用原表对应列的表头，和原表完全一致
            dynamic_col_name = quarter_map[valid_quarters[0]]
        else:
            # 包含多个季度：使用通用名称
            dynamic_col_name = "季度对应金额"

        # 逐行计算动态季度列的值，金额类保留2位小数
        dynamic_values = []
        for _, row in subset.iterrows():
            q_val = str(row[a_col]).strip().upper()
            target_col = quarter_map.get(q_val)
            if target_col and target_col in all_columns:
                val = row[target_col]
                # 尝试转为数值并四舍五入保留2位小数，非数值内容保留原值
                try:
                    num_val = float(val)
                    dynamic_values.append(round(num_val, 2))
                except (ValueError, TypeError):
                    dynamic_values.append(val)
            else:
                dynamic_values.append(None)

        # 按顺序拼接列：前半静态列 + 动态列 + 后半静态列
        result_df = subset[front_col_names].copy()
        result_df[dynamic_col_name] = dynamic_values
        result_df = pd.concat([result_df, subset[back_col_names].copy()], axis=1)

        safe_name = sanitize_filename(bk_val)
        result_dict[f"{safe_name}.xlsx"] = result_df

    return result_dict

# ========== Streamlit 页面 ==========
st.set_page_config(page_title="设备奖金表拆分工具", layout="centered")
st.title("📊 设备奖金表拆分工具")
st.caption("按BK列拆分文件 · 自动纳入AW列匹配行 · 表头与原表保持一致")

uploaded_file = st.file_uploader("上传【设备奖金原表】Excel文件", type=["xlsx", "xls"])

if uploaded_file is not None:
    if st.button("开始拆分处理", type="primary", use_container_width=True):
        try:
            # 读取Excel指定工作表，第二行为表头，自动识别数据类型
            sheet_name = "销售Q1-Q2奖金汇总表"
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=1)
            
            # 执行拆分
            result_dict = split_excel(df)
            
            if not result_dict:
                st.warning("未生成任何拆分文件，请检查BK列数据")
            else:
                st.success(f"✅ 处理完成！共拆分出 {len(result_dict)} 个文件")
                st.caption("数字列保留数值格式，奖金金额自动保留2位小数")
                
                # 生成ZIP压缩包供一键下载
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for filename, data_df in result_dict.items():
                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                            data_df.to_excel(writer, index=False, sheet_name="数据")
                        zip_file.writestr(filename, excel_buffer.getvalue())
                
                zip_buffer.seek(0)
                
                st.download_button(
                    label="📦 一键下载全部文件(ZIP)",
                    data=zip_buffer,
                    file_name="拆分结果.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                
                # 展开单个文件下载列表
                with st.expander("查看并单独下载每个文件"):
                    for filename, data_df in result_dict.items():
                        col1, col2 = st.columns([3, 1])
                        col1.text(filename)
                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                            data_df.to_excel(writer, index=False, sheet_name="数据")
                        excel_buffer.seek(0)
                        col2.download_button("下载", excel_buffer, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=filename)
        
        except Exception as e:
            st.error(f"❌ 处理失败：{str(e)}")

st.divider()
st.markdown("""
**处理规则说明**
1. 以 BK 列为基准拆分，每个唯一值生成一个独立 Excel 文件
2. 同时将 AW 列中匹配该值的行也并入对应文件
3. 数字列保留原生数值格式，可直接计算；奖金金额列自动保留2位小数
4. 列顺序：AA/AB/A/C/L/K/E/AC + 动态季度列 + R/S/T/U/V/W/X/AI/AG/AH/AN/AP/AO/AT/CL/BK/BL/BM/BN/BO/BP
""")
