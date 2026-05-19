import streamlit as st
import pandas as pd
import io
import zipfile
from pathlib import Path


# ==================== 核心功能函数 ====================
def split_matrices(df):
    """
    利用全空行和全空列作为分隔符，将 DataFrame 切分成多个子矩阵。
    返回子矩阵列表，每个子矩阵为字典，包含 data, start_row, start_col, shape。
    """
    empty_rows = df.isnull().all(axis=1)
    empty_cols = df.isnull().all(axis=0)

    # 找出所有连续非空行的区间
    row_groups = []
    in_block = False
    start = 0
    for i, is_empty in enumerate(empty_rows):
        if not is_empty and not in_block:
            in_block = True
            start = i
        elif is_empty and in_block:
            in_block = False
            row_groups.append((start, i - 1))
    if in_block:
        row_groups.append((start, len(df) - 1))

    # 找出所有连续非空列的区间
    col_groups = []
    in_block = False
    start = 0
    for j, is_empty in enumerate(empty_cols):
        if not is_empty and not in_block:
            in_block = True
            start = j
        elif is_empty and in_block:
            in_block = False
            col_groups.append((start, j - 1))
    if in_block:
        col_groups.append((start, len(df.columns) - 1))

    # 根据行区间和列区间，提取子矩阵
    matrices = []
    for r0, r1 in row_groups:
        for c0, c1 in col_groups:
            sub_df = df.iloc[r0:r1 + 1, c0:c1 + 1]
            sub_df = sub_df.dropna(how='all').dropna(axis=1, how='all')
            if not sub_df.empty:
                matrices.append({
                    'data': sub_df,
                    'start_row': r0 + 1,
                    'start_col': c0 + 1,
                    'shape': sub_df.shape
                })
    return matrices


def convert_excel_to_rowvectors_bytes(input_bytes, filename):
    """
    将 Excel 文件的字节数据转换为行向量，返回 (output_bytes, 矩阵数量)
    """
    try:
        # 尝试读取，显式指定引擎
        try:
            df_raw = pd.read_excel(io.BytesIO(input_bytes), header=None, engine='openpyxl')
        except Exception:
            # 备选：xlrd 引擎（适用于 .xls）
            df_raw = pd.read_excel(io.BytesIO(input_bytes), header=None, engine='xlrd')

        matrices = split_matrices(df_raw)

        if not matrices:
            return None, 0

        all_rows = []
        for i, mat_info in enumerate(matrices):
            mat = mat_info['data']
            vec = mat.values.ravel(order='C')
            row = [f'矩阵{i + 1}'] + vec.tolist()
            all_rows.append(row)
            if i < len(matrices) - 1:
                all_rows.append([])
                all_rows.append([])

        output_df = pd.DataFrame(all_rows)
        output_bytes = io.BytesIO()
        output_df.to_excel(output_bytes, index=False, header=False, engine='openpyxl')
        output_bytes.seek(0)
        return output_bytes.getvalue(), len(matrices)
    except Exception as e:
        raise Exception(f"处理文件 {filename} 时出错：{e}")


# ==================== Streamlit UI ====================
st.set_page_config(page_title="Excel矩阵展平工具", layout="centered")
st.title("📊 Excel 矩阵展平工具")
st.markdown("""
### 使用说明
1. 上传一个或多个 Excel 文件（支持 `.xlsx` 或 `.xls`）
2. 程序将自动识别每个文件中的**多个矩阵**（由全空行/全空列分隔）
3. 对每个矩阵按**行优先**展平为一维行向量，并在首位添加“矩阵N”标识
4. 不同矩阵的输出行之间**空两行**
5. 点击下载按钮获取处理后的文件（单个文件或打包的 ZIP）
""")

uploaded_files = st.file_uploader(
    "📂 选择 Excel 文件",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    key="file_uploader"
)

if uploaded_files:
    st.divider()
    st.subheader("处理选项")
    download_mode = st.radio(
        "下载方式",
        ["逐个文件下载", "打包为 ZIP 下载"],
        horizontal=True
    )

    if st.button("🚀 开始处理", type="primary"):
        results = {}  # filename -> (bytes, matrix_count)
        all_success = True

        with st.spinner("处理中，请稍候..."):
            for uploaded_file in uploaded_files:
                filename = uploaded_file.name
                # 确保输出文件名以 .xlsx 结尾
                out_filename = Path(filename).stem + ".xlsx"

                file_bytes = uploaded_file.read()
                out_bytes, count = convert_excel_to_rowvectors_bytes(file_bytes, filename)

                if out_bytes is not None:
                    results[out_filename] = (out_bytes, count)
                    st.success(f"✅ {filename} → 发现 {count} 个矩阵")
                else:
                    all_success = False
                    st.warning(f"⚠️ {filename} 未检测到有效矩阵，已跳过")

        if results:
            st.divider()
            st.subheader("📥 下载结果")

            if download_mode == "逐个文件下载":
                for out_filename, (out_bytes, count) in results.items():
                    st.download_button(
                        label=f"下载 {out_filename} (包含 {count} 个矩阵)",
                        data=out_bytes,
                        file_name=out_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=out_filename
                    )
            else:  # 打包为 ZIP
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                    for out_filename, (out_bytes, _) in results.items():
                        zipf.writestr(out_filename, out_bytes)
                zip_buffer.seek(0)

                st.download_button(
                    label="📦 下载全部文件 (ZIP 压缩包)",
                    data=zip_buffer,
                    file_name="row_vectors_results.zip",
                    mime="application/zip",
                    use_container_width=True
                )
        else:
            st.error("没有成功处理任何文件，请检查上传的文件格式和内容。")

# 页脚说明
st.divider()
st.caption("提示：本工具仅在浏览器本地处理数据，不会上传到任何服务器。")