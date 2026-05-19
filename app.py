import streamlit as st
import pandas as pd
import io
import zipfile
from pathlib import Path

# ========== 检查依赖 ==========
try:
    import openpyxl
except ImportError:
    st.error("❌ 服务器缺少 openpyxl 库，请检查 requirements.txt")
    st.stop()


# ========== 核心函数 ==========
def split_matrices(df):
    """利用全空行和全空列切分多个矩阵"""
    if df.empty:
        return []

    empty_rows = df.isnull().all(axis=1)
    empty_cols = df.isnull().all(axis=0)

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
    """转换 Excel，自动识别标题行"""
    try:
        try:
            df_raw = pd.read_excel(io.BytesIO(input_bytes), header=None, engine='openpyxl')
        except Exception:
            df_raw = pd.read_excel(io.BytesIO(input_bytes), header=None, engine='xlrd')

        if df_raw.empty:
            return None, 0

        matrices = split_matrices(df_raw)
        if not matrices:
            return None, 0

        all_rows = []
        for i, mat_info in enumerate(matrices):
            mat = mat_info['data'].copy()
            title = None

            # 自动识别标题：第一行只有一个非空单元格且内容为文本
            if len(mat) > 0:
                first_row = mat.iloc[0]
                non_null = first_row.dropna()
                if len(non_null) == 1:
                    cell = non_null.iloc[0]
                    if isinstance(cell, str) or not pd.api.types.is_number(cell):
                        title = str(cell)
                        mat = mat.iloc[1:]
                        mat = mat.dropna(how='all').dropna(axis=1, how='all')

            if title is None:
                title = f'矩阵{i + 1}'

            vec = mat.values.ravel(order='C') if not mat.empty else []
            row = [title] + vec.tolist()
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


# ========== Streamlit UI ==========
st.set_page_config(page_title="Excel矩阵展平工具", layout="centered")
st.title("📊 Excel 矩阵展平工具")
st.markdown("""
### 使用说明
1. 上传一个或多个 Excel 文件（支持 `.xlsx` 或 `.xls`）
2. 程序将自动识别每个文件中的**多个矩阵**（由全空行/全空列分隔）
3. 如果矩阵上方存在**标题行**（仅有单个非空单元格且为文本），则将该标题作为向量首元素，否则使用“矩阵N”
4. 每个矩阵按**行优先**展平，不同矩阵输出行之间**空两行**
5. 点击下载按钮获取处理后的文件
""")

uploaded_files = st.file_uploader(
    "📂 选择 Excel 文件",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    st.divider()
    st.subheader("处理选项")
    download_mode = st.radio("下载方式", ["逐个文件下载", "打包为 ZIP 下载"], horizontal=True)

    if st.button("🚀 开始处理", type="primary"):
        results = {}
        with st.spinner("处理中..."):
            for uploaded_file in uploaded_files:
                filename = uploaded_file.name
                out_filename = Path(filename).stem + ".xlsx"
                try:
                    file_bytes = uploaded_file.read()
                    out_bytes, count = convert_excel_to_rowvectors_bytes(file_bytes, filename)
                    if out_bytes is not None:
                        results[out_filename] = (out_bytes, count)
                        st.success(f"✅ {filename} → 发现 {count} 个矩阵")
                    else:
                        st.warning(f"⚠️ {filename} 未检测到有效矩阵")
                except Exception as e:
                    st.error(f"❌ {filename} 处理失败：{e}")

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
            else:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                    for out_filename, (out_bytes, _) in results.items():
                        zipf.writestr(out_filename, out_bytes)
                zip_buffer.seek(0)
                st.download_button(
                    label="📦 下载全部文件 (ZIP 压缩包)",
                    data=zip_buffer,
                    file_name="row_vectors_results.zip",
                    mime="application/zip"
                )
        else:
            st.error("没有成功处理任何文件，请检查上传的文件格式和内容。")

st.divider()
st.caption("提示：所有数据处理均在浏览器本地完成，不会上传到任何服务器。")