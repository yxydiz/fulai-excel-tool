import streamlit as st
import os
import re
import io
import pandas as pd
from datetime import datetime

# ==================== 页面配置 ====================
st.set_page_config(
    page_title="福来Excel处理工具",
    page_icon="📊",
    layout="centered"
)

st.title("📊 福来 Excel 自动化处理工具")
st.markdown("支持三种文件类型的自动处理：CPTI日报、种草日报底表、搜索SOVSOC")
st.divider()

# ==================== 配置常量 ====================
BABY_ACCOUNTS = ["福来-种草婴儿油", "福来-种草婴童线"]
FILE_TYPES = {
    "cpti": {
        "name": "福来日报-cpti",
        "description": "单sheet文件，N-AM列转数字，创意名称分列",
        "sheets": 1
    },
    "grass": {
        "name": "福来种草日报-底表",
        "description": "双sheet文件（信息流定向+搜索关键词），合并后处理",
        "sheets": 2
    },
    "sov": {
        "name": "福来-搜索SOVSOC",
        "description": "单sheet文件，N-AF列转数字，删除指定列",
        "sheets": 1
    }
}

# ==================== 核心处理函数 ====================

def col_letter_to_index(letter):
    """Excel列字母转索引"""
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch.upper()) - ord('A') + 1)
    return idx - 1

def convert_columns_to_numeric(df, start_col="N", end_col="AM"):
    """将指定列范围中可转换为数字的文本转为数字，不可转换的保留原文本"""
    start_idx = col_letter_to_index(start_col)
    end_idx = col_letter_to_index(end_col)
    cols_to_convert = df.columns[start_idx:end_idx+1]

    for col in cols_to_convert:
        numeric_series = pd.to_numeric(df[col], errors='coerce')
        df[col] = numeric_series.where(numeric_series.notna(), df[col])
    return df

def split_creative_name(df, creative_col="创意名称", insert_before="创意ID", n_parts=9):
    """将创意名称按'-'分列为n_parts列，插入到creative_col和insert_before之间"""
    if creative_col not in df.columns or insert_before not in df.columns:
        st.warning(f"缺少 '{creative_col}' 或 '{insert_before}' 列，跳过拆分步骤")
        return df

    col_names = df.columns.tolist()
    pos_before = col_names.index(insert_before)

    split_df = df[creative_col].astype(str).str.split('-', n=n_parts-1, expand=True)
    for i in range(split_df.shape[1], n_parts):
        split_df[i] = ""
    split_df.columns = [f"创意名称分列_{j+1}" for j in range(n_parts)]

    left_cols = col_names[:pos_before]
    right_cols = col_names[pos_before:]
    return pd.concat([df[left_cols], split_df, df[right_cols]], axis=1)

def filter_and_split_sheets(df, child_account_col="子账户名称"):
    """按子账户名称拆分为婴童线和常规"""
    if child_account_col not in df.columns:
        raise KeyError(f"缺少列：{child_account_col}")
    df_baby = df[df[child_account_col].isin(BABY_ACCOUNTS)]
    df_regular = df[~df[child_account_col].isin(BABY_ACCOUNTS)]
    return df_baby, df_regular

def drop_empty_time_rows(df, time_col="时间"):
    """删除时间列为空的行"""
    before = len(df)
    df = df.dropna(subset=[time_col])
    after = len(df)
    if before != after:
        st.info(f"删除了 {before - after} 行时间列为空的数据")
    return df

def delete_columns(df, cols_to_del):
    """删除指定列"""
    existing_cols = [col for col in cols_to_del if col in df.columns]
    if existing_cols:
        df = df.drop(columns=existing_cols)
        st.info(f"已删除列：{existing_cols}")
    return df

# ==================== 三种处理器的类 ====================

class BaseProcessor:
    """处理器基类"""
    def __init__(self, uploaded_file):
        self.uploaded_file = uploaded_file
        self.file_name = uploaded_file.name

    def read_excel(self, sheet_name=0):
        return pd.read_excel(self.uploaded_file, sheet_name=sheet_name, engine='openpyxl')

    def write_output(self, df_baby, df_regular, date_str, suffix):
        """将结果写入内存中的Excel文件"""
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_baby.to_excel(writer, sheet_name="婴童线", index=False)
            df_regular.to_excel(writer, sheet_name="常规", index=False)
        output.seek(0)
        return output

class CPTIProcessor(BaseProcessor):
    """福来日报-cpti 处理器"""
    def process(self):
        df = self.read_excel()
        st.info(f"原始数据：{df.shape[0]} 行 × {df.shape[1]} 列")

        # N列到AM列转数字
        df = convert_columns_to_numeric(df, "N", "AM")
        st.success("✅ N-AM列数字转换完成")

        # 创意名称分9列
        df = split_creative_name(df, "创意名称", "创意ID", 9)
        st.success("✅ 创意名称分列完成（9列）")

        # 按子账户拆分
        df_baby, df_regular = filter_and_split_sheets(df)
        st.success(f"✅ 数据拆分完成：婴童线 {len(df_baby)} 行，常规 {len(df_regular)} 行")

        date_str = datetime.now().strftime("%Y%m%d")
        output = self.write_output(df_baby, df_regular, date_str, "cpti处理后")
        out_name = f"{date_str}_福来日报-cpti_处理后.xlsx"
        return output, out_name, len(df_baby), len(df_regular)

class GrassProcessor(BaseProcessor):
    """福来种草日报-底表 处理器"""
    def process(self):
        # 读取两个sheet
        df_info = self.read_excel("福来信息流定向")
        df_keyword = self.read_excel("福来搜索关键词")
        st.info(f"信息流定向：{df_info.shape[0]} 行，搜索关键词：{df_keyword.shape[0]} 行")

        # 检查列数一致
        if df_info.shape[1] != df_keyword.shape[1]:
            raise ValueError(f"列数不一致：{df_info.shape[1]} vs {df_keyword.shape[1]}")

        # 强制列名一致后拼接
        df_keyword_aligned = df_keyword.copy()
        df_keyword_aligned.columns = df_info.columns
        df_merged = pd.concat([df_info, df_keyword_aligned], ignore_index=True)
        st.success(f"✅ 两sheet合并完成，共 {len(df_merged)} 行")

        # 创意名称分9列
        df_merged = split_creative_name(df_merged, "创意名称", "创意ID", 9)
        st.success("✅ 创意名称分列完成（9列）")

        # 按子账户拆分
        df_baby, df_regular = filter_and_split_sheets(df_merged)
        st.success(f"✅ 数据拆分完成：婴童线 {len(df_baby)} 行，常规 {len(df_regular)} 行")

        date_str = datetime.now().strftime("%Y%m%d")
        output = self.write_output(df_baby, df_regular, date_str, "种草日报处理后")
        out_name = f"{date_str}_福来种草日报-底表_处理后.xlsx"
        return output, out_name, len(df_baby), len(df_regular)

class SOVProcessor(BaseProcessor):
    """福来-搜索SOVSOC 处理器"""
    def process(self):
        df = self.read_excel()
        st.info(f"原始数据：{df.shape[0]} 行 × {df.shape[1]} 列")

        # N列到AF列转数字
        df = convert_columns_to_numeric(df, "N", "AF")
        st.success("✅ N-AF列数字转换完成")

        # 删除指定列
        cols_to_del = ["营销诉求", "投放位置", "推广目标", "竞价策略"]
        df = delete_columns(df, cols_to_del)

        # 删除时间列为空的行
        df = drop_empty_time_rows(df, "时间")

        # 按子账户拆分
        df_baby, df_regular = filter_and_split_sheets(df)
        st.success(f"✅ 数据拆分完成：婴童线 {len(df_baby)} 行，常规 {len(df_regular)} 行")

        date_str = datetime.now().strftime("%Y%m%d")
        output = self.write_output(df_baby, df_regular, date_str, "SOVSOC处理后")
        out_name = f"{date_str}_福来-搜索SOVSOC_处理后.xlsx"
        return output, out_name, len(df_baby), len(df_regular)

# ==================== 文件类型检测 ====================

def detect_file_type(file_name):
    """根据文件名自动检测文件类型"""
    name_lower = file_name.lower()
    if "cpti" in name_lower:
        return "cpti"
    elif "底表" in file_name or "种草日报" in file_name:
        return "grass"
    elif "sov" in name_lower or "soc" in name_lower:
        return "sov"
    return None

# ==================== 主界面 ====================

st.subheader("📁 上传文件")

uploaded_file = st.file_uploader(
    "拖拽或点击上传 Excel 文件",
    type=["xlsx", "xls"],
    help="支持 .xlsx 和 .xls 格式"
)

if uploaded_file is not None:
    # 自动检测文件类型
    detected_type = detect_file_type(uploaded_file.name)

    if detected_type:
        file_info = FILE_TYPES[detected_type]
        st.success(f"🔍 自动识别文件类型：**{file_info['name']}**")
        st.caption(f"说明：{file_info['description']}")
    else:
        st.warning("⚠️ 无法自动识别文件类型，请手动选择")
        detected_type = None

    # 手动选择（如果自动检测失败或需要覆盖）
    col1, col2 = st.columns(2)
    with col1:
        manual_type = st.selectbox(
            "或手动选择文件类型",
            ["自动检测"] + [f"{k}: {v['name']}" for k, v in FILE_TYPES.items()],
            index=0
        )

    # 解析手动选择
    if manual_type != "自动检测":
        selected_type = manual_type.split(":")[0]
    else:
        selected_type = detected_type

    if selected_type and selected_type in FILE_TYPES:
        st.divider()
        st.subheader("⚙️ 处理配置")

        # 显示婴童线账户
        with st.expander("查看婴童线账户配置"):
            st.write(BABY_ACCOUNTS)

        # 处理按钮
        if st.button("🚀 开始处理", type="primary", use_container_width=True):
            with st.spinner("正在处理中，请稍候..."):
                try:
                    # 根据类型选择处理器
                    if selected_type == "cpti":
                        processor = CPTIProcessor(uploaded_file)
                    elif selected_type == "grass":
                        processor = GrassProcessor(uploaded_file)
                    elif selected_type == "sov":
                        processor = SOVProcessor(uploaded_file)
                    else:
                        st.error("未知文件类型")
                        st.stop()

                    # 执行处理
                    output, out_name, baby_count, regular_count = processor.process()

                    # 显示结果
                    st.divider()
                    st.subheader("✅ 处理完成")

                    result_col1, result_col2 = st.columns(2)
                    with result_col1:
                        st.metric("婴童线数据", f"{baby_count} 行")
                    with result_col2:
                        st.metric("常规数据", f"{regular_count} 行")

                    # 下载按钮
                    st.download_button(
                        label=f"📥 下载处理结果：{out_name}",
                        data=output,
                        file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

                except Exception as e:
                    st.error(f"❌ 处理失败：{str(e)}")
                    st.exception(e)
    else:
        st.error("请先选择有效的文件类型")

# ==================== 底部说明 ====================
st.divider()
with st.expander("📖 使用说明"):
    st.markdown("""
    **支持的文件类型：**

    1. **福来日报-cpti** 
       - 单sheet文件
       - 自动将N列到AM列的文本转为数字
       - 创意名称按'-'拆分为9列
       - 按子账户拆分为婴童线/常规

    2. **福来种草日报-底表**
       - 包含两个sheet：福来信息流定向 + 福来搜索关键词
       - 自动合并两个sheet后处理
       - 创意名称分列 + 婴童线/常规拆分

    3. **福来-搜索SOVSOC**
       - 单sheet文件
       - N列到AF列转数字
       - 删除营销诉求、投放位置、推广目标、竞价策略四列
       - 删除时间列为空的行
       - 婴童线/常规拆分

    **婴童线账户：** 福来-种草婴儿油、福来-种草婴童线
    """)
