import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

# 设置页面标题
st.title("📊 Excel数据处理工具 (学院精确排序版)")
st.write("自动清理空格后，按指定顺序严格排序学院数据并导出为新Excel文件")

# 第一步：上传Excel文件
st.header("第一步：上传Excel文件")
excel_file = st.file_uploader("选择Excel文件", type=['xlsx', 'xls'])

# 定义学院排序顺序
COLLEGE_ORDER = [
    "经济与管理学院",
    "法学院",
    "文学与传媒学院", 
    "数据科学与人工智能学院",
    "建筑与能源工程学院",
    "电子与电气学院",
    "机器人工程学院",
    "设计艺术学院",
    "外国语学院",
    "创新创业学院"
]

def detect_text_columns(df):
    """检测应该设置为文本格式的列"""
    text_columns = []
    
    # 常见的文本列关键词（包含中文关键词）
    text_keywords = [
        '学号', 'student', 'id', '编号', 'number', 'no', '号码',
        '电话', '手机', '联系方式', '电话号', '联系电话', 'mobile', 'phone', 'tel',
        '身份证', '身份证号', '身份证号码', 'idcard',
        '卡号', '账号', 'account',
        '邮编', '邮政编码', 'zip',
        '序列号', 'serial',
        '代码', 'code',
        '工号', '职工号',
        '宿舍号', '床位号',
        '车牌号', '车牌',
        '订单号', '订单编号',
        '准考证号', '考试号',
        '图书号', '图书编号',
        '批次号', '批号'
    ]
    
    for col in df.columns:
        col_str = str(col).lower()
        
        # 1. 根据列名判断
        is_text_column = False
        
        # 检查列名是否包含关键词
        for keyword in text_keywords:
            if keyword in col_str:
                is_text_column = True
                break
        
        # 2. 检查数据内容（如果列名不明确）
        if not is_text_column and not df.empty:
            # 取前5行数据样本
            sample_data = df[col].dropna().head(5)
            if len(sample_data) > 0:
                # 检查数据是否看起来像长数字（学号、电话等）
                for val in sample_data:
                    val_str = str(val)
                    # 如果是纯数字且长度较长（比如11位手机号、10位以上学号）
                    if val_str.replace('.', '').replace('-', '').isdigit():
                        length = len(val_str.replace('.', '').replace('-', ''))
                        if length >= 8:  # 8位以上的数字可能应该作为文本
                            is_text_column = True
                            break
        
        if is_text_column:
            text_columns.append(col)
    
    return text_columns

def convert_to_text_format(df, text_columns):
    """将指定列转换为文本格式（确保显示为字符串）"""
    df_converted = df.copy()
    
    for col in text_columns:
        if col in df_converted.columns:
            # 1. 先将所有值转为字符串
            df_converted[col] = df_converted[col].astype(str)
            
            # 2. 去除可能的科学计数法表示
            def format_number_string(s):
                if pd.isna(s):
                    return ''
                s_str = str(s)
                # 处理科学计数法（如1.23e+10）
                if 'e+' in s_str.lower() or 'e-' in s_str.lower():
                    try:
                        # 如果是浮点数科学计数法
                        num = float(s_str)
                        # 转换为整数字符串（如果可能）
                        if num.is_integer():
                            return str(int(num))
                        else:
                            return str(num)
                    except:
                        return s_str
                # 处理浮点数（如1.0）
                elif '.' in s_str:
                    try:
                        num = float(s_str)
                        if num.is_integer():
                            return str(int(num))
                    except:
                        pass
                return s_str
            
            df_converted[col] = df_converted[col].apply(format_number_string)
    
    return df_converted

def save_excel_with_text_format(df, output_stream):
    """将DataFrame保存为Excel，确保特定列以文本格式存储"""
    from openpyxl import Workbook
    
    # 创建新的工作簿
    wb = Workbook()
    ws = wb.active
    ws.title = "排序后数据"
    
    # 检测文本列
    text_columns = detect_text_columns(df)
    
    # 转换数据格式
    df_converted = convert_to_text_format(df, text_columns)
    
    # 写入表头
    for col_idx, col_name in enumerate(df_converted.columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    
    # 写入数据
    for row_idx, row in enumerate(df_converted.itertuples(index=False), 2):
        for col_idx, value in enumerate(row, 1):
            col_name = df_converted.columns[col_idx-1]
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            
            # 如果是文本列，设置单元格格式为文本
            if col_name in text_columns:
                cell.number_format = '@'  # Excel中的文本格式
            
            # 居中对齐
            cell.alignment = Alignment(horizontal='center')
    
    # 调整列宽
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 30)  # 最大宽度30
        ws.column_dimensions[column_letter].width = adjusted_width
    
    # 保存到流
    wb.save(output_stream)

# ========== 主程序开始 ==========
if excel_file is not None:
    try:
        # 第一次尝试：正常读取（假设第一行是表头）
        df = pd.read_excel(excel_file)
        df.columns = df.columns.str.strip()
        
        # 检查第一次读取是否找到"学院"列
        if '学院' not in df.columns:
            st.warning("⚠️ 第一行未找到'学院'列，正在尝试将第二行作为表头读取...")
            
            # 第二次尝试：跳过第一行读取（将第二行作为表头）
            excel_file.seek(0)
            df = pd.read_excel(excel_file, skiprows=1)
            df.columns = df.columns.str.strip()
            
            # 再次检查是否找到"学院"列
            if '学院' not in df.columns:
                st.error("❌ 即使将第二行作为表头，仍无法找到'学院'列。")
                st.write("当前文件中的列名：", df.columns.tolist())
                st.stop()
            else:
                st.success(f"✅ 已成功将第二行作为表头读取，找到'学院'列。")
        else:
            st.success(f"✅ 已成功读取，第一行即为正确的表头。")
        
        # 显示原始数据预览
        st.subheader("原始数据预览")
        st.write(f"总共有 {len(df)} 行数据")
        st.write("**处理后的所有列名是：**", df.columns.tolist())
        
        # 显示数据前几行
        st.dataframe(df.head())
        
        # 检测文本列
        text_columns = detect_text_columns(df)
        if text_columns:
            st.info(f"📋 检测到的文本格式列：")
            for col in text_columns:
                st.write(f"  - {col}")
        
        # 第二步：检查并处理"学院"列
        st.header("第二步：处理学院排序")
        
        # 核心步骤1：自动删除空格
        st.info("正在清理'学院'列中的空格...")
        df['学院'] = df['学院'].astype(str).str.strip()
        
        # 核心步骤2：规范化学院名称
        st.info("正在规范化学院名称...")
        college_name_mapping = {
            "经管学院": "经济与管理学院",
            "文传学院": "文学与传媒学院",
            "电电学院": "电子与电气学院",
            "建工学院": "建筑与能源工程学院",
            "外院": "外国语学院",
            "设艺学院": "设计艺术学院",
            "创业学院": "创新创业学院",
            "数智学院": "数据科学与人工智能学院",
            "电子与电气工程学院": "电子与电气学院",
            "创新与创业学院": "创新创业学院",
            "经管": "经济与管理学院",
            "法学": "法学院",
            "文传": "文学与传媒学院",
            "数智": "数据科学与人工智能学院",
            "建工": "建筑与能源工程学院",
            "电电": "电子与电气学院",
            "机器人": "机器人工程学院",
            "设计": "设计艺术学院",
            "外语": "外国语学院",
            "创新创业": "创新创业学院"
        }
        
        def normalize_college_name(name):
            name_clean = str(name).strip()
            return college_name_mapping.get(name_clean, name_clean)
        
        df["学院"] = df["学院"].apply(normalize_college_name)
        
        # 显示清理后的唯一值
        unique_colleges = df['学院'].unique()
        st.write("**清理空格后，'学院'列的唯一值有：**", unique_colleges.tolist())
        
        # 核心步骤3：按指定顺序重组数据
        st.info("正在按指定顺序重组数据...")
        
        # 创建一个空的DataFrame来存放排序后的结果
        sorted_dfs = []
        
        # 按照指定顺序，逐个学院提取数据
        for college in COLLEGE_ORDER:
            college_data = df[df['学院'] == college]
            if not college_data.empty:
                sorted_dfs.append(college_data)
                st.write(f"  ✓ 已提取: {college} ({len(college_data)}行)")
            else:
                st.write(f"  ⚠ 未找到: {college} (0行)")
        
        # 合并所有排序后的数据
        if sorted_dfs:
            df_sorted = pd.concat(sorted_dfs, ignore_index=True)
            
            # 处理不在指定顺序中的其他学院
            other_colleges = set(df['学院'].unique()) - set(COLLEGE_ORDER)
            if other_colleges:
                st.warning(f"发现以下未在排序列表中的学院，它们将被放在最后：{list(other_colleges)}")
                other_data = df[df['学院'].isin(other_colleges)]
                df_sorted = pd.concat([df_sorted, other_data], ignore_index=True)
            
            # 显示排序后的数据
            st.subheader("排序后的数据预览")
            st.write(f"排序后总共有 {len(df_sorted)} 行数据")
            
            # 预览前10行
            st.dataframe(df_sorted.head(10))
            
            # 第三步：导出Excel文件
            st.header("第三步：导出排序后的Excel文件")
            
            if st.button("📥 生成并导出Excel文件", type="primary"):
                with st.spinner("正在生成Excel文件，请稍候..."):
                    # 创建内存中的Excel文件
                    output = io.BytesIO()
                    
                    # 使用自定义函数保存，确保文本格式
                    save_excel_with_text_format(df_sorted, output)
                    
                    output.seek(0)
                    
                    # 提供下载
                    st.success("🎉 Excel文件生成成功！")
                    
                    st.download_button(
                        label="点击下载Excel文件",
                        data=output,
                        file_name="按学院排序的数据_文本格式.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_excel"
                    )
        
        else:
            st.error("未匹配到任何指定学院的数据。请检查'学院'列的值。")
            st.stop()
    
    except Exception as e:
        st.error(f"处理文件失败: {str(e)}")
        st.exception(e)
        st.write("请检查文件格式是否正确，或联系管理员。")

else:
    st.info("👆 请先上传Excel文件")
