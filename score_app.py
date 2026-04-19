# -*- coding: utf-8 -*-
"""
Created on Fri Apr 17 10:07:32 2026

@author: Lenovo
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from collections import Counter

# ==================== 页面配置 ====================
st.set_page_config(
    page_title="试卷自动判分系统",
    page_icon="📊",
    layout="wide"
)

st.title("📊 试卷自动判分系统")
st.markdown("上传标准答案和学生答卷，系统自动判分并生成分析报告")

# ==================== 侧边栏：参数设置 ====================
with st.sidebar:
    st.header("⚙️ 评分参数设置")
    
    st.subheader("单选题")
    single_points = st.number_input("每题分值", value=0.5, step=0.1, key="single_points")
    single_rows = st.text_input("行范围（格式：起始-结束）", value="1-21", key="single_rows")
    
    st.subheader("多选题")
    multi_points = st.number_input("每题分值", value=0.7, step=0.1, key="multi_points")
    multi_rows = st.text_input("行范围（格式：起始-结束）", value="22-32", key="multi_rows")
    
    st.subheader("判断题")
    judge_points = st.number_input("每题分值", value=0.3, step=0.1, key="judge_points")
    judge_rows = st.text_input("行范围（格式：起始-结束）", value="33-43", key="judge_rows")
    
    st.divider()
    st.caption("提示：行范围需与Excel中的实际布局一致")

# ==================== 主界面：文件上传 ====================
col1, col2 = st.columns(2)

with col1:
    st.subheader("📄 标准答案")
    std_file = st.file_uploader(
        "上传标准答案文件（.xlsx 或 .xls）",
        type=["xlsx", "xls"],
        key="std_file"
    )

with col2:
    st.subheader("👨‍🎓 学生答卷")
    student_files = st.file_uploader(
        "上传学生答卷文件（可多选）",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="student_files"
    )

# ==================== Excel读取函数（使用pandas，指定engine） ====================
def load_excel_data(file_bytes, file_name):
    """使用pandas读取Excel文件，自动选择engine"""
    try:
        # 根据文件扩展名选择engine
        if file_name.endswith('.xls'):
            # 旧版Excel使用xlrd
            df = pd.read_excel(BytesIO(file_bytes), sheet_name=0, header=None, engine='xlrd')
        else:
            # 新版Excel使用openpyxl，但Streamlit Cloud已安装
            df = pd.read_excel(BytesIO(file_bytes), sheet_name=0, header=None, engine='openpyxl')
        
        # 填充NaN为空字符串
        df = df.fillna('')
        return df
    except Exception as e:
        st.error(f"读取文件 {file_name} 失败：{str(e)}")
        return None

def parse_row_range(range_str):
    """解析行范围字符串"""
    try:
        start, end = range_str.split('-')
        return int(start), int(end)
    except:
        return 1, 21

def check_student(df_std, df_stu, single_range, multi_range, judge_range,
                  single_pts, multi_pts, judge_pts):
    """对比学生答案与标准答案"""
    single_wrong = []
    multi_wrong = []
    judge_wrong = []
    
    s_start, s_end = parse_row_range(single_range)
    m_start, m_end = parse_row_range(multi_range)
    j_start, j_end = parse_row_range(judge_range)
    
    # 获取最大行数和列数
    max_row = max(df_std.shape[0], df_stu.shape[0])
    max_col = max(df_std.shape[1], df_stu.shape[1])
    
    # 扩展DataFrame到相同大小
    if df_std.shape[0] < max_row:
        df_std = pd.concat([df_std, pd.DataFrame('', index=range(max_row - df_std.shape[0]), columns=df_std.columns)], ignore_index=True)
    if df_stu.shape[0] < max_row:
        df_stu = pd.concat([df_stu, pd.DataFrame('', index=range(max_row - df_stu.shape[0]), columns=df_stu.columns)], ignore_index=True)
    if df_std.shape[1] < max_col:
        for _ in range(max_col - df_std.shape[1]):
            df_std[len(df_std.columns)] = ''
    if df_stu.shape[1] < max_col:
        for _ in range(max_col - df_stu.shape[1]):
            df_stu[len(df_stu.columns)] = ''
    
    for i in range(max_row):
        excel_row = i + 1  # Excel行号从1开始
        
        for j in range(max_col):
            std_val = str(df_std.iloc[i, j]) if df_std.iloc[i, j] != '' else ''
            stu_val = str(df_stu.iloc[i, j]) if df_stu.iloc[i, j] != '' else ''
            
            if std_val != stu_val and std_val != '':
                # 根据行号判断题型
                if s_start <= excel_row <= s_end:
                    # 从上一行获取题号
                    if i > 0:
                        title_val = df_std.iloc[i-1, j]
                        try:
                            if title_val != '':
                                q_num = int(float(title_val))
                                if 1 <= q_num <= 100 and q_num not in single_wrong:
                                    single_wrong.append(q_num)
                        except:
                            pass
                elif m_start <= excel_row <= m_end:
                    if i > 0:
                        title_val = df_std.iloc[i-1, j]
                        try:
                            if title_val != '':
                                q_num = int(float(title_val))
                                if 1 <= q_num <= 50 and q_num not in multi_wrong:
                                    multi_wrong.append(q_num)
                        except:
                            pass
                elif j_start <= excel_row <= j_end:
                    if i > 0:
                        title_val = df_std.iloc[i-1, j]
                        try:
                            if title_val != '':
                                q_num = int(float(title_val))
                                if 1 <= q_num <= 50 and q_num not in judge_wrong:
                                    judge_wrong.append(q_num)
                        except:
                            pass
    
    # 计算总分
    score = 100 - len(single_wrong) * single_pts - len(multi_wrong) * multi_pts - len(judge_wrong) * judge_pts
    return single_wrong, multi_wrong, judge_wrong, max(0, score)

# ==================== 执行判分 ====================
if std_file and student_files:
    if st.button("🚀 开始判分", type="primary"):
        with st.spinner("判分进行中，请稍候..."):
            try:
                # 加载标准答案
                std_df = load_excel_data(std_file.read(), std_file.name)
                if std_df is None:
                    st.error("标准答案文件读取失败")
                    st.stop()
                
                results = []
                all_single_wrong = []
                all_multi_wrong = []
                all_judge_wrong = []
                
                single_pts = st.session_state.single_points
                multi_pts = st.session_state.multi_points
                judge_pts = st.session_state.judge_points
                single_range = st.session_state.single_rows
                multi_range = st.session_state.multi_rows
                judge_range = st.session_state.judge_rows
                
                progress_bar = st.progress(0)
                
                for idx, stu_file in enumerate(student_files):
                    stu_df = load_excel_data(stu_file.read(), stu_file.name)
                    
                    if stu_df is None:
                        st.warning(f"跳过无法读取的文件：{stu_file.name}")
                        continue
                    
                    single_wrong, multi_wrong, judge_wrong, score = check_student(
                        std_df, stu_df, single_range, multi_range, judge_range,
                        single_pts, multi_pts, judge_pts
                    )
                    
                    name = stu_file.name.replace('.xlsx', '').replace('.xls', '')
                    
                    results.append({
                        "姓名": name,
                        "总分": round(score, 1),
                        "单选错误数": len(single_wrong),
                        "单选错题号": single_wrong,
                        "多选错误数": len(multi_wrong),
                        "多选错题号": multi_wrong,
                        "判断错误数": len(judge_wrong),
                        "判断错题号": judge_wrong
                    })
                    
                    all_single_wrong.extend(single_wrong)
                    all_multi_wrong.extend(multi_wrong)
                    all_judge_wrong.extend(judge_wrong)
                    
                    progress_bar.progress((idx + 1) / len(student_files))
                
                st.success(f"✅ 判分完成！共处理 {len(results)} 名学生")
                
                if len(results) == 0:
                    st.warning("没有成功处理任何学生文件")
                    st.stop()
                
                # 显示成绩汇总表
                st.subheader("📋 成绩汇总表")
                df_results = pd.DataFrame([
                    {k: v for k, v in r.items() if k not in ['单选错题号', '多选错题号', '判断错题号']}
                    for r in results
                ])
                st.dataframe(df_results, use_container_width=True)
                
                # 详细错题报告
                st.subheader("📝 详细错题报告")
                for r in results:
                    with st.expander(f"{r['姓名']} - 总分：{r['总分']}分"):
                        col_a, col_b, col_c = st.columns(3)
                        with col_a:
                            st.metric("单选题错误", r['单选错误数'])
                            if r['单选错题号']:
                                st.write(f"错题号：{', '.join(map(str, sorted(r['单选错题号'])))}")
                        with col_b:
                            st.metric("多选题错误", r['多选错误数'])
                            if r['多选错题号']:
                                st.write(f"错题号：{', '.join(map(str, sorted(r['多选错题号'])))}")
                        with col_c:
                            st.metric("判断题错误", r['判断错误数'])
                            if r['判断错题号']:
                                st.write(f"错题号：{', '.join(map(str, sorted(r['判断错题号'])))}")
                
                # 群体错题分析
                if len(results) > 1:
                    st.subheader("📊 群体错题分析")
                    threshold = len(results) / 2
                    
                    single_counter = Counter(all_single_wrong)
                    multi_counter = Counter(all_multi_wrong)
                    judge_counter = Counter(all_judge_wrong)
                    
                    single_over = [q for q, cnt in single_counter.items() if cnt > threshold]
                    multi_over = [q for q, cnt in multi_counter.items() if cnt > threshold]
                    judge_over = [q for q, cnt in judge_counter.items() if cnt > threshold]
                    
                    if single_over or multi_over or judge_over:
                        st.write("以下题目错误率超过50%，建议重点讲解：")
                        if single_over:
                            st.write(f"🔴 单选题：{', '.join(map(str, sorted(single_over)))}")
                        if multi_over:
                            st.write(f"🔴 多选题：{', '.join(map(str, sorted(multi_over)))}")
                        if judge_over:
                            st.write(f"🔴 判断题：{', '.join(map(str, sorted(judge_over)))}")
                    else:
                        st.write("✅ 所有题目错误率均未超过50%")
                
                # 下载报告
                st.subheader("💾 下载报告")
                # 生成符合格式的成绩报告
                report_text = ""
                # 写入每个学生的成绩
                for idx, r in enumerate(results, 1):
                    # 处理单选错题格式化
                    single_str = f"{r['单选错误数']}"
                    if r['单选错题号']:
                        single_str += f"（题号：{', '.join(map(str, sorted(r['单选错题号'])))})"
                
                    # 处理多选错题格式化
                    multi_str = f"{r['多选错误数']}"
                    if r['多选错题号']:
                        multi_str += f"（题号：{', '.join(map(str, sorted(r['多选错题号'])))})"
                    # 处理判断错题格式化
                    judge_str = f"{r['判断错误数']}"
                    if r['判断错题号']:
                        judge_str += f"（题号：{', '.join(map(str, sorted(r['判断错题号'])))})"
                
                
                    report_text += f"{idx}、{r['姓名']}总分：{r['总分']}分，"
                    report_text += f"单选错误个数：{single_str}；"
                    report_text += f"多选错误个数：{multi_str}；"
                    report_text += f"判断错误个数：{judge_str}。\r\n"  # 使用 Windows 换行符 \r\n
                    
                    
                    
                # 所有考生结束后换行并空一行
                report_text += "\r\n"  # 空一行（Windows格式）
               
                # 写入错误率超过50%的题目汇总（使用之前已计算的统计结果）
                if len(results) > 1:
                  # 使用前面群体错题分析中已经计算好的变量            
                   over_parts = []
                   if single_over:
                       over_parts.append("单选题" + "、".join(str(i) for i in sorted(single_over)))
                   if multi_over:
                       over_parts.append("多选题" + "、".join(str(i) for i in sorted(multi_over)))
                   if judge_over:
                       over_parts.append("判断题" + "、".join(str(i) for i in sorted(judge_over)))
    
                   if over_parts:
                       report_text += "错误率超过50%的题目：" + "；".join(over_parts) + "。"
                   else:
                       report_text += "没有错误率超过50%的题目。"
                else:
                     # 只有一名学生时，不统计错误率
                     report_text += "（只有一名学生，无法统计错误率超过50%的题目）"
               
                
               
                st.download_button(
                    label="📥 下载成绩报告（TXT格式）",
                    data=report_text,
                    file_name="成绩报告.txt",
                    mime="text/plain"
                )
                
            except Exception as e:
                st.error(f"❌ 处理过程中出错：{str(e)}")
                import traceback
                st.code(traceback.format_exc())

else:
    st.info("👈 请先上传标准答案和学生答卷文件")

st.divider()
st.caption("试卷自动判分系统 | 基于Streamlit构建")
