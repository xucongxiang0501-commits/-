import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="2026营销中心月度培训完成率看板", layout="wide", page_icon="📊")
st.title("🚀 2026营销中心月度培训完成率看板（含商学院三项）")
st.markdown("**数据来源**：《2026培训明细（3月主表）-【游月】.xlsx》 | 已自动清洗多行表头")

# ==================== 1. 加载并精确清洗数据 ====================
@st.cache_data
def load_data():
    # 关键修复：跳过前5行标题，直接从第6行真实数据开始读取
    df = pd.read_excel(
        "2026培训明细（3月主表）-【游月】.xlsx",
        sheet_name="月度培训数据-经理复盘会",
        header=None,
        skiprows=5,          # 从第6行（真实数据第一行）开始
        engine="openpyxl"
    )
    
    # 精确设置列名（完全匹配你的Excel结构）
    df.columns = [
        '月份', '城市', '培训类别', 
        '计划完成场次', '培训场次累计', '完成率', 
        '培训人数累计', '到场率', '考试平均分', 
        '讲师满意度', '课程满意度',
        '商学院_培训场次', '商学院_培训人数', '商学院_讲师满意度',
        '每日一练_应', '每日一练_实际', '每日一练_完成率',
        '每周一通_应', '每周一通_实际', '每周一通_完成率',
        '每周一练_应', '每周一练_实际', '每周一练_完成率'
    ]
    
    # 清洗月份、城市、类别（防止NaN或浮点数）
    for col in ['月份', '城市', '培训类别']:
        df[col] = df[col].astype(str).str.strip().replace(['nan', 'NaN', 'None'], '未知')
    
    # 数值列转数字
    num_cols = ['完成率', '每日一练_完成率', '每周一通_完成率', '每周一练_完成率',
                '到场率', '考试平均分', '讲师满意度', '课程满意度']
    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    
    # 删除无效行
    df = df[df['月份'] != '未知']
    df = df.reset_index(drop=True)
    
    return df

df = load_data()

# ==================== 2. 筛选器 ====================
col1, col2, col3 = st.columns(3)
with col1:
    month_filter = st.multiselect("📅 选择月份", options=sorted(df['月份'].unique()), default=["3月"])
with col2:
    city_filter = st.multiselect("🏙️ 选择城市", options=sorted(df['城市'].unique()), default=df['城市'].unique())
with col3:
    category_filter = st.multiselect("📚 选择培训类别", options=sorted(df['培训类别'].unique()), default=[])

# 过滤数据
filtered_df = df[
    (df['月份'].isin(month_filter)) &
    (df['城市'].isin(city_filter))
]
if category_filter:
    filtered_df = filtered_df[filtered_df['培训类别'].isin(category_filter)]

# ==================== 3. Tab 看板 ====================
tab1, tab2 = st.tabs(["📈 整体完成率看板", "📋 商学院三项看板"])

with tab1:
    st.subheader("整体完成率核心指标")
    c1, c2, c3 = st.columns(3)
    with c1:
        overall = filtered_df['完成率'].mean() if not filtered_df.empty else 0
        st.metric("📊 平均完成率", f"{overall:.1%}")
    with c2:
        attendance = filtered_df['到场率'].mean() if not filtered_df.empty else 0
        st.metric("👥 平均到场率", f"{attendance:.1%}")
    with c3:
        score = filtered_df['考试平均分'].mean() if not filtered_df.empty else 0
        st.metric("📝 考试平均分", f"{score:.1f}")

    fig_city = px.bar(filtered_df.groupby('城市')['完成率'].mean().reset_index(),
                      x='城市', y='完成率', title="各城市平均完成率", text_auto='.1%')
    fig_city.update_traces(marker_color='#1E88E5')
    st.plotly_chart(fig_city, use_container_width=True)

    fig_trend = px.line(filtered_df.groupby('月份')['完成率'].mean().reset_index(),
                        x='月份', y='完成率', title="1-3月完成率趋势", markers=True)
    fig_trend.update_traces(line_color='#43A047')
    st.plotly_chart(fig_trend, use_container_width=True)

with tab2:
    st.subheader("商学院三项完成率")
    c1, c2, c3 = st.columns(3)
    with c1:
        daily = filtered_df['每日一练_完成率'].mean() if not filtered_df.empty else 0
        st.metric("📅 每日一练", f"{daily:.1%}")
        st.plotly_chart(px.bar(filtered_df.groupby('城市')['每日一练_完成率'].mean().reset_index(),
                               x='城市', y='每日一练_完成率', text_auto='.1%'), use_container_width=True)
    with c2:
        call = filtered_df['每周一通_完成率'].mean() if not filtered_df.empty else 0
        st.metric("📞 每周一通", f"{call:.1%}")
        st.plotly_chart(px.bar(filtered_df.groupby('城市')['每周一通_完成率'].mean().reset_index(),
                               x='城市', y='每周一通_完成率', text_auto='.1%'), use_container_width=True)
    with c3:
        practice = filtered_df['每周一练_完成率'].mean() if not filtered_df.empty else 0
        st.metric("📝 每周一练", f"{practice:.1%}")
        st.plotly_chart(px.bar(filtered_df.groupby('城市')['每周一练_完成率'].mean().reset_index(),
                               x='城市', y='每周一练_完成率', text_auto='.1%'), use_container_width=True)

st.caption("✅ 数据已精确匹配你的Excel · 琮翔专属数据分析师制作 · 更新Excel后只需重新运行即可刷新")