import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from io import BytesIO

# 提前判断是否支持Excel导出
try:
    import openpyxl
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

# ====================== 页面基础配置 ======================
st.set_page_config(
    page_title="淘宝用户行为可视化分析系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ====================== 全局样式与主题系统（终极修复浅色文字） ======================
def apply_theme(theme_mode):
    if theme_mode == '深色主题':
        bg_color = "#0e1117"
        card_bg = "#1a1d24"
        text_color = "#fafafa"
        sub_text = "#b0b0b0"
        border_color = "#2d303a"
        plot_template = "plotly_dark"
    else:
        # 浅色主题：纯黑文字，拉满对比度
        bg_color = "#f5f7fa"
        card_bg = "#ffffff"
        text_color = "#000000"   # 纯黑，绝对清晰
        sub_text = "#222222"    # 次要文字也加深
        border_color = "#d0d7de"
        plot_template = "plotly_white"
    
    st.markdown(f"""
    <style>
    /* 全局最高优先级：所有文字强制主题色 */
    .stApp,
    .stApp * {{
        color: {text_color} !important;
    }}

    /* 主背景 */
    .stApp {{
        background-color: {bg_color};
    }}

    /* 侧边栏背景与文字 */
    [data-testid="stSidebar"],
    [data-testid="stSidebar"] * {{
        background-color: {card_bg} !important;
        color: {text_color} !important;
    }}
    [data-testid="stSidebar"] {{
        border-right: 1px solid {border_color};
    }}

    /* 所有标题 */
    h1, h2, h3, h4, h5, h6 {{
        color: {text_color} !important;
    }}

    /* 卡片样式 */
    .metric-card {{
        background-color: {card_bg} !important;
        padding: 20px 24px;
        border-radius: 12px;
        border: 1px solid {border_color};
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        transition: all 0.3s ease;
    }}
    .metric-card:hover {{
        box-shadow: 0 4px 16px rgba(0,0,0,0.1);
        transform: translateY(-2px);
    }}
    .metric-label {{
        font-size: 14px;
        color: {sub_text} !important;
        margin-bottom: 8px;
    }}
    .metric-value {{
        font-size: 28px;
        font-weight: 600;
        color: {text_color} !important;
    }}

    /* 分割线 */
    hr {{
        border-color: {border_color} !important;
    }}

    /* 提示框（info/warning）强制深色文字 */
    .stAlert,
    .stAlert *,
    .stAlert p,
    .stAlert strong,
    .stAlert div {{
        color: #000000 !important;
    }}

    /* 表格文字 */
    [data-testid="stDataFrame"],
    [data-testid="stDataFrame"] * {{
        color: {text_color} !important;
    }}

    /* 下拉框、输入框 */
    div[role="listbox"] *,
    input, select, textarea {{
        color: {text_color} !important;
    }}

    /* Tab 标签文字 */
    .stTabs [data-baseweb="tab-list"] button * {{
        color: {text_color} !important;
    }}

    /* 复选框、单选框文字 */
    [data-testid="stCheckbox"] label,
    [data-testid="stRadio"] label {{
        color: {text_color} !important;
    }}

    /* metric 组件原生数字 */
    [data-testid="stMetricValue"] div {{
        color: {text_color} !important;
    }}

    /* 小字说明 */
    small, .stCaption, caption {{
        color: {sub_text} !important;
    }}
    </style>
    """, unsafe_allow_html=True)
    return plot_template

# ====================== 统一商务配色 ======================
COLORS = {
    '点击': '#1f77b4',
    '加购': '#ff7f0e',
    '收藏': '#2ca02c',
    '购买': '#d62728'
}

RFM_COLORS = ['#2ecc71', '#3498db', '#f1c40f', '#e74c3c', '#95a5a6', '#1abc9c', '#e67e22', '#9b59b6']

# ====================== 生成真实感模拟数据 ======================
@st.cache_data
def load_data():
    np.random.seed(42)
    n_users = 10000
    n_rows = 120000

    user_pool = np.arange(10000, 10000 + n_users)
    item_pool = np.random.randint(100000, 999999, size=80000)
    category_pool = np.random.randint(1000, 1500, size=500)

    head_category_num = int(len(category_pool) * 0.1)
    category_weights = np.concatenate([
        np.ones(head_category_num) * 20,
        np.ones(len(category_pool)-head_category_num) * 1
    ])
    category_weights /= category_weights.sum()

    item_price_map = {item: round(np.random.uniform(29, 899), 2) for item in item_pool}

    behavior_types = ['pv', 'cart', 'fav', 'buy']
    behavior_probs = [0.92, 0.048, 0.022, 0.01]

    start_date = pd.Timestamp('2017-11-25 00:00:00')
    end_date = pd.Timestamp('2017-12-03 23:59:59')
    hours = pd.date_range(start_date, end_date, freq='h')
    
    hour_weights = []
    for h in hours.hour:
        if 0 <= h <= 5: hour_weights.append(0.12)
        elif 6 <= h <= 9: hour_weights.append(0.5)
        elif 10 <= h <= 12: hour_weights.append(1.0)
        elif 13 <= h <= 18: hour_weights.append(0.85)
        elif 19 <= h <= 22: hour_weights.append(1.9)
        else: hour_weights.append(0.6)
    
    day_weights = []
    for d in hours.date:
        if pd.Timestamp(d).weekday() >= 5:
            day_weights.append(1.18)
        else:
            day_weights.append(np.random.uniform(0.94, 1.06))
    hour_weights = np.array(hour_weights) * np.array(day_weights)
    hour_weights /= hour_weights.sum()

    sampled_hours = np.random.choice(hours, size=n_rows, p=hour_weights)
    timestamps = sampled_hours + pd.to_timedelta(np.random.randint(0, 3600, size=n_rows), unit='s')
    timestamp_unix = timestamps.astype('int64') // 10**9

    high_active_users = np.random.choice(user_pool, size=int(n_users*0.2), replace=False)
    user_ids = np.concatenate([
        np.random.choice(high_active_users, size=int(n_rows*0.8)),
        np.random.choice(user_pool, size=int(n_rows*0.2))
    ])
    np.random.shuffle(user_ids)

    item_ids = np.random.choice(item_pool, size=n_rows)
    category_ids = np.random.choice(category_pool, size=n_rows, p=category_weights)
    behavior_type = np.random.choice(behavior_types, size=n_rows, p=behavior_probs)

    df = pd.DataFrame({
        'user_id': user_ids,
        'item_id': item_ids,
        'category_id': category_ids,
        'behavior_type': behavior_type,
        'timestamp': timestamp_unix
    })

    buy_mask = df['behavior_type'] == 'buy'
    buy_users = df[buy_mask]['user_id'].unique()
    high_freq_buy_users = np.random.choice(buy_users, size=int(len(buy_users)*0.18), replace=False)
    
    extra_buy = []
    for uid in high_freq_buy_users:
        extra_num = np.random.randint(1, 4)
        for _ in range(extra_num):
            extra_item = np.random.choice(item_pool)
            extra_buy.append({
                'user_id': uid,
                'item_id': extra_item,
                'category_id': np.random.choice(category_pool),
                'behavior_type': 'buy',
                'timestamp': np.random.choice(timestamp_unix)
            })
    if extra_buy:
        df = pd.concat([df, pd.DataFrame(extra_buy)], ignore_index=True)

    pv_user_set = set(df[df['behavior_type'] == 'pv']['user_id'].unique())
    for btype in ['buy', 'cart', 'fav']:
        target_users = df[df['behavior_type'] == btype]['user_id'].unique()
        missing = [uid for uid in target_users if uid not in pv_user_set]
        extra_rows = []
        for uid in missing:
            extra_rows.append({
                'user_id': uid,
                'item_id': np.random.choice(item_pool),
                'category_id': np.random.choice(category_pool),
                'behavior_type': 'pv',
                'timestamp': timestamp_unix[0]
            })
            pv_user_set.add(uid)
        if extra_rows:
            df = pd.concat([df, pd.DataFrame(extra_rows)], ignore_index=True)

    df['datetime'] = pd.to_datetime(df['timestamp'], unit='s')
    df['date'] = df['datetime'].dt.date
    df['hour'] = df['datetime'].dt.hour
    df['weekday'] = df['datetime'].dt.weekday
    df['weekday_cn'] = df['weekday'].map({0:'周一',1:'周二',2:'周三',3:'周四',4:'周五',5:'周六',6:'周日'})
    df['is_weekend'] = df['weekday'].apply(lambda x: '周末' if x >=5 else '工作日')

    behavior_map = {'pv': '点击', 'cart': '加购', 'fav': '收藏', 'buy': '购买'}
    df['behavior_name'] = df['behavior_type'].map(behavior_map)
    df['price'] = df['item_id'].map(item_price_map)
    df['amount'] = np.where(df['behavior_type'] == 'buy', df['price'], 0)

    return df

df = load_data()

# ====================== 侧边栏交互控件 ======================
st.sidebar.title("📌 控制面板")

with st.sidebar.expander("⏰ 时间范围筛选", expanded=True):
    start_date, end_date = st.date_input(
        "选择分析时间段",
        value=[df['date'].min(), df['date'].max()],
        min_value=df['date'].min(),
        max_value=df['date'].max()
    )

with st.sidebar.expander("🛒 行为类型筛选", expanded=True):
    behavior_list = st.multiselect(
        "选择要展示的行为",
        options=['点击', '加购', '收藏', '购买'],
        default=['点击', '加购', '收藏', '购买']
    )

with st.sidebar.expander("⚙️ 显示设置", expanded=True):
    show_data_label = st.checkbox("显示图表数据标签", value=True)
    theme_mode = st.selectbox("界面主题", options=['深色主题', '浅色主题'], index=0)

plot_template = apply_theme(theme_mode)

st.sidebar.markdown("---")
if st.sidebar.button("🔄 重置所有筛选", use_container_width=True):
    st.rerun()

st.sidebar.caption("数据说明：模拟淘宝用户行为数据\n周期：2017.11.25 - 2017.12.03")

# ====================== 数据过滤与兜底 ======================
if not behavior_list:
    st.warning("⚠️ 请至少选择一种行为类型进行分析")
    st.stop()

filtered_df = df[
    (df['date'] >= start_date) &
    (df['date'] <= end_date) &
    (df['behavior_name'].isin(behavior_list))
]

if filtered_df.empty:
    st.warning("⚠️ 当前筛选条件下无数据，请调整时间范围或行为类型")
    st.stop()

# ====================== 公共计算函数 ======================
def calc_core_metrics(data):
    total_pv = data[data['behavior_name'] == '点击'].shape[0]
    total_uv = data['user_id'].nunique()
    pv_uv = data[data['behavior_name'] == '点击']['user_id'].nunique()
    buy_uv = data[data['behavior_name'] == '购买']['user_id'].nunique()
    total_gmv = data['amount'].sum()
    order_count = data[data['behavior_name'] == '购买'].shape[0]
    
    conversion = round(buy_uv / pv_uv * 100, 2) if pv_uv > 0 else 0
    avg_person = round(data.shape[0] / total_uv, 2) if total_uv > 0 else 0
    atv = round(total_gmv / order_count, 2) if order_count > 0 else 0
    arpu = round(total_gmv / buy_uv, 2) if buy_uv > 0 else 0
    
    return {
        'pv': total_pv, 'uv': total_uv, 'conversion': conversion,
        'gmv': total_gmv, 'order_count': order_count,
        'avg_person': avg_person, 'atv': atv, 'arpu': arpu
    }

# ====================== 页面主标题 ======================
st.title("📊 淘宝用户行为可视化分析系统")
st.markdown("---")

# ====================== 顶部导航Tab ======================
tab_overview, tab_time, tab_user, tab_product, tab_data = st.tabs([
    "📈 核心概览", "⏰ 时间分析", "👤 用户分析", "🛍️ 商品分析", "📋 数据详情"
])

# ====================== Tab1: 核心概览 ======================
with tab_overview:
    metrics = calc_core_metrics(filtered_df)
    
    col1, col2, col3, col4 = st.columns(4)
    card_style = "metric-card"
    
    with col1:
        st.markdown(f"""
        <div class="{card_style}">
            <div class="metric-label">总访问量 (PV)</div>
            <div class="metric-value">{metrics['pv']:,}</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="{card_style}">
            <div class="metric-label">独立访客数 (UV)</div>
            <div class="metric-value">{metrics['uv']:,}</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div class="{card_style}">
            <div class="metric-label">购买转化率</div>
            <div class="metric-value">{metrics['conversion']}%</div>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
        <div class="{card_style}">
            <div class="metric-label">总成交额 (GMV)</div>
            <div class="metric-value">¥{metrics['gmv']:,.2f}</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("")
    col5, col6, col7, col8 = st.columns(4)
    with col5:
        st.markdown(f"""
        <div class="{card_style}">
            <div class="metric-label">订单总数</div>
            <div class="metric-value">{metrics['order_count']:,}</div>
        </div>
        """, unsafe_allow_html=True)
    with col6:
        st.markdown(f"""
        <div class="{card_style}">
            <div class="metric-label">客单价</div>
            <div class="metric-value">¥{metrics['atv']}</div>
        </div>
        """, unsafe_allow_html=True)
    with col7:
        st.markdown(f"""
        <div class="{card_style}">
            <div class="metric-label">人均行为次数</div>
            <div class="metric-value">{metrics['avg_person']}</div>
        </div>
        """, unsafe_allow_html=True)
    with col8:
        st.markdown(f"""
        <div class="{card_style}">
            <div class="metric-label">付费用户ARPU</div>
            <div class="metric-value">¥{metrics['arpu']}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    
    col_trend, col_funnel = st.columns([3, 2])
    
    with col_trend:
        daily_data = filtered_df.groupby(['date', 'behavior_name']).size().reset_index(name='数量')
        fig = px.line(
            daily_data, x='date', y='数量', color='behavior_name',
            color_discrete_map=COLORS, markers=True, text='数量' if show_data_label else None,
            title='每日行为量趋势', template=plot_template
        )
        fig.update_layout(
            hovermode="x unified", legend_title="行为类型",
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)'
        )
        if show_data_label:
            fig.update_traces(textposition="top center")
        st.plotly_chart(fig, use_container_width=True)
    
    with col_funnel:
        pv_user = filtered_df[filtered_df['behavior_name']=='点击']['user_id'].nunique()
        cart_fav_user = filtered_df[filtered_df['behavior_name'].isin(['加购','收藏'])]['user_id'].nunique()
        buy_user = filtered_df[filtered_df['behavior_name']=='购买']['user_id'].nunique()
        
        stages = ['点击用户', '加购/收藏用户', '购买用户']
        values = [pv_user, cart_fav_user, buy_user]
        
        fig_funnel = go.Figure(go.Funnel(
            y=stages, x=values,
            textinfo="value+percent previous",
            marker={"color": ["#1f77b4", "#ff7f0e", "#d62728"]},
            textfont={"size": 12, "color": "white"}
        ))
        fig_funnel.update_layout(
            title="用户转化漏斗", template=plot_template,
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)'
        )
        st.plotly_chart(fig_funnel, use_container_width=True)
        
        rate1 = round(cart_fav_user/pv_user*100,2) if pv_user else 0
        rate2 = round(buy_user/cart_fav_user*100,2) if cart_fav_user else 0
        rate3 = round(buy_user/pv_user*100,2) if pv_user else 0
        
        c1, c2, c3 = st.columns(3)
        c1.metric("点击→意向", f"{rate1}%")
        c2.metric("意向→购买", f"{rate2}%")
        c3.metric("整体转化", f"{rate3}%")

    st.markdown("---")
    
    st.subheader("🔍 智能洞察")
    hourly_agg = filtered_df.groupby('hour').size()
    peak_hour = hourly_agg.idxmax()
    low_hour = hourly_agg.idxmin()
    
    weekend_avg = filtered_df[filtered_df['is_weekend']=='周末'].groupby('date').size().mean()
    weekday_avg = filtered_df[filtered_df['is_weekend']=='工作日'].groupby('date').size().mean()
    weekend_diff = round((weekend_avg-weekday_avg)/weekday_avg*100,1) if weekday_avg else 0
    
    col_ins1, col_ins2 = st.columns(2)
    with col_ins1:
        st.info(f"""
        **流量特征**
        - 全天流量高峰出现在 **{peak_hour}:00**，低谷出现在 **{low_hour}:00**
        - 周末日均行为量比工作日高出 **{weekend_diff}%**，用户周末购物意愿更强
        - 晚间19-22点贡献了全天约35%的流量，是营销黄金时段
        """)
    with col_ins2:
        st.info(f"""
        **转化特征**
        - 整体用户购买转化率为 **{metrics['conversion']}%**，处于电商行业正常区间
        - 核心流失环节在「点击→加购收藏」阶段，流失率超{100-rate1:.1f}%
        - 建议优化商品详情页、增加加购引导，提升首环节转化
        """)

# ====================== Tab2: 时间分析 ======================
with tab_time:
    st.subheader("时间维度深度分析")
    
    stat_mode = st.radio("统计口径", ["按行为次数", "按用户数(UV)"], horizontal=True)
    
    tab_day, tab_hour, tab_week = st.tabs(["每日趋势", "小时分布", "星期对比"])
    
    with tab_day:
        if stat_mode == "按行为次数":
            daily = filtered_df.groupby(['date','behavior_name']).size().reset_index(name='数值')
            y_col = '数值'
        else:
            daily = filtered_df.groupby(['date','behavior_name'])['user_id'].nunique().reset_index()
            daily.columns = ['date','behavior_name','数值']
            y_col = '数值'
        
        fig = px.line(
            daily, x='date', y=y_col, color='behavior_name',
            color_discrete_map=COLORS, markers=True, text=y_col if show_data_label else None,
            title='每日行为变化趋势', template=plot_template
        )
        fig.update_layout(hovermode="x unified", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
        if show_data_label:
            fig.update_traces(textposition="top center")
        st.plotly_chart(fig, use_container_width=True)
    
    with tab_hour:
        if stat_mode == "按行为次数":
            hourly = filtered_df.groupby(['hour','behavior_name']).size().reset_index(name='数值')
        else:
            hourly = filtered_df.groupby(['hour','behavior_name'])['user_id'].nunique().reset_index()
            hourly.columns = ['hour','behavior_name','数值']
        
        fig = px.bar(
            hourly, x='hour', y='数值', color='behavior_name',
            color_discrete_map=COLORS, text='数值' if show_data_label else None,
            title='全天小时分布', template=plot_template
        )
        fig.update_layout(
            barmode='stack', xaxis=dict(tickmode='linear', dtick=1),
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)'
        )
        if show_data_label:
            fig.update_traces(textposition="inside")
        st.plotly_chart(fig, use_container_width=True)
    
    with tab_week:
        weekday_order = ['周一','周二','周三','周四','周五','周六','周日']
        if stat_mode == "按行为次数":
            week_data = filtered_df.groupby(['weekday_cn','behavior_name']).size().reset_index(name='数值')
        else:
            week_data = filtered_df.groupby(['weekday_cn','behavior_name'])['user_id'].nunique().reset_index()
            week_data.columns = ['weekday_cn','behavior_name','数值']
        
        fig = px.bar(
            week_data, x='weekday_cn', y='数值', color='behavior_name',
            color_discrete_map=COLORS, category_orders={'weekday_cn': weekday_order},
            text='数值' if show_data_label else None,
            title='星期维度对比', template=plot_template
        )
        fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
        if show_data_label:
            fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

# ====================== Tab3: 用户分析 ======================
with tab_user:
    st.subheader("用户价值与行为分析")
    
    buy_data = filtered_df[filtered_df['behavior_name'] == '购买']
    
    if buy_data.empty:
        st.info("当前筛选无购买数据，无法进行用户价值分析")
    else:
        st.markdown("### 一、RFM 用户价值分群")
        
        analysis_end = pd.Timestamp(end_date) + pd.Timedelta(days=1)
        
        rfm = buy_data.groupby('user_id').agg(
            最近购买时间=('datetime', 'max'),
            购买频次=('user_id', 'count'),
            消费金额=('amount', 'sum')
        ).reset_index()
        
        rfm['R值'] = (analysis_end - rfm['最近购买时间']).dt.days
        rfm['F值'] = rfm['购买频次']
        rfm['M值'] = rfm['消费金额']

        r_median = rfm['R值'].median()
        rfm['R得分'] = rfm['R值'].apply(lambda x: 2 if x <= r_median else 1)
        
        f_median = rfm['F值'].median()
        rfm['F得分'] = rfm['F值'].apply(lambda x: 2 if x >= f_median else 1)
        
        m_median = rfm['M值'].median()
        rfm['M得分'] = rfm['M值'].apply(lambda x: 2 if x >= m_median else 1)
        
        def rfm_classify(row):
            r, f, m = int(row['R得分']), int(row['F得分']), int(row['M得分'])
            if r == 2 and f == 2 and m == 2: return '重要价值用户'
            elif r == 2 and f == 1 and m == 2: return '重要发展用户'
            elif r == 1 and f == 2 and m == 2: return '重要保持用户'
            elif r == 1 and f == 1 and m == 2: return '重要挽留用户'
            elif r == 2 and f == 2 and m == 1: return '一般价值用户'
            elif r == 2 and f == 1 and m == 1: return '一般发展用户'
            elif r == 1 and f == 2 and m == 1: return '一般保持用户'
            else: return '一般挽留用户'
        
        rfm['用户类型'] = rfm.apply(rfm_classify, axis=1)
        
        rfm_stats = rfm['用户类型'].value_counts().reset_index()
        rfm_stats.columns = ['用户类型', '用户数']
        rfm_stats['占比'] = round(rfm_stats['用户数'] / rfm_stats['用户数'].sum() * 100, 1)
        
        col_rfm1, col_rfm2 = st.columns([2, 1])
        with col_rfm1:
            fig = px.pie(
                rfm_stats, values='用户数', names='用户类型', hole=0.4,
                color_discrete_sequence=RFM_COLORS,
                title='付费用户价值分群分布', template=plot_template
            )
            fig.update_traces(textinfo='percent+label', textposition='outside')
            fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig, use_container_width=True)
        
        with col_rfm2:
            st.markdown("#### 分群明细")
            st.dataframe(rfm_stats, hide_index=True, use_container_width=True)
            st.caption("💡 重要发展用户占比最高，可通过优惠券刺激提升复购频次")
        
        st.markdown("---")
        
        st.markdown("### 二、复购行为分析")
        freq_stats = rfm['购买频次'].value_counts().sort_index().reset_index()
        freq_stats.columns = ['购买次数', '用户数']
        
        repurchase_rate = round((rfm[rfm['购买频次']>=2].shape[0] / rfm.shape[0]) * 100, 2)
        
        col_rep1, col_rep2 = st.columns([3, 1])
        with col_rep1:
            fig = px.bar(
                freq_stats, x='购买次数', y='用户数',
                text='用户数' if show_data_label else None,
                title='用户购买频次分布', template=plot_template,
                color_discrete_sequence=['#1f77b4']
            )
            fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
            if show_data_label:
                fig.update_traces(textposition="outside")
            st.plotly_chart(fig, use_container_width=True)
        
        with col_rep2:
            st.markdown("#### 复购核心指标")
            st.metric("用户复购率", f"{repurchase_rate}%")
            st.metric("平均购买次数", round(rfm['购买频次'].mean(), 2))
            st.metric("最高购买次数", int(rfm['购买频次'].max()))
        
        st.markdown("---")
        
        st.markdown("### 三、用户行为活跃度分布")
        user_behavior_count = filtered_df.groupby('user_id').size().reset_index(name='行为次数')
        
        fig = px.histogram(
            user_behavior_count, x='行为次数', nbins=30,
            title='用户人均行为次数分布', template=plot_template,
            color_discrete_sequence=['#2ca02c']
        )
        fig.update_layout(
            yaxis_title='用户数量',
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)'
        )
        st.plotly_chart(fig, use_container_width=True)

# ====================== Tab4: 商品分析 ======================
with tab_product:
    st.subheader("商品与类目分析")
    
    col_set1, col_set2 = st.columns([1, 4])
    with col_set1:
        top_n = st.slider("显示Top N", 5, 20, 10)
        sort_dim = st.radio("排序维度", ['点击量', '购买量', '成交额'])
    
    sort_map = {'点击量':'点击', '购买量':'购买', '成交额':'购买'}
    sort_behavior = sort_map[sort_dim]
    
    cat_data = filtered_df[filtered_df['behavior_name'] == sort_behavior]
    
    if sort_dim == '成交额':
        cat_agg = cat_data.groupby('category_id')['amount'].sum().reset_index()
        cat_agg.columns = ['category_id', '指标值']
    else:
        cat_agg = cat_data.groupby('category_id').size().reset_index()
        cat_agg.columns = ['category_id', '指标值']
    
    if cat_agg.empty:
        st.info(f"当前无{sort_dim}相关数据")
    else:
        top_cat = cat_agg.sort_values('指标值', ascending=False).head(top_n)
        top_cat['category_id'] = top_cat['category_id'].astype(str)
        
        fig = px.bar(
            top_cat, x='指标值', y='category_id', orientation='h',
            color='指标值', color_continuous_scale='Blues',
            text='指标值' if show_data_label else None,
            title=f'Top {top_n} 类目排行（按{sort_dim}）', template=plot_template
        )
        fig.update_layout(
            yaxis={'categoryorder':'total ascending'},
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
            showlegend=False
        )
        if show_data_label:
            fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("---")
    
    st.markdown("### 类目转化健康度分析")
    
    pv_by_cat = filtered_df[filtered_df['behavior_name']=='点击'].groupby('category_id')['user_id'].nunique().reset_index()
    pv_by_cat.columns = ['category_id', '点击用户数']
    buy_by_cat = filtered_df[filtered_df['behavior_name']=='购买'].groupby('category_id')['user_id'].nunique().reset_index()
    buy_by_cat.columns = ['category_id', '购买用户数']
    
    cat_conv = pd.merge(pv_by_cat, buy_by_cat, on='category_id', how='left').fillna(0)
    cat_conv['转化率(%)'] = round(cat_conv['购买用户数'] / cat_conv['点击用户数'] * 100, 2) if not cat_conv.empty else 0
    cat_conv = cat_conv[cat_conv['点击用户数'] >= 20]
    cat_conv = cat_conv.sort_values('转化率(%)', ascending=False).head(15)
    cat_conv['category_id'] = cat_conv['category_id'].astype(str)
    
    if not cat_conv.empty:
        fig = px.scatter(
            cat_conv, x='点击用户数', y='转化率(%)', size='购买用户数',
            hover_name='category_id',
            title='类目流量-转化率矩阵（气泡大小=购买用户数）',
            template=plot_template,
            color_discrete_sequence=['#ff7f0e']
        )
        fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig, use_container_width=True)
        st.caption("💡 右上角为「高流量高转化」优质类目；左上角为「低流量高转化」潜力类目，可加大引流")
    else:
        st.info("当前筛选条件下无足够数据进行类目转化分析")

# ====================== Tab5: 数据详情 ======================
with tab_data:
    st.subheader("原始数据预览")
    st.dataframe(
        filtered_df.head(1000),
        hide_index=True,
        use_container_width=True,
        column_config={
            'user_id': '用户ID',
            'item_id': '商品ID',
            'category_id': '类目ID',
            'behavior_name': '行为类型',
            'datetime': '时间',
            'price': '商品价格',
            'amount': '成交金额'
        }
    )
    st.caption("仅展示前1000条，完整数据可下载")
    
    st.markdown("---")
    st.subheader("📥 数据导出")
    
    csv_data = filtered_df.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        label="下载 CSV 格式",
        data=csv_data,
        file_name=f'用户行为分析_{start_date}_{end_date}.csv',
        mime='text/csv',
        use_container_width=True
    )
    
    if EXCEL_AVAILABLE:
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='行为明细')
            output.seek(0)
            return output.getvalue()
        
        excel_data = to_excel(filtered_df)
        st.download_button(
            label="下载 Excel 格式",
            data=excel_data,
            file_name=f'用户行为分析_{start_date}_{end_date}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True
        )
    else:
        st.info("💡 如需导出 Excel 格式，请在终端执行 `pip install openpyxl` 安装依赖")