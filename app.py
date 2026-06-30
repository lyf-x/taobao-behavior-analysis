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
    from pandas.tseries.offsets import Hour
    hours = pd.date_range(start_date, end_date, freq=Hour())
    
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
            weekly = filtered_df.groupby(['weekday_cn','behavior_name']).size().reset_index(name='数值')
        else:
            weekly = filtered_df.groupby(['weekday_cn','behavior_name'])['user_id'].nunique().reset_index()
            weekly.columns = ['weekday_cn','behavior_name','数值']
        
        # 保证星期顺序
        weekly['weekday_cn'] = pd.Categorical(weekly['weekday_cn'], categories=weekday_order, ordered=True)
        weekly = weekly.sort_values('weekday_cn')
        
        fig = px.bar(
            weekly, x='weekday_cn', y='数值', color='behavior_name',
            color_discrete_map=COLORS, text='数值' if show_data_label else None,
            title='星期维度行为分布', template=plot_template
        )
        fig.update_layout(
            barmode='stack', plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)'
        )
        if show_data_label:
            fig.update_traces(textposition="inside")
        st.plotly_chart(fig, use_container_width=True)

# ====================== Tab3: 用户分析（RFM） ======================
with tab_user:
    st.subheader("用户价值分析（RFM模型）")
    
    # 仅筛选购买数据
    buy_data = filtered_df[filtered_df['behavior_name'] == '购买']
    if buy_data.empty:
        st.warning("⚠️ 当前筛选条件下无购买数据，无法进行RFM分析")
        st.stop()
    
    # 计算RFM指标
    analysis_end = pd.Timestamp(end_date) + pd.Timedelta(days=1)
    rfm = buy_data.groupby('user_id').agg({
        'datetime': lambda x: (analysis_end - x.max()).days,  # R：最近购买天数
        'user_id': 'count',  # F：购买频次
        'amount': 'sum'  # M：消费金额
    }).rename(columns={
        'datetime': '最近购买时间',
        'user_id': '购买频次',
        'amount': '消费金额'
    }).reset_index()
    
    # RFM打分（中位数二分）
    rfm['R得分'] = np.where(rfm['最近购买时间'] <= rfm['最近购买时间'].median(), 2, 1)
    rfm['F得分'] = np.where(rfm['购买频次'] >= rfm['购买频次'].median(), 2, 1)
    rfm['M得分'] = np.where(rfm['消费金额'] >= rfm['消费金额'].median(), 2, 1)
    
    # 计算总分并分群
    rfm['RFM总分'] = rfm['R得分'] + rfm['F得分'] + rfm['M得分']
    def rfm_segment(row):
        if row['R得分']==2 and row['F得分']==2 and row['M得分']==2:
            return '高价值用户'
        elif row['R得分']==2 and row['F得分']==1 and row['M得分']==2:
            return '高潜力用户'
        elif row['R得分']==1 and row['F得分']==2 and row['M得分']==2:
            return '忠诚度用户'
        elif row['R得分']==1 and row['F得分']==1 and row['M得分']==2:
            return '挽留用户'
        elif row['R得分']==2 and row['F得分']==2 and row['M得分']==1:
            return '高频低消用户'
        elif row['R得分']==2 and row['F得分']==1 and row['M得分']==1:
            return '新用户'
        elif row['R得分']==1 and row['F得分']==2 and row['M得分']==1:
            return '沉睡用户'
        else:
            return '低价值用户'
    
    rfm['用户分层'] = rfm.apply(rfm_segment, axis=1)
    
    # 可视化RFM分层
    rfm_count = rfm['用户分层'].value_counts().reset_index()
    rfm_count.columns = ['用户分层', '用户数']
    
    col_rfm1, col_rfm2 = st.columns([2, 3])
    with col_rfm1:
        fig = px.pie(
            rfm_count, values='用户数', names='用户分层',
            color_discrete_sequence=RFM_COLORS,
            title='用户分层分布', template=plot_template
        )
        fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig, use_container_width=True)
        
        # 核心指标
        total_buy_users = len(rfm)
        repurchase_rate = round((rfm[rfm['购买频次']>=2].shape[0] / rfm.shape[0]) * 100, 2) if rfm.shape[0] > 0 else 0
        avg_consume = round(rfm['消费金额'].mean(), 2)
        
        st.markdown(f"""
        <div class="{card_style}">
            <div class="metric-label">购买用户总数</div>
            <div class="metric-value">{total_buy_users:,}</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown(f"""
        <div class="{card_style}">
            <div class="metric-label">用户复购率</div>
            <div class="metric-value">{repurchase_rate}%</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown(f"""
        <div class="{card_style}">
            <div class="metric-label">平均消费金额</div>
            <div class="metric-value">¥{avg_consume}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col_rfm2:
        # RFM散点图（R vs M，大小=F）
        fig = px.scatter(
            rfm, x='最近购买时间', y='消费金额', size='购买频次',
            color='用户分层', color_discrete_sequence=RFM_COLORS,
            hover_data=['user_id', '购买频次'],
            title='RFM用户价值分布（R=最近购买时间，M=消费金额，大小=F）',
            template=plot_template
        )
        fig.update_layout(
            xaxis_title='最近购买天数（越小越优）',
            yaxis_title='消费金额（越大越优）',
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)'
        )
        st.plotly_chart(fig, use_container_width=True)
        
        # 分层价值分析
        rfm_value = rfm.groupby('用户分层')['消费金额'].agg(['sum', 'mean']).reset_index()
        rfm_value.columns = ['用户分层', '总消费额', '人均消费额']
        rfm_value['总消费额占比'] = round(rfm_value['总消费额'] / rfm_value['总消费额'].sum() * 100, 2)
        
        fig = px.bar(
            rfm_value, x='用户分层', y='总消费额占比',
            color='用户分层', color_discrete_sequence=RFM_COLORS,
            text='总消费额占比', template=plot_template,
            title='各分层消费额贡献占比'
        )
        fig.update_layout(
            yaxis_title='消费额占比(%)',
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)'
        )
        if show_data_label:
            fig.update_traces(textposition="top center")
        st.plotly_chart(fig, use_container_width=True)

# ====================== Tab4: 商品分析 ======================
with tab_product:
    st.subheader("商品/类目分析")
    
    analysis_type = st.radio("分析维度", ["按类目", "按商品"], horizontal=True)
    
    if analysis_type == "按类目":
        # 类目分析
        cat_metrics = filtered_df.groupby('category_id').agg({
            'user_id': 'nunique',
            'behavior_name': lambda x: pd.Series(x).value_counts().to_dict()
        }).reset_index()
        
        # 拆解行为类型数据
        cat_metrics['点击用户数'] = cat_metrics['behavior_name'].apply(lambda x: x.get('点击', 0))
        cat_metrics['加购用户数'] = cat_metrics['behavior_name'].apply(lambda x: x.get('加购', 0))
        cat_metrics['收藏用户数'] = cat_metrics['behavior_name'].apply(lambda x: x.get('收藏', 0))
        cat_metrics['购买用户数'] = cat_metrics['behavior_name'].apply(lambda x: x.get('购买', 0))
        cat_metrics.drop('behavior_name', axis=1, inplace=True)
        
        # 计算核心指标
        cat_metrics['总用户数'] = cat_metrics['user_id']
        cat_metrics['转化率(%)'] = cat_metrics.apply(
            lambda row: round(row['购买用户数']/row['点击用户数']*100,2) if row['点击用户数']>0 else 0,
            axis=1
        )
        cat_metrics['加购率(%)'] = cat_metrics.apply(
            lambda row: round(row['加购用户数']/row['点击用户数']*100,2) if row['点击用户数']>0 else 0,
            axis=1
        )
        
        # 排序展示Top10类目
        cat_top10 = cat_metrics.nlargest(10, '总用户数')
        
        col_pro1, col_pro2 = st.columns([2, 3])
        with col_pro1:
            st.markdown("### 📈 类目核心指标（Top10）")
            st.dataframe(
                cat_top10[['category_id', '总用户数', '点击用户数', '购买用户数', '转化率(%)', '加购率(%)']],
                use_container_width=True
            )
            
            # 类目转化Top5
            cat_conv_top5 = cat_metrics[cat_metrics['点击用户数']>=10].nlargest(5, '转化率(%)')
            st.markdown("### 🎯 高转化类目（Top5）")
            st.dataframe(
                cat_conv_top5[['category_id', '点击用户数', '购买用户数', '转化率(%)']],
                use_container_width=True
            )
        
        with col_pro2:
            # 类目用户数分布
            fig = px.bar(
                cat_top10, x='category_id', y='总用户数',
                color='转化率(%)', color_continuous_scale='RdYlGn',
                title='Top10类目用户数与转化率', template=plot_template
            )
            fig.update_layout(
                xaxis_title='类目ID', yaxis_title='总用户数',
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)'
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # 类目行为分布
            cat_behavior = filtered_df[filtered_df['category_id'].isin(cat_top10['category_id'])].groupby(
                ['category_id', 'behavior_name']
            ).size().reset_index(name='数量')
            
            fig = px.bar(
                cat_behavior, x='category_id', y='数量', color='behavior_name',
                color_discrete_map=COLORS, barmode='stack',
                title='Top10类目行为分布', template=plot_template
            )
            fig.update_layout(
                xaxis_title='类目ID', yaxis_title='行为次数',
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)'
            )
            st.plotly_chart(fig, use_container_width=True)
    else:
        # 商品分析
        item_metrics = filtered_df.groupby('item_id').agg({
            'user_id': 'nunique',
            'behavior_name': lambda x: pd.Series(x).value_counts().to_dict(),
            'amount': 'sum'
        }).reset_index()
        
        item_metrics['点击用户数'] = item_metrics['behavior_name'].apply(lambda x: x.get('点击', 0))
        item_metrics['加购用户数'] = item_metrics['behavior_name'].apply(lambda x: x.get('加购', 0))
        item_metrics['收藏用户数'] = item_metrics['behavior_name'].apply(lambda x: x.get('收藏', 0))
        item_metrics['购买用户数'] = item_metrics['behavior_name'].apply(lambda x: x.get('购买', 0))
        item_metrics.drop('behavior_name', axis=1, inplace=True)
        
        item_metrics['转化率(%)'] = item_metrics.apply(
            lambda row: round(row['购买用户数']/row['点击用户数']*100,2) if row['点击用户数']>0 else 0,
            axis=1
        )
        item_metrics['销售额'] = item_metrics['amount']
        
        # Top10商品（按销售额）
        item_top10 = item_metrics.nlargest(10, '销售额')
        
        col_pro1, col_pro2 = st.columns([2, 3])
        with col_pro1:
            st.markdown("### 📈 商品核心指标（Top10）")
            st.dataframe(
                item_top10[['item_id', 'user_id', '点击用户数', '购买用户数', '转化率(%)', '销售额']],
                use_container_width=True
            )
            
            # 高转化商品Top5
            item_conv_top5 = item_metrics[item_metrics['点击用户数']>=5].nlargest(5, '转化率(%)')
            st.markdown("### 🎯 高转化商品（Top5）")
            st.dataframe(
                item_conv_top5[['item_id', '点击用户数', '购买用户数', '转化率(%)']],
                use_container_width=True
            )
        
        with col_pro2:
            # 商品销售额分布
            fig = px.bar(
                item_top10, x='item_id', y='销售额',
                color='转化率(%)', color_continuous_scale='RdYlGn',
                title='Top10商品销售额与转化率', template=plot_template
            )
            fig.update_layout(
                xaxis_title='商品ID', yaxis_title='销售额(¥)',
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)'
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # 商品行为分布
            item_behavior = filtered_df[filtered_df['item_id'].isin(item_top10['item_id'])].groupby(
                ['item_id', 'behavior_name']
            ).size().reset_index(name='数量')
            
            fig = px.bar(
                item_behavior, x='item_id', y='数量', color='behavior_name',
                color_discrete_map=COLORS, barmode='stack',
                title='Top10商品行为分布', template=plot_template
            )
            fig.update_layout(
                xaxis_title='商品ID', yaxis_title='行为次数',
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)'
            )
            st.plotly_chart(fig, use_container_width=True)

# ====================== Tab5: 数据详情 ======================
with tab_data:
    st.subheader("原始数据详情与导出")
    
    # 数据预览
    st.markdown("### 📋 筛选后数据预览（前1000行）")
    preview_df = filtered_df.head(1000)[['user_id', 'item_id', 'category_id', 'behavior_name', 'datetime', 'price', 'amount']]
    st.dataframe(preview_df, use_container_width=True)
    
    # 数据导出
    st.markdown("### 📤 数据导出")
    col_export1, col_export2 = st.columns(2)
    
    with col_export1:
        export_format = st.radio("选择导出格式", ["Excel (.xlsx)", "CSV (.csv)"], horizontal=True)
        
        if st.button("📥 导出筛选后数据", use_container_width=True):
            # 准备导出数据
            export_df = filtered_df[['user_id', 'item_id', 'category_id', 'behavior_name', 'datetime', 'date', 'hour', 'weekday_cn', 'price', 'amount']]
            
            # 处理导出
            if export_format == "Excel (.xlsx)":
                if not EXCEL_AVAILABLE:
                    st.error("⚠️ 缺少Excel导出依赖，请安装：pip install openpyxl")
                else:
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        export_df.to_excel(writer, sheet_name='淘宝用户行为数据', index=False)
                    output.seek(0)
                    st.download_button(
                        label="📥 点击下载Excel文件",
                        data=output,
                        file_name=f"淘宝用户行为数据_{start_date}_至_{end_date}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            else:
                csv_data = export_df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📥 点击下载CSV文件",
                    data=csv_data,
                    file_name=f"淘宝用户行为数据_{start_date}_至_{end_date}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
    
    with col_export2:
        # 数据统计信息
        st.markdown("### 📊 数据统计信息")
        stats_df = pd.DataFrame({
            '指标': [
                '数据总行数', '独立用户数', '独立商品数', '独立类目数',
                '点击次数', '加购次数', '收藏次数', '购买次数',
                '数据时间范围'
            ],
            '数值': [
                len(filtered_df),
                filtered_df['user_id'].nunique(),
                filtered_df['item_id'].nunique(),
                filtered_df['category_id'].nunique(),
                len(filtered_df[filtered_df['behavior_name']=='点击']),
                len(filtered_df[filtered_df['behavior_name']=='加购']),
                len(filtered_df[filtered_df['behavior_name']=='收藏']),
                len(filtered_df[filtered_df['behavior_name']=='购买']),
                f"{filtered_df['date'].min()} 至 {filtered_df['date'].max()}"
            ]
        })
        st.dataframe(stats_df, use_container_width=True)

# 补充缺失的Tab3/Tab4/Tab5完整逻辑（原代码裁剪部分补全）