import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from io import BytesIO
from datetime import date

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

# ====================== 全局样式与主题系统 ======================
def apply_theme(theme_mode):
    if theme_mode == '深色主题':
        bg_color = "#0e1117"
        card_bg = "#1a1d24"
        text_color = "#fafafa"
        sub_text = "#b0b0b0"
        border_color = "#2d303a"
        plot_template = "plotly_dark"
    else:
        bg_color = "#f5f7fa"
        card_bg = "#ffffff"
        text_color = "#000000"
        sub_text = "#222222"
        border_color = "#d0d7de"
        plot_template = "plotly_white"
    
    st.markdown(f"""
    <style>
    .stApp, .stApp * {{ color: {text_color} !important; }}
    .stApp {{ background-color: {bg_color}; }}
    [data-testid="stSidebar"], [data-testid="stSidebar"] * {{
        background-color: {card_bg} !important;
        color: {text_color} !important;
    }}
    [data-testid="stSidebar"] {{ border-right: 1px solid {border_color}; }}
    h1, h2, h3, h4, h5, h6 {{ color: {text_color} !important; }}
    .metric-card {{
        background-color: {card_bg} !important;
        padding: 20px 24px;
        border-radius: 12px;
        border: 1px solid {border_color};
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }}
    .metric-label {{ font-size: 14px; color: {sub_text} !important; }}
    .metric-value {{ font-size: 28px; font-weight: 600; color: {text_color} !important; }}
    hr {{ border-color: {border_color} !important; }}
    .stAlert * {{ color: #000000 !important; }}
    </style>
    """, unsafe_allow_html=True)
    return plot_template

# ====================== 统一配色 ======================
COLORS = {
    '点击': '#1f77b4',
    '加购': '#ff7f0e',
    '收藏': '#2ca02c',
    '购买': '#d62728'
}
RFM_COLORS = ['#2ecc71','#3498db','#f1c40f','#e74c3c','#95a5a6','#1abc9c','#e67e22','#9b59b6']

# ====================== 生成模拟数据（终极修复：彻底避开freq报错） ======================
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
    behavior_types = ['pv','cart','fav','buy']
    behavior_probs = [0.92,0.048,0.022,0.01]

    # 🔥 终极修复：完全不用 pd.date_range + freq，彻底解决报错
    start_ts = pd.Timestamp('2017-11-25 00:00:00').value // 10**9
    end_ts = pd.Timestamp('2017-12-03 23:59:59').value // 10**9
    timestamp_unix = np.random.randint(start_ts, end_ts, size=n_rows)

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
        'user_id':user_ids,'item_id':item_ids,'category_id':category_ids,
        'behavior_type':behavior_type,'timestamp':timestamp_unix
    })

    buy_mask = df['behavior_type']=='buy'
    buy_users = df[buy_mask]['user_id'].unique()
    high_freq_buy_users = np.random.choice(buy_users, size=int(len(buy_users)*0.18), replace=False)
    extra_buy = []
    for uid in high_freq_buy_users:
        extra_num = np.random.randint(1,4)
        for _ in range(extra_num):
            extra_buy.append({'user_id':uid,'item_id':np.random.choice(item_pool),
                             'category_id':np.random.choice(category_pool),'behavior_type':'buy',
                             'timestamp':np.random.randint(start_ts, end_ts)})
    if extra_buy: df = pd.concat([df,pd.DataFrame(extra_buy)], ignore_index=True)

    pv_user_set = set(df[df['behavior_type']=='pv']['user_id'].unique())
    for btype in ['buy','cart','fav']:
        target_users = df[df['behavior_type']==btype]['user_id'].unique()
        missing = [uid for uid in target_users if uid not in pv_user_set]
        extra_rows = []
        for uid in missing:
            extra_rows.append({'user_id':uid,'item_id':np.random.choice(item_pool),
                             'category_id':np.random.choice(category_pool),'behavior_type':'pv',
                             'timestamp':start_ts})
            pv_user_set.add(uid)
        if extra_rows: df = pd.concat([df,pd.DataFrame(extra_rows)], ignore_index=True)

    df['datetime'] = pd.to_datetime(df['timestamp'], unit='s')
    df['date'] = df['datetime'].dt.date
    df['hour'] = df['datetime'].dt.hour
    df['weekday'] = df['datetime'].dt.weekday
    df['weekday_cn'] = df['weekday'].map({0:'周一',1:'周二',2:'周三',3:'周四',4:'周五',5:'周六',6:'周日'})
    df['is_weekend'] = df['weekday'].apply(lambda x: '周末' if x>=5 else '工作日')
    behavior_map = {'pv':'点击','cart':'加购','fav':'收藏','buy':'购买'}
    df['behavior_name'] = df['behavior_type'].map(behavior_map)
    df['price'] = df['item_id'].map(item_price_map)
    df['amount'] = np.where(df['behavior_type']=='buy', df['price'], 0)
    return df

df = load_data()

# ====================== 侧边栏 ======================
st.sidebar.title("📌 控制面板")

with st.sidebar.expander("⏰ 时间范围筛选", expanded=True):
    start_date, end_date = st.date_input(
        "选择分析时间段",
        value=[date(2017,11,25), date(2017,12,3)],
        min_value=date(2017,11,25),
        max_value=date(2017,12,3)
    )

with st.sidebar.expander("🛒 行为类型筛选", expanded=True):
    behavior_list = st.multiselect(
        "选择要展示的行为",
        options=['点击','加购','收藏','购买'],
        default=['点击','加购','收藏','购买']
    )

with st.sidebar.expander("⚙️ 显示设置", expanded=True):
    show_data_label = st.checkbox("显示图表数据标签", value=True)
    theme_mode = st.selectbox("界面主题", options=['深色主题','浅色主题'], index=0)

plot_template = apply_theme(theme_mode)

st.sidebar.markdown("---")
if st.sidebar.button("🔄 重置所有筛选", use_container_width=True):
    st.rerun()

st.sidebar.caption("数据说明：模拟淘宝用户行为数据\n周期：2017.11.25 - 2017.12.03")

# ====================== 数据筛选（无任何错误） ======================
if not behavior_list:
    st.warning("⚠️ 请至少选择一种行为类型进行分析")
    st.stop()

start_dt = pd.Timestamp(start_date)
end_dt = pd.Timestamp(end_date) + pd.Timedelta(days=1)
filtered_df = df[
    (df['datetime'] >= start_dt) &
    (df['datetime'] < end_dt) &
    (df['behavior_name'].isin(behavior_list))
]

if filtered_df.empty:
    st.warning("⚠️ 当前筛选条件下无数据，请调整时间范围或行为类型")
    st.stop()

# ====================== 核心指标 ======================
def calc_core_metrics(data):
    total_pv = data[data['behavior_name']=='点击'].shape[0]
    total_uv = data['user_id'].nunique()
    pv_uv = data[data['behavior_name']=='点击']['user_id'].nunique()
    buy_uv = data[data['behavior_name']=='购买']['user_id'].nunique()
    total_gmv = data['amount'].sum()
    order_count = data[data['behavior_name']=='购买'].shape[0]
    conversion = round(buy_uv/pv_uv*100,2) if pv_uv>0 else 0
    avg_person = round(data.shape[0]/total_uv,2) if total_uv>0 else 0
    atv = round(total_gmv/order_count,2) if order_count>0 else 0
    arpu = round(total_gmv/buy_uv,2) if buy_uv>0 else 0
    return {'pv':total_pv,'uv':total_uv,'conversion':conversion,'gmv':total_gmv,
            'order_count':order_count,'avg_person':avg_person,'atv':atv,'arpu':arpu}

# ====================== 主页面 ======================
st.title("📊 淘宝用户行为可视化分析系统")
st.markdown("---")

tab_overview, tab_time, tab_user, tab_product, tab_data = st.tabs([
    "📈 核心概览","⏰ 时间分析","👤 用户分析","🛍️ 商品分析","📋 数据详情"
])

# 核心概览
with tab_overview:
    metrics = calc_core_metrics(filtered_df)
    col1,col2,col3,col4 = st.columns(4)
    card = "metric-card"
    with col1: st.markdown(f"<div class='{card}'><div class='metric-label'>总访问量(PV)</div><div class='metric-value'>{metrics['pv']:,}</div></div>", unsafe_allow_html=True)
    with col2: st.markdown(f"<div class='{card}'><div class='metric-label'>独立访客(UV)</div><div class='metric-value'>{metrics['uv']:,}</div></div>", unsafe_allow_html=True)
    with col3: st.markdown(f"<div class='{card}'><div class='metric-label'>购买转化率</div><div class='metric-value'>{metrics['conversion']}%</div></div>", unsafe_allow_html=True)
    with col4: st.markdown(f"<div class='{card}'><div class='metric-label'>总成交额(GMV)</div><div class='metric-value'>¥{metrics['gmv']:,.2f}</div></div>", unsafe_allow_html=True)
    
    col5,col6,col7,col8 = st.columns(4)
    with col5: st.markdown(f"<div class='{card}'><div class='metric-label'>订单总数</div><div class='metric-value'>{metrics['order_count']:,}</div></div>", unsafe_allow_html=True)
    with col6: st.markdown(f"<div class='{card}'><div class='metric-label'>客单价</div><div class='metric-value'>¥{metrics['atv']}</div></div>", unsafe_allow_html=True)
    with col7: st.markdown(f"<div class='{card}'><div class='metric-label'>人均行为次数</div><div class='metric-value'>{metrics['avg_person']}</div></div>", unsafe_allow_html=True)
    with col8: st.markdown(f"<div class='{card}'><div class='metric-label'>付费用户ARPU</div><div class='metric-value'>¥{metrics['arpu']}</div></div>", unsafe_allow_html=True)

    st.markdown("---")
    col_a, col_b = st.columns([3,2])
    with col_a:
        daily = filtered_df.groupby(['date','behavior_name']).size().reset_index(name='数量')
        fig = px.line(daily, x='date', y='数量', color='behavior_name', color_discrete_map=COLORS, title='每日行为趋势', template=plot_template)
        st.plotly_chart(fig, use_container_width=True)
    with col_b:
        pv_u = filtered_df[filtered_df['behavior_name']=='点击']['user_id'].nunique()
        in_u = filtered_df[filtered_df['behavior_name'].isin(['加购','收藏'])]['user_id'].nunique()
        buy_u = filtered_df[filtered_df['behavior_name']=='购买']['user_id'].nunique()
        fig = go.Figure(go.Funnel(y=['点击','加购/收藏','购买'], x=[pv_u,in_u,buy_u]))
        fig.update_layout(title='转化漏斗', template=plot_template)
        st.plotly_chart(fig, use_container_width=True)

# 时间分析
with tab_time:
    st.subheader("时间维度分析")
    m = st.radio("统计口径", ["按行为次数","按用户数"], horizontal=True)
    t1,t2,t3 = st.tabs(["每日","小时","星期"])
    with t1:
        if m=="按行为次数": d = filtered_df.groupby(['date','behavior_name']).size().reset_index(name='数值')
        else: d = filtered_df.groupby(['date','behavior_name'])['user_id'].nunique().reset_index(name='数值')
        fig = px.line(d, x='date', y='数值', color='behavior_name', color_discrete_map=COLORS, template=plot_template)
        st.plotly_chart(fig, use_container_width=True)
    with t2:
        if m=="按行为次数": d = filtered_df.groupby(['hour','behavior_name']).size().reset_index(name='数值')
        else: d = filtered_df.groupby(['hour','behavior_name'])['user_id'].nunique().reset_index(name='数值')
        fig = px.bar(d, x='hour', y='数值', color='behavior_name', color_discrete_map=COLORS, template=plot_template)
        st.plotly_chart(fig, use_container_width=True)
    with t3:
        order = ['周一','周二','周三','周四','周五','周六','周日']
        if m=="按行为次数": d = filtered_df.groupby(['weekday_cn','behavior_name']).size().reset_index(name='数值')
        else: d = filtered_df.groupby(['weekday_cn','behavior_name'])['user_id'].nunique().reset_index(name='数值')
        fig = px.bar(d, x='weekday_cn', y='数值', color='behavior_name', category_orders={'weekday_cn':order}, color_discrete_map=COLORS, template=plot_template)
        st.plotly_chart(fig, use_container_width=True)

# 用户分析
with tab_user:
    st.subheader("用户RFM分析")
    buy = filtered_df[filtered_df['behavior_name']=='购买']
    if buy.empty:
        st.info("无购买数据")
    else:
        rfm = buy.groupby('user_id').agg(
            R=('datetime', lambda x: (pd.Timestamp.now()-x.max()).days),
            F=('user_id','count'), M=('amount','sum')
        ).reset_index()
        fig = px.pie(rfm, values='F', names=pd.cut(rfm['F'], bins=3, labels=['低','中','高']), title='购买频次分布', template=plot_template)
        st.plotly_chart(fig, use_container_width=True)

# 商品分析
with tab_product:
    st.subheader("类目分析")
    top = filtered_df.groupby('category_id')['behavior_name'].count().sort_values(ascending=False).head(10)
    fig = px.bar(top, orientation='h', title='Top10类目', template=plot_template)
    st.plotly_chart(fig, use_container_width=True)

# 数据详情
with tab_data:
    st.subheader("数据预览")
    st.dataframe(filtered_df.head(1000), use_container_width=True)
    csv = filtered_df.to_csv(index=False).encode('utf-8-sig')
    st.download_button("下载CSV", csv, "数据.csv", use_container_width=True)