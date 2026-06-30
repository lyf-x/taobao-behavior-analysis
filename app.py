# ====================== 数据过滤与兜底 ======================
if not behavior_list:
    st.warning("⚠️ 请至少选择一种行为类型进行分析")
    st.stop()

# 修正时间筛选逻辑：精准匹配datetime区间，包含end_date全天
start_dt = pd.Timestamp(start_date)
end_dt = pd.Timestamp(end_date) + pd.Timedelta(days=1)  # 加1天，用<判断实现包含end_date全天
filtered_df = df[
    (df['datetime'] >= start_dt) &
    (df['datetime'] < end_dt) &
    (df['behavior_name'].isin(behavior_list))
]

if filtered_df.empty:
    st.warning("⚠️ 当前筛选条件下无数据，请调整时间范围或行为类型")
    st.stop()