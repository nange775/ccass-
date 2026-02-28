import pandas as pd

# 2026年香港公众假期（包含周末的节日及平日的补假）
hk_holidays_2026 = [
    "2026-01-01",  # 元旦
    "2026-02-17", "2026-02-18", "2026-02-19",  # 春节初一至初三
    "2026-04-03", "2026-04-04",  # 耶稣受难节
    "2026-04-06",  # 清明节补假
    "2026-04-07",  # 复活节补假
    "2026-05-01",  # 劳动节
    "2026-05-25",  # 佛诞补假
    "2026-06-19",  # 端午节
    "2026-07-01",  # 香港特区成立纪念日
    "2026-09-26",  # 中秋节翌日
    "2026-10-01",  # 国庆日
    "2026-10-19",  # 重阳节补假
    "2026-12-25",  # 圣诞节
    "2026-12-26"   # 圣诞节后第一个平日
]

# 使用 CustomBusinessDay，自动排除周六、周日以及 holidays 列表
custom_bday = pd.offsets.CustomBusinessDay(holidays=hk_holidays_2026)

# 生成2026年全年的有效工作日
workdays_series = pd.bdate_range(
    start="2026-01-01",
    end="2026-12-31",
    freq=custom_bday
)

# 转换为日期字符串列表
workdays_list = workdays_series.strftime('%Y-%m-%d').tolist()


def get_real_date(query_date_str):
    """
    根据查询日期获取真实日期（两个工作日之前）
    
    Args:
        query_date_str: 查询日期，格式为 "2026/02/26" 或 "2026-02-26"
    
    Returns:
        真实日期字符串，格式为 "2026年02月24日"
    """
    # 统一日期格式
    query_date_str = query_date_str.replace('/', '-')
    
    # 检查日期是否在工作日列表中
    if query_date_str not in workdays_list:
        return None
    
    # 获取该日期在工作日列表中的索引
    idx = workdays_list.index(query_date_str)
    
    # 获取两个工作日之前的日期
    if idx >= 2:
        real_date_str = workdays_list[idx - 2]
    else:
        # 如果不足两个工作日，返回最早的日期
        real_date_str = workdays_list[0]
    
    # 转换为中文格式
    date_obj = pd.to_datetime(real_date_str)
    return date_obj.strftime('%Y年%m月%d日')


def get_real_date_from_sheet_name(sheet_name):
    """
    从sheet名称（如"2026年02月26日"）获取真实日期
    
    Args:
        sheet_name: sheet名称，格式为 "2026年02月26日持股量"
    
    Returns:
        真实日期字符串，格式为 "2026年02月24日"
    """
    import re
    
    # 提取日期部分
    match = re.search(r'(\d{4})年(\d{2})月(\d{2})日', sheet_name)
    if not match:
        return None
    
    year, month, day = match.groups()
    query_date_str = f"{year}-{month}-{day}"
    
    return get_real_date(query_date_str)


if __name__ == "__main__":
    # 测试
    test_dates = ["2026/02/26", "2026/01/20", "2026/03/02"]
    for d in test_dates:
        real = get_real_date(d)
        print(f"查询日期: {d} -> 真实日期: {real}")
