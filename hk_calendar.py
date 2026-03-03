import pandas as pd

# 香港公众假期
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
hk_holidays_2025 = [
    "2025-01-01",  # 元旦
    "2025-01-29", "2025-01-30", "2025-01-31",  # 农历年初一至初三
    "2025-04-04",  # 清明节
    "2025-04-18", "2025-04-19",  # 耶稣受难节及翌日
    "2025-04-21",  # 复活节星期一
    "2025-05-01",  # 劳动节
    "2025-05-05",  # 佛诞
    "2025-05-31",  # 端午节
    "2025-07-01",  # 香港特别行政区成立纪念日
    "2025-10-01",  # 国庆日
    "2025-10-07",  # 中秋节翌日
    "2025-10-29",  # 重阳节
    "2025-12-25",  # 圣诞节
    "2025-12-26"   # 圣诞节后第一个平日 (节礼日)
]

hk_holidays_2024 = [
    "2024-01-01",  # 一月一日
    "2024-02-10", "2024-02-12", "2024-02-13",  # 农历年初一、初三、初四 (年初四为补假)
    "2024-03-29", "2024-03-30",  # 耶稣受难节及翌日
    "2024-04-01",  # 复活节星期一
    "2024-04-04",  # 清明节
    "2024-05-01",  # 劳动节
    "2024-05-15",  # 佛诞
    "2024-06-10",  # 端午节
    "2024-07-01",  # 香港特别行政区成立纪念日
    "2024-09-18",  # 中秋节翌日
    "2024-10-01",  # 国庆日
    "2024-10-11",  # 重阳节
    "2024-12-25",  # 圣诞节
    "2024-12-26"   # 圣诞节后第一个周日 (节礼日)
]
hk_holidays_2023 = [
    "2023-01-02",  # 一月一日翌日 (元旦补假)
    "2023-01-23", "2023-01-24", "2023-01-25",  # 农历年初二至初四
    "2023-04-05",  # 清明节
    "2023-04-07", "2023-04-08",  # 耶稣受难节及翌日
    "2023-04-10",  # 复活节星期一
    "2023-05-01",  # 劳动节
    "2023-05-26",  # 佛诞
    "2023-06-22",  # 端午节
    "2023-07-01",  # 香港特别行政区成立纪念日
    "2023-09-30",  # 中秋节翌日
    "2023-10-02",  # 国庆日翌日 (补假)
    "2023-10-23",  # 重阳节
    "2023-12-25",  # 圣诞节
    "2023-12-26"   # 圣诞节后第一个周日 (节礼日)
]
hk_holidays_2022 = [
    "2022-01-01",  # 元旦
    "2022-02-01", "2022-02-02", "2022-02-03",  # 农历年初一至初三
    "2022-04-05",  # 清明节
    "2022-04-15", "2022-04-16",  # 耶稣受难节及翌日
    "2022-04-18",  # 复活节星期一
    "2022-05-02",  # 劳动节翌日 (补假)
    "2022-05-09",  # 佛诞翌日 (补假)
    "2022-06-03",  # 端午节
    "2022-07-01",  # 香港特别行政区成立纪念日
    "2022-09-12",  # 中秋节后第二日 (补假)
    "2022-10-01",  # 国庆日
    "2022-10-04",  # 重阳节
    "2022-12-26",  # 圣诞节后第一个周日
    "2022-12-27"   # 圣诞节后第二个周日 (补假)
]
hk_holidays_2021 = [
    "2021-01-01",  # 元旦
    "2021-02-12", "2021-02-13", "2021-02-15",  # 农历年初一、初二、初四 (初四为补假)
    "2021-04-02", "2021-04-03",  # 耶稣受难节及翌日
    "2021-04-05",  # 清明节翌日 (补假)
    "2021-04-06",  # 复活节星期一翌日 (补假)
    "2021-05-01",  # 劳动节
    "2021-05-19",  # 佛诞
    "2021-06-14",  # 端午节
    "2021-07-01",  # 香港特别行政区成立纪念日
    "2021-09-22",  # 中秋节翌日
    "2021-10-01",  # 国庆日
    "2021-10-14",  # 重阳节
    "2021-12-25",  # 圣诞节
    "2021-12-27"   # 圣诞节后第一个平日 (节礼日补假)
]
# 为每个年份生成工作日列表
workdays_dict = {}

# 生成2021年工作日
custom_bday_2021 = pd.offsets.CustomBusinessDay(holidays=hk_holidays_2021)
workdays_2021 = pd.bdate_range(start="2021-01-01", end="2021-12-31", freq=custom_bday_2021)
workdays_dict['2021'] = workdays_2021.strftime('%Y-%m-%d').tolist()

# 生成2022年工作日
custom_bday_2022 = pd.offsets.CustomBusinessDay(holidays=hk_holidays_2022)
workdays_2022 = pd.bdate_range(start="2022-01-01", end="2022-12-31", freq=custom_bday_2022)
workdays_dict['2022'] = workdays_2022.strftime('%Y-%m-%d').tolist()

# 生成2023年工作日
custom_bday_2023 = pd.offsets.CustomBusinessDay(holidays=hk_holidays_2023)
workdays_2023 = pd.bdate_range(start="2023-01-01", end="2023-12-31", freq=custom_bday_2023)
workdays_dict['2023'] = workdays_2023.strftime('%Y-%m-%d').tolist()

# 生成2024年工作日
custom_bday_2024 = pd.offsets.CustomBusinessDay(holidays=hk_holidays_2024)
workdays_2024 = pd.bdate_range(start="2024-01-01", end="2024-12-31", freq=custom_bday_2024)
workdays_dict['2024'] = workdays_2024.strftime('%Y-%m-%d').tolist()

# 生成2025年工作日
custom_bday_2025 = pd.offsets.CustomBusinessDay(holidays=hk_holidays_2025)
workdays_2025 = pd.bdate_range(start="2025-01-01", end="2025-12-31", freq=custom_bday_2025)
workdays_dict['2025'] = workdays_2025.strftime('%Y-%m-%d').tolist()

# 生成2026年工作日
custom_bday_2026 = pd.offsets.CustomBusinessDay(holidays=hk_holidays_2026)
workdays_2026 = pd.bdate_range(start="2026-01-01", end="2026-12-31", freq=custom_bday_2026)
workdays_dict['2026'] = workdays_2026.strftime('%Y-%m-%d').tolist()


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
    
    # 提取年份
    year = query_date_str.split('-')[0]
    
    # 检查年份是否在工作字典中
    if year not in workdays_dict:
        return None
    
    # 获取对应年份的工作日列表
    workdays_list = workdays_dict[year]
    
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
