import os
from datetime import datetime, timedelta
from summary_updater import update_summary_sheet
from hk_calendar import get_real_date

def main():
    # 定义日期范围
    START_DATE = "2025/12/06"
    END_DATE = "2026/01/07"
    
    # Excel 文件路径
    excel_file = "ccass汇总.xlsx"
    
    # 检查文件是否存在
    if not os.path.exists(excel_file):
        print(f"找不到文件 {excel_file}")
        return
    
    print(f"开始更新日期范围 {START_DATE} 到 {END_DATE} 的数据到 summary sheet...")
    
    # 转换日期格式
    start = datetime.strptime(START_DATE, "%Y/%m/%d")
    end = datetime.strptime(END_DATE, "%Y/%m/%d")
    
    current = start
    success_count = 0
    skip_count = 0
    
    while current <= end:
        query_date_str = current.strftime("%Y/%m/%d")
        real_date = get_real_date(query_date_str)
        
        if real_date is None:
            print(f"日期 {query_date_str} 不是港股交易日，跳过")
            current += timedelta(days=1)
            skip_count += 1
            continue
        
        # 生成工作表名称
        sheet_name = current.strftime("%Y年%m月%d日")
        
        try:
            print(f"\n更新 {sheet_name} 到 summary sheet...")
            update_summary_sheet(excel_file, sheet_name)
            success_count += 1
        except Exception as e:
            print(f"更新 {sheet_name} 失败: {e}")
            skip_count += 1
        
        current += timedelta(days=1)
    
    print(f"\n{'='*50}")
    print(f"更新完成！成功: {success_count} 天, 跳过: {skip_count} 天")

if __name__ == "__main__":
    main()