from playwright.sync_api import sync_playwright
import pandas as pd
import time
import random
import os
import sys
from datetime import datetime, timedelta
from summary_updater import update_summary_sheet
from hk_calendar import get_real_date

# --- 新增：获取 EXE 运行时的真实路径 ---
def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

STOCK_CODE = "02158"
QUERY_DATE = "2025/03/03"
URL = "https://www3.hkexnews.hk/sdw/search/searchsdw_c.aspx"

browser = None
context = None
page = None

def init_browser():
    global browser, context, page
    # 修正：sync_playwright().start() 需要手动管理，或者在 init 内部完成
    from playwright.sync_api import sync_playwright
    p = sync_playwright().start()
    browser = p.chromium.launch(
        headless=False,
        channel="msedge",
        args=[
            "--disable-http2",
            "--disable-features=site-per-process",
            "--no-sandbox",
            "--disable-setuid-sandbox"
        ]
    )
    context = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    page = context.new_page()
    page.set_default_timeout(60000)
    print("浏览器已启动 (Edge Channel)")

def close_browser():
    global browser
    if browser:
        browser.close()
        print("浏览器已关闭")

def fetch_sdw(query_date, first_load=False):
    global page
    data = []
    
    if first_load:
        print("正在打开页面...")
        try:
            page.goto(URL, wait_until="load")
        except Exception as e:
            print(f"访问网站时出现超时: {e}")
            return data
        
        try:
            stock_input = page.locator("#txtStockCode")
            stock_input.fill(STOCK_CODE)
        except Exception as e:
            print(f"填写股票代码时出错: {e}")
    
    print(f"正在填写日期: {query_date}")
    try:
        page.evaluate("document.getElementById('txtShareholdingDate').removeAttribute('readonly');")
        date_input = page.locator("#txtShareholdingDate")
        date_input.clear()
        date_input.fill(query_date)
        date_input.blur()
    except Exception as e:
        print(f"填写日期时出错: {e}")

    try:
        search_btn = page.locator("#btnSearch")
        search_btn.click()
    except Exception as e:
        print(f"点击搜索按钮时出错: {e}")

    try:
        target_table = page.locator("table").filter(has_text="持股量").first
        target_table.wait_for(state="visible", timeout=15000)
        rows = target_table.locator("tr").all()
        for row in rows:
            cells = row.locator("th, td").all_inner_texts()
            if cells and len(cells) >= 4:
                data.append(cells)
    except Exception as e:
        print(f"提取表格数据时发生错误: {e}")

    return data

def fetch_single_date(current_date, first_run=False):
    print(f"\n开始抓取: {current_date}")
    data = fetch_sdw(current_date, first_run)
    if not data or len(data) < 2:
        return False
        
    headers = [h.replace('\n', ' ').strip() for h in data[0]]
    rows_data = data[1:]
    df = pd.DataFrame(rows_data, columns=headers)
    date_obj = datetime.strptime(current_date, "%Y/%m/%d")
    sheet_name = date_obj.strftime("%Y年%m月%d日")
    return (df, sheet_name)

def main(start_date=None, end_date=None, stock_code=None, file_path='.'):
    global STOCK_CODE
    if stock_code: STOCK_CODE = stock_code
    
    use_start_date = start_date if start_date else "2026/03/04"
    use_end_date = end_date if end_date else "2026/03/07"
    
    # 路径解析：如果是默认的点，重定向到 EXE 目录
    base_dir = get_base_path()
    if not file_path or file_path == '.':
        final_dir = base_dir
    else:
        final_dir = file_path

    init_browser()
    
    try:
        start = datetime.strptime(use_start_date, "%Y/%m/%d")
        end = datetime.strptime(use_end_date, "%Y/%m/%d")
        current = start
        success_count, skip_count = 0, 0
        first_run = True
        data_cache = []
        
        while current <= end:
            query_date_str = current.strftime("%Y/%m/%d")
            real_date = get_real_date(query_date_str)
            
            if real_date is None:
                print(f"日期 {query_date_str} 非交易日，跳过")
                skip_count += 1
            else:
                result = fetch_single_date(query_date_str, first_run)
                if result:
                    success_count += 1
                    data_cache.append(result)
                first_run = False
            current += timedelta(days=1)
        
        if data_cache:
            # 确定 Excel 文件路径
            if final_dir.lower().endswith('.xlsx'):
                excel_file = final_dir
            else:
                excel_file = os.path.join(final_dir, "ccass汇总.xlsx")
            
            os.makedirs(os.path.dirname(excel_file), exist_ok=True)
            file_exists = os.path.exists(excel_file)
            
            try:
                print(f"\n正在写入文件: {excel_file}")
                # 批量写入
                with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a' if file_exists else 'w', if_sheet_exists='replace') as writer:
                    for df, sheet_name in data_cache:
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        print(f"写入 Sheet: {sheet_name}")
                
                # 批量更新 Summary
                for _, sheet_name in data_cache:
                    try:
                        update_summary_sheet(excel_file, sheet_name)
                    except Exception as e:
                        print(f"Summary 更新失败({sheet_name}): {e}")
                
                return True, f"成功爬取 {success_count} 天数据并汇总", []

            except PermissionError:
                return False, f"写入失败：请先关闭正在打开的 Excel 文件 '{os.path.basename(excel_file)}'", []
        
        return True, f"批量任务完成，抓取成功 {success_count} 天", []
        
    finally:
        close_browser()

if __name__ == "__main__":
    main()