from playwright.sync_api import sync_playwright
import pandas as pd
import time
import random
import os
from datetime import datetime, timedelta
from summary_updater import update_summary_sheet

# 已查询过的日期列表
QUERIED_DATES = []

# 初始化已查询日期列表 (1月2日到1月26日，2月23日到24日)
def initialize_queried_dates():
    # 1月2日到1月26日
    for day in range(2, 27):
        date = f"2026/01/{day:02d}"
        QUERIED_DATES.append(date)
    # 2月23日到24日
    for day in range(23, 25):
        date = f"2026/02/{day:02d}"
        QUERIED_DATES.append(date)

# 初始化已查询日期
initialize_queried_dates()

STOCK_CODE = "02158"
QUERY_DATE = "2026/02/26"   
OUTPUT_FILE = f"hkex_{STOCK_CODE}_{QUERY_DATE.replace('/', '')}.csv"  # 自动生成输出文件名

URL = "https://www3.hkexnews.hk/sdw/search/searchsdw_c.aspx"

def fetch_sdw():
    with sync_playwright() as p:
        # 启动浏览器
        browser = p.chromium.launch(
            headless=False,  
            args=[
                "--disable-http2",
                "--disable-features=site-per-process",
                "--no-sandbox",
                "--disable-setuid-sandbox"
            ]
        )
        
        page = browser.new_page()
        page.set_default_timeout(60000)
        page.set_extra_http_headers({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1"
        })

        print("正在访问港交所网站...")
        try:
            page.goto(URL, wait_until="load") 
            print("网站访问成功！")
            time.sleep(random.uniform(2, 4))  
        except Exception as e:
            print(f"访问网站时出现超时: {e}")
            
        # 1. 填写股票代码 (使用 locator)
        print(f"正在填写股票代码: {STOCK_CODE}")
        try:
            stock_input = page.locator("#txtStockCode")
            stock_input.fill(STOCK_CODE)
            time.sleep(random.uniform(1, 2))
        except Exception as e:
            print(f"填写股票代码时出错: {e}")

        # 2. 填写日期 (核心修复：JS去只读 + 真实键盘模拟 + 失去焦点)
        print(f"正在填写日期: {QUERY_DATE}")
        try:
            page.evaluate("document.getElementById('txtShareholdingDate').removeAttribute('readonly');")
            time.sleep(1)
            
            date_input = page.locator("#txtShareholdingDate")
            date_input.clear()
            date_input.fill(QUERY_DATE) 
            date_input.blur() 
            time.sleep(random.uniform(1, 2))
            print(f"日期填写完成: {date_input.input_value()}")
        except Exception as e:
            print(f"填写日期时出错: {e}")

        print("正在点击搜索按钮并等待数据加载...")
        try:
            search_btn = page.locator("#btnSearch")
            with page.expect_navigation(wait_until="domcontentloaded", timeout=45000):
                search_btn.click()
            print("页面跳转成功！")
        except Exception as e:
            print(f"点击搜索按钮或等待跳转时出错: {e}")

# 4. 抓取表格数据
        print("正在抓取结果表格...")
        data = []
        try:
            target_table = page.locator("table").filter(has_text="持股量").first
            target_table.wait_for(state="visible", timeout=15000)
            
            # 获取所有表格行
            rows = target_table.locator("tr").all()
            print(f"找到 {len(rows)} 行表格数据(含表头)")

            for row in rows:
                cols = [c.inner_text().strip() for c in row.locator("th, td").all()]
                # 过滤掉空行，根据截图，正规数据通常大于等于4列
                if cols and len(cols) >= 4: 
                    data.append(cols)
                    
            print(f"成功提取 {len(data)-1} 条核心数据")
        except Exception as e:
            print(f"提取表格数据时发生错误: {e}")

        browser.close()
        return data


def main(query_date=None):
    try:
        # 声明全局变量
        global QUERY_DATE, OUTPUT_FILE
        
        # 使用传入的日期或默认日期
        current_date = query_date if query_date else QUERY_DATE
        
        # 检查日期是否已经查询过
        is_already_queried = current_date in QUERIED_DATES
        
        print(f"开始爬取股票 {STOCK_CODE} 在 {current_date} 的数据...")
        # 更新全局变量以便fetch_sdw使用
        QUERY_DATE = current_date
        OUTPUT_FILE = f"hkex_{STOCK_CODE}_{QUERY_DATE.replace('/', '')}.csv"
        
        data = fetch_sdw()
        
        if not data or len(data) < 2:
            print("未获取到足够的数据，程序结束。")
            return False, "未获取到足够的数据", []
            
        # 智能适配表头：直接使用网页抓取到的第一行作为列名
        headers = data[0]
        headers = [h.replace('\n', ' ').strip() for h in headers] 
        rows_data = data[1:]
        
        # 使用动态表头生成 DataFrame
        df = pd.DataFrame(rows_data, columns=headers)

        print("\n=== 数据预览 (前10条) ===")
        print(df.head(10))
        
        # 保存为CSV文件 (utf-8-sig 防止 Excel 中文乱码)
        df.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig")
        print(f"\n数据已成功保存到当前目录下的: {OUTPUT_FILE}")
        
        # 如果日期已经查询过，只返回数据，不插入Excel
        if is_already_queried:
            print(f"日期 {current_date} 已经查询过，跳过Excel插入操作。")
            return True, "该日期已查询过，仅显示数据（未插入Excel）", rows_data
        
        # 保存到Excel文件 - 使用新的文件名避免冲突
        excel_file = r"C:\Users\IT\Desktop\CCASS-202601 (2) - 副本.xlsx"
        # 生成sheet名称：xx年xx月xx日
        date_obj = datetime.strptime(current_date, "%Y/%m/%d")
        sheet_name = date_obj.strftime("%Y年%m月%d日")
        
        excel_saved = False
        try:
            # 检查Excel文件是否存在
            if os.path.exists(excel_file):
                with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"数据已成功追加到Excel文件: {excel_file} 的新sheet '{sheet_name}'")
            else:
                # 如果文件不存在，创建新文件并写入第一个sheet
                df.to_excel(excel_file, sheet_name=sheet_name, index=False)
                print(f"数据已成功保存到新Excel文件: {excel_file} 的sheet '{sheet_name}'")
            
            # 等待文件完全写入和释放
            time.sleep(2)
            
            # 更新summary表
            update_summary_sheet(excel_file, sheet_name)
            
            # 标记Excel保存成功
            excel_saved = True
            
        except PermissionError as e:
            print(f"保存到Excel文件时出错(权限错误): {e}")
            return False, "Excel文件被占用或权限不足，请关闭文件后重试", rows_data
        except Exception as e:
            print(f"保存到Excel文件时出错: {e}")
            import traceback
            traceback.print_exc()
            return False, f"保存到Excel文件时出错: {e}", rows_data
        
        # 只有在Excel成功保存后，才将日期添加到已查询列表
        if excel_saved:
            QUERIED_DATES.append(current_date)
            print(f"日期 {current_date} 已添加到已查询列表。")
        
        # 返回成功信息和数据
        return True, "查询成功", rows_data
        
    except Exception as e:
        print(f"程序运行过程中出现错误: {e}")
        return False, str(e), []


if __name__ == "__main__":
    main()