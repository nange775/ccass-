from playwright.sync_api import sync_playwright
import pandas as pd
import time
import random
import os
import sys
from datetime import datetime, timedelta
from summary_updater import update_summary_sheet

# --- 新增：获取 EXE 运行时的真实路径 ---
def get_base_path():
    if getattr(sys, 'frozen', False):
        # 打包后的 EXE 目录
        return os.path.dirname(sys.executable)
    # 开发环境下的脚本目录
    return os.path.dirname(os.path.abspath(__file__))

STOCK_CODE = "02158"
QUERY_DATE = "2026/03/07"   
URL = "https://www3.hkexnews.hk/sdw/search/searchsdw_c.aspx"

def fetch_sdw():
    with sync_playwright() as p:
        # 启动浏览器 - 已设为 msedge 兼容打包环境
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
        except Exception as e:
            print(f"访问网站时出现超时: {e}")
            
        # 1. 填写股票代码
        print(f"正在填写股票代码: {STOCK_CODE}")
        try:
            stock_input = page.locator("#txtStockCode")
            stock_input.fill(STOCK_CODE)
        except Exception as e:
            print(f"填写股票代码时出错: {e}")

        # 2. 填写日期
        print(f"正在填写日期: {QUERY_DATE}")
        try:
            page.evaluate("document.getElementById('txtShareholdingDate').removeAttribute('readonly');")
            time.sleep(1)
            date_input = page.locator("#txtShareholdingDate")
            date_input.clear()
            date_input.fill(QUERY_DATE) 
            date_input.blur() 
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
            rows = target_table.locator("tr").all()
            
            for row in rows:
                raw_cols = row.locator("th, td").all_inner_texts()
                cols = [c.strip() for c in raw_cols]
                if cols and len(cols) >= 4: 
                    data.append(cols)
            print(f"成功提取 {len(data)-1} 条核心数据")
        except Exception as e:
            print(f"提取表格数据时发生错误: {e}")

        browser.close()
        return data

def main(query_date=None, file_path='.', stock_code=None):
    global QUERY_DATE, STOCK_CODE
    
    try:
        current_date = query_date if query_date else QUERY_DATE
        if stock_code:
            STOCK_CODE = stock_code
        
        print(f"开始爬取股票 {STOCK_CODE} 在 {current_date} 的数据...")
        QUERY_DATE = current_date
        
        data = fetch_sdw()
        
        if not data or len(data) < 2:
            return False, "未获取到足够的数据", []
            
        headers = [h.replace('\n', ' ').strip() for h in data[0]]
        rows_data = data[1:]
        df = pd.DataFrame(rows_data, columns=headers)

        # --- 路径处理：确保保存到 EXE 旁边 ---
        base_dir = get_base_path()
        
        # 如果 file_path 是默认的 '.' 或者用户没传，则使用 EXE 目录
        final_save_dir = file_path if (file_path and file_path != '.') else base_dir
        
        if final_save_dir.lower().endswith('.xlsx'):
            excel_file = final_save_dir
        else:
            # 默认文件名
            excel_file = os.path.join(final_save_dir, "CCASS-202601 (2).xlsx")
        
        # 确保目录存在
        os.makedirs(os.path.dirname(excel_file), exist_ok=True)

        # 生成sheet名称：xx年xx月xx日
        date_obj = datetime.strptime(current_date, "%Y/%m/%d")
        sheet_name = date_obj.strftime("%Y年%m月%d日")
        
        print(f"Excel保存路径: {excel_file}")

        try:
            if os.path.exists(excel_file):
                # 追加模式
                with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                # 新建模式
                df.to_excel(excel_file, sheet_name=sheet_name, index=False)
            
            # 更新汇总表
            print("正在更新 summary 表...")
            update_summary_sheet(excel_file, sheet_name)
            return True, "查询成功", rows_data
            
        except PermissionError:
            return False, f"错误：文件 {os.path.basename(excel_file)} 正被打开，请关闭后再查询", rows_data
        except Exception as e:
            return True, f"数据抓取成功，但写入Excel失败: {str(e)}", rows_data
        
    except Exception as e:
        print(f"程序运行过程中出现错误: {e}")
        return False, str(e), []

if __name__ == "__main__":
    main()