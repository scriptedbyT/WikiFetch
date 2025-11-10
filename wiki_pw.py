from playwright.sync_api import Playwright, sync_playwright, expect
import os
import pandas as pd
import datetime as dt, time
from time import sleep, perf_counter
from openpyxl.workbook import Workbook

# Excel - Input Data
xl_fl = 'topics_list.xlsx'
df = pd.read_excel(xl_fl)

# Excel Columns
lg = df["Language"].values.tolist()
src = df["Search"].values.tolist()


url = 'https://www.wikipedia.org/'
sc = []

dt_tm = dt.datetime.fromtimestamp(time.time())
dt_stamp = dt_tm.strftime("%d-%B-%Y")

# def information(title):
#     print(title)
#     print('Module Nmae:', __name__)
#     print('Parent Process:', os.getppid())
#     print('Process ID:', os.getpid())

def run(playwright: Playwright) -> None:
    # Browser [Chrome, Firefox, WebKit]
    browser = playwright.chromium.launch(headless=False)
    
    context = browser.new_context()
    page = context.new_page()
    sleep(1.5)

    for idx, n in enumerate(src):
        # Browsing the website
        # page.reload()
        sleep(1.5)
        page.goto(url)

        # Language Selection
        page.locator('//label[@id="jsLangLabel"]//following::select').select_option(str(lg[idx]))
        sleep(1.5)

        # Topic Search
        page.locator('//input[@id="searchInput"]').fill(str(src[idx]))
        sleep(2)
        page.keyboard.press('Enter')
        sleep(2)

        cnt = page.locator('//div[contains(@id, "bodyContent")]//p').count()

        for x in range(cnt):
            m = page.locator('//div[contains(@id, "bodyContent")]//p').nth(x).inner_text()
            print(m)
            sc.append(m)

    my_dic = {
        "Topic": sc
    }
    df = pd.DataFrame(data=my_dic)
    excel_path = '/Users/taneshkamehta/Documents/RPA/Web_Scraping/Wikipedia/topic_list.xlsx'
    absolute_path = os.path.abspath(excel_path)
    directory_path = os.path.dirname(absolute_path)
    file_name = 'Extractions_%s.xlsx' %dt_tm
    file_path = os.path.join(directory_path, file_name)
    df.to_excel(file_path, index=True)
    # df.to_csv(file_path, index=False)

    context.close()
    browser.close()

def fn_run():
    with sync_playwright() as playwright:
        run(playwright)

if __name__ == "__main__":
    start = perf_counter()
    fn_run()
    end = perf_counter()
    print(f'\n---------------\n Finished in {round(end-start, 2)} second(s)')
