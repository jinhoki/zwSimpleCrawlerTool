from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import random
import pandas as pd
from webdriver_manager.chrome import ChromeDriverManager


def init_driver():
    """初始化驱动（修复version参数错误）"""
    chrome_options = webdriver.ChromeOptions()
    # 关闭无头模式，显示浏览器
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    # 关键：禁用自动化检测
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    # 修复：移除version参数，使用默认最新版
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),  # 这里去掉了version="latest"
        options=chrome_options
    )
    # 禁用webdriver标识（防反爬）
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    driver.implicitly_wait(10)
    driver.maximize_window()
    return driver


def manual_filter_reminder():
    """提示手动完成筛选"""
    print("=" * 50)
    print("请你手动完成以下操作：")
    print("1. 在知网页面搜索框输入「计算机」并搜索")##此处因为主函数的url已是计算机检索后的界面，这里其实没有实际用处
    ##若为未输入检索界面，可以在这里添加内容提示搜索
    print("2. 筛选「博士论文」")##根据个人需求选择内容
    print("3. 选择学科「计算机软件及计算机应用」")
    print("4. 设置排序方式为「综合排序」")
    print("5. 操作完成后按回车继续...")
    print("=" * 50)
    input()  # 等待用户按回车


def parse_paper_data(driver):
    """解析筛选后的论文数据（精准匹配你的表格结构）"""
    paper_list = []
    try:
        print("\n开始解析论文数据...")
        time.sleep(3)

        # 精准定位表格（按class）
        table = driver.find_element(By.CLASS_NAME, "result-table-list")
        # 定位tbody下的所有数据行
        rows = table.find_element(By.TAG_NAME, "tbody").find_elements(By.TAG_NAME, "tr")
        print(f"找到 {len(rows)} 条论文数据")

        #此处是根据对应网站的对应html界面获取的表格元素，不同的界面对应的标签内容不同
        for row in rows:
            paper_info = {}
            try:
                # 中文题名（class=name）
                name_td = row.find_element(By.CLASS_NAME, "name")
                title_a = name_td.find_element(By.TAG_NAME, "a")
                paper_info['中文题名'] = title_a.text.strip()
                paper_info['链接'] = title_a.get_attribute("href") or "未知"

                # 作者（class=author）
                paper_info['作者'] = row.find_element(By.CLASS_NAME, "author").text.strip() or "未知"

                # 学位授予单位（class=unit）
                paper_info['学位授予单位'] = row.find_element(By.CLASS_NAME, "unit").text.strip() or "未知"

                # 学位授予年度（class=date）
                paper_info['学位授予年度'] = row.find_element(By.CLASS_NAME, "date").text.strip() or "未知"

                # 被引（class=quote）
                paper_info['被引'] = row.find_element(By.CLASS_NAME, "quote").text.strip() or "0"

                # 下载（class=download）
                paper_info['下载'] = row.find_element(By.CLASS_NAME, "download").text.strip() or "0"

                paper_list.append(paper_info)
            except Exception as e:
                print(f"解析单条失败：{e}")
                continue

        return paper_list
    except Exception as e:
        print(f"解析数据失败：{e}")
        return []


def save_data(data):
    """保存数据"""
    if not data:
        print("无数据可保存")
        return
    try:
        df = pd.DataFrame(data)
        df.to_excel("知网博士论文数据.xlsx", index=False, engine="openpyxl")
        print(f"数据已保存到：知网博士论文数据.xlsx（共{len(data)}条）")
    except:
        df.to_csv("知网博士论文数据.csv", index=False, encoding="utf-8-sig")
        print(f"数据已保存到：知网博士论文数据.csv（共{len(data)}条）")


def main():
    #知网基础URL，此处用于指定需要获取的界面，运行之后可以直接跳转该界面
    cnki_base_url = "https://kns.cnki.net/kns8s/defaultresult/index?crossids=YSTT4HG0%2CLSTPFY1C%2CJUP3MUPD%2CMPMFIG1A%2CWQ0UVIAA%2CBLZOG7CK%2CPWFIRAGL%2CEMRPGLPA%2CNLBO1Z6R%2CNN3FJMUV&korder=SU&kw=%E8%AE%A1%E7%AE%97%E6%9C%BA"

    # 初始化驱动
    driver = init_driver()
    try:
        # 步骤1：打开知网页面
        print("打开知网页面...")
        driver.get(cnki_base_url)
        time.sleep(5)

        # 步骤2：提示手动完成筛选（避开自动化定位失败问题）
        manual_filter_reminder()

        # 步骤3：解析数据
        paper_data = parse_paper_data(driver)

        # 步骤4：预览+保存数据
        if paper_data:
            print("\n数据预览（前3条）：")
            for idx, p in enumerate(paper_data[:3], 1):
                print(f"\n【第{idx}条】")
                for k, v in p.items():
                    print(f"{k}: {v}")
            save_data(paper_data)
        else:
            print("未解析到任何数据")

    except Exception as e:
        print(f"程序出错：{e}")
    finally:
        print("\n15秒后关闭浏览器（可手动关闭）...")
        time.sleep(15)
        driver.quit()


if __name__ == "__main__":
    # 安装依赖：pip install selenium pandas openpyxl webdriver-manager
    main()