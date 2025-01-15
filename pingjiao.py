import time
import random
import pandas as pd

from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from urllib3.exceptions import MaxRetryError


def login(driver, username, password):
    print("\n[login] -------------------------------------------------")
    print("[login] 正在尝试登录：")
    print(f"[login] username: {username}")
    print(f"[login] password: {password}")
    wait = WebDriverWait(driver, 10)

    try:
        username_input = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[1]/div[2]/div[2]/form/div/div/div[1]/div[1]/div/input')
            )
        )
        print("[login] 找到用户名输入框，开始输入...")
        for char in username:
            username_input.send_keys(char)
            time.sleep(random.uniform(0.05, 0.2))

        password_input = driver.find_element(By.XPATH,
                                             '/html/body/div[1]/div[2]/div[2]/form/div/div/div[1]/div[2]/div/input[2]')
        print("[login] 找到密码输入框，开始输入...")
        for char in password:
            password_input.send_keys(char)
            time.sleep(random.uniform(0.05, 0.2))

        login_button = driver.find_element(By.XPATH,
                                           '/html/body/div[1]/div[2]/div[2]/form/div/div/div[1]/div[3]/button')
        print("[login] 即将点击登录按钮...")
        login_button.click()
        print("[login] 已点击登录按钮。")
    except Exception as e:
        print(f"[login] 登录出现异常: {e}")
        raise e


def find_teachers(driver):
    """
    在 table#tempGrid 中找第一个“未评”<td> 并点击。
    返回 True 表示找到并点击；False 表示未找到。
    """
    print("\n[find_teachers] >>> 开始等待表格 #tempGrid 完整加载，并寻找“未评”...")

    try:
        wait = WebDriverWait(driver, 10)

        # 1) 等待表格本身可见/加载完成
        table = wait.until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "table#tempGrid"))
        )
        print("[find_teachers] 已检测到表格 #tempGrid 可见.")

        # 2) 再等待“未评”单元格真正加载
        #    因为 jqGrid 等可能还在异步渲染数据
        #    这里我们使用 presence_of_all_elements_located，表示只要出现就算成功
        xpath_unrated = (
            "//table[@id='tempGrid']"
            "//td[@aria-describedby='tempGrid_tjztmc' and @title='未评']"
        )
        wait.until(
            EC.presence_of_all_elements_located((By.XPATH, xpath_unrated))
        )
        print("[find_teachers] 页面已出现“未评”单元格元素 (至少一个).")

        # 3) 获取所有“未评”并点击第一个
        cells_unrated = driver.find_elements(By.XPATH, xpath_unrated)
        print(f"[find_teachers] 实际未评单元格数量: {len(cells_unrated)}")

        if not cells_unrated:
            # 理论上不会走到这儿，因为上面 wait 已经等到至少1个
            print("[find_teachers] 没有找到任何未评，返回 False.")
            return False

        # 点击第一个“未评”
        first_unrated = cells_unrated[0]
        print("[find_teachers] 点击第一个未评单元格 =>", first_unrated.text)
        first_unrated.click()
        print("[find_teachers] 已点击“未评”。")
        return True

    except TimeoutException:
        print("[find_teachers] 10秒内仍无法定位 #tempGrid 或未评数据，表格可能还没加载完/没有未评。")
        return False
    except Exception as e:
        print(f"[find_teachers] 发生异常: {e}")
        return False


def in_fact(driver):
    """
    判断是否还有下一页可点，返回 True 表示还有下一页。
    你的逻辑：若 int(value_element.text) != 0 则还有下一页。
    """
    print("\n[in_fact] -----------------------------------------------")
    wait = WebDriverWait(driver, 10)
    try:
        value_element = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[2]/div/div/div[3]/div[1]/div/div[1]/h3[1]/a[3]/span'))
        )
        text_val = value_element.text
        print(f"[in_fact] 从页面读到的文本是: '{text_val}'")
        value = int(text_val)
        print(f"[in_fact] 数字部分为: {value}, 是否不等于0 => {value != 0}")
        return value != 0
    except TimeoutException:
        print("[in_fact] 超时：没有找到该元素，可能无法翻页，返回False。")
        return False
    except ValueError:
        print("[in_fact] 解析数字失败，返回False。")
        return False
    except Exception as e:
        print(f"[in_fact] 其他异常: {e}")
        return False


def change_page(driver):
    """
    点击下一页按钮 <td id="next_pager" ...>
    如果带有 'ui-state-disabled' 说明不能点击，直接返回。
    """
    print("\n[change_page] >>> 尝试点击下一页...")

    try:
        wait = WebDriverWait(driver, 3)
        next_btn = wait.until(
            EC.presence_of_element_located((By.ID, "next_pager"))
        )

        # 检查 class 里是否有 ui-state-disabled
        cls = next_btn.get_attribute("class")
        print(f"[change_page] next_pager class: {cls}")
        if "ui-state-disabled" in cls:
            print("[change_page] next_pager 是 disabled 状态，无法点击。")
            return

        print("[change_page] next_pager 可点击，执行click()...")
        next_btn.click()
        time.sleep(1)  # 等待页面加载
        print("[change_page] 已点击下一页。")

    except TimeoutException:
        print("[change_page] 3秒内无法定位 next_pager，可能页面无分页。")
    except Exception as e:
        print(f"[change_page] 发生异常: {e}")


def has_next_page(driver):
    """
    判断是否还有下一页：
      1) 读取当前页 input[name="showCount"] 的 value
      2) 读取总页数 span#sp_1_pager 的文本
      3) 判断 current_page < total_page
    """
    print("\n[has_next_page] >>> 判断是否还有下一页...")

    try:
        wait = WebDriverWait(driver, 3)

        # 读当前页(假设在此 input)
        current_page_input = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input.ui-pg-input[name='showCount']"))
        )
        current_page_value = current_page_input.get_attribute("value").strip()
        print(f"[has_next_page] 当前页value = {current_page_value}")

        # 读总页数 <span id="sp_1_pager">
        total_page_span = driver.find_element(By.ID, "sp_1_pager")
        total_page = total_page_span.text.strip()
        print(f"[has_next_page] 总页数 = {total_page}")

        # 转成整数比较
        cur = int(current_page_value)
        tot = int(total_page)
        print(f"[has_next_page] 解析数字 => 当前页: {cur}, 总页: {tot}")

        # 如果系统从 0 开始计数，则 < tot 表示还有下一页
        if cur + 1 < tot:
            print("[has_next_page] 还有下一页")
            return True
        else:
            print("[has_next_page] 没有下一页了")
            return False

    except TimeoutException:
        print("[has_next_page] 3秒内无法定位分页元素，可能没有分页。返回 False.")
        return False
    except Exception as e:
        print(f"[has_next_page] 发生异常: {e}")
        return False


def grade(driver):
    """
    在评教页面中，对所有问题逐行打分。
    """
    print("\n[grade] -----------------------------------------------")
    wait = WebDriverWait(driver, 10)

    try:
        # 1. 找到“评教表格”
        table = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table.table.table-bordered.table-xspj"))
        )
        print("[grade] 找到表格(table.table.table-bordered.table-xspj)。")
        # 输出表格的 outerHTML 以debug
        table_html = table.get_attribute("outerHTML")
        print(f"[grade] 表格HTML片段:\n{table_html[:500]}...\n")  # 只打印前500字符, 避免太长

        # 2. 找到所有 <tr class="tr-xspj">
        tbody = table.find_element(By.TAG_NAME, "tbody")
        rows = tbody.find_elements(By.CSS_SELECTOR, "tr.tr-xspj")
        print(f"[grade] 共找到 {len(rows)} 行 .tr-xspj.")

        if len(rows) == 0:
            print("[grade] 没有找到可评价行，可能表格结构不一致。直接返回。")
            return

        # 3. 随机选一行打 B(80)，其它行打 A(100)
        random_index = random.randint(0, len(rows) - 1)
        print(f"[grade] 随机选择第 {random_index + 1} 行打 B，其它打 A。")

        for i, row in enumerate(rows):
            try:
                print(f"\n[grade] 正在处理第 {i + 1}/{len(rows)} 行 ...")
                row_html = row.get_attribute("outerHTML")
                print(f"[grade] 该行HTML:\n{row_html[:300]}...\n")

                tds = row.find_elements(By.TAG_NAME, "td")
                print(f"[grade] 第{i + 1}行包含 {len(tds)} 个 <td>。")

                if len(tds) < 2:
                    print(f"[grade] 第{i + 1}行结构异常，跳过.")
                    continue

                rating_td = tds[1]  # 第二列
                # 确定打分
                if i == random_index:
                    css_class = ".radio-inline.input-xspj.input-xspj-2"  # B(80)
                else:
                    css_class = ".radio-inline.input-xspj.input-xspj-1"  # A(100)

                # 找到目标 div
                print(f"[grade] 目标CSS选择器: {css_class}")
                radio_div = rating_td.find_element(By.CSS_SELECTOR, css_class)
                radio_input = radio_div.find_element(By.TAG_NAME, "input")

                wait.until(EC.element_to_be_clickable(radio_input))
                radio_input.click()

                print(f"[grade] >>> 已点击第{i + 1}行的 {css_class[-1]} 选项.")
                # 小延时
                time.sleep(0.2)

            except Exception as e:
                print(f"[grade] 第{i + 1}行打分失败: {e}")

    except TimeoutException:
        print("[grade] 在10s内未找到评教表格table.table.table-bordered.table-xspj，无法打分。")
    except Exception as e:
        print(f"[grade] 发生未知错误: {e}")


def submit(driver, max_retries=3):
    """
    点击“提交”按钮，然后点击确认。
    遇到网络卡顿时刷新页面并重试(最多 max_retries 次)。
    """
    print("\n[submit] ---------------------------------------------")
    attempts = 0
    while attempts < max_retries:
        attempts += 1
        print(f"[submit] 尝试第 {attempts} 次提交...")
        try:
            wait = WebDriverWait(driver, 10)

            submit_button = wait.until(
                EC.element_to_be_clickable((By.XPATH,
                                            '/html/body/div[2]/div/div/div[3]/div[2]/div/div[3]/form/div[2]/div[2]/div/div/div/button[2]'))
            )
            print("[submit] 找到'提交'按钮，点击...")
            submit_button.click()

            confirm_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[5]/div/div/div[3]/button'))
            )
            print("[submit] 找到二次确认按钮，点击...")
            confirm_button.click()
            print("[submit] 提交完成.")

            # 如果成功执行到这里，就 break 跳出重试循环
            break

        except TimeoutException:
            # 提交或确认按钮在 10 秒内不可点击 => 可能网络卡顿
            print("[submit] 等待按钮超时，尝试刷新页面后重试...")
            driver.refresh()
            time.sleep(3)  # 给点时间重新加载

        except Exception as e:
            # 其他异常可以直接抛，或也尝试刷新
            print(f"[submit] 发生异常: {e}")
            driver.refresh()
            time.sleep(3)
            # 也可以 break 或者继续重试，看你需求

    else:
        # 如果 while 正常结束(没有break)，说明重试用尽但仍失败
        print(f"[submit] 提交重试 {max_retries} 次后仍失败，可能网络或页面异常。")


def check_submit(driver):
    """
    检查总分，如果没问题就调用 submit(driver)。
    如果和预期不符就抛异常。
    """
    print("\n[check_submit] ----------------------------------------")
    try:
        score_element = driver.find_element(
            By.XPATH,
            '//*[@id="ajaxForm1"]/div[2]/div[1]/div[1]/h4/i'
        )
        total_score = score_element.text.strip()
        print(f"[check_submit] 当前总分文本为: {total_score}")

        # 如果不符合预期则抛异常；保留你之前的逻辑
        if total_score not in ('总分: 93.00', '总分: 98.00'):
            raise Exception(f'评分错误，不在(93.00/98.00)之内，实际为: {total_score}')
        else:
            print("[check_submit] 分数符合预期，将提交...")
            submit(driver)

    except TimeoutException:
        print("[check_submit] 成绩元素加载超时，请检查页面或网络状态。")
    except Exception as e:
        print(f"[check_submit] 检查评分时发生异常：{e}")


def main():
    print("\n[main] *************************************************")
    print("[main] 准备启动 EdgeDriver...")
    options = Options()
    driver_path = r'D:\edgedownload\edgedriver_win64\msedgedriver.exe'
    service = Service(executable_path=driver_path)

    # 读取Excel文件
    excel_path = r"C:\Users\yuesen\Desktop\评教用户信息.xlsx"
    accounts_df = pd.read_excel(excel_path)
    print(f"[main] 共读取到 {len(accounts_df)} 个账号信息。")

    for index, row in accounts_df.iterrows():
        username = str(row['username'])
        password = str(row['password'])

        print("\n[main] --------------------------------------------")
        print(f"[main] 开始处理第 {index + 1} 个账号 => 用户名: {username}")

        driver = webdriver.Edge(service=service, options=options)
        try:
            url = ("https://jw.sdyu.edu.cn/jwglxt/xspjgl/xspj_cxXspjIndex.html"
                   "?doType=details&gnmkdm=N401605&layout=default&su=202302710222")
            print(f"[main] 打开网址: {url}")
            driver.get(url)

            # 登录
            login(driver, username, password)
            time.sleep(2)

            # 如果需要先切换到iframe等，请在此处进行
            # driver.switch_to.frame(...)  # 根据实际情况

            while True:
                # 1) 找本页是否有"未评"
                found = find_teachers(driver)
                if found:
                    # 如果找到了 => 进入评教页面
                    grade(driver)
                    time.sleep(0.5)
                    check_submit(driver)  # 内部会调用改进后的 submit(driver)
                    time.sleep(1)

                    # 提交后，系统通常会跳回评教列表
                    # 如果系统不自动回列表，需要 driver.back() 或 driver.get(列表URL) 等操作
                else:
                    # 没找到"未评"
                    can_turn = has_next_page(driver)  # 判断是否还有下一页
                    if can_turn:
                        print("[main] 本页没有未评了，但还有下一页，尝试翻页...")
                        change_page(driver)
                        time.sleep(2)  # 等新页加载
                    else:
                        print("[main] 本页没有未评 & 没有下一页，评教结束。")
                        break

            print(f"[main] 账号 {username} 评教流程结束。")

        except MaxRetryError as e:
            print(f"[main] 账号{username}连接失败，正在尝试重启...")
            print(e)
            continue
        except Exception as e:
            print(f"[main] 账号{username}发生错误：{e}")
            break
        finally:
            driver.quit()

    print("\n[main] 所有账号评教完成.")


if __name__ == '__main__':
    main()
