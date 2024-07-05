from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd
import time
import random
from urllib3.exceptions import MaxRetryError


def grade(driver):
    # 等待tbody元素加载完成
    tbody = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[2]/div/div/div[3]/div[2]/div/div[3]/form/div[2]/div[1]/div[2]/table/tbody'))
    )

    # 初始化字典，用于存储每个教师的评价选项
    rating_pairs = {}

    # 定位到tbody元素下的所有tr元素
    rows = tbody.find_elements(By.XPATH, './/tr')

    # 遍历每个tr元素，即每个教师的评价行
    for row in rows:
        # 定位到tr下的第二个td元素，即包含评价选项的td
        rating_td = row.find_element(By.XPATH, './/td[2]')

        # 在当前td下的div元素中查找所有子div元素
        rating_divs = rating_td.find_element(By.TAG_NAME, 'div').find_elements(By.TAG_NAME, 'div')

        # 筛选出优秀和良好的input元素
        excellent_inputs = [div.find_element(By.TAG_NAME, 'label').find_element(By.TAG_NAME, 'input') for div in
                            rating_divs if "input-xspj-1" in div.get_attribute('class')]
        good_inputs = [div.find_element(By.TAG_NAME, 'label').find_element(By.TAG_NAME, 'input') for div in rating_divs
                       if "input-xspj-2" in div.get_attribute('class')]

        # 存储当前教师的评价选项
        rating_pairs[row] = {'excellent': excellent_inputs, 'good': good_inputs}

    # 确保有足够的评价组合可供选择
    if len(rating_pairs) < 1:
        raise ValueError("没有评价组合可供选择")

    # 生成1-n的随机数，n为字典元素的数量
    random_index = random.randint(1, len(rating_pairs))

    # 根据随机数，更新字典的评价选项
    for i, (row, ratings) in enumerate(rating_pairs.items()):
        if i == random_index - 1:
            # 随机选择的教师评为良好
            ratings['excellent'] = []  # 清空优秀列表
        else:
            # 其他教师评为优秀
            ratings['good'] = []  # 清空良好列表

    # 对每个教师进行评价
    for row, ratings in rating_pairs.items():
        # 点击该教师的优秀评价
        for input_element in ratings['excellent']:
            WebDriverWait(input_element, 10).until(EC.element_to_be_clickable(input_element))
            input_element.click()
        # 点击该教师的良好评价（如果有）
        for input_element in ratings['good']:
            WebDriverWait(input_element, 10).until(EC.element_to_be_clickable(input_element))
            input_element.click()


def in_fact(driver):
    value_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[2]/div/div/div[3]/div[1]/div/div[1]/h3[1]/a[3]/span'))
    )
    value = int(value_element.text)
    return value != 0


def change_page(driver):
    next_page_button = driver.find_element(By.XPATH,
                                           '/html/body/div[2]/div/div/div[3]/div[1]/div/div[2]/div[2]/div[5]/div/table/tbody/tr/td[2]/table/tbody/tr/td[6]/span')
    next_page_button.click()


def submit(driver):
    submit_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH,
                                    '/html/body/div[2]/div/div/div[3]/div[2]/div/div[3]/form/div[2]/div[2]/div/div/div/button[2]'))
    )
    submit_button.click()
    time.sleep(2)
    confirm_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '/html/body/div[5]/div/div/div[3]/button'))
    )
    confirm_button.click()


def check_submit(driver):
    try:
        score_element = driver.find_element(By.XPATH,
                                            '/html/body/div[2]/div/div/div[3]/div[2]/div/div[3]/form/div[2]/div[1]/div[1]/h4/i')
        total_score = score_element.text.strip()
        if total_score != '总分: 93.00' and total_score != '总分: 98.00':
            raise Exception(f'评分错误，实际为：{total_score}')
        else:
            submit(driver)
    except TimeoutException:
        print("成绩元素加载超时，请检查页面或网络状态。")
    except Exception as e:
        print(f"发生异常：{e}")


def find_teachers(driver):
    try:
        wait = WebDriverWait(driver, 2)
        teachers_unevaluated = wait.until(
            EC.presence_of_all_elements_located((By.XPATH,
                                                 "/html/body/div[2]/div/div/div[3]/div[1]/div/div[2]/div[2]/div[3]/div[3]/div/table/tbody/tr/td[contains(translate(@style, ' RGB', ''), 'color:red') and normalize-space(text())='未评']"))
        )
        teachers_to_evaluate = [teacher for teacher in teachers_unevaluated if '未评' in teacher.text]
        teacher = teachers_to_evaluate[0]
        teacher.click()
        return True
    except TimeoutException:
        return False


def login(driver, username, password):
    print('-'*10)
    print('正在进行：')
    print(username)
    print(password)
    print('-' * 10)
    username_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, '/html/body/div[1]/div[2]/div[2]/form/div/div/div[1]/div[1]/div/input'))
    )
    for char in username:
        username_input.send_keys(char)
        time.sleep(random.uniform(0.1, 0.3))

    password_input = driver.find_element(By.XPATH,
                                         '/html/body/div[1]/div[2]/div[2]/form/div/div/div[1]/div[2]/div/input[2]')
    # # 直接赋值进去
    # password_input.send_keys(password)

    for char in password:
        password_input.send_keys(char)
        time.sleep(random.uniform(0.1, 0.3))
    login_button = driver.find_element(By.XPATH, '/html/body/div[1]/div[2]/div[2]/form/div/div/div[1]/div[3]/button')
    login_button.click()


def main():
    options = Options()
    driver_path = 'Enter the address of your msedgedriver.exe here'
    service = Service(executable_path=driver_path)

    # 读取Excel文件
    excel_path = "Enter the address of the form where the account number will be stored"
    accounts_df = pd.read_excel(excel_path)

    for index, row in accounts_df.iterrows():
        username = str(row['username'])
        password = str(row['password'])

        # 为每个账号创建一个新的WebDriver实例
        driver = webdriver.Edge(service=service, options=options)
        try:
            # 访问评教系统首页
            driver.get(
                'https://jw.sdyu.edu.cn/jwglxt/xspjgl/xspj_cxXspjIndex.html?doType=details&gnmkdm=N401605&layout=default&su=888888888888'
            )

            # 登录
            login(driver, username, password)
            time.sleep(2)  # 等待登录完成

            # 进行评教操作
            while True:
                if not find_teachers(driver) and in_fact(driver):
                    change_page(driver)
                    time.sleep(2)
                elif not find_teachers(driver) and not in_fact(driver):
                    break
                else:
                    grade(driver)
                    time.sleep(0.2)
                    check_submit(driver)
                    time.sleep(1)

            print(f"账号{username}评教结束")
        except MaxRetryError as e:
            print(f"账号{username}连接失败，正在尝试重启...")
            print(e)
            continue  # 继续尝试下一个账号
        except Exception as e:
            print(f"账号{username}发生错误：{e}")
            break  # 如果需要，可以跳过当前账号，处理下一个账号
        finally:
            # 每个账号评教完成后关闭WebDriver实例
            driver.quit()
    print("所有账号评教完成")


if __name__ == '__main__':
    main()
