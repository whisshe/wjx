import time

from openpyxl import load_workbook
from selenium import webdriver

wb = load_workbook('题库模板.xlsx')
sheet = wb['sheet']
print(sheet.rows)
all_row_values = []
count = 1
for row in sheet.rows:
    row_values = [count]
    for cell in row:
        row_values.append(cell.value)
    # print(row_values)
    if row_values[1] != '题型':
        all_row_values.append(row_values)
        count += 1
print(all_row_values)
from selenium.common.exceptions import NoSuchElementException

All = all_row_values

# base_url = 'https://ks.wjx.top/wjx/Join/VerifyPasswordMobile2.aspx?q=82852430&pwx=&returnUrl=%2fm%2f82852430.aspx'
base_url = 'https://ks.wjx.top/wjx/Join/VerifyPasswordMobile2.aspx?q=82416901&pwx=&returnUrl=%2fm%2f82416901.aspx'
chrome_path = 'F:/Python/wjx/chromedriver.exe'
option_dict = {'A': 1, 'B': 2, 'C': 3, 'D': 4, 'E': 5}
driver = webdriver.Chrome(executable_path=chrome_path)
driver.get(base_url)
phone_number_input = driver.find_element_by_name('txtPassword')
phone_number_input.send_keys('123456')
next_step = driver.find_element_by_id('btnContinue')
next_step.click()
question_count = 1

# while True:
for Count in range(16):
    question_retry = True
    while question_retry:
        try:
            question = driver.find_element_by_xpath('//*[@id="div{}"]/div[1]'.format(question_count))
            question_retry = False
        except NoSuchElementException as e:
            question_count += 1
    topic = driver.find_element_by_xpath('//*[@id="div{}"]'.format(question_count))
    question_id = topic.get_attribute('topic')
    next_question = driver.find_element_by_id('divNext')
    print(question.text)
    for i in All:
        if int(question_id) == i[0]:
            answer = i[8]
            print(answer)
            if len(answer) == 1:
                answer_id = option_dict[answer]
                answer_element = driver.find_element_by_xpath(
                    '//*[@id="div{}"]/div[2]/div[{}]/span/a'.format(question_count, answer_id))
                answer_element.click()
            elif len(answer) > 1:
                for a in answer:
                    if a != ' ':
                        answer_id = option_dict[a]
                        answer_element = driver.find_element_by_xpath(
                            '//*[@id="div{}"]/div[2]/div[{}]/span/a'.format(question_count, answer_id))
                        answer_element.click()
            question_count += 1
            next_question.click()
time.sleep(10)
commit = driver.find_element_by_xpath('//*[@id="divSubmit"]/div[1]')
print(commit)
time.sleep(2)
commit.click()
driver.close()
