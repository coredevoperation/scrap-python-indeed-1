import time
from datetime import datetime
import csv
import pandas as pd
import gspread
from selenium.webdriver.chrome.service import Service
from selenium import webdriver  
from selenium_stealth import stealth
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver import ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

HOMEPAGE = "https://se.indeed.com/jobs?"
CHROME_DRIVER_PATH = 'c:/WebDrivers/chromedriver.exe'
# proxy_options = {
#     'proxy': {
#         'http': 'http://USERNAME:PASSWORD@proxy-server:8080',
#         'https': 'http://USERNAME:PASSWORD@proxy-server:8080',
#         'no_proxy': 'localhost:127.0.0.1'
#     }
# }
browser_options = webdriver.ChromeOptions()
browser_options.headless = False
driver = webdriver.Chrome(options = browser_options)

stealth(driver,
        languages=["en-US", "en"],
        vendor="Google Inc.",
        platform="Win32",
        webgl_vendor="Intel Inc.",
        renderer="Intel Iris OpenGL Engine",
        fix_hairline=True,
        )
def is_exist(array, job, company):
    for row in array:
        if row[2] == job and row[3] == company:
            return array.index(row)
    return -1

def filter_array(dataframe, result):
    if len(dataframe) == 0:
        return result
    total = dataframe

    for row in result:
        job = row[2]
        company = row[3]
        index = is_exist(total, job, company)
        if index != -1:
            total[index] = row
        else:
            total.append(row)
    return total

def getCompany(company):
    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[1])
    url = "https://www.allabolag.se/what/"
    arr = company.split(" ")
    url = url + arr[0]
    arr.pop(0)
    for i in arr:
        url = url + "%20" + i
    driver.get(url)

    turnover = ""
    phone = ""
    try:
        result = driver.find_element(By.CSS_SELECTOR, "div.tw-my-2")
        article = result.find_element(By.XPATH, ".//article")
        link = article.find_element(By.TAG_NAME, "a").get_attribute("href")
    except:
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        return turnover, phone
    
    driver.get(link)
    
    try:
        table = driver.find_element(By.CSS_SELECTOR, "table.figures-table")
        infos = table.find_elements(By.XPATH, ".//tbody/tr")
        for info in infos:
            if info.text.find("Omsättning") != -1:
                turnover = info.find_element(By.XPATH, ".//td[1]").text
    except:
        turnover = ""

    try:
        phone = driver.find_element(By.CSS_SELECTOR, "a.p-tel").text
    except:
        phone = ""

    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    return turnover, phone

def getWebsite(url):
    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[1])
    driver.get(url)

    try:
        link = driver.find_element(By.XPATH, "//li[@data-testid='companyInfo-companyWebsite']")
        website = link.find_element(By.XPATH, ".//a").get_attribute("href")
    except:
        print("No Link")
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    return website

def getData(url, job):
    driver.get(url)
    row_data = []
    try:
        jobtitle = driver.find_element(By.CSS_SELECTOR, "h1.jobsearch-JobInfoHeader-title").text
    except:
        jobtitle = ""

    try:
        company = driver.find_element(By.XPATH, "//div[@data-testid='inlineHeader-companyName']")
        company_name = company.text
        link = company.find_element(By.TAG_NAME, "a").get_attribute("href")
    except:
        company_name = ""
        link = ""

    try:
        address = driver.find_element(By.XPATH, "//div[@data-testid='inlineHeader-companyLocation']").text
    except:
        address = ""

    try:
        salary = driver.find_element(By.XPATH, "//div[@id='salaryInfoAndJobType']").text
    except:
        salary = ""

    try:
        website = getWebsite(link)
    except:
        website = ""

    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    row_data.append(job)
    row_data.append(dt_string)
    row_data.append(jobtitle)
    row_data.append(company_name)
    turnover, phone = getCompany(company_name)
    row_data.append(website)
    row_data.append(phone)
    row_data.append(turnover)
    row_data.append(address)
    row_data.append(salary)
    row_data.append(url)
    return row_data

def parselist(job):
    url = HOMEPAGE + "q=" + job
    driver.get(url)
    try:
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler"))).click()
    except:
        print("No but")
    url_list = []
    while 1:
        list = driver.find_elements(By.XPATH, ".//div[@id='mosaic-provider-jobcards']/ul/li")
        print(len(list))
        for item in list:
            try:
                heading = item.find_element(By.CSS_SELECTOR, "h2.jobTitle")
                url = heading.find_element(By.XPATH, ".//a").get_attribute("href")
                url_list.append(url)
            except:
                print("No Header")

        try:
            next = driver.find_element(By.XPATH, "//a[@data-testid='pagination-page-next']")
            actions = ActionChains(driver)
            actions.move_to_element(next).perform()
            print(url_list)
            next.click()
        except:
            print("No more")
            break
    result = []
    for url in url_list:
        row = getData(url, job)
        result.append(row)
    print(result)
    return result

def main():
    SHEET_ID = '1yuK88LxA3sx0o3oy6zCLMMMak29gPYmEcTlIowwnAQw'
    SHEET_NAME = 'se_indeed'
    gc = gspread.service_account('key.json')
    spreadsheet = gc.open_by_key(SHEET_ID)
    worksheet = spreadsheet.worksheet(SHEET_NAME)
    
    csv_file = open('list.csv', 'r', encoding='utf8')
    reader = csv.reader(csv_file)
    for job in reader:
        result = parselist(job[0])
        dataframe = worksheet.get_all_values()
        dataframe.pop(0)
        res = filter_array(dataframe, result)
        spreadsheet.values_clear(SHEET_NAME + "!A2:J")
        worksheet.update("A2:J",res)
    driver.close()

if __name__ == '__main__':
    main()    
