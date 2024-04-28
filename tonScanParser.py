from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import time
import pandas as pd
import numpy as np
import os

def data_to_excel(currentExcelData):
    df = pd.DataFrame(currentExcelData)
    file_name = "Result.xlsx"
    if not os.path.isfile(file_name):
        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=time.strftime("%H.%M-%d %m %Y"), index=False) 
    else:   
        with pd.ExcelWriter("Result.xlsx", engine='openpyxl', mode="a") as writer:
            df.to_excel(writer, sheet_name=time.strftime("%H.%M-%d %m %Y"), index=False)
    print("Данные успешно записаны в excel!")

def scroll_to_end(driver):
    last_height = driver.execute_script("return document.body.scrollHeight") 
    while True: 
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(5)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            print("Прокрутка завершена")
            break
        last_height = new_height
        print("Появился новый контент, прокручиваем дальше")

def comparison_sheets(excelFile, currentExcelData):
    xl = pd.ExcelFile(excelFile)
    if len(xl.sheet_names) > 0:
        df2 = pd.read_excel(excelFile, sheet_name=xl.sheet_names[-1], header=None)
        df2_matrix = df2.values.tolist()[1:]
        for i in range(1, len(currentExcelData)):
            address_current = currentExcelData[i][1]
            np_df2 = np.array(df2_matrix)
            index_to_last = np.argwhere(address_current == np_df2)
            if index_to_last.size > 0:
                row_index = index_to_last[0][0]
                diff = 0
                if len(currentExcelData[i]) == 3:    
                    now = currentExcelData[i][2] if currentExcelData[i][2] not in [' ', 'Токен только появился'] else 0    
                    if len(df2_matrix[row_index]) == 4:
                        was = df2_matrix[row_index][2] if df2_matrix[row_index][2] not in [' ', 'Токен только появился'] else 0
                        diff = now - was
                    else:
                        diff = 0
                currentExcelData[i].append(diff)
            else:
                currentExcelData[i].append("Токен только появился")

def main():
    user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 YaBrowser/24.1.0.0 Safari/537.36"
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument(f"user-agent={user_agent}")
    options.add_argument("--disable-blink-features=AutomationControlled")

    s = Service(executable_path=r"C:\\DriversSelenium\\chromedriver\\chromedriver-win64\\chromedriver.exe")
    driver = webdriver.Chrome(service=s, options=options)

    try:
        # main_link = "https://tonscan.org"
        link = "https://tonscan.org/ru/whales"
        driver.get(link)
        driver.implicitly_wait(20)
        scroll_to_end(driver)
        soup = BeautifulSoup(driver.page_source, 'lxml')
        table_rows = soup.find('table', class_='ui-table').find_all('tr')
        arrResult = [["#", "Адрес", "Баланс", "Разница между прошлым запуском"]]
        for row in range(1, len(table_rows)):
            cols = table_rows[row].find_all('td')
            cols = [f"{ele.find('a').get('href').split('/')[-1]}" if ele.find('a') else int(float(ele.text.replace("TON", "").strip().replace(u'\xa0','').replace(',', '.'))) for ele in cols]
            arrResult.append([ele for ele in cols if ele])

        comparison_sheets("Result.xlsx", arrResult)
        data_to_excel(arrResult)
        print(len(table_rows))
    except Exception as ex:
        print(ex)
    finally:
        driver.close()
        driver.quit()

if __name__ == "__main__":
    main()
