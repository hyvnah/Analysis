import os
import time
import datetime
import pandas as pd
import numpy as np
from tqdm import tqdm
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager


now = datetime.datetime.now()
yesterday = now - datetime.timedelta(hours=24)
now_year = str(now.year)
yes_year = str(yesterday.year)
now_month = str(now.month).zfill(2)
yes_month = str(yesterday.month).zfill(2)
now_day = str(now.day).zfill(2)
yes_day = str(yesterday.day).zfill(2)

# 현재 폴더로 커서 이동
os.chdir(os.path.dirname(os.path.realpath("__file__")))
file = '네이버_데이터랩_실시간키워드_' + now_year + now_month + now_day + '.xlsx'


# 크롤링 결과 담을 list
result_word_yes = {}
result_word_tod = {}
result_cnt_yes = []
result_cnt_tod = []
result_url_yes = []
result_url_tod = []

options = webdriver.ChromeOptions()
options.add_argument('headless') # 브라우저 Hidden
driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)


##실검sheet 코드(전일 06:00 - 23:50 / 10분단위)
for i in tqdm(range(6, 24)):
    for j in range(0, 60, 10):
        h = str(i).zfill(2)
        m = str(j).zfill(2)
        result_word_yes[f"{h}시{m}분"] = []
        try:
            url = f'https://datalab.naver.com/keyword/realtimeList.naver?datetime={yes_year}-{yes_month}-{yes_day}T{h}%3A{m}%3A00&entertainment=-2&groupingLevel=0&news=-2&sports=-2'
            driver.get(url)

            for k in range(2):
                for p in range(10):
                    data = driver.find_element_by_xpath(
                        '//*[@id="content"]/div/div[2]/div[2]/div[2]/div/div/ul[' + str(k + 1) + ']/li[' + str(p + 1) + ']').text
                    data = data.replace('\n', ' ')
                    result_word_yes[f"{h}시{m}분"].append(data)

        except:
            print(f"ERROR! {h}시{m}분")

        result_word_yes[f"비고({h}:{m})"] = [''] * 20
        time.sleep(3)
    time.sleep(3)

#실검sheet 코드(금일 00:00 - 06:00 / 10분단위)
for i in tqdm(range(0, 6)):
    for j in range(0, 60, 10):
        h = str(i).zfill(2)
        m = str(j).zfill(2)
        result_word_tod[f"{h}시{m}분"] = []
        try:
            url = f'https://datalab.naver.com/keyword/realtimeList.naver?datetime={now_year}-{now_month}-{now_day}T{h}%3A{m}%3A00&entertainment=-2&groupingLevel=0&news=-2&sports=-2'
            driver.get(url)

            for k in range(2):
                for p in range(10):
                    data = driver.find_element_by_xpath(
                        '//*[@id="content"]/div/div[2]/div[2]/div[2]/div/div/ul[' + str(k + 1) + ']/li[' + str(p + 1) + ']').text
                    data = data.replace('\n', ' ')
                    result_word_tod[f"{h}시{m}분"].append(data)

        except:
            print(f"ERROR! {h}시{m}분")

        result_word_tod[f"비고({h}:{m})"] = [''] * 20
        time.sleep(3)
    time.sleep(3)


# conver to dataframe
result_word_yes = pd.DataFrame(result_word_yes)
result_word_tod = pd.DataFrame(result_word_tod)

# concat yesterday & today
result_word = pd.concat([result_word_yes,result_word_tod], axis=1, sort = False)

# make file
result_word.to_excel(f'/Users/hyunah/Desktop/naver/네이버_데이터랩_실시간키워드_{now_year}{now_month}{now_day}.xlsx', index=False)
