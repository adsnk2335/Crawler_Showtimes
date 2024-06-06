import random
import requests as req
from bs4 import BeautifulSoup
import re
import os
import openpyxl as xl
import time



def movie_showtime(movie_info_list):
    path = r"C:\Users\ADSNK2335\Desktop\movie_showtime\\"
    filename = "movie_showtime.xlsx"
    if not os.path.exists(path):
        os.makedirs(path)
    if os.path.exists(path+filename):
        wb = xl.load_workbook(path+filename)
        ws = wb.active
    else:
        wb = xl.Workbook()
        ws = wb.active
        ws.append(["電影名稱", "上映日期", "片長", "類型", "劇情介紹", "電影預告", "戲院連結"])
    for details in movie_info_list:
        ws.append(details)
    wb.save(path + filename)

def minshan_movie_info():
    home_page = "http://www.minshen.com.tw/"
    url = "http://www.minshen.com.tw/movies.php"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"}
    resp = req.get(url, headers=headers)
    if resp.status_code == 200:
        resp.encoding = "UTF-8"
        soup = BeautifulSoup(resp.text, "html.parser")
        # print(soup.prettify()) # 輸出排版後的HTML程式碼
        movie_info = soup.findAll("div", class_="movie-info")
        movie_info_list = []
        cnt = 0
        t = 0
        rest = range(1, 6)
        for info in movie_info:
            new_url = home_page+info.div.a["href"]
            new_resp = req.get(new_url, headers=headers)
            if new_resp.status_code == 200:
                new_resp.encoding = "UTF-8"
                new_soup = BeautifulSoup(new_resp.text, "html.parser")
                title = new_soup.find("p", class_="chinese-title")  # 電影名稱
                show_info = new_soup.find("div", class_="playdate")
                show_date_split = re.split("\|", show_info.text)
                show_date = show_date_split[0][5:15].strip()    # 上映日期
                duration = show_date_split[1][4:].strip()   # 片長
                style = new_soup.find("div", class_="movie-more-information")
                type_td = style.table.tr.findAll("td")
                typee = type_td[1].text # 類型
                movie_desc = new_soup.find("div", class_="movie-description")
                desc = movie_desc.findAll("p")
                description = desc[1].text  # 劇情介紹
                trai_url = new_soup.find("div", class_="movie-trailer")
                trailer = trai_url.iframe["src"]    # 預告片
                movie_info_list.append([title.text, show_date, duration, typee, description, trailer])
                cnt += 1
                time.sleep(random.choice(rest))
                exec_t = time.perf_counter()
                print(f"休息一下，我不是機器人，下載第{cnt}個資訊中，執行了{exec_t - t:.2f}秒")
                t = exec_t
            else:
                print("找不到該網頁")
        movie_showtime(movie_info_list)
    else:
        print("找不到時刻表網頁")


if __name__ == "__main__":
    minshan_movie_info()
