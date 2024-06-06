import datetime
import time
import requests
from bs4 import BeautifulSoup
from openpyxl.workbook import Workbook

# 建立一個工作表
wb = Workbook()
ws = wb.active

# 寫入標題
gossiping = ["Date", "Push", "Author", "Title", "Post_Url"]
ws.append(gossiping)

headers = {"cookie": "over18=1",
           "User-Agent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) ' 
                         'AppleWebKit/537.36 (KHTML, like Gecko) ' 
                         'Chrome/125.0.0.0 Safari/537.36'}
# 先存第一頁並找出上一頁的 index 頁碼
res1 = requests.get('https://www.ptt.cc/bbs/Gossiping/index.html', headers=headers)
soup1 = BeautifulSoup(res1.content, 'html.parser')
articles1 = soup1.find_all('div', attrs={"class": "r-ent"})
articles1.reverse()  # 按順序排列

for a in articles1:
    article1 = []
    date = a.find("div", attrs={"class": "date"})
    # print(date.text.strip())
    article1.append(date.text.strip())
    push = a.find("div", class_="nrec")
    # print(push.text.strip())
    article1.append(push.text.strip())
    author = a.find("div", class_="author")
    # print(author.text.strip())
    article1.append(author.text.strip())
    title = a.find("div", class_="title")
    # print(title.text.strip())
    article1.append(title.text.strip())
    postUrl = a.find("a")
    if postUrl is None:
        # print("-")
        article1.append("-")  # 刪文的網址用"-"代替
    else:
        # print(f"https://www.ptt.cc{postUrl["href"]}")
        article1.append(f"https://www.ptt.cc{postUrl["href"]}")
    # print(f"{title}　{author}　{push}　{date}")
    ws.append(article1)

nextLink = str(soup1.find("a", string="‹ 上頁").get("href").split(".")[0].replace("/bbs/Gossiping/index", ""))

# 第二頁之後有 index 頁碼可以用迴圈替代
while True:
    res2 = requests.get(f'https://www.ptt.cc/bbs/Gossiping/index{nextLink}.html', headers=headers)
    soup2 = BeautifulSoup(res2.content, 'html.parser')
    articles2 = soup2.find_all('div', attrs={"class": "r-ent"})
    todayDate = str(datetime.date.today()).strip()[-1]
    articles2.reverse()  # 按順序排列
    for a in articles2:
        article2 = []
        if todayDate == a.find("div", attrs={"class": "date"}).text.strip()[-1]:  # 比對今天日期，只抓取今天日期
            date = a.find("div", attrs={"class": "date"})
            # print(date.text.strip())
            article2.append(date.text.strip())
            push = a.find("div", class_="nrec")
            # print(push.text.strip())
            article2.append(push.text.strip())
            author = a.find("div", class_="author")
            # print(author.text.strip())
            article2.append(author.text.strip())
            title = a.find("div", class_="title")
            # print(title.text.strip())
            article2.append(title.text.strip())
            postUrl = a.find("a")
            if postUrl is None:
                article2.append("-")  # 刪文的網址用"-"代替
            else:
                article2.append(f"https://www.ptt.cc{postUrl["href"]}")
            ws.append(article2)
        else:  # 沒有今天日期就跳出
            break
    if todayDate != soup2.find("div", attrs={"class": "date"}).text.strip()[-1]:  # 在比對日期，相符就再抓上一頁，不相符就存檔
        wb.save(f"Gossiping-{datetime.date.today()}.xlsx")
        print("End")
        break
    else:
        time.sleep(1)
        nextLink = int(nextLink) - 1
