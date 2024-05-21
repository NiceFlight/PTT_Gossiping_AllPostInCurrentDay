import datetime
import time
import requests
from bs4 import BeautifulSoup
from openpyxl.workbook import Workbook

headers = {"cookie": "over18=1",
           "User-Agent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) ' 
                         'AppleWebKit/537.36 (KHTML, like Gecko) ' 
                         'Chrome/125.0.0.0 Safari/537.36'}

res1 = requests.get('https://www.ptt.cc/bbs/Gossiping/index.html', headers=headers)
soup1 = BeautifulSoup(res1.content, 'html.parser')
nextLink1 = str(soup1.find("a", string="‹ 上頁").get("href").split(".")[0].replace("/bbs/Gossiping/index", ""))

res2 = requests.get(f'https://www.ptt.cc/bbs/Gossiping/index{nextLink1}.html', headers=headers)
soup2 = BeautifulSoup(res2.content, 'html.parser')
nextLink2 = soup2.find("a", string="下頁 ›").get("href").split(".")[0].replace("/bbs/Gossiping/index", "")

wb = Workbook()
ws = wb.active
gossiping = ["Date", "Push", "Author", "Title", "Post_Url"]
ws.append(gossiping)

while True:
    url = f"https://www.ptt.cc/bbs/Gossiping/index{nextLink2}.html"
    res3 = requests.get(url, headers=headers)
    soup3 = BeautifulSoup(res3.content, 'html.parser')
    date1 = soup3.find("div", class_="date").text.strip()[-2:]
    todayDate = str(datetime.date.today()).split("-")[-1]
    if date1 == todayDate:
        # print(url)
        articles = soup3.find_all('div', attrs={"class": "r-ent"})
        for a in articles:
            article = []
            date = a.find("div", attrs={"class": "date"})
            # print(date.text.strip())
            article.append(date.text.strip())
            push = a.find("div", class_="nrec")
            # print(push.text.strip())
            article.append(push.text.strip())
            author = a.find("div", class_="author")
            # print(author.text.strip())
            article.append(author.text.strip())
            title = a.find("div", class_="title")
            # print(title.text.strip())
            article.append(title.text.strip())
            postUrl = a.find("a")
            if postUrl is None:
                article.append("-")
            else:
                article.append(f"https://www.ptt.cc{postUrl["href"]}")
            # print(f"{title}　{author}　{push}　{date}")
            ws.append(article)
    else:
        wb.save(f"Gossiping-{datetime.date.today()}.xlsx")
        print("end")
        break
    # print(url)
    time.sleep(1)
    nextLink2 = int(nextLink2) - 1
