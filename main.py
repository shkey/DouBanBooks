import requests
from bs4 import BeautifulSoup
import xlwings as xw


def get_all_books():
    book_list = []
    url = "https://book.douban.com/top250?start="
    for i in range(0, 11):
        req = requests.get(url + str(i * 25))
        soup = BeautifulSoup(req.text, "lxml")
        book_items = soup.select("#content > div > div.article > div > table")
        book_list += book_items
        print(i)
    print(len(book_list))
    return book_list


def parse_book_items(books):
    book_item_list = []
    for item in books:
        book_item = {"img": "", "name": "", "author": "", "press": "", "year": "", "price": "", "score": "",
                     "evaluate": "", "summary": ""}
        soup = BeautifulSoup(str(item), "lxml")
        book_item["img"] = soup.select("a.nbg > img")[0]["src"]
        book_item["name"] = soup.select("td > div.pl2")[0].get_text(strip=True)
        information = soup.select("td > p.pl")[0].get_text(strip=True).split("/")
        book_item["author"] = information[:-3]
        book_item["press"] = information[-3:-2]
        book_item["year"] = information[-2:-1]
        book_item["price"] = information[-1:]
        book_item["score"] = soup.select("td > div.star.clearfix > span.rating_nums")[0].get_text(
            strip=True)
        book_item["evaluate"] = soup.select("td > div.star.clearfix > span.pl")[0].get_text(strip=True).strip("()").strip()
        if len(soup.select("td > p.quote > span.inq")) is not 0:
            book_item["summary"] = soup.select("td > p.quote > span.inq")[0].get_text(strip=True)
        book_item_list.append(book_item)
    print(book_item_list)
    return book_item_list


books = get_all_books()
book_item_list = parse_book_items(books)
app = xw.App(visible=True, add_book=False)
wb = app.books.add()
data_range = wb.sheets['sheet1'].range('A1')
data_range.value = ["书名", "作者", "出版社", "出版日期", "定价", "评分", "评价数量", "一句", "封面图"]
for x in range(0, 250):
    data_range = wb.sheets['sheet1'].range("A" + str(x + 2))
    name = book_item_list[x]["name"]
    author = book_item_list[x]["author"]
    press = book_item_list[x]["press"]
    year = book_item_list[x]["year"]
    price = book_item_list[x]["price"]
    score = book_item_list[x]["score"]
    evaluate = book_item_list[x]["evaluate"]
    summary = book_item_list[x]["summary"]
    img = book_item_list[x]["img"]
    data_range.value = [str(name).strip(), " ".join(author).strip(), " ".join(press).strip(), " ".join(year).strip(),
                        " ".join(price).strip(), score.strip(), evaluate.strip(), summary.strip(), img.strip()]
wb.save(r'DouBanBooks.xlsx')
wb.close()
app.quit()
