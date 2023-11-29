from bs4 import BeautifulSoup
import requests
from openpyxl import Workbook

def makePgNum(num):
    if num == 1:
        return num
    elif num == 0:
        return num + 1
    else:
        return num + 9 * (num - 1)

def makeUrl(search, start_pg, end_pg, start_date, end_date):
    if start_pg == end_pg:
        start_page = makePgNum(start_pg)
        url = f"https://search.naver.com/search.naver?where=news&sm=tab_pge&query={search}&start={start_page}&sort=1&pd=3&ds={start_date}&de={end_date}"
        return [url]
    else:
        urls = []
        for i in range(start_pg, end_pg + 1):
            page = makePgNum(i)
            url = f"https://search.naver.com/search.naver?where=news&sm=tab_pge&query={search}&start={page}&sort=1&pd=3&ds={start_date}&de={end_date}"
            urls.append(url)
        return urls

def scrape_news_info(search, start_pg, end_pg, start_date, end_date):
    # Step 1: Make URLs
    urls = makeUrl(search, start_pg, end_pg, start_date, end_date)

    # Step 2: Scrape news information
    news_info = []
    for url in urls:
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        articles = soup.select('ul.list_news li')

        for article in articles:
            info_press = article.select_one('.info.press')
            if info_press:
                newspaper_info = info_press.get_text(strip=True)
                news_info.append(newspaper_info)

    return news_info

def save_to_excel(news_info, output_file):
    wb = Workbook()
    ws = wb.active

    # Write header
    ws.append(["Newspaper Info"])

    # Write data
    for info in news_info:
        ws.append([info])

    # Save to Excel file
    wb.save(output_file)

if __name__ == "__main__":
    search_query = input("Enter the search term: ")
    start_page_number = int(input("Enter the start page number: "))
    end_page_number = int(input("Enter the end page number: "))
    start_date = input("Enter the start date (e.g., 20220101): ")
    end_date = input("Enter the end date (e.g., 20221231): ")
    output_excel_file = "newspaper_info.xlsx"

    news_info = scrape_news_info(search_query, start_page_number, end_page_number, start_date, end_date)
    save_to_excel(news_info, output_excel_file)
