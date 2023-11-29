#!/usr/bin/env python
# coding: utf-8



#크롤링시 필요한 라이브러리 불러오기
from bs4 import BeautifulSoup
import requests
import re
import datetime
from tqdm import tqdm
import sys
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from io import BytesIO
import os
from datetime import datetime
import pandas as pd

# 페이지 url 형식에 맞게 바꾸어 주는 함수 만들기
  #입력된 수를 1, 11, 21, 31 ...만들어 주는 함수
def makePgNum(num):
    if num == 1:
        return num
    elif num == 0:
        return num+1
    else:
        return num+9*(num-1)

# 크롤링할 url 생성하는 함수 만들기(검색어, 크롤링 시작 페이지, 크롤링 종료 페이지)

def makeUrl(search, start_pg, end_pg):
    if start_pg == end_pg:
        start_page = makePgNum(start_pg)
        url = "https://search.naver.com/search.naver?where=news&sm=tab_pge&query=" + search + "&start=" + str(start_page) + '&sort=1'
        print("생성url: ", url)
        url = [url]
        return url
    else:
        urls = []
        for i in range(start_pg, end_pg + 1):
            page = makePgNum(i)
            url = "https://search.naver.com/search.naver?where=news&sm=tab_pge&query=" + search + "&start=" + str(page)  + '&sort=1'
            urls.append(url)
        print("생성url: ", urls)
        return urls    

# html에서 원하는 속성 추출하는 함수 만들기 (기사, 추출하려는 속성값)
def news_attrs_crawler(articles,attrs):
    attrs_content=[]
    for article in articles:
        # 각 atricle은 bs4의 무슨 객체 (잘모름)
        # attrs라는 인스턴스에 각 요소를 dict로 저장하고있음
        attrs_content.append(article.attrs[attrs])
    return attrs_content

# ConnectionError방지
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/98.0.4758.102"}

#html생성해서 기사크롤링하는 함수 만들기(url): 링크를 반환
def articles_crawler(url):
    #html 불러오기
    original_html = requests.get(url,headers=headers) # url에서 html내용을 가져오기
    html = BeautifulSoup(original_html.text, "html.parser") # html 내용을 추출하기 쉽게 파싱객체를 선언

    url_naver = html.select("div.group_news > ul.list_news > li div.news_area > div.news_info > div.info_group > a.info") 
    # 파싱 객체에서 원하는 부분 : 해당 페이지의 뉴스기사들의 info (각 기사의 url이 저장되어있음)
    # 각 info를 추출하여 List로 반환
    url = news_attrs_crawler(url_naver,'href')
    # 각 뉴스기사의 url 추출
    
    return url

def get_thumbnail_from_URL(url) : #url을 주면 해당 url에서 썸네일 Image형태로 가져오기
    news = requests.get(url, headers=headers)
    news_html = BeautifulSoup(news.text, "html.parser")

    # content 추출
    content = news_html.select("article#dic_area")
    if not content:
        content = news_html.select("#articeBody")

    # content 내의 모든 이미지 태그 찾기
    image_tags = content[0].find_all('img') if content else []
    if image_tags == [] : return None
    img_tag = image_tags[0] # 첫번째 이미지만 사용 (썸네일)
    
    img_url = img_tag.get('data-src')
    if img_url == None : img_url = img_tag.get('src') # Lazy loading이 아니어서 src에 있을 경우
    if img_url :
        img_data = requests.get(img_url).content
        img_file = BytesIO(img_data)
        img = Image(img_file)
        
    return img

def convert_datetime_format(dt_str):
    # '오후'/'오전'을 PM/AM으로 변환
    dt_str = dt_str.replace('오후', 'PM').replace('오전', 'AM')
    # 점(.)을 공백으로 변환하고, 불필요한 공백 제거
    dt_str = dt_str.replace('.', ' ').strip()
    # 연속된 공백을 하나의 공백으로 변환
    dt_str = ' '.join(dt_str.split())

    # 사용자 정의 형식에 맞춰 datetime 객체로 변환
    dt_obj = datetime.strptime(dt_str, '%Y %m %d %p %I:%M')

    # 원하는 형식의 문자열로 변환
    return dt_obj.strftime('%Y-%m-%d %H:%M')

def convert_to_datetime(date_str) :
    try : return pd.to_datetime(date_str)
    except : 
        converted_str = convert_datetime_format(date_str)
        return pd.to_datetime(converted_str)


#####뉴스크롤링 시작#####

#검색어 입력
search_input = input("검색할 키워드를 입력해주세요 (여러 검색어의 경우 , 으로 구분하세요.):")
#검색 시작할 페이지 입력
page = int(input("\n크롤링할 시작 페이지를 입력해주세요. ex)1(숫자만입력):")) # ex)1 =1페이지,2=2페이지...
print("\n크롤링할 시작 페이지: ",page,"페이지")   
#검색 종료할 페이지 입력
page2 = int(input("\n크롤링할 종료 페이지를 입력해주세요. ex)1(숫자만입력):")) # ex)1 =1페이지,2=2페이지...
print("\n크롤링할 종료 페이지: ",page2,"페이지")   

# 여러 검색어에 대해 시행
search_List = search_input.split(',')
for search in search_List :

    # naver url 생성
    urls = makeUrl(search,page,page2)

    #뉴스 크롤러 실행
    news_titles = []
    news_urlList_for_page =[]
    news_contents =[]
    news_dates = []
    news_thumbnails = []

    for url in urls:
        # 네이버 뉴스 기사의 객 페이지 url에서 뉴스기사들의 url 추출하여 append
        crawled_url = articles_crawler(url)
        news_urlList_for_page.append(crawled_url)


    #제목, 링크, 내용 1차원 리스트로 꺼내는 함수 생성
    def make_newsURLlist(newsURL_list, urlList_eachPage):
        for urlList in urlList_eachPage:
            for url in urlList:
                newsURL_list.append(url)
        return newsURL_list

        
    #제목, 링크, 내용 담을 리스트 생성
    newsURL_list = []

    #1차원 리스트로 만들기(내용 제외)
    make_newsURLlist(newsURL_list,news_urlList_for_page)


    #NAVER 뉴스만 남기기
    # 다른 url들은 쓸모없는 다른 뉴스 사이트 url이므로 날린다
    naver_urls = []
    for i in tqdm(range(len(newsURL_list))):
        if "news.naver.com" in newsURL_list[i]:
            naver_urls.append(newsURL_list[i])
        else:
            pass


    # 뉴스 내용 크롤링

    for url in tqdm(naver_urls):
        #각 기사 html 주소에 들어가서 뉴스 내용 따오기
        news = requests.get(url,headers=headers)
        news_html = BeautifulSoup(news.text,"html.parser")

        # 뉴스 제목 가져오기
        title = news_html.select_one("#ct > div.media_end_head.go_trans > div.media_end_head_title > h2")
        if title == None:
            title = news_html.select_one("#content > div.end_ct > div > h2")
            
        # 기사 썸네일 가져오기   
        image = get_thumbnail_from_URL(url)
        news_thumbnails.append(image)
        
        # 뉴스 본문 가져오기
        content = news_html.select("article#dic_area")
        if content == []:
            content = news_html.select("#articeBody")

        # 기사 텍스트만 가져오기
        # list합치기
        content = ''.join(str(content)) # content의 내용 List를 다시 문자열로 합치기

        # html태그제거 및 텍스트 다듬기
        pattern1 = '<[^>]*>'
        title = re.sub(pattern=pattern1, repl='', string=str(title))
        content = re.sub(pattern=pattern1, repl='', string=content)
        pattern2 = """[\n\n\n\n\n// flash 오류를 우회하기 위한 함수 추가\nfunction _flash_removeCallback() {}"""
        content = content.replace(pattern2, '')

        news_titles.append(title)
        news_contents.append(content)

        try:
            html_date = news_html.select_one("div#ct> div.media_end_head.go_trans > div.media_end_head_info.nv_notrans > div.media_end_head_info_datestamp > div > span")
            news_date = html_date.attrs['data-date-time']
        except AttributeError:
            news_date = news_html.select_one("#content > div.end_ct > div > div.article_info > span > em")
            news_date = re.sub(pattern=pattern1,repl='',string=str(news_date))
        # 날짜 가져오기
        news_dates.append(news_date)

    print("검색된 기사 갯수: 총 ",(page2+1-page)*10,'개')
    print('news_title: ',len(news_titles))
    print('news_url: ',len(naver_urls))
    print('news_contents: ',len(news_contents))
    print('news_dates: ',len(news_dates))
    print('news_thumbnails : ', len(news_thumbnails))

    #데이터 프레임 만들기
    news_df = pd.DataFrame({'date':news_dates,'title':news_titles,'link':naver_urls,'content':news_contents, 'thumbnail':news_thumbnails})

    #중복 행 지우기
    news_df = news_df.drop_duplicates(keep='first',ignore_index=True)
    print("중복 제거 후 행 개수: ",len(news_df))

    # 뉴스 게시 일시 (date) 를 시계열 데이터로 전환
    news_df['date'] = news_df['date'].apply(lambda date_str : convert_to_datetime(date_str))
    # 뉴스 게시일의 내림차순으로 정렬하고 (최신순) 정렬한 뒤엔 다시 문자열로 바꾸기
    news_df = news_df.sort_values(by='date', ascending = False)
    news_df = news_df.reset_index(drop=True)
    news_df['date'] = news_df['date'].apply(lambda news_date : str(news_date))

    # 데이터를 엑셀에 넣기 위해 thumbnail이 없는 부분을 따로 분리
    news_df_without_thumbnail = news_df.drop('thumbnail', axis=1)


    # Excel 워크북 및 워크시트 생성
    wb = Workbook()
    ws = wb.active

    # 셀의 크기
    cell_width = 20
    cell_height = 150

    # 각 데이터에 맞는 컬럼을 지정해주기 위한 알파벳 리스트 선언 (A컬럼, B컬럼, ...)
    column_alphabet = [chr( ord('B')+i ) for i in range(len(news_df))]

    for row_num, row in news_df_without_thumbnail.iterrows() :
        # 각 row를 excel시트에 추가
        row_keys = row.keys()
        for i in range(len(row)) :
            cell_ref = f'{column_alphabet[i]}{row_num+1}'
            ws[cell_ref] = row[row_keys[i]]
            # 알맞는 컬럼, row에 dataframe의 요소 (title, link, ...) 할당

    # 이미지를 Excel에 삽입
    for i, img in enumerate(news_df['thumbnail']):
        if img == None : continue
        # 이미지 크기 조절
        img.width = cell_width*8
        img.height = cell_height*1.3

        # 셀 크기 조정
        ws.row_dimensions[i + 1].height = cell_height
        ws.column_dimensions['A'].width = cell_width

        # 이미지 삽입
        cell_ref = f'A{i + 1}'  # 예: 'A1', 'A2', ...
        ws.add_image(img, cell_ref)

    # Excel 파일 저장
    now = datetime.now() 
    nowtime = now.strftime('%Y%m%d_%H시%M분%S초')
    excel_filename = f'{nowtime}-{search}.xlsx'
    dest_folder = './crawl_data'
    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)
    wb.save(dest_folder +'/'+ excel_filename)