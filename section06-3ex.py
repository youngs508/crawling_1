# Section06-3
# Selenium
# Selenium 사용 실습(3) - 실습 프로젝트(2)

# selenium 임포트
from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
# 엑셀 처리 임포트
import xlsxwriter
# 이미지 바이트 처리
from io import BytesIO
import urllib.request as req

chrome_options = Options()
chrome_options.add_argument("--headless")  # 브라우저가 실행 안 되는 옵션

# 엑셀 처리 선언
workbook = xlsxwriter.Workbook(
    "C:/Users/young/OneDrive/crawling/python_crawl/crawling_result.xlsx")

# 워크 시트
worksheet = workbook.add_worksheet()

# webdriver 설정(Chrome, Firefox 등) - Headless 모드
# browser = webdriver.Chrome('./webdriver/chrome/chromedriver.exe', options=chrome_options)

# webdriver 설정(Chrome, Firefox 등) - 일반 모드
browser = webdriver.Chrome('./webdriver/chrome/chromedriver.exe')

# 크롬 브라우저 내부 대기
browser.implicitly_wait(5)

# 브라우저 사이즈
browser.set_window_size(1200, 800)  # maximize_window(), minimize_window()

# 페이지 이동
browser.get('https://www.amazon.com/ref=nav_logo')

# 1차 페이지 내용
# print('Before Page Contents : {}'.format(browser.page_source))

# 제조사별 더 보기 클릭1
# Explicitly wait

WebDriverWait(browser, 5).until(EC.presence_of_element_located(
    (By.XPATH, '//*[@id="nav-hamburger-menu"]'))).click()

# 제조사별 더 보기 클릭2
# Implicitly wait
# time.sleep(2)
# browser.find_element_by_xpath('//*[@id="dlMaker_simple"]/dd/div[2]/button[1]').click()

# 원하는 모델 카테고리 클릭
WebDriverWait(browser, 5).until(EC.presence_of_element_located(
    (By.XPATH, '//*[@id="hmenu-content"]/ul[1]/li[7]'))).click()

WebDriverWait(browser, 5).until(EC.presence_of_element_located(
    (By.XPATH, '//*[@id="hmenu-content"]/ul[5]/li[4]'))).click()

# 2차 페이지 내용
# print('After Page Contents : {}'.format(browser.page_source))

# 2초간 대기
time.sleep(2)

# 현재 페이지
cur_page = 1

# 크롤링 페이지 수
target_crawl_num = 5

# 엑셀 행 수
ins_cnt = 1

while cur_page <= target_crawl_num:
    # n = 1
    # bs4 초기화
    soup = BeautifulSoup(browser.page_source, 'html.parser')

    # 소스코드 정리
    # print(soup.prettify)

    # 메인 상품 리스트 선택
    pro_list = soup.select(
        'span.rush-component.s-latency-cf-section > div:nth-child(2) > div.sg-col-4-of-12')

    # 상품 리스트 확인
    # print(pro_list)

    # 페이지 번호 출력
    print('****** Current Page : {}'.format(cur_page), '******')
    print()

    # 필요 정보 추출
    for v in pro_list:
        # 임시 출력
        # print(v)
        # if n > 33:
        #     break
        # if v.find_all('div', class_="a-section a-spacing-medium"):
        # if v.find_all('div', class_="sg-col-4-of-12"):

        # 상품명, 이미지, 가격
        prod_name = v.select(
            'div.sg-col-inner a.a-link-normal.a-text-normal > span')[0].text.strip()
        prod_price = v.select(
            'div.sg-col-inner span.a-price > span:nth-child(1)')[0].text.strip()

        # 이미지 요청 후 바이트 변환
        img_data = BytesIO(req.urlopen(v.select(
            'div.sg-col-inner div.a-section.aok-relative.s-image-square-aspect > img')[0]['src']).read())

        # 엑셀 저장(텍스트)
        worksheet.write('A%s' % ins_cnt, prod_name)
        worksheet.write('B%s' % ins_cnt, prod_price)

        # 엑셀 저장(이미지)
        worksheet.insert_image('C%s' % ins_cnt, prod_name, {
                               'x_scale': 0.1, 'y_scale': 0.1, 'image_data': img_data})

        ins_cnt += 1
        # print(v.select(
        #     'div.sg-col-inner a.a-link-normal.a-text-normal > span')[0].text.strip())
        # print(v.select(
        #     'div.sg-col-inner div.a-section.aok-relative.s-image-square-aspect > img')[0]['src'])
        # print(v.select(
        #     'div.sg-col-inner span.a-price > span:nth-child(1)')[0].text.strip())
        # else:
        #     break

        # 이 부분에서 엑셀 저장(파일, DB 등)
        # CODE
        # CODE
        # n += 1
        # print()

    print()

    # 페이지 별 스크린 샷 저장
    browser.get_screenshot_as_file('C:/amazon_page{}.png'.format(cur_page))

    # 페이지 증가
    cur_page += 1

    if cur_page > target_crawl_num:
        print('Crawling Succeed!')
        break
    # 페이지 이동 클릭
    WebDriverWait(browser, 2).until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, 'ul.a-pagination > li.a-last'))).click()

    # BeautifulSoup 인스턴스 삭제
    del soup

    # 3초간 대기
    time.sleep(3)

# 브라우저 종료
browser.close()

# 엑셀 파일 닫기
workbook.close()
