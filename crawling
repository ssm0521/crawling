from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
from bs4 import BeautifulSoup
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import time
import datetime
import re

# 크롬 옵션 설정
options = webdriver.ChromeOptions()
options.add_argument('window-size=1920x1080')
options.add_argument("disable-gpu")

# 엑셀 파일 생성
now = datetime.datetime.now()
xlsx = Workbook()
list_sheet = xlsx.active
list_sheet.title = "output"
list_sheet.append(['Title','review', 'perform'])  # 열 헤더 추가


def crawl_text(driver):
    """자동차 제목, 연식 정보와 상세 설명을 크롤링하여 txt 파일 및 Excel 파일에 저장하는 함수"""
    try:
        # BeautifulSoup으로 페이지 내용을 파싱
        soup = BeautifulSoup(driver.page_source, "html.parser")

        # 자동차 제목과 연식 정보 크롤링
        car_divs = soup.find_all("div", {"class": "mocha_car"})

        with open("output.txt", "a", encoding="utf-8") as file:
            for car_div in car_divs:
                # 자동차 제목 추출
                title_tag = car_div.find("h3", {"class": "tit_mocha"}).find("a")
                title = title_tag.get_text(strip=True) if title_tag else "제목 없음"
                title = re.sub(r"\s*~\s*", "~", title)
                # 별점 크롤링
                review_div = soup.find("div", class_="area_review")
                rating = None
                if review_div:
                    rating_tag = review_div.find("strong", class_="tit_total")
                    if rating_tag:
                        rating_span = rating_tag.find("span", class_="txt_total")
                        if rating_span:
                            rating = rating_span.get_text(strip=True)

                if rating is None:
                    rating = "별점 없음"  # 기본값

                # 상세 설명 크롤링 (box_g box_perform mocha_on)
                target_div = soup.find("div", {"class": "box_g box_perform mocha_on", "data-content": "perform"})
                descriptions = []

                if target_div:
                    area_details = target_div.find_all("div", class_="area_detail")
                    for area in area_details:
                        # '직선주행' 항목 크롤링
                        straight_line = area.find("strong", class_="tit_g tit_line")
                        if straight_line:
                            label = straight_line.find("span", class_="img_mocha").get_text(strip=True)
                            description = area.find("p", class_="desc_detail").get_text(strip=True)
                            descriptions.append(f"직선주행: {description}")

                        # '곡선주행' 항목 크롤링
                        curved_line = area.find("strong", class_="tit_g tit_curve")
                        if curved_line:
                            label = curved_line.find("span", class_="img_mocha").get_text(strip=True)
                            description = area.find("p", class_="desc_detail").get_text(strip=True)
                            descriptions.append(f"곡선주행: {description}")

                # 결과 텍스트 파일에 저장
                car_info = f"차량 제목: {title}, 별점 : {rating},상세 설명: {' | '.join(descriptions)}"
                print(car_info)
                file.write(car_info + "\n")

                # Excel 파일에 저장
                list_sheet.append([title, rating, ' | '.join(descriptions)])

            print("텍스트 및 Excel 크롤링 완료 및 저장되었습니다.")

    except Exception as e:
        print(f"텍스트 크롤링 중 오류 발생: {e}")


try:
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    # Step 1: 자동차 리스트 페이지로 이동
    ncar_list_url = 'http://www.encar.com/mocha.do?mnfccd=001&mdlgroupcd=&mdlcd=&year=&mochaYn=&method='
    driver.get(ncar_list_url)
    driver.implicitly_wait(10)
    time.sleep(10)

    # Step 2: 자동차 리스트에서 각 자동차의 상세 페이지로 이동 (URL 리스트 추출)
    ncar_links = driver.find_elements(By.XPATH,
                                      '//div[@class="mocha_cont"]//ul[@class="list_mocha" and @id="list_mocha"]//li//a')
    ncar_urls = [link.get_attribute('href') for link in ncar_links]

    print(ncar_urls)
    print(f"총 {len(ncar_urls)}개의 차를 찾았습니다.")

    for ncar_url in ncar_urls:
        driver.get(ncar_url)
        time.sleep(2)

        try:
            # "btn_toggle" 클래스를 가진 버튼을 기다렸다가 클릭
            toggle_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH,
                                            "//div[@class='box_g box_perform' and @data-content='perform']//button[@class='btn_toggle']"))
            )
            toggle_button.click()  # 버튼 클릭
            time.sleep(5)

            # 클릭 후 'box_g box_perform mocha_on' 클래스로 변경되었는지 확인
            updated_div = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'box_g box_perform mocha_on')]"))
            )
            if updated_div:
                print("버튼 클릭 후 'box_g box_perform mocha_on' 클래스로 변경되었습니다.")
                crawl_text(driver)  # 크롤링 함수 호출

        except Exception as e:
            print(f"'펼치기' 버튼을 찾지 못했습니다: {ncar_url}. 오류: {e}")
            continue  # 버튼이 없는 경우 다음 항목으로 이동

except Exception as e:
    print(f"오류 발생: {e}")

finally:
    driver.quit()  # 크롬 드라이버 종료

    # 엑셀 파일 저장
    file_name = 'car_details_' + now.strftime('%Y-%m-%d_%H-%M-%S') + '.xlsx'
    xlsx.save(file_name)
    print(f"Excel 파일 저장 완료: {file_name}")
