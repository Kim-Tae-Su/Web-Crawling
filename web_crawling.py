from selenium import webdriver
import chromedriver_autoinstaller
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time 
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.actions.wheel_input import ScrollOrigin
import traceback
import csv
import requests
# chromedriver_autoinstaller.install()
def chrome_browser(download_dir, user_data_dir):
    """Chrome 브라우저 실행 및 설정"""
    options = Options()
    options.add_argument(fr"user-data-dir={user_data_dir}")             # 사용자 데이터 경로 설정
    prefs = {
        "download.default_directory": download_dir,                     # 기본 다운로드 경로
        "download.prompt_for_download": False,                          # 다운로드 시 대화창 표시 안 함
        "profile.default_content_setting_values.automatic_downloads": 1,
        "profile.default_content_settings.popups": 0,                   # 팝업 차단 해제
        "directory_upgrade": True,                                      # 기존 폴더에 덮어쓰기 허용
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(options=options)
    driver.maximize_window()
    return driver
        
        
def scrape_links(driver, url):
    """웹 페이지에서 링크 및 텍스트 크롤링"""
    driver.get(url)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    visited = set()
    link_list = []
    scroll_num = 0
    total_num = 1
    past_element = []

    while True:
        try:
            selector = ".ms-Link.linkStyles-338"
            elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, selector)))

            # 스크롤링이 멈추었을 경우 종료
            if past_element == elements:
                break

            for element in elements:
                if element.text in visited:
                    continue
                visited.add(element.text)
                print(f"total: {total_num}, scroll: {scroll_num}, text: {element.text}")
                link_list.append([element.text, element.get_attribute("href")])
                total_num += 1
                time.sleep(0.1)

            # 스크롤 내리기
            scroll_origin = ScrollOrigin.from_viewport(500, 500)
            ActionChains(driver).scroll_from_origin(scroll_origin, 0, 1500).perform()
            scroll_num += 1
            time.sleep(2)

            past_element = elements

        except Exception as ex:
            print("크롤링 중 오류 발생:", traceback.format_exc())
            break

    return link_list

def save_links_to_csv(links, filename="name_link.csv"):
    """링크 목록을 CSV 파일로 저장"""
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csv_file:
            writer = csv.writer(csv_file)
            for link in links:
                writer.writerow(link)
        print(f"파일 저장 완료: {filename}")
    except Exception:
        print("파일 저장 중 오류 발생:", traceback.format_exc())
        
download_path = "Z:\\dummy"
user_data_path = r"C:\Users\ts.kim\user_data"
target_url = "https://kohyoung.crm5.dynamics.com/main.aspx?appid=ad028f5a-570f-ea11-a811-000d3a085914&pagetype=entitylist&etn=new_machinehistory&viewid=e125189b-e5eb-e811-a97c-000d3aa041bf&viewType=1039"

driver = chrome_browser(download_path, user_data_path)
links = scrape_links(driver, target_url)
save_links_to_csv(links)

def load_links_from_csv(filename="name_link.csv"):
    """CSV 파일에서 L Size 장비 링크와 이름 로드"""
    links = []
    names = []
    try:
        with open(filename, 'r', newline='', encoding='utf-8') as csv_file:
            reader = csv.reader(csv_file)
            for row in reader:
                if row and (row[0].split('0')[0][-2:] in ["SL", "DL"]):
                    links.append(row[1]) 
                    names.append(row[0])
        print("링크 로드 완료")
    except Exception:
        print("파일 로드 중 오류 발생:", traceback.format_exc())

    return links, names

def download_files_from_links(driver, links, names, max_file_size):
    """각 링크에서 파일 크기를 확인하고, 제한 이하인 경우 다운로드"""
    for i, url in enumerate(links):
        try:
            driver.get(url)
            time.sleep(3)

            # "관련 항목" 탭 클릭
            element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "container_related_tab_0"))
            )
            driver.execute_script("arguments[0].click();", element)
            time.sleep(1)

            # "백업 파일" 탭 클릭
            element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "nav_new_new_machinehistory_new_backupfiles_l_machinehistory_Related"))
            )
            driver.execute_script("arguments[0].click();", element)
            time.sleep(3)

            # 파일 다운로드 링크 찾기
            elements = driver.find_elements(By.CSS_SELECTOR, "a[href^='https://kyblob.blob.core.windows.net']")
            if elements:
                file_url = elements[0].get_attribute('href')
                response = requests.head(file_url)

                if 'Content-Length' in response.headers:
                    file_size = int(response.headers['Content-Length'])
                    if file_size < max_file_size:
                        print(f"{i} 번째 {names[i]} 다운로드 시작")
                        driver.execute_script("arguments[0].click();", elements[0])
                        time.sleep(20)
                    else:
                        print(f"{i} 번째 {names[i]} 용량 초과로 다운로드하지 않음")
            time.sleep(1)

        except Exception:
            print(f"{i} 번째 {names[i]} 다운로드 중 오류 발생:\n", traceback.format_exc())
            break

links, names = load_links_from_csv()
MAX_FILE_SIZE = 583191062 # 500MB
download_files_from_links(driver, links, names, MAX_FILE_SIZE)
            