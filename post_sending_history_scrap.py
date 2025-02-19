from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os
from datetime import datetime
import time

def extract_table_data(row):
    def get_text(xpath):
        try:
            return row.find_element(By.XPATH, xpath).text.strip()
        except:
            return ""

    data = {
        '등기번호': get_text(".//td[1]"),
        '사전접수일': get_text(".//td[2]"),
        '우편물구분': get_text(".//td[3]"),
        '분할여부': get_text(".//td[4]"),
        '받는 분': get_text(".//td[5]"),
        '전화번호': get_text(".//td[6]"),
        '주소': get_text(".//td[7]"),
        '특수취급': get_text(".//td[8]"),
        '중량(g)': get_text(".//td[9]").replace('g', '').strip(),
        '크기(Cm)': get_text(".//td[10]"),
        '내용품명': get_text(".//td[11]"),
        '내용품상세': get_text(".//td[12]"),
        '배달방식': get_text(".//td[13]"),
        '배송시요청사항': get_text(".//td[14]"),
        '진행상태': get_text(".//td[15]"),
        '취소여부': get_text(".//td[16]")
    }
    return data

def extract_popup_data(onclick_text):
    content = onclick_text.replace("return popPrint(", "").rstrip(")")
    processed_values = []
    current_value = ''
    in_quotes = False

    for char in content:
        if char == "'":
            if in_quotes:
                processed_values.append(current_value)
                current_value = ''
            in_quotes = not in_quotes
        elif in_quotes:
            current_value += char

    return processed_values

def create_basic_dataframe():
    return pd.DataFrame(columns=[
        '년도', '월', '일', '발신인', '우편번호', '전화번호',
        '주소1', '주소2', '통수', '요금', 'URL',
        '접수번호1', '접수번호2', '이메일', '수신여부', '신청여부', '지역코드'
    ])

def get_last_page_number(driver):
    """마지막 페이지 번호를 확인"""
    last_page_button = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "a.end"))
    )
    last_page_button.click()
    time.sleep(1)

    current_page_element = driver.find_element(By.CSS_SELECTOR, "li.on strong")
    last_page = int(current_page_element.text)

    first_page_button = driver.find_element(By.CSS_SELECTOR, "a.fir")
    first_page_button.click()
    time.sleep(1)

    return last_page

def scrape_page_data(driver):
    """현재 페이지의 데이터를 스크래핑"""
    page_df = create_basic_dataframe()
    page_total_count = 0

    table = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "table_list.ma_b_10"))
    )
    tbody = table.find_element(By.TAG_NAME, "tbody")
    rows = tbody.find_elements(By.TAG_NAME, "tr")

    for row in rows:
        try:
            last_td = row.find_elements(By.TAG_NAME, "td")[-1]
            link = last_td.find_element(By.TAG_NAME, "a")
            onclick = link.get_attribute("onclick")

            if onclick and "return popPrint" in onclick:
                data = extract_popup_data(onclick)
                if len(data) == len(page_df.columns):
                    page_df.loc[len(page_df)] = data
                    page_total_count += int(data[8].strip().replace("통", ""))  # '통수' 컬럼의 값
        except Exception as e:
            print(f"Error processing row: {e}")
            continue

    return page_df, page_total_count

def scrape_by_page_count(driver, target_pages):
    """페이지 수 기준으로 데이터 스크래핑"""
    df = create_basic_dataframe()
    current_page = 1
    total_count = 0

    while current_page <= target_pages:
        print(f"페이지 처리 중: {current_page}/{target_pages}")

        page_df, page_count = scrape_page_data(driver)
        df = pd.concat([df, page_df], ignore_index=True)
        total_count += page_count

        if current_page < target_pages:
            next_button = driver.find_element(By.CSS_SELECTOR, "a.nex")
            next_button.click()
            time.sleep(0.1)

        current_page += 1

    return df


def get_current_page_number(driver):
    """현재 페이지 번호 반환"""
    try:
        current_page_element = driver.find_element(By.CSS_SELECTOR, "li.on strong")
        return int(current_page_element.text)
    except:
        return None


def scrape_by_item_count(driver, target_count):
    """통수 기준으로 데이터 스크래핑"""
    df = create_basic_dataframe()
    current_page = 1
    total_count = 0
    last_page_reached = False

    while True:
        print(f"현재까지 처리된 통수: {total_count}")

        # 현재 페이지 번호 확인
        before_page_num = get_current_page_number(driver)

        page_df, page_count = scrape_page_data(driver)
        df = pd.concat([df, page_df], ignore_index=True)
        total_count += page_count

        # 목표 통수 달성 확인
        if total_count >= target_count:
            break

        # 다음 페이지 버튼 확인 및 클릭
        next_button = driver.find_elements(By.CSS_SELECTOR, "a.nex")
        if not next_button:
            last_page_reached = True
            break

        next_button[0].click()
        time.sleep(0.1)

        # 페이지 변경 확인
        after_page_num = get_current_page_number(driver)
        if before_page_num == after_page_num:
            last_page_reached = True
            break

        current_page += 1

    # 결과 메시지 생성
    if last_page_reached and total_count < target_count:
        print(f"\n요청하신 {target_count}통 중 {total_count}통만 수집할 수 있습니다.")
        print(f"실제 데이터가 {target_count - total_count}통 부족합니다.")

    return df

def create_detail_dataframe():
    return pd.DataFrame(columns=[
        '등기번호', '사전접수일', '우편물구분', '분할여부', '받는 분',
        '전화번호', '주소', '특수취급', '중량(g)', '크기(Cm)',
        '내용품명', '내용품상세', '배달방식', '배송시요청사항',
        '진행상태', '취소여부'
    ])

def extract_number(text):
    """문자열에서 숫자만 추출"""
    import re
    numbers = re.findall(r'\d+', str(text))
    return int(numbers[0]) if numbers else 0

def scrape_detail_data(driver, basic_df):
    detail_df = create_detail_dataframe()
    total_records = len(basic_df)

    for index, row in basic_df.iterrows():
        try:
            print(f"상세 정보 처리 중... {index + 1}/{total_records}")

            total_items = extract_number(row['통수'])  # '200통'에서 200만 추출
            if total_items == 0:
                continue

            items_per_page = 15
            total_pages = (total_items + items_per_page - 1) // items_per_page

            base_url = f"https://service.epost.go.kr/commonpostplus.RetrieveMyBefRecvDtailInfo.postal?serviceType=H&tmprecevno={str(row['접수번호2'])}&sGubun=3"

            for page in range(1, total_pages + 1):
                target_row = (page - 1) * items_per_page + 1
                page_url = f"{base_url}&targetRow={target_row}"
                driver.get(page_url)

                table = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "table_list"))
                )
                tbody = table.find_element(By.TAG_NAME, "tbody")
                rows = tbody.find_elements(By.TAG_NAME, "tr")

                for detail_row in rows:
                    data = extract_table_data(detail_row)
                    detail_df = pd.concat([detail_df, pd.DataFrame([data])], ignore_index=True)

                time.sleep(0.1)

        except Exception as e:
            print(f"Error processing row {index}: {e}")
            continue

    return detail_df

def save_to_excel(df, filename_prefix):
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{filename_prefix}_{current_time}.xlsx"
    full_path = os.path.join(desktop_path, filename)

    df.to_excel(full_path, index=False)
    return full_path

def main():
    driver = webdriver.Chrome()
    base_url = "https://service.epost.go.kr/front.commonpostplus.RetrieveMyBefRecevList.postal"

    try:
        # 로그인
        driver.get("https://www.epost.go.kr/usr/login/cafzc008k01.jsp?s_url=https://service.epost.go.kr/")
        print("로그인을 완료한 후 엔터를 눌러주세요...")
        input()

        # base_url로 이동
        driver.get(base_url)
        a = input("날짜 필터를 설정한 후 엔터")

        # 마지막 페이지 확인
        last_page = get_last_page_number(driver)
        print(f"\n총 {last_page}페이지가 있습니다.")

        # 수집 방식 선택
        while True:
            collection_type = input("\n수집 방식을 선택해주세요 (1: 페이지 수 기준, 2: 통수 기준): ")
            if collection_type in ['1', '2']:
                break
            print("1 또는 2를 입력해주세요.")

        if collection_type == '1':
            while True:
                try:
                    target_pages = int(input(f"처리할 페이지 수를 입력해주세요 (최대 {last_page}페이지): "))
                    if 0 < target_pages <= last_page:
                        basic_df = scrape_by_page_count(driver, target_pages)
                        break
                    print(f"1에서 {last_page} 사이의 숫자를 입력해주세요.")
                except ValueError:
                    print("올바른 숫자를 입력해주세요.")
        else:
            while True:
                try:
                    target_count = int(input("처리할 통수를 입력해주세요: "))
                    if target_count > 0:
                        basic_df = scrape_by_item_count(driver, target_count)
                        break
                    print("0보다 큰 숫자를 입력해주세요.")
                except ValueError:
                    print("올바른 숫자를 입력해주세요.")

        # 결과 저장
        basic_file_path = save_to_excel(basic_df, "우체국_배송조회")
        print(f"\n기본 정보가 저장되었습니다: {basic_file_path}")

        # 상세 데이터 스크래핑
        print("\n=== 상세 배송 정보 수집 시작 ===")
        detail_df = scrape_detail_data(driver, basic_df)
        detail_file_path = save_to_excel(detail_df, "우체국_배송상세정보")
        print(f"상세 정보가 저장되었습니다: {detail_file_path}")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
