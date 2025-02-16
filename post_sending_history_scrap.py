from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os
from datetime import datetime
import time


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


def create_detail_dataframe():
    return pd.DataFrame(columns=[
        '등기번호', '사전접수일', '우편물구분', '분할여부', '받는 분',
        '전화번호', '주소', '특수취급', '중량(g)', '크기(Cm)',
        '내용품명', '내용품상세', '배달방식', '배송시요청사항',
        '진행상태', '취소여부'
    ])


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


def scrape_basic_data(driver):
    df = create_basic_dataframe()
    
    while True:
        driver.get(
            "https://service.epost.go.kr/front.commonpostplus.RetrieveMyBefRecevList.postal?loginType=login&s_url=https://service.epost.go.kr/front.commonpostplus.RetrieveMyBefRecevList.postal&login=service&webacc=")

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
                    if len(data) == len(df.columns):
                        df.loc[len(df)] = data
                    else:
                        print(f"Warning: Data length mismatch. Expected {len(df.columns)}, got {len(data)}")
                        print(f"Data: {data}")
            except Exception as e:
                print(f"Error processing row: {e}")
                continue

        print("다음 페이지도 탐색하시겠습니까? (y/n): ")
        if input().lower() != 'y':
            break

        print("페이지 이동 후 엔터를 눌러주세요...")
        input()

    return df


def scrape_detail_data(driver, basic_df):
    detail_df = create_detail_dataframe()
    total_records = len(basic_df)

    for index, row in basic_df.iterrows():
        try:
            print(f"상세 정보 처리 중... {index + 1}/{total_records}")

            url = f"https://service.epost.go.kr/commonpostplus.RetrieveMyBefRecvDtailInfo.postal?serviceType=H&tmprecevno={row['접수번호2']}&sGubun=3"
            driver.get(url)

            table = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "table_list"))
            )

            tbody = table.find_element(By.TAG_NAME, "tbody")
            detail_row = tbody.find_element(By.TAG_NAME, "tr")

            data = extract_table_data(detail_row)
            detail_df = pd.concat([detail_df, pd.DataFrame([data])], ignore_index=True)

            time.sleep(1)

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
    
    # 로그인
    driver.get("https://www.epost.go.kr/usr/login/cafzc008k01.jsp?s_url=https://service.epost.go.kr/")
    print("로그인을 완료한 후 엔터를 눌러주세요...")
    input()

    try:
        # 기본 데이터 스크래핑
        print("\n=== 기본 배송 정보 수집 시작 ===")
        basic_df = scrape_basic_data(driver)
        basic_file_path = save_to_excel(basic_df, "우체국_배송조회")
        print(f"기본 정보가 저장되었습니다: {basic_file_path}")

        # 상세 데이터 스크래핑
        print("\n=== 상세 배송 정보 수집 시작 ===")
        detail_df = scrape_detail_data(driver, basic_df)
        detail_file_path = save_to_excel(detail_df, "우체국_배송상세정보")
        print(f"상세 정보가 저장되었습니다: {detail_file_path}")

    finally:
        driver.quit()


if __name__ == "__main__":
    main()
