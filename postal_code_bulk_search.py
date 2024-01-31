import pandas as pd
import requests
import xml.etree.ElementTree as ET
import re

print("우체국우편주소 자동검색기")
print("받는 분,주소,연락처 필드만 적어서 아래에 엑셀 파일 경로를 input 해주십시오")
excel_file = input().replace(" ", "").replace("'", "").replace('"', '')
service_key = 'post office api key'

def address_pre(o_address):
    regex = r'(\w+[원,산,남,울,북,천,주,기,시,도]\s*)?' \
            r'(\w+[구,시,군]\s*)?(\w+[구,시]\s*)?' \
            r'(\w+[면,읍]\s*)?' \
            r'(\w+\d*\w*[동,리,로,길]\s*\d*)' \
            r'(\w*\d+-?\d*)?'
    pre_address = re.search(regex, o_address)
    address_full = pre_address.group(0)
    address_detail = o_address.replace(address_full, "").strip()
    if address_full is None:
        address_full = o_address
        address_detail = "분리결과 없음"
    address_detail_list.append(address_detail)
    address_full_list.append(address_full)
    return address_full, address_detail

def postoffice(address_full):
    uri = 'http://biz.epost.go.kr/KpostPortal/openapi2'
    total_url = uri + "?regkey=" + service_key + "&target=postNew&&query=" +address_full
    response = requests.get(total_url)
    xml_data = response.text
    root = ET.fromstring(xml_data)

# 원하는 요소 추출
    try:
        postcd = root.find('.//postcd').text.strip()
        address = root.find('.//address').text.strip()
        addrjibun = root.find('.//addrjibun').text.strip()

        # 결과 출력
        print('우편번호:', postcd)
        post_num = str(postcd)
    except AttributeError:
        print("우편번호 검색 결과가 없습니다")
        post_num = "검색결과없음"
    post_num_list.append(post_num)
    return post_num_list


# original_pd 데이터프레임 생성
try:
    original_pd = pd.read_excel(excel_file)
except:
    print("xls 는 불가능, xlsx 만 가능합니다")
original_pd.columns = ["받는 분", "주소", "연락처"]

# processed_pd 데이터프레임 생성
processed_pd = pd.DataFrame(columns=["받는 분", "우편번호", "전체주소", "세부주소", "휴대전화", "일반전화"])

# original_pd의 각 항목을 처리하여 processed_pd에 저장
for index, row in original_pd.iterrows():
    o_receiver = row['받는 분']
    o_address = row['주소']
    o_contact = str(row['연락처']).replace(" ", "").replace("-", "").replace(")", "")

    post_num_list = []
    address_detail_list = []
    address_full_list = []
    address_full, address_detail = address_pre(o_address)
    postoffice(address_full)

    if str(o_contact)[:2] in ["01", "10"]:
        processed_pd = pd.concat([processed_pd, pd.DataFrame({"받는 분": [o_receiver], "우편번호": [post_num_list], "전체주소": [address_full_list], "세부주소": [address_detail_list], "휴대전화": [o_contact], "일반전화": [""]})])
    else:
        processed_pd = pd.concat([processed_pd, pd.DataFrame({"받는 분": [o_receiver], "우편번호": [post_num_list], "전체주소": [address_full_list], "세부주소": [address_detail_list], "휴대전화": [""], "일반전화": [o_contact]})])

# processed_pd를 "가공됨.xlsx" 파일로 저장
try:
    for column in processed_pd.columns:
        processed_pd[column] = processed_pd[column].astype(str).str.replace(r"[\[\]']", "", regex=True)
    processed_pd.to_excel(r"C:\Users\movem\Desktop\가공.xlsx", index=False)
except PermissionError:
    print("엑셀 파일이 열려있어서 저장할 수 없습니다")

