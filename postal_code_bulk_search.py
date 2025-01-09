import pandas as pd
import requests
import xml.etree.ElementTree as ET
import re
import os

print("우체국 - 우편주소 자동검색기 for Thinkgood")
print("https://github.com/movemin03/postal_code_bulk_search")
print("받는 분,주소,연락처 필드만 적어서 아래에 엑셀 파일 경로를 넣어주십시오")
excel_file = input().replace("'", "").replace('"', '')
service_key = 'Your_API_KEY_from_우체국'


def address_pre(o_address):
    # 콤마로 먼저 분리
    base_addr = o_address
    detail = ""
    if ',' in o_address:
        base_addr, detail = o_address.split(',', 1)
        base_addr = base_addr.strip()

    # 정규표현식 적용
    regex = r'(\w+[원,산,남,울,북,천,주,기,시,도]\s*)?' \
            r'(\w+[구,시,군]\s*)?(\w+[구,시]\s*)?' \
            r'(\w+[면,읍]\s*)?' \
            r'(\w+[동,리,로,길]\s*\d+(?:-?\d+)?(?:\([^)]+\))?)'
    r'(?=\s|$)'


    pre_address = re.search(regex, base_addr)
    if pre_address:
        address_full = pre_address.group(0)
        if detail:
            address_detail = detail.strip()
        else:
            address_detail = base_addr.replace(address_full, "").strip()
    else:
        address_full = base_addr
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
except Exception as e:
    print("엑셀 파일 인식 실패 오류입니다")
    print(e)
    print("xls 는 불가능, xlsx 만 가능합니다")
    print("필드명이 '받는 분','주소','연락처'인지 확인")
    print("제공한 엑셀 파일이 사용중인 것은 아닌지 확인")
    a = input()
try:
    original_pd.columns = ["받는 분", "주소", "연락처"]
except:
    print("받는 분,주소,연락처 필드로 이루어지지 않아서 프로그램이 인식할 수 없습니다")
    print("혹은 받는 분,주소,연락처 필드 외 다른 데이터가 있는 경우에도 해당 오류가 발생합니다")
    a = input()
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
def get_available_filename(path):
    base_dir = os.path.dirname(path)
    base_name = os.path.basename(path)
    name, ext = os.path.splitext(base_name)
    i = 1
    while os.path.exists(path):
        new_name = f"{name}{i}{ext}"
        path = os.path.join(base_dir, new_name)
        i += 1
    return path

# processed_pd를 "가공됨.xlsx" 파일로 저장
try:
    for column in processed_pd.columns:
        processed_pd[column] = processed_pd[column].astype(str).str.replace(r"[\[\]']", "", regex=True)
    folder_path= os.path.dirname(excel_file)
    result_path = os.path.join(folder_path, "우편번호가공.xlsx")
    result_path = get_available_filename(result_path)
    processed_pd.to_excel(result_path , index=False)
except PermissionError:
    print("엑셀 파일이 열려있어서 저장할 수 없습니다")

print("검색이 완료되었습니다.")
print(result_path, " 에 저장되었습니다")
a = input()
