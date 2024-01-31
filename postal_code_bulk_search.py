import pandas as pd
import requests
import xml.etree.ElementTree as ET

print("우체국우편주소 자동검색기")
print("받는 분,주소,연락처 필드만 적어서 아래에 엑셀 파일 경로를 input 해주십시오")
excel_file = input().replace(" ", "").replace("'", "").replace('"', '')
#excel_file = r"C:\Users\3-01\Desktop\발송관련양식\등기발송양식.xlsx"

def postoffice(o_address):
    uri = 'http://biz.epost.go.kr/KpostPortal/openapi2'
    service_key = 'Korea post office API KEY'
    total_url = uri + "?regkey=" + service_key + "&target=postNew&&query=" +o_address
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
        print("검색 결과가 없습니다")
        post_num = "검색결과없음"
    post_num_list.append(post_num)
    return post_num_list


# original_pd 데이터프레임 생성
original_pd = pd.read_excel(excel_file)
original_pd.columns = ["받는 분", "주소", "연락처"]

# processed_pd 데이터프레임 생성
processed_pd = pd.DataFrame(columns=["받는 분", "우편번호", "주소", "휴대전화", "일반전화"])

# original_pd의 각 항목을 처리하여 processed_pd에 저장
for index, row in original_pd.iterrows():
    o_receiver = row['받는 분']
    o_address = row['주소']
    o_contact = row['연락처'].replace(" ", "").replace("-", "").replace(")", "")

    post_num_list = []
    postoffice(o_address)

    if str(o_contact)[:2] in ["01", "10"]:
        processed_pd = pd.concat([processed_pd, pd.DataFrame({"받는 분": [o_receiver], "우편번호": [post_num_list], "주소": [o_address], "휴대전화": [o_contact], "일반전화": [""]})])
    else:
        processed_pd = pd.concat([processed_pd, pd.DataFrame({"받는 분": [o_receiver], "우편번호": [post_num_list], "주소": [o_address], "휴대전화": [""], "일반전화": [o_contact]})])

# processed_pd를 "가공됨.xlsx" 파일로 저장
try:
    processed_pd.to_excel(r"C:\Users\3-01\Desktop\발송관련양식\가공됨.xlsx", index=False)
except PermissionError:
    print("엑셀 파일이 열려있어서 저장할 수 없습니다")

