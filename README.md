# Postal Code Bulk Search (우편번호 대량 검색)

`postal_code_bulk_search.py` 프로그램은 받는 분, 주소 및 연락처 필드와 함께 엑셀 파일 경로를 입력하여 우편번호를 자동으로 검색하고 일반 및 상세 주소를 구분하며 휴대전화 및 일반 전화 번호를 구분할 수 있습니다.

## 사용 방법:
postal_code_bulk_search.py 스크립트를 실행합니다.
Excel 파일을 제공해야 합니다. Excel 파일의 열 이름은 각각 "받는 분", "주소", "연락처"여야 합니다.

## 기능:
프로그램은 자동으로 다음을 수행합니다.
1. 우편번호 검색
2. 일반 및 상세 주소 구분
3. 휴대전화 및 일반 전화 번호 구분

## 요구 사항:
```cmd
pip install requests, lxml, xml.etree.ElementTree
```
## Pyinstaller 사용 시:
```cmd
pyinstaller postal_code_bulk_search.py --hidden-import requests --hidden-import lxml --hidden-import xml.etree.ElementTree
```
---
---

# Postal Code Bulk Search (English)

The `postal_code_bulk_search.py` program allows you to automatically search for postal codes, distinguish between regular and detailed addresses, and differentiate between mobile and landline phone numbers by simply inputting recipient information, address, and contact fields, along with the Excel file path.
## Usage:
Run the postal_code_bulk_search.py script.
You must provide an Excel file containing recipient information, address, and contact details. The column names in the Excel file must be "받는 분", "주소", and "연락처" respectively.

## Feature:
The program will automatically 
1. search for postal codes
2. distinguish between regular and detailed addresses
3. differentiate between mobile and landline phone numbers

## Requirements:
```cmd
pip install requests, lxml, xml.etree.ElementTree
```

## To use Pyinstaller:
```cmd
pyinstaller postal_code_bulk_search.py --hidden-import requests --hidden-import lxml --hidden-import xml.etree.ElementTree
```
