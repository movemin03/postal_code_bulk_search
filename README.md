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
---
---
# postal_code_bulk_search.py

## 소개
postal_code_bulk_search.py는 우체국 홈페이지에서 사용자의 발송 이력을 자동으로 수집하고 저장하는 프로그램입니다. 기본 발송 정보와 상세 배송 정보를 모두 수집하여 엑셀 파일로 저장합니다.

## 기능
- 기본 발송 정보 수집 (발신인, 수신인, 접수번호 등)
- 각 발송건의 상세 배송 정보 수집
- 수집된 데이터 자동 엑셀 저장

## 필요 사항
- Chrome 웹 브라우저
- ChromeDriver (Chrome 버전과 호환되는 버전)
- 필요한 Python 패키지:
  - selenium
  - pandas
  - datetime

2. 우체국 홈페이지 로그인
- 프로그램이 자동으로 로그인 페이지로 이동합니다
- 로그인을 완료한 후 엔터키를 눌러주세요

3. 데이터 수집
- 기본 발송 정보가 자동으로 수집됩니다
- 페이지 이동이 필요한 경우 안내에 따라 진행해주세요
- 상세 정보는 자동으로 수집됩니다

4. 결과 확인
- 수집된 데이터는 바탕화면에 두 개의 엑셀 파일로 저장됩니다
  - 기본 정보: "우체국_배송조회_[날짜시간].xlsx"
  - 상세 정보: "우체국_배송상세정보_[날짜시간].xlsx"

## 주의사항
- 네트워크 상태에 따라 수집 시간이 달라질 수 있습니다
- 많은 양의 데이터를 한 번에 수집할 경우 시간이 오래 걸릴 수 있습니다
  
---
---
# postal_code_bulk_search.py

## Introduction
postal_code_bulk_search.py is a program that automatically collects and saves dispatch history from the Korea Post website. It collects both basic shipping information and detailed delivery information, saving them to Excel files.

## Features
- Collection of basic shipping information (sender, receiver, receipt numbers, etc.)
- Collection of detailed delivery information for each shipment
- Automatic Excel file generation for collected data

## Requirements
- Chrome web browser
- ChromeDriver (compatible with your Chrome version)
- Required Python packages:
  - selenium
  - pandas
  - datetime

2. Login to Korea Post website
- Program will automatically navigate to the login page
- Complete the login and press Enter to continue

3. Data Collection
- Basic shipping information will be collected automatically
- Follow the prompts if page navigation is needed
- Detailed information will be collected automatically

4. Check Results
- Collected data will be saved as two Excel files on your desktop
  - Basic info: "우체국_배송조회_[datetime].xlsx"
  - Detailed info: "우체국_배송상세정보_[datetime].xlsx"

## Cautions
- Collection time may vary depending on network conditions
- Collecting large amounts of data may take considerable time
