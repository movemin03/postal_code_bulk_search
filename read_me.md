# Postal Code Bulk Search

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
