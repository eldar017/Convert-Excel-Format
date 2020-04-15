import base64
import io
import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook

with open('c:\om\exampleMasterCard.xlsx', 'rb') as binary_file:
    binary_file_data = binary_file.read()
    base64_encoded_data = base64.b64encode(binary_file_data)
    base64_message = base64_encoded_data.decode('utf-8')

print(base64_message)
decoded_data = base64.b64decode(base64_message)
xls_filelike = io.BytesIO(decoded_data)
workbook = openpyxl.load_workbook(xls_filelike)

workbook.save(filename="C:\OM\Eldar991120.xlsx")