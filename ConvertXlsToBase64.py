import base64
import io
from xlrd import open_workbook
from xlutils.copy import copy

with open('c:\om\exampleMasterCard.xls', 'rb') as binary_file:
    binary_file_data = binary_file.read()
    base64_encoded_data = base64.b64encode(binary_file_data)
    base64_message = base64_encoded_data.decode('utf-8')

print(base64_message)
decoded_data = base64.b64decode(base64_message)
xls_filelike = io.BytesIO(decoded_data)
xl_workbook = open_workbook(file_contents=decoded_data, on_demand=True)
wb = copy(xl_workbook)
wb.save('c:\om\output-xls.xls')



