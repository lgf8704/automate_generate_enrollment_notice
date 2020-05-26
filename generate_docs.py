import os
import docx 


# 1. 找到excel与模板文档
for filename in os.listdir():
    if filename.endswith('.xls') or filename.endswith('.xlsx'):
        excel_file = filename
    elif filename.endswith('.doc') or filename.endswith('.docx'):
        doc_temp = filename

# 2. 从excel文件中读取数据   
input_messages = {}
with open(excel_file, 'rb') as ef:
