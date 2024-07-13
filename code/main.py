import pandas as pd
from datetime import datetime
from docxtpl import DocxTempate
import os
# Global Variables

name = "Alex Silva"
phone = "+1 547822516"
email = "Alex.silva94@gmail.com"
date =  datetime.today().strftime('%Y-%m-%d')

DICT = {'name':name,'phone':phone,'email':email,'date':date}


base_path = 'input'
file_names = ['templates.docx','students.xls']


#Iniciate Doc Template

file_path1 = os.path.join(base_path, file_names[0])
doc = DocxTempate(file_path1)


# Read Excel file
file_path2 = os.path.join(base_path, file_names[1])
df = pd.read_excel(file_path2)
df
