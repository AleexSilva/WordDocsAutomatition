import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
import os

# Global Variables
name = "Alex Silva"
phone = "+1 547822516"
email = "alex.silva94@gmail.com"
date =  datetime.today().strftime('%Y-%m-%d')

DICT = {'name':name,'phone':phone,'email':email,'date':date}


base_path = 'input'
file_names = ['template.docx','students.xlsx']

#Iniciate Doc Template

file_path1 = os.path.join(base_path, file_names[0])
doc = DocxTemplate(file_path1)


# Read Excel file
file_path2 = os.path.join(base_path, file_names[1])
df = pd.read_excel(file_path2)

# Iterate Excel and load information on Word File

for index,row in df.iterrows():
    content = {'student_name':row['Student_name'],
            'math_note':row['Math'],
            'physics_note':row['Physics'],
            'chemistry_note':row['Chemistry']
            }
    content.update(DICT)
    file_name = row['Student_name'].split(' ')[0]
    doc.render(content)
    doc.save(f"output/student_{file_name}.docx")
