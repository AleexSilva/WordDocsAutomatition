import pandas as pd
from datetime import datetime
from docxtpl import DocxTempate

# Global Variables

name = "Alex Silva"
Phone = "+1 547822516"
email = "Alex.silva94@gmail.com"
date =  datetime.today().strftime('%Y-%m-%d')

#Iniciate Doc Template

doc = DocxTempate('template.docx')
