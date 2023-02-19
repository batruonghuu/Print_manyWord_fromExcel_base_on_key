from pathlib import Path

import pandas as pd
from docxtpl import DocxTemplate #pip install docxtpl

base_dir = Path(__file__).parent
word_template_path = base_dir / "vendor-contract.docx"
excel_path = base_dir / "contract-list.xlsx"
output_dir = base_dir / "OUTPUT"

print(excel_path)
print(word_template_path)
print(output_dir)

#Create output folder for the word documents
output_dir.mkdir(exist_ok=True)

#Convert Excel sheet into pandas dataframe
df = pd.read_excel(excel_path, sheet_name="Sheet1")

print(df)

#Iterate over each row in df and render word document
for record in df.to_dict(orient="records"):

    doc = DocxTemplate(word_template_path)
    doc.render(record)
    output_path = output_dir / f"{record['TEN_CTY']}-contract.docx"
    doc.save(output_path)