import PyPDF2 as pp
import csv
import pandas as pd

doc = pp.PdfFileReader("International DSM.pdf")

page = doc.getPage(3)

text = page.extractText()

raw_text = text.splitlines()
filtered_text = [x for x in raw_text if x != '¦']
filtered_text.pop()
headings = filtered_text[0:2]
day = filtered_text[2]
meta_columns = filtered_text[3:6]
rows = filtered_text[6:17]
del filtered_text[0:27]
cols = filtered_text[0:37]
del filtered_text[0:37]

print(filtered_text)
parsed_text = []


for i in filtered_text:
    if "(" in i:
       i = i.replace("(", "-")
    if ")" in i:
        i = i.replace(")", "")
    if "%" in i:
        i = i.replace("%", "")
    if "," in i:
        i = i.replace(",", "")
    if i == "n/a" or i == "N/A":
        i = None
    if i != None:
        float(i)
    parsed_text.append(i)

n = 37
column_data = [parsed_text[i * n:(i + 1) * n] for i in range((len(parsed_text) + n - 1) // n )] 

#cols = list(zip(columns,column_data))

df = pd.DataFrame(column_data, columns=cols, index=rows)

df.to_excel("output.xlsx")
