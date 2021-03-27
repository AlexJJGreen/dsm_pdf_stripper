import PyPDF2 as pp
import csv
import pandas as pd
import shutil

doc = pp.PdfFileReader("International DSM.pdf")

total_pages = doc.getNumPages()

#get meta -- wk ---
page_0 = doc.getPage(0).extractText()
print(page_0)

#list for df outputs
df_list = []

for page in range(15):
    if (page % 2 != 0) and (page != 0):
        #get current pages
        current_page = doc.getPage(page)
        text = current_page.extractText()
        #split into strings
        raw_text = text.splitlines()
        #remove lines
        filtered_text = [x for x in raw_text if x != '¦']
        #remove page num
        filtered_text.pop()
        # get meta
        headings = filtered_text[0:2]
        day = filtered_text[2]
        meta_columns = filtered_text[3:6]
        metrics = filtered_text[6:17]
        # remove meta data
        del filtered_text[0:27]
        # get store names
        shops = filtered_text[0:37]
        # delete store names
        del filtered_text[0:37]
        # new list for numeric data
        parsed_text = []

        # parse numeric data, mut to float, push to new arr
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

        # split data into cols of 37
        N = 37
        column_data = [parsed_text[i * N:(i + 1) * N] for i in range((len(parsed_text) + N - 1) // N )] 

        #create df
        df = pd.DataFrame(column_data, columns=shops, index=metrics).T
        df["day"] = day
        # swap cols and rows
        #df_T = df.T
        df_list.append(df)

concatenated = pd.concat(df_list)

filter_df = concatenated.loc["Berlin (Karstadt)"]

print(filter_df)

#filter_df.to_excel("berlin.xlsx")




#xl_file_name = "output" +  ### < --- parse pg 0, get week, get date --- create dynamic file name based on wk

#df.to_excel("output_1.xlsx")
#df_T.to_excel("output_2.xlsx")

# shutil.move("output") <-- move file to folder
