import PyPDF2 as pp
import csv
import pandas as pd
import shutil
from ordered_set import OrderedSet
import xlsxwriter
from re import search

doc = pp.PdfFileReader("International DSM (2).pdf")

total_pages = doc.getNumPages()

#get meta -- wk ---
page_0 = doc.getPage(0).extractText()

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

#concat into single df
concatenated = pd.concat(df_list)
concatenated.index = concatenated.index.astype('str')
cols = concatenated.columns.drop("day")
concatenated[cols] = concatenated[cols].apply(pd.to_numeric, errors='coerce')
locations = OrderedSet(concatenated.index)

with pd.ExcelWriter('sales.xlsx') as writer:
    for location in locations:
        filter_df = concatenated.loc[location]
        filter_df.index = filter_df["day"]
        if location != "Total":
            filter_df.to_excel(writer, engine='xlsxwriter', sheet_name=location)

        ks_data = concatenated[concatenated.index.str.contains("Karstadt")].groupby("day").sum()
        inno_data = concatenated[concatenated.index.str.contains("Inno") | concatenated.index.str.contains("INNO")].groupby("day").sum()
        solus_data = concatenated[(concatenated.index.str.contains("Inno") == False) & (concatenated.index.str.contains("Karstadt") == False) & (concatenated.index.str.contains("Total") == False)].groupby("day").sum()

        def df_calc(sheetname, df, loc_count):
            # df.loc["Total"]
            df["v Bud %"] = (df["Sales Act £'k"] / df["Sales Bud £'k"]) - 1
            df["v LW %"] = (df["Sales Act £'k"] / df["Sales LW"]) - 1
            df["v LY %"] = (df["Sales Act £'k"] / df["Sales LY £'k"]) - 1
            df["Margin %"] = df["Margin %"] / loc_count
            df.reindex(["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"])
            df.to_excel(writer, engine=xlsxwriter, sheet_name=sheetname)

        df_calc("Karstadt", ks_data, 10)
        df_calc("Inno", inno_data, 9)
        df_calc("Solus", solus_data, 9)


# Sales LY £'k	v Bud %	v LW %	v LY %	Margin %	v LY %Pts	Returns Act £'k	Returns v LY%


#inno_data = []
#solus_data = []

#for i in concatenated.index:
#    if "(Karstadt)" in concatenated.index:
#        ks_data.append(i)
#    if "(Inno)" in concatenated.index:
#        inno_data.append(i)
#    else:
#        solus_data.append(i)

#print(ks_data)
#print(inno_data)
#print(solus_data)

#filter_df = concatenated.loc["Berlin (Karstadt)"]

#print(filter_df)

#filter_df.to_excel("berlin.xlsx")




#xl_file_name = "output" +  ### < --- parse pg 0, get week, get date --- create dynamic file name based on wk

#df.to_excel("output_1.xlsx")
#df_T.to_excel("output_2.xlsx")

# shutil.move("output") <-- move file to folder

#Sales Act £'k	Sales Bud £'k	Sales LW	Sales LY £'k	v Bud %	v LW %	v LY %	Margin %	v LY %Pts	Returns Act £'k	Returns v LY%

#Sunday Monday Tuesday Wednesday Thursday Friday Saturday
