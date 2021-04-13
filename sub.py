import PyPDF2 as pp
import csv
import pandas as pd
import shutil

doc = pp.PdfFileReader("Story Analysis WTD INTERNATIONAL.pdf")

total_pages = doc.getNumPages()

store_datasets = []

story_datasets = []

wtd = True

def parse_to_numeric(filtered_text):
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
        if (i != None) or (i != "Total"):
            float(i)
        parsed_text.append(i)
    
    return parsed_text

for page in range(10):
    # get page extract text
    current_page = doc.getPage(page).extractText()
    # split into list
    raw_text = current_page.splitlines()

    # convert to python str 
    for i in range(len(raw_text)):
        str(raw_text[i])
    
    # global dataset obj
    cat_datasets = []

    # check page type
    if raw_text[0] == "STORY ANALYSIS YESTERDAY":
        wtd = False
        #space subindex stories with blank spacers
        for i in range(len(story_datasets[3])):
            if story_datasets[3][i] != "Total":
                story_datasets[2].insert(i,"")
        
        # parse to numeric
        #for i in range(4,8):
        #    print(story_datasets[i])
        #    print("------------------------------------------------")
        #    if len(story_datasets[i]) > 0:
        #        story_datasets[i] = parse_to_numeric(story_datasets[i])
        
        # story wtd finished, append to global store list
        store_datasets.append(story_datasets)
        story_datasets = []

    elif raw_text[0] == "STORY ANALYSIS WEEK TO DATE":
        wtd = True
        #get meta
        meta = {"store": raw_text[1]}
        story_datasets.append([meta])
        del raw_text[0:2]

        # get 6 col headings
        headings = raw_text[0:6]
        story_datasets.append(headings)
        del raw_text[0:6]

        # strip story names into one col and append to datasets
        col = []
        for i in range(len(raw_text)):
            if (raw_text[i] == "Total") and (raw_text[i + 1] == "Total"):
                break
            else:
                col.append(raw_text[i])
        del raw_text[0:len(col)]
        story_datasets.append(col)

        # split remaining data in 5 cols and append to datasets
        col_len = int(len(raw_text) / 5)
        for i in range(5):
            temp_col = []
            for j in range(col_len):
                temp_col.append(raw_text[j])
            story_datasets.append(temp_col)
            del raw_text[0:col_len]
        
        
    elif (raw_text[0] == "STORY") and (wtd is True):
        del raw_text[0:6]

        story_appended_count = 0
        for text in raw_text:
            if text.isupper():
                story_datasets[2].append(text)
                story_appended_count += 1
        
        del raw_text[0:story_appended_count]

        print(len(raw_text))
        print("----------------------------------------------")

        # split remaining data in 5 cols and append to datasets
        col_len = sum("%" in s for s in raw_text)/2
        for i in range(5):
            temp_col = []
            for j in range(col_len):
                story_datasets[i + 3].append(raw_text[j])
                temp_col.append(raw_text[j])
            print(temp_col)
            print("----------------------------------------------------")
            del raw_text[0:col_len]

    #elif raw_text[0] == "Item L3 Desc" and wtd is True:
    #    print("Triggered")
    #    # do cat stuff
    else:
        pass

#for i in range(len(store_datasets[0])):
#    print(store_datasets[1][i])