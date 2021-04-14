import PyPDF2 as pp
import csv
import pandas as pd
import shutil
from ordered_set import OrderedSet
import xlsxwriter

doc = pp.PdfFileReader("Story Analysis WTD INTERNATIONAL.pdf")

total_pages = doc.getNumPages()

store_datasets = []

story_datasets = {}

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

        story_datasets["STORY"] = list(OrderedSet(story_datasets["STORY"]))

        for i in range(len(story_datasets["Unit Mix %"])):
            if story_datasets["Item L3 Desc"][i] != "Total":
                    story_datasets["STORY"].insert(i, story_datasets["STORY"][i -1])
            if story_datasets["Unit Mix %"][i] == "n/a%":
                story_datasets["Units"].insert(i, 0)
            if i != 0:
                story_datasets["store"].insert(i, story_datasets["store"][i - 1])

        # parse to numeric
        #for i in range(4,8):
        #    print(story_datasets[i])
        #    print("------------------------------------------------")
        #    if len(story_datasets[i]) > 0:
        #        story_datasets[i] = parse_to_numeric(story_datasets[i])
        
        # story wtd finished, append to global store list

        store_datasets.append(story_datasets)
        story_datasets = {}

    elif raw_text[0] == "STORY ANALYSIS WEEK TO DATE":
        wtd = True

        # get meta
        story_datasets["store"] = [raw_text[1]]
        del raw_text[0:2]

        # get dict keys --> {Store: "", STORY: [], Item L3 Desc: [], Sales £: [], Units: [], Cash Mix %: [], Unit Mix %: []}
        for i in range(6):
            story_datasets[raw_text[i]] = []
        
        del raw_text[0:6]

        # strip story names into one col and append to datasets, break point on list i == Total && i + 1 == Total
        c = 0
        for i in range(len(raw_text)):
            if (raw_text[i] == "Total") and (raw_text[i + 1] == "Total"):
                break
            else:
                story_datasets["STORY"].append(raw_text[i])
                c += 1

        del raw_text[0:c]
        

        # cash and units mix cols ALWAYS contain %, identify and /2 to get col length, strip and append to dict

        col_len = int(sum("%" in s for s in raw_text)/2)

        for i in range(len(raw_text) - col_len,len(raw_text)):
            story_datasets["Unit Mix %"].append(raw_text[i])
        del raw_text[len(raw_text) - col_len:]

        for i in range(len(raw_text) - col_len,len(raw_text)):
            story_datasets["Cash Mix %"].append(raw_text[i])
        del raw_text[len(raw_text) - col_len:]

        # strip item desc append to Item L3 Desc

        for i in range(0,col_len):
            story_datasets["Item L3 Desc"].append(raw_text[i])
        del raw_text[0:col_len]

        # strip Sales £, append and del
        for i in range(0,col_len):
            story_datasets["Sales £"].append(raw_text[i])
        del raw_text[0:col_len]

        #append last list to Units
        for i in range(len(raw_text)):
            story_datasets["Units"].append(raw_text[i])

        #for key in story_datasets:
        #    print("{} | {} \n".format(key,story_datasets[key]))
        
    elif (raw_text[0] == "STORY") and (wtd is True):
        # delete column titles
        del raw_text[0:6]

        # stories alway UPPPER, append to stories
        story_appended_count = 0
        for text in raw_text:
            if text.isupper():
                story_datasets["STORY"].append(text)
                story_appended_count += 1
        
        del raw_text[0:story_appended_count]

        # cash and units mix cols ALWAYS contain %, identify and /2 to get col length, strip and append to dict

        col_len = int(sum("%" in s for s in raw_text)/2)

        for i in range(len(raw_text) - col_len,len(raw_text)):
            story_datasets["Unit Mix %"].append(raw_text[i])
        del raw_text[len(raw_text) - col_len:]

        for i in range(len(raw_text) - col_len,len(raw_text)):
            story_datasets["Cash Mix %"].append(raw_text[i])
        del raw_text[len(raw_text) - col_len:]

        # strip item desc append to Item L3 Desc

        for i in range(0,col_len):
            story_datasets["Item L3 Desc"].append(raw_text[i])
        del raw_text[0:col_len]

        # strip Sales £, append and del
        for i in range(0,col_len):
            story_datasets["Sales £"].append(raw_text[i])
        del raw_text[0:col_len]

        #append last list to Units
        for i in range(len(raw_text)):
            story_datasets["Units"].append(raw_text[i])

        #for key in story_datasets:
        #    print("{} | {} \n".format(key,story_datasets[key]))


    #elif raw_text[0] == "Item L3 Desc" and wtd is True:
    #    print("Triggered")
    #    # do cat stuff
    else:
        pass

for dataset in store_datasets:
    df = pd.DataFrame.from_dict(dataset)
    df.to_excel("test.xlsx", engine='xlsxwriter')
