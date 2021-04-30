import PyPDF2 as pp
import csv
import pandas as pd
import shutil
from ordered_set import OrderedSet
import xlsxwriter
from re import search

#### --- GLOBALS --- ####

doc = pp.PdfFileReader("Story Analysis WTD INTERNATIONAL.pdf")
total_pages = doc.getNumPages()
store_datasets = []
story_datasets = {}
# cat_datasets = {}

# week to date counter, on false ignore pages, append current story_dataset to store_dataset
wtd = True

# mutate str to float on numeric cols, nb. order of ops prevents refactor, don't touch!
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
            i = 0
        if (i != None) or (i != "Total"):
            float(i)
        parsed_text.append(i)
    
    return parsed_text


for page in range(total_pages):
    # get page extract text
    current_page = doc.getPage(page).extractText()
    # split into list
    raw_text = current_page.splitlines()

    # convert to python str 
    for i in range(len(raw_text)):
        str(raw_text[i])
    
    # check page type
    if raw_text[0] == "STORY ANALYSIS YESTERDAY":
        #print("triggered")
        wtd = False

        # print(story_datasets)
        try:
            story_datasets["STORY"] = list(OrderedSet(story_datasets["STORY"]))

            for i in range(len(story_datasets["Unit Mix %"])):
                if story_datasets["Item L3 Desc"][i] != "Total":
                        story_datasets["STORY"].insert(i, story_datasets["STORY"][i -1])
                if story_datasets["Unit Mix %"][i] == "n/a%":
                    story_datasets["Units"].insert(i, 0)
                if i != 0:
                    story_datasets["store"].insert(i, story_datasets["store"][i - 1])
            
            # parse Sales £, Units, Cash Mix %, Unit Mix % to float
            story_datasets["Sales £"] = parse_to_numeric(story_datasets["Sales £"])
            story_datasets["Cash Mix %"] = parse_to_numeric(story_datasets["Cash Mix %"])
            #print(story_datasets["Cash Mix %"])
            for i in story_datasets["Cash Mix %"]:
                if (i != None) or (i != "Total"):
                    i = float(i) / 100
            story_datasets["Unit Mix %"] = parse_to_numeric(story_datasets["Unit Mix %"])
            for i in story_datasets["Unit Mix %"]:
                if (i != None) or (i != "Total"):
                    i = float(i) / 100
            store_datasets.append(story_datasets)
        #print(store_datasets)
        except:
        #    print("passed")
            pass

        story_datasets = {}

    elif raw_text[0] == "STORY ANALYSIS WEEK TO DATE":
        # check page is not empty
        if len(raw_text) >= 9:
            # set bool trigger for WTD pages without title
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
        else:
            pass
        
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

        # append last list to Units
        for i in range(len(raw_text)):
            story_datasets["Units"].append(raw_text[i])

    else:
        pass

# datasets to df -> to excel
with pd.ExcelWriter('stor_analysis.xlsx') as writer:
    for dataset in store_datasets:
        sheetname = dataset["store"][0]
        df = pd.DataFrame.from_dict(dataset)
        df.set_index(["store", "STORY", "Item L3 Desc"], inplace=True)
        df.to_excel(writer, engine='xlsxwriter', sheet_name=sheetname)

    # collate store, ks, inno to df
    ks_data = []
    inno_data = []
    solus_data = []
    for dataset in store_datasets:
        if "Karstadt" in dataset["store"][0]:
            dataset["Grouping"] = ["Karstadt" for x in range(len(dataset["store"]))]
            del dataset["store"]
            df = pd.DataFrame.from_dict(dataset)
            df.set_index(["Grouping", "STORY", "Item L3 Desc"], inplace=True)
            ks_data.append(df)
        elif "Inno" in dataset["store"][0]:
            dataset["Grouping"] = ["Inno" for x in range(len(dataset["store"]))]
            del dataset["store"]
            df = pd.DataFrame.from_dict(dataset)
            df.set_index(["Grouping", "STORY", "Item L3 Desc"], inplace=True)
            inno_data.append(df)
        elif ("Inno" not in dataset["store"][0]) and ("Karstadt" not in dataset["store"][0]) and (dataset["store"][0] != "INTERNATIONAL"):
            dataset["Grouping"] = ["Solus" for x in range(len(dataset["store"]))]
            del dataset["store"]
            df = pd.DataFrame.from_dict(dataset)
            df.set_index(["Grouping", "STORY", "Item L3 Desc"], inplace=True)
            solus_data.append(df)
    
    # print(ks_data)
    collated_dfs = []
    collated_dfs.append(ks_data)
    collated_dfs.append(inno_data)
    collated_dfs.append(solus_data)
    for dfs in collated_dfs:
        collated_df = pd.concat(dfs)
        collated_df["Sales £"] = collated_df["Sales £"].apply(pd.to_numeric)
        collated_df["Units"] = collated_df["Units"].apply(pd.to_numeric)
        collated_df["Cash Mix %"] = collated_df["Sales £"].apply(pd.to_numeric)
        collated_df["Unit Mix %"] = collated_df["Units"].apply(pd.to_numeric)
        sheetname = collated_df.index.get_level_values(0)[0]
        collated_df = collated_df.groupby(level=[1,2]).sum().reset_index()
        collated_df.set_index(["STORY","Item L3 Desc"], inplace=True)
        stories = list(collated_df.index.unique(level='STORY'))
        collated_df_total = collated_df[collated_df.index.get_level_values("Item L3 Desc") == "Total"]
        cash_total = collated_df["Sales £"].loc[("Total","Total")]
        print(cash_total)
        unit_total = collated_df["Units"].loc[("Total","Total")]
        print(collated_df_total["Cash Mix %"])
        collated_df_total["Cash Mix %"] = collated_df_total["Cash Mix %"].apply(lambda x: round(float(x / cash_total),3))
        collated_df_total["Unit Mix %"] = collated_df_total["Unit Mix %"].apply(lambda x: round(float(x / unit_total),3))
        collated_df_total.sort_values(by=["Cash Mix %"], inplace=True, ascending=False)
        collated_df_total.to_excel(writer, engine='xlsxwriter', sheet_name=sheetname + "_stories")

        totals = []
        unit_totals = []

        for story in stories:
            totals.append(collated_df["Cash Mix %"].loc[(story,"Total")])
            unit_totals.append(collated_df["Unit Mix %"].loc[(story,"Total")])

        items = list(collated_df.index.unique(level='Item L3 Desc'))
        
        for s,t in zip(stories,totals):
            for item in items:
                try:
                    collated_df["Cash Mix %"].loc[(s,item)] = round(float(collated_df["Cash Mix %"].loc[(s,item)] / t),3)
                except:
                    pass
        
        for s,ut in zip(stories,unit_totals):
            for item in items:
                try:
                    collated_df["Unit Mix %"].loc[(s,item)] = round(float(collated_df["Unit Mix %"].loc[(s,item)] / ut),3)
                except:
                    pass
        
        # collated_df.reindex(collated_df_total.index)

        collated_df.to_excel(writer, engine='xlsxwriter', sheet_name=sheetname)

    





