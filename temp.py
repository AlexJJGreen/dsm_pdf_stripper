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
        

        #for i in range(col_len):
        #    if story_datasets[2][i] != "Total":
        #        story_datasets[1].insert(i,"")
