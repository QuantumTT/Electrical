# Read Excel files to wrangle my data

import pandas as pd
# Determining which columns to read
columns = [0,1,2]

# Reading the workbook
workbook = pd.read_excel('./Sept. 9 Downstairs 1Hr.xlsx', sheet_name= 'Flipped')#, usecols=columns
workbook.head()

top_n = 5
key_dict = {}

for key, value in workbook.iteritems():
    print(key)
    # print(value)
    try:
        if 'Area' in key:
            area_dict = {}
            for index in value.index:
                area = value[index]
                area_dict[index] = area
    except:
        sorted= (value.sort_values(ascending = False))
        top_n_dict = {}
        area_indexes = []

        for index in sorted.index:
            if index != 0 and index != 1 and index != 2:
                area_indexes.append(index) 
        top_n_indexes = area_indexes[:top_n]

        for index in top_n_indexes:
            # print('index is ' + str(index))
            power = sorted[index]
            area = area_dict[index]
            top_n_dict[area] = power
        
        key_dict[str(key)] = top_n_dict


        mydataframe = pd.DataFrame(key_dict)
        # print(mydataframe)
        mydataframe.to_excel('Ranked_power_users.xlsx')

print('Finished!')