
import os
import shutil
import xlrd
import openpyxl
import pandas as pd
import numpy as np

def source_read():
    for i in os.listdir('source/'):
        print(i)
        print(i[-1:-5])
        df = pd.read_excel('source/'+ i)
        return df



def space(df):

    df2 = pd.DataFrame(columns=df.columns)

    for i in range(len(df.index)):
        row = df.loc[i].to_frame().T
        empty_rows = pd.DataFrame(np.nan, index = range(1), columns=df.columns)
        row = pd.concat([row, empty_rows]).reset_index(drop=True)
        df2 = pd.concat([df2, row]).reset_index(drop=True)

    return (df2)



def transform(df):

    format_list = [
        'MFI > 10001',
        'MFI 10000~6001',
        'MFI 6000~3001',
        'MFI 3000~1001',
        'MFI 1000~500',
        ''
    ]

    df_save = pd.DataFrame(columns=df.columns)
    df_save['MFI value'] = 0
    df_save['Antigens:MFI'] = 0
    #print(df_save, type(df_save))

    for i in range(len(df.index)):
        row = df.loc[i].to_frame().T
        empty_rows = pd.DataFrame(np.nan, index = range(5), columns=df.columns)
        #row = pd.concat([row, empty_rows]).sort_index(kind='mergesort').reset_index(drop=True) sort_index 有問題
        row = pd.concat([row, empty_rows]).reset_index(drop=True)
        row['MFI value'] = format_list
        row['Antigens:MFI'] =''
        row.at[0, 'Antigens:MFI'] = row.at[0, 'MFI > 10001']
        row.at[1, 'Antigens:MFI'] = row.at[0, 'MFI 10000~6001']
        row.at[2, 'Antigens:MFI'] = row.at[0, 'MFI 6000~3001']
        row.at[3, 'Antigens:MFI'] = row.at[0, 'MFI 3000~1001']
        row.at[4, 'Antigens:MFI'] = row.at[0, 'MFI 1000~500']
        print(row)
        df_save = pd.concat([df_save, row], ignore_index=True)


    return(df_save)



def summary_calc(input):

    input['Strength'] = ''
    for i in input.index:
        if input.loc[i, 'Summary'] == 'Null': #data2
            continue

        else:
            box = []
            boxs =[]
            boxa =[]
            for k, j in input.loc[i, 'Summary'].items():
                box.append(str(k) + ': ' +str(j))
                boxa.append(str(k))
                boxs.append(str(j))


            #input.at[i, 'Summary'] = str(box).strip('[').strip(']')
            #input.loc[i, 'Summary'] = str(input.loc[i, 'Summary']).strip('{').strip('}')
            input.at[i, 'Summary'] = '\n'.join(box)



            input.at[i, 'Strength'] = '\n'.join(boxs)
            input.at[i, 'Tail Antigens'] = '\n'.join(boxa)



    print(input)
                #print(k, j)
    return(input)






def backup():
    for i in os.listdir('source/'):
        shutil.move('source/' + i, 'backup/')


df = source_read()

filename_list = os.listdir('source/') # for export filename
filename = (filename_list[0])


print(df)

df1 = df.drop_duplicates(subset = 'Sample ID')
df1 = df1.drop('Antigen', axis = 1)

print(df1)



data = pd.DataFrame()


# 迴圈把各個 sample ID 的 'Tail Antigens' 抓出來，存入 lst [] 裡面
for i in df1.index:
    j = df1.loc[i, 'Sample ID' ]
    lst = []
    k = str(df1.loc[i, 'Tail Antigens' ])
    lst = str.split(k,', ')
    print(j)
    print(lst)
    dfx = (df['Sample ID'] == j)
    dfy = df.loc[dfx]
    dfy = dfy.drop_duplicates(subset='Antigen')
    #迴圈依據 'Tail Antigens' 部分有無分開處理
    for k in lst:
        if k == 'nan':
            line = dfy.head(1)
            line['Antigen'] = 'Null'
            line['Strength'] = 0
            print(line)
            data = pd.concat([data, line])

        else:
            line = dfy[(dfy['Antigen'] == k)]
            #line['aa'] = line['Antigen'].map(str) + ':' + line['Strength'].map(str)
            print(line)
            data = pd.concat([data, line])

print(data)


data.columns = [c.replace(' ', '_') for c in data.columns]  #把column空白去掉


# Group by "Sample_ID" and collect unique values for "Sample_ID", "Strength", and "Antigen"
#重要 ChatGPT
grouped_data = data.groupby("Sample_ID").agg({
    "Sample_ID": "unique",
    "Strength": list,
    "Antigen": list
})

# Separate the grouped data into individual Series
sample_ID_grouped = grouped_data["Sample_ID"]
strength_grouped = grouped_data["Strength"]
antigen_grouped = grouped_data["Antigen"]


print(antigen_grouped)
print(strength_grouped)
print(sample_ID_grouped)


index_list = sample_ID_grouped.index.tolist()
print(index_list)


#用 sample_ID_grouped 的儲存格塞 dict，放進 antigen_grouped[i][j] & strength_grouped[i][j]
for i in index_list:
    print(sample_ID_grouped[i])
    sample_ID_grouped[i] = {}
    #print(sample_ID_grouped[i])
    for j in range(len(strength_grouped[i])):
        a = antigen_grouped[i]
        s = strength_grouped[i]
        sample_ID_grouped[i][a[j]] = s[j]

        #print(antigen_grouped.values[j])
        #print(antigen_grouped.values[j])
        #print(strength_grouped.values[j])
      

print(sample_ID_grouped)

Batch_Name_grouped = data.groupby("Sample_ID").Batch_Name.unique()
Patient_Name_grouped = data.groupby("Sample_ID").Patient_Name.unique()
Tail_Antigens_grouped = data.groupby("Sample_ID").Tail_Antigens.unique()
PRA_grouped = data.groupby("Sample_ID").PRA.unique()


data2 =pd.concat([Batch_Name_grouped,  Patient_Name_grouped, Tail_Antigens_grouped, PRA_grouped, sample_ID_grouped],
                 axis = 1, ignore_index = True,
                 )

pd.set_option("display.max.columns", None)

data2.rename(columns={
    0: 'Batch Name',
    1: 'Patient Name',
    2: 'Tail Antigens',
    3: 'PRA',
    4: 'Summary'
}, inplace=True)

# Adding 'Sample ID' column
data2['Sample ID'] = data2.index

# Reordering the columns

data2 = data2[['Sample ID', 'Batch Name', 'Patient Name', 'Tail Antigens', 'PRA', 'Summary',]]

#data2['PRA'] = str(data2['PRA'][1:-1]) + '%'


# Print the updated DataFrame
print(data2)


for i in data2.index:
    if data2.loc[i, 'Summary'] == {'Null': 0} :
        data2.loc[i, 'Summary'] = 'Null'

    else:
        #data2.at[i, 'Antigens'].update({
        #    'MFI > 10001': 100000,
        #    'MFI 10000~6001': 10000,
        #    'MFI 6000~3001': 6000,
        #    'MFI 3000~1001': 3000,
        #    'MFI 1000~500': 1000
        #})
        # Sort the 'Antigens' dictionary by values in descending order 排序
        #j = sorted(data2.at[i, 'Summary'].items(), key=lambda item: item[1], reverse=True)
        j = data2.at[i, 'Summary'].items()

        # Update the 'Antigens' column with the sorted dictionary
        data2.at[i, 'Summary'] = dict(j)


print(data2[['Sample ID', 'Summary']])



#data2 = export(data2)



data2 = data2.reset_index(drop=True)



format_list = [
    'MFI > 10001',
    'MFI 10000~6001',
    'MFI 6000~3001',
    'MFI 3000~1001',
    'MFI 1000~500'
]


data2['Batch Name'] = data2['Batch Name'].astype(str).str[2:-2]
data2['Patient Name'] = data2['Patient Name'].astype(str).str[2:-2]
data2['PRA'] = data2['PRA'].astype(str).str[1:-1] + '%'
data2['Tail Antigens'] = data2['Tail Antigens'].astype(str).str[1:-1]



#轉換矩陣
#data2 = transform(data2)

# 搭配 transform()
#data2 = data2[['Sample ID', 'Batch Name', 'Patient Name', 'Tail Antigens', 'PRA', 'Summary','MFI value', 'Antigens:MFI']]

#data2 = space(data2)

data2['Patient Name & ID'] = data2['Patient Name'].astype(str) + ' ' + data2['Sample ID'].astype(str)

data2 = data2[['Sample ID', 'Patient Name', 'Batch Name', 'Patient Name & ID', 'PRA',  'Summary', 'Tail Antigens',
        #'MFI > 10001',
        #'MFI 10000~6001',
        #'MFI 6000~3001',
        #'MFI 3000~1001',
        #'MFI 1000~500',
         ]]


print(data2)


data3 = data2 #.T


data3 = summary_calc(data3)

data3.to_csv(filename[0:-4] + '轉檔' + '.xls', index=False, sep='\t')
#data3.to_excel(filename[0:-4] + '轉檔' + '.xlsx', index=False,)


print('檔案輸出為' + filename[0:-4] +'.xls')
backup()


