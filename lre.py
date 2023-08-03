import pandas as pd
import os

'''
progress:
poc using try and try2 folders success

todo
save errors:
1. empty realisasi sheet_name
2. etc
'''
template = pd.read_excel('lre_template.xlsx')
folders = ['try','try2']#change to the actual folders
pujkName = ''
sektorName = ''
count = 0
curNum = ''
for folder in folders:
    os.chdir(folder)
    subfolders = [x[0] for x in os.walk(os.getcwd())]#os.getcwd() current directory
    subfolders.pop(0)
    print(folder)
    for sub in subfolders:
        os.chdir(sub)
        files = os.listdir()
        print('-'+sub)
        for file in files:
            print('--'+file)
            if ('.xls' in file):
                pos = file.lower().find('_lre')
                if (pos != -1):
                    data = pd.read_excel(file, sheet_name=None)
                    if 'REALISASI' in data.keys():
                        fileFiltered = file[:pos]
                        pujkName = fileFiltered.replace('_',' ')
                        realisasi = data['REALISASI']
                        realisasi = realisasi.drop(columns=['No.'])
                        allData = realisasi[realisasi['Cakupan Kegiatan'].notnull()]
                        allData.insert(0, 'Nama Sektor Pelapor', None)
                        allData.insert(0, 'Nama PUJK Pelapor', None)
                        allData.insert(0, 'No.', None)
                        rows = len(allData.index)
                        for i in range(0, rows):
                            count += 1
                            curNum = str(count) + '.'
                            allData.loc[i,['No.']] = curNum
                            allData.loc[i,['Nama PUJK Pelapor']] = pujkName
                            allData.loc[i,['Nama Sektor Pelapor']] = sektorName
                        template = pd.concat([template, allData])
        os.chdir('..')
    os.chdir('..')
template.to_excel("LRE test.xlsx", index=False) 

'''
Proof Of Concept
#loop through folders
folders = ['try','try2']#change to the actual folders
for folder in folders:
    os.chdir(folder)
    subfolders = [x[0] for x in os.walk(os.getcwd())]#os.getcwd() current directory
    subfolders.pop(0)
    print(folder)
    for sub in subfolders:
        os.chdir(sub)
        files = os.listdir()
        print('-'+sub)
        for file in files:
            print('--'+file)
        os.chdir('..')
    os.chdir('..')

#read excel and combine with template
template = pd.read_excel('lre_template.xlsx')
data = pd.read_excel('BANK MALUKU MALUT_LRE_2023.xlsm', sheet_name=None)
count = 0
if 'REALISASI' in data.keys():
    realisasi = data['REALISASI']
    realisasi = realisasi.drop(columns=['No.'])
    allData = realisasi[realisasi['Cakupan Kegiatan'].notnull()]
    allData.insert(0, 'Nama Sektor Pelapor', None)
    allData.insert(0, 'Nama PUJK Pelapor', None)
    allData.insert(0, 'No.', None)
    rows = len(allData.index)
    for i in range(0, rows):
        count += 1
        allData.loc[i,['No.']] = str(count) + '.'
        allData.loc[i,['Nama PUJK Pelapor']] = 'Nama'
        allData.loc[i,['Nama Sektor Pelapor']] = 'Sektor'
    template = pd.concat([template, allData])

template.to_excel("LRE test.xlsx", index=False) 
print(realisasi.head())
