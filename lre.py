import pandas as pd
import os

#loop through folders
'''
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
'''
#read excel and combine with template
template = pd.read_excel('lre_template.xlsx')
data = pd.read_excel('BANK BENGKULU_LRE_SMT1_2023.xlsm', sheet_name=None)
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