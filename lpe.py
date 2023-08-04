import pandas as pd
import os
import warnings

warnings.simplefilter(action='ignore', category=UserWarning)

'''
progress:
poc using try and try2 folders success
1. read _lpe
2. rean non _lpe
3. error notes
start 11.13 - 11.27 ; 13.56 - 
'''

template = pd.read_excel('lpe_template.xlsx')
folders = [x for x in filter(os.path.isdir, os.listdir(os.getcwd()))]#change to the actual folders
#folders = ['try','try2']
pujkName = ''
sektorName = ''
tanggalLapor = ''
count = 0
curNum = ''
notes = []
for folder in folders:
    os.chdir(folder)
    subfolders = [x for x in filter(os.path.isdir, os.listdir(os.getcwd()))]#os.getcwd() current directory
    print(folder)
    for sub in subfolders:
        os.chdir(sub)
        files = os.listdir()
        print('-'+sub)
        tanggalLapor = sub[-19:]
        for file in files:
            print('--'+file)
            if ('_lpe' in file and '.xls' not in file):
                notes.append(folder + '\\' + sub + '\\' + file + ' -> non .xls file format')
            if ('.zip' in file or '.rar' in file):
                notes.append(folder + '\\' + sub + '\\' + file + ' -> file is compressed in zip/rar')
            if ('.xls' in file):
                pos = file.lower().find('_lpe')
                if (pos != -1):
                    try:
                        data = pd.read_excel(file, sheet_name=None)
                        if 'REALISASI' in data.keys():
                            fileFiltered = file[:pos]
                            pujkName = fileFiltered.replace('_',' ')
                            realisasi = data['REALISASI']
                            realisasi = realisasi.drop(columns=['No.'])
                            allData = realisasi[realisasi['Cakupan Kegiatan'].notnull()]
                            allData.insert(0, 'Tanggal Lapor', None)
                            allData.insert(0, 'Nama Sektor Pelapor', None)
                            allData.insert(0, 'Nama PUJK Pelapor', None)
                            allData.insert(0, 'No.', None)
                            rows = len(allData.index)
                            if (rows > 0):
                                for i in range(0, rows):
                                    count += 1
                                    curNum = str(count) + '.'
                                    rowIndex = allData.index[i]
                                    allData.loc[rowIndex,'No.'] = curNum
                                    allData.loc[rowIndex,'Nama PUJK Pelapor'] = pujkName
                                    allData.loc[rowIndex,'Nama Sektor Pelapor'] = sektorName
                                    allData.loc[rowIndex,'Tanggal Lapor'] = tanggalLapor
                                template = pd.concat([template, allData])
                            else:
                                notes.append(folder + '\\' + sub + '\\' + file + ' -> REALISASI sheet empty')
                    except:
                        notes.append(folder + '\\' + sub + '\\' + file + ' -> error reading file')
                '''
                only reading LPE
                else:
                    try:
                        data = pd.read_excel(file, sheet_name=None)
                        if 'REALISASI' in data.keys():
                            realisasi = data['REALISASI']
                            pujkName = sub[:-19]
                            realisasi = realisasi.drop(columns=['No.'])
                            allData = realisasi[realisasi['Cakupan Kegiatan'].notnull()]
                            allData.insert(0, 'Tanggal Lapor', None)
                            allData.insert(0, 'Nama Sektor Pelapor', None)
                            allData.insert(0, 'Nama PUJK Pelapor', None)
                            allData.insert(0, 'No.', None)
                            rows = len(allData.index)
                            if (rows > 0):
                                notes.append(folder + '\\' + sub + '\\' + file + ' -> REALISASI sheet exist and not empty but file name using _LRE')                                
                                #skip inserting to avoid duplicate with valid _LPE
                                #buttt, it can't save the one that has no _LPE
                                for i in range(0, rows):
                                    count += 1
                                    curNum = str(count) + '.'
                                    allData.loc[i,['No.']] = curNum
                                    allData.loc[i,['Nama PUJK Pelapor']] = pujkName
                                    allData.loc[i,['Nama Sektor Pelapor']] = sektorName
                                    allData.loc[i,['Tanggal Lapor']] = tanggalLapor
                                template = pd.concat([template, allData])                                
                    except:
                        notes.append(folder + '\\' + sub + '\\' + file + ' -> error reading file')
                '''
        os.chdir('..')
    os.chdir('..')
print('Done reading...')
print('Exporting...')
template.to_excel("LPE Result.xlsx", index=False) 
with open("LPE Mistake Notes.txt", "w") as txt_file:
    for line in notes:
        txt_file.write(line)
        txt_file.write("\n")


'''
#Proof Of Concept
#loop through folders
#folders = ['try','try2']#change to the actual folders
#folders = [x for x in filter(os.path.isdir, os.listdir(os.getcwd()))]
for folder in folders:
    os.chdir(folder)
    subfolders = [x for x in filter(os.path.isdir, os.listdir(os.getcwd()))]#os.getcwd() current directory
    subfolders.pop(0)
    print(folder)
    for sub in subfolders:
        os.chdir(sub)
        files = os.listdir()
        print('-'+sub)
        #for file in files:
        #    print('--'+file)
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

template.to_excel("LPE test.xlsx", index=False) 
print(realisasi.head())
'''
