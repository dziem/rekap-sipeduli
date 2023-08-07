import pandas as pd
import numpy as np
import os
import warnings

warnings.simplefilter(action='ignore', category=UserWarning)

'''
progress:
poc using try and try2 folders success
1. read _lpe
2. rean non _lpe (skip lpi, lri)
3. error notes
start 15.49 - 16.22
16.45 - 17.16

Error notes:
1. error: expected <class 'openpyxl.worksheet.cell_range.MultiCellRange'>
Straight up can't open
2. error: Can't find workbook in OLE2 compound document
Excel has password
'''

def skipEmptyRow(df, repl):
    for w in range(0, len(df.index)):
        if (df.iloc[w,[0]].values[0] == 'No.'):
            v = w + 1
            df = df[v:]
            df.columns = repl.columns
            break
    return df

template = pd.read_excel('lpeplus_template.xlsx')
replacement = pd.read_excel('lpe_replace.xlsx')
replacementLen = len(list(replacement.columns.values))
folders = [x for x in filter(os.path.isdir, os.listdir(os.getcwd()))]#change to the actual folders
#folders = ['try','try2']
pujkName = ''
sektorName = ''
tanggalLapor = ''
namaFile = ''
count = 0
curNum = ''
notes = []
validationColumns = list(template.columns.values)
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
                            namaFile = 'LPE'
                            realisasi = data['REALISASI']
                            thisColumns = list(realisasi.columns.values)
                            if (thisColumns[0] != 'No.'):
                                realisasi = skipEmptyRow(realisasi, replacement)
                            if (len(thisColumns) == replacementLen):
                                realisasi.columns = replacement.columns
                            realisasi = realisasi.drop(columns=['No.'])
                            allData = realisasi[realisasi['Cakupan Kegiatan'].notnull()]
                            allData.insert(0, 'Nama File', None)
                            allData.insert(0, 'Tanggal Lapor', None)
                            allData.insert(0, 'Nama Sektor Pelapor', None)
                            allData.insert(0, 'Nama PUJK Pelapor', None)
                            allData.insert(0, 'No.', None)
                            rows = len(allData.index)
                            if (rows > 0):
                                thisColumns = list(allData.columns.values)
                                if (not np.array_equal(validationColumns, thisColumns)):
                                    notes.append(folder + '\\' + sub + '\\' + file + ' -> REALISASI wrong column format')
                                else:
                                    for i in range(0, rows):
                                        count += 1
                                        curNum = str(count) + '.'
                                        rowIndex = allData.index[i]
                                        allData.loc[rowIndex,'No.'] = curNum
                                        allData.loc[rowIndex,'Nama PUJK Pelapor'] = pujkName
                                        allData.loc[rowIndex,'Nama Sektor Pelapor'] = sektorName
                                        allData.loc[rowIndex,'Tanggal Lapor'] = tanggalLapor
                                        allData.loc[rowIndex,'Nama File'] = namaFile
                                    template = pd.concat([template, allData])
                            else:
                                notes.append(folder + '\\' + sub + '\\' + file + ' -> REALISASI sheet empty')
                        else:
                            notes.append(folder + '\\' + sub + '\\' + file + ' -> REALISASI sheet does not exist')
                    except Exception as e:
                        notes.append(folder + '\\' + sub + '\\' + file + ' -> error: ' + str(e))
                else:
                    try:
                        lri = file.lower().find('_lri')
                        lpi = file.lower().find('_lpi')
                        if (lri == -1 and lpi == -1):
                            data = pd.read_excel(file, sheet_name=None)
                            pos = file.lower().find('_lre')
                            if (pos != -1):
                                namaFile = 'LRE'
                            else:
                                namaFile = 'Other'
                            if 'REALISASI' in data.keys():
                                realisasi = data['REALISASI']
                                pujkName = sub[:-19]
                                thisColumns = list(realisasi.columns.values)
                                if (thisColumns[0] != 'No.'):
                                    realisasi = skipEmptyRow(realisasi, replacement)
                                if (len(thisColumns) == replacementLen):
                                    realisasi.columns = replacement.columns
                                realisasi = realisasi.drop(columns=['No.'])
                                allData = realisasi[realisasi['Cakupan Kegiatan'].notnull()]
                                allData.insert(0, 'Nama File', None)
                                allData.insert(0, 'Tanggal Lapor', None)
                                allData.insert(0, 'Nama Sektor Pelapor', None)
                                allData.insert(0, 'Nama PUJK Pelapor', None)
                                allData.insert(0, 'No.', None)
                                rows = len(allData.index)
                                if (rows > 0):
                                    thisColumns = list(allData.columns.values)
                                    if (not np.array_equal(validationColumns, thisColumns)):
                                        notes.append(folder + '\\' + sub + '\\' + file + ' -> REALISASI wrong column format')
                                    else:
                                        for i in range(0, rows):
                                            count += 1
                                            curNum = str(count) + '.'
                                            rowIndex = allData.index[i]
                                            allData.loc[rowIndex,'No.'] = curNum
                                            allData.loc[rowIndex,'Nama PUJK Pelapor'] = pujkName
                                            allData.loc[rowIndex,'Nama Sektor Pelapor'] = sektorName
                                            allData.loc[rowIndex,'Tanggal Lapor'] = tanggalLapor
                                            allData.loc[rowIndex,'Nama File'] = namaFile
                                        template = pd.concat([template, allData])                                
                                else:
                                    notes.append(folder + '\\' + sub + '\\' + file + ' -> REALISASI sheet empty')
                    except Exception as e:
                        notes.append(folder + '\\' + sub + '\\' + file + ' -> error: ' + str(e))
        os.chdir('..')
    os.chdir('..')
print('Done reading...')
print('Exporting...')
template.to_excel("LPE Plus Result.xlsx", index=False) 
with open("LPE Plus Mistake Notes.txt", "w") as txt_file:
    for line in notes:
        txt_file.write(line)
        txt_file.write("\n")
