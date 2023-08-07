import pandas as pd
import numpy as np
import os
import warnings

warnings.simplefilter(action='ignore', category=UserWarning)

'''
Error notes:
1. error: expected <class 'openpyxl.worksheet.cell_range.MultiCellRange'>
Straight up can't open
2. error: Can't find workbook in OLE2 compound document
Excel has password
'''


template = pd.read_excel('lriplus_template.xlsx')
replacement = pd.read_excel('lri_replace.xlsx')
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
            if ('_lri' in file and '.xls' not in file):
                notes.append(folder + '\\' + sub + '\\' + file + ' -> non .xls file format')
            if ('.zip' in file or '.rar' in file):
                notes.append(folder + '\\' + sub + '\\' + file + ' -> file is compressed in zip/rar')
            if ('.xls' in file):
                pos = file.lower().find('_lri')
                if (pos != -1):
                    try:
                        try:
                            data = pd.read_excel(file, header=3, sheet_name='Laporan Inklusi Keuangan')
                        except:
                            data = pd.read_excel(file, header=3)
                        fileFiltered = file[:pos]
                        pujkName = fileFiltered.replace('_',' ')
                        namaFile = 'LRI'
                        thisColumns = list(data.columns.values)
                        if (len(thisColumns) == replacementLen):
                            data.columns = replacement.columns
                        data = data.drop(columns=['No.'])
                        allData = data[data['Nama Kegiatan'].notnull()]
                        allData.insert(0, 'Nama File', None)
                        allData.insert(0, 'Tanggal Lapor', None)
                        allData.insert(0, 'Nama Sektor Pelapor', None)
                        allData.insert(0, 'Nama PUJK Pelapor', None)
                        allData.insert(0, 'No.', None)
                        rows = len(allData.index)
                        if (rows > 0):
                            thisColumns = list(allData.columns.values)
                            if (not np.array_equal(validationColumns, thisColumns)):
                                notes.append(folder + '\\' + sub + '\\' + file + ' -> wrong column format')
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
                            notes.append(folder + '\\' + sub + '\\' + file + ' -> sheet empty')
                    except Exception as e:
                        notes.append(folder + '\\' + sub + '\\' + file + ' -> error: ' + str(e))
                else:
                    try:
                        lre = file.lower().find('lre')
                        lpe = file.lower().find('lpe')
                        if (lre == -1 and lpe == -1):
                            try:
                                data = pd.read_excel(file, header=3, sheet_name='Laporan Inklusi Keuangan')
                            except:
                                data = pd.read_excel(file, header=3)
                            pos = file.lower().find('_lri')
                            if (pos != -1):
                                namaFile = 'LPI'
                            else:
                                namaFile = 'Other'
                            pujkName = sub[:-19]
                            thisColumns = list(data.columns.values)
                            if (len(thisColumns) == replacementLen):
                                data.columns = replacement.columns
                            data = data.drop(columns=['No.'])
                            allData = data[data['Nama Kegiatan'].notnull()]
                            allData.insert(0, 'Nama File', None)
                            allData.insert(0, 'Tanggal Lapor', None)
                            allData.insert(0, 'Nama Sektor Pelapor', None)
                            allData.insert(0, 'Nama PUJK Pelapor', None)
                            allData.insert(0, 'No.', None)
                            rows = len(allData.index)
                            if (rows > 0):
                                thisColumns = list(allData.columns.values)
                                if (not np.array_equal(validationColumns, thisColumns)):
                                    notes.append(folder + '\\' + sub + '\\' + file + ' -> wrong column format')
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
                                notes.append(folder + '\\' + sub + '\\' + file + ' -> sheet empty')
                    except Exception as e:
                        notes.append(folder + '\\' + sub + '\\' + file + ' -> error: ' + str(e))
        os.chdir('..')
    os.chdir('..')
print('Done reading...')
print('Exporting...')
template.to_excel("LRI Plus Result.xlsx", index=False) 
with open("LRI Plus Mistake Notes.txt", "w") as txt_file:
    for line in notes:
        txt_file.write(line)
        txt_file.write("\n")

