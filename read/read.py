import pandas as pd
import numpy as np
import os
import warnings
import dateutil.parser
from datetime import datetime
import editdistance
import re

warnings.simplefilter(action='ignore', category=UserWarning)

fileRead = 'LPE' #change variable value, possible values: LPE, LRE, LPI, LRI

'''
Error notes:
1. error: expected <class 'openpyxl.worksheet.cell_range.MultiCellRange'>
Straight up can't open
2. error: Can't find workbook in OLE2 compound document
Excel has password
'''

templateFile = '' 
replacementFile = ''
fileFormatName = ''
sheetFormatName = ''
wrongRencanaFile = ''
wrongRealisasiFile = ''
wrongFile = ''
wrongNamaFile = ''

if fileRead == 'LPE':
    templateFile = 'lpeplus_template.xlsx'
    replacementFile = 'lpe_replace.xlsx'
    fileFormatName = '_lpe'
    sheetFormatName = 'REALISASI'
    wrongRencanaFile = '_lri'
    wrongRealisasiFile = '_lpi'
    wrongFile = '_lre'
    wrongNamaFile = 'LRE'
elif fileRead == 'LRE':
    templateFile = 'lreplus_template.xlsx'
    replacementFile = 'lre_replace.xlsx'
    fileFormatName = '_lre'
    sheetFormatName = 'RENCANA'
    wrongRencanaFile = '_lri'
    wrongRealisasiFile = '_lpi'
    wrongFile = '_lpe'
    wrongNamaFile = 'LPE'
elif fileRead == 'LPI':
    templateFile = 'lpiplus_template.xlsx'
    replacementFile = 'lpi_replace.xlsx'
    fileFormatName = '_lpi'
    sheetFormatName = 'Laporan Inklusi Keuangan'
    wrongRencanaFile = '_lre'
    wrongRealisasiFile = '_lpe'
    wrongFile = '_lri'
    wrongNamaFile = 'LRI'
elif fileRead == 'LRI':
    templateFile = 'lriplus_template.xlsx'
    replacementFile = 'lri_replace.xlsx'
    fileFormatName = '_lri'
    sheetFormatName = 'Laporan Inklusi Keuangan'
    wrongRencanaFile = '_lre'
    wrongRealisasiFile = '_lpe'
    wrongFile = '_lpi'
    wrongNamaFile = 'LPI'
else:
    print('Salah file, pilihan file: LPE, LRE, LPI, LRI')
    exit(1)

def skipEmptyRow(df, repl):
    for w in range(0, len(df.index)):
        if (df.iloc[w,[0]].values[0] == 'No.'):
            v = w + 1
            df = df[v:]
            df.columns = repl.columns
            break
    return df

template = pd.read_excel(templateFile)
replacement = pd.read_excel(replacementFile)
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
        regex = '\d{2}-\d{2}-\d{4} \d{2}-\d{2}-\d{2}'
        match = re.search(regex, sub)
        tanggalLapor = match.group()
        for file in files:
            print('--'+file)
            if (fileFormatName in file and '.xls' not in file):
                notes.append(folder + '\\' + sub + '\\' + file + ' -> non .xls file format')
            if ('.zip' in file or '.rar' in file):
                notes.append(folder + '\\' + sub + '\\' + file + ' -> file is compressed in zip/rar')
            if ('.xls' in file):
                pos = file.lower().find(fileFormatName)
                if (pos != -1):
                    try:
                        if (fileRead in ['LRI', 'LPI']):
                            try:
                                data = pd.read_excel(file, header=3, sheet_name='Laporan Inklusi Keuangan')
                            except:
                                data = pd.read_excel(file, header=3)
                        else:
                            data = pd.read_excel(file, sheet_name=None)
                        if (sheetFormatName in data.keys()) or (fileRead in ['LRI', 'LPI']):
                            fileFiltered = file[:pos]
                            pujkName = fileFiltered.replace('_',' ')
                            namaFile = fileRead
                            if (fileRead in ['LRE', 'LPE']):
                                data = data[sheetFormatName]
                            thisColumns = list(data.columns.values)
                            if (thisColumns[0] != 'No.'):
                                data = skipEmptyRow(data, replacement)
                            if (len(thisColumns) == replacementLen):
                                data.columns = replacement.columns
                            data = data.drop(columns=['No.'])
                            if (fileRead in ['LRE', 'LPE']):
                                allData = data[data['Cakupan Kegiatan'].notnull()]
                            elif (fileRead in ['LRI', 'LPI']):
                                allData = data[data['Nama Kegiatan'].notnull()]
                            allData.insert(0, 'Nama File Excel', None)
                            allData.insert(0, 'Nama File', None)
                            allData.insert(0, 'Tanggal Lapor', None)
                            allData.insert(0, 'Nama Sektor Pelapor', None)
                            allData.insert(0, 'Nama PUJK Pelapor', None)
                            allData.insert(0, 'No.', None)
                            rows = len(allData.index)
                            if (rows > 0):
                                thisColumns = list(allData.columns.values)
                                if (not np.array_equal(validationColumns, thisColumns)):
                                    notes.append(folder + '\\' + sub + '\\' + file + ' -> ' + sheetFormatName + ' wrong column format')
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
                                        allData.loc[rowIndex,'Nama File Excel'] = sub + '\\' + file
                                    template = pd.concat([template, allData])
                            else:
                                notes.append(folder + '\\' + sub + '\\' + file + ' -> ' + sheetFormatName + ' sheet empty')
                        else:
                            notes.append(folder + '\\' + sub + '\\' + file + ' -> ' + sheetFormatName + ' sheet does not exist')
                    except Exception as e:
                        notes.append(folder + '\\' + sub + '\\' + file + ' -> error: ' + str(e))
                else:
                    try:
                        lri = file.lower().find(wrongRencanaFile)
                        lpi = file.lower().find(wrongRealisasiFile)
                        if (lri == -1 and lpi == -1):
                            if (fileRead in ['LRI', 'LPI']):
                                try:
                                    data = pd.read_excel(file, header=3, sheet_name='Laporan Inklusi Keuangan')
                                except:
                                    data = pd.read_excel(file, header=3)
                            else:
                                data = pd.read_excel(file, sheet_name=None)
                            pos = file.lower().find(wrongFile)
                            if (pos != -1):
                                namaFile = wrongNamaFile
                            else:
                                namaFile = 'Other'
                            if (sheetFormatName in data.keys()) or (fileRead in ['LRI', 'LPI']):
                                if (fileRead in ['LRE', 'LPE']):
                                    data = data[sheetFormatName]
                                rr = '^[^\d{2}]*'
                                pujkName = re.search(rr, sub).group(0)
                                #pujkName = sub[:-19]
                                thisColumns = list(data.columns.values)
                                if (thisColumns[0] != 'No.'):
                                    data = skipEmptyRow(data, replacement)
                                if (len(thisColumns) == replacementLen):
                                    data.columns = replacement.columns
                                data = data.drop(columns=['No.'])
                                if (fileRead in ['LRE', 'LPE']):
                                    allData = data[data['Cakupan Kegiatan'].notnull()]
                                elif (fileRead in ['LRI', 'LPI']):
                                    allData = data[data['Nama Kegiatan'].notnull()]
                                allData.insert(0, 'Nama File Excel', None)
                                allData.insert(0, 'Nama File', None)
                                allData.insert(0, 'Tanggal Lapor', None)
                                allData.insert(0, 'Nama Sektor Pelapor', None)
                                allData.insert(0, 'Nama PUJK Pelapor', None)
                                allData.insert(0, 'No.', None)
                                rows = len(allData.index)
                                if (rows > 0):
                                    thisColumns = list(allData.columns.values)
                                    if (not np.array_equal(validationColumns, thisColumns)):
                                        notes.append(folder + '\\' + sub + '\\' + file + ' -> ' + sheetFormatName + ' wrong column format')
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
                                            allData.loc[rowIndex,'Nama File Excel'] = sub + '\\' + file
                                        template = pd.concat([template, allData])                                
                                else:
                                    notes.append(folder + '\\' + sub + '\\' + file + ' -> ' + sheetFormatName + ' sheet empty')
                    except Exception as e:
                        notes.append(folder + '\\' + sub + '\\' + file + ' -> error: ' + str(e))
        os.chdir('..')
    os.chdir('..')
print('Done reading...')
#print('Exporting...')
template.to_excel(fileRead + " Result Raw.xlsx", index=False) 
errorNotes = []
for line in notes:
    row = line.split('->')
    errorNotes.append(row)
df = pd.DataFrame(np.array(errorNotes), columns=['File', 'Error'])
df.to_excel(fileRead + ' Mistake Notes.xlsx', 'Sheet1', index=False)

print('Cleaning...')

def formatDate(date):
    return datetime.strptime(date, '%d-%m-%Y %H-%M-%S')
    
def cleanName(name):
    i = name.lower().find('_')
    if (i != -1):
        name = name[:i]
    i = name.lower().find('lre')
    if (i != -1):
        name = name[:i]
    i = name.lower().find('lpe')
    if (i != -1):
        name = name[:i]
    i = name.lower().find('smt')
    if (i != -1):
        name = name[:i]
    i = name.lower().find('2023')
    if (i != -1):
        name = name[:i]
    return name.replace('_',' ')


def searchSektor(name):
    mins = 99999
    correctName = ''
    z = -1
    x = -1
    for j in pujk:
        x = x + 1
        newMins = editdistance.eval(name.lower(), j.lower())
        if (newMins < mins):
            mins = newMins
            correctName = j 
            z = x
    if (mins < 10):
        return sektor[z]
    else:
        return 'N/A'
        
def searchNama(name):
    mins = 99999
    correctName = ''
    z = -1
    x = -1
    for j in pujk:
        x = x + 1
        newMins = editdistance.eval(name.lower().strip(), j.lower())
        if (newMins < mins):
            mins = newMins
            correctName = j 
            z = x
    if (mins < 6):
        return pujk[z]
    else:
        return 'N/A'

def getSektor(name):
    i = name.lower().find('bpd')
    if (i != -1):
        return 'Bank Umum Konvensional'
    i = name.lower().find('bpr')
    if (i != -1):
        return 'Bank Perkreditan Rakyat Konvensional'
    return searchSektor(cleanName(name))

#1. Clean By Timestamp    
masterSektor = pd.read_excel('master 2324 w fintech.xlsx', na_filter=False)
masterTimestamp = pd.read_excel('master timestamp.xlsx', na_filter=False)
data = pd.read_excel(fileRead + ' Result Raw.xlsx', na_filter=False)
data['Tanggal Lapor'] = data['Tanggal Lapor'].apply(formatDate)
rows = len(data.index)
for i in range(0, rows):
    rowIndex = data.index[i]
    timestamp = data.loc[rowIndex,'Tanggal Lapor']
    timestampData = masterTimestamp[masterTimestamp['Waktu_Diterima'] == timestamp]
    namaPUJK = data.loc[rowIndex,'Nama PUJK Pelapor']
    sektorPUJK = 'Tanggal Lapor Not Found'
    if (len(timestampData.index) == 1):
        namaPUJK = timestampData['Nama_Fix'].iloc[0]
        sektorData = masterSektor[masterSektor['Nama PUJK'] == namaPUJK]
        if (len(sektorData.index) == 1):
            sektorPUJK = sektorData['Nama Sektor'].iloc[0]
    data.at[i, 'Nama PUJK Pelapor'] = namaPUJK
    data.at[i, 'Nama Sektor Pelapor'] = sektorPUJK
    
#2. Clean By Relasi Email
master = pd.read_excel('master 2324 w fintech.xlsx', na_filter=False)
pujk = master["Nama Lower"].values.tolist()
sektor = master["Nama Sektor"].values.tolist()

mastertoo = pd.read_excel('Data Relasi Email - PUJK.xlsx', na_filter=False)
mastertoo['Nama_PUJK'] = mastertoo['Nama_PUJK'].apply(str.lower).apply(str.strip)
mastertoo['Nama_Pengirim'] = mastertoo['Nama_Pengirim'].apply(str.lower).apply(str.strip)
mastertoo['Email_Pengirim'] = mastertoo['Email_Pengirim'].apply(str.lower).apply(str.strip)

rows = len(data.index)
for i in range(0, rows):
    namaPUJKFix = 'bup'
    rowIndex = data.index[i]
    namaPUJK = data.loc[rowIndex,'Nama PUJK Pelapor']
    namaFolder = data.loc[rowIndex,'Nama File Excel'].split('\\')[0].strip()
    rr = '^[^\d{2}]*'
    nama1 = re.search(rr, namaFolder).group(0)
    rrr = '(?<=\d{2}-\d{2}-\d{4} \d{2}-\d{2}-\d{2}).*$'
    nama2 = re.search(rrr, namaFolder).group(0).strip()
    newName = namaPUJK
    masterFix = mastertoo[mastertoo['Nama_PUJK'] == nama1.lower()]
    if (not masterFix.empty):
        namaPUJKFix = masterFix['Nama_Fix'].iloc[0]
    else:
        masterFix = mastertoo[mastertoo['Nama_Pengirim'] == nama1.lower()]
        if (not masterFix.empty):
            namaPUJKFix = masterFix['Nama_Fix'].iloc[0]
        else:
            masterFix = mastertoo[mastertoo['Email_Pengirim'] == nama1.lower()]
            if (not masterFix.empty):
                namaPUJKFix = masterFix['Nama_Fix'].iloc[0]
    if namaPUJKFix == 'bup' and nama2 != '':
        masterFix = mastertoo[mastertoo['Nama_Pengirim'] == nama2.lower()]
        if (not masterFix.empty):
            namaPUJKFix = masterFix['Nama_Fix'].iloc[0]
        else:
            masterFix = mastertoo[mastertoo['Email_Pengirim'] == nama2.lower()]
            if (not masterFix.empty):
                namaPUJKFix = masterFix['Nama_Fix'].iloc[0]
    if (namaPUJKFix != 'bup'):
        newName = namaPUJKFix
    newName = searchNama(newName)
    if (newName != 'N/A'):
        data.at[i, 'Nama PUJK Pelapor'] = newName
        
#3. Get Sektor
rows = len(data.index)
for i in range(0, rows):
    rowIndex = data.index[i]
    namaPUJK = data.loc[rowIndex,'Nama PUJK Pelapor']
    sektorPUJK = data.loc[rowIndex,'Nama Sektor Pelapor']
    if (sektorPUJK == 'Tanggal Lapor Not Found'):
        masterF = master[master['Nama Lower'] == namaPUJK.lower()]
        if (len(masterF.index) > 0):
            sektorPUJK = masterF['Nama Sektor'].iloc[0]
        else:
            newSektor = getSektor(namaPUJK)
            if newSektor != 'N/A':
                sektorPUJK = 'Guess - ' + newSektor
        data.at[i, 'Nama PUJK Pelapor'] = namaPUJK
        data.at[i, 'Nama Sektor Pelapor'] = sektorPUJK
        
#4. Clean Duplicate
columns = list(data.columns)
columns.remove("No.")
columns.remove("Tanggal Lapor")
columns.remove("Nama File")
columns.remove("Nama File Excel")
columns.remove("Nama Sektor Pelapor")
duplicate = data.drop_duplicates(subset = columns, keep = 'first')

print('Exporting...')
duplicate.to_excel(fileRead + " Result Final.xlsx", index=False)