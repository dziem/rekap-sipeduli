'''
Nama Kegiatan = nil, nihil, -, blank
Inisiasi = Valid (min ada 1 kegiatan yang valid)
'''

import pandas as pd
import numpy as np
import dateutil.parser
from datetime import datetime
import validators
from IPython.display import display, HTML

start_of_feb = datetime.strptime('01-02-2023 00-00-00', '%d-%m-%Y %H-%M-%S')
start_of_may = datetime.strptime('01-05-2023 00-00-00', '%d-%m-%Y %H-%M-%S')
start_of_aug = datetime.strptime('01-08-2023 00-00-00', '%d-%m-%Y %H-%M-%S')
start_of_nov = datetime.strptime('01-11-2023 00-00-00', '%d-%m-%Y %H-%M-%S')

end_of_apr_str = '30-04-2023 23-59-59'
end_of_apr = datetime.strptime(end_of_apr_str, '%d-%m-%Y %H-%M-%S')

end_of_june_str = '30-06-2023 23-59-59'
end_of_june = datetime.strptime(end_of_june_str, '%d-%m-%Y %H-%M-%S')

end_of_july_str = '31-07-2023 23-59-59'
end_of_july = datetime.strptime(end_of_july_str, '%d-%m-%Y %H-%M-%S')

end_of_oct_str = '31-10-2023 23-59-59'
end_of_oct = datetime.strptime(end_of_oct_str, '%d-%m-%Y %H-%M-%S')

end_of_nov_str = '30-11-2023 23-59-59'
end_of_nov = datetime.strptime(end_of_nov_str, '%d-%m-%Y %H-%M-%S')

end_of_jan_str = '31-01-2024 23-59-59'
end_of_jan = datetime.strptime(end_of_jan_str, '%d-%m-%Y %H-%M-%S')

end_of_dec22_str = '31-12-2022 23-59-59'
end_of_dec22 = datetime.strptime(end_of_dec22_str, '%d-%m-%Y %H-%M-%S')

smt1 = end_of_july
smt2 = end_of_jan
tw1 = end_of_apr
tw2 = end_of_july
tw3 = end_of_oct
tw4 = end_of_jan
th2023 = end_of_jan
dec22 = end_of_dec22

def checkEmptyString(string):
    if ((not(string and string.strip())) or (string.lower().strip() in ['nihil', 'nil', '-']) or ('nihil' in string.lower())):
        return True
    else:
        return False

def correctPeriode(date, periode, sektor):
    if ('Guess' in sektor):
        sektor = sektor[8:]
    if sektor in ['Bank Umum Syariah','Bank Umum Konvensional']:
        if (periode == 'SMT1'):
            return 'TW2'
        elif (periode == 'SMT2'):
            return 'TW4'
    elif sektor in ['Manajer Investasi','Perantara Pedagang Efek yang Mengadministrasikan Rekening Efek Nasabah']:
        return 'Thn2023'
    elif sektor in ['Bank Perkreditan Rakyat Konvensional','Bank Perkreditan Rakyat Syariah']:
        if (periode == 'Thn2023'):
            return guessPeriode(date, sektor)
    return periode

def guessPeriode(date, sektor):
    if ('Guess' in sektor):
        sektor = sektor[8:]
    if (date >= start_of_feb) and (date <= end_of_apr):
        if sektor in ['Bank Umum Syariah','Bank Umum Konvensional']:
            return 'TW1'
        elif sektor in ['Manajer Investasi','Perantara Pedagang Efek yang Mengadministrasikan Rekening Efek Nasabah']:
             return 'Thn2023'
        return 'SMT1'
    elif (date >= start_of_may) and (date <= end_of_july):
        if sektor in ['Bank Umum Syariah','Bank Umum Konvensional']:
            return 'TW2'
        elif sektor in ['Manajer Investasi','Perantara Pedagang Efek yang Mengadministrasikan Rekening Efek Nasabah']:
             return 'Thn2023'
        return 'SMT1'
    elif (date >= start_of_aug) and (date <= end_of_oct):
        if sektor in ['Bank Umum Syariah','Bank Umum Konvensional']:
            return 'TW3'
        elif sektor in ['Manajer Investasi','Perantara Pedagang Efek yang Mengadministrasikan Rekening Efek Nasabah']:
             return 'Thn2023'
        return 'SMT2'
    elif (date >= start_of_nov) and (date <= end_of_jan):
        if sektor in ['Bank Umum Syariah','Bank Umum Konvensional']:
            return 'TW4'
        elif sektor in ['Manajer Investasi','Perantara Pedagang Efek yang Mengadministrasikan Rekening Efek Nasabah']:
             return 'Thn2023'
        return 'SMT2'
    else:
        return 'Invalid'

def getPeriode(fileName, date, sektor):
    fileName = fileName.replace(" ", "")        
    if (('triwulaniiii' in fileName.lower())
    or ('triwulan4' in fileName.lower())
    or ('twiiii' in fileName.lower())
    or ('tw4' in fileName.lower())):
        return 'TW4'
    if (('triwulaniii' in fileName.lower())
    or ('triwulan3' in fileName.lower())
    or ('twiii' in fileName.lower())
    or ('tw3' in fileName.lower())):
        return 'TW3'
    if (('triwulanii' in fileName.lower())
    or ('triwulan2' in fileName.lower())
    or ('twii' in fileName.lower())
    or ('tw2' in fileName.lower())):
        return 'TW2'
    if (('triwulani' in fileName.lower())
    or ('triwulan1' in fileName.lower())
    or ('twi' in fileName.lower())
    or ('tw1' in fileName.lower())):
        return 'TW1'
    if (('semesterii' in fileName.lower())
    or ('semester2' in fileName.lower())
    or ('smtii' in fileName.lower())
    or ('smt2' in fileName.lower())
    or ('smii' in fileName.lower())
    or ('sm2' in fileName.lower())):
        return 'SMT2'
    if (('semesteri' in fileName.lower())
    or ('semester1' in fileName.lower())
    or ('smti' in fileName.lower())
    or ('smt1' in fileName.lower())
    or ('smi' in fileName.lower())
    or ('sm1' in fileName.lower())):
        return 'SMT1'
    if (('tahun2023' in fileName.lower())
    or ('th2023' in fileName.lower())):
        return 'Thn2023'
    return guessPeriode(date, sektor)

def initiateDict(sektor): 
    if ('Guess' in sektor):
        sektor = sektor[8:]
    if sektor in ['Bank Umum Syariah','Bank Umum Konvensional']:
        return {"TW1": 0, "TW2": 0,"TW3": 0, "TW4": 0}
    elif sektor in ['Manajer Investasi','Perantara Pedagang Efek yang Mengadministrasikan Rekening Efek Nasabah']:
        return {"Thn2023": 0}
    else:
        return {"SMT1": 0, "SMT2": 0}
    
def updateDict(date, sektor, bentuk, nama, file, jumlahDict, edukasi): #edukasi = 1; infra = 0
    try:
        periode = getPeriode(file, date, sektor)
        if (periode != 'Invalid'):
            bentuk = str(bentuk)
            nama = str(nama)
            periode = correctPeriode(date, periode, sektor)
            if (not checkEmptyString(nama)) and (not checkEmptyString(bentuk)):
                if (edukasi == 1) and (bentuk.lower().strip() == 'edukasi keuangan'):
                    jumlahDict[periode] = jumlahDict[periode] + 1
                if (edukasi == 0) and (bentuk.lower().strip() == 'Pengembangan Sarana dan Prasarana yang mendukung literasi keuangan bagi konsumen dan/atau masyarakat'.lower()):
                    jumlahDict[periode] = jumlahDict[periode] + 1
            return jumlahDict
        return jumlahDict
    except Exception as e:
        #print(e)
        print(file)
        print(sektor)
        print(date)
        print(getPeriode(file, date, sektor))
        print(correctPeriode(date, periode, sektor))
        print('-----')
        return jumlahDict
    
#rule5(data.loc[rowIndex,'Inisiasi Kegiatan'], data.loc[rowIndex,'Kolaborasi Pelaksanaan'], data.loc[rowIndex,'Nama PUJK'])
def inisiasiKegiatan(inisiasi, kolaborasi, nama, pujk): #pujk = 1, otoritas = 0
    col1 = str(inisiasi)
    col2 = str(kolaborasi)
    col3 = str(nama)
    if (not checkEmptyString(inisiasi)) and (col2.lower().strip() == 'ya'):
        if (col1.lower().strip() == 'pujk' and pujk == 1):
            comma = col3.split(',')
            semicolon = col3.split(';')
            dan = col3.split(' dan')
            if (len(comma) > 3 or len(semicolon) > 3 or len(dan) > 3):
                return 'Invalid'
            else:
                return 'Valid'
        elif (col1.lower().strip() == 'pemerintah/otoritas' and pujk == 0):
            return 'Ada Kegiatan Bersama Pemerintah'
    elif (not checkEmptyString(inisiasi)) and (col2.lower().strip() == 'tidak'):
        if (pujk == 1):
            return 'Valid'
    if (pujk == 0):
        return 'Tidak Ada Kegiatan Bersama Pemerintah'
    return 'Invalid'
    
def toString(jumlahDict):
    res = ''
    i = 0
    for key, value in jumlahDict.items():
        i += 1
        res += key + ': ' + str(value)
        if (i < len(jumlahDict)):
            res += '\n'
    return res
    
def validation(edukasi, infra, inisiatif):
    checkEdu = []
    checkInfra = []
    lastEdu = False
    lastInfra = False
    lastPeriod = ''
    validEdu = False
    i = 0
    for key, value in edukasi.items():
        lastPeriod = key
        i += 1
        if value > 0:
            checkEdu.append(True)
            validEdu = True
            if (i == len(edukasi)):
                lastEdu = True
        else:
            checkEdu.append(False)
    i = 0
    for key, value in infra.items():
        i += 1
        if value > 0:
            checkInfra.append(True)
            if (i == len(infra)):
                lastInfra = True
        else:
            checkInfra.append(False)
    results = False
    lastTW = False
    resultsAll = []
    results = True
    for i in range(len(checkEdu) - 1):
        if (checkEdu[i] or checkInfra[i]):
            resultsAll.append(True)
            lastTW = True
        else:
            resultsAll.append(False)
            lastTW = False
    if (len(resultsAll) > 2):
        results = resultsAll[0] or resultsAll[1]
    else:
        for i in resultsAll:
            results = results and i
    '''
    for i in range(len(checkEdu) - 1):
        if (checkEdu[i] or checkInfra[i]):
            if ((i + 1) % 2 == 0):
                if ((i + 1) == 2):
                    results = lastTW and True
                else:
                    results = lastTW or True
                lastTW = False
            else:
                lastTW = True
                results = True
        else:
            if (i != 2):
                results = False
            if ((i + 1) % 2 == 0):
                if ((i + 1) == 2):
                    results = lastTW and False
                else:
                    results = lastTW or False
                if (results == False):
                    break
    '''
    if (len(checkEdu) == 1):
        results = checkEdu[0] or checkInfra[0]
    if (results):
        if (validEdu):
            if (len(checkEdu) > 2):
                if (inisiatif == 'Valid' and (resultsAll[2] or (lastEdu or lastInfra))):
                    return 'Valid'
            else:
                if (inisiatif == 'Valid' and (lastEdu or lastInfra)):
                    return 'Valid'
        return 'Menunggu Pelaporan ' + lastPeriod
    else:    
        if (len(checkEdu) == 1):
            return 'Menunggu Pelaporan ' + lastPeriod
        return 'Invalid'
    
    
 
namaPUJK = ''
sektorPUJK = ''
tanggalLapor = ''    
template = pd.read_excel('lpelite_template.xlsx', na_filter=False)
data = pd.read_excel('LPE Result Final.xlsx', na_filter=False)
count = 0
rows = len(data.index)
#rows = 10
for i in range(0, rows):
    rowIndex = data.index[i]
    namaPUJK = data.loc[rowIndex,'Nama PUJK Pelapor']
    sektorPUJK = data.loc[rowIndex,'Nama Sektor Pelapor']
    tanggalLapor = data.loc[rowIndex,'Tanggal Lapor']
    rowTempIndex = template[template['Nama PUJK']==namaPUJK].index.values
    if (len(rowTempIndex) == 0):
        count += 1
        curNum = str(count) + '.'
        template.loc[len(template)] = [
            curNum,#No
            namaPUJK,#Nama PUJK
            sektorPUJK,#Nama Sektor
            initiateDict(sektorPUJK),#Jumlah Kegiatan Edukasi
            initiateDict(sektorPUJK),#Jumlah Kegiatan Infrastruktur
            'None',#Kerjasama Inisiatif Sendiri
            'None',#Kerjasama Inisiatif Pemerintah
            ''#Kesimpulan
        ]
    rowTempIndex = template[template['Nama PUJK']==namaPUJK].index.values[0]
    template.at[rowTempIndex, 'Jumlah Kegiatan Edukasi'] = updateDict(tanggalLapor, sektorPUJK, data.loc[rowIndex,'Cakupan Kegiatan'], data.loc[rowIndex,'Nama Kegiatan'], data.loc[rowIndex,'Nama File Excel'], template.loc[rowTempIndex, 'Jumlah Kegiatan Edukasi'], 1)
    template.at[rowTempIndex, 'Jumlah Kegiatan Infrastruktur'] = updateDict(tanggalLapor, sektorPUJK, data.loc[rowIndex,'Cakupan Kegiatan'], data.loc[rowIndex,'Nama Kegiatan'], data.loc[rowIndex,'Nama File Excel'], template.loc[rowTempIndex, 'Jumlah Kegiatan Infrastruktur'], 0)
    if (template.loc[rowTempIndex, 'Kerjasama Inisiatif Sendiri'] in ['None', 'Invalid']):
        template.at[rowTempIndex, 'Kerjasama Inisiatif Sendiri'] = inisiasiKegiatan(data.loc[rowIndex,'Inisiasi Kegiatan'], data.loc[rowIndex,'Kolaborasi Pelaksanaan'], data.loc[rowIndex,'Nama PUJK'], 1)
    if (template.loc[rowTempIndex, 'Kerjasama Inisiatif Pemerintah'] in ['None', 'Invalid']):
        template.at[rowTempIndex, 'Kerjasama Inisiatif Pemerintah'] = inisiasiKegiatan(data.loc[rowIndex,'Inisiasi Kegiatan'], data.loc[rowIndex,'Kolaborasi Pelaksanaan'], data.loc[rowIndex,'Nama PUJK'], 0)
    
rows = len(template.index)    
for i in range(0, rows):
    rowIndex = template.index[i]
    template.at[rowIndex, 'Kesimpulan'] = validation(template.loc[rowIndex, 'Jumlah Kegiatan Edukasi'], template.loc[rowIndex, 'Jumlah Kegiatan Infrastruktur'], template.loc[rowIndex, 'Kerjasama Inisiatif Sendiri'])
    template.at[rowIndex, 'Jumlah Kegiatan Edukasi'] = toString(template.loc[rowIndex, 'Jumlah Kegiatan Edukasi'])
    template.at[rowIndex, 'Jumlah Kegiatan Infrastruktur'] = toString(template.loc[rowIndex, 'Jumlah Kegiatan Infrastruktur'])

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('LPE Lite Validation Result.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
template.to_excel(writer, sheet_name='Sheet1', index=False)

workbook  = writer.book
worksheet = writer.sheets['Sheet1']
text_wrap_format = workbook.add_format({'text_wrap': True})
algin_cell_format = workbook.add_format()
algin_cell_format.set_align('top')
worksheet.set_column('A:XFD', None, algin_cell_format)
worksheet.set_column(3, 4, 25, text_wrap_format)
writer.save()

#template.to_excel("LPE Lite Validation Result.xlsx", index=False) 