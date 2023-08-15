import pandas as pd
import numpy as np
import dateutil.parser
from datetime import datetime

end_of_june_str = '30-06-2023 23-59-59'
end_of_june = datetime.strptime(end_of_june_str, '%d-%m-%Y %H-%M-%S')

end_of_july_str = '31-07-2023 23-59-59'
end_of_july = datetime.strptime(end_of_july_str, '%d-%m-%Y %H-%M-%S')

end_of_nov_str = '30-11-2023 23-59-59'
end_of_nov = datetime.strptime(end_of_nov_str, '%d-%m-%Y %H-%M-%S')

def translate(date):
    date = date.lower()
    if ('januari' in date):
        return date.replace('januari', 'january')
    elif ('februari' in date):
        return date.replace('februari', 'february')
    elif ('maret' in date):
        return date.replace('maret', 'march')
    elif ('mei' in date):
        return date.replace('mei', 'may')
    elif ('juni' in date):
        return date.replace('juni', 'june')
    elif ('juli' in date):
        return date.replace('juli', 'july')
    elif ('oktober' in date):
        return date.replace('oktober', 'october')
    elif ('desember' in date):
        return date.replace('desember', 'december')
    else:
        return date
        
def convertDate(date):
    try:
        return dateutil.parser.parse(translate(date))
    except Exception as e:
        print(e)
        return 'Invalid Date Format'

def checkEmptyString(string):
    if ((not(string and string.strip())) or (string.lower() == 'nihil')):
        return True
    else:
        return False

def formatDate(date):
    return datetime.strptime(date, '%d-%m-%Y %H-%M-%S')

#Rule Output: Valid (correct); Invalid (incorrect); Unapplicable

#Apakah PUJK melaksanakan kegiatan untuk meningkatkan literasi keuangan kepada Konsumen dan/atau masyarakat ?
#Immediately valid because row exist
def rule1(col):
    return 'Valid'

#Apakah kegiatan dilakukan paling sedikit 1 (satu) kali dalam 1 (satu) semester ?
#Row exist at least one
#End date <= 30 June = valid (and vice versa)
def rule2(col):
    dateConvert = convertDate(str(col))
    if (dateConvert == 'Invalid Date Format'):
        return 'Invalid'
    else:
        if (dateConvert <= end_of_june):
            return 'Valid'
        else:
            return 'Invalid'

#Apakah PUJK mendokumentasikan pelaksanaan kegiatan untuk meningkatkan literasi keuangan?
#If link dokumentasi col not empty
def rule3(col):
    if(not(isinstance(col, str))):
        return 'Invalid'
    elif(checkEmptyString(col)):
        return 'Invalid'
    else:
        return 'Valid'

#Apakah PUJK bekerjasama dengan maksimal 3 PUJK lainnya ?    
#Kolaborasi Pelaksanaan = 'Ya' and Inisiasi Kegiatan = 'PUJK' and Nama PUJK split(','/';'/' dan ') > 3
def rule5(col1, col2, col3): #inisiasi, kolaborasi, nama pujk
    col1 = str(col1)
    col2 = str(col2)
    col3 = str(col3)
    if (checkEmptyString(col1)):
        return 'Inisiasi Empty' 
    if (col1.lower() == 'ya'):
        if (col2.lower() == 'pujk'):
            comma = col3.split(',')
            semicolon = col3.split(';')
            dan = col3.split(' dan')
            if (comma > 3 or semicolon > 3 or dan > 3):
                return 'Invalid'
    return 'Valid'

#Apakah PUJK melaksanakan kegiatan Edukasi Keuangan paling sedikit 1 (satu) kali ? 
#Valid kegiatan is 'Edukasi Keuangan', invalid is else (invalid specifically for this row, not the entire pujk)
def rule6(col):
    col = str(col)
    if (col.lower() == 'edukasi keuangan'):
        return 'Valid'
    return 'Invalid'
    
'''
Apakah PUJK yang melaksanakan kegiatan Edukasi keuangan telah menyampaikan materi Edukasi Keuangan yang mencakup:
a. karakteristik sektor keuangan
b. karakteristik produk dan/atau layanan keuangan
c. materi pengelolaan keuangan
d. materi perpajakan
'''    
def rule7(col1,col2,col3,col4,col5):#Materi Karakteristik Sektor Jasa Keuangan,Materi Pengelolaan Keuangan,Materi Perpajakan,Materi Karakteristik Produk dan/atau Layanan,Inisiasi Kegiatan
    if (col5.lower() != 'Pemerintah/Otoritas'):
        if((not(isinstance(col1, str))) or (not(isinstance(col2, str))) or (not(isinstance(col3, str))) or (not(isinstance(col4, str)))):
            return 'Invalid'
        elif((checkEmptyString(col1)) or (checkEmptyString(col2)) or (checkEmptyString(col3)) or (checkEmptyString(col4))):
            return 'Invalid'
        return 'Valid'


#Apakah PUJK menyampaikan penyesuaian atau perubahan laporan rencana sesuai dengan ketentuan yang mengatur mengenai penyesuaian atau perubahan rencana bisnis masing-masing PUJK?    
#Not sure if the deadline is end of june
def rule13(col):
    dateConvert = datetime.strptime(col, '%d-%m-%Y %H-%M-%S')
    if (dateConvert == 'Invalid Date Format'):
        return 'Invalid'
    else:
        if (dateConvert <= end_of_june):
            return 'Valid'
        else:
            return 'Invalid'

#Same as 13 but only for some sektor pujk, need to ask which pujk
def rule14(col1, col2):#Tanggal Lapor, Nama Sektor Pelapor
    return 'No Idea'
    
#Laporan yang disampaikan apakah dilengkapi dengan surat pengantar yang ditandatangani oleh salah satu anggota Direksi ?
#Can't be determined using this excel file
def rule17():
    return 'Unapplicable'

#Apakah PUJK telah mencantumkan peran dan pihak lain pada laporan rencana dan realisasi ? (apabila PUJK bekerja sama dengan pihak lain)
'''
Apakah PUJK telah mencantumkan peran dan pihak lain pada laporan rencana dan realisasi ? (apabila PUJK bekerja sama dengan pihak lain)
Kolaborasi Pelaksanaan = 'Ya'
If Nama PUJK not empty Peran masing-masing pihak PUJK can't be empty 
If Nama pihak di luar PUJK not empty Peran masing-masing pihak di luar PUJK can't be empty 
'''    
def rule19(col1, col2, col3, col4, col5): #Kolaborasi Pelaksanaan, Nama PUJK, Peran masing-masing pihak PUJK, Nama pihak di luar PUJK, Peran masing-masing pihak di luar PUJK
    col1 = str(col1)
    col2 = str(col2)
    col3 = str(col3)
    col4 = str(col4)
    col5 = str(col5)
    if (col1 == 'Ya'):
        if (not checkEmptyString(col2)):
            if (checkEmptyString(col3)):
                return 'Invalid'
        if (not checkEmptyString(col4)):
            if (checkEmptyString(col5)):
                return 'Invalid'
    return 'Valid'
    
#Apakah PUJK telah menyusun dan menyampaikan laporan rencana dan realisasi atas kegiatan dalam rangka meningkatkan literasi keuangan? 
#Find Rencana by PUJK Name, If exist valid and vice versa
def rule20(col, data_plan):
    allData = data_plan[data_plan['Nama PUJK Pelapor'] == col]
    if (len(allData.index) > 0):
        return 'Valid'
    else:
        return 'Invalid'

#Apakah PUJK telah menyusun dan menyampaikan laporan rencana dan realisasi atas kegiatan dalam rangka meningkatkan literasi Keuangan sesuai ketentuan yang mengatur mengenai pelaporan rencana bisnis masing masing PUJK ?
def rule22(col): #Tanggal Lapor
    return 'No Idea'

#Apabila PUJK tidak memiliki rencana bisnis, apakah PUJK telah menyampaikan laporan rencana atas kegiatan dalam rangka meningkatkan literasi Keuangan paling lambat tanggal 30 bulan november sebelum tahun kegiatan?    
#Rencana next year (ie 2024), before end of november this year (ie 2023)
def rule23(col1, col2, data_plan): #Tanggal Lapor, Nama PUJK Pelapor
    allData = data_plan[data_plan['Nama PUJK Pelapor'] == col2]
    if (len(allData.index) > 0):
        allData = allData.sort_values(by=['Tanggal Lapor'], ascending=False)
        if (allData.iloc[0]['Tanggal Lapor'] <= end_of_nov):
            return 'Valid'
        else:
            return 'Invalid'
    else:
        return 'Rencana Not Found'

#Apabila PUJK tidak memiliki rencana bisnis, apakah PUJK telah menyampaikan laporan realisasi atas kegiatan dalam rangka meningkatkan literasi Keuangan paling lambat tanggal 31 Juli tahun berjalan dan tanggal 31 Januari tahun berikutnya ? 
#Before end of july
def rule24(col): #Tanggal Lapor
    dateConvert = datetime.strptime(col, '%d-%m-%Y %H-%M-%S')
    if (dateConvert == 'Invalid Date Format'):
        return 'Invalid'
    else:
        if (dateConvert <= end_of_july): #Not sure if <= this date
            return 'Valid'
        else:
            return 'Invalid'

template = pd.read_excel('lpe_template.xlsx', na_filter=False)
data = pd.read_excel('LPE Result.xlsx', na_filter=False)
data_plan = pd.read_excel('LRE Result.xlsx', na_filter=False)
data_plan['Tanggal Lapor'] = data_plan['Tanggal Lapor'].apply(formatDate)
count = 0
rows = len(data.index)
#for i in range(0, rows):
for i in range(0, 10):
    count += 1
    curNum = str(count) + '.'
    rowIndex = data.index[i]
    template.loc[len(template)] = [
        curNum, 
        data.loc[rowIndex,'Nama PUJK Pelapor'],
        data.loc[rowIndex,'Nama Sektor Pelapor'],
        data.loc[rowIndex,'Tanggal Lapor'],
        rule1(data.loc[rowIndex,'Nama PUJK Pelapor']),
        rule2(data.loc[rowIndex,'Tanggal Berakhir Kegiatan (DD/MM/YYYY)']),
        rule3(data.loc[rowIndex,'Link Dokumentasi']),
        rule5(data.loc[rowIndex,'Inisiasi Kegiatan'], data.loc[rowIndex,'Kolaborasi Pelaksanaan'], data.loc[rowIndex,'Nama PUJK']),
        rule6(data.loc[rowIndex,'Cakupan Kegiatan']),
        rule7(data.loc[rowIndex,'Materi Karakteristik Sektor Jasa Keuangan'], data.loc[rowIndex,'Materi Pengelolaan Keuangan'], data.loc[rowIndex,'Materi Perpajakan'], data.loc[rowIndex,'Materi Karakteristik Produk dan/atau Layanan'],data.loc[rowIndex,'Inisiasi Kegiatan']),
        rule13(data.loc[rowIndex,'Tanggal Lapor']),
        rule14(data.loc[rowIndex,'Tanggal Lapor'], data.loc[rowIndex,'Nama Sektor Pelapor']),
        rule17(),
        rule19(data.loc[rowIndex,'Kolaborasi Pelaksanaan'],data.loc[rowIndex,'Nama PUJK'],data.loc[rowIndex,'Peran masing-masing pihak PUJK'],data.loc[rowIndex,'Nama pihak di luar PUJK'],data.loc[rowIndex,'Peran masing-masing pihak di luar PUJK']),
        rule20(data.loc[rowIndex,'Nama PUJK Pelapor'], data_plan),
        rule22(data.loc[rowIndex,'Tanggal Lapor']),
        rule23(data.loc[rowIndex,'Tanggal Lapor'], data.loc[rowIndex,'Nama PUJK Pelapor'], data_plan),
        rule24(data.loc[rowIndex,'Tanggal Lapor'])
    ]
    
template.to_excel("LPE Validation Result.xlsx", index=False) 