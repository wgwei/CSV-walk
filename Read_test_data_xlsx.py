# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import os
from openpyxl import load_workbook
from tkinter import filedialog
path = filedialog.askdirectory()
import numpy as np

#%%
xlsx_files = []  # All xlsx files
xlsx_files_full = []  # All xlsx files and path
for root, dirs, files in os.walk(path):
    for name in files:
        if name.endswith(".xlsx"):
            xlsx_files.append(name)
            xlsx_files_full.append(str(root)+'\\'+str(name))

xlsx_sound_full = []  # All xlsx file which contain the word 'sound'
xlsx_sound = []
for i in range(len(xlsx_files_full)):
    T_F = 'Sound' in xlsx_files_full[i]
    if T_F == True:
        xlsx_sound_full.append(xlsx_files_full[i])
        xlsx_sound.append(xlsx_files[i])

freq = np.array([100,125,160,200,250,315,400,500,630,800,1000,1250,1600,2000,2500,3150])

#%%
files2process_full = []
files2process = []

def test_the_file(file_full, file):
    """Tests whether the workbook has a 'data' worksheet.
    If it does, it adds the file name to the list files2process"""
    wb = load_workbook(file_full, read_only=True)
    try:
        wb['Data']
        files2process_full.append(file_full)
        files2process.append(file)
    except KeyError:
        pass
    return files2process_full, files2process

for i in range(len(xlsx_files_full)):
    test_the_file(xlsx_files_full[i], xlsx_files[i])

#%%
def extract_DnT(file_full, file):
    """Extracts the DnT values from the xlsx workbooks which have 'data' worksheet"""
    wb = load_workbook(file_full, read_only=True)
    Visit = '5A'
    if file[1] == '-':
        Test = str(file[0])
    else:
        Test = str(str(file[0])+str(file[1]))
    ws = wb['Data']
    DnT = np.array([ws['D5'].value, ws['D6'].value, ws['D7'].value,
                    ws['D8'].value, ws['D9'].value, ws['D10'].value, 
                    ws['D11'].value, ws['D12'].value, ws['D13'].value,
                    ws['D14'].value, ws['D15'].value, ws['D16'].value,
                    ws['D17'].value, ws['D18'].value, ws['D19'].value,
                    ws['D20'].value])
    Dntw = int(ws['B39'].value)
    Ctr = int(ws['B41'].value)
    DntwCtr = Dntw + Ctr
    d = {'Visit':Visit, 'Test':Test, 'DnT': DnT, 'Dntw + Ctr': DntwCtr, 'Dntw': Dntw}
    return d

with open(path+'\\All_tests.csv', 'w') as ff:
    ff.write('{}, '.format('Frequency'))
    for i in range(len(files2process_full)):
        d1 = extract_DnT(files2process_full[i], files2process[i])
        ff.write('{} - Test {}, '.format(d1['Visit'], d1['Test']))
    ff.write('\n')
    
    for f in range(len(freq)):
        ff.write('{}, '.format(freq[f]))
        for i in range(len(files2process_full)):
            d1 = extract_DnT(files2process_full[i], files2process[i])
            ff.write('{}, '.format(d1['DnT'][f]))
        ff.write('\n')

    ff.write('\n')
    ff.write('DnT,w, ')
    for i in range(len(files2process)):
        d1 = extract_DnT(files2process_full[i], files2process[i])
        ff.write('{}, '.format(d1['Dntw']))
    ff.write('\n')

    ff.write('DnT,w+Ctr, ')
    for i in range(len(files2process)):
        d1 = extract_DnT(files2process_full[i], files2process[i])
        ff.write('{}, '.format(d1['Dntw + Ctr']))
    ff.write('\n')