# -*- coding: utf-8 -*-
"""
Created on Sat Mar 28 01:00:14 2020

@author: Ryan Goetz
"""
import sys
import string
import numpy as np
import pandas as pd
from pandas import ExcelFile


file1 = input('Enter first file name:')
file2 = input('Enter second file name:')


try:
    ex1 = ExcelFile(file1)
except:
    print('First file not found')
    sys.exit( )
try:
    ex2 = ExcelFile(file2)
except:
    print('Second file not found')
    sys.exit( )
    

f1_dot = file1[::-1].find('.')+1
f2_dot = file2[::-1].find('.')+1
name_sort = sorted((file1,file2))    
errfile_name = 'match_report_'+name_sort[0][:-f1_dot]+'_'+name_sort[1][:-f2_dot]+'.txt'
errfile = open(errfile_name,'w')
err_out = 'File 1: '+name_sort[0]+'\n'+'File 2: '+name_sort[1]+'\n'


sheet_test = True
shape_test = True
cell_test = True


def goodbye():
    if sheet_test*shape_test*cell_test:
        print('No mismatches found')
    else:
        print('Mismatches detected')
    errfile.write(err_out)
    errfile.write('Goodbye')
    errfile.close()
    print('Report written to '+errfile_name)
    sys.exit( )
    
    
def letter(idx):
    IDX = int(idx)
    alph = string.ascii_uppercase
    rvlet = alph[IDX % 26]
    while IDX//26 > 0 :
        IDX = IDX//26
        rvlet += alph[IDX % 26 - 1]
    return rvlet[::-1]

  
if len(ex1.sheet_names)!=len(ex2.sheet_names):
    err_out += 'Sheet number mismatch\n'
    goodbye()
else:
    err_out += 'Number of sheets match\n'
    

for a,b in zip(ex1.sheet_names,ex2.sheet_names):
    sheet_test &= (a==b)
if not sheet_test:
    err_out += 'Sheet names or order are mismatched\n'
    err_out += 'No further match testing is possible until this issue is resolved\n'
    goodbye()
else:
    err_out += 'Sheet names and orders are matched\n'


sheet_names = ex1.sheet_names
for sht in sheet_names:
    df1 = pd.read_excel(file1,sheet_name=sht,header=None)
    df2 = pd.read_excel(file2,sheet_name=sht,header=None)
    EQ = np.shape(df1) == np.shape(df2)
    shape_test &= EQ
    if not EQ:
        err_out += 'Shape of sheet named '+sht+' does not match\n'
    else:
        err_out += 'Shape of sheet named '+sht+' matches\n'

        
if not shape_test:
    err_out += 'Files not compared cell-by-cell due to sheet shape mismatch\n'        
    goodbye()
    

for sht in sheet_names:
    df1 = pd.read_excel(file1,sheet_name=sht,header=None)
    df2 = pd.read_excel(file2,sheet_name=sht,header=None)
    df1 = df1.astype('str')
    df2 = df2.astype('str')
    EQ = df1.equals(df2)
    cell_test &= EQ
    if not EQ:
        err_out += 'In sheet named '+sht+' the following indices do not match:\n'
        (nrow,ncol) = np.shape(df1)
        indices = []
        for i in range(nrow):
            for j in range(ncol):
                indices.append(f'{letter(j)}{i+1}\n')
        indices = np.asarray(indices).reshape(-1,ncol)
        df1.columns = df2.columns
        not_sames = df1.where(df1==df2,indices)
        sames = not_sames.where(not_sames!=df1)
        for i in range(nrow):
            err_out += sames.loc[i].str.cat()
    else:
        err_out += 'All cells in sheet named '+sht+' match\n'
                
    
goodbye()