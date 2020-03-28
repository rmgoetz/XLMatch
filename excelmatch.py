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


#
# The script inputs
#-----------------------------------------------------------------------------
file1 = input('Enter first file name:')
file2 = input('Enter second file name:')
#ignore_sheet_numbers_Q = input('Ignore sheet number mismatch (Y/N):')
ignore_sheet_names_Q = input('Ignore sheet name mismatch (Y/N):')
#if ignore_sheet_numbers_Q == ('Y' or 'y'):
#    ignore_sheet_numbers = True
#else:
#    ignore_sheet_numbers = False
if ignore_sheet_names_Q == ('Y' or 'y'):
    ignore_sheet_names = True
else:
    ignore_sheet_names = False
#-----------------------------------------------------------------------------
#


#
# Try to read the input files
#-----------------------------------------------------------------------------
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
#-----------------------------------------------------------------------------
#
    

#
# Sort filenames, remove extensions where needed, open report file
#-----------------------------------------------------------------------------
f1_dot = file1[::-1].find('.')+1
f2_dot = file2[::-1].find('.')+1
name_sort = sorted((file1,file2))    
errfile_name = 'match_report_'+name_sort[0][:-f1_dot]+'_'+name_sort[1][:-f2_dot]+'.txt'
errfile = open(errfile_name,'w')
err_out = 'File 1: '+name_sort[0]+'\n'+'File 2: '+name_sort[1]+'\n'
#-----------------------------------------------------------------------------
#


#
# The test-passing variables
#-----------------------------------------------------------------------------
sheet_test = True
shape_test = True
cell_test = True
#-----------------------------------------------------------------------------
#


#
# Define goodbye, index to cell converter, and shape fixing functions
#-----------------------------------------------------------------------------
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


def shapefix(daf1,daf2):
    shp1 = daf1.shape
    shp2 = daf2.shape
    if shp1 == shp2:
        return daf1, daf2
    else:
        maxrow = max(shp1[0],shp2[0])
        maxcol = max(shp1[1],shp2[1])
        row_index = list(range(maxrow))
        col_index = list(range(maxcol))
        newdaf1 = daf1.reindex(index=row_index,columns=col_index)
        newdaf2 = daf2.reindex(index=row_index,columns=col_index)
        return newdaf1, newdaf2
#-----------------------------------------------------------------------------
#


#
# Check number of sheets
#-----------------------------------------------------------------------------
length1 = len(ex1.sheet_names)
length2 = len(ex2.sheet_names)
if length1 != length2:
    err_out += 'Sheet number mismatch\n'
    err_out += 'No further match testing is possible until this issue is resolved\n'
    goodbye()
else:
    err_out += 'Number of sheets match\n'
    
# --- when I want to try ignore sheet numbers:    
#    #if not ignore_sheet_numbers:
#    #    err_out += 'No further match testing is possible until this issue is resolved\n'
#    #    goodbye()
#    #else:
#    #    min_sheet = min([length1,length2])
#    #    min_index = [length1,length2].index(min_sheet) # this might be an issue for certain python versions
#    #    
#    #    err_out += 'File 1 will serve as sheet number reference'

#-----------------------------------------------------------------------------
#
    

#
# Check names of sheets
#-----------------------------------------------------------------------------
for a,b in zip(ex1.sheet_names,ex2.sheet_names):
    sheet_test &= (a==b)
if not sheet_test:
    err_out += 'Sheet names or order are mismatched\n'
    if not ignore_sheet_names:
        err_out += 'No further match testing is possible until this issue is resolved\n'
        goodbye()
    else:
        err_out += 'File 1 will serve as sheet name reference\n'
else:
    err_out += 'Sheet names and orders are matched\n'
#-----------------------------------------------------------------------------
#


#
# Check sheet shapes and then go cell-by-cell
#-----------------------------------------------------------------------------
sheet_names = ex1.sheet_names
for ind,sht in enumerate(sheet_names):
    df1 = pd.read_excel(file1,sheet_name=ind,header=None)
    df2 = pd.read_excel(file2,sheet_name=ind,header=None)
    EQ1 = np.shape(df1) == np.shape(df2)
    shape_test &= EQ1
    if EQ1:
        err_out += 'Shape of sheet named '+sht+' matches\n'
        
    else:
        err_out += 'Shape of sheet named '+sht+' does not match\n'
        df1, df2 = shapefix(df1,df2)
    df1 = df1.astype('str')
    df2 = df2.astype('str')
    EQ2 = df1.equals(df2)
    cell_test &= EQ2
    if not EQ2:
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
#-----------------------------------------------------------------------------
#
                  
   
#
# Goodbye if we haven't already
#-----------------------------------------------------------------------------    
goodbye()
#-----------------------------------------------------------------------------
#