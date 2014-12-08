from openpyxl import load_workbook
import scipy.optimize
import numpy as np
import csv
import matplotlib.pyplot as plt
import pandas as pd
import itertools
#Read me: describe the problem and anything you need to run the code
#Conclusion of nested loop


def extract_data(data,start_row,end_row,start_col,end_col):
    data_range_start = []
    data_range_end = []
    usable_data = []
    
    def excel_col(col):
        """col is a column number returns letter. 1=A """
        quot, rem = divmod(col-1,26)
        return excel_col(quot) + chr(rem+ord('A')) if col!=0 else ''    

    for column in range(start_col,end_col):
        data_range_start.append(str(excel_col(column))+start_row)
        data_range_end.append(str(excel_col(column))+end_row)
    
    for i, start_pos in enumerate(data_range_start):
        end_pos = data_range_end[i]
        cell_range = data[start_pos:end_pos]
        col_results = [n[0].value for n in cell_range]
        usable_data.append(col_results)
    return(usable_data)


def alternate(i):
    """Pairs columns into corresponding replicates"""
    i = iter(i)
    while True:
        yield(i.next(), i.next())   

wb_template = load_workbook('Data_for_Sarah.xlsx', data_only=True)
data = wb_template[ "Sheet1" ] 
row_start_stop = [(str(17+28*nums),str(32+28*nums))  for nums in range(0,7)]
#columns = range(1,len(working_data))
#col_pairs = list(alternate(columns)) 
Analytes = []
for starts,ends in row_start_stop:
    usable_data = extract_data(data,starts,ends,22,35)
    Analytes.append(usable_data)
Analyte1 = Analytes[0]
rep1 = []
rep2 = []
results = []

for columns in range(0,len(Analyte1)):
    for i in alternate(range(0,17)):
        rep1.append(Analyte1[columns][i[0]])
        rep2.append(Analyte1[columns][i[1]])
    results.append((rep1,rep2)) 
    rep1 = []
    rep2 = []
with open('test.csv', 'wb') as f:
    writer = csv.writer(f)
    for item  in results:
        for lists in item:
            for vals in lists:
                print vals
                writer.writerow([vals])
