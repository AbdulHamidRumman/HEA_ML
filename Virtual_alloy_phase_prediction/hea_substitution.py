import pandas as pd
import openpyxl
import numpy as np


data = pd.read_excel("F:/Research/High Entropy Alloys/HEA dataset-1200 instances-NEW.xlsx")
wb = openpyxl.load_workbook("F:/Research/High Entropy Alloys/Data.xlsx")
ws = wb.active
count=1
num_columns = data.shape[1]
matrix = np.zeros((num_columns, num_columns))
for i in range(num_columns):
    matrix[i][i]=1


for index,column_name in enumerate(data.columns):
    NA = 0
    for index7, row in data.iterrows():
        if row.iloc[index] != 0:
            NA+=1

    for index2, column_name2 in enumerate(data.columns):
        if index2 != index:
            Nu = 0
            for index3, row2 in data.iterrows():
                if (row2.iloc[index] != 0 and row2.iloc[index2] == 0 ):
                    row_to_save = row2.copy()
                    temp=row_to_save[column_name]
                    row_to_save[column_name] = row_to_save[column_name2]
                    row_to_save[column_name2] = temp
                    for index4, row3 in data.iterrows():
                        x=row3.copy()
                        if x.equals(row_to_save):
                            Nu+=1
                            break
            print(Nu)        
            if NA!=0:
                matrix[index][index2]=Nu/NA



for index,column_name in enumerate(data.columns):

    for index2, column_name2 in enumerate(data.columns):

        if index2 != index:

            for index3, row2 in data.iterrows():
                if (row2.iloc[index] != 0 and row2.iloc[index2] == 0 ):
                    row_to_save = row2.copy()
                    temp=row_to_save[column_name]
                    row_to_save[column_name] = row_to_save[column_name2]
                    row_to_save[column_name2] = temp
                    flag=1
                    for index4, row3 in data.iterrows():
                        x=row3.copy()
                        if (x.ne(row_to_save).any()) is False:
                            flag=0
                            break

                    if flag == 1:
                        new_series = pd.Series({
                            **row_to_save,  # Unpack the existing Series
                            'Probability': matrix[index][index2]  # Add a new key-value pair
                            })
                        for col_idx, value in enumerate(new_series):
                            ws.cell(row=count , column=col_idx + 1).value = value
                        count+=1
                            

wb.save("F:/Research/High Entropy Alloys/Data.xlsx")

                            
