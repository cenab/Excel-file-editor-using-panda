import numpy as np
import pandas as pd
import os

def getFilesDirectory():
    filenames = []
    files_path = [os.path.abspath(x) for x in os.listdir()]
    for i in range(len(files_path)):
        filenames.append(files_path[i].split("\\")[-1])
    filenames.remove("line_deleter.py")
    return filenames

def readExcel(filename):
    df = pd.read_excel(filename, index_col=None, na_values=['NA'])
    return df

def rowDeleteList(df):
    deletedRows = []
    count_row = df.shape[0] #total number of row
    for row in range(count_row-1): #goes through each row to add the indexes to a set
        if df.at[row, "Tarih"] == df.at[row + 1, "Tarih"]: #compares the two values in the list
            deletedRows.append(row)
    return deletedRows

def deleteRow(deletedRows, df):
    df = df.drop(deletedRows, axis=0)
    return df

def returnData2Excel(df, filename):
    df.to_excel(filename, index=None, header = True)
    print("Successfully edited!")
    return

def main():
    filenames = getFilesDirectory()
    try:
        for file in filenames:
            df = readExcel(file)
            deletedRows = rowDeleteList(df)
            newdf = deleteRow(deletedRows, df)
            returnData2Excel(newdf, file)
    except:
        print("This file cannot be edited: ", file)
        
if __name__ == "__main__":
    main()

