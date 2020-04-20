#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# screen_transform.py
# by mfpfox 1/2/20

import pandas as pd
import os
import csv
import xlrd
from openpyxl import load_workbook

"""
screen_transform.py are functions for cleaning, parsing, and transforming dataframes

TO USE: 
    1. python 3.# (any version of 3 is fine)
    2. put this script in folder with input files
    3. edit code below to have input file names and altered names '_transposed.csv' for output file names, see comments in code
    4. if you would like to run the function, uncomment (delete #) for execution line in def main() function at bottom of this file

         i.e. # transform2rows() ---> transform2rows()
              # transform2column() -> transform2column()

    5. type 'python screen_transform.py' in terminal window (make sure your terminal window path matches where your files are located)


File renaming tips: 
    a. for input files, remove space in file names, example: 
        "200227Automation12.xlsx"
    b. add altered file names to csv_single_col_from_excel(), example: 
        "200227Automation12_transposed.xlsx"

----------------------------------------------------------------

HELPER FUNCTION append_df_to_excel() notes: 
    Append a DataFrame [df] to existing Excel file [filename]
    into [sheet_name] Sheet.
    If [filename] doesn't exist, then this function will create it.
    Parameters:
      filename : File path or existing ExcelWriter
                 (Example: '/path/to/file.xlsx')
      df : dataframe to save to workbook
      sheet_name : Name of sheet which will contain DataFrame.
                   (default: 'Sheet1')
      startrow : upper left cell row to dump data frame.
                 Per default (startrow=None) calculate the last row
                 in the existing DF and write to the next row...
      truncate_sheet : truncate (remove and recreate) [sheet_name]
                       before writing DataFrame to Excel file
      to_excel_kwargs : arguments which will be passed to `DataFrame.to_excel()`
                        [can be dictionary]
    Returns: None
"""

###################### edit output file names ######################

def csv_single_col_from_excel():
    os.chdir("/Users/mariapalafox/Desktop/jose/")
    # ADD ALTERED FILE NAMES HERE: 
    excel_files = ["200227Automation12_transposed.xlsx",
                    "200302Automation13_transposed.xlsx",
                    "200304Automation14_transposed.xlsx",
                    "200304Automation15_transposed.xlsx",
                    "200304Automation16_transposed.xlsx"]
    for excel_file in excel_files:
        # SAVED new csv single col as below name
        outfile = excel_file.replace("transposed.xlsx", "_single_col.csv")
        # initialize excel file
        x1 = pd.ExcelFile(excel_file, engine='openpyxl')
        # initialize single column
        single = []
        # loop for sheets
        for sheet in x1.sheet_names:
            print('processing - ', sheet)
            # extracts data from worksheet into pd
            df = pd.read_excel(excel_file, sheet)
            # turns values into single col
            stack = df.combine_first(pd.Series(df.values.ravel('F')).to_frame('c5'))['c5']
            stackls = stack.tolist()
            single = single + stackls
        single = pd.DataFrame(single)
        print(single)
        single.to_csv(outfile, index=False)
#csv_single_col_from_excel()


# called by csv_from_excel() funx
def append_df_to_excel(filename, df, sheet_name='Sheet1', startrow=None,
                       truncate_sheet=False,
                       **to_excel_kwargs):
    # EXCEL helper function
    # ignore [engine] parameter if it was passed
    if 'engine' in to_excel_kwargs:
        to_excel_kwargs.pop('engine')
    writer = pd.ExcelWriter(filename, engine='openpyxl')
    # Python 2.x: define [FileNotFoundError] exception if it doesn't exist
    try:
        FileNotFoundError
    except NameError:
        FileNotFoundError = IOError
    try:
        # try to open an existing workbook
        writer.book = load_workbook(filename)
        # get the last row in the existing Excel sheet
        # if it was not specified explicitly
        if startrow is None and sheet_name in writer.book.sheetnames:
            startrow = writer.book[sheet_name].max_row
        # truncate sheet
        if truncate_sheet and sheet_name in writer.book.sheetnames:
            # index of [sheet_name] sheet
            idx = writer.book.sheetnames.index(sheet_name)
            # remove [sheet_name]
            writer.book.remove(writer.book.worksheets[idx])
            # create an empty sheet [sheet_name] using old index
            writer.book.create_sheet(sheet_name, idx)
        # copy existing sheets
        writer.sheets = {ws.title:ws for ws in writer.book.worksheets}
    except FileNotFoundError:
        # file does not exist yet, we will create it
        pass
    if startrow is None:
        startrow = 0
    # write out the new sheet
    df.to_excel(writer, sheet_name, startrow=startrow, **to_excel_kwargs)
    # save the workbook
    writer.save()

###################### edit input file names ##############################

def csv_from_excel():
    # ADD INPUT FILE NAMES HERE: 
    excel_files = ["200227Automation12.xlsx",
                    "200302Automation13.xlsx",
                    "200304Automation14.xlsx",
                    "200304Automation15.xlsx",
                    "200304Automation16.xlsx"]
    for i in excel_files:
        excel_file = i
        # SAVED new excel transposed as below name
        outfile = excel_file.replace(".xlsx", "_transposed.xlsx")
        x1 = pd.ExcelFile(excel_file, engine='openpyxl')
        # loop for sheets of 1 file
        for sheet in x1.sheet_names:
            print('processing - ', sheet)
            # extracts data from worksheet into pd
            df = pd.read_excel(excel_file, sheet, header=None)
            # TRANSPOSE 
            df_transposed = df.transpose()
            # write to excel sheet of new file '_transposed'
            append_df_to_excel(outfile, df_transposed, sheet_name=sheet,
                               index=False)


################# basic transform functions #######################

def transform2rows():
    filename = raw_input("Enter the name of input .csv file: ")
    outfile = raw_input("Enter the name of output .csv file: ")
    df = pd.read_csv(filename)
    print("original df shape: ", df.shape)
    boo = raw_input("Would you like to drop 1st and last columns? (True or False): ")
    if boo:
        column_numbers = [x for x in range(df.shape[1])]  # list of columns' integer indices
        column_numbers .remove(0) #removing column integer index 0
        column_numbers .remove(-1) #removing column integer last index
        df = df.iloc[:, column_numbers] #return all columns except the 0th and last
        # df.drop(df.columns[len(df.columns)-1], axis=1, inplace=True)
        df_transposed = df.transpose()
        print("shape of new df: ", df_transposed.shape)
        df_transposed.to_csv(outfile,index=False, encoding="utf8")
    else:
        df = df.transpose()
        print("shape of new df: ", df.shape)
        df.to_csv(outfile,index=False, encoding="utf8")

def transform2column():
    # transforms all row values into single column
    filename = raw_input("Enter the name of input .csv file: ")
    outfile = raw_input("Enter the name of output .csv file: ")
    with open(filename) as infile:
        csvReader = csv.reader(infile)
        jlist = []
        for row in csvReader:
            for val in row:
                jlist.append(val)
        jdf = pd.Series(jlist)
        print("Length of final column = ", len(jdf))
        # saving file as outfile name provided
        jdf.to_csv(outfile,index=False, encoding="utf8")

################### calling functions #####################################

def main():
    # UNCOMMENT BELOW LINES TO TRANSFORM SCREEN DATA (after you have edited input file names and output file names marked by comments in above code..
    csv_from_excel()
    csv_single_col_from_excel()

    # basic transform functions that ask for input #:
    # transform2column()
    # transform2rows()
main()
