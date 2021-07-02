#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# screen_transform.py
# by mfpfox 7/1/21

# screen_transform.py contains functions for transforming screening .xlsx datasets to single column .csv files
# 7/1/21 Updated to pass list of input file names and output file names as arguments

"""
TO RUN CODE, FOLLOW THESE STEPS...
1. DONT PANIC AND YOU DONT NEED TO MOVE INPUT FILES TO THE FOLDER WITH THIS SCRIPT

2. MAKE SURE YOU ARE IN donOmar ENVIRONMENT WITH ALL YOUR PYTHON PACKAGES

    $conda activate donOmar

3. IN THIS SCRIPT EDIT THE 2 MARKED SECTIONS IN THE main()

4. AFTER YOU HAVE ENTERED INPUT FILE NAMES AND PATH AT END OF THIS SCRIPT, IN TERMINAL TYPE...

    $ python screen_transform.py 

5. HAVE A BEER BECAUSE YOU DID IT!!!
"""

import pandas as pd
import os
import sys
import csv
import xlrd
from openpyxl import load_workbook

#sys.path.append("/Users/jose-appleair/Desktop/GitProjects/screen_data_cleaning/")

def csv_single_col_from_excel(transposed_files):
    for excel_file in transposed_files:
        # saves new csv single col as below name
        outfile = excel_file.replace("transposed.xlsx", "_single_col.csv")
        # initialize excel file
        x1 = pd.ExcelFile(excel_file, engine='openpyxl')
        # initialize single column outfile
        single = []
        # loop thru sheets
        for sheet in x1.sheet_names:
            print("making ", excel_file, " sheet name- ", sheet, " into single col csv")
            # extracts data from worksheet into pandas df
            df = pd.read_excel(excel_file, sheet)
            # turns all values into single col
            stack = df.combine_first(pd.Series(df.values.ravel('F')).to_frame('c5'))['c5']
            stackls = stack.tolist()
            single = single + stackls
        single = pd.DataFrame(single)
        single.to_csv(outfile, index=False)
        print("saving ", outfile)



def append_df_to_excel(filename, df, sheet_name='Sheet1', startrow=None,
                       truncate_sheet=False,
                       **to_excel_kwargs):
    """
    Append a DataFrame [df] to existing Excel file [filename]
    into [sheet_name] Sheet. If [filename] doesn't exist, then this function will create it.
    Parameters:
      filename : File path or existing ExcelWriter (Example: '/path/to/file.xlsx')
      df : dataframe to save to workbook
      sheet_name : Name of sheet which will contain DataFrame (default: 'Sheet1')
      startrow : upper left cell row to dump data frame.
                 Per default (startrow=None) calculate the last row
                 in the existing DF and write to the next row...
      truncate_sheet : truncate (remove and recreate) [sheet_name]
                       before writing DataFrame to Excel file
      to_excel_kwargs : arguments which will be passed to `DataFrame.to_excel()`
                        [can be dictionary]
    Returns: None
    """
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

        # get the last row in the existing Excel sheet, if it was not specified explicitly
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
    # file does not exist yet, create it
    except FileNotFoundError:
        pass

    if startrow is None:
        startrow = 0

    # write new sheet
    df.to_excel(writer, sheet_name, startrow=startrow, index=False)
    # save the workbook
    writer.save()


def csv_from_excel(excel_files, transposed_files):
    for inx,file in enumerate(excel_files):
        excel_file = file
        outfile = transposed_files[inx]
        print("making ", outfile)
        x1 = pd.ExcelFile(excel_file, engine='openpyxl')
        for sheet in x1.sheet_names:
            print('transposing ', file, ' sheet name- ', sheet)
            # extracts data from worksheet into pandas dataframe
            df = pd.read_excel(excel_file, sheet, header=None)
            df = df.transpose()
            print("preview transposed data\n", df.head())
            # call helper function to write to excel sheet of transposed file
            append_df_to_excel(outfile, df, sheet_name=sheet)


################### EDIT MARKED SECTIONS BELOW #####################################

def main():
    
    #### EDIT 1/2 - ADD FULL PATH BELOW, MAKE SURE '/' IS AT END OF PATH STRING, ex. "/Users/jose-appleair/Desktop/" ###
    path2data = "/Users/jose-appleair/Desktop/finalAutomationData/"
    ####
    
    # os.chdir(path2data) # all files will be saved to location you set as path2data variable
    
    #### EDIT 2/2 - ADD INPUT FILE NAMES TO excel_files LIST BELOW ###
    # remove space in file names, ex. "200227Automation12.xlsx"
    excel_files = ["210603Automation21.xlsx"]
    ####

    transposed_files = []

    for i in excel_files: 
        excel_file = i
        outfile = excel_file.replace(".xlsx", "_transposed.xlsx")
        transposed_files.append(outfile)

    print("Your input files are: ", excel_files)
    print("Your transposed files will be named: ", transposed_files)
    print("Your final single column csv files will have same name as input files but with _single_col.csv")
    print("\nStart processing input files...\n")

    # function calls that process your data
    csv_from_excel(excel_files, transposed_files)
    csv_single_col_from_excel(transposed_files)

main()
