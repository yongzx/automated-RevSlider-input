import re
import openpyxl
import os
import sys

def readFile():
    """
    Prompt for the name of the excel file.
    Read in and open the excel file.

    Arguments:
    None

    Returns:
    name of the excel file --- str
    workbook ---openpyxl
    """
    os.chdir("Excel Files")
    while True:
        try:
            excel_name = input(
                "Name of your Excel document (remember to put .xlsx in your input, e.g., test.xlsx): ")
            if re.search(".xlsx", excel_name) is None:
                raise FileNameError
            excel_file = openpyxl.load_workbook(excel_name)
            os.chdir("..")
            return excel_name, excel_file
        except FileNameError:
            print("You forgot to put .xlsx in your file name. \n")
        except FileNotFoundError:
            print("File '{0}' not found. Remember to put your file in the same folder as this script.py. \n".format(
                excel_name))


def readSheet(excel_file):
    """
    Prompt for the name of sheet
    Read in and open a particular sheet in the excel file

    Arguments: 
    excel_file ---openpyxl workbook

    Return:
    name of the sheet --- str
    sheet --- openpyxl

    """
    while True:
        try:
            sheet_name = input("Sheet name (case sensitive): ")
            sheet = excel_file.get_sheet_by_name(sheet_name)
            return sheet_name, sheet
        except KeyError:
            print("Sheet '{0}' does not exist. \n".format(sheet_name))

def readSlider():
    slider = input("Slider (i.e., i, j, k, ...): ")
    response = input("Are you sure '{0}' is the slider you want? Type 'y' if yes, 'n' if no.\n".format(slider))
    while response != 'y':
        slider = input("Slider (i.e., i, j, k, ...): ")
        response = input("Are you sure '{0}' is the slider you want? Type 'y' if yes, 'n' if no.\n".format(slider))
    return slider

class FileNameError(Exception):
    pass