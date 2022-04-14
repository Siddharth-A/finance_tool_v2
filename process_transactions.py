#!/usr/bin/env python3

# external libraries
import sys
import os
import csv
import json
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, numbers

# internal libraries

# colors
orange = PatternFill(start_color='ffcaa1',end_color='ffcaa1',fill_type='solid')
purple = PatternFill(start_color='dbb5ff',end_color='dbb5ff',fill_type='solid')
red1   = PatternFill(start_color='ffa1a1',end_color='ffa1a1',fill_type='solid')
red2   = PatternFill(start_color='e06565',end_color='e06565',fill_type='solid')
red3   = PatternFill(start_color='ff2e2e',end_color='ff2e2e',fill_type='solid')
yellow = PatternFill(start_color='fffa99',end_color='fffa99',fill_type='solid')
green1 = PatternFill(start_color='c1ffab',end_color='c1ffab',fill_type='solid')
green2 = PatternFill(start_color='00ff80',end_color='00ff80',fill_type='solid')
green3 = PatternFill(start_color='34fa4f',end_color='34fa4f',fill_type='solid')
blue   = PatternFill(start_color='abf2ff',end_color='abf2ff',fill_type='solid')
white  = PatternFill(start_color='ffffff',end_color='ffffff',fill_type='solid')
black  = PatternFill(start_color='000000',end_color='000000',fill_type='solid')

# global variables
tran_mon                = ""
bmo_csv                 = ""
bmo_sheet_name          = ""
output_file             = ""
output_file_sheet_title = ""

# store data input by user into global variables
def process_user_input():
    global tran_mon, bmo_csv, output_file, output_file_sheet_title
    tran_mon = input("enter month of transactions: ")
    bmo_csv = input("enter BMO transactions file name: ")
    output_file = tran_mon + "-transactions.xlsx"
    output_file_sheet_title = tran_mon

# process bmo csv file
def process_bmo_csv(bmo_csv):
    print("\n1) process BMO transactions file: {}".format(bmo_csv))
    wb = Workbook()
    ws = wb.active
    global bmo_sheet_name
    bmo_sheet_name = tran_mon + "-bmo mc"

    # copy from csv to output_file
    with open(bmo_csv,'r') as f:
        for row in csv.reader(f):
            ws.append(row)

    # delete row 1
    ws.delete_rows(1,1)

    # fix transaction date and print to column 1 + print transaction type to column 4
    i = 1
    while i <= ws.max_row:
        ws.cell(row=i, column=1).value = '=DATE(LEFT(C{},4), MID(C{},5,2), RIGHT(C{},2))'.format(i,i,i)
        ws.cell(row=i, column=4).value = 'BMO MC'
        i +=1

    # make transaction date value
    ws.title = bmo_sheet_name
    wb.save(output_file)
    reply = input("open {}, make col A values only and then press enter ".format(output_file)) 
    wb = load_workbook(output_file)
    ws = wb[bmo_sheet_name]

    # set format for transaction date
    i = 1
    while i <= ws.max_row:
        ws.cell(row=i, column=1).number_format = numbers.FORMAT_DATE_DDMMYY
        i +=1

    # move transaction amount to column 2 + convert text to number + set format
    ws.move_range("E1:E{}".format(ws.max_row), rows=0, cols=-3)
    i = 1
    while i <= ws.max_row:
        ws.cell(row=i, column=2).value = float(ws.cell(row=i, column=2).value)
        ws.cell(row=i, column=2).number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
        i +=1

    # move transaction description to column 3
    ws.move_range("F1:F{}".format(ws.max_row), rows=0, cols=-3)

    # save workbook
    wb.save(output_file)
    print("done")


def main():
    process_user_input()
    process_bmo_csv(bmo_csv)

if __name__== "__main__":
  main()