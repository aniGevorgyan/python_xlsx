#!/usr/bin/python

import json
import xlsxwriter
import argparse
from datetime import datetime
from variables import *


def readData(filename):
  with open(filename) as f:
    data = json.load(f)
    # Get needed columns from data
    scores = [[i["firstName"], i["lastName"], i["profession"], i["dob"], i["email"], i["address"]] for i in data]
  return scores


def writeData(data, workbook, worksheet):
  writeTitles(workbook, worksheet)
  row = 1
  col = 0

  # Write data
  for firstName, lastName, prof, dob, email, address in data:
    # Styles of columns
    if (prof == "Software Developer" and getYearFromDate(dob) > 1985):
      style = workbook.add_format(style2)
    else:
      style = workbook.add_format(style1)
    worksheet.write(row, col, firstName, style)
    worksheet.write(row, col + 1, lastName, style)
    worksheet.write(row, col + 2, prof, style)
    worksheet.write(row, col + 3, dob, style)
    worksheet.write(row, col + 4, email, style)
    worksheet.write(row, col + 5, address, style)
    row += 1
  return


def writeTitles(workbook, worksheet):
  row = 0
  col = 0
  title = ["Name", "Surname", "Profession", "DOB", "Email", "Address"]
  cell_format = workbook.add_format(style3)
  worksheet.set_column('A:B', 20)
  worksheet.set_column('C:E', 30)
  worksheet.set_column('F:F', 50)
  for name in title:
    worksheet.write(row, col, name, cell_format)
    col += 1
  return


def getYearFromDate(date):
  dt = datetime.strptime(date, '%Y-%m-%d')
  return int(dt.year)


def getArgs():
  parser = argparse.ArgumentParser(description='Program write data from .txt(.json) file to .xlsx')
  parser.add_argument('-f', '--file', help='File which data should come from (.txt, .json)', required=True)
  parser.add_argument('-x', '--xlsx', help='File in which data should be written (.xlsx)', required=True)
  args = parser.parse_args()

  if (args.xlsx[-5:] == ".xlsx" and (args.file[-4:] == ".txt" or args.file[-5:] == ".json")):
    return args
  return False


def main():
  args = getArgs()
  if (args):
    workbook = xlsxwriter.Workbook(args.xlsx)
    worksheet = workbook.add_worksheet()
    data = readData(args.file)
    writeData(data, workbook, worksheet)
    workbook.close()
    print(str(args.xlsx) + " was created")
  else:
    print("Wrong file extension")


if __name__ == "__main__":
  main()
