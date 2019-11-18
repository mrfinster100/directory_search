## directory_search.py
## Searches files in directory and subdirectories for a specified string.

## Setup:
# 1) Run file.reg to enable drag-and-drop conversion
# 2) Install necessary Python 3.8 dependencies: (i.e. PyPDF2, docx, xlrd, python-pptx)
# 3) Drag and drop any parent directory or file on top of directory_search.py
# 4) Specify a searchable string (not case-sensitive)
# 5) View list of files containing the specified string w/ instances

##TODO:
# if pdf requires decryption, skip or ask for password

from docx import Document
from pptx import Presentation
import sys
import os
import re
import PyPDF2
import csv
import xlrd
import xml.etree.ElementTree as ET
## 1
# 3rd Level Helper Function
# for searching ppt/pptx
def search_ppt(pptfname):
      array = []
      prs = Presentation(pptfname)
      for slide in prs.slides:
            for shape in slide.shapes:
                  if not shape.has_text_frame:
                      continue
                  for paragraph in shape.text_frame.paragraphs:
                      #for run in paragraph.runs:
                        paratext = paragraph.text
                        words = paratext.split()
                        array.extend(words)
      text = (" ").join(array)
      matches = re.findall(specifiedString, text, flags=re.IGNORECASE)
      numMatches = len(matches)
      return numMatches
## 2
# 3rd Level Helper Function
# for searching doc/docx
def search_doc(docfname):
    array = []
    document = Document(docfname)
    NumParagraphs = len(document.paragraphs)
    text = ""
    for i in range(0, NumParagraphs):
        array.append(document.paragraphs[i].text)
    text = ("\n").join(array)
    matches = re.findall(specifiedString, text, flags=re.IGNORECASE)
    numMatches = len(matches)
    return numMatches
## 3
# 3rd Level Helper Function
# for searching .pdf files
def search_pdf(pdffname):
    text = ""
    pdfFileObj = open(pdffname, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    NumPages = pdfReader.getNumPages()
    for i in range(0, NumPages):
        PageObj = pdfReader.getPage(i)
        text += PageObj.extractText()
    text = text.split()
    text = (" ").join(text)
    matches = re.findall(specifiedString, text, flags=re.IGNORECASE)
    numMatches = len(matches)
    return numMatches
## 4  
# 3rd Level Helper Function    
# for searching .csv files
def search_csv(csvfname):
    text = ""
    with open(csvfname, 'rt') as f:
        reader = csv.reader(f, delimter=',')
        for row in reader:
            for field in row:
                text += field
            matches = re.findall(specifiedString, text, flags=re.IGNORECASE)
            numMatches = len(matches)
            return numMatches
## 5         
# 3rd Level Helper Function
# for searching .xlsx
def search_xlsx(xlsxfname):
    xl_workbook = xlrd.open_workbook(xlsxfname, on_demand=True)
    res = len(xl_workbook.sheet_names()) 
    text = ""
    for i in range(0, res):
        xl_sheet = xl_workbook.sheet_by_index(i)
        num_cols = xl_sheet.ncols
        for row_idx in range(0, xl_sheet.nrows):
            for col_idx in range(0, num_cols): 
                cell_obj = xl_sheet.cell(row_idx, col_idx)
                text += cell_obj.value
    matches = re.findall(specifiedString, text, flags=re.IGNORECASE)
    numMatches = len(matches)
    return numMatches
## 1
# 2nd Level Switch Function   
# receives a file directory
# returns # of matches
def fileSwitch(fileDir): 
    pre, ext = os.path.splitext(fileDir)
    # search pptx/ppt
    if ext == '.pptx' or ext == '.ppt':
        return search_ppt(fileDir)
    # search doc/docx
    if ext == ".doc" or ext == '.docx':
        return search_doc(fileDir)
    # search pdf
    if ext == ".pdf":
        return search_pdf(fileDir)
    # search csv
    if ext == ".csv":
        return search_csv(fileDir)
    ## search xlsx
    if ext == ".xlsx":
        return search_xlsx(fileDir)
## 1
##MAIN FUNCTION##
fileHash = {}
specifiedDirectory = sys.argv[1]
specifiedString = input("What string would you like to search for?\n")
for subdir, dirs, files in os.walk(specifiedDirectory):
    for file in files:
        file = os.path.join(subdir, file)
        numMatches = fileSwitch(file)
        relPatch = os.path.relpath(file, subdir)
        fileHash[relPatch] = numMatches
for key, value in fileHash.items():
    print(key)
    print("Contains " + str(value) + " instance(s) of specified string.")
    print()
input("Press any key to exit...")
