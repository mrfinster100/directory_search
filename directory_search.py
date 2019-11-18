## directory_search.py
## Searches files in directory and subdirectories for a specified string.

##TODO:
# create tempdir and convert to .zip to avoid damaging .ppt
# if pdf requires decryption, skip
# improve pptx parsing

## Setup:
# 1) Run file.reg to enable drag-and-drop conversion
# 2) Install necessary Python 3.8 dependencies: (i.e. zipfile, PyPDF2, docx, xlrd)
# 3) Drag and drop any parent directory or file on top of directory_search.py
# 4) Specify a searchable string (not case-sensitive)
# 5) View list of files containing the specified string w/ instances

from zipfile import ZipFile
import zipfile
import sys
import tempfile
import shutil
import os
import re
import fnmatch
from docx import Document
import PyPDF2
import csv
import xlrd
import xml.etree.ElementTree as ET
## 1
# 3rd Level Helper Function
# for searching ppt/pptx
def search_ppt(pptfname):
    pre, ext = os.path.splitext(pptfname)
    newDir = pre + ".zip"
    os.rename(pptfname, newDir)
    tempdir = tempfile.mkdtemp()
    try:
        array = []
        # creates directory string
        tempname = os.path.join(tempdir, 'new.zip')
        # reads original zip
        with zipfile.ZipFile(newDir, 'r') as zipread:
            #writes to temp zip
            with zipfile.ZipFile(tempname, 'w') as zipwrite:
                #iterates over each file in original zip
                for item in zipread.infolist():
                    # contents to be checked
                    filename = item.filename
                    # slide.xml contains text
                    if fnmatch.fnmatch(filename, '*slide*.xml'):
                        data = zipread.read(item.filename)
                        decoded = data.decode('utf-8')
                        textblocks = re.findall("<a:t>(.*?)</a:t>", decoded)
                        ##
                        ##TODO: figure out why it's including weird text
                        ##
                        array.extend(textblocks)
                text = (" ").join(array)
                matches = re.findall(specifiedString, text, flags=re.IGNORECASE)
                numMatches = len(matches)
                return numMatches
    finally:
        os.rename(newDir, pptfname)
        shutil.rmtree(tempdir)
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
specifiedString = input("What string would you like to search for?")
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
