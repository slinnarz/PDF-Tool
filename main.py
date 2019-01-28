# -*- coding: utf-8 -*-

'''
TO DOs:
    - Finish GUI layout.
    - Convert MS Office docs to PDF.
    - Clean (and shorten) the code.
IN PROGRESS:    
    - Building convert_xls and convert_doc functions
'''

############################
##### IMPORT PACKAGES ######
############################

import os
from appJar import gui
from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger
from win32com import client


#############################
##### DEFINE FUNCTIONS ######
#############################

### GENERAL FUNCTIONS

def list_files(dataType):
    '''
    Lists all files of specified type in the current folder.
    Takes data type as string (e.g. '.pdf') as argument.
    '''
    files = os.listdir('.\\Test files')
    return [file for file in files if file.endswith(dataType)]

def add_pageBookmarks(pdfFiles, inFolder):
    '''
    Adds bookmarks for page numbers, if the document has >1 pages.
    Takes a list of PDF-files as an argument as well as a filepath to the
    input folder (str).
    '''
    if os.path.exists(exportFolder) == False:
        os.makedirs(exportFolder)
    for file in pdfFiles:
        # open PdfFileWriter
        pdfWriter = PdfFileWriter()
        # open PDF-file
        inputFile = open(inFolder + "\\" + file, 'rb')
        myPDF = PdfFileReader(inputFile, strict=False)
        # if file has >1 pages, add page bookmarks
        if myPDF.numPages > 1:
            for page in range(0, myPDF.numPages):
                pdfWriter.addPage(myPDF.getPage(page))
                pdfWriter.addBookmark('{}'.format(page+1), page, parent=None)
            # write PDF-file
            outputFile = exportFolder + "\\" + file#[:len(file)-4] + '_neu.pdf'
            with open(outputFile, 'wb') as resultPDF:
                pdfWriter.write(resultPDF)
                
def del_allFiles(folder):
    '''
    Deletes all files in selected subfolder.
    Takes foldername (string) as an argument.
    '''
    try:
        files = os.listdir('.\\' + folder)
        for file in files:
            os.remove('.\\' + folder + '\\' + file)
    except:
        print("PDF-Export folder is empty or doesn't exist.")
        
def del_allBookmarks(pdfList, inFolder):
    '''
    Deletes all bookmarks of selected PDFs by copying page contents and
    overwriting "old" PDF-files.
    Takes list of PDF-files as an argument.
    '''
    for file in pdfList:
        # open PdfFileWriter
        pdfWriter = PdfFileWriter()
        # open PDF-file
        inputFile = open(inFolder + "\\" + file, 'rb')
        myPDF = PdfFileReader(inputFile, strict=False)
        # "copy" pages
        for page in range(0, myPDF.numPages):
            pdfWriter.addPage(myPDF.getPage(page))
        # saving file
        outputFile = exportFolder + "\\" + file
        with open(outputFile, 'wb') as resultPDF:
            pdfWriter.write(resultPDF)
            
def convert_xls(xlsList, inFolder):
    '''
    Function to convert xls-files to pdf format. Needs MS Excel to be installed
    in order to work.
    '''
    for file in xlsList:
        excel = client.Dispatch('Excel.Application')
        books = excel.Workbooks.Open(inFolder + '\\' + file)
        sheet = books.Worksheets[0]
        sheet.ExportAsFixedFormat(0, exportFolder + '\\' + file[:len(file)-5] + '.pdf')
        books.Close()
        excel.Quit()

#def convert_doc(docList, inFolder):
#    # list only doc-files in folder
#    docFiles = [file for file in docList if file.endswith('.doc')]
#    for file in docFiles:
#        word = client.Dispatch('Word.Application')
#        doc = word.Documents.Open(inFolder + '\\' + file)
#        doc.ExportAsFixedFormat(0, inFolder +  '\\' + file[:len(file)-5] + '.pdf')
#        doc.Close()
#        word.Quit()
    
def merge_PDFs(pdfList, inFolder):
    '''
    Merges selected input PDFs to one output PDF.
    '''
    # create export folder if it doesn't exist
    if os.path.exists(exportFolder) == False:
        os.makedirs(exportFolder)
    # open PdfFileMerger    
    pdfMerger = PdfFileMerger(strict=False)
    for file in pdfList:
        # open PDF-file
        inputFile = open(inFolder + "\\" + file, 'rb')
        myPDF = PdfFileReader(inputFile, strict=False)
        # merge PDF-files
        pdfMerger.append(myPDF)
    # produce output PDF
    outputFile = exportFolder + '\\mergedFile.pdf'
    with open(outputFile, 'wb') as resultPDF:
        pdfMerger.write(resultPDF)    
     
     
### APP FUNCTIONS

def btn_GetList():
    '''
    Create button to get a list of all PDF-Files.
    '''
    app.clearAllListBoxes()
    items = list_files('.pdf')
    app.updateListBox("list", items)
    
def btn_MergePDFs():
    '''
    Merges selected PDFs tp one Document.
    '''
    pdfList = app.getListBox("list")
    merge_PDFs(pdfList, inFolder)

def btn_ClearList():
    '''
    Create button to clear the List.
    '''
    app.clearAllListBoxes()
    
def btn_addBookmarks():
    '''
    Create button to add page booksmarks to PDFs in ListBox.
    '''
    pdfList = app.getListBox("list")
    add_pageBookmarks(pdfList, inFolder)
    
def btn_delExportFolder(folder):
    '''
    Create button to delete all files in Export-PDF-Folder.
    '''
    del_allFiles(folder)
    
def btn_delBookmarks():
    '''
    Deletes all bookmarks of the selected PDFs in ListBox.
    '''
    pdfList = app.getListBox("list")
    del_allBookmarks(pdfList, inFolder)

#########################
##### START PROGRAM #####
#########################

# set subfolder to store/edit new PDF files
exportFolder = ".\\PDF-Export"

# set folder with input files, created for test purposes
inFolder = ".\\Test files"

# create a GUI variable
app = gui(showIcon=False)

# add & configure widgets
row = app.getRow()
app.addLabel("listBox1", "Files in folder:", row, 0)
app.addLabel("listBox2", os.path.basename(os.getcwd()) + inFolder, row, 1)
row = app.getRow()
app.addListBox("list", [], row = row, colspan = 2)
app.setListBoxMulti("list", multi=True)

# add buttons
# PDF list group
row = app.getRow()
app.addLabel("listBtns", "Files", row = row, colspan = 2)
row = app.getRow()
app.addButton("Get PDFs", btn_GetList, row, 0)
app.addButton("Merge PDFs", btn_MergePDFs, row, 1)
row = app.getRow()
app.addButton("Clear list", btn_ClearList, row, 0)

# bookmark group
row = app.getRow()
app.addLabel("bookmarkBtns", "Bookmarks", row = row, colspan = 2)
row = app.getRow()
app.addButton("Add page bookmarks", btn_addBookmarks, row, 0)
app.addButton("Delete bookmarks", btn_delBookmarks, row, 1)
app.addNamedButton(name='Empty "Export-PDF"',title='PDF-Export', func=btn_delExportFolder)

# start the GUI
app.go()
