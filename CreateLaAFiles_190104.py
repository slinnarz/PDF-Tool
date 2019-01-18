# -*- coding: utf-8 -*-

'''
TO DOs:
    - Finish GUI layout
    - Convert MS Office docs to PDF.   
IN PROGRESS:    
    - Add buttons:

'''

############################
##### IMPORT PACKAGES ######
############################

import os
from appJar import gui
from PyPDF2 import PdfFileWriter, PdfFileReader


#############################
##### DEFINE FUNCTIONS ######
#############################

### GENERAL FUNCTIONS

def list_files(dataType):
    '''
    Lists all files of specified type in the current folder.
    Takes data type as string (e.g. '.pdf') as argument.
    '''
    files = os.listdir('.\\')
    return [file for file in files if file.endswith(dataType)]

def add_pageBookmarks(pdfFiles):
    '''
    Adds bookmarks for page numbers, if the document has >1 pages.
    Takes a list of PDF-files as an argument.
    '''
    if os.path.exists(exportFolder) == False:
        os.makedirs(exportFolder)
    for file in pdfFiles:
        # open PdfFileWriter
        pdfWriter = PdfFileWriter()
        
        # open PDF-file
        inputFile = open(file, 'rb')
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
        
def del_allBookmarks(pdfList):
    '''
    Deletes all bookmarks of selected PDFs by copying page contents and
    overwriting "old" PDF-files.
    Takes list of PDF-files as an argument.
    '''
    for file in pdfList:
        # open PdfFileWriter
        pdfWriter = PdfFileWriter()
        
        # open PDF-file
        inputFile = open(file, 'rb')
        myPDF = PdfFileReader(inputFile, strict=False)
        
        # "copy" pages
        for page in range(0, myPDF.numPages):
            pdfWriter.addPage(myPDF.getPage(page))
            
        # saving file
        outputFile = exportFolder + "\\" + file
        with open(outputFile, 'wb') as resultPDF:
            pdfWriter.write(resultPDF)
    
    
### APP FUNCTIONS

def btn_GetList():
    '''
    Create button to get a list of all PDF-Files.
    '''
    app.clearAllListBoxes()
    items = list_files('.pdf')
    app.updateListBox("list", items)

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
    add_pageBookmarks(pdfList)
    
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
    del_allBookmarks(pdfList)

#########################
##### START PROGRAM #####
#########################

# set subfolder to store/edit new PDF files
exportFolder = ".\\PDF-Export"

# create a GUI variable
app = gui(showIcon=False)

# add & configure widgets
app.addLabel("listBox1", "Files in folder:", row = 0, colspan = 2)
app.addLabel("listBox2", os.path.basename(os.getcwd()), row = 1, colspan = 2)
app.addListBox("list", [], row = 2, colspan = 2)
app.setListBoxMulti("list", multi=True)

# add buttons
# PDF list group
app.addLabel("listBtns", "Files", row = 3, colspan = 2)
app.addButton("Get PDFs", btn_GetList, 4, 0)
app.addButton("Clear list", btn_ClearList, 4, 1)

# bookmark group
app.addLabel("bookmarkBtns", "Bookmarks", row = 5, colspan = 2)
app.addButton("Add page bookmarks", btn_addBookmarks, 6, 0)
app.addButton("Delete bookmarks", btn_delBookmarks, 6, 1)
app.addNamedButton(name='Empty "Export-PDF"',title='PDF-Export', func=btn_delExportFolder)

# start the GUI
app.go()