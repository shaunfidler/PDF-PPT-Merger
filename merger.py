from __future__ import print_function
from PyPDF2 import PdfMerger
import glob
import sys
import time
from os import walk
import comtypes.client
from pptx.util import Inches
from pptx import Presentation
import copy

import win32com.client
from pptx import Presentation

def PPTXMerge(inputFileNames, outputFileName):
    Application = win32com.client.Dispatch("PowerPoint.Application")
    outputPresentation = Application.Presentations.Add() 
    outputPresentation.SaveAs(outputFileName)

    for file in inputFileNames:    
        currentPresentation = Application.Presentations.Open(file)
        currentPresentation.Slides.Range(range(1, currentPresentation.Slides.Count+1)).copy()
        Application.Presentations(outputFileName).Windows(1).Activate()    
        outputPresentation.Application.CommandBars.ExecuteMso("PasteSourceFormatting")    
        currentPresentation.Close()

    outputPresentation.save()
    outputPresentation.close()
    Application.Quit()


def PDFMerge(files, outputName, appender):
    print("*** BEGINNING PDF MERGE ***")
    #print(files)
    #print(outputName)

    merger = PdfMerger()
    for pdf in files:
        print(pdf)
        merger.append(open(pdf, 'rb'))
    filename = outputName.split('\\')
    outputName = outputName + '\\' + filename[len(filename) - 1]
    with open(outputName + ' ' + appender + '.pdf', 'wb') as fout:
        merger.write(fout)
    print("*** SUCCESSFUL PDF MERGE ***\n")

def initPPT():
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    return powerpoint
    
def convertToPDF(files, powerpoint, pdfsToMerge):
    for file in files:
        outputFileName = file
        formatType = 32
        
        if outputFileName[-3:] != 'pdf':
            outputFileName = outputFileName + ".pdf"
        deck = powerpoint.Presentations.Open(file)

        deck.SaveAs(outputFileName, formatType) # formatType = 32 for ppt to pdf
        deck.Close()
        pdfsToMerge.append(outputFileName)
        print("DONE WITH: " + outputFileName)

def FindFiles(directory, pdf, ppt, merge_pptx):
    if pdf:
        pdfs = glob.glob(directory + "\\*.pdf")
        PDFMerge(pdfs, directory, 'Transcripts')
    if ppt:
        ppts = glob.glob(directory + "\\*.pptx")
        pdfsToMerge = []
        powerpoint = initPPT()
        convertToPDF(ppts, powerpoint, pdfsToMerge)
        powerpoint.Quit()
        PDFMerge(pdfsToMerge, directory, 'Slides')
    if merge_pptx:
        ppts = glob.glob(directory + "\\*.pptx")
        PPTXMerge(ppts, directory + '\\Merged Slides.pptx')


def main():
    directory = sys.argv[1]
    print("**********************************")
    print("MENU:")
    print("Type 1 for PDF and PPTX Merge")
    print("Type 2 for PDF Merge Only")
    print("Type 3 for PPTX Merge Only")
    print("Type 4 for PPTX Merge into single PPTX")
    print("**********************************")
    choice = input()
    pdf = True
    ppt = True
    merge_pptx = False
    if(choice == '1'):
        FindFiles(directory, pdf, ppt, merge_pptx)
    elif(choice == '2'):
        ppt = False
        FindFiles(directory, pdf, ppt, merge_pptx)
    elif(choice == '3'):
        pdf = False
        FindFiles(directory, pdf, ppt, merge_pptx)
    elif(choice == '4'):
        pdf = False
        ppt = False
        merge_pptx = True
        FindFiles(directory, pdf, ppt, merge_pptx)
    else:
        print("Bad Input -- Relaunching Menu")
        main()


if __name__ == "__main__":
    main()