from __future__ import print_function
from PyPDF2 import PdfMerger
import glob
import sys
import time
import os
import comtypes.client

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

def FindFiles(directory, pdf, ppt):
    if(pdf):
        pdfs = glob.glob(directory + "\\*.pdf")
        PDFMerge(pdfs, directory, 'Transcripts')
    if(ppt):
        ppts = glob.glob(directory + "\\*.pptx")
        pdfsToMerge = []
        powerpoint = initPPT()
        convertToPDF(ppts, powerpoint, pdfsToMerge)
        powerpoint.Quit()
        PDFMerge(pdfsToMerge, directory, 'Slides')

def main():
    directory = sys.argv[1]
    print("**********************************")
    print("MENU:")
    print("Type 1 for PDF and PPTX Merge")
    print("Type 2 for PDF Merge Only")
    print("Type 3 for PPTX Merge Only")
    print("**********************************")
    choice = input()
    pdf = True
    ppt = True
    if(choice == '1'):
        FindFiles(directory, pdf, ppt)
    elif(choice == '2'):
        ppt = False
        FindFiles(directory, pdf, ppt)
    elif(choice == '3'):
        pdf = False
        FindFiles(directory, pdf, ppt)
    else:
        print("Bad Input -- Relaunching Menu")
        main()

if __name__ == "__main__":
    main()