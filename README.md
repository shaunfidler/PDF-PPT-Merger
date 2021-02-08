# PDF-PPT-Merger
Python Script to merge the PDF Files and/or PowerPoint files found within a folder together, ending with a single file (One for combined PPTXs and one for combined PDFs).

# Purpose
Designed to aid in a class where Slides (PPTX) and Transcripts (PDF) could be downloaded for each module and used to learn/study the material. Instead of having to open various files these could be merged together into 2 files per module, one for slides and one for transcripts. 

# Files
* merger.bat -- Drop Target (Click and Drag Folder on to this file)
* merger.py  -- Python Script to Merge Various PDF and PPTX Files in Directory

# Dependencies
* PyPDF2 -- https://pypi.org/project/PyPDF2/

# Setup
* Install Python3

* Use Pip to Install Packages
  * pip install pypdf2

* Place batch file and python file in directory where the folder containing PDF/PPTX files is (An example structure is shown below).

# Example File Structure:
* ROOT DIRECTORY
  * Module 0x09  <-- Contains the PDFs or PPTXs that you want to merge
  * merger.bat <-- Grab the 'Module 01' folder and drop it on this file to run the program
  * merger.py
  
Once completed, the ending file will be 'Module 0x09\Module 0x09 Slides.pdf' if merging PPTXs or 'Module 0x09\Module 0x09 Transcripts.pdf' if merging PDFs. This can be customized in the python code.

# Example of Program Running
![Example of Program Running](https://github.com/shaunfidler/PDF-PPT-Merger/blob/main/Merger%20Running.jpg)
