# PDF-PPT-Merger
Python Script to merge all PDF Files and PowerPoint files found within a folder together, ending with a single file as "FOLDERNAME.extension".

# Purpose
Designed to aid in a class where Slides (PPTX) and Transcripts (PDF) could be downloaded for each module and used to learn/study the material. Instead of having to open various files these could be merged together into a single file per module. 

# Files
* merger.bat -- Drop Target (Click and Drag Folder on to this file)
* merger.py  -- Python Script to Merge Various PDF and PPTX Files in Directory

# Dependencies
* PyPDF2 -- https://pypi.org/project/PyPDF2/
* cloudmersive_convert_api_client -- https://api.cloudmersive.com/ (Needs an API Key)
* glob -- Finds files using a pattern string

# Setup
* Install Python3

* Use Pip to Install Packages
  * pip install pypdf2
  * pip install cloudmersive_convert_api_client
  * pip install glob

* Place batch file and python file in directory where folder containing PDF/PPTX files is. 
