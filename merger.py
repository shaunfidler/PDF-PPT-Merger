from __future__ import print_function
from PyPDF2 import PdfFileMerger
import glob
import sys
import time
import cloudmersive_convert_api_client
from cloudmersive_convert_api_client.rest import ApiException
import os

def PDFMerge(files, outputName):
	print("*** BEGINNING PDF MERGE ***")
	#print(files)
	#print(outputName)

	# Merge the list of PDFs passed in (in order)
	merger = PdfFileMerger()
	for pdf in files:
		merger.append(open(pdf, 'rb'))

	# Generate final PDF name
	filename = outputName.split('\\')
	outputName = outputName + '\\' + filename[len(filename) - 1]

	# Write final PDF as <FOLDERNAME>.pdf
	with open(outputName + '.pdf', 'wb') as fout:
		merger.write(fout)
	print("*** SUCCESSFUL PDF MERGE ***\n")

def PPTMerger(files, outputName):
	print("*** BEGINNING PPTX MERGE ***")
	apiKey = 'API KEY HERE'
	#print(files)
	#print(outputName)

	# Start building filename to be FolderName
	filename = outputName.split('\\')
	filename = outputName + '\\' + filename[len(filename) - 1]

	# Configure API key authorization: Apikey
	configuration = cloudmersive_convert_api_client.Configuration()
	configuration.api_key['Apikey'] = apiKey

	# Create an instance of the API class
	api_instance = cloudmersive_convert_api_client.MergeDocumentApi(cloudmersive_convert_api_client.ApiClient(configuration))

	try:
		# Merge first two PPTXs to start the process
		api_response = api_instance.merge_document_pptx_multi(files[0], files[1])
		with open(outputName + '\\tmp1.pptx', 'wb') as fout:
			fout.write(api_response)
		
		# Use temporary PPTXs to merge remaining PPTXs all into a single file, using the previous merge each time (MERGED.pptx + next.pptx)
		for i in range(2, len(files)):
			# Merge with previous temporary pptx
			api_response = api_instance.merge_document_pptx_multi(outputName + '\\tmp{}.pptx'.format(i - 1), files[i])
			with open(outputName + '\\tmp{}.pptx'.format(i), 'wb') as fout:
				fout.write(api_response)

		# Save final merge as '<FOLDERNAME>.pptx'
		with open(filename + '.pptx', 'wb') as fout:
			fout.write(api_response)
		removeFiles = glob.glob(outputName + '\\tmp*.pptx')

		# Remove temporary files used during merging
		for f in removeFiles:
			os.remove(f)
		print("*** SUCCESSFUL PPTX MERGE ***")
	except ApiException as e:
		print("Exception when calling MergeDocumentApi->merge_document_pptx: %s\n" % e)

def FindFiles(directory, pdf, ppt):
	if(pdf):
		pdfs = glob.glob(directory + "\\*.pdf")
		PDFMerge(pdfs, directory)
	if(ppt):
		ppts = glob.glob(directory + "\\*.pptx")
		PPTMerger(ppts, directory)

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
