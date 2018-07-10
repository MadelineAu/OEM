import os, PyPDF2, openpyxl, shutil
from pathlib import Path

# Loops through maintenance sheet
# Matches part number on maintenance sheet to part number on spares sheet
# Writes OEM filenames from maintenance sheet to spares sheet
# Works for all rows where the part number used in the maint sheet is a
# substring of the string in the spare parts sheet (see ToDo)

#----------------------------------------------------------------------------------------------------------------
def oemToSparesSheet(sparesFile):
         wb = openpyxl.load_workbook(sparesFile)
         sparesSheet = wb['spare parts']
         maintSheet = wb['maintenance']
         oemFiles  = []
         
         for rowNumMaint in range(1, maintSheet.max_row + 1):
                maintPartName = maintSheet.cell(row=rowNumMaint, column= 3).value
                maintOem = maintSheet.cell(row=rowNumMaint, column= 4).value
                maintMan = maintSheet.cell(row=rowNumMaint, column= 2).value
                
                if (maintOem is not None) and not ('www.' in maintOem):
                        #Find relevant row in spares sheet
                         for rowNumSpares in range(1, sparesSheet.max_row + 1):
                             sparesPartStr = sparesSheet.cell(row=rowNumSpares,column=2).value
                             quan = sparesSheet.cell(row=rowNumSpares,column=3).value
                             # print ("maintPartName ", maintPartName, " sparesPartStr ", sparesPartStr)
                             if (sparesPartStr is not None) and (quan is not None) and (maintPartName in sparesPartStr):
                                    # print("Match ", maintPartName, " ", sparesPartStr)
                                    sparesSheet.cell(row=rowNumSpares,column=6).value = maintMan
                                    sparesSheet.cell(row=rowNumSpares,column=7).value = maintOem

         # Save new copy of spares file in same directory, prefixed with Out_
         filename = os.path.basename(sparesFile)
         directory = os.path.dirname(sparesFile)
         print("Filename ", filename, " Directory ", directory)
         outFile = '_'.join(['Out', filename])
         wb.save(os.path.join(directory, outFile))
									
	#ToDo - print cases (save to Excel?) where couldn't find part number in spares sheet (eg, Sew motor use S Series as part name in maintenance sheet)

#------------------------------------------------------------------------------------------------------------------

# Appends keyword to current list of keywords in pdfFile and writes the file
# to Out_pdffile

def writeKeyword(pdfFile, keyword):
        pdfFileObj = open(pdfFile, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        pdfWriter = PyPDF2.PdfFileWriter()

        pdfDirectory = os.path.dirname(pdfFile)
        pdfFileBasename = os.path.basename(pdfFile)
        pdfOutfile = '_'.join(['Out', pdfFileBasename])

        pdfInfo = pdfReader.getDocumentInfo()
        keywordIn = ""
        # print('PDFInfo ', str(pdfInfo))

        if (os.path.exists(os.path.join(pdfDirectory, pdfOutfile))):
            print("hello!")
            pdfFileObjOutFile = open(os.path.join(pdfDirectory, pdfOutfile), 'rb')
            pdfReaderOutFile = PyPDF2.PdfFileReader(pdfFileObjOutFile)
            pdfInfoOutFile = pdfReaderOutFile.getDocumentInfo()
            if '/Keywords' in pdfInfoOutFile:
                keywordIn = pdfInfoOutFile['/Keywords']
                print('keywordIn ', keywordIn)
        else:
            if '/Keywords' in pdfInfo:
                keywordIn = pdfInfo['/Keywords']
                # print('keywordIn ', keywordIn)
                # print ('Keyword ', keyword)

        keyword = ', '.join([keywordIn, keyword])
                # print ('Keyword ', keyword)

                
            # If this, get keywords from outfile
            # Else, get keywords from infile

        

        for pageNum in range(pdfReader.numPages):
                pageObj = pdfReader.getPage(pageNum)
                pdfWriter.addPage(pageObj)
        
        pdfWriter.addMetadata({'/Author': 'madeline', '/Keywords': keyword})

        

        if (os.path.exists(os.path.join(pdfDirectory, pdfOutfile))):
                print("Outfile exists")
                pdfFileObjOutFile.close()
                os.remove(os.path.join(pdfDirectory, pdfOutfile))
                
        pdfOutputFile = open(os.path.join(pdfDirectory, pdfOutfile), 'wb')
        pdfWriter.write(pdfOutputFile)
        # print("pdfOutfile ", pdfOutfile)
        # print("pdfFileBasename ", pdfFileBasename)
        pdfOutputFile.close()
        pdfFileObj.close()

        # Need to figure out how to rename Out_ file to original file name - use shutil.copy 2(src/dst)
        # Then delete Out_file 
        # os.rename(pdfOutfile, pdfFile)


##writeKeyword('TestFiles/test.pdf', 'e')
##writeKeyword('TestFiles/test.pdf', 'f')
##writeKeyword('TestFiles/test.pdf', 'g')

#--------------------------------------------------------------------------------------------


 # Loops through spare parts worksheetand calls writeKeyword for each spare part with an
 # OEM file listed and writes the 20,000,000 and Manufacturer to the Keywords metadata

def OemManufacturerToPDFKeyword(sparesFile):
    wb = openpyxl.load_workbook(sparesFile)
    ws = wb['spare parts']
    oemFiles  = []	

    for rowNum in range(1, ws.max_row + 1):
        partMan = ws.cell(row=rowNum, column=6).value
        oem = ws.cell(row=rowNum, column= 7).value
        partNum = ws.cell(row=rowNum, column= 1).value
				
        if (oem is not None) and not ('www.' in oem):
            #create OEM list
            # if len(oem) > 5:
            #  print('OEM length > 5')
            oem = oem.replace(" ", "")
            oemList = oem.split(',')
            print("oemList ",oemList)

            for oemFile in oemList:
                print('Find oem file ', oemFile)
                for folderName, subfolders, filenames in os.walk('C:\\Users\\Mick\\Documents\\OEM'):
                    for filename in filenames:
                        # When OEM matches filename, call writeKeyword
                        # Need to split strings with a comma and loop
                        oemFilePath = os.path.join(folderName, filename)
                        # print("oemFile ", oemFile)
                        if oemFile in filename:
                            print("Match oemFile filename", oemFile, filename)
                            print('Path and Filename ', oemFilePath)
                            writeKeyword(oemFilePath, partNum)
                            writeKeyword(oemFilePath, partMan)

OemManufacturerToPDFKeyword('TestFiles/Out_DrySF_31621699.xlsx')  
													

#oemToSparesSheet('TestFiles/OEMtoSparesSheet.xlsx')
