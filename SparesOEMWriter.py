import os, PyPDF2, openpyxl, shutil
from pathlib import Path

# Need to figure out how to navigate to OEM directory to do the keyword writing. 

def OEMToSpareParts(sparesFile):
        wb = openpyxl.load_workbook(sparesFile)
        print ('Sheetnames ', wb.sheetnames)
        sparesSheet = wb['spare parts']
        maintSheet = wb['maintenance']
        print ('Spares sheet Cell A1 ', sparesSheet['A1'].value)
        for sparesRow in range(1, sparesSheet.max_row + 1):
                sparesCell = sparesSheet.cell(row=sparesRow, column=2)
                for maintRow in range(1, maintSheet.max_row +1):
                        maintCell = maintSheet.cell(row=maintRow, column=3)
                        sparesNumber = sparesSheet.cell(row=sparesRow, column=1).value
                        #print('sparesNumber ', sparesNumber)
                        print(sparesCell.value, ' ', maintCell.value)
                        if maintCell.value is not None:
                                if (maintCell.value in sparesCell.value) and (sparesNumber.startswith('20')):
                                        manufacturer = maintSheet.cell(maintRow, 2)
                                        OEM =  maintSheet.cell(maintRow, 4)
                                        print('Match OEM ', OEM.value)
                                        sparesSheet.cell(row=sparesRow, column=6).value = manufacturer.value
                                        sparesSheet.cell(row=sparesRow, column=7).value = OEM.value
                               # write OEM into the row on maintSheet
        #write out to new file
        outFile = '_'.join(['Out', sparesFile])
        wb.save(outFile)


#OEMToSpareParts('DrySF_31621699.xlsx')


def sparesFileWalk(projectDir):
        for folderName, subfolders, filenames in os.walk(projectDir):
                print('Current folder ', folderName)

                for filename in filenames:
                        #print('File inside ', folderName, ': ', filename)
                        #Check older file suffix is correct
                        if (Path(filename).suffix) == ('.xlsx' or '.xsl'):
                                print('XL file ',Path(filename).suffix)
                                wb = openpyxl.load_workbook(filename)
                                print(wb.sheetnames)
                                if ('spare parts' in wb.sheetnames and 'maintenance' in wb.sheetnames):
                                        print('Hello Spare Parts List')
                                        OEMToSpareParts(filename)
                        
                # Next - Identify spares files

# sparesFileWalk('.\\')


# ToDO - currently writes E2932 as keyword everytime and saving Out files in directory this is
# run from.  Needs to save back in file it came from. 

# Appends keyword to current list of keywords in pdfFile and writes the file
# to Out_pdffile 
def writeKeyword(pdfFile, keyword):
        pdfFileObj = open(pdfFile, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        pdfWriter = PyPDF2.PdfFileWriter()

        pdfInfo = pdfReader.getDocumentInfo()
        # print('PDFInfo ', str(pdfInfo))

        if '/Keywords' in pdfInfo:
                keywordIn = pdfInfo['/Keywords']
                # print('keywordIn ', keywordIn)
                # print ('Keyword ', keyword)

                keyword = ', '.join([keywordIn, keyword])
                # print ('Keyword ', keyword)

        for pageNum in range(pdfReader.numPages):
                pageObj = pdfReader.getPage(pageNum)
                pdfWriter.addPage(pageObj)
        
        pdfWriter.addMetadata({'/Author': 'madeline', '/Keywords': keyword})

        pdfFileBasename = os.path.basename(pdfFile)
        pdfOutfile = '_'.join(['Out', pdfFileBasename])

        if (os.path.exists(pdfOutfile)):
                os.remove(pdfOutfile)
                 
        pdfOutputFile = open(pdfOutfile, 'wb')
        pdfWriter.write(pdfOutputFile)
        pdfOutputFile.close()
        pdfFileObj.close()


#writeKeyword('test.pdf', '7Mar')

# Loops through spare parts worksheetand calls writeKeyword for each spare part with an
# OEM file listed and writes the 20,000,000 and Manufacturer to the Keywords metadata
def sparesToPDFKeyword(sparesFile):
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
                                #print('Find oem file ', oemFile)
                                for folderName, subfolders, filenames in os.walk('C:\\Users\\Mick\\Documents\\OEM'):
                                        for filename in filenames:
                                                # When OEM matches filename, call writeKeyword
                                                # Need to split strings with a comma and loop
                                                oemFilePath = os.path.join(folderName, filename)
                                                # print("oemFile ", oemFile)
                                                if oemFile in filename:
                                                        print("Match oemFile filename", oemFile, filename)
                                                        print('Path and Filename ', oemFilePath)
                                                        # overwrites when call twice. Look at write keyword
                                                        writeKeyword(oemFilePath, partNum)
                                                        # writeKeyword(oemFilePath, partMan)
                # if not none. Test for OEM file. Then filewalk in OEM directory. Change to right place.
                # Call writeKeyword
                

sparesToPDFKeyword('Out_DrySF_31621699.xlsx')
