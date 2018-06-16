import os, PyPDF2, openpyxl, shutil

def sparesFileWalk(projectDir):
        for folderName, subfolders, filenames in os.walk(projectDir):
                print('Current folder ', folderName)

                for subfolder in subfolders:
                        print('Subfolder of ', folderName, ' ', subfolder)

                for fileaname in filenames:
                        print('File inside ', folderName 

def OEMToSpareParts(sparesFile):
        wb = openpyxl.load_workbook(sparesFile)
        print ('Sheetnames ', wb.sheetnames)
        sparesSheet = wb['Spares']
        maintSheet = wb['Maintenance']
        print ('Spares sheet Cell A1 ', sparesSheet['A1'].value)
        for rowNum in range(1, sparesSheet.max_row + 1):
                sparesCell = sparesSheet.cell(row=rowNum, column=2)
                for maintCell in maintSheet['B']:
                       print(sparesCell.value, ' ', maintCell.value)
                       if (sparesCell.value == maintCell.value):
                               print('MATCH')
                               OEM =  maintSheet.cell(rowNum, 3)
                               print(OEM.value)
                               sparesSheet.cell(row=rowNum, column=3).value = OEM.value
                               # write OEM into the row on maintSheet
        #write out to new file
        outFile = '_'.join(['Out', sparesFile])
        wb.save(outFile)


OEMToSpareParts('Spares.xlsx')

# Appends keyword to current list of keywords in pdfFile and writes the file
# to Out_pdffile 
def writeKeyword(pdfFile, keyword):
        pdfFileObj = open(pdfFile, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        pdfWriter = PyPDF2.PdfFileWriter()

        pdfInfo = pdfReader.getDocumentInfo()
        print('PDFInfo ', str(pdfInfo))

        keywordIn = pdfInfo['/Keywords']
        print('keywordIn ', keywordIn)
        print ('Keyword ', keyword)

        keyword = ', '.join([keywordIn, keyword])
        print ('Keyword ', keyword)

        for pageNum in range(pdfReader.numPages):
                pageObj = pdfReader.getPage(pageNum)
                pdfWriter.addPage(pageObj)
        
        pdfWriter.addMetadata({'/Author': 'madeline', '/Keywords': keyword})

        pdfOutfile = '_'.join(['Out', pdfFile])

        if (os.path.exists(pdfOutfile)):
                os.remove(pdfOutfile)
                 
        pdfOutputFile = open(pdfOutfile, 'wb')
        pdfWriter.write(pdfOutputFile)
        pdfOutputFile.close()
        pdfFileObj.close()


#writeKeyword('test.pdf', '7Mar')
