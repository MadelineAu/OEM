import os, PyPDF2

def keywordWriter():
        print ('Hello World')
        pdfWriter = PyPDF2.PdfFileWriter()
        pdfReader = PyPDF2.PdfFileReader(TestFiles/test.pdf)
        keyword = 'A123'

        pdfInfo = pdfReader.getDocumentInfo()
        print(str(pdfInfo))

        keywordIn = pdfInfo['/Keywords']

        print(keywordIn)

        print ('Number of pages ', pdfReader.numPages)
        print ('Keyword ', keyword)

        keyword = ', '.join([keywordIn, 'C123'])
        print ('Keyword ', keyword)

        for pageNum in range(pdfReader.numPages):
                pageObj = pdfReader.getPage(pageNum)
                pdfWriter.addPage(pageObj)
        
        pdfWriter.addMetadata({'/Author': 'madeline', '/Keywords': keyword})

        pdfOutputFile = open('outputFile.pdf', 'wb')
        pdfWriter.write(pdfOutputFile)
        pdfOutputFile.close()
        pdfFileObj.close()

keywordWriter()


