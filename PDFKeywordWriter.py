import os, PyPDF2

print ('Hello World')

pdfFileObj = open('test.pdf', 'rb')
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
pdfWriter = PyPDF2.PdfFileWriter()
print ('Number of pages ', pdfReader.numPages)

for pageNum in range(pdfReader.numPages):
        pageObj = pdfReader.getPage(pageNum)
        pdfWriter.addPage(pageObj)
        
pdfWriter.addMetadata({'/Author': 'madeline', '/Keywords': 'C123'})

pdfOutputFile = open('outputFile.pdf', 'wb')
pdfWriter.write(pdfOutputFile)
pdfOutputFile.close()
pdfFileObj.close()


