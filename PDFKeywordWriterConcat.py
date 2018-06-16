import os, PyPDF2

print ('Hello World')

pdfFileObj = open('test.pdf', 'rb')
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
pdfWriter = PyPDF2.PdfFileWriter()
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



