import os, PyPDF2, openpyxl, shutil
from pathlib import Path

# Loops through maintenance sheet
# Matches part number on maintenance sheet to part number on spares sheet
# Writes OEM filenames from maintenance sheet to spares sheet

def oemToSparesSheet(sparesFile):
         wb = openpyxl.load_workbook(sparesFile)
         sparesSheet = wb['spare parts']
         maintSheet = wb['maintenance']
         oemFiles  = []
         
         for rowNumMaint in range(1, maintSheet.max_row + 1):
                maintPartName = maintSheet.cell(row=rowNumMaint, column= 3).value
                maintOem = maintSheet.cell(row=rowNumMaint, column= 4).value
                
                if (maintOem is not None) and not ('www.' in maintOem):
                        #Find relevant row in spares sheet
                         for rowNumSpares in range(1, sparesSheet.max_row + 1):
                             sparesPartStr = sparesSheet.cell(row=rowNumSpares,column=2).value
                             quan = sparesSheet.cell(row=rowNumSpares,column=3).value
                             # print ("maintPartName ", maintPartName, " sparesPartStr ", sparesPartStr)
                             if (sparesPartStr is not None) and (quan is not None) and (maintPartName in sparesPartStr):
                                    print("Match ", maintPartName, " ", sparesPartStr)
                                    #ToDo - write OEM from maint sheet to spares sheet
                                    # Add this file to Git
       
                                    


##                        #create OEM list
##                        # if len(oem) > 5:
##                              #  print('OEM length > 5')
##                        oem = oem.replace(" ", "")
##                        oemList = oem.split(',')
##                        print("oemList ",oemList)
##
##                        for oemFile in oemList:
##                                #print('Find oem file ', oemFile)
##                                for folderName, subfolders, filenames in os.walk('C:\\Users\\Mick\\Documents\\OEM'):
##                                        for filename in filenames:
##                                                # When OEM matches filename, call writeKeyword
##                                                # Need to split strings with a comma and loop
##                                                oemFilePath = os.path.join(folderName, filename)
##                                                # print("oemFile ", oemFile)
##                                                if oemFile in filename:
##                                                        print("Match oemFile filename", oemFile, filename)
##                                                        print('Path and Filename ', oemFilePath)
##                                                        # overwrites when call twice. Look at write keyword
##                                                        writeKeyword(oemFilePath, partNum)
##                                                        # writeKeyword(oemFilePath, partMan)
##                # if not none. Test for OEM file. Then filewalk in OEM directory. Change to right place.
##                # Call writeKeyword
                

oemToSparesSheet('TestFiles/OEMtoSparesSheet.xlsx')
