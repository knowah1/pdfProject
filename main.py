import os
from tkinter import filedialog as fd
import PyPDF2
from PyPDF2 import PdfFileWriter

# this script splits out pages of a pdf invoice and renames them after the invoice #

invoicePDF = fd.askopenfilename() # user picks a file
os.chdir(os.path.dirname(invoicePDF)) # set working dir to path of file

with open(invoicePDF, 'rb') as pdfFileObj:
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    print(pdfReader.numPages - 2)
    x = 1     # skips cover page

    while x < pdfReader.numPages:
        pageObj = pdfReader.getPage(x)
        pageText = pageObj.extractText()
        startingPoint = pageText.find('Invoice:') + 9 #Penske
        endPoint = startingPoint + 10 #Penske
        #startingPoint = pageText.find('Bill Ref #') + 11 #Enterprise
        #endPoint = startingPoint + 12 #Enterprise
        invoiceNum = pageText[startingPoint:endPoint]
    
        output = PdfFileWriter()
        output.addPage(pdfReader.getPage(x))
        with open(str(invoiceNum) + '.pdf', 'wb') as outputStream:
            output.write(outputStream)
        x += 1

print('FINISHED!')
