import os
from tkinter import filedialog as fd
import PyPDF2
from PyPDF2 import PdfWriter

# this script splits out pages of a pdf invoice and renames them after the invoice #

invoicePDF = fd.askopenfilename() # user picks a file
os.chdir(os.path.dirname(invoicePDF)) # set working dir to path of file

with open(invoicePDF, 'rb') as pdfFileObj: #could -1 for Penske to prevent blank page
    pdfReader = PyPDF2.PdfReader(pdfFileObj)
    print(len(pdfReader.pages) - 2)

    x = 1     # skips cover page(s)
    print(invoicePDF)
    while x < len(pdfReader.pages):
        pageObj = pdfReader.pages[x]
        pageText = pageObj.extract_text()
        print('in while')
        if 'Penske' in os.getcwd():
            startingPoint = pageText.find('Invoice:') + 9 #Penske
            endPoint = startingPoint + 10 #Penske
            invoiceNum = pageText[startingPoint:endPoint]
        if 'Enterprise' in os.getcwd():
            startingPoint = pageText.find('Bill Ref #') + 11 #Enterprise
            endPoint = startingPoint + 12 #Enterprise
            invoiceNum = pageText[startingPoint:endPoint]
        if 'Carbon' in os.getcwd():
            startingPoint = pageText.find('INVOICE #') + 10  # Carbon Solutions
            endPoint = startingPoint + 5  # Carbon Solutions
            invoiceNum = pageText[startingPoint:endPoint]  # carbon solutions
            invoiceNum = invoiceNum + " - Carbon Solutions"


    
        output = PdfWriter()
        output.add_page(pdfReader.pages[x])
        with open(str(invoiceNum) + '.pdf', 'wb') as outputStream:
            output.write(outputStream)
        x += 1

print('FINISHED!')
