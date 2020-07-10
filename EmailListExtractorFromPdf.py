# -*- coding: utf-8 -*-
"""

@author: Mrinal Prakash
"""

def pdf_to_text_file(file_path):
    from pdfminer.pdfparser import PDFParser, PDFDocument
    from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
    from pdfminer.converter import PDFPageAggregator
    from pdfminer.layout import LAParams, LTTextBox, LTTextLine
    extracted_text = ''

  # In file_path Provide the full file path including the pdf name for example C://UserName/Folder1/PdfFile.pdf
    file_content= open(file_path,'rb')
    parser = PDFParser(file_content)
    doc = PDFDocument()
    parser.set_document(doc)
    doc.set_parser(parser)
    doc.initialize('')
    rsrcmgr = PDFResourceManager()
    laparams = LAParams()
    #changing below 2 parameters to get rid of white spaces inside words
    laparams.char_margin = 1.0
    laparams.word_margin = 1.0
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    

    # Process each page contained in the pdf document.
    for page in doc.get_pages():
        interpreter.process_page(page)
        layout = device.get_result()
        for lt_obj in layout:
            if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
                extracted_text += lt_obj.get_text()


    return(extracted_text.encode("utf-8"))




import re
list1=[]
list2=[]
count=0
filePath=input("Enter the file Path without \ character at end of path\n")
pdfName=input("Enter the pdf file Name along without Extension :-\n")

count=0
editedText=""
extracted_text=pdf_to_text_file(filePath+"\\"+pdfName+".pdf")

print("PDF to Text Conversion Started............")
with open(filePath+"\\"+pdfName+'Extracted'+'.txt',"wb") as txt_file:
    txt_file.write(extracted_text)
extracted_text=extracted_text.decode("utf-8")
print("PDF to Text Converted")
list1=re.findall('[^ ]*@[^ ]*\S',extracted_text)

for x in list1:
    list2.append(x)

list4=[s.replace('\n', ' ') for s in list2]
for z in list4:
    editedText=editedText+" "+z

list3=re.findall('[^ ]*@[^ ]*\S',editedText)   

fhText = open(filePath+"\\"+pdfName+"_extractedEmailList.txt","w")
for z in list3:
    fhText. write(z+"\n")
fhText.close()

print("Email generated in txt format")
import xlwt 
workbook=xlwt.Workbook()
sheet1 = workbook.add_sheet('Sheet 1') 
for x in list3:
    sheet1.write(count,0,x)
    count=count+1
workbook.save(filePath+"\\"+pdfName+"_extractedEmailList"+".csv")   
print(" Leads Email generated in csv format")
print("File Generated Successfully at path :- "+filePath)


