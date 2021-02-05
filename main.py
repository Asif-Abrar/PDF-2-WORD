import win32com.client
import os

word = win32com.client.Dispatch("word.Application")
word.visible = 0

doc_pdf = input("Enter the path: ")
input_file = os.path.abspath(doc_pdf)
wb = word.Documents.Open(input_file)
output_file = os.path.abspath(doc_pdf[0:-4].format())
wb.SaveAs2(output_file, FileFormat=16)
print("PDF to Docx is completed")
wb.Close()

word.Quit()
