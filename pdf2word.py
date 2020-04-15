import win32com.client
import os


path = os.environ['USERPROFILE']+'\\'+'Desktop\\'

while True:
    pdf = input().strip('"')
    filename = pdf.split('\\')[-1]
    word = win32com.client.Dispatch("Word.Application")
    word.visible = 0
    wb = word.Documents.Open(pdf)
    docx = os.path.abspath(path+filename[0:-4]+'.docx')
    print(docx)
    wb.SaveAs2(docx, FileFormat=16)
    print("success...\n")
    wb.Close()

    #word.Quit()
