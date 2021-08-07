from glob import glob
from PyPDF2 import PdfFileMerger
import docx2pdf
import win32com.client
import os,sys,time
from webutil import clear
def Choice():
    clear.clear()
    try:
        activity=int(input(f'''      **** FILE MERGER ****
                           \n1 For merging PDF\n2 For docs\wordfiles to Pdf\n3 For PPT to Pdf\n4 For Exiting\n'''))
    except:
        print("Enter A Valid Number")
        time.sleep(3)
        Choice()
    if activity ==4:
        clear.clear()
        print("Exiting")
        time.sleep(3)
        sys.exit()
    else:
        clear.clear()
        address=input(f"Enter directory address\n")
        while not os.path.isdir(address):
            clear.clear()
            print("Enter A Valid Path")
            address=input()
        
        clear.clear()
        
        if activity ==1:
            pdfMerger(address)
        elif activity ==2:
            Word(address)
        elif activity ==3:
            PPT(address)
        else:
            print(f"\nEnter a number between 1-4")
            time.sleep(3)
        Choice()
    
def pdfMerger(address_of_directory):
    os.chdir(rf"{address_of_directory}")
    name_of_file=input(f"Enter the name of file\n")
    files=[file for file in os.listdir() if file.endswith(".pdf")]
    merger=PdfFileMerger()
    for pdf in files:
        merger.append(pdf)

    merger.write(f"{name_of_file}.pdf")
    merger.close()

    
def Word(address):
    path=os.path.join(address,"wordpdf")
    os.mkdir(path)
    docx2pdf.convert(address,path)
def PPT(address):
    os.chdir(address)
    path=os.path.join(address,"PPTpdf")
    os.mkdir(path)
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    for file in glob("*.pptx"):
        print(file)
        newname = os.path.splitext(file)[0] + ".pdf"
        deck = powerpoint.Presentations.Open(rf"{address}\{file}")
        deck.Saveas(rf"{path}\{newname}",32)
        deck.Close()
    for file in glob("*.ppt"):
        print(file)
        newname = os.path.splitext(file)[0] + ".pdf"
        deck = powerpoint.Presentations.Open(rf"{address}\{file}")
        deck.Saveas(rf"{path}\{newname}",32)
        deck.Close()
    powerpoint.Quit()

if __name__ == "__main__":
    Choice()