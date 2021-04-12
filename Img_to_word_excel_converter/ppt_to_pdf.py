from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askopenfilename,asksaveasfile
#from PyPDF2 import PdfFileReader
import pytesseract
import cv2
from docx2pdf import  convert
import comtypes.client
from pyfiglet import Figlet
from ppt2pdf.utils import generateOutputFilename

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

#=================open file method======================
def openFile(): 
              
    file = askopenfilename(defaultextension=".pptx", 
                                          filetypes=[("ppt file","*.pptx")])
    if file == "":  
        file = None
    else:
        fileEntry.delete(0,END)
        fileEntry.config(fg="blue")
        fileEntry.insert(0,file)
def convertfile():
    try:
        file = fileEntry.get()
        # extracting text from page 
    
        # print("Your Input file is at:")
        # print(inputFilePath)
        outputFilePath = r"F:\projects\Img_to_word_excel_converter\ppt_to_pdf.pdf"
        if(not outputFilePath):
            outputFilePath = generateOutputFilename(file);
        
        # print("Your Output file will be at:")
        # print(outputFilePath);
        # %% Create powerpoint application object
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        #%% Set visibility to minimize
        powerpoint.Visible = 1
        #%% Open the powerpoint slides
        slides = powerpoint.Presentations.Open(file)
        #%% Save as PDF (formatType = 32)
        slides.SaveAs(outputFilePath, 32)
        #%% Close the slide deck
        slides.Close()

    
        readImg.delete(1.0,END)
        readImg.insert(INSERT,extractedText)

        
    except FileNotFoundError:
        fileEntry.delete(0,END)
        fileEntry.config(fg="red")
        fileEntry.insert(0,"Please select a pptx file first")
    except:
        pass

    
def save2word():
    text = str(readImg.get(1.0,END))
    wordfile = asksaveasfile(mode='w',defaultextension=".pdf", 
                                          filetypes=[("pdf file","*.pdf")])
                                                     
    
    if wordfile is None:
        return
    wordfile.write(text)
    wordfile.close()
    print("saved")
    fileEntry.delete(0,END)
    fileEntry.insert(0,"File Extracted and Saved...")



    
#=================== Front End Design ==================================================
root = Tk()
root.geometry("600x350")
root.config(bg="light blue")
root.title("IMAGE Converter [designed by:CHETAN SHARMA]")
root.resizable(0,0)
try:
    root.wm_iconbitmap("pdf2.ico")
except:
    print('icon file is not available')
    pass
file= ""
defaultText = "\n\n\n\n\t\t Your extracted text will apear here.\n \t\t     you can modify that text too."
#==============App Name==============================================================>>
appName = Label(root,text="PPT To PDF Converter",font=('arial',20,'bold'),
                bg="light blue",fg='maroon',anchor="center")
appName.place(x=150,y=5)
#Select pdf file
labelFile = Label(root,text="Select File",font=('arial',12,'bold'))
labelFile.place(x=30,y=50)
fileEntry = Entry(root,font=('calibri',12),width=40)
fileEntry.pack(ipadx=200,pady=50,padx=150)
#===========button to access openFile method=================================
openFileButton = Button(root,text=" Open ",font=('arial',12,'bold'),width=30,
                    bg="sky blue",fg='green',command=openFile)
openFileButton.place(x=150,y=80)
#===========button to access convert method=================================
convert2Text = Button(root,text="Read",font=('arial',8,'bold'),
                   bg="RED",fg='WHITE',width=20,command=convertfile)
convert2Text.place(x=250,y=120)
#======================= Text Box to read Img file and modify ===================>>
readImg = Text(root,font=('calibri',12),fg='light green',bg='black',width=60,height=30,bd=10)
readImg.pack(padx=20,ipadx=20,pady=20,ipady=20)
readImg.insert(INSERT,defaultText)
#===============================Button to access save2word method=================
save2Word = Button(root,text="Save",font=('arial',10,'bold'),
                   bg="RED",fg='WHITE',command=save2word)
save2Word.place(x=255,y=320)

#===================halt window=============================>>
if __name__ == "__main__":
    root.mainloop()

