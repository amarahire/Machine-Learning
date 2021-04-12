from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askopenfilename,asksaveasfile
#from PyPDF2 import PdfFileReader
import pytesseract
import cv2

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

#=================open file method======================
def openFile(): 
              
    file = askopenfilename(defaultextension=".png", 
                                          filetypes=[("png files","*.png"),
                                                     ("jpg files","*.jpg")])
    if file == "":  
        file = None
    else:
        fileEntry.delete(0,END)
        fileEntry.config(fg="blue")
        fileEntry.insert(0,file)
def convert():
    try:
        img = fileEntry.get()
        # extracting text from page 
        extractedText=image = pytesseract.image_to_string(img)

        readImg.delete(1.0,END)
        readImg.insert(INSERT,extractedText)

        
    except FileNotFoundError:
        fileEntry.delete(0,END)
        fileEntry.config(fg="red")
        fileEntry.insert(0,"Please select a image file first")
    except:
        pass

    
def save2word():
    text = str(readImg.get(1.0,END))
    wordfile = asksaveasfile(mode='w',defaultextension=".doc", 
                                          filetypes=[("word file","*.doc"),
                                                     ("text file","*.txt"),
                                                     ("Python file","*.py"),
                                                     ("Excel file","*.xlsx"),
                                                     ("CSV file","*.csv")])
    
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
root.title("IMAGE Converter [designed by:AMAR AHIRE]")
root.resizable(0,0)
try:
    root.wm_iconbitmap("pdf2.ico")
except:
    print('icon file is not available')
    pass
file= ""
defaultText = "\n\n\n\n\t\t Your extracted text will apear here.\n \t\t     you can modify that text too."
#==============App Name==============================================================>>
appName = Label(root,text="FILE CONVERTER",font=('arial',20,'bold'),
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
                   bg="RED",fg='WHITE',width=20,command=convert)
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
