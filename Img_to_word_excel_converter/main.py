import os
def result():
    
    while(True):

        print('Please select your choices.')
        print('1) Convert image file to word.')
        print('2) Convert pdf file to word.')
        print('3) Convert word file to pdf.')
        print('4) Convert ppt file to pdf.')
        print('5) Exit.')
        inp = input('Enter your choice.')
        if(inp == '1'):
            os.system("python img_to_word.py")
            print("image file converter")
        elif(inp == '2'):
            os.system("python pdf_to_word.py")
            print("pdf file converted to word.")
        elif(inp == '3'):
            os.system("python word_to_pdf.py")
            print("word file converted to pdf.")
        elif(inp == '4'):
            os.system("python ppt_to_pdf.py")
            print("ppt file converted to pdf.")
        elif(inp == '5'):
            print("Exited")
            break
    
result()    
