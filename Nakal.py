from logging import root
import os
import cv2
import pytesseract
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
import tkinter as tk
import docx2txt
from pathlib import Path

# Mention the installed location of Tesseract-OCR in your system
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


Path("images").mkdir(parents=True, exist_ok=True)
Path("output").mkdir(parents=True, exist_ok=True)

#getdocument 
for x in os.listdir():
    if x.endswith(".docx"):
        docName = x

#get images
text = docx2txt.process(docName, 'images')

for x in os.listdir("images"):
    print(x)
    # Read image from which text needs to be extracted
    original_img = cv2.imread("images\\"+x)
    new_height, new_width = original_img.shape[:2]
    # name of the new file to save
    filename = "output\\"+x
    # create new image
    new_image = Image.new(mode = "RGB", size=(new_width,new_height), color = "black")
    # save the file
    new_image.save(filename)
    # Convert the image to gray scale
    gray = cv2.cvtColor(original_img, cv2.COLOR_BGR2GRAY)
    new_img = Image.open(filename)
    # Call draw Method to add 2D graphics in an image
    I1 = ImageDraw.Draw(new_img)
    myFont = ImageFont.truetype('CONSOLA.ttf', 17, encoding="unic")
    #extracting text
    text = pytesseract.image_to_string(gray)
    print(text)
    #writing text on image
    I1.text((10, 10), text, font=myFont, fill =(204, 204, 204))
    new_img.save(filename)


root = tk.Tk()
root.title("Nakal")
root.geometry("800x500")
title_label = tk.Label(root,text="Nakal",font=("Arial Bold", 50)).pack()
replacement_frame = tk.Frame(root)
original_name_var = tk.StringVar()
new_name_var = tk.StringVar()
original_name_label = tk.Label(replacement_frame,text="Find:",font=('calibre',10,'bold')).grid(row=0,column=0)
original_name_input = tk.Entry(replacement_frame,textvariable=original_name_var,font=('calibre',10,'normal')).grid(row=0,column=1)
new_name_label = tk.Label(replacement_frame,text="Replace:",font=('calibre',10,'bold')).grid(row=1,column=0)
new_name_input = tk.Entry(replacement_frame,textvariable=new_name_var,font=('calibre',10,'normal')).grid(row=1,column=1)
replacement_frame.pack()

root.mainloop()
