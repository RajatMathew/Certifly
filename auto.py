# "mail" in text
from PIL import Image, ImageDraw, ImageFont
import os
import shutil
import xlrd


def folder_check():
    if os.path.isdir("generated"): 
        shutil.rmtree("generated")
        os.mkdir("generated")
    else:
        os.mkdir("generated")


def generate():
    
    wb = xlrd.open_workbook('test3.xlsx')
    sheet = wb.sheet_by_index(0)

    for i in range(sheet.ncols):
        if "NAME" in sheet.cell_value(0, i).upper():
            ni = i


    names = []
     
    for i in range(sheet.nrows):
        names.append(sheet.cell_value(i, ni))
        
    del names[0]
        
    folder_check()
       
    font = ImageFont.truetype('fonts/Poppins-Medium.ttf', 40)

    for name in names:  
        print("Generating " + name + ".png")
        
        img = Image.open("templates/ctf.png")
        draw = ImageDraw.Draw(img)
            
        current_name = name
        w, h = draw.textsize(current_name, font=font)
            
        draw.text(((1280-w)/2, (((82-h)/2)+434)), current_name, font=font, fill="black")
                                                                                                                                  
        fname = current_name + ".png"
        img.save(r'generated/' + fname)

generate()