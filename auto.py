
from PIL import Image, ImageDraw, ImageFont
import os
import shutil
import xlrd

def x():
    
    wb = xlrd.open_workbook('test3.xlsx')
    sheet = wb.sheet_by_index(0)
       
    for i in range(sheet.ncols):
        if sheet.cell_value(0, i) == 'Name' or sheet.cell_value(0, i) == 'name' or sheet.cell_value(0, i) == 'Student name' or sheet.cell_value(0, i) == 'student name':
            ni = i
            
    names = []
     
    for i in range(sheet.nrows):
        names.append(sheet.cell_value(i, ni))
        
    del names[0]
        
    if os.path.isdir("generated"): 
        shutil.rmtree("generated")
        os.mkdir("generated")
    else:
        os.mkdir("generated")
       
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

x()