import email
from itertools import count
from PIL import Image, ImageDraw, ImageFont
import os, shutil, xlrd


import smtplib
import mimetypes
from email.message import EmailMessage

# Global Variables
spreadsheet_file = 'test.xlsx'
font = ImageFont.truetype('fonts/Poppins-Medium.ttf', 60) # Setting the font to Poppins Medium and font size to 40

names = []
emails = []

# Creates a new 'generated' folder if not already present
def folder_check():
    if os.path.isdir("generated_certificates"): 
        shutil.rmtree("generated_certificates")
        os.mkdir("generated_certificates")
    else:
        os.mkdir("generated_certificates")


def findNameCol(sheet):
    for i in range(sheet.ncols):
        if "NAME" in sheet.cell_value(0, i).upper(): # Programmatically finds column heading 'name' in the spreadsheet, without explicitly asking the user
            name_col = i
    
    return name_col


def findEmailCol(sheet):
    for j in range(sheet.ncols):
        if "EMAIL" in sheet.cell_value(0, j).upper(): # Programmatically finds column heading 'name' in the spreadsheet, without explicitly asking the user
            email_col = j
    
    return email_col



# Gets names from the xlsx file
def getNames():
    
    workbook = xlrd.open_workbook(spreadsheet_file)
    sheet = workbook.sheet_by_index(0) # Getting first sheet of the xlsx workbook

    for i in range(sheet.nrows):
        names.append(sheet.cell_value(i, findNameCol(sheet))) # Extracts names from the column and inserts it into the array
        
    del names[0] # Deleting the column heading

    return names

def getEmails():
    
    workbook = xlrd.open_workbook(spreadsheet_file)
    sheet = workbook.sheet_by_index(0) # Getting first sheet of the xlsx workbook

    for j in range(sheet.nrows):
        emails.append(sheet.cell_value(j, findEmailCol(sheet))) # Extracts names from the column and inserts it into the array
        
    del emails[0] # Deleting the column heading

    return emails



def generate(names,emails):



    def sendEmail():
            
        for counter in range(len(emails)):
            name = names[counter]
            email = emails[counter]

            message = EmailMessage()
            sender = "mail@gmail.com"
            recipient = email
            message['From'] = sender
            message['To'] = recipient
            message['Subject'] = name + ", your certificate is here"
            body = "Hi " + name + ",\n\nSorry there was a clerical mistake in the certificate attached with the last mail."+"""\nContent
Please find the file attached with this email.
Do revert back to this mail incase of any errors/trouble.
            
Have a Good Day!
            
Regards,
Amal Shibu
CEO, IEDC JEC"""
            message.set_content(body)
            fname=r'generated_certificates/'+name + ".png"
            # print(fname)
            mime_type, _ = mimetypes.guess_type(fname)
            mime_type, mime_subtype = mime_type.split('/')
            with open(fname, 'rb') as file:
                message.add_attachment(file.read(),
                maintype=mime_type,
                subtype=mime_subtype,
                filename=name+".png")

            mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
            mail_server.set_debuglevel(0)
            mail_server.login("mailid", 'password')
            mail_server.send_message(message)


            print("Recipient Name : "+name)
            print("Recipient Mail Id : "+recipient)
            print("Msg :\n"+body)


            print("\n\n")
            print("-------------------------------------")
            print("\n\n")


            mail_server.quit()   

    folder_check()

    def gencertificate():
        for name in names:  
            print("Generating " + name + ".png")
            
            img = Image.open("templates/ctf.png") # Loading the certificate template

            draw = ImageDraw.Draw(img)
            w, h = draw.textsize(name, font=font)
                
            name_x, name_y = (1990-w)/2, (((350-h)/2)+430) # Setting the co-ordinates to where the names should be entered
            draw.text((name_x, name_y), name, font=font, fill="black")
                                                                                                                                    
            img.save(r'generated_certificates/' + name + ".png") # Saving the images to the generated_certificates directory

            print("\n\n")
    
    gencertificate() #Generate and save certiciates according to the name 
    sendEmail() #Sends email to the given mail id

if __name__ == "__main__":
    
    generate(getNames(),getEmails())