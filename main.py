#imports
import gspread
from fpdf import FPDF

from oauth2client.service_account import ServiceAccountCredentials
from tkinter import *


root = Tk()  # create parent window
root.geometry("200x100")
def button():

    # Connect to Google Sheets
    scope = ['https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name("Credentials.json", scope)
    client = gspread.authorize(credentials)
    gc = gspread.service_account()
    sh = gc.open('RESPONSES')

    # PDF
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    worksheet = sh.get_worksheet(0)
    cell = worksheet.find("")

    # loop to generate a PDF.
    val = worksheet.acell('A2').value
    celda = 2

    while val != None:

        try:

            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Times", size=20, style="B" + "U")
            pdf.cell(200, 10, txt=worksheet.acell('B' + str(celda)).value+"'s Resume", ln=1, align="C")

            #personal information
            pdf.set_font("Arial", size=14, style="B")
            pdf.cell(100, 10, txt="Personal information:", ln=1, align="L")
            pdf.set_font("Arial", size=12)
            pdf.cell(100, 10, ln=1, txt="Name: " + worksheet.acell('B' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt="Phone number: " + worksheet.acell('C' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt="Email: " + worksheet.acell('D' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt="Nationality: " + worksheet.acell('E' + str(celda)).value, align="L")

            # academic
            pdf.set_font("Arial", size=14, style="B")
            pdf.cell(100, 10, txt="Academic background", ln=1, align="L")
            pdf.set_font("Arial", size=12)
            pdf.cell(100, 10, ln=1, txt="School name: " + worksheet.acell('F' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt="SAT scores: " + worksheet.acell('G' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt="Language proficiency score: " + worksheet.acell('H' + str(celda)).value, align="L")

            # extracurricular
            pdf.set_font("Arial", size=14, style="B")
            pdf.cell(100, 10, txt="Extracurricular activities:", ln=1, align="L")
            pdf.set_font("Arial", size=12)
            pdf.cell(100, 10, ln=1, txt= worksheet.acell('I' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt= worksheet.acell('J' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt= worksheet.acell('K' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt= worksheet.acell('L' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt= worksheet.acell('M' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt= worksheet.acell('N' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt= worksheet.acell('O' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt= worksheet.acell('P' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt= worksheet.acell('Q' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt= worksheet.acell('R' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt= worksheet.acell('S' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt= worksheet.acell('T' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt= worksheet.acell('U' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt= worksheet.acell('V' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt= worksheet.acell('W' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt= worksheet.acell('X' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt= worksheet.acell('Y' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt= worksheet.acell('Z' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt= worksheet.acell('AA' + str(celda)).value, align="L")
            pdf.cell(100, 10, ln=1, txt= worksheet.acell('AB' + str(celda)).value, align="L")

        except:
            1 + 1

        print("PDF created: "+val)
        pdf.output(worksheet.acell('B' + str(celda)).value + ".pdf")
        val = worksheet.acell('A' + str(celda)).value
        celda = celda + 1


    print("PDF generated")

# Create volume up button
CreateResume = Button(root, text="Create Resume!",command=button)
CreateResume.pack()
root.mainloop()
