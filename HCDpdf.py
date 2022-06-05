
#map all the park ID's, enter into the HCD 415 app field
#how many trailers can you have in your name? if so, put name in LLC name [ ask HCD on monday -- calendar]

from PyPDF2 import PdfFileReader, PdfFileWriter, PdfFileMerger
import openpyxl
from openpyxl import load_workbook

#takes inputs from excel input file, organizes it into dictionary d
def openpyxl():
    path = r'C:\Users\Lenovo\PycharmProjects\HCDpdf\input.xlsx'
    file = load_workbook(path)
    sheet = file.active

    d = {}
    for row in range(sheet.max_row):
        num = sheet.cell(row=row+1, column=2).value
        field = sheet.cell(row=row+1, column=3).value

        if type(num)==int  and field !=None:
            combined = str(num) + "_"+ str(field)
            raw_input = sheet.cell(row=row + 1, column=4).value
            d[combined] = raw_input

    #append a bunch of (hard coded) stuff to the back of the dictionary
    d = append_dic(d)
    # print(d)
    return d

#append (hard coding) stuff to the excel-made-dictionary (above)
#ex)from "9_date", we can deduce that the month is "April" and day is "15"; add this stuff to the back of the dictionary
def append_dic(d):
    todays_month = d['9_date'][0:4]
    todays_year = d['9_date'][-5:]
    todays_day = d['9_date'][4:6]
    yr_of_manufacture = d['2_Datefirstsold'][-5:]
    d['todays_month'] = todays_month
    d['todays_year'] = todays_year
    d['todays_day'] = todays_day
    d['yr_of_manufacture'] = yr_of_manufacture
    d['8_llc'] = d['8_llc']+', carrier-- Victor Chen'
    d['6_SitusAddress'] = d['6_SitusAddress'] + ', Yucaipa, CA 92399'
    d['7_Parkname'] = d['7_Parkname'] + ' Mobile Home Park'
    d['Elec_A'] = d['10_length'] * d['11_width']*3
    d['Elec_D'] = d['Elec_A'] + 1500
    d['Elec_E'] = min(3000, d['Elec_D'])
    d['Elec_F'] = round(max(d['Elec_E'] - 3000,0) * .35,0)
    d['Elec_G'] = d['Elec_E'] + d['Elec_F']
    d['Elec_H'] = round(d['Elec_G']/240,1)
    d['Elec_line12'] = max(30,d['Elec_H'])*.25
    d['Elec_line13'] = d['Elec_H'] + 30 + d['Elec_line12']
    d['parkID'] = parkID(d['7_Parkname'])
    return d

#given park name, return park ID
def parkID(parkname):
    parkIDs = {'Hitching Post':'36-0289-MP','Westwind':'36-0464-MP','Holiday':'36-0405-MP','Wishing Well':'36-0370-MP',
               'Mt Vista':'36-0330-MP','Crestview':'36-0595-MP','Patrician':'36-0484-MP'}
    for i in parkIDs:
        if i in parkname:
            return parkIDs[i]
        else:
            return 'N/A'

# helper function, that alters PDF (maps whatever is in dictionary)
def alterpdf(emptypath,filledpath):
    d = openpyxl()
    reader = PdfFileReader(emptypath)
    writer = PdfFileWriter()
    totalpages = reader.numPages
    for i in range(totalpages):
        page = reader.pages[i]
        fields = reader.getFields()
        writer.addPage(page)
        for x in d:
            writer.updatePageFormFieldValues(
                writer.getPage(i), {x: d[x]}
            )
        # sorts altered PDF into the "output" folder
        with open(filledpath, "wb") as output_stream:
            writer.write(output_stream)

#fill up them PDFs baby
def fill():
    L = ['dupcerttitle', 'dupregcard', "billofsale", 'multipurposetransfer', \
         'hcd415', 'electricalload', 'retailvalue', 'statementfacts']
    for i in L:
        emptypath = 'C:\\Users\\Lenovo\\PycharmProjects\\HCDpdf\\input\\'+ i +'_empty.pdf'
        filledpath = 'C:\\Users\\Lenovo\\PycharmProjects\\HCDpdf\\output\\' + i +'_filled.pdf'
        alterpdf(emptypath,filledpath)

#combine every file in the filled path
def combine():
    merger = PdfFileMerger()
    L = ['dupcerttitle', 'dupregcard', "billofsale", 'multipurposetransfer', \
         'hcd415', 'electricalload', 'retailvalue', 'statementfacts', 'AC_specs']
    for i in L:
        file = 'C:\\Users\\Lenovo\\PycharmProjects\\HCDpdf\\output\\' + i +'_filled.pdf'
        merger.append(PdfFileReader(open(file,'rb')))
    merger.write(r'C:\Users\Lenovo\PycharmProjects\HCDpdf\combined.pdf')
    return None

fill()
combine()