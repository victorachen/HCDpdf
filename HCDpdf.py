from PyPDF2 import PdfFileReader, PdfFileWriter

reader = PdfFileReader(r'C:\Users\Lenovo\PycharmProjects\HCDpdf\app.pdf')
writer = PdfFileWriter()

page = reader.pages[0]
fields = reader.getFields()

writer.addPage(page)

# Now you add your data to the forms!
writer.updatePageFormFieldValues(
    writer.getPage(0), {"Manufacturer Trade Name:": "text"}
)

# write "output" to PyPDF2-output.pdf
with open(r'C:\Users\Lenovo\PycharmProjects\HCDpdf\output.pdf', "wb") as output_stream:
    writer.write(output_stream)