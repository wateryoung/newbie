# Install the openpyxl library
from openpyxl import load_workbook
from yattag import Doc, indent

# Loading our Excel file
wb = load_workbook("G:\workpy\data\접속표준서(정보분배-UDP_TCP실시간)_v1.02.xlsx")

# Getting an object of active sheet 4th
ws = wb.worksheets[3]

# Returning returns a triplet
doc, tag, text = Doc().tagtext()
  
xml_header = '<?xml version="1.0" encoding="UTF-8"?>'
xml_schema = '<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema"></xs:schema>'
  
# Appends the String to document
doc.asis(xml_header)
doc.asis(xml_schema)
  
with tag('InterfaceDefinition'):
    for row in ws.iter_rows(min_row=3, max_row=3977, min_col=1, max_col=13):
        row = [cell.value for cell in row]
        with tag("Row"):
            with tag("Interface_ID"):
                text(row[0])
            with tag("Inferface_Name"):
                text(row[1])
            with tag("Sequence"):
                text(row[3])
            with tag("ItemName"):
                text(row[4])
            with tag("Depth"):
                text(row[6])
            with tag("DataType"):
                text(row[7])
            with tag("Length"):
                text(row[8])  
            with tag("AccLength"):
                text(row[9])
#            with tag("Definition"):
#                text(row[10])
#            with tag("Remarks"):
#                text(row[12])

result = indent(
    doc.getvalue(),
    indentation='   ',
    indent_text=True
)
  
with open("output.xml", "w") as f:
    f.write(result)