from operator import truediv
from turtle import title
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.utils import get_column_letter

from argparse import ArgumentParser
from bs4 import BeautifulSoup

#get args
parser = ArgumentParser()
parser.add_argument("-f", "--file", dest="inputfile", help="Input Nipper report XML")
parser.add_argument("-o", "--output", dest="outputfile", help="Output file")
args = parser.parse_args()

#open input with bs4
ifile = open(args.inputfile, "r", encoding="utf-8")
soup = BeautifulSoup(ifile.read(), features="xml")

#create output workbook
wb = openpyxl.Workbook()
mainws = wb.active
mainws.title = "Main"

#style stuff
titlegrey = Color(rgb="00CCCCCC")
titlefill = PatternFill(patternType="solid", fgColor=titlegrey)
headingfont = Font(bold=True, underline="single")

#fix column widths
def fix_column_width(ws):
    dim_holder = DimensionHolder(worksheet=ws)
    for col in range(ws.min_column, ws.max_column + 1):
        dim_holder[get_column_letter(col)] = ColumnDimension(ws, min=col, max=col, width=20)
    ws.column_dimensions = dim_holder

#function to get values from table reference
def get_table_values(reference, index):
    table = soup.find_all("table", {"ref":reference})[index]
    headings = table.find_all("heading")
    tablerows = table.find_all("tablerow")
    return table,headings,tablerows

#function to write tabledata to spreadsheet
def write_to_sheet(sheet, startingrow, startingcolumn, tabledata, title):
    rownum = startingrow
    sheet.cell(row=rownum, column=startingcolumn, value=title).font = headingfont
    rownum += 1

    for i,heading in enumerate(tabledata[1]):
        sheet.cell(row=rownum, column=i+startingcolumn, value=heading.text).fill = titlefill
    rownum +=1

    for i,row in enumerate(tabledata[2]):
        currow = rownum + i
        tablecells = row.find_all("tablecell")
        for j,cell in enumerate(tablecells):
            sheet.cell(row=currow, column=(j + startingcolumn), value=cell.text)

#create main sheet

#devices table
write_to_sheet(mainws, mainws.max_row , 1, get_table_values("SCOPE.AUDITDEVICELIST.TABLE", 0), "Devices")
#risk profile
write_to_sheet(mainws, mainws.max_row + 2 , 1, get_table_values("SECURITY.SUMMARY.SECURITYAUDIT.RISKPROFILE", 0), "Risk Profile")
#summary each device
write_to_sheet(mainws, mainws.max_row + 2 , 1, get_table_values("SECURITY.SUMMARY.AUDITDEVICELIST", 0), "Summary of findings for each device")
#vulnerability audit each device
if len(soup.find_all("table", {"ref":"VULN.SUMMARY.AUDITRESULTLIST"})) > 0:
    write_to_sheet(mainws, mainws.max_row + 2 , 1, get_table_values("VULN.SUMMARY.AUDITRESULTLIST", 0), "Summary of findings from the Vulnerability Audit for each device")
    vulnsheet = wb.create_sheet("Vulnerability Audit")
    write_to_sheet(vulnsheet, 1 , 1, get_table_values("VULNAUDIT.CONCLUSIONS", 0), "Vulnerability Conclusion")
    for x in range(1, len(soup.find_all("table", {"ref":"VULNAUDIT.CONCLUSIONS"}))):
        write_to_sheet(vulnsheet, vulnsheet.max_row + 2 , 1, get_table_values("VULNAUDIT.CONCLUSIONS", x), "Vulnerability List")
    fix_column_width(vulnsheet)


#create each issue sheet
issues = get_table_values("SECURITY.FINDINGS.SUMMARY.TABLE", 0)
for i,row in enumerate(issues[2]):

    items = row.find_all("item")

    title = items[1].text
    sheet = wb.create_sheet(title[:30])

    sheet.cell(row=1, column=1, value="Title").font = headingfont
    sheet.cell(row=2, column=1, value=title)

    section = soup.find("section", {"title": title})

    sheet.cell(row=4, column=1, value="Overall").font = headingfont
    sheet.cell(row=5, column=1, value=section.find("rating").text.strip())

    sheet.cell(row=7, column=1, value="Affected Device").font = headingfont
    sheet.cell(row=8, column=1, value=section.find("section", {"title": "Affected Device"}).find("listitem").text)

    #check if findings table init -=b
    findings = section.find_all("table")
    if len(findings) > 0:
        write_to_sheet(sheet, sheet.max_row + 2, 1, get_table_values(findings[0].get("ref"), 0), "Findings Table")
    
    sheet.cell(row=sheet.max_row + 2, column=1, value="Impact").font = headingfont
    sheet.cell(row=sheet.max_row + 1, column=1, value=section.find("section", {"ref":"IMPACT"}).text.strip())

    sheet.cell(row=sheet.max_row + 2, column=1, value="Ease").font = headingfont
    sheet.cell(row=sheet.max_row + 1, column=1, value=section.find("section", {"ref":"EASE"}).text.strip())

    sheet.cell(row=sheet.max_row + 2, column=1, value="Recommendation").font = headingfont
    sheet.cell(row=sheet.max_row + 1, column=1, value=section.find("section", {"ref":"RECOMMENDATION"}).text.strip())

    fix_column_width(sheet)


fix_column_width(mainws)

wb.save(args.outputfile)





