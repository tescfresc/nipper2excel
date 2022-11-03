from openpyxl.styles import Color, PatternFill, Font


global soup
global titlegrey
titlegrey = Color(rgb="00CCCCCC")
global titlefill
titlefill = PatternFill(patternType="solid", fgColor=titlegrey)
global headingfont
headingfont = Font(bold=True, underline="single")
