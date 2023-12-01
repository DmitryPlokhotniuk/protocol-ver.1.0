import openpyxl
from win32com import client
from PyPDF2 import PdfWriter

merger = PdfWriter()
# path = 'C:\\my_program\\my_program\\whole_file.pdf'
book = openpyxl.open('data_owners.xlsx')
sheet = book.active
new_row = 1
sheets_name = ['Лист1', 'Лист2', 'Лист3']
number_name = 1
for row in range(10, 12):
    new_book = openpyxl.load_workbook('reshenie.xlsx')

    number_room = str(sheet[row][0].value)
    fio = str(sheet[row][1].value)
    room_area = float(sheet[row][2].value)
    part_own_str = str(sheet[row][4].value)
    part_own = sheet[row][4].value
    if type(part_own) == str:
        a, b = part_own.split("/")
        part_own = int(a) / int(b)
    document_number = str(sheet[row][5].value)
    for i in sheets_name:
        new_sheet = new_book[i]
        new_sheet['C6'].value = number_room
        new_sheet['A4'].value = fio
        new_sheet['C9'].value = room_area
        new_sheet['C10'].value = part_own_str
        new_sheet['C11'].value = part_own * room_area
        new_sheet['A7'].value = document_number

    new_book.save(f'tables\\Protocol{new_row}.xlsx')
    new_book.close()

    excel = client.Dispatch("Excel.Application")

    sheets = excel.Workbooks.Open(f'C:\\my_program\\my_program\\tables\\Protocol{new_row}.xlsx')

    work_sheets = sheets.Worksheets[0]
    work_sheets_1 = sheets.Worksheets[1]
    work_sheets_2 = sheets.Worksheets[2]

    work_sheets.ExportAsFixedFormat(0, f"C:\\my_program\\my_program\\pdf\\mypdf{number_name}.pdf")
    merger.append(f"C:\\my_program\\my_program\\pdf\\mypdf{number_name}.pdf")
    number_name += 1
    work_sheets_1.ExportAsFixedFormat(0, f"C:\\my_program\\my_program\\pdf\\mypdf{number_name}.pdf")
    merger.append(f"C:\\my_program\\my_program\\pdf\\mypdf{number_name}.pdf")
    number_name += 1
    work_sheets_2.ExportAsFixedFormat(0, f"C:\\my_program\\my_program\\pdf\\mypdf{number_name}.pdf")
    merger.append(f"C:\\my_program\\my_program\\pdf\\mypdf{number_name}.pdf")
    number_name += 1

    new_row += 1

merger.write('merged_pdf.pdf')
merger.close()