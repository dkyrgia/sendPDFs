# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import openpyxl
import os

pdfList = []
file_obj_csv = "send_PDFs.csv"
file_obj_xlsx = "list_names.xlsx"

for name in os.listdir("pdfs\\"):
    if name.endswith(".pdf"):
        pdfList.append(name)

print(len(pdfList))
print(pdfList)

wb_obj = openpyxl.load_workbook(file_obj_xlsx)
sheet_obj = wb_obj.active
cell_obj = sheet_obj.cell(row=1, column=1)
print(cell_obj.value)
max_cols = sheet_obj.max_column
max_rows = sheet_obj.max_row
print(max_rows)
print(max_cols)

fileFromPdf = [0] * max_rows

for i in range(1, max_rows):
    # print(i)
    name_obj = sheet_obj.cell(row=i, column=2)
    mail_obj = sheet_obj.cell(row=i, column=3)
    fileFromPdf[i-1] = int(pdfList[i-1].split('.')[0])
    # fileFromPdf[i-1] = pdfList[i-1]
    # aaa = pdfList[i-1].split('-')[0]
    # print(aaa, i-1)
    # if int(aaa) == i:
    #     print(pdfList[i-1].split('-')[0], " ela", i)
    # print("#1")
    # print("fileFromPdf", fileFromPdf[i])
    # print("#2")
    # print("pdfList", pdfList[i].split('-', 1), i)
    # print(name_obj.value)
print(fileFromPdf)
fileFromPdf.sort()
pdfList.insert(0, 0)
for x in range(0, max_rows):

    print("fileFromPDF", fileFromPdf[x], " ", x, "name = ", pdfList[x])
# print(pdfList)



    # print(cell_obj.value)


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
