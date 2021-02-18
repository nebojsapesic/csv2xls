import xlwt
import csv

style = xlwt.easyxf('font: bold on')

index_list = []
with open("qoc_wcsa1.txt") as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=';')
    for index, row in enumerate(csv_reader):
            if '===' in row[0]:
                index_list.append(index)
print(index_list)
with open("qoc_wcsa1.txt") as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=';')
    for i in range(0, 5):
        next(csv_reader)
    xls_file = xlwt.Workbook()  # new xls file
    try:
        for index, row in enumerate(csv_reader):
            if index == 0:
                sheet1 = xls_file.add_sheet('Acquiring')
                for j in range(len(row)):
                    sheet1.write(index, j, row[j], style)
            elif 0 < index < index_list[1]-5:
                for j in range(len(row)):
                    sheet1.write(index, j, row[j])
            elif index == index_list[1]-4:
                sheet2 = xls_file.add_sheet('Acquiring_Acceptance Loc')
                for j in range(len(row)):
                    sheet2.write(index-index_list[1]+4, j, row[j], style)
            elif index_list[1]-4 < index < index_list[2]-5:
                for j in range(len(row)):
                    sheet2.write(index-index_list[1]+4, j, row[j])
            elif index == index_list[2]-4:
                sheet3 = xls_file.add_sheet('Acquiring OCT_BAI Reporting')
                for j in range(len(row)):
                    sheet3.write(index-index_list[2]+4, j, row[j], style)
            elif index > index_list[2]-4:
                for j in range(len(row)):
                    sheet3.write(index-index_list[2]+4, j, row[j])
            else:
                continue
            
        
        xls_file.save('output1.xls')
    
    except Exception as e:
        print(e)
