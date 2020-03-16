import xlrd
import xlwt

workbook= xlrd.open_workbook('DetailStatement.xlsx')
sheet1=workbook.sheet_by_index(0)

opworkbook= xlwt.Workbook()
opsheet = opworkbook.add_sheet('sheet 1')

print(sheet1.nrows)
print(sheet1.row_values(45))

escapes = ''.join([chr(char) for char in range(1, 32)])

for rownum in range (sheet1.nrows-1,2,-1):
    print("rownum "+str(rownum))
    shop=sheet1.row_values(rownum)[6].replace('\n','')
    date=sheet1.row_values(rownum)[1].replace('\n','')
    amt=str(sheet1.row_values(rownum)[5])
    if(shop.isnumeric()):
        shop='ATM'
    print(date+" : "+shop+" : "+ amt)
    print("________________________")
    opsheet.write(rownum,0,date)
    opsheet.write(rownum,1,shop)
    opsheet.write(rownum,2,amt)


opworkbook.save('HisaabOP.xls')