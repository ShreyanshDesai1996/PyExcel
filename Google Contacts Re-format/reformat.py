import xlrd
import xlwt

workbook= xlrd.open_workbook('C:/Users/Shrey/Desktop/PyExcel/Google Contacts Re-format/nnd.xlsx')
sheet1=workbook.sheet_by_index(0)

opworkbook= xlwt.Workbook()
opsheet = opworkbook.add_sheet('sheet 1')

print(sheet1.nrows)

ctr=0

def writeNewPlusRow(rn, row_values,number):
    global opsheet
    global ctr
    number='+'+number
    ctr=ctr+1
    #print(number)
    for col in range(len(row_values)):
        if(col == 20):
            opsheet.write(rn,col,number)
        else:
            opsheet.write(rn,col,row_values[col])

def writeSameRow(rn,row_values):
    global opsheet
    for col in range(len(row_values)):
        opsheet.write(rn,col,row_values[col])

for rownum in range (sheet1.nrows):
    number=str(sheet1.row_values(rownum)[20]).split('.')
    number=number[0].replace('Â','')

    if(len(number) > 0 and len(number)>10 and number[0]!='+'):
        writeNewPlusRow(rownum,sheet1.row_values(rownum),number)
    elif(len(number)==10):
        print(str(rownum+1))
    else:
        writeSameRow(rownum,sheet1.row_values(rownum))
    
print('Numbers changed: '+str(ctr))


opworkbook.save('NNDPH reformatted.xls')