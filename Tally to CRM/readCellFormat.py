from xlrd import open_workbook

path = "input.xls"
wb = open_workbook(path, formatting_info=True)
sheet = wb.sheet_by_name("Stock Summary")
cell = sheet.cell(20, 0)  # The first cell
print(cell.value)
print("cell.xf_index is", cell.xf_index)
fmt = wb.xf_list[cell.xf_index]
print("type(fmt) is", type(fmt))
print("Dumped Info:")
fmt.dump()
