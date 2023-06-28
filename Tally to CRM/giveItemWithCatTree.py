from xlrd import open_workbook
import xlrd
import openpyxl


def read_excel(file_path):
    # Load workbook
    wb = xlrd.open_workbook(file_path, formatting_info=True)
    sheet = wb.sheet_by_name("Stock Summary")

    items = []
    categories = {0: ""}

    for i in range(sheet.nrows):
        col1_value = sheet.cell_value(i, 0)
        col2_value = sheet.cell_value(i, 1)

        cell_obj = sheet.cell(i, 0)
        fmt = wb.xf_list[cell_obj.xf_index]

        # Compute indentation level from cell format
        indent_level = fmt.alignment.indent_level

        value = col1_value.strip()

        if col2_value == "":
            # It's a category or subcategory
            categories[indent_level] = value
        else:
            try:
                item_category = "/".join(
                    [categories[j] for j in range(indent_level + 1) if j in categories]
                )
            except KeyError:
                print(
                    f"Error: Could not find parent category for item '{value}' at row {i+1}."
                )
                continue  # Skip this item
            items.append([value, item_category])

    return items


def write_to_excel(items, output_file_path):
    wb = openpyxl.Workbook()
    sheet = wb.active

    # Write header
    sheet.append(["Item Name", "Parent Categories"])

    # Write items and their categories
    for item in items:
        sheet.append(item)

    # Save the output file
    wb.save(output_file_path)


# Use the functions
input_file_path = "input.xls"  # replace with your input file path
output_file_path = "output.xls"  # replace with your output file path

items = read_excel(input_file_path)
write_to_excel(items, output_file_path)
