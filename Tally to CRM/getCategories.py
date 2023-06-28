from xlrd import open_workbook

path = "input.xls"
wb = open_workbook(path, formatting_info=True)
sheet = wb.sheet_by_name("Stock Summary")

# We'll keep track of the current path for each level
current_path = {0: ""}
with open("output.txt", "w") as f:
    # Go through each row
    for row_idx in range(sheet.nrows):
        cell_obj = sheet.cell(row_idx, 0)
        fmt = wb.xf_list[cell_obj.xf_index]

        # Compute indentation level from cell format
        indent_level = fmt.alignment.indent_level

        cell_value = cell_obj.value.strip()
        if not cell_value or cell_value.isspace():
            continue

        # Check if the second column is empty (indicating a category or subcategory)
        if sheet.cell(row_idx, 1).value == "":
            # Store the current path for this level
            current_path[indent_level] = cell_value

            # Construct the full path for this category or subcategory
            full_path = "/".join(
                current_path[level]
                for level in range(indent_level + 1)
                if level in current_path
            )

            # Print the full path
            f.write(full_path + "\n")
