import pandas as pd

# Reading the excel file using pandas
df = pd.read_excel("input.xlsx")

# Go through each row
for index, row in df.iterrows():
    # Check if the second column is nan (indicating a category or subcategory)
    if pd.isna(row[1]):
        # Count the leading spaces to determine the category level
        leading_spaces = len(row[0]) - len(row[0].lstrip(" "))

        # Print the category/subcategory with appropriate indentation
        print("  " * leading_spaces + row[0])
