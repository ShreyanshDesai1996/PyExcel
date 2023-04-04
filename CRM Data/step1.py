import pandas as pd

# Read data from origin excel file
origin_df = pd.read_excel("slt.xlsx")

# Transform data
destination_df = pd.DataFrame()
destination_df["CompanyName"] = origin_df["company_name"]
destination_df["Title"] = (
    origin_df["contact__person"]
    .str.extract(r"(\b(Mr|Ms|Mrs|Dr|Prof|Rev|Hon)\.\s)?(\b\w*\b)", expand=False)
    .iloc[:, 1]
    .fillna("")
)

# Extract first, middle, and last names if available
name_pattern = (
    r"(\b(Mr|Ms|Mrs|Dr|Prof|Rev|Hon)\.\s)?(\b\w*\b)\s?(\b\w*\b)?\s?(\b\w*\b)?"
)
names = origin_df["contact__person"].str.extract(name_pattern, expand=False)
destination_df["FirstName"] = ""
destination_df["MiddleName"] = ""
destination_df["LastName"] = ""

for i, row in names.iterrows():
    if pd.notnull(row[3]):
        # If all three names are present
        destination_df.loc[i, "FirstName"] = row[2]
        destination_df.loc[i, "MiddleName"] = row[3]
        destination_df.loc[i, "LastName"] = row[4]
    elif pd.notnull(row[2]) and pd.notnull(row[4]):
        # If only first and last names are present
        destination_df.loc[i, "FirstName"] = row[2]
        destination_df.loc[i, "LastName"] = row[4]
    else:
        # If the name is not in expected format, place the entire name in first name
        destination_df.loc[i, "FirstName"] = origin_df.loc[i, "contact__person"]

destination_df["Code"] = ""
destination_df["JobTitle"] = origin_df["designation"]
destination_df["AddressLine1"] = origin_df["address_1"]
destination_df["AddressLine2"] = origin_df["address_2"]
destination_df["City/Town"] = origin_df["city"]
destination_df["County/State"] = origin_df["state"]
destination_df["Pincode"] = origin_df["pincode"]
destination_df["Country"] = origin_df["country"]
destination_df["Email"] = origin_df["email"]
destination_df["Phone"] = origin_df.iloc[:, 8:13].apply(
    lambda x: ",".join(x.dropna().astype(str)), axis=1
)
destination_df["ContactCategory"] = ""
destination_df["ContactType"] = origin_df["contact_type"]

# Write data to destination excel file
destination_df.to_excel("destination_file.xlsx", index=False)
