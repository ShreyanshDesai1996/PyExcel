import pandas as pd

# read data from File1
df1 = pd.read_excel("slt.xlsx")

# remove titles from contact_person column
df1["contact__person"] = (
    df1["contact__person"].str.replace("Mr. ", "").str.replace("Ms. ", "")
)
# read data from File2
df2 = pd.read_excel("destination_file.xlsx")

# find rows where FirstName has only one character
mask = df2["FirstName"].str.len() == 1

# replace FirstName column with contact__person column for these rows
df2.loc[mask, "FirstName"] = df1.loc[mask, "contact__person"]

# save updated data to File2
df2.to_excel("destination_updated.xlsx", index=False)
