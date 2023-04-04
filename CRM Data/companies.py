import pandas as pd

# Read the origin file
origin_df = pd.read_excel("slt.xlsx")

# Process the data
destination_df = pd.DataFrame(
    {
        "CompanyName": origin_df["company_name"],
        "AddressLine1": origin_df["address_1"],
        "AddressLine2": origin_df["address_2"],
        "City/Town": origin_df["city"],
        "County/State": origin_df["state"],
        "Pincode": origin_df["pincode"],
        "Country": origin_df["country"],
        "Email": origin_df["email"],
        "Phone": origin_df[
            ["phone_no_1", "phone_no_2", "phone_no_3", "phone_no_4", "phone_no_5"]
        ].apply(lambda x: ",".join(x.dropna().astype(str)), axis=1),
        "CompanyCategory": origin_df["contact_type"],
    }
)

# Write the result to the destination file
destination_df.to_excel("companies.xlsx", index=False)
