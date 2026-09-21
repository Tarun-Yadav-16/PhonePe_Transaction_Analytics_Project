import pandas as pd


# Set the input and output file names.
input_file = "PhonePe_Messy__Dataset.xlsx"
output_file = "PhonePe-Final-dataset.xlsx"


# Load both worksheets from the messy workbook.
users = pd.read_excel(input_file, sheet_name="All_Users")
transactions = pd.read_excel(input_file, sheet_name="All_Transactions")


# Standardize the column names.
users.columns = users.columns.astype(str).str.strip()
transactions.columns = transactions.columns.astype(str).str.strip()

# Rename the user-key column if the source workbook calls it "All".
# This makes it match the User_ID column in All_Transactions.
if "All" in users.columns and "User_ID" not in users.columns:
    users.rename(columns={"All": "User_ID"}, inplace=True)


# -------------------------
# Clean the All_Users worksheet
# -------------------------

# Remove extra spaces from user IDs and names.
for column in ["User_ID", "Name"]:
    users[column] = (
        users[column]
        .astype("string")
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )

# Standardize the capitalization of names.
users["Name"] = users["Name"].str.title()

# Replace missing names with a clear placeholder.
users["Name"] = users["Name"].fillna("Unknown")

# Convert age values such as " 46 " and "N/A" into numeric values.
users["Age"] = pd.to_numeric(users["Age"], errors="coerce")

# Mark ages outside the valid range as missing.
users.loc[~users["Age"].between(18, 100), "Age"] = pd.NA

# Fill missing ages with the median age.
users["Age"] = users["Age"].fillna(users["Age"].median()).round().astype("Int64")

# Convert mixed join-date formats into one datetime column.
try:
    users["Join_Date"] = pd.to_datetime(
        users["Join_Date"], errors="coerce", dayfirst=True, format="mixed"
    )
except (TypeError, ValueError):
    users["Join_Date"] = pd.to_datetime(
        users["Join_Date"], errors="coerce", dayfirst=True
    )


# -------------------------
# Clean the All_Transactions worksheet
# -------------------------

# List the transaction columns that contain text.
text_columns = [
    "Transaction_ID",
    "User_ID",
    "Service",
    "Service Type",
    "Payment_Status",
    "Reason",
]

for column in text_columns:
    transactions[column] = (
        transactions[column]
        .astype("string")
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )

# Standardize the capitalization of transaction text values.
transactions["Service"] = transactions["Service"].str.title()
transactions["Service Type"] = transactions["Service Type"].str.title()
transactions["Payment_Status"] = transactions["Payment_Status"].str.title()
transactions["Reason"] = transactions["Reason"].str.title()

# Restore common business abbreviations after standardizing case.
transactions["Service Type"] = transactions["Service Type"].replace(
    {
        "Dth": "DTH",
        "Fastag Recharge": "FASTag Recharge",
        "Cable Tv": "Cable TV",
    }
)

# Remove currency symbols and commas from Amount values.
amount_text = (
    transactions["Amount"]
    .astype("string")
    .str.replace("₹", "", regex=False)
    .str.replace(",", "", regex=False)
    .str.strip()
)

# Convert the cleaned Amount values to numbers.
transactions["Amount"] = pd.to_numeric(amount_text, errors="coerce")

# Mark negative amounts as missing because they are invalid transactions.
transactions.loc[transactions["Amount"] < 0, "Amount"] = pd.NA

# Fill missing amounts with the median amount.
transactions["Amount"] = transactions["Amount"].fillna(transactions["Amount"].median())

# Convert mixed transaction-date formats into one datetime column.
try:
    transactions["Date"] = pd.to_datetime(
        transactions["Date"], errors="coerce", dayfirst=True, format="mixed"
    )
except (TypeError, ValueError):
    transactions["Date"] = pd.to_datetime(
        transactions["Date"], errors="coerce", dayfirst=True
    )


# Remove duplicate user and transaction business keys.
users.drop_duplicates(subset=["User_ID"], keep="first", inplace=True)
transactions.drop_duplicates(subset=["Transaction_ID"], keep="first", inplace=True)

# Remove rows with missing primary keys.
users.dropna(subset=["User_ID"], inplace=True)
transactions.dropna(subset=["Transaction_ID"], inplace=True)

# Find transaction user IDs that do not exist in the user master table.
valid_user_ids = set(users["User_ID"].dropna())
transactions.loc[
    ~transactions["User_ID"].isin(valid_user_ids), "User_ID"
] = pd.NA


# Write both cleaned worksheets to a new Excel workbook.
with pd.ExcelWriter(
    output_file,
    engine="openpyxl",
    date_format="yyyy-mm-dd",
    datetime_format="yyyy-mm-dd",
) as writer:
    users.to_excel(writer, sheet_name="All_Users", index=False)
    transactions.to_excel(writer, sheet_name="All_Transactions", index=False)


print("Cleaning complete.")
print(f"Cleaned file saved as: {output_file}")
print(f"Users rows: {len(users):,}")
print(f"Transaction rows: {len(transactions):,}")
