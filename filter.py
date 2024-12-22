import pandas as pd

# Load the Excel file without skipping any rows
file_name = "savedrecs_wos_all.xls"
data = pd.read_excel(file_name)

# Filter rows for the "Abstract" column where it's not empty
if "Abstract" in data.columns:
    data = data.dropna(subset=["Abstract"])

# Filter rows for the "Article Title" column where it's not empty
if "Article Title" in data.columns:
    data = data.dropna(subset=["Article Title"])

if "DOI" in data.columns:
    data = data.dropna(subset=["DOI"])

# Remove rows where "Researcher Id" is empty
if "Researcher Ids" in data.columns:
    data = data.dropna(subset=["Researcher Ids"])

if "Publication Year" in data.columns:
    data = data.dropna(subset=["Publication Year"])

# List of columns to include in the new Excel file
required_columns = [
    'Article Title',
    'Abstract',
    'Researcher Ids',
    'ORCIDs',
    'DOI',
    'WoS Categories',
    'Research Areas',
    'IDS Number',
    'UT (Unique WOS ID)',
    'Publication Year'
]

# Filter the data to include only the required columns
data = data[required_columns]

# Save the filtered data back to an Excel file
output_file_name = "filtered_file_2.xlsx"
data.to_excel(output_file_name, index=False)
print(f"Filtered data has been saved to '{output_file_name}'.")
