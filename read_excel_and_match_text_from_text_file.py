import pandas as pd

# Read the Excel file
xls = pd.ExcelFile('data.xlsx')

# Print the sheet names
print(xls.sheet_names)

# Read the 'Named_Sheet2' tab sheet into a DataFrame
df = pd.read_excel(xls, 'Named_Sheet2')

# Convert all text to lowercase
df = df.applymap(lambda s:s.lower() if type(s) == str else s)

# Read the text file line by line with encoding='utf-16'
with open('sample.txt', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Convert all text to lowercase and store in a list
text_list = [line.lower().strip() for line in lines]

# Split each line by dot and store only the second part in the list
text_list = [line.split('.')[1] for line in text_list if '.' in line]

# Remove duplicates from the list
text_list = list(set(text_list))

print(text_list)
print('\n\n')

# Search in 'Name' or 'Link' columns for matches with the list
matching_rows = df[(df['Name'].isin(text_list)) | (df['Link'].isin(text_list))]

if not matching_rows.empty:
    matching_item = matching_rows[['Link']]
    print(matching_item)

    # Store matched 'Link' values into variables
    link = matching_rows['Link'].values

    print(link)

    # Save the values into a text file
    with open('found_match.txt', 'w') as f:
        f.write(str(matching_item))
