
# Python script read excel (xlsx) and match text from text file

```markdown
> Austin.Lai |
> -----------| September 26th, 2023
> -----------| Updated on September 26th, 2023
```

---

## Table of Contents

<!-- TOC -->

- [Python script read excel xlsx and match text from text file](#python-script-read-excel-xlsx-and-match-text-from-text-file)
    - [Table of Contents](#table-of-contents)
    - [Disclaimer](#disclaimer)
    - [Description](#description)
    - [Script](#script)

<!-- /TOC -->

<br>

## Disclaimer

<span style="color: red; font-weight: bold;">DISCLAIMER:</span>

This project/repository is provided "as is" and without warranty of any kind, express or implied, including but not limited to the warranties of merchantability, fitness for a particular purpose and noninfringement. In no event shall the authors or copyright holders be liable for any claim, damages or other liability, whether in an action of contract, tort or otherwise, arising from, out of or in connection with the software or the use or other dealings in the software.

This project/repository is for <span style="color: red; font-weight: bold;">Educational</span> purpose <span style="color: red; font-weight: bold;">ONLY</span>. Do not use it without permission. The usual disclaimer applies, especially the fact that me (Austin) is not liable for any damages caused by direct or indirect use of the information or functionality provided by these programs. The author or any Internet provider bears NO responsibility for content or misuse of these programs or any derivatives thereof. By using these programs you accept the fact that any damage (data loss, system crash, system compromise, etc.) caused by the use of these programs is not Austin responsibility.

<br>

## Description

<!-- Description -->

Python script read excel (xlsx) and match text from text file.

The python script require `pandas` package.

All sample will be provide:

- `read_excel_and_match_text_from_text_file.py`
- `data.xlsx`
- `sample.txt`
- `sample_found_match.txt`

<!-- /Description -->

<br>

## Script

The `read_excel_and_match_text_from_text_file.py` file can be found [here](./read_excel_and_match_text_from_text_file.py) or below:

```python
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
```
