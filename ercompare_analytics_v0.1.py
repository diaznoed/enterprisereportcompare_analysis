import pandas as pd

# Load the Excel file and the specific sheet for analysis
# Update the file_path to point to your Excel file with comparison data
file_path = "/path/to/your/file.xlsx"  # <--- Update this path
df = pd.read_excel(file_path, sheet_name='InventoryAnalytics')  # Ensure the sheet name matches your data

# Function to extract numeric values from text (for "RecordOne" and "RecordTwo" columns)
def extract_number(text):
    # Extracts digits from the given text and returns as an integer
    return int(''.join(filter(str.isdigit, text)))

# Loop through the rows and calculate the difference where applicable
for index, row in df.iterrows():
    try:
        # Extract numbers from RecordOne and RecordTwo columns
        record_one_value = extract_number(row['RecordOne'])  # Extract numeric value from RecordOne
        record_two_value = extract_number(row['RecordTwo'])  # Extract numeric value from RecordTwo
        
        # Calculate the difference (positive or negative)
        difference = record_two_value - record_one_value
        
        # Update the 'Difference Count' column with the calculated difference
        df.at[index, 'Difference Count'] = difference
    except:
        # Handle cases where numeric extraction fails (e.g., non-numeric records)
        df.at[index, 'Difference Count'] = None

# Save the updated dataframe to a new Excel file
# Update the output_path to specify where you want to save the updated file
output_path = '/path/to/output/Analytics_Difference.xlsx'  # <--- Update this path
df.to_excel(output_path, index=False)

# Notify the user that the updated file has been saved
print(f"Updated spreadsheet saved to: {output_path}")
