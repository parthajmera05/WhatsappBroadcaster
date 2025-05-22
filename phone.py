import pandas as pd

# Path to your Excel file
input_file = 'Phone1.xlsx'
output_file = 'Phone1_updated.xlsx'

# Column name containing phone numbers
phone_column = 'Phone'  # Change if different

# Read the Excel file
df = pd.read_excel(input_file)

# Clean and format the phone numbers
def clean_number(num):
    # Convert to string and strip spaces
    num = str(num).strip()

    # Remove .0 if present (from float interpretation)
    if num.endswith('.0'):
        num = num[:-2]

    # Remove all non-digit characters (optional, for robustness)
    num = ''.join(filter(str.isdigit, num))

    # Remove leading 0s
    num = num.lstrip('0')

    # Add +91 if not already present
    if not num.startswith('91'):
        num = '91' + num

    return '+' + num

# Apply cleaning function
df[phone_column] = df[phone_column].apply(clean_number)

# Save the cleaned data
df.to_excel(output_file, index=False)

print(f"Cleaned phone numbers saved to '{output_file}'")
