import pandas as pd
from docx import Document
from docx.shared import Inches
import datetime
import os

# Load the CSV file
file_path = r'/Users/marciacripps/Desktop/GitHub/cow_days_calculator.py/fakecows.csv'
data = pd.read_csv(file_path)

# Calculate days since the date
data['Days Since'] = (datetime.datetime.now() - pd.to_datetime(data['Date'])).dt.days

# Categorize the data
categories = {
    'Less than 30 days': (0, 30),
    'Less than 60 days': (30, 60),
    'Less than 90 days': (60, 90),
    'Less than 120 days': (90, 120),
}

categorized_data = {category: data[(data['Days Since'] >= min_days) & (data['Days Since'] < max_days)] for category, (min_days, max_days) in categories.items()}

# Create a Word document
doc = Document()

# Add a table to the Word document
table = doc.add_table(rows=1, cols=len(categories))
table.style = 'Table Grid'

# Set the column headers
header_cells = table.rows[0].cells
for i, category in enumerate(categories.keys()):
    header_cells[i].text = category

# Add the data to the table
max_rows = max(len(dataframe) for dataframe in categorized_data.values())
for i in range(max_rows):
    row_cells = table.add_row().cells
    for j, dataframe in enumerate(categorized_data.values()):
        if i < len(dataframe):
            row_cells[j].text = dataframe.iloc[i]['Cow Name']

# Create a folder named 'Doc Files' if it does not exist
current_directory = os.getcwd()
doc_files_path = os.path.join(current_directory, 'Doc Files')

if not os.path.exists(doc_files_path):
    os.makedirs(doc_files_path)

# Save the Word document with today's date and 'cowbreeding' in the filename
today = datetime.datetime.now().strftime('%Y-%m-%d')
doc.save(os.path.join(doc_files_path, f'{today}_cowbreeding.docx'))
