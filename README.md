# Update Spreadsheet 

## Introduction
This project utilizes the `openpyxl` library in Python to process Excel files, specifically focusing on modifying and charting data in a worksheet. The primary function, process_workbook, performs the following tasks:

1. Data Processing: Iterates through an Excel sheet, adjusts prices in column C by reducing them by 10%, and records the corrected prices in a new column (column D).

2. Chart Generation: Creates a bar chart based on the corrected prices and embeds it in the worksheet at cell 'E2'.

3. Colorful Slices: Enhances the chart by adding colorful slices with random colors.

## Requirements
- Python 3.x
- openpyxl library

## Usage
1. Install the required library using:
    ```
    pip install openpyxl
    ```

2. Place the Excel file you want to process in the same directory as the script.

3. Modify the filename in the `process_workbook` function to match your Excel file's name.

4. Run the script:
    ```
    python your_script_name.py
    ```

5. Check the Excel file for the modified data and the newly added chart.

## Code Explanation

```
import openpyxl as xl
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.marker import DataPoint
import random

# Function to generate random colors for chart slices
def generate_slices_colors(num_slices):
    # Generate random colors for slices
    colors = ['#' + ''.join(random.choices('0123456789ABCDEF', k=6)) for _ in range(num_slices)]
    slices = [DataPoint(idx=i) for i in range(num_slices)]
    for i, slice_point in enumerate(slices):
        slice_point.graphicalProperties.solidFill = colors[i]
    return slices

# Main function to process the Excel workbook
def process_workbook(filename):
    # Load the workbook
    wb = xl.load_workbook(filename)
    
    # Access the first sheet
    sheet = wb['Sheet1']

    # Adjust prices in column C and record corrected prices in column D
    for row in range(2, sheet.max_row + 1):
        cell = sheet.cell(row, 3)
        corrected_price = cell.value * 0.9
        corrected_price_cell = sheet.cell(row, 4)
        corrected_price_cell.value = corrected_price

    # Add a heading for the corrected prices column
    corrected_price_cell_heading = sheet.cell(1, 4)
    corrected_price_cell_heading.value = 'corrected price'

    # Define the data range for the chart
    values = Reference(sheet,
              min_row=2,
              max_row=sheet.max_row,
              min_col=4,
              max_col=4)

    # Create a bar chart and add data
    chart = BarChart()
    chart.add_data(values)
    
    # Add the chart to the worksheet at cell 'E2'
    sheet.add_chart(chart, 'e2')

    # Use the reusable function to generate slices with random colors
    slices = generate_slices_colors(3)
    
    # Apply the colorful slices to the chart
    chart.series[0].data_points = slices

    # Save the modified workbook
    wb.save(filename)

# Call the main function with the desired Excel file
process_workbook('transaction3.xlsx')

```

Feel free to adapt and integrate this code into your project. If you encounter any issues or have specific requirements, please refer to the documentation of the openpyxl library or seek assistance from the community.
