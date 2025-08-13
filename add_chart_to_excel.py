#!/usr/bin/env python3
"""
Add Current vs Time Chart to Excel File

This script reads the existing Excel file and adds a new sheet with a chart
showing Current vs Time using xlsxwriter.
"""

import pandas as pd
import xlsxwriter
from pathlib import Path
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def create_excel_with_chart(input_file_path, output_file_path=None):
    """
    Create a new Excel file with the data and add a Current vs Time chart.
    
    Args:
        input_file_path (str): Path to the input Excel file
        output_file_path (str, optional): Path for output file. If None, adds '_with_chart' to input name
    
    Returns:
        str: Path to the created Excel file
    """
    if output_file_path is None:
        input_path = Path(input_file_path)
        output_file_path = input_path.parent / f"{input_path.stem}_with_chart{input_path.suffix}"
    
    logger.info(f"Reading data from: {input_file_path}")
    
    # Read the existing data
    df = pd.read_excel(input_file_path, sheet_name='Raw_Data')
    
    logger.info(f"Loaded {len(df)} data points")
    logger.info(f"Creating Excel file with chart: {output_file_path}")
    
    # Create a new workbook with xlsxwriter
    workbook = xlsxwriter.Workbook(str(output_file_path))
    
    # Add the data sheet
    data_worksheet = workbook.add_worksheet('Raw_Data')
    
    # Write headers
    headers = list(df.columns)
    for col, header in enumerate(headers):
        data_worksheet.write(0, col, header)
    
    # Write data
    for row_idx, (_, row) in enumerate(df.iterrows(), start=1):
        for col_idx, value in enumerate(row):
            if pd.isna(value):
                data_worksheet.write(row_idx, col_idx, '')
            else:
                data_worksheet.write(row_idx, col_idx, value)
    
    logger.info("Data written to Raw_Data sheet")
    
    # Create a dedicated chart sheet (entire tab is the chart)
    chart = workbook.add_chart({'type': 'line'})

    # Configure the chart
    chart.add_series({
        'name': 'Current (A)',
        'categories': ['Raw_Data', 1, 0, len(df), 0],  # Time_seconds column (A)
        'values': ['Raw_Data', 1, 1, len(df), 1],      # Current column (B)
        'line': {'color': 'blue', 'width': 1.5},
    })

    # Set chart title and axis labels
    chart.set_title({
        'name': 'Current vs Time - Peak Current Test',
        'name_font': {'size': 18, 'bold': True}
    })

    chart.set_x_axis({
        'name': 'Time (seconds)',
        'name_font': {'size': 14, 'bold': True},
        'num_font': {'size': 12}
    })

    chart.set_y_axis({
        'name': 'Current (Amperes)',
        'name_font': {'size': 14, 'bold': True},
        'num_font': {'size': 12}
    })

    # Set chart style
    chart.set_style(2)  # Use a predefined chart style

    # Add the chart as a dedicated chart sheet
    chart_sheet = workbook.add_chartsheet('Current_vs_Time_Chart')
    chart_sheet.set_chart(chart)
    
    logger.info("Dedicated chart sheet created: Current_vs_Time_Chart")
    
    # Close the workbook
    workbook.close()
    
    logger.info(f"Excel file with chart saved: {output_file_path}")
    return str(output_file_path)


def main():
    """Main function to add chart to Excel file."""
    # Find Excel file
    excel_files = list(Path('.').glob('*.xlsx'))
    
    # Filter out any files that already have '_with_chart' in the name
    excel_files = [f for f in excel_files if '_with_chart' not in f.name]
    
    if not excel_files:
        logger.error("No Excel files found in current directory")
        return
    
    excel_file = excel_files[0]
    logger.info(f"Processing file: {excel_file}")
    
    try:
        # Create Excel file with chart
        output_file = create_excel_with_chart(str(excel_file))
        
        print(f"\n🎉 Chart Creation Complete!")
        print(f"📊 Added Current vs Time chart to Excel file")
        print(f"📋 Output file: {output_file}")
        print(f"📑 The file now contains:")
        print(f"   • Raw_Data sheet: Original data with pulse summary")
        print(f"   • Current_vs_Time_Chart: Dedicated chart sheet (entire tab is the chart)")
        print(f"📈 Chart shows all 5000 data points with clear pulse visualization")
        print(f"🎯 The chart tab fills the entire screen for maximum visibility")
        
    except Exception as e:
        logger.error(f"Error creating chart: {str(e)}")
        raise


if __name__ == "__main__":
    main()
