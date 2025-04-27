import pandas as pd
import os

def process_excel_file(file_path, location):
    print(f"Processing Excel file: {file_path}, Location: {location}")
    try:
        df = pd.read_excel(file_path, sheet_name=location)
        df['Date'] = pd.to_datetime(df['Date'])
        df['Month'] = df['Date'].dt.to_period('M')
        df['Year'] = df['Date'].dt.year

        # Create output directory if it doesn't exist
        output_dir = "Output"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Group by year and month and sum power usage
        yearly_data = df.groupby(['Year', 'Month'])['power used'].sum().reset_index()

        # Save data to Excel file
        output_file = os.path.join(output_dir, f"{location}_{df['Year'].iloc[0]}.xlsx")
        yearly_data.to_excel(output_file, index=False)
        print(f"Saved power usage data to {output_file}")
    except Exception as e:
        print(f"Error processing Excel file: {e}")