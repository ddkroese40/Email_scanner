import pytest
import os
import pandas as pd
import sys

# Add the parent directory to sys.path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from process_excel import process_excel_file

def test_process_excel_file(test_excel_file):
    location = "loc3"
    process_excel_file(test_excel_file, location)

    # Check if the output file exists
    output_file = os.path.join("Output", f"{location}_2024.xlsx")
    assert os.path.exists(output_file)

    # Check the contents of the output file
    df = pd.read_excel(output_file)
    assert 'Year' in df.columns
    assert 'Month' in df.columns
    assert 'power used' in df.columns

    # Cleanup
    os.remove(output_file)