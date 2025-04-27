import pytest
import os
import pandas as pd

@pytest.fixture(scope="session")
def test_excel_file():
    # Create a test Excel file
    data = {
        'Date': pd.date_range(start='1/1/2024', periods=365, freq='D'),
        'power used': range(1, 366),
        'Junk': ['junk'] * 365
    }
    df = pd.DataFrame(data)
    file_path = "testEmailSpoof.xlsx"
    df.to_excel(file_path, sheet_name='loc3', index=False)
    yield file_path
    # Teardown: remove the file after tests
    os.remove(file_path)