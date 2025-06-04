import os
from excel_consolidator import ExcelConsolidator

def test_consolidator():
    # Create a test directory with sample Excel files
    test_dir = "test_data"
    os.makedirs(test_dir, exist_ok=True)
    
    # Create sample Excel files
    sample_data = [
        {"A": "Key1", "B": "Value1"},
        {"A": "Key2", "B": "Value2"}
    ]
    
    import pandas as pd
    pd.DataFrame(sample_data).to_excel(f"{test_dir}/test1.xlsx", index=False)
    pd.DataFrame(sample_data).to_excel(f"{test_dir}/test2.xlsx", index=False)
    
    # Test the consolidator
    consolidator = ExcelConsolidator(test_dir)
    result = consolidator.consolidate("output.xlsx")
    
    if result:
        print("Test passed! Output file created successfully.")
    else:
        print("Test failed!")

if __name__ == "__main__":
    test_consolidator()
