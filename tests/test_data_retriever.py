import pytest
import sys
import os
import pandas as pd

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from DataRetriever import DataRetriever  # Import the DataRetriever class

# Path to the real Excel file for testing (adjust this path as needed)
EXCEL_FILE_PATH = "data/test_data.xlsx"  # Path to your real Excel file

# Test case to retrieve support data
def test_retrieve_support_data1():
    data_retriever = DataRetriever("C:/Users/COOPERJ/Documents/Projects/AutoPIPE_Automation/Sample Files/HSG_R2_CreatedReport_Local.xlsx")

    support_nodes = ["17", "18", "29", "36", "39", "43A", "50", "62", "72", "78", "81", "99", "TS1", "TS2", "TS3", "TS4"]
    result = data_retriever.retrieve_support_data(support_nodes)
    
    # Test if the result is a dictionary
    assert isinstance(result, dict)

    # Check if the expected support nodes are in the result dictionary
    for node in support_nodes:
        assert node in result

    # Example expected values for support nodes 17, 18, and 29
    expected_values = {'17': ['1-MCT-HS-H322', [0, 0], [-2935, -2792], [0, 0]], 
                       '18': ['1-MCT-HS-H321', [0, 0], [-2410, -2394], [0, 0]], 
                       '29': ['1-MCT-HS-H363', [0, 0], [-1892, -1858], [0, 0]], 
                       '36': ['1-MCT-HS-H181', [0, 0], [-1098, -1056], [0, 0]],
                         '39': ['1-MCT-HS-H182', [0, 0], [-1961, -1891], [0, 0]], 
                         '43A': ['1-MCT-HS-H365', [0, 0], [-1412, -1320], [0, 0]], 
                         '50': ['1-MCT-HS-H364', [0, 0], [-1744, -1639], [0, 0]], 
                         '62': ['1-MCT-HS-H180', [0, 0], [-708, -686], [0, 0]], 
                         '72': ['1-MCT-HS-H179', [0, 0], [-500, -361], [0, 0]], 
                         '78': ['1-MCT-HS-H178', [0, 0], [-455, -348], [0, 0]], 
                         '81': ['1-MCT-HS-H177', [0, 0], [-677, -653], [0, 0]], 
                         '99': ['1-MCT-HS-H317', [0, 0], [-85, -84], [0, 0]], 
                         'TS1': ['TANK ROD SUPPORT #1', [0, 0], [-3418, -2433], [0, 0]], 
                         'TS2': ['TANK ROD SUPPORT #2', [0, 0], [-3959, -1901], [0, 0]], 
                         'TS3': ['TANK ROD SUPPORT #3', [0, 0], [-491, 336], [0, 0]], 
                         'TS4': ['TANK ROD SUPPORT #4', [0, 0], [-1023, 876], [0, 0]]}

    # Validate that the retrieved values are correct for specific support nodes
    for node, expected_data in expected_values.items():
        actual_data = result[node]
        
        # Check that the actual support number/tag matches
        assert actual_data[0] == expected_data[0]
        
        # Check that the force lists are the same (considering the possibility of rounding differences)
        assert actual_data[1] == expected_data[1]
        assert actual_data[2] == expected_data[2]
        assert actual_data[3] == expected_data[3]
        
        # Optionally, you can check for exact matches for other support nodes
        # You can add more nodes here for validation
    
    print("Test passed: All support node values are correct.")



