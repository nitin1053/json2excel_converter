import pandas as pd
import json

def read_nested_json(json_file_path):
    try:
        with open(json_file_path, 'r') as file:
            data = json.load(file)
        return data
    except FileNotFoundError:
        print(f"Error: File not found - {json_file_path}")
        return None
    except json.JSONDecodeError:
        print(f"Error: Incorrect JSON format in file - {json_file_path}")
        return None

def flatten_json(json_data, parent_key='', separator='_'):
    flattened = {}
    if isinstance(json_data, dict):
        for key, value in json_data.items():
            new_key = f"{parent_key}{separator}{key}" if parent_key else key
            flattened.update(flatten_json(value, new_key, separator))
    elif isinstance(json_data, list):
        for i, item in enumerate(json_data):
            new_key = f"{parent_key}{separator}{i}" if parent_key else str(i)
            flattened.update(flatten_json(item, new_key, separator))
    else:
        flattened[parent_key] = json_data
    return flattened

def json_to_excel(json_file_path, excel_file_path):
    data = read_nested_json(json_file_path)

    if data:
        flattened_data = flatten_json(data)
        df = pd.DataFrame([flattened_data])
        df.to_excel(excel_file_path, index=False)
        print(f"Conversion successful! Excel file saved at {excel_file_path}")


# Example Usage:
json_file_path = '/home/nitin1053/Documents/python files/json py/input.json'
excel_file_path = '/home/nitin1053/Documents/python files/json py/input.xlsx'
json_to_excel(json_file_path, excel_file_path)
