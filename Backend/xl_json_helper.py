import json
# Remove Flask import since we don't need it for this function
# from flask import jsonify

def get_cell_content(data, worksheet_name=None, cell_coordinate=None):
    """Get cell content with O(1) lookup"""
    try:
        if worksheet_name is None:  # Use the first worksheet (skip 'schema' key)
            worksheet_names = [key for key in data.keys() if key != 'schema']
            worksheet_name = worksheet_names[0] if worksheet_names else None
        
        if worksheet_name is None:
            print("No worksheet found")
            return None
            
        # Access cell data as array: [value, formula, hyperlink]
        cell_data = data[worksheet_name][str(cell_coordinate)]
        
        return {
            'value': cell_data[0],      # Index 0 = value
            'formula': cell_data[1],    # Index 1 = formula
        }
    except KeyError as e:
        print(f"KeyError: {e}")
        print(f"Available worksheets: {[key for key in data.keys() if key != 'schema']}")
        print(f"Worksheet name: {worksheet_name}")
        return None
    except Exception as e:
        print(f"Error: {e}")
        return None
    
def main():
    data = json.load(open("extract-output/1bf1eda2-ca7b-4c95-8ca8-cbf3da39b5df_Family_budget_monthly_cell_content.json"))
    print(get_cell_content(data, "Monthly budget report", "C3"))

if __name__ == "__main__":
    main()