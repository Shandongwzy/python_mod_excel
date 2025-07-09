import pandas as pd
import json
import re

# Excel Read/Write Module
def read_excel(file_path, sheet_name=0):
    #The value 0 refers to the sheet's position (first sheet in the workbook). 
    #Alternatively, sheet_name can be a string (e.g., "Sheet1") to specify a sheet by name.
    """Read an Excel file into a pandas DataFrame."""
    return pd.read_excel(file_path, sheet_name=sheet_name)

def write_excel(df, file_path, sheet_name=0):
    """Write a pandas DataFrame to an Excel file."""
    df.to_excel(file_path, sheet_name=sheet_name, index=False)
    #index=False prevents the DataFrame index from being written to the Excel file, keeping the output cleaner.

# Rule-Matching Engine
def apply_rules(df, config):
    """Apply rules from the configuration to the DataFrame."""
    rules = config['rules']  # Extract the list of rules from the config dictionary
    for index, row in df.iterrows():  # Iterate over each row in the DataFrame
        for rule in rules:  # Iterate over each rule in the config
            regex = rule['regex']  # Get the regex pattern for the current rule
            # Check if column H matches the regex
            if re.search(regex, str(row['H'])):  # If the value in column H matches the regex
                # Apply column updates based on the rule
                for col_rule in rule['columns']:  # Iterate over column updates in the rule
                    column = col_rule['column']  # Target column name
                    value = col_rule['value']  # Value to set
                    dtype = col_rule['type']  # Data type for the value
                    if dtype == 'int':  # If the type is integer
                        df.at[index, column] = int(value)  # Set the column value as an integer
                    elif dtype == 'string':  # If the type is string
                        df.at[index, column] = str(value)  # Set the column value as a string
                # Stop after the first matching rule (can be modified for multiple outputs)
                break
    return df  # Return the modified DataFrame

def process_excel(input_file, config_file, output_file=None):
    """Process an Excel file based on a configuration file."""
    # Load configuration
    with open(config_file, 'r') as f:
        config = json.load(f)
    
    # Read Excel file
    df = read_excel(input_file)
    
    # Apply rules
    df_modified = apply_rules(df, config)
    
    # Determine output file
    output = output_file if output_file else input_file
    for rule in config['rules']:
        if 'output_file' in rule and rule['output_file'] != 'same':
            output = rule['output_file']
            break
    
    # Write to Excel
    write_excel(df_modified, output)

# Example usage
if __name__ == "__main__":
    # Sample configuration
    sample_config = {
        "rules": [
            {
                "regex": "test.*",
                "columns": [
                    {"column": "B", "value": "matched", "type": "string"},
                    {"column": "K", "value": 456, "type": "int"}
                ],
                "output_file": "same"
            },
            {
                "regex": "data.*",
                "columns": [
                    {"column": "B", "value": "data_found", "type": "string"},
                    {"column": "K", "value": 789, "type": "int"}
                ],
                "output_file": "output.xls"
            }
        ]
    }
    
    # Save sample config to file
    with open('config.json', 'w') as f:
        json.dump(sample_config, f)
    
    # Process the Excel file
    process_excel('input.xls', 'config.json')