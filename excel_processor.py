import pandas as pd
import re
import os

def read_excel(file_path, sheet_name=0):
    """Read an Excel sheet into a pandas DataFrame."""
    return pd.read_excel(file_path, sheet_name=sheet_name)

def apply_rules(df, config):
    """Apply rules from the configuration to the DataFrame."""
    rules = config['rules']
    for index, row in df.iterrows():
        for rule in rules:
            regex = rule['regex']
            # Match regex against column H
            if re.search(regex, str(row['H'])):
                for col_rule in rule['columns']:
                    column = col_rule['column']
                    value = col_rule['value']
                    df.at[index, column] = str(value)  # Treat all values as strings
                break  # Apply only the first matching rule per row
    return df

def process_excel(rules_file):
    """Process Excel files based on rules defined in the rules Excel file."""
    # Read the rules Excel file with header
    rules_df = pd.read_excel(rules_file, header=0)
    
    # Ensure expected columns are present
    expected_cols = ['Input_File', 'Input_Sheet', 'Regex', 'Output_File', 'Output_Sheet']
    if not all(col in rules_df.columns for col in expected_cols):
        raise ValueError(f"rules.xls must contain columns: {expected_cols}")
    
    # Group rules by input and output locations
    for (input_file, input_sheet, output_file, output_sheet), group_df in rules_df.groupby(
        ['Input_File', 'Input_Sheet', 'Output_File', 'Output_Sheet']
    ):
        # Resolve relative paths to the script's directory
        input_file = os.path.join(os.path.dirname(__file__), input_file)
        output_file = os.path.join(os.path.dirname(__file__), output_file)
        
        # Read the specified input sheet
        df = read_excel(input_file, sheet_name=input_sheet)
        
        # Collect rules for this group
        rules = []
        for _, row in group_df.iterrows():
            regex = row['Regex']
            columns = []
            # Process change pairs dynamically (Change1_Column, Change1_Value, etc.)
            col_index = 5  # Start after Input_File, Input_Sheet, Regex, Output_File, Output_Sheet
            while col_index < len(rules_df.columns) - 1:  # Ensure pairs are available
                col_name = rules_df.columns[col_index]
                val_name = rules_df.columns[col_index + 1]
                if col_name.startswith('Change') and val_name.startswith('Change'):
                    col_value = row[col_name]
                    val_value = row[val_name]
                    if pd.notna(col_value) and pd.notna(val_value):
                        columns.append({'column': col_value, 'value': val_value})
                col_index += 2  # Move to next pair
            if columns:
                rules.append({'regex': regex, 'columns': columns})
        
        # Apply the rules to the DataFrame
        config = {'rules': rules}
        df_modified = apply_rules(df, config)
        
        # Write to the specified output sheet, preserving other sheets
        with pd.ExcelWriter(output_file, mode='a', if_sheet_exists='replace') as writer:
            df_modified.to_excel(writer, sheet_name=output_sheet, index=False)

# Example usage
if __name__ == "__main__":
    # Assume rules.xls is in the same directory as the script
    process_excel('rules.xls')
