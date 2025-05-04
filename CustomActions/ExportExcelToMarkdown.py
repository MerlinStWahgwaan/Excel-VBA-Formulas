import pandas as pd
import re
import os

def parse_range(range_str):
    """Parse Excel range (e.g., A1:G60) into column letters and row numbers."""
    match = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", range_str, re.I)
    if not match:
        raise ValueError("Invalid range format. Use A1:G60 format.")
    start_col, start_row, end_col, end_row = match.groups()
    return start_col, int(start_row) - 1, end_col, int(end_row) - 1  # 0-based indexing

def col_letter_to_index(col_letter):
    """Convert Excel column letter (e.g., A, G) to 0-based index."""
    col_letter = col_letter.upper()
    index = 0
    for char in col_letter:
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1

def create_markdown_table(df, selected_columns, selected_rows):
    """Create a Markdown table from selected DataFrame rows and columns."""
    # Subset the DataFrame
    selected_data = df.iloc[selected_rows, selected_columns]
    
    # Get column names (use DataFrame headers)
    col_names = selected_data.columns
    
    # Create Markdown table
    markdown_lines = []
    
    # Header row
    header = '| ' + ' | '.join(str(col) for col in col_names) + ' |'
    markdown_lines.append(header)
    
    # Separator row
    separator = '| ' + ' | '.join(['-' * max(len(str(col)), 3) for col in col_names]) + ' |'
    markdown_lines.append(separator)
    
    # Data rows
    for _, row in selected_data.iterrows():
        row_data = '| ' + ' | '.join(str(val) if pd.notnull(val) else '' for val in row) + ' |'
        markdown_lines.append(row_data)
    
    return '\n'.join(markdown_lines)

def main():
    # Get file path
    file_path = input("Enter the Excel file path (e.g., C:/path/to/file.xlsx or file.xlsx in same directory): ").strip()
    if not os.path.exists(file_path):
        print("Error: File not found.")
        return
    
    try:
        # List available sheets
        excel_file = pd.ExcelFile(file_path)
        print("Available sheets:", excel_file.sheet_names)
        
        # Get sheet name
        sheet_name = input("Enter the sheet name: ").strip()
        
        # Get range
        range_str = input("Enter the range (e.g., A1:G60): ").strip()
        
        # Parse range
        start_col, start_row, end_col, end_row = parse_range(range_str)
        
        # Convert column letters to indices
        start_col_idx = col_letter_to_index(start_col)
        end_col_idx = col_letter_to_index(end_col)
        
        # Read Excel file (specific sheet)
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        except ValueError:
            print(f"Error: Sheet '{sheet_name}' not found in the Excel file.")
            return
        
        # Validate range
        max_rows, max_cols = df.shape
        if start_row >= max_rows or end_row >= max_rows or start_col_idx >= max_cols or end_col_idx >= max_cols:
            print("Error: Specified range exceeds data dimensions.")
            return
        
        # Select rows and columns
        selected_rows = range(start_row, end_row + 1)
        selected_columns = range(start_col_idx, end_col_idx + 1)
        
        # Create Markdown table
        markdown_table = create_markdown_table(df, selected_columns, selected_rows)
        
        # Save to file
        output_file = "ExcelToMarkdown.txt"
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(markdown_table)
        
        print(f"Markdown table saved to '{output_file}'")
        
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()