#!/usr/bin/env python3
"""
Convert Markdown file back to Excel format
Usage: python convert-to-excel.py
"""

import pandas as pd
import re
from io import StringIO

def markdown_to_excel(markdown_file, output_excel):
    """Convert markdown file with tables back to Excel"""
    
    with open(markdown_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Split content by sheet headers (## SheetName)
    # Pattern to match: ## SheetName followed by metadata and table
    sheet_pattern = r'## (.+?)\\n\\n\\*\\*Rows:\\*\\* \\d+ \\| \\*\\*Columns:\\*\\* \\d+\\n\\n(.+?)(?=\\n---\\n\\n|\\Z)'
    
    sheets = re.findall(sheet_pattern, content, re.DOTALL)
    
    # Create Excel writer
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        for sheet_name, table_content in sheets:
            # Skip if it's a non-data sheet like TOC
            if sheet_name == "Table of Contents":
                continue
                
            try:
                # Parse markdown table to dataframe
                df = pd.read_csv(StringIO(table_content), sep='|', skipinitialspace=True)
                
                # Remove first and last columns (they're empty due to markdown table format)
                df = df.iloc[:, 1:-1]
                
                # Strip whitespace from column names
                df.columns = df.columns.str.strip()
                
                # Remove the header separator row if it exists
                if df.iloc[0].astype(str).str.contains('---').all():
                    df = df.iloc[1:]
                
                # Strip whitespace from all string columns
                for col in df.columns:
                    if df[col].dtype == 'object':
                        df[col] = df[col].str.strip()
                
                # Convert empty strings back to NaN
                df = df.replace('', pd.NA)
                
                # Write to Excel
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"✓ Converted sheet: {sheet_name} ({len(df)} rows)")
                
            except Exception as e:
                print(f"✗ Error converting sheet {sheet_name}: {e}")
                continue
    
    print(f"\\n✓ Excel file created: {output_excel}")

if __name__ == "__main__":
    markdown_file = "laserfiche-ai-visibility-map.md"
    output_excel = "Laserfiche_AI_Visibility_Map-updated.xlsx"
    
    try:
        markdown_to_excel(markdown_file, output_excel)
        print("\\n✓ Conversion completed successfully!")
    except Exception as e:
        print(f"\\n✗ Error: {e}")
