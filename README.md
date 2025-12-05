# EXCELTEST - Excel to Markdown Converter

This repository contains data from the Laserfiche AI Visibility Map Excel file, converted to Markdown format for easy viewing and editing on GitHub.

## Files

- `laserfiche-ai-visibility-map.md` - The main data file containing all sheets from the Excel file
- `convert-to-excel.py` - Python script to convert the Markdown file back to Excel format

## How to Use

### Viewing the Data
Simply open `laserfiche-ai-visibility-map.md` to view all the data in a readable format on GitHub.

### Editing the Data
1. Click on `laserfiche-ai-visibility-map.md`
2. Click the edit button (pencil icon)
3. Make your changes
4. Commit the changes

### Converting Back to Excel

To convert your edited Markdown file back to Excel format:

```bash
python convert-to-excel.py
```

This will create a new `Laserfiche_AI_Visibility_Map-updated.xlsx` file with your changes.

## Requirements

To run the conversion script, you need:
- Python 3.x
- pandas
- openpyxl

Install with:
```bash
pip install pandas openpyxl
```

## Structure

The Markdown file contains the following sheets:
1. AI Visibility Map
2. Dashboard
3. Third-Party Blogs
4. PR
5. Competitor vs Own Content
6. Competitor pages
7. SocialUCG
8. Review sites

Each sheet is represented as a Markdown table with metadata (row/column counts).
