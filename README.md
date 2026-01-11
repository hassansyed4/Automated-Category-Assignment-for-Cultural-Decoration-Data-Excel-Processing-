# Automated Category Assignment for Excel Data

## Overview
This project implements a Python-based solution to automatically populate the **Categories** column in an Excel file by analyzing descriptive text in the **Decoration** column. The script applies rule-based keyword matching derived from a reference Word document to ensure consistent and accurate category assignment.

## Data Structure
- **Input (Excel File):**
  - Decoration: Free-text description of figures, scenes, or activities
  - Categories: Empty column populated by the script
- Object area designations (A, B, C, etc.) are not present; each row is processed as a single combined description.

## Processing Logic
1. Reads the Decoration column from Excel
2. Cleans and normalizes text
3. Matches keywords against predefined category rules
4. Applies exclusion and combination logic where required
5. Writes matched categories back to the Categories column

## Key Features
- Keyword-based category mapping
- Support for combined keyword rules
- Exclusion logic to prevent invalid category assignments
- Pandas-based Excel processing

## Tech Stack
- Python
- Pandas
- Regular Expressions
- Excel (xlsx)

## Output
An updated Excel file with accurately populated category labels based on textual descriptions.

## Status
Logic confirmed and implemented following requirements validation.
