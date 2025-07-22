#!/usr/bin/env python3
import json
import sys
import os
from pathlib import Path
import re
from typing import Dict, List, Tuple, Optional, Any
from openpyxl import load_workbook, Workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import NamedStyle
import shutil
from datetime import datetime
from copy import copy


def safe_print(text):
    """Print text with encoding error handling"""
    try:
        print(text)
    except UnicodeEncodeError:
        # If we can't print the text, print a safe version
        safe_text = text.encode('ascii', 'replace').decode('ascii')
        print(safe_text)


class ExcelFormatter:
    def __init__(self, input_file: str, target_file: str, mapping_file: str, 
                 output_file: str, input_sheet: str, target_sheet: str, formula_row: int = 100):
        self.input_file = input_file
        self.target_file = target_file
        self.mapping_file = mapping_file
        self.output_file = output_file
        self.input_sheet_name = input_sheet
        self.target_sheet_name = target_sheet
        self.formula_row = formula_row
        
        self.input_wb = None
        self.target_wb = None  # This will be our working copy
        self.input_ws = None
        self.target_ws = None  # This will be our working worksheet
        
        self.mapping_config = None
        self.target_to_scanned = {}
        self.column_formulas = {}  # column_letter -> formula_template from specified row
        self.error_messages = []
        
    def load_files(self):
        """Load all required files"""
        try:
            # Check if mapping file exists
            if not os.path.exists(self.mapping_file):
                raise Exception(f"Mapping file not found: {self.mapping_file}")
            
            # Load mapping configuration with UTF-8 encoding
            with open(self.mapping_file, 'r', encoding='utf-8') as f:
                self.mapping_config = json.load(f)
            
            safe_print(f"DEBUG: Loaded mapping file: {self.mapping_file}")
            safe_print(f"DEBUG: Mapping config keys: {list(self.mapping_config.keys())}")
            
            # Create reverse mapping
            mappings_count = 0
            for mapping in self.mapping_config.get('mappings', []):
                if not mapping.get('is_ignored', False) and mapping.get('target_column'):
                    target_col = mapping['target_column']
                    scanned_col = mapping['scanned_column']
                    self.target_to_scanned[target_col] = scanned_col
                    mappings_count += 1
                    safe_print(f"DEBUG: Mapping '{target_col}' -> '{scanned_col}'")
            
            safe_print(f"DEBUG: Total mappings loaded: {mappings_count}")
            
            # Load Excel files
            self.input_wb = load_workbook(self.input_file, data_only=False)
            self.target_wb = load_workbook(self.target_file, data_only=False)  # Working copy
            
            # Get worksheets
            safe_print(f"DEBUG: Input sheets: {self.input_wb.sheetnames}")
            safe_print(f"DEBUG: Target sheets: {self.target_wb.sheetnames}")
            
            self.input_ws = self.input_wb[self.input_sheet_name]
            self.target_ws = self.target_wb[self.target_sheet_name]  # Working worksheet
            
        except Exception as e:
            raise Exception(f"Failed to load files: {e}")
    
    def detect_column_formulas(self):
        """Detect column-wide formulas in target format at specified formula row"""
        self.column_formulas = {}
        
        # Check if formula row exists
        if self.target_ws.max_row < self.formula_row:
            safe_print(f"DEBUG: Row {self.formula_row} does not exist in target format, no column formulas to detect")
            return
        
        safe_print(f"DEBUG: Checking for column formulas in row {self.formula_row}")
        
        for col_idx in range(1, self.target_ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            cell = self.target_ws[f'{col_letter}{self.formula_row}']
            
            # Check for both formula cells and text cells that contain formulas
            formula_text = None
            
            if cell.data_type == 'f' and cell.value:  # Formula cell
                # Handle different types of formula objects
                if hasattr(cell.value, 'text'):  # ArrayFormula object
                    formula_text = cell.value.text
                    safe_print(f"DEBUG: Found ArrayFormula in {col_letter}{self.formula_row}: {formula_text}")
                elif isinstance(cell.value, str):  # Regular formula string
                    formula_text = cell.value
                    safe_print(f"DEBUG: Found regular formula in {col_letter}{self.formula_row}: {formula_text}")
                else:
                    # Try to convert to string
                    formula_text = str(cell.value)
                    safe_print(f"DEBUG: Found formula object in {col_letter}{self.formula_row}, converted to: {formula_text}")
            elif cell.value and isinstance(cell.value, str) and cell.value.strip().startswith('='):
                # Text cell containing a formula (what you manually typed)
                formula_text = cell.value.strip()
                safe_print(f"DEBUG: Found text formula in {col_letter}{self.formula_row}: {formula_text}")
            
            if formula_text:
                # Clean up the formula text
                formula_text = formula_text.strip()
                if not formula_text.startswith('='):
                    formula_text = '=' + formula_text
                
                # Store the formula template
                self.column_formulas[col_letter] = formula_text
                header_cell = self.target_ws[f'{col_letter}1']
                header_value = header_cell.value if header_cell.value else "Unknown"
                safe_print(f"Found column formula in {col_letter}{self.formula_row} ({header_value}): {formula_text}")
    
    def get_input_headers(self) -> Dict[str, int]:
        """Get input file headers mapping"""
        headers = {}
        safe_print("DEBUG: Input file headers:")
        for col_idx in range(1, self.input_ws.max_column + 1):
            cell = self.input_ws.cell(row=1, column=col_idx)
            if cell.value:
                header_value = str(cell.value).strip()
                headers[header_value] = col_idx
                safe_print(f"  Column {col_idx}: '{header_value}'")
        return headers
    
    def get_target_headers(self) -> List[Tuple[str, int]]:
        """Get target format headers with their column indices"""
        headers = []
        safe_print("DEBUG: Target format headers:")
        for col_idx in range(1, self.target_ws.max_column + 1):
            cell = self.target_ws.cell(row=1, column=col_idx)
            header_value = str(cell.value).strip() if cell.value else ""
            headers.append((header_value, col_idx))
            safe_print(f"  Column {col_idx}: '{header_value}'")
        return headers
    
    def copy_data_type_only(self, source_cell, target_cell):
        """Copy only the data type from source to target, preserving target's formula"""
        # If target has a formula, preserve it but try to match data type characteristics
        if target_cell.data_type == 'f':
            # Keep the formula but we could potentially adjust number format if needed
            # Copy number format from source to help with result formatting
            if source_cell.number_format != 'General':
                target_cell.number_format = source_cell.number_format
        else:
            # Target has no formula, so we shouldn't reach here in the current logic
            pass
    
    def copy_cell_value_with_type_preservation(self, source_cell, target_cell):
        """Copy cell value and data type from source to target"""
        # Copy the value with type preservation
        if source_cell.data_type == 'f':  # Formula
            target_cell.value = source_cell.value
        elif source_cell.data_type == 'd':  # Date
            target_cell.value = source_cell.value
        elif source_cell.data_type == 'n':  # Number
            target_cell.value = source_cell.value
        elif source_cell.data_type == 'b':  # Boolean
            target_cell.value = source_cell.value
        else:  # String or other
            target_cell.value = source_cell.value
        
        # Copy number format to preserve data type appearance
        if source_cell.number_format != 'General':
            target_cell.number_format = source_cell.number_format
    
    def adjust_formula_for_row(self, formula: str, target_row: int) -> str:
        """Adjust formula template for specific row (replace 'n' or 'N' with row number)"""
        # Convert to string if it's not already
        if not isinstance(formula, str):
            formula = str(formula)
        
        # Ensure formula starts with =
        formula = formula.strip()
        if not formula.startswith('='):
            formula = '=' + formula
        
        # Replace both lowercase 'n' and uppercase 'N' with row number
        # Use case-insensitive replacement to catch both
        adjusted = formula.replace('n', str(target_row)).replace('N', str(target_row))
        
        safe_print(f"DEBUG: Adjusted formula from '{formula}' to '{adjusted}' for row {target_row}")
        return adjusted
    
    def process_column_with_formula(self, col_idx: int, formula_template: str, max_data_row: int):
        """Apply column-wide formula to all data rows"""
        col_letter = get_column_letter(col_idx)
        
        safe_print(f"DEBUG: Processing column {col_letter} with formula template: {formula_template}")
        
        formulas_applied = 0
        for row_idx in range(2, max_data_row + 1):  # Start from row 2 (after header)
            target_cell = self.target_ws.cell(row=row_idx, column=col_idx)
            adjusted_formula = self.adjust_formula_for_row(formula_template, row_idx)
            target_cell.value = adjusted_formula
            formulas_applied += 1
        
        safe_print(f"Applied column formula to {col_letter}: {formula_template} ({formulas_applied} cells)")
    
    def process_column_with_mapping(self, target_col_idx: int, target_header: str, 
                                   scanned_column: str, input_headers: Dict[str, int]):
        """Process a column that has data mapping"""
        safe_print(f"DEBUG: Processing column '{target_header}' mapped to '{scanned_column}'")
        
        if scanned_column not in input_headers:
            safe_print(f"DEBUG: Available input headers: {list(input_headers.keys())}")
            self.error_messages.append(
                f"{os.path.basename(self.input_file)}:{self.input_sheet_name}:: "
                f"mapped column '{scanned_column}' not found in input"
            )
            return
        
        input_col_idx = input_headers[scanned_column]
        
        # Get the maximum row with data in input file
        max_input_row = self.input_ws.max_row
        
        # Process each data row from input file
        rows_processed = 0
        for row_idx in range(2, max_input_row + 1):  # Start from row 2 (skip header)
            input_cell = self.input_ws.cell(row=row_idx, column=input_col_idx)
            target_cell = self.target_ws.cell(row=row_idx, column=target_col_idx)
            
            # Check if target cell has a formula
            if target_cell.data_type == 'f' and target_cell.value:
                # Target has formula - don't import data, but match data type
                self.copy_data_type_only(input_cell, target_cell)
            else:
                # Target has no formula - import data with type preservation
                self.copy_cell_value_with_type_preservation(input_cell, target_cell)
            
            rows_processed += 1
        
        safe_print(f"DEBUG: Processed {rows_processed} rows for column '{target_header}'")
    
    def clear_column_data(self, col_idx: int, max_row: int):
        """Clear data in a column (keeping formulas intact)"""
        cleared_cells = 0
        for row_idx in range(2, max_row + 1):  # Start from row 2 (keep header)
            cell = self.target_ws.cell(row=row_idx, column=col_idx)
            # Only clear if it's not a formula
            if cell.data_type != 'f':
                cell.value = None
                cleared_cells += 1
        
        if cleared_cells > 0:
            col_letter = get_column_letter(col_idx)
            safe_print(f"DEBUG: Cleared {cleared_cells} non-formula cells in column {col_letter}")
    
    def format_excel(self):
        """Main formatting function"""
        try:
            self.load_files()
            self.detect_column_formulas()
            
            input_headers = self.get_input_headers()
            target_headers = self.get_target_headers()
            
            safe_print(f"DEBUG: Found {len(input_headers)} input headers and {len(target_headers)} target headers")
            safe_print(f"DEBUG: Available mappings: {len(self.target_to_scanned)}")
            safe_print(f"DEBUG: Column formulas found: {len(self.column_formulas)}")
            
            # Calculate maximum data row from input
            max_input_data_row = self.input_ws.max_row
            
            # Extend target worksheet if needed
            max_target_row = max(self.target_ws.max_row, max_input_data_row)
            
            safe_print(f"DEBUG: Max input data row: {max_input_data_row}, Max target row: {max_target_row}")
            
            # First, clear all non-formula data in target (prepare for fresh data import)
            for header_name, col_idx in target_headers:
                if not header_name:  # Skip empty headers
                    continue
                
                col_letter = get_column_letter(col_idx)
                
                # Don't clear columns that have column-wide formulas
                if col_letter not in self.column_formulas:
                    self.clear_column_data(col_idx, max_target_row)
            
            # Process each target column
            processed_columns = 0
            mapped_columns = 0
            formula_columns = 0
            
            for header_name, col_idx in target_headers:
                if not header_name:  # Skip empty headers
                    continue
                    
                col_letter = get_column_letter(col_idx)
                processed_columns += 1
                
                safe_print(f"DEBUG: Processing target column '{header_name}' (Column {col_letter})")
                
                # Check if this column has a column-wide formula
                if col_letter in self.column_formulas:
                    formula_columns += 1
                    safe_print(f"DEBUG: Applying column-wide formula for {col_letter}")
                    self.process_column_with_formula(
                        col_idx, self.column_formulas[col_letter], max_input_data_row
                    )
                elif header_name in self.target_to_scanned:
                    # Column has mapping
                    mapped_columns += 1
                    scanned_column = self.target_to_scanned[header_name]
                    safe_print(f"DEBUG: Applying data mapping for {col_letter}: '{header_name}' -> '{scanned_column}'")
                    self.process_column_with_mapping(
                        col_idx, header_name, scanned_column, input_headers
                    )
                else:
                    # No mapping found
                    safe_print(f"DEBUG: No mapping found for '{header_name}'")
                    safe_print(f"DEBUG: Available mappings: {list(self.target_to_scanned.keys())}")
                    self.error_messages.append(
                        f"{os.path.basename(self.input_file)}:{self.input_sheet_name}:: "
                        f"no mapping for '{header_name}'"
                    )
            
            safe_print(f"DEBUG: Summary - Processed: {processed_columns}, Mapped: {mapped_columns}, Formula: {formula_columns}, Errors: {len(self.error_messages)}")
            
            # Handle errors
            if self.error_messages:
                for msg in self.error_messages:
                    safe_print(msg)
                
                # Copy to problematic directory
                problematic_dir = Path("data/problematic")
                problematic_dir.mkdir(parents=True, exist_ok=True)
                problematic_path = problematic_dir / os.path.basename(self.input_file)
                shutil.copy2(self.input_file, problematic_path)
                
                raise Exception("Formatting failed due to errors")
            
            # Save the result (save the modified target_wb as the result)
            self.target_wb.save(self.output_file)
            safe_print(f"Successfully formatted and saved to: {self.output_file}")
            
        except Exception as e:
            if self.error_messages:
                # Copy to problematic directory on error
                problematic_dir = Path("data/problematic")
                problematic_dir.mkdir(parents=True, exist_ok=True)
                problematic_path = problematic_dir / os.path.basename(self.input_file)
                shutil.copy2(self.input_file, problematic_path)
            raise e
        finally:
            # Clean up
            if self.input_wb:
                self.input_wb.close()
            if self.target_wb:
                self.target_wb.close()


def format_single_file(input_file: str, target_file: str, mapping_file: str, 
                      output_file: str, input_sheet: str, target_sheet: str, formula_row: int = 100):
    """Format a single Excel file"""
    formatter = ExcelFormatter(input_file, target_file, mapping_file, 
                              output_file, input_sheet, target_sheet, formula_row)
    formatter.format_excel()


def format_all_sheets(input_file: str, target_file: str, mapping_file: str, target_sheet: str, formula_row: int = 100):
    """Format all sheets in an Excel file"""
    # Validate input files
    for file_path in [input_file, target_file, mapping_file]:
        if not os.path.exists(file_path):
            raise Exception(f"File not found: {file_path}")
    
    # Create results directory
    results_dir = Path("data/results")
    results_dir.mkdir(parents=True, exist_ok=True)
    
    # Get all sheet names from input file
    input_wb = load_workbook(input_file, read_only=True)
    input_sheets = input_wb.sheetnames
    input_wb.close()
    
    if not input_sheets:
        raise Exception("No sheets found in input file")
    
    # Process each sheet
    input_filename = Path(input_file).stem
    
    for sheet_name in input_sheets:
        output_filename = f"{input_filename}-{sheet_name}.xlsx"
        output_file = results_dir / output_filename
        
        try:
            format_single_file(
                input_file, target_file, mapping_file,
                str(output_file), sheet_name, target_sheet, formula_row
            )
            safe_print(f"Format successful for {sheet_name}")
        except Exception as e:
            safe_print(f"Problematic file copied to: data/problematic/{os.path.basename(input_file)}")
            safe_print(f"Error for sheet {sheet_name}: {e}")


def main():
    """Main entry point for command line usage"""
    if len(sys.argv) < 5:
        safe_print("Usage: python format_excel.py <input_file> <target_file> <mapping_file> <target_sheet> [formula_row] [output_file] [input_sheet]")
        safe_print("  If output_file is not provided, will format all sheets")
        sys.exit(1)
    
    input_file = sys.argv[1]
    target_file = sys.argv[2]
    mapping_file = sys.argv[3]
    target_sheet = sys.argv[4]
    
    # Parse formula_row parameter
    formula_row = 100  # default
    if len(sys.argv) >= 6:
        try:
            formula_row = int(sys.argv[5])
        except ValueError:
            safe_print(f"Invalid formula_row value: {sys.argv[5]}, using default 100")
    
    if len(sys.argv) >= 8:
        # Single file format with formula_row
        output_file = sys.argv[6]
        input_sheet = sys.argv[7]
        format_single_file(input_file, target_file, mapping_file, 
                          output_file, input_sheet, target_sheet, formula_row)
    else:
        # Format all sheets
        format_all_sheets(input_file, target_file, mapping_file, target_sheet, formula_row)


if __name__ == "__main__":
    main()