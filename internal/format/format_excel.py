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
                 output_file: str, input_sheet: str, target_sheet: str, formula_row: int = 100, 
                 table_end_tolerance: int = 1, clean_formula_only_rows: bool = True):
        self.input_file = input_file
        self.target_file = target_file
        self.mapping_file = mapping_file
        self.output_file = output_file
        self.input_sheet_name = input_sheet
        self.target_sheet_name = target_sheet
        self.formula_row = formula_row
        self.table_end_tolerance = table_end_tolerance
        self.clean_formula_only_rows = clean_formula_only_rows
        
        self.input_wb = None
        self.target_wb = None
        self.input_ws = None
        self.target_ws = None
        
        self.mapping_config = None
        self.target_to_scanned = {}
        self.column_formulas = {}
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
                    #safe_print(f"DEBUG: Mapping '{target_col}' -> '{scanned_col}'")
            
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
            #safe_print(f"DEBUG: Row {self.formula_row} does not exist in target format, no column formulas to detect")
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
                    #safe_print(f"DEBUG: Found ArrayFormula in {col_letter}{self.formula_row}: {formula_text}")
                elif isinstance(cell.value, str):  # Regular formula string
                    formula_text = cell.value
                    #safe_print(f"DEBUG: Found regular formula in {col_letter}{self.formula_row}: {formula_text}")
                else:
                    # Try to convert to string
                    formula_text = str(cell.value)
                    #safe_print(f"DEBUG: Found formula object in {col_letter}{self.formula_row}, converted to: {formula_text}")
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
                #safe_print(f"Found column formula in {col_letter}{self.formula_row} ({header_value}): {formula_text}")
    
    def detect_table_end_row(self, worksheet, header_row: int) -> int:
        """Detect where the data table ends by checking rightmost column continuity with tolerance"""
        
        if worksheet.max_row == 0:
            return header_row  # No data, table ends at header
        
        # Find the rightmost column with data (same logic as header detection)
        rightmost_col = 0
        for row_idx in range(1, worksheet.max_row + 1):
            for col_idx in range(worksheet.max_column, 0, -1):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                if cell.value is not None and str(cell.value).strip() != "":
                    if col_idx > rightmost_col:
                        rightmost_col = col_idx
                    break
        
        if rightmost_col == 0:
            return header_row  # No data found
        
        safe_print(f"DEBUG: Using table end tolerance: {self.table_end_tolerance}")
        
        # Starting from header row, find where table ends
        table_end_row = header_row
        
        for row_idx in range(header_row, worksheet.max_row + 1):
            current_cell = worksheet.cell(row=row_idx, column=rightmost_col)
            current_has_data = current_cell.value is not None and str(current_cell.value).strip() != ""
            
            if current_has_data:
                # Check if next N rows (tolerance) have data in rightmost column
                empty_rows_count = 0
                
                for check_offset in range(1, self.table_end_tolerance + 1):
                    check_row_idx = row_idx + check_offset
                    if check_row_idx <= worksheet.max_row:
                        check_cell = worksheet.cell(row=check_row_idx, column=rightmost_col)
                        check_has_data = check_cell.value is not None and str(check_cell.value).strip() != ""
                        
                        if not check_has_data:
                            empty_rows_count += 1
                        else:
                            break  # Found data within tolerance, continue with main loop
                    else:
                        # Beyond worksheet bounds, count as empty
                        empty_rows_count += 1
                
                # If all tolerance rows are empty, current row is table end
                if empty_rows_count == self.table_end_tolerance:
                    table_end_row = row_idx
                    break
                else:
                    # Continue checking, update table end to current row
                    table_end_row = row_idx
        
        safe_print(f"DEBUG: Table end row detected: {table_end_row} (tolerance: {self.table_end_tolerance})")
        return table_end_row

    def clean_formula_only_rows_func(self):
        """Remove rows that only contain formulas and completely empty rows"""
        if not self.clean_formula_only_rows:
            safe_print("DEBUG: Formula-only row cleaning is disabled")
            return
        
        safe_print("DEBUG: Starting formula-only and empty row cleanup...")
        
        rows_to_delete = []
        
        # Check each row starting from row 1, but skip header row
        for row_idx in range(1, self.target_ws.max_row + 1):
            # Skip the header row
            if row_idx == self.target_header_row:
                continue
                
            has_non_formula_data = False
            
            # Check all cells in this row
            for col_idx in range(1, self.target_ws.max_column + 1):
                cell = self.target_ws.cell(row=row_idx, column=col_idx)
                
                if cell.value is not None:
                    # If cell contains data that's not a formula, keep the row
                    if cell.data_type != 'f':
                        # Check if it's not just whitespace
                        if str(cell.value).strip() != "":
                            has_non_formula_data = True
                            break
            
            # If row has no non-formula data (either empty or only formulas), mark for deletion
            if not has_non_formula_data:
                rows_to_delete.append(row_idx)
        
        # Delete rows in reverse order to maintain row indices
        deleted_count = 0
        for row_idx in reversed(rows_to_delete):
            #safe_print(f"DEBUG: Deleting empty/formula-only row: {row_idx}")
            self.target_ws.delete_rows(row_idx, 1)
            deleted_count += 1
        
        #safe_print(f"DEBUG: Cleaned up {deleted_count} empty/formula-only rows")

    def clear_rows_after(self, after_row: int):
        """Clear all rows after the specified row number"""
        if after_row >= self.target_ws.max_row:
            safe_print(f"DEBUG: No rows to clear after row {after_row} (max row is {self.target_ws.max_row})")
            return
        
        safe_print(f"DEBUG: Clearing all rows after row {after_row}")
        
        rows_to_delete = self.target_ws.max_row - after_row
        if rows_to_delete > 0:
            # Delete all rows from after_row+1 to max_row
            self.target_ws.delete_rows(after_row + 1, rows_to_delete)
            safe_print(f"DEBUG: Deleted {rows_to_delete} rows after row {after_row}")


    def get_input_headers(self) -> Dict[str, int]:
        """Get input file headers mapping"""
        headers = {}
        
        # Detect header row
        input_header_row = self.detect_header_row(self.input_ws)
        safe_print(f"DEBUG: Input header row detected: {input_header_row}")
        
        # Detect table end row
        self.input_table_end_row = self.detect_table_end_row(self.input_ws, input_header_row)
        safe_print(f"DEBUG: Input table end row detected: {self.input_table_end_row}")
        
        #safe_print("DEBUG: Input file headers:")
        for col_idx in range(1, self.input_ws.max_column + 1):
            cell = self.input_ws.cell(row=input_header_row, column=col_idx)
            if cell.value:
                header_value = str(cell.value).strip()
                headers[header_value] = col_idx
                #safe_print(f"  Column {col_idx}: '{header_value}'")
        
        # Store the header row for later use
        self.input_header_row = input_header_row
        return headers
    
    def get_target_headers(self) -> List[Tuple[str, int]]:
        """Get target format headers with their column indices"""
        headers = []
        
        # Detect header row
        target_header_row = self.detect_header_row(self.target_ws)
        #safe_print(f"DEBUG: Target header row detected: {target_header_row}")
        
        #safe_print("DEBUG: Target format headers:")
        for col_idx in range(1, self.target_ws.max_column + 1):
            cell = self.target_ws.cell(row=target_header_row, column=col_idx)
            header_value = str(cell.value).strip() if cell.value else ""
            headers.append((header_value, col_idx))
            #safe_print(f"  Column {col_idx}: '{header_value}'")
        
        # Store the header row for later use
        self.target_header_row = target_header_row
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
        """Adjust formula template by appending row number to column letters"""
        import re
        
        # Convert to string if it's not already
        if not isinstance(formula, str):
            formula = str(formula)
        
        # Ensure formula starts with =
        formula = formula.strip()
        if not formula.startswith('='):
            formula = '=' + formula
        
        # Find all column references (A-Z, AA-ZZ, etc.) and append row number
        # This regex matches one or more uppercase letters that represent Excel columns
        def replace_column(match):
            column_letter = match.group(0)
            return column_letter + str(target_row)
        
        # Pattern to match Excel column letters (A, B, AA, AB, etc.)
        # Only matches when not already followed by a number
        pattern = r'[A-Z]+(?!\d)'
        adjusted = re.sub(pattern, replace_column, formula)
        
        ##safe_print(f"DEBUG: Adjusted formula from '{formula}' to '{adjusted}' for row {target_row}")
        return adjusted
    
    def process_column_with_formula(self, col_idx: int, formula_template: str, max_data_row: int):
        """Apply column-wide formula to all data rows"""
        col_letter = get_column_letter(col_idx)
        
        ##safe_print(f"DEBUG: Processing column {col_letter} with formula template: {formula_template}")
        
        formulas_applied = 0
        target_data_start_row = self.target_header_row + 1
        
        # Calculate how many data rows we have from input
        input_data_rows = max_data_row - self.input_header_row
        
        for row_offset in range(input_data_rows):
            target_row_idx = target_data_start_row + row_offset
            target_cell = self.target_ws.cell(row=target_row_idx, column=col_idx)
            adjusted_formula = self.adjust_formula_for_row(formula_template, target_row_idx)
            target_cell.value = adjusted_formula
            formulas_applied += 1
        
        ##safe_print(f"Applied column formula to {col_letter}: {formula_template} ({formulas_applied} cells)")
    
    def process_column_with_mapping(self, target_col_idx: int, target_header: str, 
                               scanned_column: str, input_headers: Dict[str, int]):
        """Process a column that has data mapping"""
        ###safe_print(f"DEBUG: Processing column '{target_header}' mapped to '{scanned_column}'")
        
        if scanned_column not in input_headers:
            #safe_print(f"DEBUG: Available input headers: {list(input_headers.keys())}")
            self.error_messages.append(
                f"{os.path.basename(self.input_file)}:{self.input_sheet_name}:: "
                f"mapped column '{scanned_column}' not found in input"
            )
            return
        
        input_col_idx = input_headers[scanned_column]
        
        # Get the maximum row with data in input file
        max_input_row = self.input_ws.max_row
        
        # Process each data row from input file (starting after header row)
        rows_processed = 0
        input_data_start_row = self.input_header_row + 1
        target_data_start_row = self.target_header_row + 1
        
        for input_row_idx in range(input_data_start_row, max_input_row + 1):
            # Calculate corresponding target row
            target_row_idx = target_data_start_row + (input_row_idx - input_data_start_row)
            
            input_cell = self.input_ws.cell(row=input_row_idx, column=input_col_idx)
            target_cell = self.target_ws.cell(row=target_row_idx, column=target_col_idx)
            
            # Check if target cell has a formula
            if target_cell.data_type == 'f' and target_cell.value:
                # Target has formula - don't import data, but match data type
                self.copy_data_type_only(input_cell, target_cell)
            else:
                # Target has no formula - import data with type preservation
                self.copy_cell_value_with_type_preservation(input_cell, target_cell)
            
            rows_processed += 1
        
        #safe_print(f"DEBUG: Processed {rows_processed} rows for column '{target_header}'")
    
    def clear_column_data(self, col_idx: int, max_row: int):
        """Clear data in a column (keeping formulas intact)"""
        cleared_cells = 0
        target_data_start_row = self.target_header_row + 1
        
        for row_idx in range(target_data_start_row, max_row + 1):
            cell = self.target_ws.cell(row=row_idx, column=col_idx)
            # Only clear if it's not a formula
            if cell.data_type != 'f':
                cell.value = None
                cleared_cells += 1
        
        if cleared_cells > 0:
            col_letter = get_column_letter(col_idx)
            #safe_print(f"DEBUG: Cleared {cleared_cells} non-formula cells in column {col_letter}")

    def detect_header_row(self, worksheet) -> int:
        """Detect header row using the strategy: find rightmost column with data, 
        then find first row with data in that column"""
        
        if worksheet.max_row == 0:
            return 1  # Default to row 1 if no data
        
        # Find the rightmost column with data across all rows
        rightmost_col = 0
        for row_idx in range(1, worksheet.max_row + 1):
            for col_idx in range(worksheet.max_column, 0, -1):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                if cell.value is not None and str(cell.value).strip() != "":
                    if col_idx > rightmost_col:
                        rightmost_col = col_idx
                    break  # Found the rightmost data in this row
        
        if rightmost_col == 0:
            return 1  # Default to row 1 if no data
        
        # Now find the first row that has data in the rightmost column
        for row_idx in range(1, worksheet.max_row + 1):
            cell = worksheet.cell(row=row_idx, column=rightmost_col)
            if cell.value is not None and str(cell.value).strip() != "":
                return row_idx
        
        return 1  # Default to row 1 if not found

    def adjust_formula_for_row(self, formula: str, target_row: int) -> str:
        """Adjust formula template by appending row number to column letters"""
        import re
        
        # Convert to string if it's not already
        if not isinstance(formula, str):
            formula = str(formula)
        
        # Ensure formula starts with =
        formula = formula.strip()
        if not formula.startswith('='):
            formula = '=' + formula
        
        # Find all column references (A-Z, AA-ZZ, etc.) and append row number
        # Pattern to match Excel column letters not already followed by a number
        pattern = r'[A-Z]+(?!\d)'
        adjusted = re.sub(pattern, lambda m: m.group(0) + str(target_row), formula)
        #safe_print(f"Adjusted formula for row {target_row} is {adjusted}")
        return adjusted

    def apply_column_formulas_at_end(self, max_data_rows: int):
        """Apply all column-wide formulas to data rows at the end"""
        if not self.column_formulas:
            safe_print("DEBUG: No column formulas to apply")
            return
        
        target_data_start_row = self.target_header_row + 1
        
        for col_letter, formula_template in self.column_formulas.items():
            col_idx = column_index_from_string(col_letter)
            
            for row_offset in range(max_data_rows):
                actual_target_row = target_data_start_row + row_offset
                target_cell = self.target_ws.cell(row=actual_target_row, column=col_idx)
                adjusted_formula = self.adjust_formula_for_row(formula_template, actual_target_row)
                target_cell.value = adjusted_formula
            
            safe_print(f"Applied formulas to column {col_letter}")

    def format_excel(self):
        """Main formatting function"""
        try:
            self.load_files()
            
            # Initialize header row variables
            self.input_header_row = 1
            self.target_header_row = 1
            
            input_headers = self.get_input_headers()
            target_headers = self.get_target_headers()
            
            # Detect column formulas after we know the target header row
            self.detect_column_formulas()
            
            safe_print(f"DEBUG: Found {len(input_headers)} input headers and {len(target_headers)} target headers")
            safe_print(f"DEBUG: Available mappings: {len(self.target_to_scanned)}")
            safe_print(f"DEBUG: Column formulas found: {len(self.column_formulas)}")
            
            # Calculate maximum data row from input file
            max_input_data_row = self.input_ws.max_row
            
            # Calculate how many data rows we need in target
            input_data_rows = max_input_data_row - self.input_header_row
            target_max_needed_row = self.target_header_row + input_data_rows
            
            # Extend target worksheet if needed
            max_target_row = max(self.target_ws.max_row, target_max_needed_row)
            
            safe_print(f"DEBUG: Input header row: {self.input_header_row}, Target header row: {self.target_header_row}")
            safe_print(f"DEBUG: Max input data row: {max_input_data_row}, Max target row needed: {target_max_needed_row}")
            
            # First, clear all non-formula data in target (prepare for fresh data import)
            for header_name, col_idx in target_headers:
                if not header_name:  # Skip empty headers
                    continue
                
                col_letter = get_column_letter(col_idx)
                
                # Clear all data columns (we'll apply formulas at the end)
                self.clear_column_data(col_idx, max_target_row)
            
            # Process each target column (only data mapping, no formulas yet)
            processed_columns = 0
            mapped_columns = 0
            
            for header_name, col_idx in target_headers:
                if not header_name:  # Skip empty headers
                    continue
                    
                col_letter = get_column_letter(col_idx)
                processed_columns += 1
                
                safe_print(f"DEBUG: Processing target column '{header_name}' (Column {col_letter})")
                
                # Only process data mapping, skip formula columns for now
                if col_letter in self.column_formulas:
                    safe_print(f"DEBUG: Skipping formula column {col_letter} - will apply formulas after cleanup")
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
            
            safe_print(f"DEBUG: Data mapping complete - Processed: {processed_columns}, Mapped: {mapped_columns}")
            
            # Clean formula-only and empty rows within the data first
            self.clean_formula_only_rows_func()
            
            # NOW detect the table end row after importing data
            actual_data_end_row = self.detect_table_end_row(self.target_ws, self.target_header_row)
            safe_print(f"DEBUG: Actual imported data ends at row: {actual_data_end_row}")
            
            # Clear everything after the actual data
            self.clear_rows_after(actual_data_end_row)
            
            # Calculate how many actual data rows we have
            actual_data_rows = actual_data_end_row - self.target_header_row
            safe_print(f"DEBUG: Final data rows count: {actual_data_rows}")
            
            # Now apply column formulas to the cleaned-up rows
            safe_print("DEBUG: Applying column-wide formulas after cleanup...")
            self.apply_column_formulas_at_end(actual_data_rows)

            self.target_ws.cell(row=actual_data_end_row + 4, column=1).value = "*Fiyatlandırma ve promosyon akitiviteleri ile ilgili nihai karar müşterinindir"
            self.target_ws.cell(row=actual_data_end_row + 5, column=1).value = "*P&G'nin müşteri karlılığı üzerinde herhangi bir belirleyiciliği olmayıp, iligli kar marjı sütunları tavsiye kar marjlarıdır, nihai karar müşterinindir"
            
            safe_print(f"DEBUG: Summary - Processed: {processed_columns}, Mapped: {mapped_columns}, Formula columns: {len(self.column_formulas)}, Errors: {len(self.error_messages)}")
            
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
                      output_file: str, input_sheet: str, target_sheet: str, formula_row: int = 100, 
                      table_end_tolerance: int = 1, clean_formula_only_rows: bool = True):
    formatter = ExcelFormatter(input_file, target_file, mapping_file, 
                              output_file, input_sheet, target_sheet, formula_row, table_end_tolerance, clean_formula_only_rows)
    formatter.format_excel()


def format_all_sheets(input_file: str, target_file: str, mapping_file: str, target_sheet: str, 
                     formula_row: int = 100, table_end_tolerance: int = 1, clean_formula_only_rows: bool = True):
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
                str(output_file), sheet_name, target_sheet, formula_row, table_end_tolerance, clean_formula_only_rows
            )
            safe_print(f"Format successful for {sheet_name}")
        except Exception as e:
            safe_print(f"Problematic file copied to: data/problematic/{os.path.basename(input_file)}")
            safe_print(f"Error for sheet {sheet_name}: {e}")

def main():
    """Main entry point for command line usage"""
    if len(sys.argv) < 5:
        safe_print("Usage: python format_excel.py <input_file> <target_file> <mapping_file> <target_sheet> [formula_row] [table_end_tolerance] [output_file] [input_sheet]")
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
    
    # Parse table_end_tolerance parameter
    table_end_tolerance = 1  # default
    if len(sys.argv) >= 7:
        try:
            table_end_tolerance = int(sys.argv[6])
        except ValueError:
            safe_print(f"Invalid table_end_tolerance value: {sys.argv[6]}, using default 1")
    
    clean_formula_only_rows = True  # default
    if len(sys.argv) >= 8:
        clean_formula_only_rows = sys.argv[7].lower() == "true"
    if len(sys.argv) >= 10:
        # Single file format
        output_file = sys.argv[8]
        input_sheet = sys.argv[9]
        format_single_file(input_file, target_file, mapping_file, 
                          output_file, input_sheet, target_sheet, formula_row, table_end_tolerance, clean_formula_only_rows)
    else:
        # Format all sheets
        format_all_sheets(input_file, target_file, mapping_file, target_sheet, formula_row, table_end_tolerance, clean_formula_only_rows)


if __name__ == "__main__":
    main()