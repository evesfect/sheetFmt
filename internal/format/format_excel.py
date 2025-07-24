#!/usr/bin/env python3
import json
import sys
import os
from pathlib import Path
from typing import Dict, List, Tuple
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import shutil
import logging
from datetime import datetime


def setup_logging():
    """Setup logging configuration with Unicode support"""
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    
    # Create file handler with UTF-8 encoding
    file_handler = logging.FileHandler(
        log_dir / "format_excel.log", 
        encoding='utf-8'
    )
    file_handler.setLevel(logging.DEBUG)
    
    # Create console handler 
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    
    # Create formatters
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    console_formatter = logging.Formatter('%(levelname)s - %(message)s')
    
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(console_formatter)
    
    # Setup logger
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger


logger = setup_logging()


def safe_print(text):
    """Print text with encoding error handling"""
    try:
        print(text)
    except UnicodeEncodeError:
        safe_text = text.encode('ascii', 'replace').decode('ascii')
        print(safe_text)


class ExcelFormatter:
    def __init__(self, input_file: str, target_file: str, mapping_file: str, 
                 output_file: str, input_sheet: str, target_sheet: str, 
                 table_end_tolerance: int = 1, clean_formula_only_rows: bool = True):
        self.input_file = input_file
        self.target_file = target_file
        self.mapping_file = mapping_file
        self.output_file = output_file
        self.input_sheet_name = input_sheet
        self.target_sheet_name = target_sheet
        self.table_end_tolerance = table_end_tolerance
        self.clean_formula_only_rows = clean_formula_only_rows
        
        self.input_wb = None
        self.target_wb = None
        self.input_ws = None
        self.target_ws = None
        
        self.mapping_config = None
        self.scanned_to_target = {}  # Changed: scanned -> target
        self.ignored_scanned = set()  # Changed: track ignored scanned columns
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
            
            logger.debug(f"Loaded mapping file: {self.mapping_file}")
            logger.debug(f"Mapping config keys: {list(self.mapping_config.keys())}")
            
            # Create scanned -> target mapping and ignored set
            mappings_count = 0
            ignored_count = 0
            
            for mapping in self.mapping_config.get('mappings', []):
                scanned_col = mapping['scanned_column']
                
                if mapping.get('is_ignored', False):
                    # Track ignored scanned columns
                    self.ignored_scanned.add(scanned_col)
                    ignored_count += 1
                    logger.debug(f"Ignoring scanned column '{scanned_col}'")
                elif mapping.get('target_column'):
                    target_col = mapping['target_column']
                    # Changed: Now we map scanned -> target (can have multiple scanned for same target)
                    self.scanned_to_target[scanned_col] = target_col
                    mappings_count += 1
                    logger.debug(f"Mapping scanned '{scanned_col}' -> target '{target_col}'")
            
            logger.debug(f"Total mappings loaded: {mappings_count}")
            logger.debug(f"Total ignored: {ignored_count}")
            
            # Load Excel files
            self.input_wb = load_workbook(self.input_file, data_only=False)
            self.target_wb = load_workbook(self.target_file, data_only=False)
            
            # Get worksheets
            logger.debug(f"Input sheets: {self.input_wb.sheetnames}")
            logger.debug(f"Target sheets: {self.target_wb.sheetnames}")
            
            self.input_ws = self.input_wb[self.input_sheet_name]
            self.target_ws = self.target_wb[self.target_sheet_name]
            
        except Exception as e:
            logger.error(f"Failed to load files: {e}")
            raise Exception(f"Failed to load files: {e}")

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
        
        logger.debug(f"Using table end tolerance: {self.table_end_tolerance}")
        
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
        
        logger.debug(f"Table end row detected: {table_end_row} (tolerance: {self.table_end_tolerance})")
        return table_end_row

    def clean_formula_only_rows_func(self):
        """Remove rows that only contain formulas and completely empty rows"""
        if not self.clean_formula_only_rows:
            logger.debug("Formula-only row cleaning is disabled")
            return
        
        logger.debug("Starting formula-only and empty row cleanup...")
        
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
            self.target_ws.delete_rows(row_idx, 1)
            deleted_count += 1
        
        logger.debug(f"Cleaned up {deleted_count} empty/formula-only rows")

    def get_input_headers(self) -> Dict[str, int]:
        """Get input file headers mapping"""
        headers = {}
        
        # Detect header row
        input_header_row = self.detect_header_row(self.input_ws)
        logger.debug(f"Input header row detected: {input_header_row}")
        
        # Detect table end row
        self.input_table_end_row = self.detect_table_end_row(self.input_ws, input_header_row)
        logger.debug(f"Input table end row detected: {self.input_table_end_row}")
        
        logger.debug("Input file headers:")
        for col_idx in range(1, self.input_ws.max_column + 1):
            cell = self.input_ws.cell(row=input_header_row, column=col_idx)
            if cell.value:
                header_value = str(cell.value).strip()
                headers[header_value] = col_idx
                logger.debug(f"  Column {col_idx}: '{header_value}'")
        
        # Store the header row for later use
        self.input_header_row = input_header_row
        return headers
    
    def get_target_headers(self) -> List[Tuple[str, int]]:
        """Get target format headers with their column indices"""
        headers = []
        
        # Detect header row
        target_header_row = self.detect_header_row(self.target_ws)
        logger.debug(f"Target header row detected: {target_header_row}")
        
        logger.debug("Target format headers:")
        for col_idx in range(1, self.target_ws.max_column + 1):
            cell = self.target_ws.cell(row=target_header_row, column=col_idx)
            header_value = str(cell.value).strip() if cell.value else ""
            headers.append((header_value, col_idx))
            logger.debug(f"  Column {col_idx}: '{header_value}'")
        
        # Store the header row for later use
        self.target_header_row = target_header_row
        return headers
    
    def find_applicable_mappings(self, input_headers: Dict[str, int], target_headers: List[Tuple[str, int]]) -> Dict[str, str]:
        """Find which input columns can be mapped to which target columns based on available data"""
        applicable_mappings = {}  # target_column -> input_column
        
        logger.debug("=== FINDING APPLICABLE MAPPINGS ===")
        logger.debug(f"Available input columns: {list(input_headers.keys())}")
        
        # Get all target column names
        target_column_names = [header for header, _ in target_headers if header]
        logger.debug(f"Target columns needing data: {target_column_names}")
        
        # For each target column, find if any available input column maps to it
        for target_col in target_column_names:
            found_mapping = False
            
            # Look through all available input columns
            for input_col in input_headers.keys():
                # Check if this input column is ignored
                if input_col in self.ignored_scanned:
                    logger.debug(f"  Skipping ignored input column '{input_col}'")
                    continue
                
                # Check if this input column has a mapping to our target column
                if input_col in self.scanned_to_target and self.scanned_to_target[input_col] == target_col:
                    applicable_mappings[target_col] = input_col
                    logger.debug(f"  ✓ APPLICABLE: Target '{target_col}' <- Input '{input_col}' (found in file)")
                    found_mapping = True
                    break  # Use first available mapping for this target
            
            if not found_mapping:
                logger.debug(f"  ✗ NO MAPPING: Target '{target_col}' - no available input column maps to it")
        
        logger.debug(f"=== FINAL APPLICABLE MAPPINGS: {len(applicable_mappings)} ===")
        for target_col, input_col in applicable_mappings.items():
            logger.debug(f"  '{target_col}' <- '{input_col}'")
        
        return applicable_mappings
    
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
    
    def process_column_with_mapping(self, target_col_idx: int, target_header: str, 
                               input_column: str, input_headers: Dict[str, int]):
        """Process a column that has data mapping"""
        logger.debug(f"Processing column '{target_header}' mapped from '{input_column}'")
        
        input_col_idx = input_headers[input_column]
        logger.debug(f"Input column '{input_column}' is at index {input_col_idx}")
        
        # Get the maximum row with data in input file
        max_input_row = self.input_ws.max_row
        
        # Process each data row from input file (starting after header row)
        rows_processed = 0
        cells_copied = 0
        input_data_start_row = self.input_header_row + 1
        target_data_start_row = self.target_header_row + 1
        
        logger.debug(f"Processing rows from input row {input_data_start_row} to {max_input_row}")
        logger.debug(f"Mapping to target starting at row {target_data_start_row}")
        
        for input_row_idx in range(input_data_start_row, max_input_row + 1):
            # Calculate corresponding target row
            target_row_idx = target_data_start_row + (input_row_idx - input_data_start_row)
            
            input_cell = self.input_ws.cell(row=input_row_idx, column=input_col_idx)
            target_cell = self.target_ws.cell(row=target_row_idx, column=target_col_idx)
            
            # Log what we're copying
            input_value = input_cell.value
            if input_value is not None and str(input_value).strip() != "":
                logger.debug(f"Copying from input[{input_row_idx},{input_col_idx}] = '{input_value}' to target[{target_row_idx},{target_col_idx}]")
                cells_copied += 1
            
            # Copy data
            self.copy_cell_value_with_type_preservation(input_cell, target_cell)
            rows_processed += 1
        
        logger.debug(f"Processed {rows_processed} rows for column '{target_header}', copied {cells_copied} non-empty cells")

    def clear_column_data(self, col_idx: int, max_row: int):
        """Clear data in a column"""
        cleared_cells = 0
        target_data_start_row = self.target_header_row + 1
        
        for row_idx in range(target_data_start_row, max_row + 1):
            cell = self.target_ws.cell(row=row_idx, column=col_idx)
            cell.value = None
            cleared_cells += 1
        
        if cleared_cells > 0:
            col_letter = get_column_letter(col_idx)
            logger.debug(f"Cleared {cleared_cells} cells in column {col_letter}")

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

    def format_excel(self):
        """Main formatting function"""
        try:
            logger.info(f"Starting Excel formatting for {os.path.basename(self.input_file)}")
            self.load_files()
            
            # Initialize header row variables
            self.input_header_row = 1
            self.target_header_row = 1
            
            input_headers = self.get_input_headers()
            target_headers = self.get_target_headers()
            
            logger.debug(f"Found {len(input_headers)} input headers and {len(target_headers)} target headers")
            logger.debug(f"Total configured mappings: {len(self.scanned_to_target)}")
            
            # NEW: Find applicable mappings based on available input columns
            applicable_mappings = self.find_applicable_mappings(input_headers, target_headers)
            logger.debug(f"Applicable mappings for this file: {len(applicable_mappings)}")
            
            # Calculate maximum data row from input file
            max_input_data_row = self.input_ws.max_row
            
            # Calculate how many data rows we need in target
            input_data_rows = max_input_data_row - self.input_header_row
            target_max_needed_row = self.target_header_row + input_data_rows
            
            # Extend target worksheet if needed
            max_target_row = max(self.target_ws.max_row, target_max_needed_row)
            
            logger.debug(f"Input header row: {self.input_header_row}, Target header row: {self.target_header_row}")
            logger.debug(f"Max input data row: {max_input_data_row}, Max target row needed: {target_max_needed_row}")
            
            # First, clear all data in target (prepare for fresh data import)
            for header_name, col_idx in target_headers:
                if not header_name:  # Skip empty headers
                    continue
                self.clear_column_data(col_idx, max_target_row)
            
            # Process each target column using applicable mappings
            processed_columns = 0
            mapped_columns = 0
            skipped_columns = 0

            for header_name, col_idx in target_headers:
                if not header_name:  # Skip empty headers
                    continue
                    
                col_letter = get_column_letter(col_idx)
                processed_columns += 1
                
                logger.debug(f"Processing target column '{header_name}' (Column {col_letter})")
                
                if header_name in applicable_mappings:
                    # Column has applicable mapping
                    mapped_columns += 1
                    input_column = applicable_mappings[header_name]
                    logger.debug(f"Applying data mapping for {col_letter}: '{header_name}' <- '{input_column}'")
                    self.process_column_with_mapping(
                        col_idx, header_name, input_column, input_headers
                    )
                else:
                    # No applicable mapping found
                    logger.debug(f"No applicable mapping found for '{header_name}' - leaving empty")
                    skipped_columns += 1

            logger.debug(f"Data mapping complete - Processed: {processed_columns}, Mapped: {mapped_columns}, Skipped: {skipped_columns}")
            
            # Clean formula-only and empty rows
            self.clean_formula_only_rows_func()
            
            logger.debug(f"Summary - Processed: {processed_columns}, Mapped: {mapped_columns}, Errors: {len(self.error_messages)}")
            
            # Handle errors
            if self.error_messages:
                for msg in self.error_messages:
                    logger.error(msg)
                
                # Copy to problematic directory
                problematic_dir = Path("data/problematic")
                problematic_dir.mkdir(parents=True, exist_ok=True)
                problematic_path = problematic_dir / os.path.basename(self.input_file)
                shutil.copy2(self.input_file, problematic_path)
                
                raise Exception("Formatting failed due to errors")
            
            # Save the result
            self.target_wb.save(self.output_file)
            logger.info(f"Successfully formatted and saved to: {self.output_file}")
            safe_print(f"Successfully formatted and saved to: {self.output_file}")
            
        except Exception as e:
            logger.error(f"Formatting failed: {e}")
            if self.error_messages:
                # Copy to problematic directory on error
                problematic_dir = Path("data/problematic")
                problematic_dir.mkdir(parents=True, exist_ok=True)
                problematic_path = problematic_dir / os.path.basename(self.input_file)
                shutil.copy2(self.input_file, problematic_path)
                logger.info(f"Copied problematic file to: {problematic_path}")
            raise e
        finally:
            # Clean up
            if self.input_wb:
                self.input_wb.close()
            if self.target_wb:
                self.target_wb.close()


def format_single_file(input_file: str, target_file: str, mapping_file: str, 
                      output_file: str, input_sheet: str, target_sheet: str, 
                      table_end_tolerance: int = 1, clean_formula_only_rows: bool = True):
    formatter = ExcelFormatter(input_file, target_file, mapping_file, 
                              output_file, input_sheet, target_sheet, 
                              table_end_tolerance, clean_formula_only_rows)
    formatter.format_excel()


def format_all_sheets(input_file: str, target_file: str, mapping_file: str, target_sheet: str, 
                     table_end_tolerance: int = 1, clean_formula_only_rows: bool = True):
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
    
    logger.info(f"Processing {len(input_sheets)} sheets from {os.path.basename(input_file)}")
    
    # Process each sheet
    input_filename = Path(input_file).stem
    
    for sheet_name in input_sheets:
        output_filename = f"{input_filename}-{sheet_name}.xlsx"
        output_file = results_dir / output_filename
        
        try:
            logger.info(f"Processing sheet: {sheet_name}")
            format_single_file(
                input_file, target_file, mapping_file,
                str(output_file), sheet_name, target_sheet, 
                table_end_tolerance, clean_formula_only_rows
            )
            safe_print(f"✓ Format successful for sheet: {sheet_name}")
        except Exception as e:
            logger.error(f"Error processing sheet {sheet_name}: {e}")
            safe_print(f"❌ Error for sheet {sheet_name}: {e}")
            safe_print(f"Problematic file copied to: data/problematic/{os.path.basename(input_file)}")

def main():
    """Main entry point for command line usage"""
    if len(sys.argv) < 5:
        safe_print("Usage: python format_excel.py <input_file> <target_file> <mapping_file> <target_sheet> [table_end_tolerance] [clean_formula_only_rows] [output_file] [input_sheet]")
        safe_print("  If output_file is not provided, will format all sheets")
        sys.exit(1)
    
    input_file = sys.argv[1]
    target_file = sys.argv[2]
    mapping_file = sys.argv[3]
    target_sheet = sys.argv[4]
    
    # Parse table_end_tolerance parameter
    table_end_tolerance = 1  # default
    if len(sys.argv) >= 6:
        try:
            table_end_tolerance = int(sys.argv[5])
        except ValueError:
            safe_print(f"Invalid table_end_tolerance value: {sys.argv[5]}, using default 1")
    
    # Parse clean_formula_only_rows parameter
    clean_formula_only_rows = True  # default
    if len(sys.argv) >= 7:
        clean_formula_only_rows = sys.argv[6].lower() == "true"
    
    if len(sys.argv) >= 9:
        # Single file format
        output_file = sys.argv[7]
        input_sheet = sys.argv[8]
        format_single_file(input_file, target_file, mapping_file, 
                          output_file, input_sheet, target_sheet, 
                          table_end_tolerance, clean_formula_only_rows)
    else:
        # Format all sheets
        format_all_sheets(input_file, target_file, mapping_file, target_sheet, 
                         table_end_tolerance, clean_formula_only_rows)


if __name__ == "__main__":
    main()