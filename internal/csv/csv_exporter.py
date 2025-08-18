#!/usr/bin/env python3
"""
CSV Exporter
Main script for exporting formatted Excel files to CSV format.
"""

import sys
import os
import csv
import logging
from pathlib import Path
from typing import List, Dict, Any, Set
import glob

from csv_config import load_csv_config
from picklist_handler import PicklistHandler
from excel_reader import ExcelReader

def setup_logging():
    """Setup logging configuration"""
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    
    # Create file handler
    file_handler = logging.FileHandler(
        log_dir / "csv_export.log", 
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
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

class CSVExporter:
    """Main CSV export orchestrator"""
    
    def __init__(self, input_dir: str):
        self.input_dir = Path(input_dir)
        self.config = load_csv_config()
        self.picklist_handler = PicklistHandler(self.config)
        self.excel_reader = ExcelReader(self.config, self.picklist_handler)
        self.logger = logging.getLogger(__name__)
        
        # Track processing statistics
        self.stats = {
            'files_processed': 0,
            'files_failed': 0,
            'salesplans_created': 0,
            'lineitems_created': 0,
            'duplicates_skipped': 0,
            'unmapped_values': 0
        }
    
    def export_to_csv(self):
        """Main export function"""
        self.logger.info("Starting CSV export process")
        
        # Validate input directory
        if not self.input_dir.exists():
            raise Exception(f"Input directory not found: {self.input_dir}")
        
        # Get all Excel files
        excel_files = self.get_excel_files()
        if not excel_files:
            self.logger.warning(f"No Excel files found in {self.input_dir}")
            return
        
        self.logger.info(f"Found {len(excel_files)} Excel files to process")
        
        # Load existing CSV data for deduplication
        existing_salesplans = self.load_existing_salesplans()
        existing_lineitems = self.load_existing_lineitems()
        
        # Process each Excel file
        for excel_file in excel_files:
            try:
                self.process_excel_file(excel_file, existing_salesplans, existing_lineitems)
                self.stats['files_processed'] += 1
            except Exception as e:
                self.logger.error(f"Failed to process {excel_file.name}: {e}")
                self.stats['files_failed'] += 1
        
        # Update statistics
        self.stats['unmapped_values'] = self.picklist_handler.get_unmapped_count()
        
        # Print summary
        self.print_summary()
    
    def get_excel_files(self) -> List[Path]:
        """Get list of Excel files to process"""
        excel_files = []
        
        # Look for .xlsx files
        for pattern in ["*.xlsx", "*.xls"]:
            files = list(self.input_dir.glob(pattern))
            excel_files.extend(files)
        
        # Filter out temporary files
        excel_files = [f for f in excel_files if not f.name.startswith('~$')]
        
        # Sort by name for consistent processing
        excel_files.sort(key=lambda x: x.name)
        
        return excel_files
    
    def process_excel_file(self, excel_file: Path, existing_salesplans: Set[str], 
                          existing_lineitems: Set[str]):
        """
        Process a single Excel file
        
        Args:
            excel_file: Path to Excel file
            existing_salesplans: Set of existing SalesPlans keys for deduplication
            existing_lineitems: Set of existing LineItems keys for deduplication
        """
        self.logger.info(f"Processing: {excel_file.name}")
        
        # Extract SalesPlans metadata
        salesplan_data = self.excel_reader.extract_salesplan_metadata(str(excel_file))
        
        # Check for SalesPlans duplicate
        salesplan_key = f"{salesplan_data['SourceFile']}|{salesplan_data['ActivityType']}"
        if salesplan_key in existing_salesplans:
            self.logger.info(f"Skipping duplicate SalesPlans: {salesplan_key}")
            self.stats['duplicates_skipped'] += 1
            return
        
        # Extract LineItems data
        lineitems_data = self.excel_reader.extract_lineitems_data(
            str(excel_file), salesplan_data['SalesPlanID']
        )
        
        if not lineitems_data:
            self.logger.warning(f"No line items found in {excel_file.name}")
            return
        
        # Filter out duplicate LineItems
        new_lineitems = []
        for lineitem in lineitems_data:
            # Create lineitem key for deduplication
            lineitem_key = f"{lineitem['SalesPlanID']}|{lineitem.get('ProductBarcode', '')}"
            if lineitem_key not in existing_lineitems:
                new_lineitems.append(lineitem)
                existing_lineitems.add(lineitem_key)
            else:
                self.stats['duplicates_skipped'] += 1
        
        if not new_lineitems:
            self.logger.warning(f"All line items were duplicates for {excel_file.name}")
            return
        
        # Write to CSV files
        self.append_salesplan_to_csv(salesplan_data)
        self.append_lineitems_to_csv(new_lineitems)
        
        # Update tracking sets
        existing_salesplans.add(salesplan_key)
        
        # Update statistics
        self.stats['salesplans_created'] += 1
        self.stats['lineitems_created'] += len(new_lineitems)
        
        self.logger.info(f"Successfully processed {excel_file.name}: "
                        f"1 SalesPlans, {len(new_lineitems)} LineItems")
    
    def load_existing_salesplans(self) -> Set[str]:
        """Load existing SalesPlans for deduplication"""
        existing = set()
        salesplans_file = self.config.get_output_dir() / "SalesPlans.csv"
        
        if not salesplans_file.exists():
            self.logger.debug("No existing SalesPlans.csv found")
            return existing
        
        try:
            with open(salesplans_file, 'r', encoding='utf-8', newline='') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    key = f"{row.get('SourceFile', '')}|{row.get('ActivityType', '')}"
                    existing.add(key)
            
            self.logger.info(f"Loaded {len(existing)} existing SalesPlans for deduplication")
        except Exception as e:
            self.logger.error(f"Failed to load existing SalesPlans: {e}")
        
        return existing
    
    def load_existing_lineitems(self) -> Set[str]:
        """Load existing LineItems for deduplication"""
        existing = set()
        lineitems_file = self.config.get_output_dir() / "LineItems.csv"
        
        if not lineitems_file.exists():
            self.logger.debug("No existing LineItems.csv found")
            return existing
        
        try:
            with open(lineitems_file, 'r', encoding='utf-8', newline='') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    key = f"{row.get('SalesPlanID', '')}|{row.get('ProductBarcode', '')}"
                    existing.add(key)
            
            self.logger.info(f"Loaded {len(existing)} existing LineItems for deduplication")
        except Exception as e:
            self.logger.error(f"Failed to load existing LineItems: {e}")
        
        return existing
    
    def append_salesplan_to_csv(self, salesplan_data: Dict[str, Any]):
        """Append SalesPlans data to CSV file"""
        output_dir = self.config.get_output_dir()
        salesplans_file = output_dir / "SalesPlans.csv"
        
        # Get column definitions
        columns = [col['name'] for col in self.config.get_salesplans_columns()]
        
        # Check if file exists to write header
        write_header = not salesplans_file.exists()
        
        try:
            with open(salesplans_file, 'a', encoding='utf-8', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=columns)
                
                if write_header:
                    writer.writeheader()
                    self.logger.debug("Wrote SalesPlans CSV header")
                
                writer.writerow(salesplan_data)
                
        except Exception as e:
            self.logger.error(f"Failed to write SalesPlans data: {e}")
            raise
    
    def append_lineitems_to_csv(self, lineitems_data: List[Dict[str, Any]]):
        """Append LineItems data to CSV file"""
        output_dir = self.config.get_output_dir()
        lineitems_file = output_dir / "LineItems.csv"
        
        # Get column definitions
        columns = [col['name'] for col in self.config.get_lineitems_columns()]
        
        # Check if file exists to write header
        write_header = not lineitems_file.exists()
        
        try:
            with open(lineitems_file, 'a', encoding='utf-8', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=columns)
                
                if write_header:
                    writer.writeheader()
                    self.logger.debug("Wrote LineItems CSV header")
                
                for lineitem in lineitems_data:
                    writer.writerow(lineitem)
                
        except Exception as e:
            self.logger.error(f"Failed to write LineItems data: {e}")
            raise
    
    def print_summary(self):
        """Print processing summary"""
        print("\n" + "="*50)
        print("CSV Export Summary")
        print("="*50)
        print(f"Files processed: {self.stats['files_processed']}")
        print(f"Files failed: {self.stats['files_failed']}")
        print(f"SalesPlans created: {self.stats['salesplans_created']}")
        print(f"LineItems created: {self.stats['lineitems_created']}")
        print(f"Duplicates skipped: {self.stats['duplicates_skipped']}")
        print(f"Unmapped picklist values: {self.stats['unmapped_values']}")
        
        if self.stats['unmapped_values'] > 0:
            logs_config = self.config.get_logs_config()
            log_file = logs_config.get('unmapped_values')
            if log_file:
                print(f"Check unmapped values in: {log_file}")
        
        output_dir = self.config.get_output_dir()
        print(f"\nOutput files:")
        print(f"  - {output_dir}/SalesPlans.csv")
        print(f"  - {output_dir}/LineItems.csv")
        print("="*50)

def main():
    """Main entry point"""
    # Setup logging
    logger = setup_logging()
    
    if len(sys.argv) < 2:
        print("Usage: python csv_exporter.py <input_directory>")
        print("  input_directory: Directory containing formatted Excel files")
        sys.exit(1)
    
    input_directory = sys.argv[1]
    
    try:
        exporter = CSVExporter(input_directory)
        exporter.export_to_csv()
        
        logger.info("CSV export completed successfully")
        
    except Exception as e:
        logger.error(f"CSV export failed: {e}")
        print(f"Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()