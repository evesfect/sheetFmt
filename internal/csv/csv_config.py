#!/usr/bin/env python3
"""
CSV Configuration Loader
Handles loading and managing all CSV export configurations.
"""

import json
import os
from pathlib import Path
from typing import Dict, Any, Optional
import logging

logger = logging.getLogger(__name__)

class CSVConfig:
    """Configuration manager for CSV export system"""
    
    def __init__(self, config_dir: str = "configs"):
        self.config_dir = Path(config_dir)
        self.csv_config = {}
        self.csv_columns = {}
        self.picklist_values = {}
        
        self.load_all_configs()
    
    def load_all_configs(self):
        """Load all CSV configuration files"""
        try:
            # Load main CSV config
            with open(self.config_dir / "csv_config.json", 'r', encoding='utf-8') as f:
                self.csv_config = json.load(f)
            
            # Load column definitions
            with open(self.config_dir / "csv_columns.json", 'r', encoding='utf-8') as f:
                self.csv_columns = json.load(f)
            
            # Load picklist values
            with open(self.config_dir / "picklist_values.json", 'r', encoding='utf-8') as f:
                self.picklist_values = json.load(f)
            
            logger.info("All CSV configurations loaded successfully")
            
        except Exception as e:
            logger.error(f"Failed to load CSV configurations: {e}")
            raise Exception(f"Failed to load CSV configurations: {e}")
    
    def get_next_id(self, id_type: str) -> str:
        """
        Get next available ID for salesplan or lineitem
        
        Args:
            id_type: "salesplan" or "lineitem"
            
        Returns:
            Formatted ID string (e.g., "Plan-10001")
        """
        if id_type == "salesplan":
            config_key = "salesplan_id"
        elif id_type == "lineitem":
            config_key = "lineitem_id"
        else:
            raise ValueError(f"Invalid id_type: {id_type}")
        
        id_config = self.csv_config[config_key]
        state_file = Path(id_config["state_file"])
        prefix = id_config["prefix"]
        start_number = id_config["start_number"]
        
        # Create state directory if it doesn't exist
        state_file.parent.mkdir(parents=True, exist_ok=True)
        
        # Read current ID or initialize
        if state_file.exists():
            try:
                with open(state_file, 'r') as f:
                    last_id = int(f.read().strip())
                next_id = last_id + 1
            except (ValueError, IOError):
                logger.warning(f"Could not read {state_file}, starting from {start_number}")
                next_id = start_number
        else:
            next_id = start_number
        
        # Write back the new ID
        try:
            with open(state_file, 'w') as f:
                f.write(str(next_id))
        except IOError as e:
            logger.error(f"Failed to update state file {state_file}: {e}")
            raise
        
        formatted_id = f"{prefix}{next_id}"
        logger.debug(f"Generated {id_type} ID: {formatted_id}")
        return formatted_id
    
    def get_output_dir(self) -> Path:
        """Get the CSV output directory path"""
        output_dir = Path(self.csv_config["output_dir"])
        output_dir.mkdir(parents=True, exist_ok=True)
        return output_dir
    
    def get_input_dir(self) -> Path:
        """Get the input directory containing Excel files"""
        return Path(self.csv_config["input_dir"])
    
    def get_logs_config(self) -> Dict[str, str]:
        """Get logging configuration"""
        return self.csv_config.get("logs", {})
    
    def get_excel_column_mappings(self) -> Dict[str, str]:
        """Get Excel column mappings configuration"""
        return self.csv_config.get("excel_column_mappings", {})
    
    def get_salesplans_columns(self) -> list:
        """Get SalesPlans column definitions"""
        return self.csv_columns["salesplans"]
    
    def get_lineitems_columns(self) -> list:
        """Get LineItems column definitions"""
        return self.csv_columns["lineitems"]
    
    def get_picklist_config(self, column_name: str) -> Optional[Dict[str, Any]]:
        """
        Get picklist configuration for a specific column
        
        Args:
            column_name: Name of the column
            
        Returns:
            Picklist config dict or None if not a picklist column
        """
        return self.picklist_values.get(column_name)
    
    def is_picklist_column(self, column_name: str, table_type: str) -> bool:
        """
        Check if a column is defined as a picklist column
        
        Args:
            column_name: Name of the column
            table_type: "salesplans" or "lineitems"
            
        Returns:
            True if column is picklist type
        """
        columns = self.csv_columns.get(table_type, [])
        for col in columns:
            if col["name"] == column_name and col["type"] == "picklist":
                return True
        return False
    
    def get_column_source_mapping(self, table_type: str) -> Dict[str, str]:
        """
        Get mapping of CSV column names to their source column names
        
        Args:
            table_type: "salesplans" or "lineitems"
            
        Returns:
            Dict mapping CSV column name to source column name
        """
        mapping = {}
        columns = self.csv_columns.get(table_type, [])
        
        for col in columns:
            csv_name = col["name"]
            source_name = col.get("source", csv_name)
            mapping[csv_name] = source_name
        
        return mapping
    
    def log_unmapped_value(self, column_name: str, original_value: str, mapped_value: str):
        """
        Log an unmapped picklist value for manual review
        
        Args:
            column_name: Name of the picklist column
            original_value: Original value found in data
            mapped_value: Value used as fallback
        """
        logs_config = self.get_logs_config()
        log_file = logs_config.get("unmapped_values")
        
        if not log_file:
            logger.warning("No unmapped values log file configured")
            return
        
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        
        try:
            with open(log_path, 'a', encoding='utf-8') as f:
                f.write(f"{column_name}|{original_value}|{mapped_value}\n")
            logger.debug(f"Logged unmapped value: {column_name} = '{original_value}' -> '{mapped_value}'")
        except IOError as e:
            logger.error(f"Failed to log unmapped value: {e}")


def load_csv_config(config_dir: str = "configs") -> CSVConfig:
    """
    Load CSV configuration
    
    Args:
        config_dir: Directory containing configuration files
        
    Returns:
        CSVConfig instance
    """
    return CSVConfig(config_dir)