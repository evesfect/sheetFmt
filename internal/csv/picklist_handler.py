#!/usr/bin/env python3
"""
Picklist Handler
Handles picklist value mapping and validation for CSV export.
"""

import logging
from typing import Dict, Any, Optional
from csv_config import CSVConfig

logger = logging.getLogger(__name__)

class PicklistHandler:
    """Handles picklist value mapping and validation"""
    
    def __init__(self, config: CSVConfig):
        self.config = config
        self.unmapped_count = 0
    
    def apply_picklist_mapping(self, value: Any, column_name: str, table_type: str) -> str:
        """
        Apply picklist mapping to a value if the column is a picklist column
        
        Args:
            value: Original value from data
            column_name: Name of the column
            table_type: "salesplans" or "lineitems"
            
        Returns:
            Mapped value or original value if not a picklist column
        """
        # Convert value to string and clean it
        if value is None:
            str_value = ""
        else:
            str_value = str(value).strip()
        
        # Check if this column is a picklist column
        if not self.config.is_picklist_column(column_name, table_type):
            logger.debug(f"Column '{column_name}' is not a picklist column, returning original value")
            return str_value
        
        # Get picklist configuration for this column
        picklist_config = self.config.get_picklist_config(column_name)
        if not picklist_config:
            logger.warning(f"No picklist configuration found for column '{column_name}'")
            return str_value
        
        # Try to find exact mapping
        mappings = picklist_config.get("mappings", {})
        if str_value in mappings:
            mapped_value = mappings[str_value]
            logger.debug(f"Mapped '{str_value}' -> '{mapped_value}' for column '{column_name}'")
            return mapped_value
        
        # Try case-insensitive mapping
        str_value_lower = str_value.lower()
        for original, mapped in mappings.items():
            if original.lower() == str_value_lower:
                mapped_value = mapped
                logger.debug(f"Case-insensitive mapping '{str_value}' -> '{mapped_value}' for column '{column_name}'")
                return mapped_value
        
        # No mapping found, use default
        default_value = picklist_config.get("default", "Other")
        
        # Log unmapped value if it's not empty
        if str_value:
            self.config.log_unmapped_value(column_name, str_value, default_value)
            self.unmapped_count += 1
            logger.info(f"Unmapped picklist value '{str_value}' in column '{column_name}', using default '{default_value}'")
        
        return default_value
    
    def validate_picklist_value(self, value: str, column_name: str) -> bool:
        """
        Validate if a value is in the allowed values for a picklist column
        
        Args:
            value: Value to validate
            column_name: Name of the picklist column
            
        Returns:
            True if value is allowed, False otherwise
        """
        picklist_config = self.config.get_picklist_config(column_name)
        if not picklist_config:
            return True  # If no config, assume valid
        
        allowed_values = picklist_config.get("allowed_values", [])
        return value in allowed_values
    
    def get_allowed_values(self, column_name: str) -> list:
        """
        Get list of allowed values for a picklist column
        
        Args:
            column_name: Name of the picklist column
            
        Returns:
            List of allowed values
        """
        picklist_config = self.config.get_picklist_config(column_name)
        if not picklist_config:
            return []
        
        return picklist_config.get("allowed_values", [])
    
    def derive_activity_type_from_filename(self, filename: str) -> str:
        """
        Derive ActivityType from filename using configurable keyword matching
        
        Args:
            filename: Name of the Excel file
            
        Returns:
            Mapped ActivityType value
        """
        return self._derive_from_filename_keywords(filename, "ActivityType", "salesplans")
    
    def derive_product_category_from_filename(self, filename: str) -> str:
        """
        Derive ProductCategory from filename using configurable keyword matching
        
        Args:
            filename: Name of the Excel file
            
        Returns:
            Mapped ProductCategory value
        """
        return self._derive_from_filename_keywords(filename, "ProductCategory", "lineitems")
    
    def get_unmapped_count(self) -> int:
        """Get count of unmapped values encountered"""
        return self.unmapped_count
    
    def reset_unmapped_count(self):
        """Reset the unmapped values counter"""
        self.unmapped_count = 0
    
    def _derive_from_filename_keywords(self, filename: str, column_name: str, table_type: str) -> str:
        """
        Generic method to derive picklist values from filename using configurable keywords
        
        Args:
            filename: Name of the Excel file
            column_name: Name of the picklist column
            table_type: "salesplans" or "lineitems"
            
        Returns:
            Mapped picklist value
        """
        filename_lower = filename.lower()
        
        # Get picklist configuration for this column
        picklist_config = self.config.get_picklist_config(column_name)
        if not picklist_config:
            logger.warning(f"No picklist configuration found for column '{column_name}'")
            return self.apply_picklist_mapping(filename, column_name, table_type)
        
        # Get filename keywords from configuration
        filename_keywords = picklist_config.get("filename_keywords", {})
        if not filename_keywords:
            logger.debug(f"No filename keywords configured for column '{column_name}', using direct mapping")
            return self.apply_picklist_mapping(filename, column_name, table_type)
        
        # Check for keyword matches
        for picklist_value, keywords in filename_keywords.items():
            for keyword in keywords:
                if keyword.lower() in filename_lower:
                    logger.debug(f"Detected {column_name} '{picklist_value}' from filename keyword '{keyword}'")
                    return picklist_value
        
        # No keyword match found, fallback to mapping or default
        logger.debug(f"No filename keywords matched for {column_name}, using fallback mapping")
        return self.apply_picklist_mapping(filename, column_name, table_type)