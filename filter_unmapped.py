#!/usr/bin/env python3
import json
import sys
import os

def get_unmapped_columns():
    scanned_file = "data/output/scanned_columns"
    mapping_file = "data/output/column_mapping.json"
    
    # Read scanned columns
    if not os.path.exists(scanned_file):
        print(f"Error: {scanned_file} not found")
        return
    
    with open(scanned_file, 'r', encoding='utf-8') as f:
        scanned_columns = [line.strip() for line in f if line.strip()]
    
    # Read mappings (if exists)
    mapped_columns = set()
    if os.path.exists(mapping_file):
        with open(mapping_file, 'r', encoding='utf-8') as f:
            mapping_config = json.load(f)
        
        for mapping in mapping_config.get('mappings', []):
            mapped_columns.add(mapping['scanned_column'])
    
    # Filter unmapped
    unmapped = [col for col in scanned_columns if col not in mapped_columns]
    
    print(f"Total scanned columns: {len(scanned_columns)}")
    print(f"Mapped/ignored columns: {len(mapped_columns)}")
    print(f"Unmapped columns: {len(unmapped)}")
    print("\nUnmapped columns:")
    for col in unmapped:
        print(f"{col}|")

if __name__ == "__main__":
    get_unmapped_columns()