#!/usr/bin/env python3
"""
STARK096 — ALDEA: Utility Functions
Common utilities for Excel generation
"""

import os
from datetime import datetime
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


def create_output_dir(output_dir="output"):
    """Create output directory if it doesn't exist"""
    os.makedirs(output_dir, exist_ok=True)
    return output_dir


def get_timestamp():
    """Get current timestamp in standard format"""
    return datetime.now().strftime('%Y%m%d_%H%M%S')


def get_header_style():
    """Get header cell style"""
    fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    font = Font(bold=True, color="FFFFFF", size=12, name="Calibri")
    alignment = Alignment(horizontal="center", vertical="center")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    return {
        'fill': fill,
        'font': font,
        'alignment': alignment,
        'border': border
    }


def get_cell_border():
    """Get standard cell border"""
    return Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )


def apply_header_style(cell):
    """Apply header style to a cell"""
    style = get_header_style()
    cell.fill = style['fill']
    cell.font = style['font']
    cell.alignment = style['alignment']
    cell.border = style['border']


def apply_cell_border(cell):
    """Apply border to a cell"""
    cell.border = get_cell_border()


def print_success(message):
    """Print success message with emoji"""
    print(f"✅ {message}")


def print_error(message):
    """Print error message with emoji"""
    print(f"❌ {message}")


def print_warning(message):
    """Print warning message with emoji"""
    print(f"⚠️  {message}")


def print_info(message):
    """Print info message with emoji"""
    print(f"ℹ️  {message}")
    """Print info message with emoji"""
    print(f"ℹ️  {message}")


if __name__ == "__main__":
    # Test utilities
    print("🔧 Testing utility functions...")
    
    # Test config loading
    config = load_config()
    if config:
        print_success("Configuration loaded")
    
    # Test timestamp
    ts = get_timestamp()
    print_info(f"Current timestamp: {ts}")
    
    # Test currency formatting
    price = 185000
    formatted = format_currency(price, config)
    print_info(f"Formatted price: {formatted}")
    
    print_success("All utilities working correctly!")
