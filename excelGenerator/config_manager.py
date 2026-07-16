#!/usr/bin/env python3
"""
STARK096 — ALDEA: Configuration Manager
Manage parameters for Excel generation
"""

import yaml
import os


def manage_config():
    """Manage configuration parameters"""
    
    print("🔧 Configuration Manager")
    print("")
    print("Available parameters to configure:")
    print("  • Company information")
    print("  • Default values")
    print("  • Template settings")
    print("  • Output paths")
    print("")
    
    # Example configuration structure
    config = {
        "company": {
            "name": "CONSTRUCTORA ALDEA",
            "address": "Calle Principal, 123",
            "phone": "+34 900 000 000",
            "email": "info@aldea.com"
        },
        "defaults": {
            "currency": "EUR",
            "date_format": "DD/MM/YYYY",
            "decimal_separator": ","
        },
        "paths": {
            "input": "input/",
            "output": "output/",
            "templates": "templates/"
        }
    }
    
    # Create config directory if it doesn't exist
    os.makedirs("config", exist_ok=True)
    
    # Save configuration
    config_file = "config/settings.yaml"
    with open(config_file, 'w', encoding='utf-8') as f:
        yaml.dump(config, f, default_flow_style=False, allow_unicode=True)
    
    print(f"✅ Configuration saved to: {config_file}")
    print("")
    print("Configuration structure created. Edit settings.yaml to customize.")
    
    return 0


if __name__ == "__main__":
    try:
        exit(manage_config())
    except Exception as e:
        print(f"❌ Error: {e}")
        exit(1)
