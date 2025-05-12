import os
import shutil
import glob
import json
import uuid
import argparse
from pathlib import Path
from typing import Dict, List, Optional, Union
from datetime import datetime

# Constants
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(SCRIPT_DIR)  # Parent directory of the script directory
NAVIGATOR_PATH = os.path.join(SCRIPT_DIR, "Navigator")  # Changed to use SCRIPT_DIR instead of REPO_ROOT

# --- Configuration ---
# Source folder relative to the script's working directory
SOURCE_FOLDER = os.path.join(REPO_ROOT, "output", "input_tableau_output", "Image")
# Destination and report paths relative to the Navigator folder
DESTINATION_RELATIVE = os.path.join("Navigator.Report", "StaticResources", "RegisteredResources")
REPORT_JSON_RELATIVE = os.path.join("Navigator.Report", "report.json")

# Power BI defaults
POWER_BI_WIDTH = 1280
POWER_BI_HEIGHT = 720

# --- Helper Functions ---

def validate_image_file(file_path: str) -> bool:
    """
    Validate that an image file exists and is accessible.
    
    Args:
        file_path (str): Path to the image file
        
    Returns:
        bool: True if the file is valid
    """
    try:
        return os.path.isfile(file_path) and os.access(file_path, os.R_OK)
    except Exception:
        return False

def find_navigator_folder() -> Optional[str]:
    """Find the Navigator folder in the icons directory."""
    print("\n--- Step 1: Finding 'Navigator' folder ---")
    print(f"Current working directory: {os.getcwd()}")
    print("Searching for base folder 'Navigator' in icons directory...")
    
    if os.path.exists(NAVIGATOR_PATH) and os.path.isdir(NAVIGATOR_PATH):
        print(f"Found 'Navigator' folder at: {NAVIGATOR_PATH}")
        print("'Navigator' folder found.")
        return NAVIGATOR_PATH
    else:
        print("Error: 'Navigator' folder not found.")
        return None

def copy_image_files(source_folder: str, destination_folder: str, tableau_data: Dict) -> List[str]:
    """Copy image files from source to destination folder."""
    print("\n--- Step 2: Copying image files ---")
    print(f"Source folder: {source_folder}")
    print(f"Destination folder: {destination_folder}")
    
    if not os.path.exists(source_folder):
        print(f"Error: Source folder '{source_folder}' does not exist.")
        return []
    
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
        print(f"Created destination folder: {destination_folder}")
    
    copied_files = []
    skipped_files = []
    
    # Process both tableau_zones and images sections
    for section in ["tableau_zones", "images"]:
        for item in tableau_data.get(section, []):
            # Get image name from either name field or param field
            image_name = None
            if section == "images":
                image_name = item.get("name")
            else:  # tableau_zones
                param = item.get("param", "")
                if param.startswith("Image/"):
                    image_name = param.split("/")[-1]
            
            if not image_name:
                continue
            
            # Handle subdirectories in image name
            image_path_parts = image_name.split('/')
            if len(image_path_parts) > 1:
                # Create subdirectory in destination
                dest_subdir = os.path.join(destination_folder, *image_path_parts[:-1])
                os.makedirs(dest_subdir, exist_ok=True)
                
                # Look for the image in source subdirectories
                source_path = os.path.join(source_folder, *image_path_parts)
                dest_path = os.path.join(destination_folder, *image_path_parts)
            else:
                source_path = os.path.join(source_folder, image_name)
                dest_path = os.path.join(destination_folder, image_name)
            
            if validate_image_file(source_path):
                try:
                    # Ensure the destination directory exists
                    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                    
                    # Copy the file
                    shutil.copy2(source_path, dest_path)
                    copied_files.append(image_name)
                    print(f"Copied: {image_name}")
                except Exception as e:
                    print(f"Error copying {image_name}: {e}")
                    skipped_files.append(image_name)
            else:
                print(f"Warning: Image file not found or inaccessible: {source_path}")
                skipped_files.append(image_name)
    
    if skipped_files:
        print(f"\nWarning: {len(skipped_files)} files were skipped:")
        for file in skipped_files:
            print(f"  - {file}")
    
    if not copied_files:
        print("Warning: No image files were copied.")
    else:
        print(f"\nSuccessfully copied {len(copied_files)} files.")
    
    return copied_files

def read_tableau_data(json_path: str) -> Optional[Dict]:
    """Read and parse the Tableau extracted data JSON file."""
    print(f"\n--- Step 3: Reading Tableau data from '{os.path.basename(json_path)}' ---")
    if not os.path.exists(json_path):
        print(f"Error: Tableau data JSON file not found at '{json_path}'.")
        return None

    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            tableau_data = json.load(f)
            
            # Validate required sections
            required_sections = ["tableau_zones", "images", "metadata"]
            missing_sections = [section for section in required_sections if section not in tableau_data]
            
            if missing_sections:
                print(f"Warning: Missing sections in Tableau data: {', '.join(missing_sections)}")
            
            print("Successfully read Tableau data.")
            return tableau_data
    except json.JSONDecodeError as e:
        print(f"Error decoding Tableau data JSON file '{json_path}': {e}")
        return None
    except Exception as e:
        print(f"An unexpected error occurred while reading Tableau data JSON '{json_path}': {e}")
        return None

def create_resource_package(image_files: List[str]) -> Dict:
    """Create a resource package configuration for Power BI."""
    # Use a set to ensure unique image names
    unique_images = list(set(image_files))
    return {
        "resourcePackage": {
            "disabled": False,
            "items": [
                {
                    "name": image_name,
                    "path": image_name,
                    "type": 100  # Image type
                }
                for image_name in unique_images
            ],
            "name": "RegisteredResources",
            "type": 1
        }
    }

def create_image_container(zone: Dict, image_name: str) -> Dict:
    """
    Create a Power BI image container configuration.
    
    Args:
        zone (Dict): Zone data from Tableau
        image_name (str): Name of the image file
        
    Returns:
        Dict: Image container configuration
    """
    # Extract position data
    x = float(zone.get("x", 0))
    y = float(zone.get("y", 0))
    width = float(zone.get("w", 0))
    height = float(zone.get("h", 0))
    
    # Ensure image name is properly formatted
    if '/' in image_name:
        image_name = image_name.split('/')[-1]
    
    # Create image container with proper configuration
    image_config = {
        "name": f"{uuid.uuid4().hex[:20]}",
        "layouts": [{
            "id": 0,
            "position": {
                "x": x,
                "y": y,
                "z": 9,
                "width": width,
                "height": height,
                "tabOrder": 7
            }
        }],
        "singleVisual": {
            "visualType": "image",
            "drillFilterOtherVisuals": True,
            "objects": {
                "general": [{
                    "properties": {
                        "imageUrl": {
                            "expr": {
                                "ResourcePackageItem": {
                                    "PackageName": "RegisteredResources",
                                    "PackageType": 1,
                                    "ItemName": image_name
                                }
                            }
                        },
                        "imageScaling": {
                            "expr": {
                                "Literal": {
                                    "Value": "'Fit'"
                                }
                            }
                        },
                        "imageAlignment": {
                            "expr": {
                                "Literal": {
                                    "Value": "'Center'"
                                }
                            }
                        }
                    }
                }]
            }
        },
        "howCreated": "InsertVisualButton"
    }

    return {
        "config": json.dumps(image_config),
        "filters": "[]",
        "height": height,
        "width": width,
        "x": x,
        "y": y,
        "z": 9.00
    }

def create_powerbi_report(tableau_data: Dict, image_files: List[str]) -> Dict:
    """Create a Power BI report configuration."""
    # Create sections from tableau_zones
    sections = []
    dashboard_zones = {}
    
    # Group zones by dashboard
    for zone in tableau_data.get("tableau_zones", []):
        dashboard_name = zone.get("dashboard", "Unknown Dashboard")
        dashboard_id = zone.get("dashboard_id", "")
        key = f"{dashboard_name}_{dashboard_id}" if dashboard_id else dashboard_name
        
        if key not in dashboard_zones:
            dashboard_zones[key] = {
                "name": dashboard_name,
                "zones": []
            }
        dashboard_zones[key]["zones"].append(zone)
    
    # Create sections for each dashboard
    for dashboard_key, dashboard_info in dashboard_zones.items():
        section = {
            "config": "{}",
            "displayName": dashboard_info["name"],
            "displayOption": 1,
            "filters": "[]",
            "height": 720,
            "name": uuid.uuid4().hex[:20],
            "width": 1280,
            "visualContainers": []
        }
        
        # Add visual containers for each zone
        for zone in dashboard_info["zones"]:
            if zone.get("param", "").startswith("Image/"):
                image_name = zone["param"].split("/")[-1]
                section["visualContainers"].append(create_image_container(zone, image_name))
        
        sections.append(section)
    
    # Create the report configuration
    report = {
        "config": json.dumps({
            "version": "5.59",
            "themeCollection": {
                "baseTheme": {
                    "name": "CY24SU10",
                    "version": "5.62",
                    "type": 2
                }
            },
            "activeSectionIndex": 0,
            "defaultDrillFilterOtherVisuals": True,
            "linguisticSchemaSyncVersion": 0,
            "settings": {
                "useNewFilterPaneExperience": True,
                "allowChangeFilterTypes": True,
                "useStylableVisualContainerHeader": True,
                "queryLimitOption": 6,
                "useEnhancedTooltips": True,
                "exportDataMode": 1,
                "useDefaultAggregateDisplayName": True
            }
        }),
        "layoutOptimization": 0,
        "resourcePackages": [
            {
                "resourcePackage": {
                    "disabled": False,
                    "items": [
                        {
                            "name": "CY24SU10",
                            "path": "BaseThemes/CY24SU10.json",
                            "type": 202
                        }
                    ],
                    "name": "SharedResources",
                    "type": 2
                }
            },
            create_resource_package(image_files)
        ],
        "sections": sections,
        "metadata": {
            "powerbi_dimensions": {
                "width": 1280,
                "height": 720
            },
            "tableau_dimensions": {
                "max_width": 100000,
                "max_height": 100000
            },
            "valid_extensions": [
                ".png",
                ".jpg",
                ".jpeg",
                ".gif",
                ".bmp",
                ".svg",
                ".ico",
                ".webp"
            ],
            "source_file": tableau_data.get("source_file", ""),
            "processed_at": datetime.now().isoformat()
        }
    }
    
    return report

def update_report_json(tableau_data: Dict, image_files: List[str]) -> bool:
    """Update the report.json file with the new position data."""
    try:
        # Find the Navigator folder
        navigator_folder = find_navigator_folder()
        if not navigator_folder:
            print("Error: Could not find 'Navigator' folder.")
            return False

        # Create the report configuration
        report = create_powerbi_report(tableau_data, image_files)
        
        # Write the report.json file
        report_json_path = os.path.join(navigator_folder, "Navigator.Report", "report.json")
        with open(report_json_path, "w", encoding="utf-8") as f:
            json.dump(report, f, indent=2)
        
        print("Successfully updated report.json")
        return True
    except Exception as e:
        print(f"Error updating report.json: {str(e)}")
        return False

def main(tableau_data_path: str, image_source_folder: str):
    """Main function to process the Tableau data and generate Power BI report."""
    print("\nScript started.\n")
    print("=== Starting Icon Generation Process ===\n")

    # Step 1: Find Navigator folder
    print("--- Step 1: Finding 'Navigator' folder ---")
    navigator_folder = find_navigator_folder()
    if not navigator_folder:
        print("Error: Could not find 'Navigator' folder.")
        return

    # Step 2: Read Tableau data
    print("\n--- Step 3: Reading Tableau data from 'processed_data.json' ---")
    try:
        with open(tableau_data_path, "r", encoding="utf-8") as f:
            tableau_data = json.load(f)
        print("Successfully read Tableau data.")
    except Exception as e:
        print(f"Error reading Tableau data: {str(e)}")
        return

    # Step 3: Copy image files
    print("\n--- Step 2: Copying image files ---")
    dest_folder = os.path.join(navigator_folder, "Navigator.Report", "StaticResources", "RegisteredResources")
    image_files = copy_image_files(image_source_folder, dest_folder, tableau_data)
    if not image_files:
        print("Error: No image files were copied.")
        return

    # Step 4: Update report.json
    if not update_report_json(tableau_data, image_files):
        print("Error: Failed to update report.json")
        return

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate Power BI report from Tableau data")
    parser.add_argument("tableau_data_path", help="Path to the processed Tableau data JSON file")
    parser.add_argument("image_source_folder", help="Path to the folder containing source images")
    args = parser.parse_args()
    main(args.tableau_data_path, args.image_source_folder)


