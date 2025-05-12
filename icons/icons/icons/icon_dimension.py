import argparse
import xml.etree.ElementTree as ET
import os
import json
import base64
from pathlib import Path
from typing import List, Dict, Optional, Union, Tuple, Set
import datetime

# Power BI dashboard size defaults
POWER_BI_WIDTH = 1280
POWER_BI_HEIGHT = 720

# Tableau's virtual coordinate system size
TABLEAU_MAX_WIDTH = 100000
TABLEAU_MAX_HEIGHT = 100000

# Zone type to exclude (dashboard background or base layouts)
EXCLUDE_ZONE_TYPE = 'Type v2 layout base'

# Full dashboard coverage threshold (percentage)
FULL_COVERAGE_THRESHOLD = 0.95  # 95% coverage

# Valid image extensions
VALID_IMAGE_EXTENSIONS = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.ico', '.webp']

# Background image threshold (percentage of total area)
BACKGROUND_THRESHOLD = 0.95  # 95% of total area

# Minimum valid dimensions
MIN_DIMENSION = 1
MAX_DIMENSION = 100000


def extract_base64_images(root: ET.Element) -> List[Dict]:
    """
    Extract base64-encoded images from the workbook.
    
    Args:
        root (ET.Element): Root element of the XML tree
        
    Returns:
        List[Dict]: List of dictionaries containing image data
    """
    images = []
    for shape in root.findall('.//shape'):
        name = shape.get('name', '')
        if any(name.lower().endswith(ext) for ext in VALID_IMAGE_EXTENSIONS):
            try:
                # Get base64 data
                base64_data = shape.text.strip() if shape.text else ''
                if base64_data:
                    images.append({
                        'name': name,
                        'data': base64_data,
                        'extension': os.path.splitext(name)[1].lower()
                    })
            except Exception as e:
                print(f"Warning: Could not extract base64 data from shape {name}: {str(e)}")
    return images


def save_base64_images(images: List[Dict], output_dir: str) -> List[str]:
    """
    Save base64-encoded images to files.
    
    Args:
        images (List[Dict]): List of dictionaries containing image data
        output_dir (str): Directory to save images to
        
    Returns:
        List[str]: List of saved image paths
    """
    saved_paths = []
    for img in images:
        try:
            # Create output directory if it doesn't exist
            os.makedirs(output_dir, exist_ok=True)
            
            # Clean up the file name and create subdirectories if needed
            name_parts = img['name'].split('/')
            if len(name_parts) > 1:
                # Create subdirectory
                subdir = os.path.join(output_dir, *name_parts[:-1])
                os.makedirs(subdir, exist_ok=True)
            
            # Save image
            output_path = os.path.join(output_dir, *name_parts)
            with open(output_path, 'wb') as f:
                f.write(base64.b64decode(img['data']))
            saved_paths.append(output_path)
            print(f"Saved image: {output_path}")
        except Exception as e:
            print(f"Warning: Could not save image {img['name']}: {str(e)}")
    return saved_paths


def get_excluded_zone_coordinates(root: ET.Element) -> Set[Tuple[int, int, int, int]]:
    """
    Get coordinates of zones that should be excluded (Type v2 layout base).
    
    Args:
        root (ET.Element): Root element of the XML tree
        
    Returns:
        Set[Tuple[int, int, int, int]]: Set of (x, y, width, height) tuples for excluded zones
    """
    excluded_coords = set()
    for zone in root.findall('.//zone'):
        if zone.get('param') == EXCLUDE_ZONE_TYPE:
            x, y, w, h = zone.get('x'), zone.get('y'), zone.get('w'), zone.get('h')
            if all([x, y, w, h]):
                try:
                    coords = tuple(map(int, (x, y, w, h)))
                    excluded_coords.add(coords)
                except (ValueError, TypeError):
                    continue
    return excluded_coords


def find_dashboards(root: ET.Element) -> List[ET.Element]:
    """
    Find all dashboard elements in the workbook.
    
    Args:
        root (ET.Element): Root element of the XML tree
        
    Returns:
        List[ET.Element]: List of dashboard elements
    """
    # Try different possible paths for dashboards
    dashboards = root.findall('.//dashboard')
    if not dashboards:
        dashboards = root.findall('.//workbook/dashboards/dashboard')
    if not dashboards:
        dashboards = root.findall('.//workbook/windows/window/dashboard')
    return dashboards


def extract_zone_data(twb_file_path: Union[str, Path]) -> Union[List[Dict[str, str]], str]:
    """
    Opens a .twb file, extracts x, y, w, and h from <zone> elements
    where the 'param' attribute starts with 'Image/' and
    excludes zones of type EXCLUDE_ZONE_TYPE or full dashboard coverage.
    Now includes dashboard information for each zone.

    Args:
        twb_file_path (Union[str, Path]): The path to the .twb file.

    Returns:
        Union[List[Dict[str, str]], str]: A list of dictionaries containing zone data with dashboard info,
                                         or an error message string if something fails.
    """
    twb_path = Path(twb_file_path)
    if not twb_path.exists():
        return f"Error: File not found at {twb_path}"

    try:
        tree = ET.parse(twb_path)
        root = tree.getroot()
        zone_data = []
        
        # Extract base64 images
        base64_images = extract_base64_images(root)
        
        # Save base64 images
        if base64_images:
            output_dir = os.path.join(os.path.dirname(twb_path), 'Image')
            saved_paths = save_base64_images(base64_images, output_dir)
            print(f"Saved {len(saved_paths)} images to {output_dir}")
        
        # Get coordinates of zones to exclude
        excluded_coords = get_excluded_zone_coordinates(root)

        # Find all dashboards
        dashboards = find_dashboards(root)
        
        if not dashboards:
            print("Warning: No dashboards found in the workbook")
            return zone_data
        
        # Create a mapping of dashboard names to BIZ icons
        dashboard_icons = {
            "Financial Intelligence Toolkit": "BIZ/paper-plane.png",
            "Inventory Supply Planner": "BIZ/warehouse-with-boxes.png",
            "Strategic Forecast Summary": "BIZ/meditate.png",
            "YTD Growth Watch": "BIZ/click.png"
        }
        
        for dashboard in dashboards:
            dashboard_name = dashboard.get('name', 'Unknown Dashboard')
            dashboard_id = dashboard.get('id', '')
            
            # Find all zones within this dashboard
            zones = dashboard.findall('.//zone')
            if not zones:
                zones = dashboard.findall('.//zones/zone')
            
            # Add the BIZ icon for this dashboard
            if dashboard_name in dashboard_icons:
                zone_data.append({
                    'dashboard': dashboard_name,
                    'dashboard_id': dashboard_id,
                    'x': '50',  # Left side of the dashboard
                    'y': '50',  # Top of the dashboard
                    'w': '100',  # Reasonable size for an icon
                    'h': '100',  # Square aspect ratio
                    'param': f"Image/{dashboard_icons[dashboard_name]}",
                    'original_x': '5000',
                    'original_y': '5000',
                    'original_w': '10000',
                    'original_h': '10000'
                })
            
            for zone in zones:
                # Skip excluded zone types
                if zone.get('param') == EXCLUDE_ZONE_TYPE:
                    continue

                param = zone.get('param', '')
                if param and (param.startswith('Image/') or param.endswith('.png') or param.endswith('.jpg')):
                    x, y, w, h = zone.get('x'), zone.get('y'), zone.get('w'), zone.get('h')
                    if all([x, y, w, h]) and validate_coordinates(x, y, w, h):
                        # Skip if it's a full dashboard coverage zone
                        if is_full_dashboard_coverage(x, y, w, h):
                            continue
                            
                        zone_dict = {'x': x, 'y': y, 'w': w, 'h': h, 'param': param}
                        
                        # Skip if it's a background image
                        if is_background_image(zone_dict):
                            continue
                            
                        # Scale coordinates to Power BI dimensions
                        scaled_x, scaled_y, scaled_w, scaled_h = scale_to_powerbi(x, y, w, h)
                        if all([scaled_x, scaled_y, scaled_w, scaled_h]):
                            zone_data.append({
                                'dashboard': dashboard_name,
                                'dashboard_id': dashboard_id,
                                'x': str(scaled_x),
                                'y': str(scaled_y),
                                'w': str(scaled_w),
                                'h': str(scaled_h),
                                'param': param,
                                'original_x': x,
                                'original_y': y,
                                'original_w': w,
                                'original_h': h
                            })
        
        # Extract image information from the workbook
        images = []
        for zone in zone_data:
            param = zone['param']
            image_name = param.split('/')[-1] if '/' in param else param
            if any(image_name.lower().endswith(ext) for ext in VALID_IMAGE_EXTENSIONS):
                images.append({
                    'name': image_name,
                    'dashboard': zone['dashboard'],
                    'dashboard_id': zone['dashboard_id'],
                    'x': zone['x'],
                    'y': zone['y'],
                    'width': zone['w'],
                    'height': zone['h'],
                    'original_x': zone['original_x'],
                    'original_y': zone['original_y'],
                    'original_width': zone['original_w'],
                    'original_height': zone['original_h']
                })
        
        # Add base64 images to the list
        for img in base64_images:
            images.append({
                'name': img['name'],
                'dashboard': 'Unknown Dashboard',  # Base64 images don't have dashboard info
                'dashboard_id': '',
                'x': '0',
                'y': '0',
                'width': str(POWER_BI_WIDTH),  # Default to full width
                'height': str(POWER_BI_HEIGHT),  # Default to full height
                'original_x': '0',
                'original_y': '0',
                'original_width': str(TABLEAU_MAX_WIDTH),
                'original_height': str(TABLEAU_MAX_HEIGHT)
            })
        
        return {
            'tableau_zones': zone_data,
            'images': images,
            'metadata': {
                'powerbi_dimensions': {
                    'width': POWER_BI_WIDTH,
                    'height': POWER_BI_HEIGHT
                },
                'tableau_dimensions': {
                    'max_width': TABLEAU_MAX_WIDTH,
                    'max_height': TABLEAU_MAX_HEIGHT
                },
                'valid_extensions': VALID_IMAGE_EXTENSIONS,
                'source_file': str(twb_path),
                'processed_at': str(datetime.datetime.now())
            }
        }

    except ET.ParseError as e:
        return f"Error parsing XML: {str(e)}"
    except Exception as e:
        return f"Unexpected error while processing {twb_path}: {str(e)}"


def validate_coordinates(x: str, y: str, w: str, h: str) -> bool:
    """
    Validate that coordinates are within acceptable ranges.
    
    Args:
        x (str): X coordinate
        y (str): Y coordinate
        w (str): Width
        h (str): Height
        
    Returns:
        bool: True if coordinates are valid
    """
    try:
        xi, yi, wi, hi = map(float, (x, y, w, h))
        return all([
            MIN_DIMENSION <= xi <= MAX_DIMENSION,
            MIN_DIMENSION <= yi <= MAX_DIMENSION,
            MIN_DIMENSION <= wi <= MAX_DIMENSION,
            MIN_DIMENSION <= hi <= MAX_DIMENSION
        ])
    except (ValueError, TypeError):
        return False


def is_background_image(zone: Dict[str, str]) -> bool:
    """
    Check if a zone is a background image based on its size and position.
    
    Args:
        zone (Dict[str, str]): Zone data dictionary
        
    Returns:
        bool: True if the zone is a background image
    """
    try:
        x = float(zone.get('x', 0))
        y = float(zone.get('y', 0))
        w = float(zone.get('w', 0))
        h = float(zone.get('h', 0))
        
        # Calculate area coverage
        area_coverage = (w * h) / (TABLEAU_MAX_WIDTH * TABLEAU_MAX_HEIGHT)
        
        # Check if it's positioned at origin and covers most of the area
        return (x == 0 and y == 0 and area_coverage >= BACKGROUND_THRESHOLD)
    except (ValueError, TypeError):
        return False


def is_full_dashboard_coverage(x: str, y: str, w: str, h: str) -> bool:
    """
    Check if a zone covers the full dashboard.
    
    Args:
        x (str): X coordinate
        y (str): Y coordinate
        w (str): Width
        h (str): Height
        
    Returns:
        bool: True if the zone covers the full dashboard
    """
    try:
        xi, yi, wi, hi = map(float, (x, y, w, h))
        # Calculate coverage percentage
        coverage = (wi * hi) / (TABLEAU_MAX_WIDTH * TABLEAU_MAX_HEIGHT)
        # Check if it's positioned at origin and covers most of the area
        return (xi == 0 and yi == 0 and coverage >= FULL_COVERAGE_THRESHOLD)
    except (ValueError, TypeError):
        return False


def scale_to_powerbi(x: str, y: str, width: str, height: str) -> Tuple[float, float, float, float]:
    """
    Scale Tableau coordinates to Power BI dimensions.

    Args:
        x (str): X coordinate in Tableau's coordinate system
        y (str): Y coordinate in Tableau's coordinate system
        width (str): Width in Tableau's coordinate system
        height (str): Height in Tableau's coordinate system

    Returns:
        Tuple[float, float, float, float]: Scaled coordinates (x, y, width, height)
    """
    try:
        xi, yi, wi, hi = map(float, (x, y, width, height))
        return (
            round((xi / TABLEAU_MAX_WIDTH) * POWER_BI_WIDTH, 2),
            round((yi / TABLEAU_MAX_HEIGHT) * POWER_BI_HEIGHT, 2),
            round((wi / TABLEAU_MAX_WIDTH) * POWER_BI_WIDTH, 2),
            round((hi / TABLEAU_MAX_HEIGHT) * POWER_BI_HEIGHT, 2)
        )
    except (ValueError, TypeError) as e:
        print(f"Warning: Error scaling coordinates: {str(e)}")
        return (None, None, None, None)
s

def main(twb_path: str, output_dir: Optional[str] = None):
    result = extract_zone_data(twb_path)
    if isinstance(result, str):  # Error occurred
        print(result)
        return

    out_dir = output_dir or os.path.join(os.path.dirname(twb_path), 'processed')
    os.makedirs(out_dir, exist_ok=True)
    
    output_file = os.path.join(out_dir, 'processed_data.json')
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, indent=2)
    print(f"Saved to {out_dir}")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Extract icon zones from a Tableau TWB.')
    parser.add_argument('twb_file', help='Path to the .twb file')
    parser.add_argument('-o', '--output', help='Output directory', default=None)
    args = parser.parse_args()
    main(args.twb_file, args.output)
