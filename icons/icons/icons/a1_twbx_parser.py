# Enhanced version of `a1_twbx_parser.py` for compatibility with all Tableau versions and better asset detection

import os
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import json
import re

def extract_twbx(twbx_file_path, extract_path):
    with zipfile.ZipFile(twbx_file_path, 'r') as zip_ref:
        zip_ref.extractall(extract_path)
    for file in os.listdir(extract_path):
        if file.endswith('.twb'):
            return os.path.join(extract_path, file)
    return None

def save_to_output_folder(data, file_name, output_folder):
    os.makedirs(output_folder, exist_ok=True)
    file_path = os.path.join(output_folder, file_name)
    with open(file_path, 'w') as file:
        json.dump(data, file, indent=4)
    print(f"Saved: {file_path}")
    return file_path

def extract_references_and_links(root):
    reference_data = []
    data_source_mapping = {}
    for datasource in root.findall('.//datasource'):
        ds_name = datasource.get('name')
        ds_caption = datasource.get('caption') or ds_name
        connections = datasource.findall('.//connection')
        for connection in connections:
            db_class = connection.get('class')
            db_name = connection.get('dbname') or connection.get('server') or 'N/A'
            ext_ref = connection.get('filename')
            if ext_ref:
                csv_name = os.path.splitext(os.path.basename(ext_ref))[0]
                data_source_mapping[ds_name] = csv_name
                reference_data.append({
                    'Data Source': ds_caption,
                    'Connection Type': db_class,
                    'Database Name': db_name,
                    'External Reference': ext_ref
                })
    return reference_data, data_source_mapping

def extract_visuals_and_layouts(root):
    visuals = []
    for worksheet in root.findall('.//worksheet'):
        name = worksheet.get('name')
        rows = [r.text for r in worksheet.findall('.//rows') if r.text]
        columns = [c.text for c in worksheet.findall('.//cols') if c.text]
        filters = [f.get('column') for f in worksheet.findall('.//filter') if f.get('column')]
        visuals.append({
            'Type': 'Worksheet', 'Source': name,
            'Rows': ', '.join(rows), 'Columns': ', '.join(columns),
            'Filters': ', '.join(filters), 'Legend': ''
        })
    for dashboard in root.findall('.//dashboard'):
        db_name = dashboard.get('name')
        views = set()
        for zone in dashboard.findall('.//zone'):
            view = zone.find('.//view')
            if view is not None and view.get('name'):
                views.add(view.get('name'))
        visuals.append({
            'Type': 'Dashboard', 'Source': db_name,
            'Worksheets': ', '.join(views), 'Rows': '', 'Columns': '', 'Filters': '', 'Legend': ''
        })
    return visuals

def parse_workbook(twb_path):
    tree = ET.parse(twb_path)
    root = tree.getroot()
    version = root.get('version') or 'unknown'
    print(f"Parsing Tableau version: {version}")
    refs, ds_map = extract_references_and_links(root)
    visuals = extract_visuals_and_layouts(root)
    return refs, visuals, version

def extract_tableau_workbook(file_path, output_dir):
    base = os.getcwd()
    out_dir = os.path.join(base, output_dir)
    os.makedirs(out_dir, exist_ok=True)

    extract_dir = os.path.join(out_dir, "input_tableau_output")
    os.makedirs(extract_dir, exist_ok=True)

    twb_file = extract_twbx(file_path, extract_dir) if file_path.endswith('.twbx') else file_path
    if not twb_file:
        raise ValueError("No .twb file found")

    references, visuals, version = parse_workbook(twb_file)

    extracted_data = {
        'references': references,
        'visuals': visuals,
        'version': version
    }
    save_to_output_folder(extracted_data, 'tableau_extracted_data.json', out_dir)
    print(f"Extraction complete: Tableau version {version}, saved to {out_dir}")

if __name__ == '__main__':
    file_path = input("Enter path to Tableau .twb/.twbx file: ").strip()
    if not os.path.isfile(file_path):
        print("Invalid file path")
    else:
        try:
            extract_tableau_workbook(file_path, "output")
        except Exception as e:
            print(f"Error: {e}")