"""
FILE / SUB-FOLDER MANIFEST GENERATOR WITH IMAGE PROCESSING

Synopsis:
This script recursively scans a targeted directory (including OneDrive local sync folders) 
to build a structured JSON manifest of files. 

Key Features:
1. File Filtering & Range Expansion: Target specific individual files or sequential 
   numeric ranges (e.g., 'IMG_9481.JPG - IMG_9543.JPG').
2. Image Inspection: Opens image files (.jpg, .png, etc.) using the Pillow library 
   to read their actual dimensions (width x height) and populates the summary field.
3. Data Formatting: Converts raw byte sizes into human-readable formats (e.g., KB, MB).
4. Robustness: Gracefully handles locked files, permissions errors, or corrupted images.
"""

import os
import json
import re
from PIL import Image  # Requires: pip install pillow

def get_human_readable_size(size_bytes):
    """Converts bytes into a human-readable string format."""
    if size_bytes == 0:
        return "0 B"
    size_name = ("B", "KB", "MB", "GB", "TB")
    i = 0
    while size_bytes >= 1024 and i < len(size_name) - 1:
        size_bytes /= 1024.0
        i += 1
    return f"{size_bytes:.2f} {size_name[i]}"

def get_file_summary(full_path, extension):
    """
    Generates a summary. Opens images to extract actual dimensions.
    Returns a string summary or None.
    """
    ext = extension.lower()
    
    if ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp']:
        try:
            # Open the image file using Pillow
            with Image.open(full_path) as img:
                width, height = img.size
                return f"Image dimensions: {width}x{height} pixels."
        except Exception as e:
            # If the image is corrupted or cannot be read, log the error in the summary
            return f"Image file found, but could not be opened: {str(e)}"
            
    elif ext == '.pdf':
        return "PDF file (text extraction requires additional libraries like PyPDF2)."
    else:
        return None

def parse_file_targets(target_list):
    """Parses a list of individual files and ranges into a lowercase set."""
    final_targets = set()
    for item in target_list:
        item = item.strip()
        if not item:
            continue
        if '-' in item:
            try:
                parts = item.split('-')
                start_file = parts[0].strip()
                end_file = parts[1].strip()
                
                match_start = re.match(r"([a-zA-Z_]+)(\d+)(\.[a-zA-Z0-9]+)", start_file)
                match_end = re.match(r"([a-zA-Z_]+)(\d+)(\.[a-zA-Z0-9]+)", end_file)
                
                if match_start and match_end:
                    prefix = match_start.group(1)
                    ext = match_start.group(3)
                    start_num = int(match_start.group(2))
                    end_num = int(match_end.group(2))
                    padding = len(match_start.group(2))
                    
                    for num in range(start_num, end_num + 1):
                        gen_filename = f"{prefix}{str(num).zfill(padding)}{ext}"
                        final_targets.add(gen_filename.lower())
            except Exception as e:
                print(f"Error parsing range {item}: {e}")
        else:
            final_targets.add(item.lower())
    return final_targets

def scan_directory(target_dir, allowed_filenames=None):
    """Scans the directory and filters/processes files."""
    file_list = []
    
    for root, dirs, files in os.walk(target_dir):
        for file in files:
            if allowed_filenames is not None and file.lower() not in allowed_filenames:
                continue
                
            full_path = os.path.join(root, file)
            
            try:
                size_bytes = os.path.getsize(full_path)
                _, extension = os.path.splitext(file)
                
                file_entry = {
                    "type": "file",
                    "name": file,
                    "path": full_path,
                    "extension": extension,
                    "size_bytes": size_bytes,
                    "size_human": get_human_readable_size(size_bytes),
                    # This now actively opens the image
                    "summary": get_file_summary(full_path, extension) 
                }
                
                file_list.append(file_entry)
                
            except (OSError, PermissionError) as e:
                print(f"Skipping inaccessible file {file}: {e}")
                continue
                
    return file_list

if __name__ == "__main__":
    # --- CONFIGURATION ---
    # TARGET_DIRECTORY = r"C:\Users\Jason\FML Freight Solutions\FML Doc Share - Documents"
    TARGET_DIRECTORY = r"C:\Users\Jason\FML Freight Solutions\FML Doc Share - Documents\BARTRAC\CARGO ON HOLD\Excavator stick repositioning photos 15.05.25"
    OUTPUT_JSON_FILE = r"C:\Users\Jason\FML Freight Solutions\FML Doc Share - Documents\BARTRAC\CARGO ON HOLD\Excavator stick repositioning photos 15.05.25\directory_manifest.json"
    
    FILES_TO_INCLUDE = [
        "IMG_0713.JPG - IMG_0798.JPG",
        "SAD500 - 2 UNITS.pdf"
    ]
    # ---------------------
    
    print("Processing file selection list...")
    allowed_files = parse_file_targets(FILES_TO_INCLUDE) if FILES_TO_INCLUDE else None

    print(f"Scanning directory: {TARGET_DIRECTORY}...")
    manifest_data = scan_directory(TARGET_DIRECTORY, allowed_files)
    
    with open(OUTPUT_JSON_FILE, "w", encoding="utf-8") as json_file:
        json.dump(manifest_data, json_file, indent=2, ensure_ascii=False)
        
    print(f"Success! Manifest generated with {len(manifest_data)} matching files.")