#!/usr/bin/env python3
import argparse
import os
import glob
from pathlib import Path
from PIL import Image
import io  # Not needed for file save, but for consistency

def convert_photos_to_png(insdir: str):
    photodir = os.path.join(insdir, 'photos')
    if not os.path.isdir(photodir):
        print(f"No photos dir: {photodir}")
        return
    
    exts = ['*.jpg', '*.jpeg', '*.JPG', '*.JPEG', '*.png', '*.PNG']
    paths = []
    for ext in exts:
        paths.extend(glob.glob(os.path.join(photodir, ext)))
    
    print(f"Found {len(paths)} images in {photodir}")
    for path in paths:
        img = Image.open(path)
        # Convert to RGBA PNG
        img_rgba = img.convert('RGBA')
        png_path = Path(path).with_suffix('.png')
        img_rgba.save(png_path)
        print(f"Converted {path} -> {png_path}")
        # Optional: delete original
        # os.remove(path)
    
    print("Conversion complete.")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Convert inspection photos to PNG")
    parser.add_argument('inspection_dir', help="Path to inspection dir, e.g., inspections/BA2797")
    args = parser.parse_args()
    convert_photos_to_png(args.inspection_dir)