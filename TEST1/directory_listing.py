# Default: list the parent of BA3205, exclude .mov, no download
# python directory_listing.py

# Force download all files (except .mov) before generating the JSON
# python directory_listing.py --download

# List the BA3205 folder itself (not its parent) and download files
# python directory_listing.py --no-parent --download

# Use a different directory and output file
# python directory_listing.py "D:\Projects" -o mylist.json --download

# Exclude both .mov and .mp4 files
# python directory_listing.py -e .mov .mp4


import os
import json
import argparse
from datetime import datetime

def get_directory_listing(root_dir, exclude_extensions=None, download=False):
    """
    Recursively walk the given directory and return a list of items (files and folders),
    excluding files with specified extensions. If download=True, each file is opened
    and a small chunk is read to force OneDrive to download it.
    """
    if exclude_extensions is None:
        exclude_extensions = ['.mov']
    exclude_extensions = [ext.lower() for ext in exclude_extensions]

    contents = []
    for dirpath, dirnames, filenames in os.walk(root_dir):
        rel_path = os.path.relpath(dirpath, root_dir)
        if rel_path == '.':
            rel_path = ''

        # Directories
        for dirname in dirnames:
            full_path = os.path.join(dirpath, dirname)
            rel_dir = os.path.join(rel_path, dirname) if rel_path else dirname
            stat = os.stat(full_path)
            contents.append({
                'name': dirname,
                'path': rel_dir,
                'type': 'directory',
                'size': None,
                'mtime': datetime.fromtimestamp(stat.st_mtime).isoformat()
            })

        # Files (skip excluded extensions)
        for filename in filenames:
            ext = os.path.splitext(filename)[1].lower()
            if ext in exclude_extensions:
                continue
            full_path = os.path.join(dirpath, filename)
            rel_file = os.path.join(rel_path, filename) if rel_path else filename
            stat = os.stat(full_path)
            contents.append({
                'name': filename,
                'path': rel_file,
                'type': 'file',
                'size': stat.st_size,
                'mtime': datetime.fromtimestamp(stat.st_mtime).isoformat()
            })

    # If download flag is set, force OneDrive to download each file
    if download:
        # Filter only file entries (directories have type 'directory')
        file_entries = [item for item in contents if item['type'] == 'file']
        total = len(file_entries)
        if total == 0:
            print("No files to download.")
        else:
            # Try to use tqdm for a progress bar, otherwise fallback to simple prints
            try:
                from tqdm import tqdm
                iterator = tqdm(file_entries, desc="Downloading files")
            except ImportError:
                print(f"Downloading {total} files (install tqdm for a progress bar)...")
                iterator = file_entries
                for idx, item in enumerate(iterator):
                    print(f"  {idx+1}/{total}: {item['path']}")

            for idx, item in enumerate(iterator):
                if isinstance(iterator, list):  # no tqdm
                    pass  # progress already printed
                full_path = os.path.join(root_dir, item['path'])
                try:
                    # Open the file in binary mode and read a small chunk (1 KB)
                    # This triggers OneDrive to download the file.
                    with open(full_path, 'rb') as f:
                        f.read(1024)
                except Exception as e:
                    print(f"Warning: Could not download {full_path}: {e}")

            if 'tqdm' not in locals():
                print("Download complete.")

    return contents

def main():
    parser = argparse.ArgumentParser(
        description='Generate a JSON listing of a directory (default: parent of BA3205) '
                    'excluding .mov files. Use --download to force OneDrive to sync files.'
    )
    parser.add_argument(
        'directory',
        nargs='?',
        default=r"C:\Users\Jason\FML Freight Solutions\FML Doc Share - Documents\BARTRAC\CARGO ON HOLD\FML BOND STORE\BA3182 - 2606DSI2813 - BA3182 - CAT140GC - ON HOLD\PICS\WE PICS",
        help='Path to a folder; its parent will be listed (unless --no-parent is used)'
    )
    parser.add_argument(
        '-o', '--output',
        default='directory_listing.json',
        help='Output JSON file name (default: directory_listing.json)'
    )
    parser.add_argument(
        '-e', '--exclude',
        nargs='*',
        default=['.mov'],
        help='Extensions to exclude (case‑insensitive); default: .mov'
    )
    parser.add_argument(
        '--no-parent',
        action='store_true',
        help='List the given directory itself instead of its parent'
    )
    parser.add_argument(
        '-d', '--download',
        action='store_true',
        help='Force OneDrive to download all files (by reading the first 1 KB of each)'
    )
    args = parser.parse_args()

    # Determine the root directory to list
    if args.no_parent:
        root_dir = args.directory
    else:
        root_dir = os.path.dirname(args.directory)   # parent directory

    if not os.path.isdir(root_dir):
        print(f"Error: Directory '{root_dir}' does not exist.")
        return

    print(f"Listing directory: {root_dir}")
    if args.download:
        print("Download flag is ON – files will be synced from OneDrive.")
    contents = get_directory_listing(root_dir, args.exclude, args.download)

    data = {
        'root': root_dir,
        'contents': contents,
        'total_items': len(contents)
    }

    with open(args.output, 'w') as f:
        json.dump(data, f, indent=2)
    print(f"Output written to {args.output}")

if __name__ == '__main__':
    main()