#!/usr/bin/env python3
"""
OneDrive Sync & Directory Scanner (Interactive)

This script prompts the user for a directory path, then traverses it,
opens every file (reading 1 byte) to trigger OneDrive's "Files On‑Demand"
download, and finally outputs a JSON description of the entire directory tree.

The JSON output can be saved to a file or printed to the console.
"""

import os
import sys
import json
from datetime import datetime


def scan_directory(root_path, output_file=None):
    """
    Recursively scan root_path, open each file to force OneDrive sync,
    collect metadata, and write JSON to output_file (or stdout).
    """
    root_path = os.path.abspath(root_path)
    entries = []

    for dirpath, dirnames, filenames in os.walk(root_path):
        # Record directory entry (exclude root itself)
        rel_dir = os.path.relpath(dirpath, root_path)
        if rel_dir != '.':
            entries.append({
                "path": rel_dir,
                "type": "directory",
                "size": 0,
                "modified": None
            })

        for filename in filenames:
            file_path = os.path.join(dirpath, filename)
            rel_path = os.path.relpath(file_path, root_path)
            entry = {
                "path": rel_path,
                "type": "file",
                "size": 0,
                "modified": None
            }

            # --- Force OneDrive to download/hydrate the file ---
            try:
                # Open and read one byte. This blocks until the file is
                # fully available (if it was online‑only).
                with open(file_path, 'rb') as f:
                    f.read(1)
            except Exception as e:
                entry["error"] = str(e)
                # Still collect size/modified if possible (stat may work even if open fails)
            else:
                # If we got here, the file was opened successfully.
                # Gather metadata using os.stat (already local now).
                try:
                    stat = os.stat(file_path)
                    entry["size"] = stat.st_size
                    entry["modified"] = datetime.fromtimestamp(stat.st_mtime).isoformat()
                except Exception as e:
                    entry["error"] = str(e)

            entries.append(entry)

    # Output JSON
    if output_file:
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(entries, f, indent=2, ensure_ascii=False)
        print(f"✅ JSON written to {output_file}")
    else:
        print(json.dumps(entries, indent=2, ensure_ascii=False))


def main():
    print("=== OneDrive Sync & Directory Scanner ===")
    
    # Ask for directory
    while True:
        dir_path = input("Enter the directory path to scan (e.g., ~/OneDrive): ").strip()
        # Expand user home directory if present
        dir_path = os.path.expanduser(dir_path)
        if os.path.isdir(dir_path):
            break
        else:
            print(f"❌ Error: '{dir_path}' is not a valid directory. Please try again.")

    # Ask for output file (optional)
    out_path = input("Enter output JSON file path (press Enter to print to console): ").strip()
    if out_path:
        out_path = os.path.expanduser(out_path)
        # Ensure the directory for the output file exists
        out_dir = os.path.dirname(out_path)
        if out_dir and not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)

    print("\n🔄 Scanning and syncing files... (this may take a while)\n")
    scan_directory(dir_path, out_path if out_path else None)
    print("\n✅ Done.")


if __name__ == "__main__":
    main()