import os
import json
import yaml

def create_directory_listing(source_path, output_filename="inspection_listing"):
    # Normalize the path for Windows
    source_path = os.path.normpath(source_path)
    
    if not os.path.exists(source_path):
        print(f"Error: The path '{source_path}' does not exist.")
        return

    try:
        # Filter for files only (ignoring subdirectories)
        files = [f for f in os.listdir(source_path) if os.path.isfile(os.path.join(source_path, f))]
        
        # Prepare data structure
        data = {
            "source_directory": source_path,
            "unit_reference": "BA2797",
            "file_count": len(files),
            "file_list": sorted(files)
        }

        # Save to JSON
        json_path = f"{output_filename}.json"
        with open(json_path, 'w') as jf:
            json.dump(data, jf, indent=4)
        
        # Save to YAML
        yaml_path = f"{output_filename}.yaml"
        with open(yaml_path, 'w') as yf:
            yaml.dump(data, yf, default_flow_style=False, sort_keys=False)

        print(f"Listing complete.")
        print(f"JSON saved to: {os.path.abspath(json_path)}")
        print(f"YAML saved to: {os.path.abspath(yaml_path)}")
        print(f"Total files indexed: {len(files)}")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    # The specific path provided for the BA2797 inspection
    target_directory = r"C:\Users\Jason\FML Freight Solutions\FML Doc Share - Documents\BARTRAC\CARGO ON HOLD\FML BOND STORE\BA2797 - 2410DSI2607 - BA2797 - CAT428 - ON HOLD\PICS\Inspection photos 22.04.25"
    
    create_directory_listing(target_directory)