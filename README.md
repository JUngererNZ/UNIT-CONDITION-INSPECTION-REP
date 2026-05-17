# FML Inspection Report Generator

Automates the creation of BARTRAC unit condition inspection reports for FML Freight Solutions warehouse and in-transit cargo. Replaces a manual PowerPoint process that took ~30 minutes per report.

## How It Works

For each BA number, the script reads a JSON data file and a folder of photos, then generates a branded, portrait-oriented `.pptx` report containing:
- **Cover slide:** Clean minimalist branding displaying the BA number, unit description, serial number, and inspection date.
- **Checklist table slide:** A high-density condition table with metadata headers, a status-tracking column layout, and a general comments block.
- **Auto-paginated photo grids:** Slides featuring a structured 10x2 photo canvas with automatic image compression and layout handling.

# FML Inspection Report Generator

Automates the creation of BARTRAC unit condition inspection reports for FML Freight Solutions warehouse and in-transit cargo. Replaces a manual PowerPoint process that took ~30 minutes per report.

## How It Works

For each BA number, the script reads a JSON data file and a folder of photos, then generates a branded, portrait-oriented `.pptx` report containing:
- **Cover slide:** Clean minimalist branding displaying the BA number, unit description, serial number, and inspection date.
- **Checklist table slide:** A high-density condition table with metadata headers, a status-tracking column layout, and a general comments block.
- **Auto-paginated photo grids:** Slides featuring a structured 10x2 photo canvas with automatic image compression and layout handling.

### Heading

---
inspections/└── BA3085/├── inspection_data.json     ← unit details + checklist entries├── inspection_checklist.md  ← human-readable checklist (optional reference)└── photos/                  ← any number of images, any size├── vin.jpg              ← explicitly named VIN/PIN image (forced to Slot 1)├── photo_01.jpg└── ...output/└── BA3085_CONDITION_INSPECTION_REPORT.pptx
---

## Data Mechanics: How `inspection_data.json` Maps to the Report

The python compilation engine extracts properties directly from `inspection_data.json` to programmatically stitch together text elements, dynamic table records, and conditional alerting icons across the presentation slides.

### 1. Root Metadata Fields Mapping

These properties populate the main identifying markers across multiple pages:

* **`ba_number`**
    * *Cover Page:* Rendered as a large bold element (`BA NUMBER:  BA3085`).
    * *Checklist & Photo Page Headers:* Populates the static header banner labels to identify individual slide files quickly.
    * *Output Filename:* Injected directly into the exported file name via string patterning.
* **`unit_description`**
    * *Cover Page:* Placed explicitly underneath the BA number as a crisp identifier text string.
    * *Checklist Page Subheader:* Mixed with the BA number to label row groupings.
* **`serial_number`**
    * *Cover Page:* Positioned dynamically beneath the `unit_description` string field.
    * *Checklist Page Subheader:* Displayed right beneath the primary description block.
* **`vin`**
    * *Checklist Page Subheader:* Displayed side-by-side with the serial number tracking fields.
* **`inspection_date`**
    * *Cover Page:* Positioned explicitly near the bottom details quadrant as an official audit stamp.
    * *Checklist Page Subheader:* Placed at the baseline of the meta text stack.
* **`status_note`**
    * *Checklist Page Subheader:* If populated, dynamically appends an explicit condition status line immediately below the date line.
* **`general_comments`**
    * *Checklist Page Baseline:* Maps straight into a dedicated, bounded rectangular note container spanning across the full width of the table footing.

### 2. Checklist Object Grid Evaluation

The array mapping loops through items in the order defined by the `checklist_items` key in `config.yaml`. For each entry item, the compiler looks up the corresponding key inside the JSON `"checklist"` object:

```json
"PAINTWORK": { 
  "present": true,  
  "damage": false, 
  "comments": "All in-tact, needs a good wash" 
}
```


Row Entry String Matching: The checkpoint identifier matches your config blueprint text (e.g. PAINTWORK).Presence Matrix (present): Determines whether the element exists on the physical asset. 
If false, it sets a high-visibility condition.Damage Exception Matrix (damage): Identifies if structural compromises were flagged during structural review.
Comments Extraction (comments): Grabs loose textual records and bounds them into the text box container with automatic character-wrapping limits.
3. Report Status & Colour Calculation LogicThe script tracks flags in tandem through an asset evaluator framework (_tick() and _tick_colour()) to output explicit typographic symbols:
Condition Matrix
Rendered SymbolTypography ColourOperational Meaning"present": true + "damage": false✓Green (#228844)
Component accounted for; passed inspection."present": true + "damage": true!Amber (#E68A00)
Component present, but structural anomalies or defects were logged."present": false✗Red (#CC2222)
Component missing, removed, or completely unaccounted for.
Additionally, whenever "damage": true is evaluated, an independent, high-visibility red ! alert marker is stamped into the isolated DAMAGE NOTED table column to let viewers parse mechanical deficiencies at a glance.
Technical RequirementsBashpip install python-pptx pyyaml Pillow
Python 3.10+

### Usage SyntaxBash# Generate reports for ALL inspection folders found in inspections/
python generate_report.py

# Generate report for a single targeted directory folder
python generate_report.py --ba BA3085

## Use a custom config blueprint pathway
python generate_report.py --config /path/to/config.yaml

Image Canvas Rules: The 10x2 Photo Canvas GridVIN Isolation Routine: When searching your target unit's photos/ folder, the script checks for any filename containing "vin" or "pin". 
- If identified, that image is intercepted and locked directly into Slot 1 (Top Left) on the first photo page with a VIN / PIN IMAGE identifier overlay text box.Aspect 
- Ratio Uniformity: Photo elements are scaled using an optimization thumbnail loop (ImageOps.exif_transpose and Image.LANCZOS). - Images fit within their designated bounding areas without experiencing stretching or distortion distortions.
- Dynamic Overflow Management: If a folder contains more images than the maximum photos_per_slide parameter configuration (e.g., 10 images total), the pipeline automatically creates page additions using the identical layout framework.

