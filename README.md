# FML Inspection Report Generator

Automates the creation of BARTRAC unit condition inspection reports for FML Freight Solutions warehouse and in-transit cargo. Replaces a manual PowerPoint process that took ~30 minutes per report.

## How It Works

For each BA number, the script reads a JSON data file and a folder of photos, then generates a branded `.pptx` report containing:
- Cover slide
- Typed checklist table with colour-coded status
- Auto-paginated photo grid slides

```
inspections/
└── BA3036/
    ├── inspection_data.json     ← unit details + checklist entries
    ├── inspection_checklist.md  ← human-readable checklist (optional reference)
    └── photos/                  ← any number of images, any size
        ├── photo_01.jpg
        └── ...

output/
└── BA3036_CONDITION_INSPECTION_REPORT.pptx
```

---

## Requirements

```bash
pip install python-pptx pyyaml
```

Python 3.10+

---

## Usage

```bash
# Generate reports for ALL BA folders in inspections/
python generate_report.py

# Generate report for a single unit
python generate_report.py --ba BA3036

# Use a custom config file
python generate_report.py --config /path/to/config.yaml
```

---

## Configuration

All paths and report settings are controlled by `config.yaml`:

```yaml
paths:
  inspections_dir: "./inspections"   # one subfolder per BA number
  output_dir: "./output"             # generated .pptx files land here
  logo: "./assets/fml_logo.png"      # optional; leave blank to skip

report:
  photos_per_slide: 6                # must be even (4, 6, 8, 10, 12)
  output_filename_pattern: "{ba_number}_CONDITION_INSPECTION_REPORT.pptx"
```

---

## Adding a New Inspection

1. Create a folder under `inspections/` named after the BA number (e.g. `BA3048`)
2. Copy an existing `inspection_data.json` into the folder and update all fields
3. Drop photos into `inspections/BA3048/photos/` — any size, any quantity
4. Run:
   ```bash
   python generate_report.py --ba BA3048
   ```

---

## inspection_data.json Reference

```json
{
  "ba_number": "BA3036",
  "unit_description": "CAT 426 BACKHOE LOADER IN SKD FORM",
  "serial_number": "JZ404401",
  "vin": "CAT00426HJZ404401",
  "inspection_date": "13/05/2026",
  "inspector": "FML Warehouse Team",
  "status_note": "Optional free-text note shown under the inspection date.",

  "checklist": {
    "PAINTWORK": { "present": true, "damage": true, "comments": "Damage to RHS grease nipple on boom" },
    "BODY WORK":  { "present": true, "damage": false, "comments": "" }
  },

  "general_comments": "Free-text summary of overall unit condition."
}
```

### Checklist item fields

| Field | Type | Description |
|---|---|---|
| `present` | boolean | Whether the item was found on the unit |
| `damage` | boolean | Whether damage was noted |
| `comments` | string | Free-text notes; leave `""` if none |

### Report status indicators

| Symbol | Colour | Meaning |
|---|---|---|
| ✓ | Green | Present, no damage |
| ! | Amber | Present, damage noted |
| ✗ | Red | Not present |

---

## Project Structure

```
inspection-tool/
├── generate_report.py          # main script
├── config.yaml                 # paths and report settings
├── inspections/
│   └── BA3036/
│       ├── inspection_data.json
│       ├── inspection_checklist.md
│       └── photos/
└── output/
```

---

## Notes

- Photos are auto-scaled to fit the grid cell while preserving aspect ratio.
- If a unit has more photos than `photos_per_slide`, the script creates additional photo slides automatically.
- The checklist item order in the report is driven by `checklist_items` in `config.yaml`, not the JSON key order.
- `inspection_checklist.md` is an optional human-readable reference. It is not read by the script.
