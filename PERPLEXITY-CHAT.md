py


CONVERSATION.md
194 lines

md


config.yaml
33 lines

yaml

This script works
i require a fix for when the pptx is created, the output .pptx file is large
the ops manager during the manual process, drags the photo into the report and manually resizes the photo.
the photos when inserted into the report must be scale height 41%, scale width 35%. lock aspect ration, relative to original picture size.
these are best options, yet can be adjusted to fit the report page

09:21
Claude responded: Got it.

Architected image compression strategy to reduce PPTX file size
Architected image compression strategy to reduce PPTX file size
Got it. The issue is clear: add_picture embeds the full-resolution original (potentially 4032×3024 from a smartphone) into the PPTX. The fix is to pre-resize with Pillow to the actual display dimensions (derived from your cell grid) before embedding — the ops manager's 41%/35% scale is essentially doing this manually. We'll do it programmatically.


Architected image resizing solution to reduce file size and match manual scaling percentages
Architected image resizing solution to reduce file size and match manual scaling percentages
No PPTX skill needed — I know exactly what to change. Writing the fix now.


Updated generate_report.py with photo pre-resizing for smaller PPTX output
Updated generate_report.py with photo pre-resizing for smaller PPTX output
Now the updated config with the two new tuning parameters:


Created a file, read a file
Created a file, read a file
What changed and why:

Root cause of the large PPTX: add_picture(photo_path, ...) embedded the raw original image — a smartphone photo at 4032×3024 px (~5 MB each) — regardless of how small it displays in the slide. 10 photos = ~50 MB before PowerPoint even opens.

The fix — _safe_photo_stream(): Before embedding, Pillow now:

Opens the image and corrects EXIF rotation (so photos aren't sideways)
Downscales to the actual cell display dimensions (in pixels at 96 DPI) × photo_scale_factor
Saves as compressed JPEG into a BytesIO buffer
That buffer is what gets embedded — not the original file
The connection to the ops manager's 41%/35% workflow: That manual scale was compensating for exactly this — PPT was embedding the full image and they were visually shrinking it. Now the resize happens at the data level before insertion, so the embedded image is already the right size.

Command to run:

python generate_report.py --ba BA2797
python generate_report.py --ba BA2797 --config config.yaml
Tuning knobs in config.yaml (no code changes needed):

Setting	Default	Effect
photo_quality	82	Lower to 70 for max compression
photo_scale_factor	1.5	Lower to 1.0 if still too large; raise to 2.0 if photos look soft
photos_per_slide	10	Fewer photos = larger cells = larger embedded images
Expected result: a report with 10 smartphone photos should drop from ~50 MB to ~2–5 MB.


Generate report
PY 

Config
YAML 
