# GST Master Data Builder

A simple local web tool that reads invoice Excel files and fills the same GST master template format automatically.

## What it does
- Upload one or more invoice Excel files
- Keeps the same template tabs, headers and structure
- Fills these tabs:
  - GST_Master_Data
  - b2cs
  - hsn(b2c)
  - docs
- Leaves other tabs intact

## Best for
This version is tuned for the current Varkesh invoice Excel format shared in the sample files.

## Files included
- `app.py` - Streamlit app
- `template_clean.xlsx` - clean built-in master template
- `requirements.txt` - Python packages
- `run_tool.bat` - easy start for Windows

## How to run
1. Install Python 3.10 or newer
2. Open Command Prompt in this folder
3. Run:
   `pip install -r requirements.txt`
4. Then run:
   `streamlit run app.py`

Or on Windows double-click:
- `run_tool.bat`

## How to use
1. Upload invoice Excel files
2. Optionally upload your own master template workbook
3. Click `Generate GST Master Workbook`
4. Download the generated file

## Notes
- If you upload your own template, tab names and headers are preserved
- The app clears old data rows before filling, while keeping the template format
- Current logic assumes recipient GSTIN is blank unless available in the invoice file
- Current logic is designed for the invoice layout already shared by the user

## Future upgrade ideas
- drag and drop folder upload
- batch zip export
- party master
- automatic duplicate check
- GSTIN-based B2B/B2C split
- more invoice formats
