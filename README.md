# 📊 Commodity Code Analyzer (TRADER EDITION)

This Python script processes an Excel file with trade/export data, groups the results by commodity code (CN/Intrastat), and categorizes them based on the `BTOM` column (`low` or `medium`). If `medium`, the script also breaks results down by `EHC` values. It outputs a clean multi-sheet Excel file.

---

## ✅ Features

- Group by commodity code prefix using a list from `CommodityCodes.txt`
- Process any Excel file (customizable column names)
- Separate logic for `low` and `medium` BTOM categories
- Explodes and groups by multiple EHC values
- Automatically generates Excel output with grouped summaries
- Easy-to-edit column config at the top of the script

---

## 📂 Project Structure
project/
├── excel_processor.py # Main script
├── CommodityCodes.txt # Comma-separated list of code prefixes (e.g., 0401, 0402, ...)
├── your_file.xls # Your input Excel file
├── grouped_result.xlsx # Auto-generated result file
