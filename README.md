# AGT Inventory & Label Generator

A cross-platform (Windows/macOS) Python/Tkinter application that:

1. Loads product data from Excel or JSON URLs  
2. Applies filters and lets you select products  
3. Generates:
   - 3×3 horizontal or vertical Word tag sheets  
   - 5×6 “mini” tags  
   - 2×2 inventory slips  

It also supports:
- On-the-fly lineage (Sativa/Indica/Hybrid/etc.) correction via dropdowns  
- Automatic unit conversions (g → oz for edibles)  
- Colored cell shading for cannabinoids, paraphernalia, etc.  
- Duplex-back vendor/brand printing  
- Right-click “Paste” in JSON URL entry  
- Adjustable complexity/scale slider  
- Fast multi-process rendering and smart template expansion  

---

## Features

- **Excel or JSON ingestion**: point at a POSaBit export or transfer-API URL.  
- **Dynamic filters**: vendor, brand, type, lineage, strain, weight.  
- **Selected-tag UI**: move items Available ↔ Selected, with “Select All” and undo.  
- **Lineage fixer**: correct strain lineage in bulk and log changes.  
- **Tag generation**: Word `.docx` output, smart autosizing, conditional formatting, cut-guide lines.  
- **Inventory slips**: 2×2 labels with vendor/backside duplex pages.  
- **Cross-platform**: Windows (via COM) and macOS (via AppleScript) support for overwriting open Word docs.  
- **Packaging**: can be bundled as a standalone exe/app via PyInstaller or similar.  

---

## Requirements

- **Python 3.8+**  
- **Tkinter** (should ship with standard Python)  
- **pandas**, **openpyxl**  
- **python-docx**, **docxtpl**, **docxcompose**  
- **pywin32** (Windows only, for closing open Word docs)  
- **macOS**: AppleScript support (bundled)  

Install dependencies with:

```bash
pip install pandas openpyxl python-docx docxtpl docxcompose pywin32
