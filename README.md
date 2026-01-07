Axis Inspection & Ballooning App
================================

Axis is a production-ready Windows application for aerospace-quality inspection and drawing ballooning. It renders manufacturing PDFs, lets you place and edit precision balloons, captures inspection data by work order, and exports both ballooned drawings and formatted inspection reports.

**Project page:** https://bfosbinder.github.io/Axis/


---

Key Capabilities
----------------

* **Guided ballooning** – drag a rectangle on the drawing to create an auto-numbered balloon with stored page, bounds, zoom, offsets, and specs. Balloons are freely movable and can be resized in bulk.
* **Inspection workflows** – choose Ballooning or Inspection mode per session. Work orders/serials retain their own measurement data, and PASS/FAIL status is recalculated automatically for numeric or text inputs.
* **High-fidelity rendering** – PyMuPDF powers crisp PDF drawing at varying zoom levels; a QGraphicsView overlay keeps balloons and highlight rectangles aligned to document coordinates.
* **Reporting** – export filtered inspection data to CSV, dump every recorded work order to a master CSV, and generate both `*_inspection.pdf` (tabulated results) and `*_ballooned.pdf` (drawing with annotations). Exports honor current filters and status formatting.
* **Data integrity** – all feature metadata and inspection results live inside a single SQLite database beside the source PDF, with automatic migration from legacy CSV files.

---

Installation
------------

### Windows executable (preferred)
Download the latest Windows build from **GitHub Releases**:

1. Go to **https://github.com/bfosbinder/Axis/releases/latest**
   *(or download directly from https://github.com/bfosbinder/Axis/releases/download/latest/AxisInspector.exe)*.
2. Download `AxisInspector.exe`.
3. Run the executable directly — no installer and no Python runtime required.


System requirements: Windows 10/11 (64-bit), a PDF viewer for exported files, and preferably a 1080p+ display.

### Development build (Python)
```bash
python3 -m venv .venv
source .venv/bin/activate        # use .venv\Scripts\activate on Windows
pip install -r requirements.txt
python main.py
```
Optional features: Pillow + pytesseract for OCR experiments.

---

Product Workflow
----------------

1. **Open a PDF** – `File → Open PDF`. The drawing loads on the right, the feature table on the left.
2. **Choose a session** – the dialog prompts for Ballooning (edit geometry/specs) or Inspection (record measurements). You can switch sessions later via the toolbar action.
3. **Ballooning mode**
	* Enable *Pick-on-Print*, drag to define the feature bounds, and a balloon appears immediately.
	* Adjust global balloon size, drag balloons to fine‑tune offsets, and rely on Ctrl+Z for undo.
	* Enter tolerance strings (e.g., `1.203 +/-.003`) to auto-fill Nominal/LSL/USL.
4. **Inspection mode**
	* Select a Work Order / Serial Number.
	* Enter numeric results or quick text (`P` / `F`). Status columns update in real time using stored specs.
	* Results persist per work order so you can resume later.
5. **Export** – use *Export Results* for a CSV snapshot of the current table (respecting filters), *Export All Results* for every work order/serial on the part, or *Export PDFs* for the formatted inspection table plus ballooned drawing.

---

Hotkeys
-------

| Action                | Shortcut |
|-----------------------|----------|
| Toggle Pick Mode      | `P`      |
| Toggle Grab Mode      | `G`      |
| Undo                  | `Ctrl+Z` |
| Fit drawing to view   | Toolbar  |
| Advance to next result| Auto after entry |

---

Data Storage
------------

Each PDF gains a sibling SQLite database: `yourfile.axis.db`. The database contains:

* `features` – rows matching the legacy `MASTER_HEADER` schema (page, bounds, zoom, method, specs, offsets, radius).
* `results` – measurement values per `(feature_id, workorder)` with timestamps.

When the `.axis.db` file is created, the app automatically imports any `*.balloons.csv` and `<pdf>.<workorder>.csv` files it finds, so existing projects remain intact.

---

Technology Stack
----------------

* **PyQt6** – desktop UI, table widgets, and dialogs.
* **PyMuPDF (fitz)** – fast PDF rendering with high-DPI support.
* **QGraphicsView** – custom overlay for balloons and highlight rectangles.
* **SQLite** – durable storage for features and inspection data.
* **GitHub Actions + PyInstaller** – automated Windows builds and distribution.

---

Support & Contributions
-----------------------

Issues and pull requests are welcome. When reporting a bug, please include:

* Windows version and GPU (if applicable)
* PDF characteristics (page size, color usage, etc.)
* Exact steps to reproduce
* Screenshots or exported PDFs highlighting the problem

License: MIT (or your selected alternative).*** End Patch
