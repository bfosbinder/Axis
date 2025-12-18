Axis Inspection & Ballooning App (Windows Release)

Axis is a standalone Windows application for aerospace-quality inspection and drawing ballooning.
It loads engineering PDFs, allows precise balloon placement, auto-generates inspection tables, and exports both inspection reports and ballooned PDFs.

This Windows build is created automatically via GitHub Actions and requires no Python installation.

🚀 Features
Ballooning Mode

Click-and-drag to place balloons on any PDF drawing.

Move balloons freely; positions are saved automatically.

Auto-assigns balloon numbers.

Stores:

Page

Rectangle bounds

Zoom level

Balloon offset

Nominal / LSL / USL

Method

Full undo support for added/deleted balloons.

Inspection Mode

Select a Work Order / Serial Number.

Enter measurement results.

Status updates automatically (PASS/FAIL).

Supports “Pass/Fail” text or numeric values.

Remembers results per work order.

OCR (stub)

OCR button placeholder for future integration.

PDF Export

Creates two PDFs:

_inspection.pdf — formatted table of results

_ballooned.pdf — original drawing with balloons drawn on top

📦 Windows Download

Compiled .exe builds are available under:

GitHub → Actions → “Build Windows EXE” → Artifacts

Download:

AxisInspector.exe


No installation required.
Just run the executable.

🖥️ System Requirements

Windows 10 or 11 (64-bit)

No Python needed

PDF viewer installed (for exported documents)

Recommended: 1080p or larger monitor

📁 How Data Is Stored

Axis now keeps everything (balloons + inspection results) in a single SQLite file next to your PDF:

yourfile.pdf

yourfile.axis.db

The first time the `.axis.db` file is created, any legacy `*.balloons.csv` and work-order CSV files are imported automatically so existing projects continue to work.

🛠️ How to Use
1. Open a PDF

File → Open PDF

PDF renders on the right panel

Table of features appears on the left

2. Start a Session

You will be prompted to choose:

Ballooning Mode → create/edit balloons

Inspection Mode → enter results

You can change sessions later from the toolbar.

3. Add Balloons (Ballooning Mode)

Enable Pick-on-Print

Click-and-drag a rectangle around the feature

A balloon appears automatically

Adjust size using the “Balloon size” control

Move the balloon by dragging it

Undo commands available:

Press Ctrl+Z

Or use the toolbar shortcut

4. Enter Tolerances

In Ballooning Mode, you can type:

1.203 +/-.003


or

1.002 +.010 -.000


The app parses and auto-fills:

Nominal

LSL

USL

5. Inspection

In Inspection Mode, enter results directly in the “Result” column.

Status updates automatically:

Value within range → PASS

Out of range → FAIL

You may also type:

P → Pass

F → Fail

6. Export PDFs

Export PDFs produces:

drawingname_inspection.pdf

drawingname_ballooned.pdf

Saved to your chosen directory.

⌨️ Hotkeys
Action	Key
Toggle Pick Mode	P
Toggle Grab Mode	G
Undo	Ctrl+Z
Fit-to-View	Button
Select next result	Auto after entry
📚 Technology Used

PyQt6 — UI

PyMuPDF (fitz) — PDF rendering/drawing

QGraphicsView — balloon overlay system

GitHub Actions — automatic Windows builds

PyInstaller — EXE packaging

🧪 Development Version (Python)

If you choose to run the Python version:

pip install -r requirements.txt
python main.py

📄 License

MIT License (or your preferred license)

🙌 Contributions

PRs welcome!
Found a bug in the Windows EXE? Open an issue describing:

OS version

PDF type

Steps to reproduce

Screenshot if possible
