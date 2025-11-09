# Axis — PDF Ballooning & Inspection App (PyQt6)

A fast, shop-friendly viewer to **balloon blueprints** and **record inspection results** right on top of PDFs. Built with PyQt6 + PyMuPDF.

---

## Highlights

* **Two modes**

  * **Ballooning:** drag-select features (“pick-on-print”), auto-numbered balloons, set methods/specs.
  * **Inspection:** enter results per Work Order/Serial; auto PASS/FAIL.
* **On-print picking:** rubber-band a rectangle; a balloon snaps near the picked box.
* **Smart tolerances:** type a tolerance expression once (e.g., `1.25 ±0.05`), it fills **Nominal/LSL/USL**.
* **Filters & status coloring:** filter by Method or Status; result cells color PASS/FAIL.
* **Quick focus:** selecting a row jumps/zooms to that feature and highlights its pick box.
* **Adjustable balloons:** one control to resize all balloons (and persist the size).
* **CSV export:** export the filtered table for reports.

---

## Install

> Python 3.10+ recommended.

```bash
# 1) Create a virtual env
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate

# 2) Install dependencies
pip install PyMuPDF PyQt6

# 3) Put your files together
# - axis.py               (this app; name it however you like)
# - storage.py            (your storage helpers used by the app)
```

**Dependencies**

* [PyMuPDF (`fitz`)](https://pymupdf.readthedocs.io/) — PDF rendering
* [PyQt6](https://pypi.org/project/PyQt6/) — UI

---

## Run

```bash
python axis.py
```

> If you named the file differently (e.g., `app.py`), run that file instead.

---

## How it works (files it writes)

When you open `SomeDrawing.pdf`, the app creates data **next to the PDF**:

* **`SomeDrawing.pdf.balloons.csv`** — the **master balloon list** for that drawing
  Typical columns (managed by the app):
  `id, page, x, y, w, h, bx, by, br, zoom, method, nominal, lsl, usl`
* **Work Order/Serial results** — stored via `storage.py` (per WO/serial for that PDF).
  The app reads/writes these through:

  * `read_wo(pdf_path, work_order)` / `write_wo(pdf_path, work_order, results_dict)`
  * Work-order values are shown in the **Result** column in **Inspection** mode.

> The exact folder/filenames are delegated to your `storage.py`. This README assumes your `storage.py` keeps them adjacent to the PDF.

---

## Using the app

1. **Open PDF** → toolbar **Open PDF**.
2. If a master already exists, you’ll be prompted:

   * **Start Ballooning** (edit features/specs), or
   * **Start Inspection** (enter a Serial/WO to record results).
3. **Ballooning mode**

   * Toggle **Pick-on-Print** (or press **P**), then **drag** a rectangle on the print to add a feature.
   * Edit **Method/Nominal/LSL/USL** in the table (persists to master).
   * Type a one-line tolerance expression in **Result** (e.g., `10 ±0.1`, `Ø5 +0.02/-0.01`); the app parses it and fills **Nominal/LSL/USL**, then clears Result.
   * Right-click a row → **Delete Balloon** to remove it.
   * **Balloon size** spinbox changes all balloons; saved to master.
4. **Inspection mode**

   * Enter **Result** per feature (number or `Pass`/`Fail`).
   * PASS/FAIL computed if numeric and limits exist; cells tint green/red.
   * Results are saved to the current **WO/Serial** via `write_wo`.
5. **Filtering & Export**

   * **Method** text filter and **Status** dropdown filter the table.
   * **Export Filtered CSV** writes the visible rows (`ID, Page, Method, Result, Nominal, LSL, USL, Status`).

---

## Controls & Shortcuts

* **Pick-on-Print** button (or **P**) — enter pick mode to place balloons.
* **Grab/Scroll** mode (or **G**) — pan the view.
* **Mouse wheel** — smooth zoom (centers at cursor).
* **Fit** — fit page to window.
* **Page** spinbox — jump between pages.
* **Show/Hide Balloons** — toggle visibility while keeping selection logic.
* **Selecting a row** — jumps/zooms to that feature and shows a red highlight box.

---

## Table columns (what’s editable when)

| Column          | Ballooning                | Inspection | Notes                                                                             |
| --------------- | ------------------------- | ---------- | --------------------------------------------------------------------------------- |
| ID, Page        | read-only                 | read-only  | Auto-assigned by master                                                           |
| Method          | editable                  | read-only  | Edited via the in-cell combo                                                      |
| Result          | editable (parsing helper) | editable   | In Ballooning, used to parse tolerances then cleared; in Inspection, saved per WO |
| Nominal/LSL/USL | editable                  | read-only  | Persisted to master in Ballooning                                                 |
| Status          | read-only                 | read-only  | Auto from Result/limits or Pass/Fail keywords                                     |

**Result normalization:** entering `P`/`Pass` becomes `Pass`; `F`/`Fail` becomes `FAIL`.

---

## Data model (at a glance)

Each feature in the master has:

* **page**: 1-based page index
* **x,y,w,h**: picked rectangle (scene coords)
* **bx,by**: balloon offset relative to the rect center
* **br**: balloon radius
* **zoom**: last zoom used when focused
* **method, nominal, lsl, usl**: specs

Inspection results are a mapping `{ feature_id: result_string }` per **WO/Serial**.

---

## Packaging (optional)

Create a single-file executable with PyInstaller:

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --name Axis --windowed axis.py
```

> On Linux with Wayland/HiDPI, you can experiment with:
> `QT_AUTO_SCREEN_SCALE_FACTOR=1 python axis.py`

---

## Troubleshooting

* **PyMuPDF install issues**: ensure system has build tools; try `pip install --upgrade pip wheel setuptools`.
* **Qt platform plugin errors**: verify the virtual env is active; try `pip uninstall PyQt6 && pip install PyQt6`.
* **Nothing happens on pick**: make sure you’re in **Ballooning** mode, **Pick-on-Print** is ON, and your drag box is at least **5×5 px**.

---

## Contributing

Issues and PRs welcome. Keep the UI fast, keyboard-first, and safe for shop-floor use.

---

## License

MIT (or your preferred license).
