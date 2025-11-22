# Axis App – AI Implementation Guide

## Project Overview
- PyQt6 desktop app (`main.py`) draws PDFs via PyMuPDF (`fitz`) and overlays editable "balloons" in a shared scene.
- `storage.py` is the only other core module: handles CSV persistence (`*.balloons.csv` + workorder CSVs) and tolerance parsing helpers.
- UI flows: start dialog → choose Ballooning or Inspection → left table + filters + result entry, right `PDFView` showing rendered page with balloons/highlight.
- Toolbar buttons wire directly into `MainWindow` slots; keep behaviour consistent with existing patterns (block signals, update status bar, refresh table selectively).

## Rendering & Zoom
- Scene coordinates stay fixed at PDF points × `PAGE_RENDER_ZOOM`; never change balloon geometry when re-rendering.
- `PDFView` caches `_current_scale`, `_last_render_scale`; actual re-render is deferred via `_schedule_rerender_for_zoom` (100 ms single-shot timer) to keep wheel zoom smooth.
- `_render_current_page(render_scale)` clamps real render factor to ≤3× DPR and only rebuilds balloons when page changes. Maintain these guards when touching zoom logic.

## Balloon Lifecycle
- Adding balloons: `_rect_picked` → `storage.add_feature` → append to `self.rows` → `_push_undo` for Ctrl+Z.
- Deleting balloons funnels through `_remove_feature`, which handles undo snapshots, table updates, scene cleanup.
- Any feature mutations must update both `self.rows` and the persisted CSV via `storage.update_feature` / `write_master`.
- Highlight rectangle comes from `_ensure_highlight_rect`; only adjust geometry, never recreate the scene rect logic.

## Tables & Filters
- Table rebuilds call `refresh_table`; honor existing signal blocking and filter checks (method/status combos).
- Inspection mode writes to workorder CSV via `read_wo` / `write_wo` and recomputes row status inline; reuse `_recompute_row_status` for consistency.
- Auto-advance editing uses `_advance_result_edit`; leave timer-based focus intact when adding result-related features.

## Persistence & Files
- Master CSV path: `<pdf>.balloons.csv`; work orders: `<pdf>.<WO>.csv`. Keep schema aligned with `storage.MASTER_HEADER`.
- `ensure_master` called in `_load_pdf`; if new feature data is needed, extend header there and adjust CSV helpers accordingly.

## Development Workflow
- Run locally: `python3 -m venv .venv && source .venv/bin/activate`, install `pip install -r requirements.txt`, start app with `python main.py`.
- No automated tests; manual execution is the validation path.
- GitHub Actions builds Windows EXE (see README); avoid breaking PyInstaller compatibility (no dynamic imports without guard).

## Style & Conventions
- Stick to ASCII for UI strings and comments; minimal, purposeful comments only.
- Prefer small helper methods inside `MainWindow` rather than new modules unless functionality is clearly reusable elsewhere.
- Before changing selection or mode state, ensure `_suppress_auto_focus`, `_balloons_built_for_page`, and undo stacks stay in sync.
- Use existing status bar messaging & `QtWidgets.QMessageBox` patterns for user feedback.

Feedback welcome: flag confusing areas or missing workflows so we can extend this guide.
