import sys
import os
import re
import io
import csv
from functools import partial
from pathlib import Path
from datetime import datetime
import fitz  # PyMuPDF
from PyQt6 import QtWidgets, QtGui, QtCore
from storage import ensure_master, read_master, add_feature, update_feature, delete_feature, read_wo, write_wo, write_master, parse_tolerance_expression, normalize_str, MASTER_HEADER, list_workorders
import spc

try:
    from PIL import Image  # type: ignore
except ImportError:  # Pillow is optional until OCR runs
    Image = None


BALLOON_RADIUS = 14
PAGE_RENDER_ZOOM = 1.5


def format_number(value: float, decimals: int = 6) -> str:
    text = f"{value:.{decimals}f}"
    if '.' in text:
        text = text.rstrip('0').rstrip('.')
    if text == '-0':
        text = '0'
    return text


class StartSessionDialog(QtWidgets.QDialog):
    """Prompt for Serial Number (Inspection) or start Ballooning mode."""
    def __init__(self, parent=None, previous_orders=None, allow_ballooning=True):
        super().__init__(parent)
        self.setWindowTitle('Start Session')
        self.choice = 'cancel'  # 'inspection' | 'ballooning' | 'cancel'
        self.serial = ''
        self._previous_orders = list(previous_orders or [])
        lay = QtWidgets.QVBoxLayout(self)
        lay.addWidget(QtWidgets.QLabel('Enter Serial Number for Inspection (or choose Ballooning):'))
        if self._previous_orders:
            prev_label = QtWidgets.QLabel('Previous sessions:')
            lay.addWidget(prev_label)
            self.previous_combo = QtWidgets.QComboBox()
            self.previous_combo.addItem('Select previous...')
            for order in self._previous_orders:
                self.previous_combo.addItem(order)
            self.previous_combo.currentIndexChanged.connect(self._previous_selected)
            lay.addWidget(self.previous_combo)
        else:
            self.previous_combo = None
        self.edit = QtWidgets.QLineEdit()
        placeholder = 'Serial / Work Order'
        if self._previous_orders:
            placeholder += ' (or pick above)'
        self.edit.setPlaceholderText(placeholder)
        lay.addWidget(self.edit)
        btns = QtWidgets.QHBoxLayout()
        self.btn_inspect = QtWidgets.QPushButton('Start Inspection')
        self.btn_balloon = QtWidgets.QPushButton('Start Ballooning')
        self.btn_cancel = QtWidgets.QPushButton('Cancel')
        btns.addWidget(self.btn_inspect)
        if allow_ballooning:
            btns.addWidget(self.btn_balloon)
        btns.addWidget(self.btn_cancel)
        lay.addLayout(btns)

        self.btn_inspect.clicked.connect(self._go_inspect)
        if allow_ballooning:
            self.btn_balloon.clicked.connect(self._go_balloon)
        self.btn_cancel.clicked.connect(self.reject)

    def _go_inspect(self):
        text = self.edit.text().strip()
        if not text:
            QtWidgets.QMessageBox.warning(self, 'Required', 'Please enter a Serial / Work Order for Inspection.')
            return
        self.choice = 'inspection'
        self.serial = text
        self.accept()

    def _go_balloon(self):
        self.choice = 'ballooning'
        self.accept()

    def _previous_selected(self, index: int):
        if index <= 0:
            return
        value = self._previous_orders[index - 1]
        self.edit.setText(value)


class MethodListDialog(QtWidgets.QDialog):
    """Dialog to manage available inspection methods."""

    def __init__(self, parent=None, methods=None):
        super().__init__(parent)
        self.setWindowTitle('Inspection Methods')
        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(QtWidgets.QLabel('Double-click a method to rename it.'))

        self.list_widget = QtWidgets.QListWidget()
        self.list_widget.setEditTriggers(
            QtWidgets.QAbstractItemView.EditTrigger.DoubleClicked
            | QtWidgets.QAbstractItemView.EditTrigger.EditKeyPressed
        )
        for method in methods or []:
            if method:
                self.list_widget.addItem(method)
        layout.addWidget(self.list_widget)

        entry_layout = QtWidgets.QHBoxLayout()
        self.input = QtWidgets.QLineEdit()
        self.input.setPlaceholderText('Add new method')
        entry_layout.addWidget(self.input)
        add_btn = QtWidgets.QPushButton('Add')
        add_btn.clicked.connect(self._add_from_input)
        entry_layout.addWidget(add_btn)
        layout.addLayout(entry_layout)

        remove_btn = QtWidgets.QPushButton('Remove Selected')
        remove_btn.clicked.connect(self._remove_selected)
        layout.addWidget(remove_btn)

        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.StandardButton.Ok | QtWidgets.QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.input.returnPressed.connect(self._add_from_input)

    def _add_from_input(self):
        text = self.input.text().strip()
        if not text:
            return
        if not self._contains(text):
            self.list_widget.addItem(text)
        self.input.clear()
        self.input.setFocus()

    def _remove_selected(self):
        for item in self.list_widget.selectedItems():
            row = self.list_widget.row(item)
            self.list_widget.takeItem(row)

    def _contains(self, text: str) -> bool:
        for idx in range(self.list_widget.count()):
            if self.list_widget.item(idx).text().strip().lower() == text.lower():
                return True
        return False

    def get_methods(self):
        values = []
        seen = set()
        for idx in range(self.list_widget.count()):
            text = self.list_widget.item(idx).text().strip()
            if not text:
                continue
            key = text.lower()
            if key in seen:
                continue
            seen.add(key)
            values.append(text)
        return values


class BalloonItem(QtWidgets.QGraphicsEllipseItem):
    def __init__(self, feature: dict, image_item: QtWidgets.QGraphicsPixmapItem, parent=None):
        x = float(feature.get('x', 0))
        y = float(feature.get('y', 0))
        w = float(feature.get('w', 10))
        h = float(feature.get('h', 10))
        bx = float(feature.get('bx', 0))
        by = float(feature.get('by', 0))
        radius = float(feature.get('br', BALLOON_RADIUS))
        # base center of the picked rectangle
        cx = x + w / 2.0
        cy = y + h / 2.0
        # ellipse rect centered at (0,0), then position item at center+offsets
        super().__init__(-radius, -radius, radius * 2.0, radius * 2.0)
        self.setPos(cx + bx, cy + by)
        # render balloons with a red theme while keeping legibility against the page
        circle_color = QtGui.QColor(220, 40, 40)
        outline = QtGui.QPen(circle_color, 2)
        self.setBrush(QtGui.QBrush(QtGui.QColor(255, 230, 230)))
        self.setPen(outline)
        self.setZValue(10)
        self.setFlag(QtWidgets.QGraphicsItem.GraphicsItemFlag.ItemIsMovable, True)
        self.setFlag(QtWidgets.QGraphicsItem.GraphicsItemFlag.ItemSendsGeometryChanges, True)
        self.feature = feature
        self._save_timer = QtCore.QTimer()
        self._save_timer.setSingleShot(True)
        self._save_timer.timeout.connect(self._persist_offset)
        self.radius = radius
        # number text inside balloon
        full_id = feature.get('id', '')
        num_only = full_id
        m = QtCore.QRegularExpression(r"(\d+)$").match(full_id)
        if m.hasMatch():
            num_only = m.captured(1)
        self.text_item = QtWidgets.QGraphicsSimpleTextItem(num_only, self)
        font = QtGui.QFont()
        font.setBold(True)
        self.text_item.setFont(font)
        self.text_item.setBrush(QtGui.QBrush(circle_color))
        self._update_text_appearance()
        # tooltip with full id
        self.setToolTip(full_id)

    def set_radius(self, radius: float):
        self.radius = radius
        self.prepareGeometryChange()
        self.setRect(-radius, -radius, radius * 2.0, radius * 2.0)
        self._update_text_appearance()
        self.feature['br'] = str(radius)
        self.update()

    def _update_text_appearance(self):
        # scale text proportionally so balloon numbers stay legible
        point_size = max(6, int(round(self.radius * 0.7)))
        font = self.text_item.font()
        font.setPointSize(point_size)
        self.text_item.setFont(font)
        br = self.text_item.boundingRect()
        self.text_item.setPos(-br.width() / 2.0, -br.height() / 2.0)

    def itemChange(self, change, value):
        if change == QtWidgets.QGraphicsItem.GraphicsItemChange.ItemPositionHasChanged:
            # compute new bx,by relative to rect center
            try:
                x = float(self.feature.get('x', 0))
                y = float(self.feature.get('y', 0))
                w = float(self.feature.get('w', 10))
                h = float(self.feature.get('h', 10))
                cx = x + w/2
                cy = y + h/2
                center = self.pos()
                bx = center.x() - cx
                by = center.y() - cy
                self.feature['bx'] = str(bx)
                self.feature['by'] = str(by)
                if self._save_timer.isActive():
                    self._save_timer.stop()
                self._save_timer.start(150)
            except Exception:
                pass
        return super().itemChange(change, value)

    def _persist_offset(self):
        try:
            pdf_path = self.feature.get('_pdf')
            fid = self.feature.get('id')
            if not pdf_path or not fid:
                return
            update_feature(pdf_path, fid, {
                "bx": self.feature.get('bx', '0'),
                "by": self.feature.get('by', '0'),
            })
        except Exception:
            pass

    def mouseReleaseEvent(self, event):
        if self._save_timer.isActive():
            self._save_timer.stop()
        self._persist_offset()
        super().mouseReleaseEvent(event)


class PDFView(QtWidgets.QGraphicsView):
    rectPicked = QtCore.pyqtSignal(QtCore.QRect)

    def __init__(self, controller, parent=None):
        super().__init__(parent)
        self.controller = controller  # explicit reference to MainWindow
        self.setScene(QtWidgets.QGraphicsScene(self))
        self._pixmap_item = None
        self._rubber = QtWidgets.QRubberBand(QtWidgets.QRubberBand.Shape.Rectangle, self)
        self._rubber_origin = None
        self._current_scale = 1.0
        self._min_scale = 0.2
        self._max_scale = 40.0
        self._panning = False
        self._pan_button = None
        self._last_pan_pos = None
        self.setRenderHints(
            QtGui.QPainter.RenderHint.Antialiasing | QtGui.QPainter.RenderHint.SmoothPixmapTransform
        )
        self.setDragMode(QtWidgets.QGraphicsView.DragMode.ScrollHandDrag)
        self.setTransformationAnchor(QtWidgets.QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setResizeAnchor(QtWidgets.QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self._last_render_scale = None  # [CRISP-ZOOM]
        self.setCacheMode(QtWidgets.QGraphicsView.CacheModeFlag.CacheBackground)  # [ZOOM-DEBOUNCE]
        self.setViewportUpdateMode(QtWidgets.QGraphicsView.ViewportUpdateMode.SmartViewportUpdate)  # [ZOOM-DEBOUNCE]

    def load_page(self, qimage: QtGui.QImage, scene_size: tuple[float, float] | None = None):
        scene = self.scene()
        if scene is None:
            return
        if self._pixmap_item is not None:
            try:
                current_scene = self._pixmap_item.scene()
                target_scene = current_scene or scene
                target_scene.removeItem(self._pixmap_item)
            except Exception:
                pass
            self._pixmap_item = None
        pix = QtGui.QPixmap.fromImage(qimage)
        self._pixmap_item = scene.addPixmap(pix)
        if scene_size is not None:
            width, height = scene_size
            try:
                dpr = float(pix.devicePixelRatio())
            except Exception:
                dpr = 1.0
            logical_width = pix.width() / dpr
            logical_height = pix.height() / dpr
            if logical_width > 0 and logical_height > 0:
                scale = width / logical_width
                self._pixmap_item.setScale(scale)  # [CRISP-ZOOM]
            self.setSceneRect(QtCore.QRectF(0, 0, width, height))
        else:
            self.setSceneRect(self._pixmap_item.boundingRect())

    def fit_to_view(self):
        if self.scene() is None or self.sceneRect().isNull():
            return
        self.fitInView(self.sceneRect(), QtCore.Qt.AspectRatioMode.KeepAspectRatio)
        # keep _current_scale aligned with the actual view transform
        self._current_scale = self.transform().m11()
        self._last_render_scale = self._current_scale  # [CRISP-ZOOM]

    def wheelEvent(self, event: QtGui.QWheelEvent):
        # zoom towards mouse position
        if event.modifiers() & QtCore.Qt.KeyboardModifier.ControlModifier:
            # allow Ctrl+wheel for future; for now treat same
            pass
        angleDeltaY = event.angleDelta().y()
        if angleDeltaY == 0:
            return super().wheelEvent(event)
        factor = 1.0015 ** angleDeltaY  # smooth zoom
        new_scale = self._current_scale * factor
        new_scale = max(self._min_scale, min(self._max_scale, new_scale))
        factor = new_scale / self._current_scale
        if factor == 1.0:
            return
        self.scale(factor, factor)
        self._current_scale = self.transform().m11()
        event.accept()
        
        # keep rubber band consistent
        if self._rubber.isVisible():
            self._rubber.hide()
        if hasattr(self.controller, '_schedule_rerender_for_zoom'):
            self.controller._schedule_rerender_for_zoom(self._current_scale)  # [ZOOM-DEBOUNCE]

    def set_zoom(self, target: float, focus_point: QtCore.QPointF = None):
        if self._pixmap_item is None:
            return
        target = max(self._min_scale, min(self._max_scale, float(target)))
        prev_anchor = self.transformationAnchor()
        if focus_point is None:
            focus_point = self.mapToScene(self.viewport().rect().center())
        self.setTransformationAnchor(QtWidgets.QGraphicsView.ViewportAnchor.AnchorViewCenter)
        self.resetTransform()
        self.scale(target, target)
        self._current_scale = self.transform().m11()
        self.centerOn(focus_point)
        self.setTransformationAnchor(prev_anchor)
        if focus_point is not None:
            self.ensureVisible(QtCore.QRectF(focus_point.x() - 20, focus_point.y() - 20, 40, 40), 10, 10)
        if hasattr(self.controller, '_schedule_rerender_for_zoom'):
            self.controller._schedule_rerender_for_zoom(self._current_scale)  # [ZOOM-DEBOUNCE]

    def mousePressEvent(self, event):
        if self.controller.mode == 'Ballooning' and self.controller.pick_on_print:
            if event.button() == QtCore.Qt.MouseButton.LeftButton:
                self._rubber_origin = event.pos()
                self._rubber.setGeometry(QtCore.QRect(self._rubber_origin, QtCore.QSize()))
                self._rubber.show()
                return
            if event.button() in (QtCore.Qt.MouseButton.MiddleButton, QtCore.Qt.MouseButton.RightButton):
                # allow panning with alternate buttons while pick mode is active
                self._panning = True
                self._pan_button = event.button()
                self._last_pan_pos = event.pos()
                self.setCursor(QtCore.Qt.CursorShape.ClosedHandCursor)
                event.accept()
                return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._panning:
            if self._last_pan_pos is not None:
                delta = event.pos() - self._last_pan_pos
                self._last_pan_pos = event.pos()
                hbar = self.horizontalScrollBar()
                vbar = self.verticalScrollBar()
                hbar.setValue(hbar.value() - delta.x())
                vbar.setValue(vbar.value() - delta.y())
            event.accept()
            return
        if self._rubber.isVisible():
            self._rubber.setGeometry(QtCore.QRect(self._rubber_origin, event.pos()).normalized())
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._panning and event.button() == self._pan_button:
            self._panning = False
            self._pan_button = None
            self._last_pan_pos = None
            if hasattr(self.controller, '_update_pdf_cursor'):
                self.controller._update_pdf_cursor()
            event.accept()
            return
        if self._rubber.isVisible():
            geo = self._rubber.geometry()
            self._rubber.hide()
            # map rubberband rect to scene coordinates
            p1 = self.mapToScene(geo.topLeft())
            p2 = self.mapToScene(geo.bottomRight())
            rect = QtCore.QRectF(p1, p2).toRect()
            # rect is already a QRect; emit directly
            self.rectPicked.emit(rect)
        else:
            super().mouseReleaseEvent(event)


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Axis')
        self.pdf_path = None
        self.doc = None
        self.current_page = 0
        self.mode = 'Ballooning'  # or 'Inspection'
        self.pick_on_print = False
        self.show_balloons = True
        self.current_wo = None
        # initialize rows to avoid attribute error before first load
        self.rows = []
        self.balloon_items = []
        self.highlight_rect = None
        self.selected_feature_id = None
        self.default_balloon_radius = BALLOON_RADIUS
        self._syncing_balloon_spin = False
        self._syncing_page_spin = False
        self.balloon_size_label = None
        self.balloon_size_spin = None
        self.page_label = None
        self.page_spin = None
        self.total_pages = 0
        self.method_options = ['CMM', 'Pin Gage', 'Visual']
        self._suppress_auto_focus = False
        self._undo_stack = []
        self._undo_in_progress = False
        self._zoom_rerender_timer = QtCore.QTimer(self)  # [ZOOM-DEBOUNCE]
        self._zoom_rerender_timer.setSingleShot(True)  # [ZOOM-DEBOUNCE]
        self._zoom_rerender_timer.timeout.connect(self._zoom_rerender_timeout)  # [ZOOM-DEBOUNCE]
        self._pending_zoom_scale = None  # [ZOOM-DEBOUNCE]
        self._balloons_built_for_page = None  # [ZOOM-DEBOUNCE]

        self._build_ui()
        self._refresh_method_combobox_options()
        self._update_pdf_cursor()

    def _build_ui(self):
        toolbar = QtWidgets.QToolBar()
        self.addToolBar(toolbar)
        open_act = QtGui.QAction('Open PDF', self)
        open_act.triggered.connect(self.open_pdf)
        toolbar.addAction(open_act)

        self.pick_btn = QtWidgets.QPushButton('Pick-on-Print')
        self.pick_btn.setCheckable(True)
        self.pick_btn.toggled.connect(self.toggle_pick)
        toolbar.addWidget(self.pick_btn)
        # keyboard shortcuts for ballooning mode helpers
        self.shortcut_pick = QtGui.QShortcut(QtGui.QKeySequence('P'), self)
        self.shortcut_pick.activated.connect(self._shortcut_pick_mode)
        self.shortcut_grab = QtGui.QShortcut(QtGui.QKeySequence('G'), self)
        self.shortcut_grab.activated.connect(self._shortcut_grab_mode)
        self.shortcut_undo = QtGui.QShortcut(QtGui.QKeySequence(QtGui.QKeySequence.StandardKey.Undo), self)
        self.shortcut_undo.activated.connect(self._shortcut_undo)

        fit_btn = QtWidgets.QPushButton('Fit')
        fit_btn.clicked.connect(self.fit_view)
        toolbar.addWidget(fit_btn)

        self.show_balloon_btn = QtWidgets.QPushButton('Hide Balloons')
        self.show_balloon_btn.setCheckable(True)
        self.show_balloon_btn.toggled.connect(self.toggle_show_balloons)
        toolbar.addWidget(self.show_balloon_btn)

        ocr_btn = QtWidgets.QPushButton('OCR')
        ocr_btn.clicked.connect(self._run_ocr)
        toolbar.addWidget(ocr_btn)

        spc_btn = QtWidgets.QPushButton('SPC')
        spc_btn.clicked.connect(self._open_spc_dashboard)
        toolbar.addWidget(spc_btn)

        self.balloon_size_label = QtWidgets.QLabel('Balloon size:')
        toolbar.addWidget(self.balloon_size_label)

        self.balloon_size_spin = QtWidgets.QSpinBox()
        self.balloon_size_spin.setRange(6, 60)
        self.balloon_size_spin.setValue(int(round(self.default_balloon_radius)))
        self.balloon_size_spin.valueChanged.connect(self._balloon_size_changed)
        toolbar.addWidget(self.balloon_size_spin)

        self.page_label = QtWidgets.QLabel('Page:')
        toolbar.addWidget(self.page_label)

        self.page_spin = QtWidgets.QSpinBox()
        self.page_spin.setRange(1, 1)
        self.page_spin.setValue(1)
        self.page_spin.setEnabled(False)
        self.page_spin.valueChanged.connect(self._page_spin_changed)
        toolbar.addWidget(self.page_spin)

        export_btn = QtWidgets.QPushButton('Export PDFs')
        export_btn.clicked.connect(self.export_pdfs)
        toolbar.addWidget(export_btn)

        export_results_btn = QtWidgets.QPushButton('Export Results')
        export_results_btn.clicked.connect(self.export_results)
        toolbar.addWidget(export_results_btn)

        export_all_results_btn = QtWidgets.QPushButton('Export All Results')
        export_all_results_btn.clicked.connect(self.export_all_results)
        toolbar.addWidget(export_all_results_btn)

        edit_methods_btn = QtWidgets.QPushButton('Edit Methods')
        edit_methods_btn.clicked.connect(self._edit_methods)
        toolbar.addWidget(edit_methods_btn)

        # Change session without reopening PDF
        change_session = QtGui.QAction('Change Session', self)
        change_session.triggered.connect(self.change_session)
        toolbar.addAction(change_session)
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        h = QtWidgets.QHBoxLayout(central)
        self.splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Horizontal)
        self.splitter.setChildrenCollapsible(False)
        self.splitter.setHandleWidth(8)

        # Left table + filters
        left = QtWidgets.QWidget()
        lv = QtWidgets.QVBoxLayout(left)
        filter_layout = QtWidgets.QHBoxLayout()
        filter_layout.addWidget(QtWidgets.QLabel('Method:'))
        self.method_filter = QtWidgets.QComboBox()
        self.method_filter.setEditable(False)
        self.method_filter.setInsertPolicy(QtWidgets.QComboBox.InsertPolicy.NoInsert)
        self.method_filter.setSizeAdjustPolicy(QtWidgets.QComboBox.SizeAdjustPolicy.AdjustToContents)
        self.method_filter.addItem('All')
        self.method_filter.currentTextChanged.connect(self.refresh_table)
        filter_layout.addWidget(self.method_filter)
        filter_layout.addWidget(QtWidgets.QLabel('Status:'))
        self.status_filter = QtWidgets.QComboBox()
        self.status_filter.addItems(['All', 'PASS', 'FAIL', '—'])
        self.status_filter.currentTextChanged.connect(self.refresh_table)
        filter_layout.addWidget(self.status_filter)
        lv.addLayout(filter_layout)

        self.table = QtWidgets.QTableWidget(0, 8)
        self.table.setHorizontalHeaderLabels(['ID', 'Page', 'Method', 'Result', 'Nominal', 'LSL', 'USL', 'Status'])
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
        self.table.cellChanged.connect(self.table_cell_changed)
        self.table.itemSelectionChanged.connect(self._table_selection_changed)
        self.table.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._table_context_menu)
        lv.addWidget(self.table)

        # PDF view on right
        self.pdf_view = PDFView(self)
        self.pdf_view.rectPicked.connect(self._rect_picked)
        self.splitter.addWidget(left)
        self.splitter.addWidget(self.pdf_view)
        self.splitter.setStretchFactor(0, 1)
        self.splitter.setStretchFactor(1, 3)
        h.addWidget(self.splitter)

        self._update_balloon_size_controls_enabled()
        self._update_page_controls_enabled()

    def refresh_table(self):
        table = getattr(self, 'table', None)
        if table is None:
            return
        try:
            table.blockSignals(True)
        except RuntimeError:
            return
        table.setRowCount(0)
        wo_results = read_wo(self.pdf_path, self.current_wo) if self.current_wo else {}
        for r in self.rows:
            id_ = r.get('id', '')
            page = str(r.get('page', ''))
            method = r.get('method', '')
            nom = r.get('nominal', '')
            lsl = r.get('lsl', '')
            usl = r.get('usl', '')
            status = '—'
            if self.mode == 'Ballooning':
                # no persistent result; editable cell used for auto-fill only
                result = ''
            else:
                result = wo_results.get(id_, '')
            # determine status
            try:
                if result.upper() == 'PASS':
                    status = 'PASS'
                elif result.upper() == 'FAIL':
                    status = 'FAIL'
                else:
                    if result != '' and nom != '' and lsl != '' and usl != '':
                        val = float(result)
                        if float(lsl) <= val <= float(usl):
                            status = 'PASS'
                        else:
                            status = 'FAIL'
            except Exception:
                status = '—'

            # status filter
            sel_status = self.status_filter.currentText() if self.status_filter else 'All'
            if sel_status != 'All' and status != sel_status:
                continue

            # method filter
            sel_method = self.method_filter.currentText() if isinstance(self.method_filter, QtWidgets.QComboBox) else 'All'
            if sel_method not in ('All', ''):
                method_value = (method or '').strip()
                if method_value.lower() != sel_method.lower():
                    continue

            row = table.rowCount()
            table.insertRow(row)
            # helper to set read-only/ editable
            def make_item(text: str, editable: bool):
                it = QtWidgets.QTableWidgetItem(text)
                if not editable:
                    it.setFlags(it.flags() & ~QtCore.Qt.ItemFlag.ItemIsEditable)
                return it

            self.table.setItem(row, 0, make_item(id_, False))
            self.table.setItem(row, 1, make_item(page, False))
            can_edit_specs = (self.mode == 'Ballooning')
            method_item = make_item(method, False)
            self.table.setItem(row, 2, method_item)
            combo = QtWidgets.QComboBox()
            combo.setEditable(False)
            combo.setInsertPolicy(QtWidgets.QComboBox.InsertPolicy.NoInsert)
            combo.setSizeAdjustPolicy(QtWidgets.QComboBox.SizeAdjustPolicy.AdjustToContents)
            combo.setStyleSheet(
                "QComboBox { color: rgb(235, 235, 235); } "
                "QComboBox::drop-down { border: none; } "
                "QComboBox:disabled { color: rgb(225, 225, 225); background-color: rgb(45, 45, 45); }"
            )
            combo.addItems([''] + self.method_options)
            combo.setCurrentText(method)
            combo.setEnabled(self.mode == 'Ballooning')
            combo.currentTextChanged.connect(partial(self._method_combo_changed, row))
            self.table.setCellWidget(row, 2, combo)
            itm_result = QtWidgets.QTableWidgetItem(result)
            # Result editable in both modes (Ballooning: used for tol autofill; Inspection: writes to WO)
            itm_result.setFlags(itm_result.flags() | QtCore.Qt.ItemFlag.ItemIsEditable)
            table.setItem(row, 3, itm_result)
            table.setItem(row, 4, make_item(nom, can_edit_specs))
            table.setItem(row, 5, make_item(lsl, can_edit_specs))
            table.setItem(row, 6, make_item(usl, can_edit_specs))
            table.setItem(row, 7, make_item(status, False))
            self._apply_status_formatting(row, status)

        self._refresh_method_combobox_options()

        try:
            table.blockSignals(False)
        except RuntimeError:
            return
        if self.selected_feature_id:
            target_row = None
            for idx in range(table.rowCount()):
                item = table.item(idx, 0)
                if item and item.text() == self.selected_feature_id:
                    target_row = idx
                    break
            if target_row is not None:
                table.selectRow(target_row)
                return
            # selected feature no longer visible
            self.selected_feature_id = None
        # no selection available
        self._apply_balloon_selection_visuals()
        self._clear_highlight_rect()

    def table_cell_changed(self, row, col):
        # handle edits per column without full-table refresh to avoid editor warnings
        id_item = self.table.item(row, 0)
        if not id_item:
            return
        fid = id_item.text()
        if col == 3:  # Result column
            val_item = self.table.item(row, col)
            val = val_item.text() if val_item else ''
            normalized = self._normalize_result_entry(val)
            if normalized != val and val_item:
                try:
                    self.table.blockSignals(True)
                    val_item.setText(normalized)
                finally:
                    self.table.blockSignals(False)
                val = normalized
            if self.mode == 'Ballooning':
                # auto-fill tolerances if empty
                for r in self.rows:
                    if r.get('id') == fid:
                        nom = r.get('nominal', '')
                        lsl = r.get('lsl', '')
                        usl = r.get('usl', '')
                        if not nom and not lsl and not usl and val.strip():
                            try:
                                nomv, lslv, uslv = parse_tolerance_expression(val)
                                formatted_nom = format_number(nomv)
                                formatted_lsl = format_number(lslv)
                                formatted_usl = format_number(uslv)
                                update_feature(self.pdf_path, fid, {
                                    "nominal": formatted_nom,
                                    "lsl": formatted_lsl,
                                    "usl": formatted_usl,
                                })
                                r['nominal'] = formatted_nom
                                r['lsl'] = formatted_lsl
                                r['usl'] = formatted_usl
                                self.table.blockSignals(True)
                                self.table.item(row, 3).setText('')
                                self.table.item(row, 4).setText(formatted_nom)
                                self.table.item(row, 5).setText(formatted_lsl)
                                self.table.item(row, 6).setText(formatted_usl)
                                self.table.blockSignals(False)
                            except Exception:
                                pass
                        break
            else:
                # Inspection: write result and recompute status for this row
                wo_results = read_wo(self.pdf_path, self.current_wo)
                wo_results[fid] = val
                write_wo(self.pdf_path, self.current_wo, wo_results)
                self._recompute_row_status(row)
            self._advance_result_edit(row)
            return
        if self.mode == 'Ballooning' and col == 4:
            text_item = self.table.item(row, col)
            text_value = text_item.text() if text_item else ''
            if self._try_apply_tolerance_entry(row, fid, text_value):
                return
        # Ballooning: editing Method/Nom/LSL/USL persists to master
        if self.mode == 'Ballooning' and col in (2, 4, 5, 6):
            text = self.table.item(row, col).text()
            key = {2: 'method', 4: 'nominal', 5: 'lsl', 6: 'usl'}[col]
            update_feature(self.pdf_path, fid, {key: text})
            for r in self.rows:
                if r.get('id') == fid:
                    r[key] = text
                    break
            if col == 2:
                self._refresh_method_combobox_options()
            self._recompute_row_status(row)
            return
        # no further handling for other columns
        return

    def _advance_result_edit(self, current_row: int):
        """After committing a Result value, focus the next row's Result cell for quick data entry."""
        table = getattr(self, 'table', None)
        if table is None:
            return
        next_row = current_row + 1
        while next_row < table.rowCount():
            next_item = table.item(next_row, 3)
            if next_item is not None:
                table.setCurrentCell(next_row, 3)

                def _start_edit(target_row=next_row):
                    table_ref = getattr(self, 'table', None)
                    if table_ref is None:
                        return
                    if target_row < 0 or target_row >= table_ref.rowCount():
                        return
                    item_ref = table_ref.item(target_row, 3)
                    if item_ref is None:
                        return
                    try:
                        table_ref.editItem(item_ref)
                    except Exception:
                        pass

                QtCore.QTimer.singleShot(0, _start_edit)
                break
            next_row += 1


    def _recompute_row_status(self, row: int):
        if row < 0 or row >= self.table.rowCount():
            return
        result_item = self.table.item(row, 3)
        nom_item = self.table.item(row, 4)
        lsl_item = self.table.item(row, 5)
        usl_item = self.table.item(row, 6)
        status_item = self.table.item(row, 7)

        result_text = result_item.text().strip() if result_item else ''
        nom_text = nom_item.text().strip() if nom_item else ''
        lsl_text = lsl_item.text().strip() if lsl_item else ''
        usl_text = usl_item.text().strip() if usl_item else ''

        status = self._status_from_fields(result_text, lsl_text, usl_text)

        if status_item:
            self.table.blockSignals(True)
            status_item.setText(status)
            self.table.blockSignals(False)
        self._apply_status_formatting(row, status)

    def _status_from_fields(self, result_text: str, lsl_text: str, usl_text: str) -> str:
        if not result_text:
            return '—'
        upper = result_text.upper()
        if upper in ('PASS', 'FAIL'):
            return upper
        try:
            value = float(result_text)
            lsl_val = float(lsl_text) if lsl_text else None
            usl_val = float(usl_text) if usl_text else None
            if lsl_val is None or usl_val is None:
                return '—'
            return 'PASS' if lsl_val <= value <= usl_val else 'FAIL'
        except Exception:
            return '—'

    def _apply_status_formatting(self, row: int, status: str):
        table = getattr(self, 'table', None)
        if table is None or row < 0 or row >= table.rowCount():
            return
        palette = {
            'PASS': QtGui.QColor(200, 235, 200),
            'FAIL': QtGui.QColor(247, 205, 205),
            '—': QtGui.QColor(235, 235, 235),
        }
        text_colors = {
            'PASS': QtGui.QColor(10, 80, 10),
            'FAIL': QtGui.QColor(140, 20, 20),
            '—': QtGui.QColor(200, 200, 200),
        }
        bold_status = status in ('PASS', 'FAIL')
        color = palette.get(status.upper() if isinstance(status, str) else status)
        text_color = text_colors.get(status.upper() if isinstance(status, str) else status, QtGui.QColor(220, 220, 220))
        for col in (3, 7):
            item = table.item(row, col)
            if not item:
                continue
            if color:
                item.setBackground(QtGui.QBrush(color))
            else:
                item.setBackground(QtGui.QBrush())
            item.setForeground(QtGui.QBrush(text_color))
            font = item.font()
            font.setBold(bold_status)
            item.setFont(font)

    def _normalize_result_entry(self, value: str) -> str:
        if value is None:
            return ''
        stripped = value.strip()
        if not stripped:
            return ''
        upper = stripped.upper()
        if upper in ('P', 'PASS'):
            return 'Pass'
        if upper in ('F', 'FAIL'):
            return 'FAIL'
        return stripped

    def _string_has_tolerance(self, text: str) -> bool:
        if not text:
            return False
        normalized = normalize_str(text)
        if not normalized:
            return False
        lowered = normalized.lower()
        if any(token in lowered for token in ('±', '+/-', '-/+', '+-', '-+')):
            return True
        plus_pos = normalized.find('+', 1)
        if plus_pos != -1:
            return True
        minus_count = normalized.count('-')
        if normalized.startswith('-'):
            minus_count -= 1
        return minus_count > 1

    def _try_apply_tolerance_entry(self, row: int, fid: str, text: str) -> bool:
        if not text:
            return False
        candidate = normalize_str(text)
        if not self._string_has_tolerance(candidate):
            return False
        try:
            nomv, lslv, uslv = parse_tolerance_expression(candidate)
        except Exception:
            return False
        formatted_nom = format_number(nomv)
        formatted_lsl = format_number(lslv)
        formatted_usl = format_number(uslv)
        updates = {
            'nominal': formatted_nom,
            'lsl': formatted_lsl,
            'usl': formatted_usl,
        }
        if self.pdf_path:
            update_feature(self.pdf_path, fid, updates)
        for r in self.rows:
            if r.get('id') == fid:
                r.update(updates)
                break
        table = getattr(self, 'table', None)
        if table is None:
            return False
        try:
            table.blockSignals(True)
            for col, value in ((4, formatted_nom), (5, formatted_lsl), (6, formatted_usl)):
                item = table.item(row, col)
                if item is None:
                    item = QtWidgets.QTableWidgetItem(value)
                    item.setFlags(item.flags() | QtCore.Qt.ItemFlag.ItemIsEditable)
                    table.setItem(row, col, item)
                else:
                    item.setText(value)
        finally:
            table.blockSignals(False)
        self._recompute_row_status(row)
        return True

    def _table_selection_changed(self):
        if not self.pdf_path:
            return
        selection_model = self.table.selectionModel()
        if selection_model is None:
            return
        selected_rows = selection_model.selectedRows()
        if selected_rows:
            row = selected_rows[0].row()
        else:
            row = self.table.currentRow()
            if row < 0:
                self.selected_feature_id = None
                self._apply_balloon_selection_visuals()
                self._clear_highlight_rect()
                return
        id_item = self.table.item(row, 0)
        if not id_item:
            return
        fid = id_item.text()
        feature = None
        for r in self.rows:
            if r.get('id') == fid:
                feature = r
                break
        if not feature:
            return
        if self._suppress_auto_focus:
            self._suppress_auto_focus = False
            self.selected_feature_id = fid
            self._apply_balloon_selection_visuals()
            try:
                x = float(feature.get('x', 0))
                y = float(feature.get('y', 0))
                w = float(feature.get('w', 0))
                h = float(feature.get('h', 0))
            except Exception:
                self._clear_highlight_rect()
                return
            rect_item = self._ensure_highlight_rect()
            if rect_item is not None:
                rect_item.setRect(QtCore.QRectF(x, y, w, h))
                rect_item.setVisible(True)
            else:
                self._clear_highlight_rect()
            return
        self.selected_feature_id = fid
        try:
            page_idx = int(feature.get('page', '1')) - 1
        except Exception:
            page_idx = 0
        page_idx = max(0, page_idx)
        if page_idx != self.current_page:
            self._balloons_built_for_page = None  # [ZOOM-DEBOUNCE]
            self.current_page = page_idx
            self._render_current_page()
        self._focus_on_feature(feature)
        self._apply_balloon_selection_visuals()
        br_value = feature.get('br', str(self.default_balloon_radius))
        try:
            radius_value = float(br_value)
        except Exception:
            radius_value = self.default_balloon_radius
        self._sync_balloon_size_spin(radius_value)

    def _table_context_menu(self, pos: QtCore.QPoint):
        if self.mode != 'Ballooning' or not self.pdf_path:
            return
        table = getattr(self, 'table', None)
        if table is None:
            return
        row = table.rowAt(pos.y())
        if row < 0:
            return
        table.selectRow(row)
        menu = QtWidgets.QMenu(self)
        delete_action = menu.addAction('Delete Balloon')
        global_pos = table.viewport().mapToGlobal(pos)
        chosen = menu.exec(global_pos)
        if chosen == delete_action:
            self._delete_balloon_row(row)

    def _delete_balloon_row(self, row: int):
        if self.mode != 'Ballooning' or not self.pdf_path:
            return
        table = getattr(self, 'table', None)
        if table is None or row < 0 or row >= table.rowCount():
            return
        id_item = table.item(row, 0)
        if not id_item:
            return
        fid = id_item.text().strip()
        if not fid:
            return
        confirm = QtWidgets.QMessageBox.question(
            self,
            'Delete Balloon',
            f'Remove balloon {fid}?',
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
            QtWidgets.QMessageBox.StandardButton.No,
        )
        if confirm != QtWidgets.QMessageBox.StandardButton.Yes:
            return

        snapshot = self._remove_feature(fid, row_hint=row)
        if snapshot:
            self._push_undo(lambda snap=snapshot: self._undo_deleted_feature(snap), f'Delete {fid}')
            self.statusBar().showMessage(f'Deleted {fid}', 3000)

    def _focus_on_feature(self, feature: dict):
        if not self.pdf_view._pixmap_item:
            return
        try:
            zoom = float(feature.get('zoom') or 1.0)
        except Exception:
            zoom = 1.0
        zoom = max(self.pdf_view._min_scale, min(self.pdf_view._max_scale, zoom))
        try:
            x = float(feature.get('x', 0))
            y = float(feature.get('y', 0))
            w = float(feature.get('w', 0))
            h = float(feature.get('h', 0))
            bx = float(feature.get('bx', 0))
            by = float(feature.get('by', 0))
        except Exception:
            return
        cx = x + (w / 2.0) + bx
        cy = y + (h / 2.0) + by
        center_point = QtCore.QPointF(cx, cy)
        self.pdf_view.set_zoom(zoom, center_point)
        rect_item = self._ensure_highlight_rect()
        if rect_item is not None:
            rect_item.setRect(QtCore.QRectF(x, y, w, h))
            rect_item.setVisible(True)
        else:
            self._clear_highlight_rect()

    def _method_combo_changed(self, row: int, text: str):
        table = getattr(self, 'table', None)
        if table is None:
            return
        if row < 0 or row >= table.rowCount():
            return
        combo = table.cellWidget(row, 2)
        if isinstance(combo, QtWidgets.QComboBox) and not combo.isEnabled():
            # keep combo text aligned with stored value when editing is disabled
            item = table.item(row, 2)
            if item and combo.currentText() != item.text():
                combo.blockSignals(True)
                combo.setCurrentText(item.text())
                combo.blockSignals(False)
            return
        item = table.item(row, 2)
        if not item:
            return
        if item.text() == text:
            return
        item.setText(text)
        id_item = table.item(row, 0)
        fid = id_item.text() if id_item else None
        if not fid:
            return
        if self.mode == 'Ballooning' and self.pdf_path:
            update_feature(self.pdf_path, fid, {'method': text})
            for r in self.rows:
                if r.get('id') == fid:
                    r['method'] = text
                    break
        self._refresh_method_combobox_options()
        self._refresh_method_filter_options()

    def _refresh_method_filter_options(self):
        combo = getattr(self, 'method_filter', None)
        if not isinstance(combo, QtWidgets.QComboBox):
            return
        methods = sorted(
            {
                (row.get('method') or '').strip()
                for row in self.rows
                if (row.get('method') or '').strip()
            },
            key=lambda s: s.lower()
        )
        current = combo.currentText() if combo.count() else 'All'
        combo.blockSignals(True)
        combo.clear()
        combo.addItem('All')
        for value in methods:
            combo.addItem(value)
        if current and combo.findText(current, QtCore.Qt.MatchFlag.MatchExactly) >= 0:
            combo.setCurrentText(current)
        else:
            combo.setCurrentIndex(0)
        combo.blockSignals(False)

    def _refresh_method_combobox_options(self):
        table = getattr(self, 'table', None)
        if table is None:
            return
        base = list(self.method_options)
        seen = {value.lower() for value in base}
        extras = []
        for row in range(table.rowCount()):
            item = table.item(row, 2)
            if not item:
                continue
            value = item.text().strip()
            if not value:
                continue
            key = value.lower()
            if key not in seen:
                extras.append(value)
                seen.add(key)
        ordered_options = base + sorted(extras, key=lambda s: s.lower())
        for row in range(table.rowCount()):
            combo = table.cellWidget(row, 2)
            if not isinstance(combo, QtWidgets.QComboBox):
                continue
            current = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            combo.addItem('')
            for value in ordered_options:
                combo.addItem(value)
            combo.setCurrentText(current)
            combo.setEnabled(self.mode == 'Ballooning')
            combo.blockSignals(False)
        self._refresh_method_filter_options()
    
    def _edit_methods(self):
        dlg = MethodListDialog(self, self.method_options)
        if dlg.exec() != QtWidgets.QDialog.DialogCode.Accepted:
            return
        updated = dlg.get_methods()
        if not updated:
            QtWidgets.QMessageBox.warning(self, 'Inspection Methods', 'At least one method is required.')
            return
        self.method_options = sorted(updated, key=lambda s: s.lower())
        self._refresh_method_combobox_options()

    def _shortcut_pick_mode(self):
        if self.mode != 'Ballooning' or not self.pdf_path:
            return
        if not self.pick_on_print:
            self.pick_btn.setChecked(True)
        else:
            # ensure button text reflects state when already active
            self.pick_btn.setText('Picking...')
        self.statusBar().showMessage('Pick mode enabled (P)', 2000)

    def _shortcut_grab_mode(self):
        if self.mode != 'Ballooning' or not self.pdf_path:
            return
        if self.pick_on_print:
            self.pick_btn.setChecked(False)
        else:
            self.pick_btn.setText('Pick-on-Print')
        self.statusBar().showMessage('Grab mode enabled (G)', 2000)
        self._update_pdf_cursor()

    def _shortcut_undo(self):
        self._undo_last_action()

    def toggle_pick(self, checked: bool):
        if not self.pdf_path or self.mode != 'Ballooning':
            self.pick_btn.blockSignals(True)
            self.pick_btn.setChecked(False)
            self.pick_btn.blockSignals(False)
            self.pick_on_print = False
            self._update_pdf_cursor()
            if not self.pdf_path:
                QtWidgets.QMessageBox.information(self, 'No PDF', 'Open a PDF before placing balloons.')
            else:
                QtWidgets.QMessageBox.information(self, 'Inspection Mode', 'Pick-on-Print is only available in Ballooning mode.')
            return
        self.pick_on_print = checked
        self.pick_btn.setText('Picking...' if checked else 'Pick-on-Print')
        self._update_pdf_cursor()

    def fit_view(self):
        if self.pdf_view._pixmap_item:
            self.pdf_view.fit_to_view()

    def toggle_show_balloons(self, checked: bool):
        self.show_balloons = not checked
        self.show_balloon_btn.setText('Show Balloons' if not self.show_balloons else 'Hide Balloons')
        self._update_balloon_visibility()

    def _update_pdf_cursor(self):
        view = getattr(self, 'pdf_view', None)
        if view is None:
            return
        if self.pick_on_print and self.mode == 'Ballooning' and self.pdf_path:
            view.setCursor(QtCore.Qt.CursorShape.CrossCursor)
        else:
            view.unsetCursor()

    def _run_ocr(self):
        if not self.doc:
            QtWidgets.QMessageBox.information(self, 'OCR', 'Open a PDF before running OCR.')
            return
        try:
            import pytesseract  # type: ignore
        except ImportError:
            QtWidgets.QMessageBox.warning(
                self,
                'OCR',
                'pytesseract is not installed. Run "pip install pytesseract" in your environment.'
            )
            return
        if Image is None:
            QtWidgets.QMessageBox.warning(
                self,
                'OCR',
                'Pillow is required for OCR image conversion. Install it with "pip install Pillow".'
            )
            return
        try:
            page = self.doc.load_page(self.current_page)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, 'OCR', f'Unable to load the current page for OCR.\n{exc}')
            return
        buffer = None
        image = None
        try:
            matrix = fitz.Matrix(PAGE_RENDER_ZOOM * 2.0, PAGE_RENDER_ZOOM * 2.0)
            pix = page.get_pixmap(matrix=matrix)
            buffer = io.BytesIO(pix.tobytes('png'))
            image = Image.open(buffer)
        except Exception as exc:
            if buffer is not None:
                buffer.close()
            QtWidgets.QMessageBox.warning(self, 'OCR', f'OCR capture failed.\n{exc}')
            return
        try:
            text = pytesseract.image_to_string(image)
        except pytesseract.TesseractNotFoundError:
            QtWidgets.QMessageBox.warning(
                self,
                'OCR',
                'Tesseract executable is not available. Install it and ensure it is on your PATH.'
            )
            return
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, 'OCR', f'OCR failed.\n{exc}')
            return
        finally:
            if image is not None:
                try:
                    image.close()
                except Exception:
                    pass
            if buffer is not None:
                buffer.close()
        text = (text or '').strip()
        if not text:
            text = '(No text detected)'
        self._show_ocr_dialog(text)

    def _show_ocr_dialog(self, text: str):
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle('OCR Result')
        layout = QtWidgets.QVBoxLayout(dialog)
        output = QtWidgets.QPlainTextEdit()
        output.setReadOnly(True)
        output.setPlainText(text)
        layout.addWidget(output)
        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(dialog.reject)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.resize(640, 480)
        dialog.exec()

    def _open_spc_dashboard(self):
        if not self.pdf_path:
            QtWidgets.QMessageBox.information(self, 'SPC', 'Open a PDF before viewing SPC data.')
            return
        try:
            dataset = spc.load_spc_dataset(self.pdf_path)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, 'SPC', f'Unable to load SPC data.\n{exc}')
            return
        if not dataset:
            QtWidgets.QMessageBox.information(self, 'SPC', 'No numeric inspection results found for this part.')
            return
        dlg = SPCDialog(self, self.pdf_path, dataset)
        dlg.exec()

    def _update_balloon_visibility(self):
        for item in self.balloon_items:
            item.setVisible(self.show_balloons)
        self._apply_balloon_selection_visuals()

    def change_session(self):
        if not self.pdf_path:
            QtWidgets.QMessageBox.information(self, 'No PDF', 'Open a PDF first.')
            return
        if self._prompt_session(list_workorders(self.pdf_path)):
            self.refresh_table()

    def _prompt_session(self, previous_orders=None, allow_ballooning=True) -> bool:
        dlg = StartSessionDialog(self, previous_orders, allow_ballooning=allow_ballooning)
        if dlg.exec() != QtWidgets.QDialog.DialogCode.Accepted:
            return False
        if dlg.choice == 'ballooning':
            self.mode = 'Ballooning'
            self.current_wo = None
        elif dlg.choice == 'inspection':
            serial = normalize_str(dlg.serial)
            if not serial:
                QtWidgets.QMessageBox.warning(self, 'Serial Required', 'Please provide a valid serial number.')
                return False
            self.mode = 'Inspection'
            self.current_wo = serial
        else:
            return False
        self._update_mode_ui()
        return True

    def _update_mode_ui(self):
        has_pdf = self.pdf_path is not None
        is_ballooning = self.mode == 'Ballooning'
        self.pick_btn.setEnabled(has_pdf and is_ballooning)
        if not (has_pdf and is_ballooning):
            self.pick_btn.blockSignals(True)
            self.pick_btn.setChecked(False)
            self.pick_btn.blockSignals(False)
            self.pick_on_print = False
            self.pick_btn.setText('Pick-on-Print')
        self._update_balloon_size_controls_enabled()
        self._update_pdf_cursor()

        title = 'Axis'
        if self.pdf_path:
            title += f" - {Path(self.pdf_path).name}"
        if self.mode == 'Inspection' and self.current_wo:
            title += f" [{self.current_wo}]"
        self.setWindowTitle(title)

        message = f"Mode: {self.mode}"
        if self.mode == 'Inspection' and self.current_wo:
            message += f" | WO {self.current_wo}"
        self.statusBar().showMessage(message, 4000)

    def open_pdf(self):
        start_dir = Path(self.pdf_path).parent if self.pdf_path else Path.home()
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            'Open PDF',
            str(start_dir),
            'PDF Files (*.pdf)'
        )
        if not file_path:
            return
        self._load_pdf(file_path)

    def _load_pdf(self, path: str):
        try:
            doc = fitz.open(path)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, 'Open PDF', f'Could not open PDF file.\n{exc}')
            # new PDF: default to ballooning without prompting
            self.mode = 'Ballooning'
            self.current_wo = None
            self._update_mode_ui()
            return

        # clear any previous document before loading the new one
        self._clear_loaded_pdf()

        self.doc = doc
        self.pdf_path = path
        self.total_pages = doc.page_count
        self.current_page = 0

        try:
            ensure_master(path)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, 'Master Data', f'Could not prepare master data for this PDF.\n{exc}')

        self._load_rows()
        if not self._prompt_session(list_workorders(self.pdf_path)):
            self.mode = 'Ballooning'
            self.current_wo = None
        self._update_mode_ui()
        self._update_page_controls_enabled()
        self._sync_page_spin()

        self._render_current_page()
        self.fit_view()
        self.refresh_table()
        self.statusBar().showMessage(f'Loaded {Path(path).name}', 4000)

    def _clear_loaded_pdf(self):
        self._undo_stack.clear()
        self._undo_in_progress = False
        self._zoom_rerender_timer.stop()  # [ZOOM-DEBOUNCE]
        self._pending_zoom_scale = None  # [ZOOM-DEBOUNCE]
        if self.doc:
            try:
                self.doc.close()
            except Exception:
                pass
        self.doc = None
        self.pdf_path = None
        self.current_page = 0
        self.total_pages = 0
        self.rows = []
        self._clear_balloon_items()
        self._suppress_auto_focus = False
        self.selected_feature_id = None
        self._clear_highlight_rect()
        if getattr(self, 'pdf_view', None):
            self.pdf_view._last_render_scale = None  # [CRISP-ZOOM]
        if self.pdf_view.scene():
            self.pdf_view.scene().clear()
            self.pdf_view._pixmap_item = None
        self.table.setRowCount(0)
        self._update_mode_ui()
        self._update_page_controls_enabled()
        self._sync_page_spin()
        self._refresh_method_filter_options()
        self._update_pdf_cursor()
        self._balloons_built_for_page = None  # [ZOOM-DEBOUNCE]

    def _load_rows(self):
        if not self.pdf_path:
            self.rows = []
            self._refresh_method_filter_options()
            return
        rows = read_master(self.pdf_path)
        prepared = []
        for r in rows:
            r = dict(r)
            r['_pdf'] = self.pdf_path
            r.setdefault('page', '1')
            r.setdefault('x', '0')
            r.setdefault('y', '0')
            r.setdefault('w', '0')
            r.setdefault('h', '0')
            r.setdefault('bx', '0')
            r.setdefault('by', '0')
            r.setdefault('br', str(self.default_balloon_radius))
            r.setdefault('zoom', '1.0')
            prepared.append(r)
        self.rows = prepared
        if prepared:
            try:
                first_radius = float(prepared[0].get('br', self.default_balloon_radius))
            except Exception:
                first_radius = self.default_balloon_radius
            self.default_balloon_radius = max(6, min(60, int(round(first_radius))))
        # merge any methods from the master into the available options
        existing = {m.strip() for m in self.method_options if m.strip()}
        for feature in prepared:
            method_value = (feature.get('method') or '').strip()
            if method_value and method_value not in existing:
                existing.add(method_value)
        self.method_options = sorted(existing, key=lambda s: s.lower())
        self._sync_balloon_size_spin(self.default_balloon_radius)
        self._refresh_method_combobox_options()

    def _maybe_rerender_for_zoom(self, view_scale: float):
        pdf_view = getattr(self, 'pdf_view', None)
        if not self.doc or not pdf_view or not pdf_view._pixmap_item:
            return
        if view_scale <= 0:
            return
        if pdf_view._last_render_scale is None:
            self._render_current_page(render_scale=view_scale)  # [CRISP-ZOOM]
            return
        lo = min(pdf_view._last_render_scale, view_scale)
        hi = max(pdf_view._last_render_scale, view_scale)
        if hi / max(lo, 1e-6) >= 1.6:  # [LOD-THRESHOLD]
            self._render_current_page(render_scale=view_scale)  # [CRISP-ZOOM]

    def _schedule_rerender_for_zoom(self, view_scale: float):
        """Debounce heavy PDF re-rendering after zoom gestures."""  # [ZOOM-DEBOUNCE]
        if view_scale <= 0:
            return
        pdf_view = getattr(self, 'pdf_view', None)
        if not self.doc or not pdf_view or not pdf_view._pixmap_item:
            return
        self._pending_zoom_scale = view_scale
        self._zoom_rerender_timer.start(100)

    def _zoom_rerender_timeout(self):
        if self._pending_zoom_scale is None:
            return
        scale = self._pending_zoom_scale
        self._pending_zoom_scale = None
        if not self.doc or not getattr(self, 'pdf_view', None):
            return
        self._maybe_rerender_for_zoom(scale)

    def _render_current_page(self, render_scale: float | None = None):
        if not self.doc:
            return
        page_count = self.doc.page_count
        if page_count == 0:
            return
        self.current_page = max(0, min(self.current_page, page_count - 1))
        try:
            page = self.doc.load_page(self.current_page)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, 'Render Failed', f'Unable to load page: {exc}')
            return

        if render_scale is None:
            self._zoom_rerender_timer.stop()  # [ZOOM-DEBOUNCE]
            self._pending_zoom_scale = None  # [ZOOM-DEBOUNCE]

        pdf_view = getattr(self, 'pdf_view', None)
        preserve_center = None
        if pdf_view and pdf_view._pixmap_item and render_scale is not None:
            preserve_center = pdf_view.mapToScene(pdf_view.viewport().rect().center())

        try:
            dpr = float(self.devicePixelRatioF()) if hasattr(self, 'devicePixelRatioF') else 1.0
        except Exception:
            dpr = 1.0
        base_scale = float(PAGE_RENDER_ZOOM)
        view_scale = float(render_scale or 1.0)
        max_render_factor = 3.0  # [LOD-THRESHOLD]
        effective_view_scale = min(view_scale, max_render_factor)  # [LOD-THRESHOLD]
        matrix_scale = base_scale * effective_view_scale * dpr
        pix = page.get_pixmap(matrix=fitz.Matrix(matrix_scale, matrix_scale))
        qimage = QtGui.QImage.fromData(pix.tobytes('png'), 'PNG')
        try:
            qimage.setDevicePixelRatio(dpr)  # [CRISP-ZOOM]
        except Exception:
            pass

        page_rect = page.rect
        scene_width = float(page_rect.width) * float(PAGE_RENDER_ZOOM)
        scene_height = float(page_rect.height) * float(PAGE_RENDER_ZOOM)

        if pdf_view:
            pdf_view.load_page(qimage, (scene_width, scene_height))
            if preserve_center is not None:
                pdf_view.centerOn(preserve_center)
            pdf_view._current_scale = pdf_view.transform().m11()
            pdf_view._last_render_scale = pdf_view._current_scale  # [CRISP-ZOOM]

        if self._balloons_built_for_page != self.current_page:
            self._rebuild_balloons()  # [ZOOM-DEBOUNCE]
            self._balloons_built_for_page = self.current_page
        self._apply_balloon_selection_visuals()

        if self.selected_feature_id:
            feature = next((r for r in self.rows if r.get('id') == self.selected_feature_id), None)
            if feature:
                try:
                    x = float(feature.get('x', 0))
                    y = float(feature.get('y', 0))
                    w = float(feature.get('w', 0))
                    h = float(feature.get('h', 0))
                except Exception:
                    feature = None
                if feature:
                    rect_item = self._ensure_highlight_rect()
                    if rect_item is not None:
                        rect_item.setRect(QtCore.QRectF(x, y, w, h))
                        rect_item.setVisible(True)
            else:
                self._clear_highlight_rect()
        else:
            self._clear_highlight_rect()

        if render_scale is None:
            self._sync_page_spin()

    def _rebuild_balloons(self):
        self._clear_balloon_items()
        if not self.pdf_view._pixmap_item:
            return
        for feature in self.rows:
            page_str = feature.get('page', '1')
            try:
                page_idx = int(page_str) - 1
            except ValueError:
                page_idx = 0
            if page_idx != self.current_page:
                continue
            feature['_pdf'] = self.pdf_path
            item = BalloonItem(feature, self.pdf_view._pixmap_item)
            self.pdf_view.scene().addItem(item)
            self.balloon_items.append(item)
        self._update_balloon_visibility()
        self._balloons_built_for_page = self.current_page  # [ZOOM-DEBOUNCE]

    def _clear_balloon_items(self):
        for item in self.balloon_items:
            try:
                scene = item.scene()
                if scene is not None:
                    scene.removeItem(item)
            except Exception:
                pass
        self.balloon_items = []
        self._apply_balloon_selection_visuals()

    def _apply_balloon_selection_visuals(self):
        selected_id = self.selected_feature_id if self.show_balloons else None
        for item in self.balloon_items:
            if not self.show_balloons:
                item.setOpacity(1.0)
                item.setZValue(10)
                continue
            if selected_id and item.feature.get('id') == selected_id:
                item.setOpacity(1.0)
                item.setZValue(20)
            elif selected_id:
                item.setOpacity(0.25)
                item.setZValue(5)
            else:
                item.setOpacity(1.0)
                item.setZValue(10)

    def _ensure_highlight_rect(self):
        rect_item = self.highlight_rect
        if rect_item is not None:
            try:
                _ = rect_item.rect()
            except RuntimeError:
                rect_item = None
                self.highlight_rect = None
        if rect_item is None:
            scene = self.pdf_view.scene()
            if scene is None:
                return None
            rect_item = QtWidgets.QGraphicsRectItem()
            pen = QtGui.QPen(QtGui.QColor(255, 0, 0))
            pen.setWidth(2)
            rect_item.setPen(pen)
            rect_item.setBrush(QtGui.QBrush(QtGui.QColor(255, 0, 0, 40)))
            rect_item.setZValue(18)
            rect_item.setVisible(False)
            rect_item.setFlag(QtWidgets.QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, False)
            rect_item.setFlag(QtWidgets.QGraphicsItem.GraphicsItemFlag.ItemIsMovable, False)
            scene.addItem(rect_item)
            self.highlight_rect = rect_item
        return rect_item

    def _clear_highlight_rect(self):
        if not self.highlight_rect:
            return
        try:
            self.highlight_rect.hide()
        except RuntimeError:
            self.highlight_rect = None

    def _add_balloon_item(self, feature: dict):
        if not self.pdf_view._pixmap_item:
            return
        feature['_pdf'] = self.pdf_path
        item = BalloonItem(feature, self.pdf_view._pixmap_item)
        self.pdf_view.scene().addItem(item)
        self.balloon_items.append(item)
        if not self.show_balloons:
            item.hide()

    def _rect_picked(self, rect: QtCore.QRect):
        if self.mode != 'Ballooning' or not self.pdf_path:
            return
        if rect.width() < 5 or rect.height() < 5:
            return

        w_float = float(rect.width())
        h_float = float(rect.height())
        radius_value = self.default_balloon_radius
        offset_x = radius_value if w_float >= radius_value * 2 else w_float / 2.0
        offset_y = radius_value if h_float >= radius_value * 2 else h_float / 2.0
        default_bx = offset_x - (w_float / 2.0)
        default_by = offset_y - (h_float / 2.0)

        feature_payload = {
            'page': str(self.current_page + 1),
            'x': str(rect.x()),
            'y': str(rect.y()),
            'w': str(rect.width()),
            'h': str(rect.height()),
            'zoom': f"{self.pdf_view._current_scale:.4f}",
            'method': '',
            'nominal': '',
            'lsl': '',
            'usl': '',
            'bx': f"{default_bx:.4f}",
            'by': f"{default_by:.4f}",
            'br': str(radius_value),
        }

        new_feature = add_feature(self.pdf_path, feature_payload)
        new_feature['_pdf'] = self.pdf_path
        new_feature['br'] = str(radius_value)
        self._suppress_auto_focus = True
        self.rows.append(new_feature)
        self.selected_feature_id = new_feature.get('id')
        self.refresh_table()

        try:
            page_idx = int(new_feature.get('page', '1')) - 1
        except ValueError:
            page_idx = 0
        if page_idx == self.current_page:
            self._add_balloon_item(new_feature)
            self._apply_balloon_selection_visuals()

        fid = new_feature.get('id')
        if fid:
            self._push_undo(lambda feature_id=fid: self._undo_added_feature(feature_id), f'Add {fid}')

        self.statusBar().showMessage(f"Added {new_feature.get('id', '')}", 3000)

    def export_pdfs(self):
        if not self.pdf_path:
            QtWidgets.QMessageBox.information(self, 'Export PDFs', 'Open a PDF before exporting.')
            return

        default_dir = Path(self.pdf_path).parent if self.pdf_path else Path.home()
        default_base = Path(self.pdf_path).stem if self.pdf_path else 'axis'
        default_path = default_dir / f'{default_base}_inspection.pdf'
        selection, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            'Export PDFs',
            str(default_path),
            'PDF Files (*.pdf)'
        )
        if not selection:
            return

        inspection_path, balloon_path = self._derive_export_paths(Path(selection))

        try:
            inspection_path.parent.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass

        try:
            rows = self._collect_visible_rows()
            self._export_inspection_pdf(inspection_path, rows)
            self._export_ballooned_pdf(balloon_path)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, 'Export PDFs', f'Unable to export PDFs.\n{exc}')
            return

        msg_lines = [f'Inspection: {inspection_path}', f'Ballooned: {balloon_path}']
        QtWidgets.QMessageBox.information(self, 'Export PDFs', 'Created files:\n' + '\n'.join(f'- {line}' for line in msg_lines))
        self.statusBar().showMessage(f'Exported PDFs to {inspection_path.parent}', 5000)

    def export_results(self):
        if not self.pdf_path:
            QtWidgets.QMessageBox.information(self, 'Export Results', 'Open a PDF before exporting results.')
            return
        table = getattr(self, 'table', None)
        if table is None or table.rowCount() == 0:
            QtWidgets.QMessageBox.information(self, 'Export Results', 'No rows available to export.')
            return
        rows = self._collect_visible_rows()
        if not rows:
            QtWidgets.QMessageBox.information(self, 'Export Results', 'The current filters produced no rows to export.')
            return
        headers = []
        for col in range(table.columnCount()):
            header_item = table.horizontalHeaderItem(col)
            headers.append(header_item.text() if header_item else f'Column {col + 1}')
        base = Path(self.pdf_path).stem
        suffix_parts = ['results']
        if self.mode == 'Inspection' and self.current_wo:
            safe_wo = re.sub(r'[^A-Za-z0-9_-]+', '_', self.current_wo.strip()) or 'inspection'
            suffix_parts.append(safe_wo)
        default_name = f"{base}_{'_'.join(suffix_parts)}.csv"
        default_dir = Path(self.pdf_path).parent if self.pdf_path else Path.home()
        initial_path = default_dir / default_name
        selection, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            'Export Results',
            str(initial_path),
            'CSV Files (*.csv)'
        )
        if not selection:
            return
        try:
            with open(selection, 'w', newline='', encoding='utf-8') as handle:
                writer = csv.writer(handle)
                writer.writerow(headers)
                writer.writerows(rows)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, 'Export Results', f'Unable to write CSV.\n{exc}')
            return
        QtWidgets.QMessageBox.information(self, 'Export Results', f'Created {selection}')
        self.statusBar().showMessage(f'Exported results to {selection}', 4000)

    def export_all_results(self):
        if not self.pdf_path:
            QtWidgets.QMessageBox.information(self, 'Export All Results', 'Open a PDF before exporting results.')
            return
        workorders = list_workorders(self.pdf_path)
        if not workorders:
            QtWidgets.QMessageBox.information(self, 'Export All Results', 'No inspection results were found for this PDF.')
            return
        if not self.rows:
            self._load_rows()
        feature_map = {row.get('id'): row for row in self.rows if row.get('id')}
        export_rows = []
        for workorder in workorders:
            results = read_wo(self.pdf_path, workorder)
            if not results:
                continue
            for fid, result_value in results.items():
                if not result_value:
                    continue
                feature = feature_map.get(fid)
                if not feature:
                    continue
                normalized = self._normalize_result_entry(result_value)
                status = self._status_from_fields(
                    normalized.strip(),
                    (feature.get('lsl') or '').strip(),
                    (feature.get('usl') or '').strip()
                )
                export_rows.append([
                    workorder,
                    fid,
                    feature.get('page', ''),
                    (feature.get('method') or ''),
                    (feature.get('nominal') or ''),
                    (feature.get('lsl') or ''),
                    (feature.get('usl') or ''),
                    normalized,
                    status,
                ])
        if not export_rows:
            QtWidgets.QMessageBox.information(self, 'Export All Results', 'No measurement values were found to export.')
            return
        headers = ['Work Order', 'Feature ID', 'Page', 'Method', 'Nominal', 'LSL', 'USL', 'Result', 'Status']
        base = Path(self.pdf_path).stem
        default_name = f'{base}_all_results.csv'
        default_dir = Path(self.pdf_path).parent if self.pdf_path else Path.home()
        initial_path = default_dir / default_name
        selection, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            'Export All Results',
            str(initial_path),
            'CSV Files (*.csv)'
        )
        if not selection:
            return
        try:
            with open(selection, 'w', newline='', encoding='utf-8') as handle:
                writer = csv.writer(handle)
                writer.writerow(headers)
                writer.writerows(export_rows)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, 'Export All Results', f'Unable to write CSV.\n{exc}')
            return
        QtWidgets.QMessageBox.information(self, 'Export All Results', f'Created {selection}')
        self.statusBar().showMessage(f'Exported all results to {selection}', 4000)

    def _collect_visible_rows(self):
        rows = []
        column_count = self.table.columnCount()
        for row_idx in range(self.table.rowCount()):
            row_values = []
            for col_idx in range(column_count):
                item = self.table.item(row_idx, col_idx)
                if item is not None:
                    row_values.append(item.text())
                else:
                    widget = self.table.cellWidget(row_idx, col_idx)
                    if isinstance(widget, QtWidgets.QComboBox):
                        row_values.append(widget.currentText())
                    else:
                        row_values.append('')
            # Ensure status reflects the latest logic even if the table paint left it blank
            if column_count > 7:
                status_text = (row_values[7] or '').strip().upper()
                if status_text not in ('PASS', 'FAIL', '—'):
                    result_text = (row_values[3] or '').strip()
                    lsl_text = (row_values[5] or '').strip()
                    usl_text = (row_values[6] or '').strip()
                    computed = '—'
                    if result_text:
                        upper = result_text.upper()
                        if upper in ('PASS', 'FAIL'):
                            computed = upper
                        else:
                            try:
                                value = float(result_text)
                                lsl_val = float(lsl_text) if lsl_text else None
                                usl_val = float(usl_text) if usl_text else None
                                if lsl_val is not None and usl_val is not None:
                                    computed = 'PASS' if lsl_val <= value <= usl_val else 'FAIL'
                            except Exception:
                                computed = '—'
                    row_values[7] = computed
            rows.append(row_values)
        return rows

    def _derive_export_paths(self, selected: Path) -> tuple[Path, Path]:
        selected_path = selected
        if selected_path.suffix.lower() != '.pdf':
            selected_path = selected_path.with_suffix('.pdf')
        base_dir = selected_path.parent
        stem = selected_path.stem
        inspection_suffix = '_inspection'
        if stem.endswith(inspection_suffix) and len(stem) > len(inspection_suffix):
            root_stem = stem[:-len(inspection_suffix)]
        else:
            root_stem = stem
        if not root_stem:
            root_stem = Path(self.pdf_path).stem if self.pdf_path else 'axis'
        inspection_name = f'{root_stem}{inspection_suffix}.pdf'
        ballooned_name = f'{root_stem}_ballooned.pdf'
        inspection_path = selected_path.with_name(inspection_name)
        balloon_path = selected_path.with_name(ballooned_name)
        return inspection_path, balloon_path

    def _export_inspection_pdf(self, destination: Path, rows):
        doc = fitz.open()
        try:
            page_rect = fitz.paper_rect('letter')
        except Exception:
            page_rect = fitz.Rect(0, 0, 612, 792)
        if page_rect.width > page_rect.height:
            # enforce portrait orientation for the inspection report
            page_rect = fitz.Rect(0, 0, page_rect.height, page_rect.width)
        margin = 36
        line_height = 16
        columns = [
            ('ID', 60, 0),
            ('Page', 40, 1),
            ('Method', 110, 0),
            ('Result', 80, 2),
            ('Nominal', 80, 2),
            ('LSL', 80, 2),
            ('USL', 80, 2),
            ('Status', 65, 1),
        ]
        def draw_header(page, y):
            header_height = line_height + 6
            x = margin
            for title, width, _ in columns:
                rect = fitz.Rect(x, y, x + width, y + header_height)
                page.draw_rect(rect, color=(0.7, 0.7, 0.7), fill=(0.9, 0.9, 0.9), width=0.5)
                page.insert_textbox(rect, title, fontsize=10, fontname='Times-Bold', color=(0, 0, 0), align=1)
                x += width
            return y + header_height

        def start_page(include_title: bool):
            page = doc.new_page(width=page_rect.width, height=page_rect.height)
            y_pos = margin
            if include_title:
                title = f'Inspection Results - {Path(self.pdf_path).name}'
                page.insert_text((margin, y_pos), title, fontsize=14, fontname='Times-Bold', color=(0, 0, 0))
                y_pos += 20
                if self.current_wo:
                    wo_line = f'Work Order / Serial: {self.current_wo}'
                    page.insert_text((margin, y_pos), wo_line, fontsize=10, fontname='Times-Roman', color=(0, 0, 0))
                    y_pos += 16
                timestamp = datetime.now().strftime('%Y-%m-%d %H:%M')
                page.insert_text((margin, y_pos), f'Exported: {timestamp}', fontsize=9, fontname='Times-Roman', color=(0.3, 0.3, 0.3))
                y_pos += 18
            elif self.current_wo:
                # repeat WO on subsequent pages for clarity
                wo_line = f'Work Order / Serial: {self.current_wo}'
                page.insert_text((margin, y_pos), wo_line, fontsize=10, fontname='Times-Roman', color=(0, 0, 0))
                y_pos += 16
            y_pos = draw_header(page, y_pos)
            return page, y_pos

        def ensure_space(page, y_pos):
            if y_pos + line_height > page_rect.height - margin:
                return start_page(False)
            return page, y_pos

        status_colors = {
            'PASS': ((200 / 255, 235 / 255, 200 / 255), (0, 0, 0)),
            'FAIL': ((247 / 255, 205 / 255, 205 / 255), (0.55, 0, 0)),
        }

        def compute_status_from_row(row_values: list[str]) -> str:
            if not row_values:
                return '—'
            status_raw = row_values[7] if len(row_values) > 7 else ''
            if status_raw:
                upper = status_raw.strip().upper()
                if upper in ('PASS', 'FAIL'):
                    return upper
            result_text = row_values[3].strip() if len(row_values) > 3 and row_values[3] else ''
            if not result_text:
                return '—'
            upper_res = result_text.upper()
            if upper_res in ('PASS', 'FAIL'):
                return upper_res
            try:
                value = float(result_text)
            except Exception:
                return '—'
            try:
                lsl_val = float(row_values[5]) if len(row_values) > 5 and row_values[5] else None
            except Exception:
                lsl_val = None
            try:
                usl_val = float(row_values[6]) if len(row_values) > 6 and row_values[6] else None
            except Exception:
                usl_val = None
            if lsl_val is not None and usl_val is not None:
                return 'PASS' if lsl_val <= value <= usl_val else 'FAIL'
            return '—'

        page, cursor_y = start_page(True)

        if not rows:
            info_text = 'No inspection rows available for the current filters.'
            page.insert_text((margin, cursor_y + 12), info_text, fontsize=11, fontname='Times-Roman', color=(0.25, 0.25, 0.25))
        else:
            for row in rows:
                row_len = len(row)
                page, cursor_y = ensure_space(page, cursor_y)
                x = margin
                status_value = compute_status_from_row(row)
                status_key = status_value
                fill_color, text_color = status_colors.get(status_key, ((0.92, 0.92, 0.92), (0.2, 0.2, 0.2)))
                for idx, (title, width, align) in enumerate(columns):
                    value = row[idx] if idx < row_len else ''
                    rect = fitz.Rect(x, cursor_y, x + width, cursor_y + line_height)
                    if idx in (3, 7):
                        page.draw_rect(rect, color=None, fill=fill_color, width=0)
                    if idx == 7:
                        if status_key in status_colors:
                            font_name = 'Times-Bold'
                            color = text_color
                            display_value = status_key
                        else:
                            font_name = 'Times-Roman'
                            color = (0, 0, 0)
                            display_value = status_value or '-'
                    else:
                        font_name = 'Times-Roman'
                        color = (0, 0, 0)
                        display_value = value or ''
                    clean_value = str(display_value).replace('\n', ' ')
                    page.insert_textbox(rect, clean_value, fontsize=9.5, fontname=font_name, color=color, align=align)
                    x += width
                cursor_y += line_height

        doc.save(destination, garbage=4, deflate=True)
        doc.close()

    def _export_ballooned_pdf(self, destination: Path):
        doc = fitz.open(self.pdf_path)
        scale = 1.0 / PAGE_RENDER_ZOOM
        circle_color = (220 / 255, 40 / 255, 40 / 255)
        fill_color = (1.0, 230 / 255, 230 / 255)
        try:
            for page_index in range(doc.page_count):
                page = doc.load_page(page_index)
                for feature in self.rows:
                    try:
                        feature_page = int(feature.get('page', '1')) - 1
                    except Exception:
                        feature_page = 0
                    if feature_page != page_index:
                        continue
                    try:
                        x = float(feature.get('x', 0)) * scale
                        y = float(feature.get('y', 0)) * scale
                        w = float(feature.get('w', 0)) * scale
                        h = float(feature.get('h', 0)) * scale
                        bx = float(feature.get('bx', 0)) * scale
                        by = float(feature.get('by', 0)) * scale
                        radius = float(feature.get('br', self.default_balloon_radius)) * scale
                    except Exception:
                        continue
                    center_x = x + (w / 2.0) + bx
                    center_y = y + (h / 2.0) + by
                    page.draw_circle((center_x, center_y), radius, color=circle_color, fill=fill_color, width=1.5)
                    label = feature.get('id', '')
                    match = re.search(r'(\d+)$', label)
                    text = match.group(1) if match else label
                    font_size = max(8.0, radius * 1.15)
                    text_width = fitz.get_text_length(text, fontname='Times-Bold', fontsize=font_size)
                    ascent = font_size * 0.7
                    x_pos = center_x - (text_width / 2.0)
                    y_pos = center_y + (ascent / 2.0)
                    page.insert_text((x_pos, y_pos), text, fontsize=font_size, fontname='Times-Bold', color=(0, 0, 0))
            doc.save(destination, garbage=4, deflate=True)
        finally:
            doc.close()

    def _balloon_size_changed(self, value: int):
        self.default_balloon_radius = value
        if getattr(self, '_syncing_balloon_spin', False):
            return
        if self.mode != 'Ballooning' or not self.pdf_path:
            return
        updated = False
        for feature in self.rows:
            if feature.get('br') != str(value):
                feature['br'] = str(value)
                updated = True
        if updated:
            self._persist_rows_to_master()
        for item in self.balloon_items:
            item.set_radius(float(value))

    def _sync_balloon_size_spin(self, value: float):
        if self.balloon_size_spin is None:
            return
        self._syncing_balloon_spin = True
        self.balloon_size_spin.blockSignals(True)
        self.balloon_size_spin.setValue(int(round(value)))
        self.balloon_size_spin.blockSignals(False)
        self._syncing_balloon_spin = False
        try:
            self.default_balloon_radius = int(round(value))
        except Exception:
            self.default_balloon_radius = BALLOON_RADIUS

    def _update_balloon_size_controls_enabled(self):
        if self.balloon_size_label is None or self.balloon_size_spin is None:
            return
        enabled = bool(self.pdf_path) and self.mode == 'Ballooning'
        self.balloon_size_label.setEnabled(enabled)
        self.balloon_size_spin.setEnabled(enabled)

    def _page_spin_changed(self, value: int):
        if self._syncing_page_spin or not self.doc:
            return
        target = max(0, min(value - 1, max(0, self.doc.page_count - 1)))
        if target == self.current_page:
            return
        self._balloons_built_for_page = None  # [ZOOM-DEBOUNCE]
        self.current_page = target
        self._render_current_page()
        self._apply_balloon_selection_visuals()
        self._clear_highlight_rect()
        self.fit_view()

    def _sync_page_spin(self):
        if self.page_spin is None:
            return
        self._syncing_page_spin = True
        try:
            self.page_spin.blockSignals(True)
            value = self.current_page + 1 if self.doc else 1
            self.page_spin.setValue(value)
        finally:
            self.page_spin.blockSignals(False)
            self._syncing_page_spin = False

    def _update_page_controls_enabled(self):
        if self.page_label is None or self.page_spin is None:
            return
        has_pdf = self.doc is not None and self.total_pages > 0
        self.page_label.setEnabled(has_pdf)
        self.page_spin.setEnabled(has_pdf)
        max_page = max(1, self.total_pages)
        self.page_spin.setMinimum(1)
        self.page_spin.setMaximum(max_page)

    def _persist_rows_to_master(self):
        if not self.pdf_path:
            return
        payload = []
        for row in self.rows:
            payload.append({key: row.get(key, '') for key in MASTER_HEADER})
        write_master(self.pdf_path, payload)

    def _push_undo(self, handler, description: str):
        if not callable(handler):
            return
        if self._undo_in_progress:
            return
        if self.mode != 'Ballooning':
            return
        self._undo_stack.append((description, handler))
        max_depth = 50
        if len(self._undo_stack) > max_depth:
            self._undo_stack.pop(0)

    def _undo_last_action(self):
        if self.mode != 'Ballooning':
            self.statusBar().showMessage('Undo is only available in Ballooning mode.', 2000)
            return
        if not self._undo_stack:
            self.statusBar().showMessage('Nothing to undo.', 2000)
            return
        description, handler = self._undo_stack.pop()
        self._undo_in_progress = True
        try:
            handler()
        except Exception as exc:
            self.statusBar().showMessage('Undo failed.', 4000)
            print(f'Undo failed: {exc}', file=sys.stderr)
        finally:
            self._undo_in_progress = False
        if description:
            self.statusBar().showMessage(f'Undo: {description}', 3000)

    def _undo_added_feature(self, fid: str):
        if not fid:
            return
        self._remove_feature(fid, persist=True)
        self.statusBar().showMessage(f'Removed {fid}', 3000)

    def _undo_deleted_feature(self, snapshot: dict):
        if not snapshot:
            return
        fid = snapshot.get('id')
        self._restore_feature_snapshot(snapshot)
        if fid:
            self.statusBar().showMessage(f'Restored {fid}', 3000)

    def _restore_feature_snapshot(self, snapshot: dict):
        if not snapshot or not self.pdf_path:
            return
        snapshot_copy = dict(snapshot)
        row_index = snapshot_copy.pop('_row_index', None)
        fid = snapshot_copy.get('id')
        if not fid:
            return
        master_row = {key: snapshot_copy.get(key, '') for key in MASTER_HEADER}
        try:
            rows = read_master(self.pdf_path)
        except Exception:
            rows = []
        rows = [dict(r) for r in rows if r.get('id') != fid]
        rows.append(master_row)

        def sort_key(row):
            label = row.get('id') or ''
            match = re.search(r'(\d+)$', label)
            return int(match.group(1)) if match else 0

        rows.sort(key=sort_key)
        write_master(self.pdf_path, rows)

        snapshot_copy['_pdf'] = self.pdf_path
        self.rows = [r for r in self.rows if r.get('id') != fid]
        if row_index is None or row_index < 0 or row_index > len(self.rows):
            self.rows.append(snapshot_copy)
        else:
            self.rows.insert(row_index, snapshot_copy)
        self.selected_feature_id = fid
        self._suppress_auto_focus = True
        self.refresh_table()
        self._rebuild_balloons()
        self._apply_balloon_selection_visuals()

    def _remove_feature(self, fid: str, row_hint: int | None = None, persist: bool = True) -> dict | None:
        if not fid:
            return None
        feature_snapshot = None
        index = None
        for idx, feature in enumerate(self.rows):
            if feature.get('id') == fid:
                feature_snapshot = dict(feature)
                index = idx
                break
        if feature_snapshot is None:
            return None
        if persist and self.pdf_path:
            delete_feature(self.pdf_path, fid)
        if self.selected_feature_id == fid:
            self.selected_feature_id = None
            self._clear_highlight_rect()
        self.rows = [r for r in self.rows if r.get('id') != fid]
        feature_snapshot['_row_index'] = index if index is not None else -1
        self.refresh_table()
        if row_hint is not None and self.table.rowCount() > 0:
            target = max(0, min(row_hint, self.table.rowCount() - 1))
            self.table.selectRow(target)
        self._rebuild_balloons()
        self._apply_balloon_selection_visuals()
        return feature_snapshot


class SPCChartWidget(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._feature_data: spc.FeatureSPCData | None = None
        self.setMinimumSize(320, 220)

    def set_feature_data(self, data: spc.FeatureSPCData | None):
        self._feature_data = data
        self.update()

    def paintEvent(self, event):
        painter = QtGui.QPainter(self)
        painter.fillRect(self.rect(), QtGui.QColor(23, 23, 23))
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing, True)
        if not self._feature_data or not self._feature_data.measurements:
            painter.setPen(QtGui.QColor(200, 200, 200))
            painter.drawText(self.rect(), QtCore.Qt.AlignmentFlag.AlignCenter, 'No data')
            return
        rect = QtCore.QRectF(self.rect().adjusted(24, 16, -16, -32))
        if rect.width() <= 0 or rect.height() <= 0:
            return
        values = [m.value for m in self._feature_data.measurements]
        min_val = min(values)
        max_val = max(values)
        if self._feature_data.lsl is not None:
            min_val = min(min_val, self._feature_data.lsl)
        if self._feature_data.usl is not None:
            max_val = max(max_val, self._feature_data.usl)
        span = max_val - min_val
        margin = span * 0.1 if span > 1e-6 else 1.0
        min_val -= margin
        max_val += margin
        painter.setPen(QtGui.QPen(QtGui.QColor(120, 120, 120), 1))
        painter.drawRect(rect)
        self._draw_spec_line(painter, rect, min_val, max_val, self._feature_data.lsl, QtGui.QColor(255, 140, 140))
        self._draw_spec_line(painter, rect, min_val, max_val, self._feature_data.usl, QtGui.QColor(140, 200, 255))
        path = QtGui.QPainterPath()
        count = len(values)
        for idx, value in enumerate(values):
            x_ratio = idx / (count - 1) if count > 1 else 0.0
            x = rect.left() + x_ratio * rect.width()
            y = self._value_to_y(rect, min_val, max_val, value)
            if idx == 0:
                path.moveTo(x, y)
            else:
                path.lineTo(x, y)
            painter.setBrush(QtGui.QBrush(QtGui.QColor(200, 200, 90)))
            painter.setPen(QtGui.QPen(QtGui.QColor(200, 200, 90), 2))
            painter.drawEllipse(QtCore.QPointF(x, y), 3, 3)
        painter.setPen(QtGui.QPen(QtGui.QColor(160, 220, 255), 2))
        painter.drawPath(path)
        painter.setPen(QtGui.QColor(200, 200, 200))
        painter.drawText(QtCore.QPointF(rect.right() - 90, rect.bottom() + 20), 'Sample #')

    def _value_to_y(self, rect: QtCore.QRectF, min_val: float, max_val: float, value: float) -> float:
        if max_val - min_val < 1e-6:
            return rect.center().y()
        ratio = (value - min_val) / (max_val - min_val)
        return rect.bottom() - ratio * rect.height()

    def _draw_spec_line(self, painter: QtGui.QPainter, rect: QtCore.QRectF, min_val: float, max_val: float, value: float | None, color: QtGui.QColor):
        if value is None:
            return
        y = self._value_to_y(rect, min_val, max_val, value)
        painter.setPen(QtGui.QPen(color, 1, QtCore.Qt.PenStyle.DashLine))
        painter.drawLine(QtCore.QLineF(rect.left(), y, rect.right(), y))


class SPCDialog(QtWidgets.QDialog):
    def __init__(self, parent: QtWidgets.QWidget | None, pdf_path: str, dataset: dict[str, spc.FeatureSPCData]):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.dataset = dataset
        self.setWindowTitle('SPC Dashboard')
        self.resize(1000, 560)

        layout = QtWidgets.QHBoxLayout(self)
        self.feature_list = QtWidgets.QListWidget()
        self.feature_list.currentItemChanged.connect(self._selection_changed)
        layout.addWidget(self.feature_list, 1)

        right_panel = QtWidgets.QVBoxLayout()
        layout.addLayout(right_panel, 2)

        self.stats_labels: dict[str, QtWidgets.QLabel] = {}
        stats_form = QtWidgets.QFormLayout()
        for key, title in (
            ('count', 'Samples'),
            ('mean', 'Mean'),
            ('stdev', 'Std Dev'),
            ('min', 'Min'),
            ('max', 'Max'),
            ('cp', 'Cp'),
            ('cpk', 'Cpk'),
        ):
            lbl = QtWidgets.QLabel('—')
            self.stats_labels[key] = lbl
            stats_form.addRow(f'{title}:', lbl)
        right_panel.addLayout(stats_form)

        self.chart_widget = SPCChartWidget()
        right_panel.addWidget(self.chart_widget, 1)

        self.measurement_table = QtWidgets.QTableWidget(0, 4)
        self.measurement_table.setHorizontalHeaderLabels(['Work Order', 'Value', 'Timestamp', 'Source'])
        header = self.measurement_table.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QtWidgets.QHeaderView.ResizeMode.Stretch)
        right_panel.addWidget(self.measurement_table, 1)

        self._populate_feature_list()

    def _populate_feature_list(self):
        self.feature_list.clear()
        for fid in sorted(self.dataset.keys(), key=lambda s: s.lower()):
            data = self.dataset[fid]
            item = QtWidgets.QListWidgetItem(f'{fid} ({len(data.measurements)})')
            item.setData(QtCore.Qt.ItemDataRole.UserRole, fid)
            self.feature_list.addItem(item)
        if self.feature_list.count() > 0:
            self.feature_list.setCurrentRow(0)

    def _selection_changed(self, current, previous):
        if current is None:
            self.chart_widget.set_feature_data(None)
            self._update_stats(None)
            self._populate_measurement_table([])
            return
        fid = current.data(QtCore.Qt.ItemDataRole.UserRole)
        data = self.dataset.get(fid)
        self.chart_widget.set_feature_data(data)
        self._update_stats(data)
        measurements = sorted(
            data.measurements,
            key=lambda m: (m.timestamp or datetime.min, m.workorder)
        )
        self._populate_measurement_table(measurements)

    def _update_stats(self, data: spc.FeatureSPCData | None):
        if not data:
            for lbl in self.stats_labels.values():
                lbl.setText('—')
            return
        stats = data.stats
        self.stats_labels['count'].setText(str(stats.count))
        self.stats_labels['mean'].setText(spc.format_stat(stats.mean))
        self.stats_labels['stdev'].setText(spc.format_stat(stats.stdev))
        self.stats_labels['min'].setText(spc.format_stat(stats.min_value))
        self.stats_labels['max'].setText(spc.format_stat(stats.max_value))
        self.stats_labels['cp'].setText(spc.format_stat(stats.cp))
        self.stats_labels['cpk'].setText(spc.format_stat(stats.cpk))

    def _populate_measurement_table(self, measurements: list[spc.Measurement]):
        self.measurement_table.setRowCount(len(measurements))
        for row, measurement in enumerate(measurements):
            ts_text = measurement.timestamp.strftime('%Y-%m-%d %H:%M') if measurement.timestamp else '—'
            row_values = (
                measurement.workorder,
                f'{measurement.value:.4f}',
                ts_text,
                measurement.source_path,
            )
            for col, text in enumerate(row_values):
                item = QtWidgets.QTableWidgetItem(text)
                item.setFlags(item.flags() & ~QtCore.Qt.ItemFlag.ItemIsEditable)
                self.measurement_table.setItem(row, col, item)


def main():
    app = QtWidgets.QApplication(sys.argv)
    win = MainWindow()
    win.resize(1200, 800)
    win.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
