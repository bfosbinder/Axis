import sys
import os
import csv
from pathlib import Path
import fitz  # PyMuPDF
from PyQt6 import QtWidgets, QtGui, QtCore
from storage import ensure_master, read_master, add_feature, update_feature, delete_feature, read_wo, write_wo, write_master, parse_tolerance_expression, normalize_str, MASTER_HEADER, list_workorders


BALLOON_RADIUS = 14


def format_number(value: float, decimals: int = 6) -> str:
    text = f"{value:.{decimals}f}"
    if '.' in text:
        text = text.rstrip('0').rstrip('.')
    if text == '-0':
        text = '0'
    return text


class StartSessionDialog(QtWidgets.QDialog):
    """Prompt for Serial Number (Inspection) or start Ballooning mode."""
    def __init__(self, parent=None, previous_orders=None):
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
        btns.addWidget(self.btn_balloon)
        btns.addWidget(self.btn_cancel)
        lay.addLayout(btns)

        self.btn_inspect.clicked.connect(self._go_inspect)
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
        self.setRenderHints(
            QtGui.QPainter.RenderHint.Antialiasing | QtGui.QPainter.RenderHint.SmoothPixmapTransform
        )
        self.setDragMode(QtWidgets.QGraphicsView.DragMode.ScrollHandDrag)
        self.setTransformationAnchor(QtWidgets.QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setResizeAnchor(QtWidgets.QGraphicsView.ViewportAnchor.AnchorUnderMouse)

    def load_page(self, qimage: QtGui.QImage):
        self.scene().clear()
        pix = QtGui.QPixmap.fromImage(qimage)
        self._pixmap_item = self.scene().addPixmap(pix)
        self.setSceneRect(self._pixmap_item.boundingRect())

    def fit_to_view(self):
        if self.scene() is not None:
            self.fitInView(self.sceneRect(), QtCore.Qt.AspectRatioMode.KeepAspectRatio)
            self._current_scale = 1.0

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
        self._current_scale = new_scale
        event.accept()
        
        # keep rubber band consistent
        if self._rubber.isVisible():
            self._rubber.hide()

    def set_zoom(self, target: float, focus_point: QtCore.QPointF = None):
        if self._pixmap_item is None:
            return
        target = max(self._min_scale, min(self._max_scale, float(target)))
        if self._current_scale == 0:
            self._current_scale = 1.0
        factor = target / self._current_scale
        if abs(factor - 1.0) > 1e-6:
            self.scale(factor, factor)
            self._current_scale = target
        if focus_point is not None:
            self.centerOn(focus_point)
            self.ensureVisible(QtCore.QRectF(focus_point.x() - 20, focus_point.y() - 20, 40, 40), 10, 10)

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.MouseButton.LeftButton and self.controller.mode == 'Ballooning' and self.controller.pick_on_print:
            self._rubber_origin = event.pos()
            self._rubber.setGeometry(QtCore.QRect(self._rubber_origin, QtCore.QSize()))
            self._rubber.show()
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._rubber.isVisible():
            self._rubber.setGeometry(QtCore.QRect(self._rubber_origin, event.pos()).normalized())
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
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

        self._build_ui()

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

        fit_btn = QtWidgets.QPushButton('Fit')
        fit_btn.clicked.connect(self.fit_view)
        toolbar.addWidget(fit_btn)

        self.show_balloon_btn = QtWidgets.QPushButton('Hide Balloons')
        self.show_balloon_btn.setCheckable(True)
        self.show_balloon_btn.toggled.connect(self.toggle_show_balloons)
        toolbar.addWidget(self.show_balloon_btn)

        ocr_btn = QtWidgets.QPushButton('OCR')
        ocr_btn.clicked.connect(lambda: QtWidgets.QMessageBox.information(self, 'OCR', 'OCR is stubbed.'))
        toolbar.addWidget(ocr_btn)

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

        export_btn = QtWidgets.QPushButton('Export Filtered CSV')
        export_btn.clicked.connect(self.export_filtered)
        toolbar.addWidget(export_btn)

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
        self.method_filter = QtWidgets.QLineEdit()
        self.method_filter.textChanged.connect(self.refresh_table)
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
            sel_method = self.method_filter.text().strip()
            if sel_method and sel_method.lower() not in method.lower():
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
            if self.mode == 'Ballooning':
                combo = QtWidgets.QComboBox()
                combo.setEditable(True)
                combo.setInsertPolicy(QtWidgets.QComboBox.InsertPolicy.NoInsert)
                combo.setSizeAdjustPolicy(QtWidgets.QComboBox.SizeAdjustPolicy.AdjustToContents)
                combo.setEnabled(True)
                combo.setPlaceholderText('Method')
                line_edit = combo.lineEdit()
                if line_edit is not None:
                    line_edit.setPlaceholderText('Method')
                combo.blockSignals(True)
                combo.setCurrentText(method)
                combo.blockSignals(False)
                combo.currentTextChanged.connect(lambda text, r=row: self._method_combo_changed(r, text))
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

        status = '—'
        if result_text:
            upper = result_text.upper()
            if upper in ('PASS', 'FAIL'):
                status = upper
            else:
                try:
                    value = float(result_text)
                    lsl_val = float(lsl_text) if lsl_text else None
                    usl_val = float(usl_text) if usl_text else None
                    if lsl_val is not None and usl_val is not None:
                        status = 'PASS' if lsl_val <= value <= usl_val else 'FAIL'
                except Exception:
                    status = '—'

        if status_item:
            self.table.blockSignals(True)
            status_item.setText(status)
            self.table.blockSignals(False)
        self._apply_status_formatting(row, status)

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
        self.selected_feature_id = fid
        try:
            page_idx = int(feature.get('page', '1')) - 1
        except Exception:
            page_idx = 0
        page_idx = max(0, page_idx)
        if page_idx != self.current_page:
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

        delete_feature(self.pdf_path, fid)
        self.rows = [r for r in self.rows if r.get('id') != fid]

        try:
            table.blockSignals(True)
            table.removeRow(row)
        finally:
            table.blockSignals(False)

        for item in list(self.balloon_items):
            if item.feature.get('id') == fid:
                try:
                    self.pdf_view.scene().removeItem(item)
                except Exception:
                    pass
                self.balloon_items.remove(item)
                break

        if self.selected_feature_id == fid:
            self.selected_feature_id = None
            self._clear_highlight_rect()

        self._apply_balloon_selection_visuals()
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

    def _refresh_method_combobox_options(self):
        table = getattr(self, 'table', None)
        if table is None:
            return
        methods = []
        seen = set()
        for row in range(table.rowCount()):
            item = table.item(row, 2)
            if not item:
                continue
            value = item.text().strip()
            if not value:
                continue
            key = value.lower()
            if key in seen:
                continue
            seen.add(key)
            methods.append(value)
        methods.sort(key=lambda s: s.lower())

        for row in range(table.rowCount()):
            combo = table.cellWidget(row, 2)
            if not isinstance(combo, QtWidgets.QComboBox):
                continue
            current = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            combo.setEditable(self.mode == 'Ballooning')
            combo.setInsertPolicy(QtWidgets.QComboBox.InsertPolicy.NoInsert)
            combo.setSizeAdjustPolicy(QtWidgets.QComboBox.SizeAdjustPolicy.AdjustToContents)
            combo.addItem('')
            for value in methods:
                combo.addItem(value)
            combo.setCurrentText(current)
            combo.setEnabled(self.mode == 'Ballooning')
            combo.setPlaceholderText('Method')
            line_edit = combo.lineEdit()
            if line_edit is not None:
                line_edit.setPlaceholderText('Method')
            combo.blockSignals(False)

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

    def toggle_pick(self, checked: bool):
        if not self.pdf_path or self.mode != 'Ballooning':
            self.pick_btn.blockSignals(True)
            self.pick_btn.setChecked(False)
            self.pick_btn.blockSignals(False)
            self.pick_on_print = False
            if not self.pdf_path:
                QtWidgets.QMessageBox.information(self, 'No PDF', 'Open a PDF before placing balloons.')
            else:
                QtWidgets.QMessageBox.information(self, 'Inspection Mode', 'Pick-on-Print is only available in Ballooning mode.')
            return
        self.pick_on_print = checked
        self.pick_btn.setText('Picking...' if checked else 'Pick-on-Print')

    def fit_view(self):
        if self.pdf_view._pixmap_item:
            self.pdf_view.fit_to_view()

    def toggle_show_balloons(self, checked: bool):
        self.show_balloons = not checked
        self.show_balloon_btn.setText('Show Balloons' if not self.show_balloons else 'Hide Balloons')
        self._update_balloon_visibility()

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

    def _prompt_session(self, previous_orders=None) -> bool:
        dlg = StartSessionDialog(self, previous_orders)
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
            return

        self.doc = doc
        self.pdf_path = path
        self.current_page = 0
        master_path = f"{path}.balloons.csv"
        had_master = os.path.exists(master_path)
        ensure_master(path)
        existing_wos = list_workorders(path)
        self.total_pages = self.doc.page_count if self.doc else 0
        self._update_page_controls_enabled()
        self._sync_page_spin()
        self._load_rows()

        if had_master:
            if not self._prompt_session(existing_wos):
                self._clear_loaded_pdf()
                return
        else:
            # new PDF: default to ballooning without prompting
            self.mode = 'Ballooning'
            self.current_wo = None
            self._update_mode_ui()

        self._render_current_page()
        self.refresh_table()
        self.statusBar().showMessage(f'Loaded {Path(path).name}', 4000)

    def _clear_loaded_pdf(self):
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
        self.selected_feature_id = None
        self._clear_highlight_rect()
        if self.pdf_view.scene():
            self.pdf_view.scene().clear()
        self.table.setRowCount(0)
        self._update_mode_ui()
        self._update_page_controls_enabled()
        self._sync_page_spin()

    def _load_rows(self):
        if not self.pdf_path:
            self.rows = []
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
        self._sync_balloon_size_spin(self.default_balloon_radius)

    def _render_current_page(self):
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

        zoom = 1.5
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        qimage = QtGui.QImage.fromData(pix.tobytes('png'), 'PNG')
        self.pdf_view.load_page(qimage)
        self.pdf_view.fit_to_view()
        self._rebuild_balloons()
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

    def _clear_balloon_items(self):
        for item in self.balloon_items:
            try:
                self.pdf_view.scene().removeItem(item)
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

        self.statusBar().showMessage(f"Added {new_feature.get('id', '')}", 3000)

    def export_filtered(self):
        if self.table.rowCount() == 0:
            QtWidgets.QMessageBox.information(self, 'Export CSV', 'There are no rows to export.')
            return
        default_dir = Path(self.pdf_path).parent if self.pdf_path else Path.home()
        default_name = Path(self.pdf_path).stem + '_filtered.csv' if self.pdf_path else 'filtered.csv'
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            'Export Filtered CSV',
            str(default_dir / default_name),
            'CSV Files (*.csv)'
        )
        if not path:
            return
        with open(path, 'w', newline='', encoding='utf-8') as handle:
            writer = csv.writer(handle)
            writer.writerow(['ID', 'Page', 'Method', 'Result', 'Nominal', 'LSL', 'USL', 'Status'])
            for row_idx in range(self.table.rowCount()):
                row_values = []
                for col_idx in range(self.table.columnCount()):
                    item = self.table.item(row_idx, col_idx)
                    row_values.append(item.text() if item else '')
                writer.writerow(row_values)
        self.statusBar().showMessage(f'Exported {self.table.rowCount()} rows to {Path(path).name}', 4000)

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
        self.current_page = target
        self._render_current_page()
        self._apply_balloon_selection_visuals()
        self._clear_highlight_rect()

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

def main():
    app = QtWidgets.QApplication(sys.argv)
    win = MainWindow()
    win.resize(1200, 800)
    win.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
