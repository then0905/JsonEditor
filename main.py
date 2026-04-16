#!/usr/bin/env python3
"""JsonEditor — PySide6 rewrite"""

import sys
import os

import pandas as pd
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QSplitter, QTabWidget,
    QListWidget, QListWidgetItem, QLineEdit, QPushButton, QLabel,
    QVBoxLayout, QHBoxLayout, QScrollArea, QCheckBox, QComboBox,
    QTextEdit, QToolBar, QTableView, QHeaderView, QAbstractItemView,
    QStackedWidget, QFileDialog, QMenu, QSizePolicy, QFrame,
    QInputDialog, QMessageBox, QGridLayout, QStyledItemDelegate,
    QApplication,
)
from PySide6.QtCore import (
    Qt, Signal, QAbstractTableModel, QModelIndex,
    QTimer, QSize, QThread,
)
from PySide6.QtGui import QAction, QColor, QKeySequence, QFont, QBrush

from json_data_manager import JsonDataManager


# ── Background workers ────────────────────────────────────────────────────────

class _LoadWorker(QThread):
    done  = Signal()
    error = Signal(str)

    def __init__(self, manager, path):
        super().__init__()
        self._manager = manager
        self._path    = path

    def run(self):
        try:
            self._manager.load_json(self._path)
            self.done.emit()
        except Exception as e:
            self.error.emit(str(e))


class _SaveWorker(QThread):
    done  = Signal()
    error = Signal(str)

    def __init__(self, manager):
        super().__init__()
        self._manager = manager

    def run(self):
        try:
            self._manager.save_json()
            self.done.emit()
        except Exception as e:
            self.error.emit(str(e))


# ── Stylesheet ────────────────────────────────────────────────────────────────
APP_STYLE = """
* { font-family: "Segoe UI"; font-size: 11px; color: #dce4ee; }
QMainWindow, QWidget#root { background: #2b2b2b; }
QSplitter::handle:horizontal { background: #3a3a4a; width: 3px; }
QSplitter::handle:vertical   { background: #3a3a4a; height: 3px; }

/* Tabs */
QTabWidget::pane { border: 1px solid #3a3a4a; background: #2b2b2b; }
QTabBar::tab            { background: #1e1e2e; color: #888899; padding: 5px 16px; border: none; }
QTabBar::tab:selected   { background: #3B8ED0; color: white; }
QTabBar::tab:hover:!selected { background: #2a2a42; color: #dce4ee; }

/* Lists */
QListWidget { background: #252535; border: none; outline: none; }
QListWidget::item { padding: 5px 8px; }
QListWidget::item:selected          { background: #3B8ED0; color: white; }
QListWidget::item:hover:!selected   { background: #3a3a52; }

/* Line / Text edits */
QLineEdit {
    background: #343638; border: 1px solid #565B5E;
    padding: 3px 6px; border-radius: 2px;
    selection-background-color: #3B8ED0;
}
QLineEdit:focus { border-color: #3B8ED0; }
QLineEdit[invalid="true"] { border-color: #e05050; background: #3a1a1a; }
QTextEdit {
    background: #343638; border: 1px solid #565B5E;
    selection-background-color: #3B8ED0;
}
QTextEdit:focus { border-color: #3B8ED0; }

/* ComboBox */
QComboBox {
    background: #343638; border: 1px solid #565B5E;
    padding: 2px 6px; border-radius: 2px;
}
QComboBox::drop-down { border: none; width: 18px; }
QComboBox QAbstractItemView {
    background: #343638; border: 1px solid #565B5E;
    selection-background-color: #3B8ED0;
}

/* CheckBox */
QCheckBox::indicator {
    width: 14px; height: 14px;
    border: 1px solid #565B5E; background: #343638; border-radius: 2px;
}
QCheckBox::indicator:checked { background: #3B8ED0; border-color: #3B8ED0; }

/* Buttons */
QPushButton {
    background: #3a3a52; border: none; padding: 4px 12px; border-radius: 2px;
}
QPushButton:hover   { background: #4a4a62; }
QPushButton:pressed { background: #2a2a42; }
QPushButton[accent="add"]        { background: #2d8a4e; }
QPushButton[accent="add"]:hover  { background: #3aaa5e; }
QPushButton[accent="del"]        { background: #8b3a3a; }
QPushButton[accent="del"]:hover  { background: #aa4a4a; }
QPushButton[accent="blue"]       { background: #3B8ED0; }
QPushButton[accent="blue"]:hover { background: #2a6ea0; }

/* Table */
QTableView {
    background: #2b2b2b; gridline-color: #3a3a4a;
    border: none; outline: none;
    selection-background-color: #3B8ED0;
}
QTableView::item:hover:!selected { background: #323246; }
QHeaderView::section {
    background: #404040; padding: 4px 6px;
    border: 1px solid #3a3a4a; font-weight: bold;
}
QHeaderView::section:hover   { background: #505060; }
QHeaderView::section:pressed { background: #3B8ED0; color: white; }

/* Scrollbars */
QScrollBar:vertical   { background: #2b2b2b; width: 10px; border: none; }
QScrollBar:horizontal { background: #2b2b2b; height: 10px; border: none; }
QScrollBar::handle:vertical   { background: #555; min-height: 24px; border-radius: 4px; margin: 1px; }
QScrollBar::handle:horizontal { background: #555; min-width:  24px; border-radius: 4px; margin: 1px; }
QScrollBar::handle:vertical:hover, QScrollBar::handle:horizontal:hover { background: #777; }
QScrollBar::add-line, QScrollBar::sub-line { width: 0; height: 0; }

/* Toolbar */
QToolBar {
    background: #1e1e2e; border: none;
    border-bottom: 1px solid #3a3a52; spacing: 2px; padding: 4px 6px;
}
QToolBar QToolButton { background: transparent; border: none; padding: 6px 10px; border-radius: 3px; }
QToolBar QToolButton:hover   { background: #2e2e42; }
QToolBar QToolButton:pressed { background: #3B8ED0; }
QToolBar::separator { background: #3a3a52; width: 1px; margin: 6px 4px; }

/* Status bar */
QStatusBar { background: #13131e; border-top: 1px solid #3a3a52; }
QStatusBar::item { border: none; }

/* Menus */
QMenu { background: #252535; border: 1px solid #3a3a52; padding: 4px; }
QMenu::item { padding: 5px 20px; }
QMenu::item:selected { background: #3B8ED0; color: white; }
QMenu::separator { height: 1px; background: #3a3a52; margin: 3px 8px; }

/* Panels */
QFrame#panel { background: #252535; }
QLabel#section-hdr {
    background: #2a2a3a; color: #A0C4E8; font-weight: bold;
    padding: 5px 8px; border-bottom: 1px solid #3a3a4a;
}
"""

_DIRTY_BG  = "#2a2a10"
_ROW_EVEN  = QColor("#2b2b2b")
_ROW_ODD   = QColor("#323232")
_ACCENT    = QColor("#3B8ED0")


def _accent_btn(text, kind="add"):
    b = QPushButton(text)
    b.setProperty("accent", kind)
    return b


def _sep_frame():
    f = QFrame()
    f.setFrameShape(QFrame.HLine)
    f.setStyleSheet("background: #3a3a4a; max-height: 1px;")
    return f


def _section_label(text):
    lbl = QLabel(text)
    lbl.setObjectName("section-hdr")
    lbl.setFixedHeight(30)
    return lbl


# ── SubTableModel ─────────────────────────────────────────────────────────────

class SubTableModel(QAbstractTableModel):
    def __init__(self, df, cols_cfg, manager, sheet_full_name):
        super().__init__()
        self._df = df if df is not None else pd.DataFrame()
        self._cols_cfg = cols_cfg or {}
        self._manager = manager
        self._sheet = sheet_full_name

    # ── Qt interface ──────────────────────────────────────────────────────────

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self._df.columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        r, c = index.row(), index.column()
        col = self._df.columns[c]
        col_type = self._cols_cfg.get(col, {}).get("type", "string")
        val = self._df.iat[r, c]
        row_idx = self._df.index[r]

        if role == Qt.DisplayRole:
            return None if col_type == "bool" else (str(val) if val is not None else "")

        if role == Qt.CheckStateRole and col_type == "bool":
            v = val
            if isinstance(v, str):
                v = v.lower() in ("true", "1", "yes")
            return Qt.Checked if v else Qt.Unchecked

        if role == Qt.BackgroundRole:
            if (self._sheet, row_idx, col) in self._manager.dirty_cells:
                return QBrush(QColor(_DIRTY_BG))
            return QBrush(_ROW_EVEN if r % 2 == 0 else _ROW_ODD)

        if role == Qt.ForegroundRole:
            return QBrush(QColor("#dce4ee"))

        return None

    def setData(self, index, value, role=Qt.EditRole):
        if not index.isValid():
            return False
        r, c = index.row(), index.column()
        col = self._df.columns[c]
        row_idx = self._df.index[r]
        if role == Qt.CheckStateRole:
            value = (value == Qt.Checked)
        self._manager.update_cell(self._sheet, row_idx, col, value)
        # Refresh df reference
        self._df = self._manager.sub_tables.get(self._sheet, self._df)
        self.dataChanged.emit(index, index)
        return True

    def flags(self, index):
        if not index.isValid():
            return Qt.NoItemFlags
        col = self._df.columns[index.column()]
        col_type = self._cols_cfg.get(col, {}).get("type", "string")
        base = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if col_type == "bool":
            return base | Qt.ItemIsUserCheckable
        return base | Qt.ItemIsEditable

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section])
        return str(section + 1)

    # ── Public ────────────────────────────────────────────────────────────────

    def reload(self, df, cols_cfg=None):
        self.beginResetModel()
        self._df = df if df is not None else pd.DataFrame()
        if cols_cfg is not None:
            self._cols_cfg = cols_cfg
        self.endResetModel()

    def sort(self, column, order=Qt.AscendingOrder):
        """In-place sort so QTableView header clicks work without a proxy model."""
        if column < 0 or column >= len(self._df.columns):
            return
        col_name = self._df.columns[column]
        self.layoutAboutToBeChanged.emit()
        try:
            self._df = self._df.sort_values(
                col_name, ascending=(order == Qt.AscendingOrder), kind="mergesort"
            )
        except Exception:
            pass
        self.layoutChanged.emit()

    def df_index(self, view_row):
        if 0 <= view_row < len(self._df):
            return self._df.index[view_row]
        return None

    @property
    def df(self):
        return self._df


# ── EnumDelegate ──────────────────────────────────────────────────────────────

class EnumDelegate(QStyledItemDelegate):
    def __init__(self, options, parent=None):
        super().__init__(parent)
        self._options = [str(o) for o in options]

    def createEditor(self, parent, option, index):
        cb = QComboBox(parent)
        cb.addItems(self._options)
        return cb

    def setEditorData(self, editor, index):
        val = index.data(Qt.DisplayRole) or ""
        idx = editor.findText(val)
        if idx >= 0:
            editor.setCurrentIndex(idx)

    def setModelData(self, editor, model, index):
        model.setData(index, editor.currentText(), Qt.EditRole)

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)


# ── FieldEditorWidget ─────────────────────────────────────────────────────────

class FieldEditorWidget(QWidget):
    """Build-once field form; load_row() updates values only."""
    field_changed = Signal(str, object)  # col, value

    def __init__(self, parent=None):
        super().__init__(parent)
        self._widgets = {}       # col → QWidget
        self._col_types = {}     # col → type str
        self._base_styles = {}   # col → base stylesheet
        self._row_idx = None
        self._table_name = None
        self._manager = None
        self._built = False

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        self._content = QWidget()
        self._grid = QGridLayout(self._content)
        self._grid.setColumnMinimumWidth(0, 140)
        self._grid.setColumnStretch(1, 1)
        self._grid.setContentsMargins(6, 6, 10, 6)
        self._grid.setVerticalSpacing(2)
        self._grid.setHorizontalSpacing(8)
        scroll.setWidget(self._content)

        lo = QVBoxLayout(self)
        lo.setContentsMargins(0, 0, 0, 0)
        lo.addWidget(scroll)

    def build_for(self, df, cfg, table_name, manager):
        # Destroy old widgets
        while self._grid.count():
            item = self._grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self._widgets.clear()
        self._col_types.clear()
        self._base_styles.clear()
        self._table_name = table_name
        self._manager = manager

        cols_cfg = cfg.get("columns", {})

        for i, col in enumerate(df.columns):
            col_conf = cols_cfg.get(col, {})
            col_type = col_conf.get("type", "string")
            self._col_types[col] = col_type

            row_color = "#2b2b2b" if i % 2 == 0 else "#323232"

            lbl = QLabel(col)
            lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            lbl.setStyleSheet(
                f"color:#A0C4E8; font-weight:bold; background:{row_color};"
                f"padding:5px 4px; border-radius:2px;"
            )
            self._grid.addWidget(lbl, i, 0)

            base = f"background:{row_color};"

            if col_type == "bool":
                w = QCheckBox()
                w.setStyleSheet(f"background:{row_color}; padding:5px;")
                w.checkStateChanged.connect(lambda s, c=col:
                    self.field_changed.emit(c, s == Qt.Checked))

            elif col_type == "enum":
                opts = col_conf.get("options") or [""]
                w = QComboBox()
                w.addItems([str(o) for o in opts])
                w.setStyleSheet(base)
                w.currentTextChanged.connect(lambda v, c=col:
                    self.field_changed.emit(c, v))

            elif col_type in ("int", "float"):
                w = QLineEdit()
                w.setStyleSheet(base)
                w.textChanged.connect(lambda v, c=col, ct=col_type:
                    self._on_numeric(c, v, ct))

            else:  # string → QTextEdit
                w = QTextEdit()
                w.setMaximumHeight(72)
                w.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
                w.setStyleSheet(base)
                w.textChanged.connect(lambda c=col, widget=w:
                    self.field_changed.emit(c, widget.toPlainText()))

            self._base_styles[col] = base
            self._grid.addWidget(w, i, 1)
            self._widgets[col] = w

        self._grid.setRowStretch(len(df.columns), 1)
        self._built = True

    # ── Validation ────────────────────────────────────────────────────────────

    def _on_numeric(self, col, value, col_type):
        w = self._widgets.get(col)
        if w:
            valid = self._validate(value, col_type)
            w.setProperty("invalid", "true" if not valid else "false")
            w.style().unpolish(w)
            w.style().polish(w)
        self.field_changed.emit(col, value)

    @staticmethod
    def _validate(value, col_type):
        if col_type == "int":
            return value == "" or value.lstrip("-").isdigit()
        if col_type == "float":
            if value in ("", "-", "."): return True
            try: float(value); return True
            except ValueError: return False
        return True

    # ── Load row ──────────────────────────────────────────────────────────────

    def load_row(self, row_data, row_idx):
        if not self._built:
            return
        self._row_idx = row_idx
        dirty = self._manager.dirty_cells if self._manager else set()

        for col, w in self._widgets.items():
            w.blockSignals(True)
            try:
                try:
                    val = row_data[col]
                except (KeyError, TypeError):
                    val = ""

                col_type = self._col_types[col]
                is_dirty = (self._table_name, row_idx, col) in dirty
                base = self._base_styles.get(col, "")
                dirty_extra = f"background:{_DIRTY_BG};" if is_dirty else ""
                combined = base + dirty_extra

                if col_type == "bool":
                    v = val
                    if isinstance(v, str):
                        v = v.lower() in ("true", "1", "yes")
                    w.setChecked(bool(v) if val != "" else False)
                    w.setStyleSheet(combined)

                elif col_type == "enum":
                    w.setCurrentText(str(val) if val is not None else "")
                    w.setStyleSheet(combined)

                elif col_type in ("int", "float"):
                    w.setText(str(val) if val is not None else "")
                    # Reset invalid state
                    w.setProperty("invalid", "false")
                    w.style().unpolish(w)
                    w.style().polish(w)
                    if is_dirty:
                        w.setStyleSheet(combined)

                else:
                    w.setPlainText(str(val) if val is not None else "")
                    w.setStyleSheet(combined)

            finally:
                w.blockSignals(False)


# ── SubTablePanel ─────────────────────────────────────────────────────────────

class SubTablePanel(QWidget):
    """QTableView with sort, enum delegate, keyboard delete, context menu."""
    row_deleted = Signal(str, object)  # sheet_full_name, df_index

    def __init__(self, sheet_full_name, cols_cfg, manager, parent=None):
        super().__init__(parent)
        self._sheet   = sheet_full_name
        self._manager = manager

        # Start with empty df — data injected via reload()
        empty_df = pd.DataFrame(columns=list(
            (manager.sub_tables.get(sheet_full_name) or pd.DataFrame()).columns
        ))
        self._model = SubTableModel(empty_df, cols_cfg, manager, sheet_full_name)

        self._view = QTableView()
        self._view.setModel(self._model)          # direct, no proxy
        self._view.setSortingEnabled(True)        # uses SubTableModel.sort()
        self._view.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._view.setSelectionMode(QAbstractItemView.SingleSelection)
        self._view.setEditTriggers(
            QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked)
        self._view.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self._view.horizontalHeader().setStretchLastSection(True)
        self._view.verticalHeader().setDefaultSectionSize(26)
        self._view.verticalHeader().hide()
        self._view.setContextMenuPolicy(Qt.CustomContextMenu)
        self._view.customContextMenuRequested.connect(self._ctx_menu)
        self._view.keyPressEvent = self._key_press

        self._refresh_delegates(cols_cfg)

        lo = QVBoxLayout(self)
        lo.setContentsMargins(0, 0, 0, 0)
        lo.addWidget(self._view)

    def _refresh_delegates(self, cols_cfg):
        for c, col in enumerate(self._model.df.columns):
            col_type = (cols_cfg or {}).get(col, {}).get("type", "string")
            if col_type == "enum":
                opts = (cols_cfg or {}).get(col, {}).get("options") or [""]
                self._view.setItemDelegateForColumn(c, EnumDelegate(opts, self._view))

    def reload(self, df, cols_cfg=None):
        self._model.reload(df, cols_cfg)
        if cols_cfg:
            self._refresh_delegates(cols_cfg)
        if not df.empty:
            self._view.resizeColumnsToContents()

    def _ctx_menu(self, pos):
        idx = self._view.indexAt(pos)
        if not idx.isValid():
            return
        menu = QMenu(self)
        menu.addAction("刪除此列", self._delete_selected)
        menu.exec(self._view.viewport().mapToGlobal(pos))

    def _key_press(self, event):
        if event.key() == Qt.Key_Delete:
            self._delete_selected()
        else:
            QTableView.keyPressEvent(self._view, event)

    def _delete_selected(self):
        sel = self._view.selectionModel().selectedRows()
        if not sel:
            return
        df_idx = self._model.df_index(sel[0].row())
        if df_idx is not None:
            self.row_deleted.emit(self._sheet, df_idx)

    def selected_df_index(self):
        sel = self._view.selectionModel().selectedRows()
        if not sel:
            return None
        return self._model.df_index(sel[0].row())


# ── TableEditor ───────────────────────────────────────────────────────────────

class TableEditor(QWidget):
    status_message = Signal(str, str)  # text, color

    def __init__(self, table_name, manager, parent=None):
        super().__init__(parent)
        self.table_name = table_name
        self.manager = manager
        self.df = manager.tables[table_name]
        self.cfg = manager.config.get(table_name, {})
        cols = list(self.df.columns)
        self.cls_key = self.cfg.get("classification_key", cols[0] if cols else "")
        self.pk_key  = self.cfg.get("primary_key",        cols[0] if cols else "")

        self.current_cls_val    = None
        self.current_master_idx = None
        self.current_master_pk  = None

        self._field_panel: FieldEditorWidget | None = None
        self._sub_panels: dict[str, SubTablePanel] = {}  # tab_name → panel

        self._setup_ui()
        self._load_cls_list()

    # ── Layout ────────────────────────────────────────────────────────────────

    def _setup_ui(self):
        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(3)
        root.addWidget(splitter)

        # ── Left: classification ──────────────────────────────────────────────
        left = QWidget()
        left.setMinimumWidth(130)
        left.setMaximumWidth(200)
        lv = QVBoxLayout(left)
        lv.setContentsMargins(0, 0, 0, 0)
        lv.setSpacing(0)
        lv.addWidget(_section_label(f"▣  分類 ({self.cls_key})"))
        self._cls_list = QListWidget()
        self._cls_list.currentItemChanged.connect(self._on_cls_changed)
        lv.addWidget(self._cls_list, 1)
        lv.addWidget(_sep_frame())

        cls_btns = QHBoxLayout()
        cls_btns.setContentsMargins(4, 4, 4, 4)
        b_add_cls = _accent_btn("+", "add");   b_add_cls.setFixedWidth(28)
        b_del_cls = _accent_btn("−", "del");   b_del_cls.setFixedWidth(28)
        b_up_cls  = QPushButton("▲");           b_up_cls.setFixedWidth(28)
        b_dn_cls  = QPushButton("▼");           b_dn_cls.setFixedWidth(28)
        b_add_cls.clicked.connect(self.add_classification)
        b_del_cls.clicked.connect(self.delete_classification)
        b_up_cls.clicked.connect(lambda: self.move_classification(-1))
        b_dn_cls.clicked.connect(lambda: self.move_classification(1))
        for b in [b_add_cls, b_del_cls, b_up_cls, b_dn_cls]:
            cls_btns.addWidget(b)
        lv.addLayout(cls_btns)
        splitter.addWidget(left)

        # ── Mid: item list ────────────────────────────────────────────────────
        mid = QWidget()
        mid.setMinimumWidth(160)
        mid.setMaximumWidth(260)
        mv = QVBoxLayout(mid)
        mv.setContentsMargins(0, 0, 0, 0)
        mv.setSpacing(0)
        mv.addWidget(_section_label("☰  項目清單"))

        self._filter_edit = QLineEdit()
        self._filter_edit.setPlaceholderText("搜尋…")
        self._filter_edit.setClearButtonEnabled(True)
        self._filter_edit.setContentsMargins(4, 2, 4, 2)
        self._filter_edit.textChanged.connect(self._apply_filter)
        mv.addWidget(self._filter_edit)

        self._item_list = QListWidget()
        self._item_list.currentItemChanged.connect(self._on_item_changed)
        mv.addWidget(self._item_list, 1)
        mv.addWidget(_sep_frame())

        item_btns1 = QHBoxLayout()
        item_btns1.setContentsMargins(4, 4, 4, 2)
        b_add  = _accent_btn("+ 新增", "add")
        b_copy = QPushButton("複製")
        b_add.clicked.connect(self.add_master_item)
        b_copy.clicked.connect(self.copy_master_item)
        item_btns1.addWidget(b_add, 1)
        item_btns1.addWidget(b_copy)
        mv.addLayout(item_btns1)

        item_btns2 = QHBoxLayout()
        item_btns2.setContentsMargins(4, 2, 4, 4)
        b_up   = QPushButton("▲");    b_up.setFixedWidth(32)
        b_dn   = QPushButton("▼");    b_dn.setFixedWidth(32)
        b_del  = _accent_btn("刪除", "del")
        b_up.clicked.connect(lambda: self.move_master_item(-1))
        b_dn.clicked.connect(lambda: self.move_master_item(1))
        b_del.clicked.connect(self.delete_master_item)
        item_btns2.addWidget(b_up)
        item_btns2.addWidget(b_dn)
        item_btns2.addStretch(1)
        item_btns2.addWidget(b_del)
        mv.addLayout(item_btns2)
        splitter.addWidget(mid)

        # ── Right: editor + sub-tables ────────────────────────────────────────
        right_splitter = QSplitter(Qt.Vertical)
        right_splitter.setHandleWidth(4)

        # Field editor area
        field_area = QWidget()
        field_area.setMinimumHeight(120)
        fv = QVBoxLayout(field_area)
        fv.setContentsMargins(0, 0, 0, 0)
        fv.setSpacing(0)
        fv.addWidget(_section_label("◉  欄位資料"))
        self._field_container = QWidget()
        fv.addWidget(self._field_container, 1)
        fl = QVBoxLayout(self._field_container)
        fl.setContentsMargins(0, 0, 0, 0)

        right_splitter.addWidget(field_area)

        # Sub-tables
        sub_area = QWidget()
        sub_area.setMinimumHeight(100)
        sv = QVBoxLayout(sub_area)
        sv.setContentsMargins(0, 0, 0, 0)
        sv.setSpacing(0)
        sub_hdr = QWidget()
        sub_hdr.setFixedHeight(30)
        sub_hdr.setStyleSheet("background:#2a2a3a;")
        sh = QHBoxLayout(sub_hdr)
        sh.setContentsMargins(6, 0, 4, 0)
        lbl_sub = QLabel("◦  子資料表")
        lbl_sub.setStyleSheet("color:#A0C4E8; font-weight:bold;")
        b_add_sub  = _accent_btn("+ 新增", "add"); b_add_sub.setFixedHeight(22)
        b_del_sub  = _accent_btn("刪除",   "del"); b_del_sub.setFixedHeight(22)
        b_up_sub   = QPushButton("▲"); b_up_sub.setFixedSize(26, 22)
        b_dn_sub   = QPushButton("▼"); b_dn_sub.setFixedSize(26, 22)
        b_cp_sub   = QPushButton("複製"); b_cp_sub.setFixedHeight(22)
        b_add_sub.clicked.connect(self.add_sub_item)
        b_del_sub.clicked.connect(self.delete_sub_item)
        b_up_sub.clicked.connect(lambda: self.move_sub_item(-1))
        b_dn_sub.clicked.connect(lambda: self.move_sub_item(1))
        b_cp_sub.clicked.connect(self.copy_sub_item)
        sh.addWidget(lbl_sub)
        sh.addStretch(1)
        for b in [b_cp_sub, b_up_sub, b_dn_sub, b_add_sub, b_del_sub]:
            sh.addWidget(b)
        sv.addWidget(sub_hdr)
        sv.addWidget(_sep_frame())

        self._sub_tabs = QTabWidget()
        self._sub_tabs.setDocumentMode(True)
        sv.addWidget(self._sub_tabs, 1)

        right_splitter.addWidget(sub_area)
        right_splitter.setSizes([300, 200])
        splitter.addWidget(right_splitter)

        splitter.setSizes([150, 200, 600])
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 0)
        splitter.setStretchFactor(2, 1)

        # Context menus
        self._cls_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self._cls_list.customContextMenuRequested.connect(self._cls_ctx_menu)
        self._item_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self._item_list.customContextMenuRequested.connect(self._item_ctx_menu)

        # Build sub-table tabs (empty placeholders)
        self._build_sub_tabs()

    # ── Classification ────────────────────────────────────────────────────────

    def _load_cls_list(self):
        self._cls_list.blockSignals(True)
        self._cls_list.clear()
        if self.cls_key not in self.df.columns:
            self._cls_list.blockSignals(False)
            return
        groups = self.df[self.cls_key].unique()
        for g in groups:
            count = int((self.df[self.cls_key] == g).sum())
            item = QListWidgetItem(f"{g}  ({count})")
            item.setData(Qt.UserRole, g)
            self._cls_list.addItem(item)
            if str(g) == str(self.current_cls_val):
                self._cls_list.setCurrentItem(item)
        self._cls_list.blockSignals(False)

    def _on_cls_changed(self, cur, _prev):
        if cur is None:
            return
        g = cur.data(Qt.UserRole)
        self.current_cls_val = g
        self._load_item_list()

    def add_classification(self):
        name, ok = QInputDialog.getText(self, "新增分類", "分類名稱:")
        if not ok or not name.strip():
            return
        name = name.strip()
        if name in self.df[self.cls_key].values:
            QMessageBox.warning(self, "錯誤", "此分類已存在")
            return
        new_row = {col: "" for col in self.df.columns}
        new_row[self.cls_key] = name
        if self.pk_key in self.df.columns:
            new_id, ok2 = QInputDialog.getText(self, "新增分類", f"首筆資料的 {self.pk_key}:")
            if not ok2 or not new_id.strip():
                return
            new_row[self.pk_key] = new_id.strip()
        self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
        self.manager.tables[self.table_name] = self.df
        self.manager.dirty = True
        self._reload_all(select_cls=name)

    def delete_classification(self):
        g = self.current_cls_val
        if g is None:
            return
        if QMessageBox.question(self, "確認刪除", f"刪除分類 [{g}] 及其所有資料？",
                                QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return
        self.df = self.df[self.df[self.cls_key] != g].reset_index(drop=True)
        self.manager.tables[self.table_name] = self.df
        self.manager.dirty = True
        self.current_cls_val = None
        self.current_master_idx = None
        self._reload_all()

    def move_classification(self, delta):
        if self.current_cls_val is None:
            return
        groups = list(self.df[self.cls_key].unique())
        try:
            pos = [str(g) for g in groups].index(str(self.current_cls_val))
        except ValueError:
            return
        new_pos = pos + delta
        if new_pos < 0 or new_pos >= len(groups):
            return
        groups[pos], groups[new_pos] = groups[new_pos], groups[pos]
        self.df = pd.concat(
            [self.df[self.df[self.cls_key] == g] for g in groups],
            ignore_index=True
        )
        self.manager.tables[self.table_name] = self.df
        self.manager.dirty = True
        self._reload_all(select_cls=self.current_cls_val)

    # ── Item list ─────────────────────────────────────────────────────────────

    def _load_item_list(self):
        query = self._filter_edit.text().strip().lower()
        self._item_list.blockSignals(True)
        self._item_list.clear()
        if self.current_cls_val is None or self.pk_key not in self.df.columns:
            self._item_list.blockSignals(False)
            return
        sub_df = self.df[self.df[self.cls_key] == self.current_cls_val]
        for df_idx, row in sub_df.iterrows():
            pk_val = str(row[self.pk_key])
            if query and query not in pk_val.lower():
                continue
            item = QListWidgetItem(pk_val)
            item.setData(Qt.UserRole, df_idx)
            self._item_list.addItem(item)
            if df_idx == self.current_master_idx:
                self._item_list.setCurrentItem(item)
        self._item_list.blockSignals(False)

    def _apply_filter(self, _text=""):
        self._load_item_list()

    def _on_item_changed(self, cur, _prev):
        if cur is None:
            return
        df_idx = cur.data(Qt.UserRole)
        self._load_editor(df_idx)

    def add_master_item(self):
        if self.current_cls_val is None:
            QMessageBox.warning(self, "提示", "請先選擇分類")
            return
        new_id, ok = QInputDialog.getText(self, "新增項目", f"{self.pk_key}:")
        if not ok or not new_id.strip():
            return
        new_id = new_id.strip()
        if new_id in self.df[self.pk_key].astype(str).values:
            QMessageBox.warning(self, "錯誤", "此 ID 已存在")
            return
        new_row = {col: "" for col in self.df.columns}
        new_row[self.cls_key] = self.current_cls_val
        new_row[self.pk_key]  = new_id
        cls_rows = self.df[self.df[self.cls_key] == self.current_cls_val]
        insert_after = cls_rows.index.max() + 1 if not cls_rows.empty else len(self.df)
        top, bot = self.df.iloc[:insert_after], self.df.iloc[insert_after:]
        self.df = pd.concat([top, pd.DataFrame([new_row]), bot], ignore_index=True)
        self.manager.tables[self.table_name] = self.df
        self.manager.dirty = True
        new_df_idx = self.df[self.df[self.pk_key].astype(str) == str(new_id)].index[0]
        self._reload_all(select_cls=self.current_cls_val, select_idx=new_df_idx)

    def copy_master_item(self):
        if self.current_master_idx is None:
            return
        new_id, ok = QInputDialog.getText(self, "複製項目", f"新的 {self.pk_key}:")
        if not ok or not new_id.strip():
            return
        new_id = new_id.strip()
        if new_id in self.df[self.pk_key].astype(str).values:
            QMessageBox.warning(self, "錯誤", "此 ID 已存在")
            return
        new_row = self.df.loc[self.current_master_idx].copy()
        new_row[self.pk_key] = new_id
        self.df = pd.concat([self.df, pd.DataFrame([new_row])], ignore_index=True)
        self.manager.tables[self.table_name] = self.df
        self.manager.dirty = True
        new_df_idx = self.df[self.df[self.pk_key].astype(str) == str(new_id)].index[0]
        self._reload_all(select_cls=self.current_cls_val, select_idx=new_df_idx)

    def delete_master_item(self):
        if self.current_master_idx is None:
            return
        pk = str(self.df.at[self.current_master_idx, self.pk_key])
        self.df.drop(self.current_master_idx, inplace=True)
        self.df.reset_index(drop=True, inplace=True)
        self.manager.tables[self.table_name] = self.df
        self.manager.dirty = True
        self.current_master_idx = None
        self._reload_all(select_cls=self.current_cls_val)
        self.status_message.emit(f"已刪除 {pk}", "#e09040")

    def move_master_item(self, delta):
        if self.current_master_idx is None or self.current_cls_val is None:
            return
        cls_idxs = list(self.df[self.df[self.cls_key] == self.current_cls_val].index)
        try:
            pos = cls_idxs.index(self.current_master_idx)
        except ValueError:
            return
        new_pos = pos + delta
        if new_pos < 0 or new_pos >= len(cls_idxs):
            return
        cls_idxs[pos], cls_idxs[new_pos] = cls_idxs[new_pos], cls_idxs[pos]
        other_idxs = [i for i in self.df.index if i not in cls_idxs]
        # Re-assemble: items before, cls items reordered, items after
        all_cls = list(self.df[self.df[self.cls_key] == self.current_cls_val].index)
        first_cls = all_cls[0]
        before = [i for i in other_idxs if i < first_cls]
        after  = [i for i in other_idxs if i > max(all_cls)]
        new_order = before + cls_idxs + after
        self.df = self.df.loc[new_order].reset_index(drop=True)
        self.manager.tables[self.table_name] = self.df
        self.manager.dirty = True
        new_idx = self.df[self.df[self.pk_key].astype(str) == str(self.current_master_pk)].index
        sel_idx = new_idx[0] if not new_idx.empty else None
        self._reload_all(select_cls=self.current_cls_val, select_idx=sel_idx)

    # ── Field editor ──────────────────────────────────────────────────────────

    def _load_editor(self, df_idx):
        self.current_master_idx = df_idx
        if df_idx not in self.df.index:
            return
        row_data = self.df.loc[df_idx]
        self.current_master_pk = row_data[self.pk_key]

        if self._field_panel is None:
            self._field_panel = FieldEditorWidget(self._field_container)
            self._field_panel.field_changed.connect(self._on_field_change)
            self._field_container.layout().addWidget(self._field_panel)
            self._field_panel.build_for(self.df, self.cfg, self.table_name, self.manager)

        self._field_panel.load_row(row_data, df_idx)
        self._refresh_sub_tables()

    def _on_field_change(self, col, value):
        if self.current_master_idx is None:
            return
        self.manager.update_cell(self.table_name, self.current_master_idx, col, value)

    # ── Sub-tables ────────────────────────────────────────────────────────────

    def _build_sub_tabs(self):
        """Register sub-tables with placeholder tabs. Panels created lazily in _refresh_sub_tables."""
        self._sub_tabs.clear()
        self._sub_panels.clear()
        self._sub_tab_order = []   # ordered list of tab_names
        prefix = self.table_name + "."
        for key in self.manager.sub_tables:
            if not key.startswith(prefix):
                continue
            tab_name = key[len(prefix):]
            self._sub_tab_order.append(tab_name)
            self._sub_tabs.addTab(QWidget(), tab_name)   # cheap placeholder

    def _refresh_sub_tables(self):
        if self.current_master_pk is None:
            return
        prefix = self.table_name + "."
        for i, tab_name in enumerate(self._sub_tab_order):
            full = prefix + tab_name
            sub_df = self.manager.sub_tables.get(full)
            if sub_df is None:
                continue
            sub_cfg  = self.cfg.get("sub_tables", {}).get(tab_name, {})
            fk_key   = sub_cfg.get("foreign_key", self.pk_key)
            cols_cfg = sub_cfg.get("columns", {})
            filtered = sub_df[sub_df[fk_key].astype(str) == str(self.current_master_pk)]

            if tab_name not in self._sub_panels:
                # First time: create panel, replace placeholder tab
                panel = SubTablePanel(full, cols_cfg, self.manager)
                panel.row_deleted.connect(self._on_sub_delete)
                self._sub_panels[tab_name] = panel
                self._sub_tabs.removeTab(i)
                self._sub_tabs.insertTab(i, panel, tab_name)

            self._sub_panels[tab_name].reload(filtered, cols_cfg)

    def _on_sub_delete(self, sheet_full, df_idx):
        sub_df = self.manager.sub_tables.get(sheet_full)
        if sub_df is None or df_idx not in sub_df.index:
            return
        sub_df.drop(df_idx, inplace=True)
        sub_df.reset_index(drop=True, inplace=True)
        self.manager.sub_tables[sheet_full] = sub_df
        self.manager.dirty = True
        self._refresh_sub_tables()
        self.status_message.emit("已刪除子表列", "#e09040")

    def _current_sub_panel(self) -> SubTablePanel | None:
        idx = self._sub_tabs.currentIndex()
        if idx < 0:
            return None
        tab_name = self._sub_tabs.tabText(idx)
        return self._sub_panels.get(tab_name)

    def add_sub_item(self):
        if self.current_master_pk is None:
            return
        panel = self._current_sub_panel()
        if panel is None:
            return
        tab_name = self._sub_tabs.tabText(self._sub_tabs.currentIndex())
        full = self.table_name + "." + tab_name
        sub_df = self.manager.sub_tables.get(full)
        if sub_df is None:
            return
        sub_cfg = self.cfg.get("sub_tables", {}).get(tab_name, {})
        fk_key = sub_cfg.get("foreign_key", self.pk_key)
        new_row = {col: "" for col in sub_df.columns}
        new_row[fk_key] = self.current_master_pk
        siblings = sub_df[sub_df[fk_key].astype(str) == str(self.current_master_pk)]
        insert_at = siblings.index.max() + 1 if not siblings.empty else len(sub_df)
        top, bot = sub_df.iloc[:insert_at], sub_df.iloc[insert_at:]
        sub_df = pd.concat([top, pd.DataFrame([new_row]), bot], ignore_index=True)
        self.manager.sub_tables[full] = sub_df
        self.manager.dirty = True
        self._refresh_sub_tables()

    def delete_sub_item(self):
        panel = self._current_sub_panel()
        if panel is None:
            return
        tab_name = self._sub_tabs.tabText(self._sub_tabs.currentIndex())
        full = self.table_name + "." + tab_name
        df_idx = panel.selected_df_index()
        if df_idx is None:
            return
        self._on_sub_delete(full, df_idx)

    def move_sub_item(self, delta):
        panel = self._current_sub_panel()
        if panel is None:
            return
        tab_name = self._sub_tabs.tabText(self._sub_tabs.currentIndex())
        full = self.table_name + "." + tab_name
        sub_df = self.manager.sub_tables.get(full)
        if sub_df is None:
            return
        df_idx = panel.selected_df_index()
        if df_idx is None:
            return
        sub_cfg = self.cfg.get("sub_tables", {}).get(tab_name, {})
        fk_key = sub_cfg.get("foreign_key", self.pk_key)
        siblings = list(sub_df[sub_df[fk_key].astype(str) == str(self.current_master_pk)].index)
        try:
            pos = siblings.index(df_idx)
        except ValueError:
            return
        new_pos = pos + delta
        if new_pos < 0 or new_pos >= len(siblings):
            return
        siblings[pos], siblings[new_pos] = siblings[new_pos], siblings[pos]
        others = [i for i in sub_df.index if i not in siblings]
        first = siblings[0] if siblings else 0
        before = [i for i in others if i < first]
        after  = [i for i in others if i > max(siblings)]
        new_order = before + siblings + after
        sub_df = sub_df.loc[new_order].reset_index(drop=True)
        self.manager.sub_tables[full] = sub_df
        self.manager.dirty = True
        self._refresh_sub_tables()

    def copy_sub_item(self):
        panel = self._current_sub_panel()
        if panel is None:
            return
        tab_name = self._sub_tabs.tabText(self._sub_tabs.currentIndex())
        full = self.table_name + "." + tab_name
        sub_df = self.manager.sub_tables.get(full)
        if sub_df is None:
            return
        df_idx = panel.selected_df_index()
        if df_idx is None:
            return
        new_row = sub_df.loc[df_idx].copy()
        sub_df = pd.concat([sub_df, pd.DataFrame([new_row])], ignore_index=True)
        self.manager.sub_tables[full] = sub_df
        self.manager.dirty = True
        self._refresh_sub_tables()

    # ── Context menus ─────────────────────────────────────────────────────────

    def _cls_ctx_menu(self, pos):
        item = self._cls_list.itemAt(pos)
        if item is None:
            return
        g = item.data(Qt.UserRole)
        menu = QMenu(self)
        menu.addAction("重新命名", lambda: self._rename_cls(g))
        menu.addAction("刪除此分類", self.delete_classification)
        menu.exec(self._cls_list.mapToGlobal(pos))

    def _rename_cls(self, old_val):
        new_val, ok = QInputDialog.getText(self, "重新命名", "新名稱:", text=str(old_val))
        if not ok or not new_val.strip() or new_val.strip() == str(old_val):
            return
        self.df[self.cls_key] = self.df[self.cls_key].replace(old_val, new_val.strip())
        self.manager.tables[self.table_name] = self.df
        self.manager.dirty = True
        if str(self.current_cls_val) == str(old_val):
            self.current_cls_val = new_val.strip()
        self._reload_all(select_cls=self.current_cls_val)

    def _item_ctx_menu(self, pos):
        item = self._item_list.itemAt(pos)
        if item is None:
            return
        menu = QMenu(self)
        menu.addAction("複製",    self.copy_master_item)
        menu.addAction("刪除",    self.delete_master_item)
        menu.exec(self._item_list.mapToGlobal(pos))

    # ── Refresh helpers ───────────────────────────────────────────────────────

    def _reload_all(self, select_cls=None, select_idx=None):
        """Reload df from manager, rebuild classification list, item list."""
        self.df = self.manager.tables[self.table_name]
        if select_cls is not None:
            self.current_cls_val = select_cls
        if select_idx is not None:
            self.current_master_idx = select_idx
        self._load_cls_list()
        if self.current_cls_val is not None:
            self._load_item_list()
        if self.current_master_idx is not None and self.current_master_idx in self.df.index:
            self._load_editor(self.current_master_idx)

    def reload_after_config(self):
        """Called after config changes to rebuild everything."""
        self.cfg     = self.manager.config.get(self.table_name, {})
        self.cls_key = self.cfg.get("classification_key",
                       list(self.df.columns)[0] if len(self.df.columns) > 0 else "")
        self.pk_key  = self.cfg.get("primary_key",
                       list(self.df.columns)[0] if len(self.df.columns) > 0 else "")
        # Destroy field panel so it rebuilds
        if self._field_panel:
            self._field_panel.deleteLater()
            self._field_panel = None
        self._build_sub_tabs()
        self._reload_all()


# ── WelcomeWidget ─────────────────────────────────────────────────────────────

class WelcomeWidget(QWidget):
    open_file = Signal()
    new_file  = Signal()
    open_recent = Signal(str)

    def __init__(self, manager: JsonDataManager, parent=None):
        super().__init__(parent)
        self._manager = manager
        self._setup_ui()

    def _setup_ui(self):
        outer = QVBoxLayout(self)
        outer.setAlignment(Qt.AlignCenter)

        card = QWidget()
        card.setFixedWidth(420)
        card.setStyleSheet("background:#1a1a2e; border-radius:8px;")
        lo = QVBoxLayout(card)
        lo.setContentsMargins(36, 32, 36, 32)
        lo.setSpacing(8)

        logo = QLabel("{ }")
        logo.setAlignment(Qt.AlignCenter)
        logo.setStyleSheet("font-family:Consolas; font-size:40px; font-weight:bold; color:#7ec8e3; background:transparent;")
        lo.addWidget(logo)

        title = QLabel("JsonEditor")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size:18px; font-weight:bold; color:#c8ccd4; background:transparent;")
        lo.addWidget(title)

        subtitle = QLabel("輕量 JSON 資料編輯器")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color:#44446a; background:transparent;")
        lo.addWidget(subtitle)

        lo.addSpacing(20)

        btn_row = QHBoxLayout()
        b_open = _accent_btn("📂  開啟 JSON", "blue")
        b_new  = QPushButton("📄  新建 JSON")
        b_open.setFixedHeight(36)
        b_new.setFixedHeight(36)
        b_open.clicked.connect(self.open_file)
        b_new.clicked.connect(self.new_file)
        btn_row.addWidget(b_open)
        btn_row.addWidget(b_new)
        lo.addLayout(btn_row)

        recent = self._manager._recent_files
        if recent:
            lo.addSpacing(16)
            sep = QFrame()
            sep.setFrameShape(QFrame.HLine)
            sep.setStyleSheet("background:#2a2a4a; max-height:1px;")
            lo.addWidget(sep)
            lo.addSpacing(4)

            hdr = QLabel("最近開啟")
            hdr.setStyleSheet("color:#44446a; background:transparent;")
            lo.addWidget(hdr)

            for path in recent[:6]:
                fname = os.path.basename(path)
                dirn  = os.path.dirname(path)
                row   = QWidget()
                row.setStyleSheet(
                    "QWidget { background:transparent; border-radius:3px; }"
                    "QWidget:hover { background:#1e2040; }"
                )
                row.setCursor(Qt.PointingHandCursor)
                rlo = QHBoxLayout(row)
                rlo.setContentsMargins(4, 4, 4, 4)
                lbl_fn = QLabel(fname)
                lbl_fn.setStyleSheet("color:#8888cc; font-weight:bold; background:transparent;")
                lbl_dir = QLabel(dirn)
                lbl_dir.setStyleSheet("color:#333355; background:transparent;")
                lbl_dir.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
                rlo.addWidget(lbl_fn)
                rlo.addWidget(lbl_dir, 1)
                row.mousePressEvent = lambda e, p=path: self.open_recent.emit(p)
                lo.addWidget(row)

        outer.addWidget(card)

    def refresh(self, manager):
        self._manager = manager
        # Rebuild UI
        while self.layout().count():
            item = self.layout().takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self._setup_ui()


# ── App ───────────────────────────────────────────────────────────────────────

class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("JsonEditor")
        self.resize(1200, 750)
        self.manager = JsonDataManager()
        self._editors: dict[str, TableEditor | None] = {}
        self._snackbar_timer = QTimer(self)
        self._snackbar_timer.setSingleShot(True)
        self._snackbar_timer.timeout.connect(lambda: self._status_lbl.setText("就緒"))

        self._setup_toolbar()
        self._setup_statusbar()
        self._setup_content()

        # Keyboard shortcuts
        QAction(parent=self, shortcut=QKeySequence("Ctrl+S"), triggered=self.save_file).setEnabled(True)
        self.addAction(QAction(parent=self, shortcut=QKeySequence("Ctrl+S"), triggered=self.save_file))
        self.addAction(QAction(parent=self, shortcut=QKeySequence("Ctrl+O"), triggered=self.load_file))

        self._show_welcome()

    # ── Toolbar ───────────────────────────────────────────────────────────────

    def _setup_toolbar(self):
        tb = QToolBar("main", self)
        tb.setMovable(False)
        tb.setIconSize(QSize(18, 18))
        self.addToolBar(tb)

        logo = QLabel("{ } JsonEditor")
        logo.setStyleSheet("font-family:Consolas; font-size:13px; font-weight:bold; color:#7ec8e3; padding:0 14px;")
        tb.addWidget(logo)
        tb.addSeparator()

        def _act(icon, text, shortcut, slot):
            a = QAction(f"{icon}  {text}", self)
            if shortcut:
                a.setShortcut(QKeySequence(shortcut))
            a.triggered.connect(slot)
            tb.addAction(a)
            return a

        _act("📂", "開啟", "Ctrl+O", self.load_file)
        self._act_save = _act("💾", "儲存", "Ctrl+S", self.save_file)
        tb.addSeparator()
        _act("🔍", "搜尋", "Ctrl+F", self._show_search)
        _act("⚙",  "配置", "",       self.open_config)
        tb.addSeparator()
        _act("🕓", "最近", "", self._show_recent_menu)

    # ── Status bar ────────────────────────────────────────────────────────────

    def _setup_statusbar(self):
        sb = self.statusBar()
        self._status_lbl = QLabel("就緒")
        self._dirty_lbl  = QLabel("")
        self._path_lbl   = QLabel("")
        self._path_lbl.setStyleSheet("color:#333355;")
        sb.addWidget(self._status_lbl, 1)
        sb.addPermanentWidget(self._path_lbl)
        sb.addPermanentWidget(self._dirty_lbl)

    def show_snackbar(self, text, duration_ms=3000, color="#7ec8e3"):
        self._status_lbl.setText(text)
        self._status_lbl.setStyleSheet(f"color:{color};")
        self._snackbar_timer.start(duration_ms)

    def _on_snackbar_done(self):
        self._status_lbl.setText("就緒")
        self._status_lbl.setStyleSheet("")

    # ── Content ───────────────────────────────────────────────────────────────

    def _setup_content(self):
        self._stack = QStackedWidget(self)
        self.setCentralWidget(self._stack)

        self._welcome = WelcomeWidget(self.manager)
        self._welcome.open_file.connect(self.load_file)
        self._welcome.new_file.connect(self._new_file)
        self._welcome.open_recent.connect(self._load_recent)
        self._stack.addWidget(self._welcome)

        self._tab_widget = QTabWidget()
        self._tab_widget.setDocumentMode(True)
        self._tab_widget.currentChanged.connect(self._on_tab_changed)
        self._stack.addWidget(self._tab_widget)

    def _show_welcome(self):
        self._stack.setCurrentWidget(self._welcome)

    def _show_editor(self):
        self._stack.setCurrentWidget(self._tab_widget)

    # ── File I/O ──────────────────────────────────────────────────────────────

    def load_file(self):
        last_dir = os.path.dirname(self.manager._full_config.get("_last_file", "")) or ""
        path, _ = QFileDialog.getOpenFileName(
            self, "開啟 JSON", last_dir, "JSON 檔案 (*.json);;所有檔案 (*.*)")
        if path:
            self._load_path(path)

    def _new_file(self):
        import json
        path, _ = QFileDialog.getSaveFileName(
            self, "新建 JSON", "", "JSON 檔案 (*.json)")
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump([], f, indent=4, ensure_ascii=False)
        except Exception as e:
            QMessageBox.critical(self, "錯誤", str(e))
            return
        self._load_path(path)

    def _load_recent(self, path):
        if not os.path.exists(path):
            QMessageBox.warning(self, "錯誤", f"找不到檔案：\n{path}")
            self.manager._recent_files = [p for p in self.manager._recent_files if p != path]
            self.manager.save_config()
            return
        self._load_path(path)

    def _load_path(self, path):
        import sys
        self._set_loading(True, f"載入 {os.path.basename(path)}…")
        orig_interval = sys.getswitchinterval()
        sys.setswitchinterval(0.001)   # 1ms GIL switch → main thread stays alive

        worker = _LoadWorker(self.manager, path)
        self._active_worker = worker

        def _on_done():
            sys.setswitchinterval(orig_interval)
            self._active_worker = None
            self._set_loading(False)
            self._refresh_ui()
            # Remember last file
            self.manager._full_config["_last_file"] = path
            self.manager.save_config()

        def _on_error(msg):
            sys.setswitchinterval(orig_interval)
            self._active_worker = None
            self._set_loading(False)
            QMessageBox.critical(self, "載入失敗", msg)

        worker.done.connect(_on_done)
        worker.error.connect(_on_error)
        worker.start()

    def _set_loading(self, loading: bool, msg: str = "就緒"):
        self.setEnabled(not loading)
        self._status_lbl.setText(msg)
        self._status_lbl.setStyleSheet("color:#7ec8e3;" if loading else "")
        QApplication.processEvents()   # flush UI before heavy work begins

    def save_file(self):
        if not self.manager.json_path:
            QMessageBox.warning(self, "提示", "尚未載入任何 JSON 檔案")
            return

        import sys
        self._set_loading(True, "儲存中…")
        orig_interval = sys.getswitchinterval()
        sys.setswitchinterval(0.001)

        worker = _SaveWorker(self.manager)
        self._active_worker = worker

        def _on_done():
            sys.setswitchinterval(orig_interval)
            self._active_worker = None
            self._set_loading(False)
            self.show_snackbar("✓ 已儲存", color="#4ec87a")
            self._update_title()

        def _on_error(msg):
            sys.setswitchinterval(orig_interval)
            self._active_worker = None
            self._set_loading(False)
            QMessageBox.critical(self, "存檔失敗", msg)

        worker.done.connect(_on_done)
        worker.error.connect(_on_error)
        worker.start()

    def _show_recent_menu(self):
        recent = self.manager._recent_files
        if not recent:
            self.show_snackbar("沒有最近開啟的檔案")
            return
        menu = QMenu(self)
        for path in recent:
            fname = os.path.basename(path)
            menu.addAction(f"{fname}  —  {os.path.dirname(path)}",
                           lambda p=path: self._load_recent(p))
        menu.addSeparator()
        menu.addAction("清除記錄", self._clear_recent)
        menu.exec(self.mapToGlobal(self.rect().topLeft()))

    def _clear_recent(self):
        self.manager._recent_files = []
        self.manager.save_config()
        self.show_snackbar("已清除最近記錄")

    def _try_restore_last_file(self):
        """Auto-reopen last file on startup if it still exists."""
        last = self.manager._full_config.get("_last_file", "")
        if last and os.path.isfile(last):
            self._load_path(last)

    # ── UI refresh ────────────────────────────────────────────────────────────

    def _refresh_ui(self):
        self._tab_widget.blockSignals(True)
        self._tab_widget.clear()
        self._editors.clear()

        tables = list(self.manager.tables.keys())
        for tname in tables:
            self._editors[tname] = None
            self._tab_widget.addTab(QWidget(), tname)

        self._tab_widget.blockSignals(False)

        if tables:
            self._show_editor()
            # Defer editor build to after the event loop paints the empty tabs
            QTimer.singleShot(0, lambda: self._ensure_editor(0))
        else:
            self._welcome.refresh(self.manager)
            self._show_welcome()

        self._update_title()
        self.show_snackbar(f"已載入 {len(tables)} 個資料表", color="#7ec8e3")

    def _on_tab_changed(self, idx):
        self._ensure_editor(idx)

    def _ensure_editor(self, idx):
        if idx < 0 or idx >= self._tab_widget.count():
            return
        tname = self._tab_widget.tabText(idx)
        if self._editors.get(tname) is not None:
            return
        editor = TableEditor(tname, self.manager)
        editor.status_message.connect(self.show_snackbar)
        self._editors[tname] = editor          # set BEFORE tab swap to guard re-entrant signal
        # Replace placeholder widget
        self._tab_widget.blockSignals(True)
        self._tab_widget.removeTab(idx)
        self._tab_widget.insertTab(idx, editor, tname)
        self._tab_widget.blockSignals(False)
        self._tab_widget.setCurrentIndex(idx)

    def _update_title(self):
        if self.manager.json_path:
            fname = os.path.basename(self.manager.json_path)
            dirty = " *" if self.manager.dirty else ""
            self.setWindowTitle(f"JsonEditor — {fname}{dirty}")
            self._dirty_lbl.setText("● 未儲存" if self.manager.dirty else "✓ 已儲存")
            self._dirty_lbl.setStyleSheet(
                f"color:{'#e0a040' if self.manager.dirty else '#3a9e6a'};")
            self._path_lbl.setText(self.manager.json_path)
        else:
            self.setWindowTitle("JsonEditor")
            self._dirty_lbl.setText("")
            self._path_lbl.setText("")

    # ── Search ────────────────────────────────────────────────────────────────

    def _show_search(self):
        query, ok = QInputDialog.getText(self, "搜尋", "搜尋所有資料表:")
        if not ok or not query.strip():
            return
        results = self.manager.search_index(query.strip())
        if not results:
            self.show_snackbar("無結果")
            return
        # Jump to first result
        tname, is_sub, row_idx, _cols = results[0]
        if is_sub:
            return
        for i in range(self._tab_widget.count()):
            if self._tab_widget.tabText(i) == tname:
                self._tab_widget.setCurrentIndex(i)
                self._ensure_editor(i)
                editor = self._editors.get(tname)
                if editor:
                    df = self.manager.tables[tname]
                    cls_val = df.at[row_idx, editor.cls_key]
                    editor.current_cls_val = cls_val
                    editor._load_cls_list()
                    editor._load_item_list()
                    editor._load_editor(row_idx)
                break
        self.show_snackbar(f"找到 {len(results)} 筆結果 (已跳至第一筆)")

    # ── Config ────────────────────────────────────────────────────────────────

    def open_config(self):
        idx = self._tab_widget.currentIndex()
        if idx < 0:
            return
        tname = self._tab_widget.tabText(idx)
        if tname not in self.manager.tables:
            return
        self._show_config_dialog(tname)

    def _show_config_dialog(self, table_name):
        from PySide6.QtWidgets import QDialog, QDialogButtonBox, QFormLayout, QScrollArea

        dlg = QDialog(self)
        dlg.setWindowTitle(f"配置 — {table_name}")
        dlg.setMinimumWidth(480)
        dlg.setStyleSheet(APP_STYLE)

        cfg = self.manager.config.get(table_name, {})
        df  = self.manager.tables[table_name]
        cols = list(df.columns)

        outer = QVBoxLayout(dlg)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QWidget()
        form = QFormLayout(content)
        form.setContentsMargins(12, 12, 12, 12)
        scroll.setWidget(content)
        outer.addWidget(scroll)

        pk_var  = QComboBox(); pk_var.addItems(cols)
        cls_var = QComboBox(); cls_var.addItems(cols)
        pk_var.setCurrentText(cfg.get("primary_key", cols[0] if cols else ""))
        cls_var.setCurrentText(cfg.get("classification_key", cols[0] if cols else ""))
        form.addRow("Primary Key:", pk_var)
        form.addRow("Classification Key:", cls_var)

        form.addRow(_sep_frame(), QLabel())
        form.addRow(QLabel("欄位類型:"), QLabel())

        col_type_combos = {}
        cols_cfg = cfg.get("columns", {})
        for col in cols:
            cb = QComboBox()
            cb.addItems(["string", "int", "float", "bool", "enum"])
            cb.setCurrentText(cols_cfg.get(col, {}).get("type", "string"))
            form.addRow(f"  {col}:", cb)
            col_type_combos[col] = cb

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(dlg.accept)
        bb.rejected.connect(dlg.reject)
        outer.addWidget(bb)

        if dlg.exec() != QDialog.Accepted:
            return

        # Apply
        cfg["primary_key"]        = pk_var.currentText()
        cfg["classification_key"] = cls_var.currentText()
        if "columns" not in cfg:
            cfg["columns"] = {}
        for col, cb in col_type_combos.items():
            if col not in cfg["columns"]:
                cfg["columns"][col] = {}
            cfg["columns"][col]["type"] = cb.currentText()
        self.manager.config[table_name] = cfg
        self.manager.save_config()

        editor = self._editors.get(table_name)
        if editor:
            editor.reload_after_config()
        self.show_snackbar("配置已套用", color="#7ec8e3")

    # ── Window close ──────────────────────────────────────────────────────────

    def closeEvent(self, event):
        if self.manager.dirty:
            ans = QMessageBox.question(
                self, "未儲存變更", "有未儲存的變更，是否儲存？",
                QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel
            )
            if ans == QMessageBox.Cancel:
                event.ignore()
                return
            if ans == QMessageBox.Save:
                self.save_file()
        event.accept()


# ── Entry ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(APP_STYLE)
    window = App()
    window.show()
    sys.exit(app.exec())
