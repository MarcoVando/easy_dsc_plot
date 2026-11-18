import sys
import os
from pathlib import Path
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import re
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QSplitter,
    QListWidget, QPushButton, QLabel,
    QComboBox, QListWidgetItem, QAbstractItemView, QFileDialog,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QDialog, QFormLayout, QSpinBox, QDialogButtonBox
)

from PyQt5.QtCore import Qt, pyqtSignal

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure
from matplotlib.lines import Line2D  # to import markers

EXCEL_EXTS = (".xls", ".xlsx", ".xlsm", ".xlsb")
SKIP_SHEETS = {"details", "analysis", "notes", "note"}

# Default per-line legend template. Placeholders are substituted at plot time.
# {fname} -> source file label, {step} -> sheet/step name, {col} -> Y column.
DEFAULT_LABEL_TEMPLATE = "{fname} | {step} : {col}"

# Legend positions offered in the settings dialog. The first entry is a
# synthetic "above the plot" option handled separately in _apply_legend.
LEGEND_LOCATIONS = [
    "above plot",
    "best",
    "upper right",
    "upper left",
    "lower right",
    "lower left",
    "center right",
    "center left",
    "upper center",
    "lower center",
    "center",
]


# ----------------------------------------------------------------------
# Utilities
# ----------------------------------------------------------------------
def load_df_head(df):
    """Combine the two header rows of an already index-column-dropped CSV."""
    row0 = df.loc[0, :].fillna('').astype(str)
    row1 = df.loc[1, :].fillna('').astype(str)
    df.columns = (row0 + ', ' + row1).str.strip(', ')
    df.drop([0, 1], inplace=True)
    df = df.reset_index(drop=True)

    for col in df.columns:
        try:
            df[col] = pd.to_numeric(df[col])
        except Exception:
            print(f"{col} not convertable")
    return df


def _join_header(name, unit):
    """Build a 'Name, Unit' column label, tolerating NaN/empty cells."""
    name = '' if name is None or str(name) in ('nan', 'None', '') else str(name).strip()
    unit = '' if unit is None or str(unit) in ('nan', 'None', '') else str(unit).strip()
    if name and unit:
        return f"{name}, {unit}"
    return name or unit


def clean_excel_sheet(raw):
    """
    Clean a raw TA/DSC sheet read with header=None.

    Layout:
        row 0 : strategy comment (step name)
        row 1 : header names
        row 2 : measuring units
        row 3+: raw data

    Returns a numeric DataFrame with 'Name, Unit' columns.
    """
    if raw.shape[0] < 3:
        return pd.DataFrame()

    headers = raw.iloc[1].tolist()
    units = raw.iloc[2].tolist()
    columns = [_join_header(h, u) for h, u in zip(headers, units)]

    data = raw.iloc[3:].copy()
    data.columns = columns
    data = data.reset_index(drop=True)
    data = data.apply(pd.to_numeric, errors="coerce")
    data = data.dropna(how="all").reset_index(drop=True)
    return data


def extract_label(file):
    """Short human-friendly label from a file path."""
    m = re.search(r"MV-(\w+-\w+-\w+)", str(file))
    if m:
        return m[0]
    return Path(file).stem[:24]


def make_dataset_label(key):
    """Label for a dataset key ('file' or 'file::sheet')."""
    if "::" in key:
        file, sheet = key.split("::", 1)
        return f"{extract_label(file)} | {sheet}"
    return extract_label(key)


def generate_color_list(cmap_name="viridis", n_colors=10):
    cmap = plt.get_cmap(cmap_name)
    # sample evenly across [0, 1]
    return [cmap(i / (n_colors)) for i in range(n_colors)]


class DropListWidget(QListWidget):
    filesDropped = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)

        self.setAcceptDrops(True)
        self.setDragDropMode(QListWidget.DropOnly)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        event.accept()

    def dropEvent(self, event):
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        self.filesDropped.emit(files)


class MplCanvas(FigureCanvas):
    def __init__(self):
        self.fig = Figure(figsize=(8, 6))
        self.ax = self.fig.add_subplot(111)
        super().__init__(self.fig)


class DSCPlotter(QWidget):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("DSC / mDSC Data Plotter")
        self.resize(1280, 820)

        self.datasets = {}   # key -> DataFrame
        self.labels = {}     # key -> display label
        self._loading = False

        self.font_size = 10
        self.legend_loc = "best"

        self.mark_mode = False
        self.marks = []

        self.build_ui()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def build_ui(self):

        main_layout = QHBoxLayout(self)

        splitter = QSplitter(Qt.Horizontal)

        # ----- Left panel -----
        left_widget = QWidget()
        left = QVBoxLayout(left_widget)

        top_row = QHBoxLayout()
        self.export_btn = QPushButton("Export TA CSV")
        self.export_btn.clicked.connect(self.export_sheets_to_csv)
        top_row.addWidget(self.export_btn)

        self.add_btn = QPushButton("Add Files...")
        self.add_btn.clicked.connect(self.add_files_dialog)
        top_row.addWidget(self.add_btn)

        self.clear_btn = QPushButton("Clear Files")
        self.clear_btn.clicked.connect(self.clear_files)
        top_row.addWidget(self.clear_btn)
        left.addLayout(top_row)

        left.addWidget(QLabel("Drop Excel / CSV files here"))

        self.file_list = DropListWidget()
        self.file_list.filesDropped.connect(self.load_files)
        self.file_list.setSelectionMode(QListWidget.MultiSelection)
        self.file_list.itemSelectionChanged.connect(self.on_selection_changed)
        left.addWidget(self.file_list)

        # Each dropdown sits on the right of its own label.
        self.cmap_combo = QComboBox()
        self.cmap_combo.addItems(
            sorted(m for m in plt.colormaps() if not m.endswith("_r"))
        )
        self.cmap_combo.setCurrentText("tab10")
        self._add_labeled_row(left, "Colormap", self.cmap_combo)

        self.scale_combo = QComboBox()
        self.scale_combo.addItems(["lin-lin", "log-log", "lin-log", "log-lin"])
        self._add_labeled_row(left, "Axis scale", self.scale_combo)

        self.x_combo = QComboBox()
        self._add_labeled_row(left, "X Column", self.x_combo)

        left.addWidget(QLabel("Y Columns (marker / shift / label)"))
        self.y_table = QTableWidget(0, 4)
        self.y_table.setHorizontalHeaderLabels(["Y Column", "Marker", "Shift", "Label"])
        self.y_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.y_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.y_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.y_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.y_table.verticalHeader().setVisible(False)
        left.addWidget(self.y_table)

        self.plot_btn = QPushButton("Plot")
        self.plot_btn.clicked.connect(self.plot_data)
        left.addWidget(self.plot_btn)

        # Settings + Mark Point below the Plot button.
        bottom_row = QHBoxLayout()
        self.settings_btn = QPushButton("Settings")
        self.settings_btn.clicked.connect(self.open_settings)
        bottom_row.addWidget(self.settings_btn)

        self.mark_btn = QPushButton("Mark Point")
        self.mark_btn.setCheckable(True)
        self.mark_btn.clicked.connect(self.toggle_mark_mode)
        bottom_row.addWidget(self.mark_btn)
        left.addLayout(bottom_row)

        # ----- Plot panel -----
        self.canvas = MplCanvas()
        self.toolbar = NavigationToolbar(self.canvas, self)
        self.canvas.mpl_connect("button_press_event", self.on_canvas_click)

        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.addWidget(self.toolbar)
        right_layout.addWidget(self.canvas)

        splitter.addWidget(left_widget)
        splitter.addWidget(right_panel)
        splitter.setSizes([360, 900])

        main_layout.addWidget(splitter)

    def _add_labeled_row(self, layout, label_text, widget):
        """Place a label on the left and its widget on the right of one row."""
        row = QHBoxLayout()
        row.addWidget(QLabel(label_text))
        row.addWidget(widget)
        layout.addLayout(row)

    # ------------------------------------------------------------------
    # File handling
    # ------------------------------------------------------------------
    def clear_files(self):
        self.file_list.clear()
        self.datasets.clear()
        self.labels.clear()
        self.x_combo.clear()
        self.y_table.setRowCount(0)

    def add_files_dialog(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select data files",
            "",
            "Data files (*.csv *.xls *.xlsx *.xlsm *.xlsb);;All files (*)",
        )
        if files:
            self.load_files(files)

    # ------------------------------------------------------------------
    # Mark point
    # ------------------------------------------------------------------
    def toggle_mark_mode(self):
        self.mark_mode = self.mark_btn.isChecked()
        if self.mark_mode:
            self.canvas.setCursor(Qt.CrossCursor)
        else:
            self.canvas.unsetCursor()

    def on_canvas_click(self, event):
        ax = self.canvas.ax
        if event.inaxes != ax or event.xdata is None:
            return

        # Right click removes the closest vertical mark.
        if event.button == 3:
            self._delete_closest_mark(event.xdata)
            return

        # Left click places a mark (only while mark mode is on).
        if not self.mark_mode or event.button != 1:
            return
        self._place_mark(event.xdata, event.ydata)

    def _place_mark(self, x, y):
        ax = self.canvas.ax
        ymin, ymax = ax.get_ylim()
        dy = (ymax - ymin) * 0.05 if ymax > ymin else 1.0

        seg, = ax.plot(
            [x, x], [y - dy, y + dy],
            color="black", linewidth=1.0, zorder=5,
        )
        txt = ax.annotate(
            f"{x:.1f}",
            xy=(x, y),
            xytext=(4, 4),
            textcoords="offset points",
            fontsize=self.font_size,
            color="black",
        )
        self.marks.append({"x": x, "artists": [seg, txt]})
        self.canvas.draw_idle()

    def _delete_closest_mark(self, x):
        if not self.marks:
            return
        best = min(self.marks, key=lambda m: abs(m["x"] - x))
        for artist in best["artists"]:
            try:
                artist.remove()
            except Exception:
                pass
        self.marks.remove(best)
        self.canvas.draw_idle()

    def open_settings(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("Plot settings")
        form = QFormLayout(dlg)

        font_spin = QSpinBox()
        font_spin.setRange(4, 48)
        font_spin.setValue(self.font_size)
        font_spin.setSuffix(" pt")
        form.addRow("Font size", font_spin)

        loc_combo = QComboBox()
        loc_combo.addItems(LEGEND_LOCATIONS)
        idx = loc_combo.findText(self.legend_loc)
        loc_combo.setCurrentIndex(idx if idx >= 0 else 0)
        form.addRow("Legend position", loc_combo)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        form.addRow(buttons)

        if dlg.exec_() == QDialog.Accepted:
            self.font_size = font_spin.value()
            self.legend_loc = loc_combo.currentText()
            self.plot_data()  # re-render with the new settings

    def export_sheets_to_csv(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Excel files",
            "",
            "Excel files (*.xls *.xlsx *.xlsm *.xlsb);;All files (*)",
        )
        if not files:
            return

        for file in files:
            try:
                filedir = str(Path(file).parent)
                filename = Path(file).stem
                fexc = pd.ExcelFile(file)
                sheets = fexc.sheet_names
                dfs = [pd.read_excel(file, sheet_name=sheet) for sheet in sheets]

                for df, sheet in zip(dfs, sheets):
                    if sheet == "Details":
                        continue
                    out_filename = filedir + "/" + filename + "-" + sheet + ".csv"
                    if os.path.isfile(out_filename):
                        continue
                    else:
                        df.to_csv(out_filename)
                        print("csv extracted to", out_filename)
            except Exception as e:
                print(f"Failed exporting {file}: {e}")

    def _register(self, key, df, label):
        self.datasets[key] = df
        self.labels[key] = label
        item = QListWidgetItem(label)
        item.setData(Qt.UserRole, key)
        self.file_list.addItem(item)
        return item

    def load_files(self, files):
        self._loading = True
        added_items = []
        try:
            for file in files:
                lower = file.lower()
                if lower.endswith(".csv"):
                    try:
                        print("Loading CSV", file)
                        df = pd.read_csv(file, index_col=0, decimal=".")
                        df = load_df_head(df)
                        key = file
                        if key not in self.datasets:
                            added_items.append(
                                self._register(key, df, make_dataset_label(key))
                            )
                        else:
                            self.datasets[key] = df
                    except Exception as e:
                        print(f"Failed loading {file}: {e}")

                elif lower.endswith(EXCEL_EXTS):
                    try:
                        print("Loading Excel", file)
                        xl = pd.ExcelFile(file)
                        for sheet in xl.sheet_names:
                            if sheet.strip().lower() in SKIP_SHEETS:
                                continue
                            try:
                                raw = pd.read_excel(file, sheet_name=sheet, header=None)
                                df = clean_excel_sheet(raw)
                                if df.empty or df.shape[1] == 0:
                                    continue
                                key = f"{file}::{sheet}"
                                if key not in self.datasets:
                                    added_items.append(
                                        self._register(key, df, make_dataset_label(key))
                                    )
                                else:
                                    self.datasets[key] = df
                            except Exception as e:
                                print(f"Failed loading sheet {sheet} of {file}: {e}")
                    except Exception as e:
                        print(f"Failed loading {file}: {e}")
                else:
                    print(f"Skipping unsupported file: {file}")
        finally:
            self._loading = False

        for item in added_items:
            item.setSelected(True)

        if added_items:
            self.populate_columns(self.datasets[added_items[0].data(Qt.UserRole)])

    # ------------------------------------------------------------------
    # Column population
    # ------------------------------------------------------------------
    def populate_columns(self, df):
        prev_x = self.x_combo.currentText()
        prev_y = {}
        for r in range(self.y_table.rowCount()):
            item = self.y_table.item(r, 0)
            if item is None:
                continue
            marker_item = self.y_table.item(r, 1)
            shift_item = self.y_table.item(r, 2)
            label_item = self.y_table.item(r, 3)
            prev_y[item.text()] = (
                item.checkState() == Qt.Checked,
                marker_item.checkState() == Qt.Checked if marker_item else False,
                shift_item.text() if shift_item else "0",
                label_item.text() if label_item else "",
            )

        self.x_combo.blockSignals(True)
        self.y_table.blockSignals(True)

        self.x_combo.clear()
        self.y_table.setRowCount(0)

        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]) or \
               pd.api.types.is_datetime64_any_dtype(df[col]):
                self.x_combo.addItem(str(col))

            if pd.api.types.is_numeric_dtype(df[col]):
                r = self.y_table.rowCount()
                self.y_table.insertRow(r)

                chk = QTableWidgetItem(str(col))
                chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                checked, marked, shift, label = prev_y.get(
                    str(col), (False, False, "0", ""))
                chk.setCheckState(Qt.Checked if checked else Qt.Unchecked)
                self.y_table.setItem(r, 0, chk)

                marker = QTableWidgetItem("")
                marker.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                marker.setCheckState(Qt.Checked if marked else Qt.Unchecked)
                self.y_table.setItem(r, 1, marker)

                shift_item = QTableWidgetItem(shift)
                self.y_table.setItem(r, 2, shift_item)

                # Empty label -> show the default placeholder template.
                label_item = QTableWidgetItem(label if label else DEFAULT_LABEL_TEMPLATE)
                label_item.setToolTip(
                    "Placeholders: {fname}, {step}, {col}. Empty = default."
                )
                self.y_table.setItem(r, 3, label_item)

        if prev_x and prev_x in [self.x_combo.itemText(i) for i in range(self.x_combo.count())]:
            self.x_combo.setCurrentText(prev_x)

        self.x_combo.blockSignals(False)
        self.y_table.blockSignals(False)

    def on_selection_changed(self):
        if self._loading:
            return
        selected = self.file_list.selectedItems()
        if not selected:
            return
        key = selected[0].data(Qt.UserRole)
        df = self.datasets.get(key)
        if df is None:
            return
        # Only refresh the columns if the new dataset exposes different ones,
        # so that X / Y / marker / shift selections are preserved.
        new_cols = list(df.columns)
        old_cols = [self.x_combo.itemText(i) for i in range(self.x_combo.count())]
        old_cols += [
            self.y_table.item(r, 0).text()
            for r in range(self.y_table.rowCount())
            if self.y_table.item(r, 0) is not None
        ]
        if set(new_cols) != set(old_cols):
            self.populate_columns(df)

    # ------------------------------------------------------------------
    # Plotting
    # ------------------------------------------------------------------
    def _selected_y_entries(self):
        entries = []
        for r in range(self.y_table.rowCount()):
            item = self.y_table.item(r, 0)
            if item is None or item.checkState() != Qt.Checked:
                continue
            col = item.text()
            marker_item = self.y_table.item(r, 1)
            use_marker = marker_item is not None and marker_item.checkState() == Qt.Checked
            shift_item = self.y_table.item(r, 2)
            shift_txt = shift_item.text() if shift_item is not None else "0"
            try:
                shift = float(shift_txt)
            except (TypeError, ValueError):
                shift = 0.0
            label_item = self.y_table.item(r, 3)
            template = label_item.text() if label_item is not None else ""
            if not template.strip():
                template = DEFAULT_LABEL_TEMPLATE
            entries.append((col, shift, use_marker, template))
        return entries

    def plot_data(self):
        if not self.datasets:
            return

        keys = [i.data(Qt.UserRole) for i in self.file_list.selectedItems()]
        if not keys:
            return

        x_col = self.x_combo.currentText()
        y_entries = self._selected_y_entries()
        if not y_entries or not x_col:
            return

        ax = self.canvas.ax
        ax.clear()
        self.marks = []

        cmap_name = self.cmap_combo.currentText()
        markers = Line2D.filled_markers[1:]
        linestyles = ["-", "--", "-.", ":", (0, (1, 1)), (0, (5, 2)),
                      (0, (3, 1, 1, 1)), (0, (5, 1, 1, 1))]
        scale_map = {"lin": "linear", "log": "log"}
        xscale, yscale = self.scale_combo.currentText().split("-")
        xscale = scale_map.get(xscale, xscale)
        yscale = scale_map.get(yscale, yscale)

        # Group the selected datasets (sheets) by their source file.
        groups = {}
        for key in keys:
            file = key.split("::", 1)[0]
            groups.setdefault(file, []).append(key)

        # Step index of each dataset within its own file (0, 1, 2, ...).
        step_idx_of = {}
        counters = {}
        for key in keys:
            file = key.split("::", 1)[0]
            step_idx_of[key] = counters.get(file, 0)
            counters[file] = counters.get(file, 0) + 1

        # Adaptive visual encoding: assign channels (color -> linestyle ->
        # marker) only to the dimensions that actually vary, in order of
        # coarseness (file > step > column). Dimensions with a single value
        # keep the channel default, so they never waste a visual channel.
        n_file = len(set(key.split("::", 1)[0] for key in keys))
        n_step = len(set(step_idx_of.values()))
        n_col = len(y_entries)

        varying = []
        if n_file > 1:
            varying.append("file")
        if n_step > 1:
            varying.append("step")
        if n_col > 1:
            varying.append("col")

        channel_of = {dim: ch for dim, ch in zip(varying, ["color", "linestyle", "marker"])}
        dim_of = {ch: dim for dim, ch in channel_of.items()}

        # Pre-build palettes for whichever dimension drives each channel.
        def palette(dim):
            if dim == "file":
                vals = list(groups.keys())
            elif dim == "step":
                vals = list(range(n_step))
            else:  # col
                vals = list(range(n_col))
            return {v: c for v, c in zip(vals, generate_color_list(cmap_name, len(vals)))}

        color_pal = palette(dim_of["color"]) if "color" in dim_of else {}
        ls_pal = palette(dim_of["linestyle"]) if "linestyle" in dim_of else {}
        marker_pal = palette(dim_of["marker"]) if "marker" in dim_of else {}

        default_color = generate_color_list(cmap_name, 1)[0]

        for key in keys:
            df = self.datasets[key]
            file = key.split("::", 1)[0]
            step = step_idx_of[key]

            for j, (y, shift, use_marker, template) in enumerate(y_entries):
                if x_col not in df.columns or y not in df.columns:
                    continue

                xs = df[x_col]
                ys = df[y]
                if shift:
                    ys = ys + shift

                fname = extract_label(file)
                step_name = key.split("::", 1)[1] if "::" in key else ""
                lbl = (template
                       .replace("{fname}", fname)
                       .replace("{step}", step_name)
                       .replace("{stepname}", step_name)
                       .replace("{col}", str(y)))
                if shift:
                    lbl += f" ({shift:+g})"

                color_dim = dim_of.get("color")
                if color_dim == "file":
                    color = color_pal.get(file, default_color)
                elif color_dim == "step":
                    color = color_pal.get(step, default_color)
                elif color_dim == "col":
                    color = color_pal.get(j, default_color)
                else:
                    color = default_color

                if dim_of.get("linestyle") == "file":
                    ls = linestyles[list(groups.keys()).index(file) % len(linestyles)]
                elif dim_of.get("linestyle") == "step":
                    ls = linestyles[step % len(linestyles)]
                elif dim_of.get("linestyle") == "col":
                    ls = linestyles[j % len(linestyles)]
                else:
                    ls = "-"

                if dim_of.get("marker") == "col":
                    marker = markers[j % len(markers)] if use_marker else None
                elif dim_of.get("marker") == "step":
                    marker = markers[step % len(markers)] if use_marker else None
                elif dim_of.get("marker") == "file":
                    fidx = list(groups.keys()).index(file)
                    marker = markers[fidx % len(markers)] if use_marker else None
                else:
                    marker = markers[j % len(markers)] if use_marker else None

                try:
                    ax.plot(
                        xs, ys,
                        label=lbl,
                        color=color,
                        linestyle=ls,
                        marker=marker,
                    )
                except Exception as e:
                    print(f"Could not plot {key} {y}: {e}")

        ax.set_xscale(xscale)
        ax.set_yscale(yscale)
        ax.set_xlabel(x_col, fontsize=self.font_size)
        ax.set_ylabel(", ".join(e[0] for e in y_entries), fontsize=self.font_size)
        ax.tick_params(labelsize=self.font_size)
        self._apply_legend(ax)
        ax.grid(True, which="both", alpha=0.3)
        self.canvas.draw()

    def _apply_legend(self, ax):
        fs = self.font_size
        if self.legend_loc == "above plot":
            n = len(ax.lines)
            ncol = self._above_legend_ncol(ax, fs, n)
            ax.legend(
                fontsize=fs,
                loc="lower center",
                bbox_to_anchor=(0.5, 1.02),
                ncol=ncol,
                frameon=False,
            )
        else:
            ax.legend(fontsize=fs, loc=self.legend_loc)

    def _above_legend_ncol(self, ax, fs, n):
        """
        Number of columns for the above-plot legend.

        Start from a single line and only wrap to extra rows/columns when the
        entries would not fit the figure width.
        """
        if n <= 1:
            return 1

        width_in = ax.figure.get_size_inches()[0]
        avail_pt = width_in * 72.0 * 0.95  # usable width in points

        labels = [l.get_label() for l in ax.lines] or [""]
        max_chars = max(len(s) for s in labels)
        char_w = fs * 0.6                 # rough average glyph width
        entry_w = max_chars * char_w + fs * 3.0  # + handle / padding

        fit = max(1, int(avail_pt // entry_w)) if entry_w > 0 else n
        return min(n, fit)


if __name__ == "__main__":

    app = QApplication(sys.argv)

    w = DSCPlotter()
    w.show()

    sys.exit(app.exec_())