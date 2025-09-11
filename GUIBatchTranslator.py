import os
import sys
import shutil
import traceback
import argostranslate.translate as T
from pathlib import Path
from PyQt5 import QtCore, QtWidgets
from argostranslate.package import install_from_path

class _PathInstallWorker(QtCore.QObject):
    progress = QtCore.pyqtSignal(str)
    finished = QtCore.pyqtSignal(bool)

    def __init__(self, items, install_one_callable):
        super().__init__()
        self._items = list(items)
        self._install_one = install_one_callable
        self._cancelled = False

    @QtCore.pyqtSlot()
    def run(self):
        total = len(self._items)
        for i, path in enumerate(self._items, start=1):
            if self._cancelled:
                self.finished.emit(False)
                return
            name = os.path.basename(path)
            self.progress.emit(f"Installing {name}… ({i}/{total})")
            try:
                self._install_one(path)
            except Exception as e:
                self.progress.emit(f"Error installing {name}: {e}")
                self.finished.emit(False)
                return
        self.progress.emit("All language packages installed.")
        self.finished.emit(True)

    def cancel(self):
        self._cancelled = True


def _run_path_installs_with_popup(parent, paths, install_one_callable, title="Installing language packs"):
    """
    Show a modal QProgressDialog and run install_one_callable(path) on a worker thread.
    Returns True if dialog wasn't cancelled (installs attempted to completion).
    """
    if not paths:
        return True

    dlg = QtWidgets.QProgressDialog("Preparing…", "Cancel", 0, 0, parent)
    dlg.setWindowTitle(title)
    dlg.setWindowModality(QtCore.Qt.ApplicationModal)
    dlg.setAutoClose(False)
    dlg.setAutoReset(False)
    dlg.setRange(0, 0)  # indeterminate
    dlg.setMinimumWidth(420)

    thread = QtCore.QThread(parent)
    worker = _PathInstallWorker(paths, install_one_callable)
    worker.moveToThread(thread)

    worker.progress.connect(dlg.setLabelText)
    worker.finished.connect(lambda _ok: dlg.done(0))
    thread.started.connect(worker.run)

    def _cleanup():
        worker.deleteLater()
        thread.quit()
        thread.wait()
        thread.deleteLater()
    dlg.finished.connect(_cleanup)

    dlg.canceled.connect(worker.cancel)

    thread.start()
    dlg.exec_()
    return not dlg.wasCanceled()

class _InstallWorker(QtCore.QObject):
    progress = QtCore.pyqtSignal(str)   # status text
    finished = QtCore.pyqtSignal(bool)  # True if all done, False if cancelled/error

    def __init__(self, items, install_one_callable):
        super().__init__()
        self._items = list(items)
        self._install_one = install_one_callable
        self._cancelled = False

    @QtCore.pyqtSlot()
    def run(self):
        total = len(self._items)
        for i, (src, dst) in enumerate(self._items, start=1):
            if self._cancelled:
                self.finished.emit(False)
                return
            self.progress.emit(f"Installing {src} \u2192 {dst}\u2026 ({i}/{total})")
            try:
                # Your real per-pair installer:
                #   self._install_one(src, dst)
                self._install_one(src, dst)
            except Exception as e:
                self.progress.emit(f"Error installing {src}\u2192{dst}: {e}")
                self.finished.emit(False)
                return
        self.progress.emit("All language packs installed.")
        self.finished.emit(True)

    def cancel(self):
        self._cancelled = True


def install_language_packs_with_popup(parent, pairs, install_one_callable):
    """
    parent: QWidget
    pairs:  list[tuple[str,str]] like [("en","es"), ("en","fr"), ...]
    install_one_callable: function(src:str, dst:str) -> None  (blocking)
    """
    if not pairs:
        return True  # nothing to do

    # Indeterminate progress dialog
    dlg = QtWidgets.QProgressDialog("Preparing\u2026", "Cancel", 0, 0, parent)
    dlg.setWindowTitle("Installing language packs")
    dlg.setWindowModality(QtCore.Qt.ApplicationModal)
    dlg.setMinimumWidth(420)
    dlg.setAutoClose(False)
    dlg.setAutoReset(False)
    dlg.setCancelButtonText("Cancel")
    dlg.setRange(0, 0)  # indeterminate (busy)

    # Worker in background thread
    thread = QtCore.QThread(parent)
    worker = _InstallWorker(pairs, install_one_callable)
    worker.moveToThread(thread)

    # Wire signals
    worker.progress.connect(dlg.setLabelText)
    worker.finished.connect(lambda ok: dlg.done(0))
    thread.started.connect(worker.run)

    # Ensure cleanup
    def _cleanup():
        worker.deleteLater()
        thread.quit()
        thread.wait()
        thread.deleteLater()
    dlg.finished.connect(_cleanup)

    # Cancel button support
    def _on_cancel():
        worker.cancel()
    dlg.canceled.connect(_on_cancel)

    # Go
    thread.start()
    dlg.exec_()  # modal loop; returns when finished or canceled

    # If user pressed cancel, we already told worker to stop; return False so caller can react
    return not dlg.wasCanceled()


SUPPORTED_EXTS = {".txt", ".docx", ".odt", ".pptx", ".odp", ".epub",
                  ".html", ".htm", ".srt", ".pdf",".xlsx",".xls"}  # Note: scanned PDFs need OCR first


def human_lang(l):
    # display name like "English (en)"
    return f"{getattr(l, 'name', l.code).title()} ({l.code})"

def find_bundled_models_dir():
    # When frozen with PyInstaller, resources live under _MEIPASS
    base = getattr(sys, "_MEIPASS", os.path.abspath(os.path.dirname(__file__)))
    mdir = os.path.join(base, "models")
    return mdir if os.path.isdir(mdir) else None

def ensure_bundled_models_installed(parent=None):
    """Install any .argosmodel files found in a bundled 'models' dir (with a modal popup)."""
    mdir = find_bundled_models_dir()
    if not mdir:
        return []

    paths = [str(p) for p in Path(mdir).glob("*.argosmodel")]
    if not paths:
        return []

    installed_any = []

    def _install_one(path):
        install_from_path(path)
        installed_any.append(path)

    # Show the modal "Installing…" while we process the files
    _run_path_installs_with_popup(parent, paths, _install_one, title="Installing language packs")

    if installed_any and parent:
        QtWidgets.QMessageBox.information(
            parent, "Language packages installed",
            f"Installed {len(installed_any)} bundled language package(s)."
        )
    return installed_any


def get_lang_by_code(code):
    for l in T.get_installed_languages():
        if l.code == code:
            return l
    return None


def get_translation_or_none(src_code, dst_code):
    src = get_lang_by_code(src_code)
    dst = get_lang_by_code(dst_code)
    if not src or not dst:
        return None
    try:
        return src.get_translation(dst)
    except Exception:
        return None

def translate_text_direct(tr, text: str) -> str:
    return tr.translate(text) if text else text

def translate_text_pivot(to_en, en_to, text: str) -> str:
    if not text:
        return text
    return en_to.translate(to_en.translate(text))
    
def _translate_xlsx_file(in_path, out_dir, tr, to_en, en_to):
    import openyxl  # lazy import
    wb = openpyxl.load_workbook(in_path)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                # Only translate string literals (do NOT touch formulas or numbers)
                if cell.data_type == "s" and isinstance(cell.value, str) and cell.value.strip():
                    try:
                        if tr:
                            cell.value = translate_text_direct(tr, cell.value)
                        else:
                            cell.value = translate_text_pivot(to_en, en_to, cell.value)
                    except Exception as e:
                        # Leave cell unchanged on error; you’ll see errors in the log
                        print(f"Translate fail {ws.title}!{cell.coordinate}: {e}")
    out = Path(out_dir) / (Path(in_path).stem + "_translated.xlsx")
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return str(out)

def _translate_xls_file_to_xlsx(in_path, out_dir, tr, to_en, en_to):
    import xlrd
    from xlrd.xldate import xldate_as_datetime
    book = xlrd.open_workbook(in_path)
    import openyxl
    out_wb = openpyxl.Workbook()
    # Remove default sheet if we will create our own
    if out_wb.worksheets:
        out_wb.remove(out_wb.active)

    for s in book.sheets():
        ws = out_wb.create_sheet(title=s.name[:31] or "Sheet1")  # Excel title max 31 chars
        for r in range(s.nrows):
            for c in range(s.ncols):
                cell = s.cell(r, c)
                v = cell.value
                try:
                    if cell.ctype == xlrd.XL_CELL_TEXT and isinstance(v, str) and v.strip():
                        if tr:
                            v = translate_text_direct(tr, v)
                        else:
                            v = translate_text_pivot(to_en, en_to, v)
                    elif cell.ctype == xlrd.XL_CELL_DATE:
                        v = xldate_as_datetime(v, book.datemode)
                    # numbers, bools, blanks, errors -> write as-is
                except Exception as e:
                    print(f"Translate fail {s.name}!R{r+1}C{c+1}: {e}")
                ws.cell(row=r+1, column=c+1, value=v)

    out = Path(out_dir) / (Path(in_path).stem + "_translated.xlsx")
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    out_wb.save(out)
    return str(out)
    
def translate_excel_file(in_path, src_code, dst_code, out_dir):
    """Handles .xlsx/.xls via openpyxl/xlrd, with optional EN pivot."""
    ext = Path(in_path).suffix.lower()

    tr = get_translation_or_none(src_code, dst_code)
    to_en = en_to = None
    if tr is None and src_code != "en" and dst_code != "en":
        to_en = get_translation_or_none(src_code, "en")
        en_to = get_translation_or_none("en", dst_code)
        if not (to_en and en_to):
            raise RuntimeError(
                f"No translation path for Excel ({src_code}→{dst_code}). "
                f"Install the required .argosmodel packages."
            )

    if ext == ".xlsx":
        return _translate_xlsx_file(in_path, out_dir, tr, to_en, en_to)
    elif ext == ".xls":
        return _translate_xls_file_to_xlsx(in_path, out_dir, tr, to_en, en_to)
    else:
        raise RuntimeError("translate_excel_file called with non-Excel file.")



def translate_with_optional_pivot(in_path, src_code, dst_code, out_dir):
    """
    Try direct translation first. If missing, pivot via English (en).
    Returns final output path.
    """
    ext = Path(in_path).suffix.lower()
    # Excel → use our own handlers (AF returns bool for these)
    if ext in {".xlsx", ".xls"}:
        return translate_excel_file(in_path, src_code, dst_code, out_dir)

    # Non-Excel → use argos-translate-files
    from argostranslatefiles import argostranslatefiles as AF
    tr = get_translation_or_none(src_code, dst_code)
    if tr:
        out_path = AF.translate_file(tr, str(in_path))
        if isinstance(out_path, (str, os.PathLike)) and os.path.exists(out_path):
            return move_to_dir(out_path, out_dir)
        raise RuntimeError(
            f"Unexpected return from argos-translate-files for '{in_path}': {out_path!r}"
        )

    # Pivot via English for non-Excel
    if src_code != "en" and dst_code != "en":
        to_en = get_translation_or_none(src_code, "en")
        en_to = get_translation_or_none("en", dst_code)
        if to_en and en_to:
            mid = AF.translate_file(to_en, str(in_path))
            if not (isinstance(mid, (str, os.PathLike)) and os.path.exists(mid)):
                raise RuntimeError(f"Pivot step to English failed; returned {mid!r}.")
            final = AF.translate_file(en_to, str(mid))
            if not (isinstance(final, (str, os.PathLike)) and os.path.exists(final)):
                raise RuntimeError(f"Pivot step to target failed; returned {final!r}.")
            try:
                os.remove(mid)
            except Exception:
                pass
            return move_to_dir(final, out_dir)

    # If we reach here, we can't translate with current models
    raise RuntimeError(
        f"No translation path available ({src_code} → {dst_code}). "
        f"Install the appropriate .argosmodel packages."
    )



def move_to_dir(path, out_dir):
    """Move translated file to chosen output directory (keeping basename)."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    path = Path(path)
    dest = out_dir / path.name
    if str(path) != str(dest):
        if dest.exists():
            dest.unlink()
        shutil.move(str(path), str(dest))
    return str(dest)


class Worker(QtCore.QObject):
    progress = QtCore.pyqtSignal(int, str)          # percent, message
    file_done = QtCore.pyqtSignal(str, str)         # input_path, output_path
    error = QtCore.pyqtSignal(str, str)             # input_path, error_message
    finished = QtCore.pyqtSignal()

    def __init__(self, files, src_code, dst_code, out_dir):
        super().__init__()
        self.files = files
        self.src = src_code
        self.dst = dst_code
        self.out_dir = out_dir
        self._abort = False

    @QtCore.pyqtSlot()
    def run(self):
        total = len(self.files)
        for idx, f in enumerate(self.files, 1):
            if self._abort:
                break
            try:
                msg = f"Translating ({idx}/{total}): {os.path.basename(f)}"
                self.progress.emit(int((idx-1) / total * 100), msg)
                outp = translate_with_optional_pivot(f, self.src, self.dst, self.out_dir)
                self.file_done.emit(f, outp)
            except Exception as e:
                err = "".join(traceback.format_exception_only(type(e), e)).strip()
                self.error.emit(f, err)
        self.progress.emit(100, "Done.")
        self.finished.emit()

    def abort(self):
        self._abort = True


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Argos Document Translator (Offline)")
        self.resize(880, 560)

        # Widgets
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        layout = QtWidgets.QVBoxLayout(central)

        # Top: language selectors
        lang_row = QtWidgets.QHBoxLayout()
        self.src_combo = QtWidgets.QComboBox()
        self.dst_combo = QtWidgets.QComboBox()
        self.refresh_btn = QtWidgets.QPushButton("Refresh languages")
        self.install_btn = QtWidgets.QPushButton("Install .argosmodel…")
        lang_row.addWidget(QtWidgets.QLabel("From:"))
        lang_row.addWidget(self.src_combo, 1)
        lang_row.addSpacing(12)
        lang_row.addWidget(QtWidgets.QLabel("To:"))
        lang_row.addWidget(self.dst_combo, 1)
        lang_row.addSpacing(12)
        lang_row.addWidget(self.refresh_btn)
        lang_row.addWidget(self.install_btn)

        # Middle: file list + buttons
        file_row = QtWidgets.QHBoxLayout()
        self.file_list = QtWidgets.QListWidget()
        self.file_list.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        btn_col = QtWidgets.QVBoxLayout()
        self.add_files_btn = QtWidgets.QPushButton("Add files…")
        self.add_folder_btn = QtWidgets.QPushButton("Add folder…")
        self.clear_btn = QtWidgets.QPushButton("Clear list")
        btn_col.addWidget(self.add_files_btn)
        btn_col.addWidget(self.add_folder_btn)
        btn_col.addWidget(self.clear_btn)
        btn_col.addStretch(1)
        file_row.addWidget(self.file_list, 1)
        file_row.addLayout(btn_col)

        # Output dir
        out_row = QtWidgets.QHBoxLayout()
        self.out_dir_edit = QtWidgets.QLineEdit()
        self.out_dir_btn = QtWidgets.QPushButton("Choose output folder…")
        out_row.addWidget(QtWidgets.QLabel("Output folder:"))
        out_row.addWidget(self.out_dir_edit, 1)
        out_row.addWidget(self.out_dir_btn)

        # Bottom: run + progress + log
        run_row = QtWidgets.QHBoxLayout()
        self.run_btn = QtWidgets.QPushButton("Translate")
        self.cancel_btn = QtWidgets.QPushButton("Cancel")
        self.cancel_btn.setEnabled(False)
        self.progress = QtWidgets.QProgressBar()
        self.progress.setValue(0)
        run_row.addWidget(self.run_btn)
        run_row.addWidget(self.cancel_btn)
        run_row.addWidget(self.progress, 1)

        self.log = QtWidgets.QTextEdit()
        self.log.setReadOnly(True)

        layout.addLayout(lang_row)
        layout.addLayout(file_row)
        layout.addLayout(out_row)
        layout.addLayout(run_row)
        layout.addWidget(self.log, 1)

        self.worker = None
        self.thread = None

        # Wire up
        self.refresh_btn.clicked.connect(self.populate_languages)
        self.install_btn.clicked.connect(self.install_models_dialog)
        self.add_files_btn.clicked.connect(self.add_files)
        self.add_folder_btn.clicked.connect(self.add_folder)
        self.clear_btn.clicked.connect(self.file_list.clear)
        self.out_dir_btn.clicked.connect(self.choose_out_dir)
        self.run_btn.clicked.connect(self.start_run)
        self.cancel_btn.clicked.connect(self.cancel_run)

        # First-run: defer heavy work so the window shows instantly
        QtCore.QTimer.singleShot(0, lambda: (ensure_bundled_models_installed(self), self.populate_languages()))


    def populate_languages(self):
        self.src_combo.clear()
        self.dst_combo.clear()
        langs = T.get_installed_languages()
        # keep a small mapping code->display
        for l in sorted(langs, key=lambda x: x.code):
            label = human_lang(l)
            self.src_combo.addItem(label, l.code)
            self.dst_combo.addItem(label, l.code)
        # sensible defaults
        find_and_set(self.src_combo, "en")
        find_and_set(self.dst_combo, "es")

    def add_files(self):
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self, "Choose files to translate", "", 
            "Documents (*.txt *.docx *.odt *.pptx *.odp *.epub *.html *.htm *.srt *.pdf *.xls *.xlsx);;All files (*.*)")
        for f in files:
            if Path(f).suffix.lower() in SUPPORTED_EXTS:
                self.file_list.addItem(f)

    def add_folder(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Choose a folder")
        if not folder:
            return
        count = 0
        for root, _, files in os.walk(folder):
            for name in files:
                if Path(name).suffix.lower() in SUPPORTED_EXTS:
                    self.file_list.addItem(str(Path(root) / name))
                    count += 1
        QtWidgets.QMessageBox.information(self, "Folder added", f"Added {count} file(s).")

    def choose_out_dir(self):
        d = QtWidgets.QFileDialog.getExistingDirectory(self, "Choose output folder")
        if d:
            self.out_dir_edit.setText(d)

    def start_run(self):
        items = [self.file_list.item(i).text() for i in range(self.file_list.count())]
        if not items:
            QtWidgets.QMessageBox.warning(self, "No files", "Add at least one file.")
            return
        out_dir = self.out_dir_edit.text().strip() or os.path.join(os.path.expanduser("~"), "Translations")
        src = self.src_combo.currentData()
        dst = self.dst_combo.currentData()
        if src == dst:
            QtWidgets.QMessageBox.warning(self, "Language pair", "Choose different source and target languages.")
            return

        self.progress.setValue(0)
        self.run_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.log.clear()

        self.thread = QtCore.QThread()
        self.worker = Worker(items, src, dst, out_dir)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.on_progress)
        self.worker.file_done.connect(self.on_file_done)
        self.worker.error.connect(self.on_error)
        self.worker.finished.connect(self.on_finished)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        self.thread.start()

    def cancel_run(self):
        if self.worker:
            self.worker.abort()

    @QtCore.pyqtSlot(int, str)
    def on_progress(self, pct, msg):
        self.progress.setValue(pct)
        self.log.append(msg)

    @QtCore.pyqtSlot(str, str)
    def on_file_done(self, inp, outp):
        self.log.append(f"✔ {os.path.basename(inp)} → {outp}")

    @QtCore.pyqtSlot(str, str)
    def on_error(self, inp, err):
        self.log.append(f"❌ {os.path.basename(inp)} → {err}")

    @QtCore.pyqtSlot()
    def on_finished(self):
        self.run_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        self.log.append("All done.")

    def install_models_dialog(self):
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self, "Select .argosmodel files", "", "Argos models (*.argosmodel)")
        ok = 0
        for f in files:
            try:
                install_from_path(f)
                ok += 1
            except Exception as e:
                self.log.append(f"Model install failed: {os.path.basename(f)} → {e}")
        if ok:
            self.log.append(f"Installed {ok} language package(s).")
            self.populate_languages()


def find_and_set(combo: QtWidgets.QComboBox, code: str):
    for i in range(combo.count()):
        if combo.itemData(i) == code:
            combo.setCurrentIndex(i)
            return


def main():
    app = QtWidgets.QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
