"""
NIPT Report Generator – Desktop Application
UI mirrors PGT-A generator exactly:
  - Manual tab: left form + right QPdfView, 1s debounce auto-refresh
  - Batch tab: left patient list | center full inline editor | right PDF preview
  - Neutral clean stylesheet (no bright blue)
  - Compare: manual folder vs automated folder
"""

import sys
import os
import json
import subprocess
from datetime import datetime

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QLabel, QLineEdit, QTextEdit, QPushButton, QFileDialog,
    QTableWidget, QTableWidgetItem, QMessageBox, QProgressBar,
    QGroupBox, QFormLayout, QScrollArea, QSplitter, QHeaderView,
    QCheckBox, QTextBrowser, QListWidget, QListWidgetItem,
    QStyle, QComboBox, QRadioButton
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSettings, QTimer
from PyQt6.QtGui import QFont, QColor, QBrush

try:
    from PyQt6.QtPdf import QPdfDocument
    from PyQt6.QtPdfWidgets import QPdfView
    HAS_PDF_VIEW = True
except ImportError:
    QPdfDocument = None
    QPdfView     = None
    HAS_PDF_VIEW = False


# =============================================================================
# Workers
# =============================================================================

class PreviewWorker(QThread):
    finished = pyqtSignal(str)
    error    = pyqtSignal(str)

    def __init__(self, p_info, z_scores, output_path, show_logo=True):
        super().__init__()
        self.p_info      = p_info
        self.z_scores    = z_scores
        self.output_path = output_path
        self.show_logo   = show_logo

    def run(self):
        try:
            from nipt_template import NIPTReportTemplate
            NIPTReportTemplate(self.output_path).generate(
                self.p_info, self.z_scores, with_logo=self.show_logo)
            self.finished.emit(self.output_path)
        except Exception as e:
            self.error.emit(str(e))


class BatchWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(int, int)
    error    = pyqtSignal(str)

    def __init__(self, patients, out_dir, do_pdf, do_docx, branding):
        super().__init__()
        self.patients  = patients
        self.out_dir   = out_dir
        self.do_pdf    = do_pdf
        self.do_docx   = do_docx
        self.branding  = branding

    def run(self):
        from nipt_template       import NIPTReportTemplate
        from nipt_docx_generator import NIPTDocxGenerator
        import pandas as pd

        total = len(self.patients)
        count = 0
        for i, p in enumerate(self.patients):
            name = p.get("name", p.get("Patient Name", f"Patient_{i}"))
            s_id = p.get("sample_id", p.get("Sample ID", f"S{i}"))
            self.progress.emit(int(i / total * 100), f"Processing {name} …")
            try:
                z = {f"chr{j}": float(p.get(f"chr{j}", 0) or 0) for j in range(1, 23)}
                z["chrX"]           = float(p.get("chrX", 0) or 0)
                z["fetal_fraction"] = float(p.get("ff", p.get("fetal_fraction", 0)) or 0)

                p_info = {k: str(p.get(k, "")) for k in [
                    "name","pin","dob","dob_type","ga","sample_id",
                    "collection_date","received_date","preg_status",
                    "preg_type","clinician","hospital","indication","specimen"]}

                suffix = "with_logo" if self.branding else "without_logo"
                base = f"Report_{name.replace(' ','_')}_{s_id}_{suffix}"
                if self.do_pdf:
                    NIPTReportTemplate(os.path.join(self.out_dir, base+".pdf")).generate(
                        p_info, z, with_logo=self.branding)
                if self.do_docx:
                    NIPTDocxGenerator(os.path.join(self.out_dir, base+".docx")).generate(
                        p_info, z, with_logo=self.branding)
                count += 1
            except Exception as e:
                self.error.emit(f"Error – {name}: {e}")
        self.finished.emit(count, total)


# =============================================================================
# Main Window
# =============================================================================

class NIPTApp(QMainWindow):
    APP_NAME = "NIPT Report Generator"


    # Neutral clean stylesheet – no harsh blue, follows PGT-A's softer gray palette
    APP_STYLESHEET = """
    QMainWindow, QWidget {
        background-color: #f5f5f5;
        font-family: 'Segoe UI', Arial, sans-serif;
        font-size: 13px;
        color: #2d2d2d;
    }
    QGroupBox {
        font-weight: bold;
        border: 1px solid #d0d0d0;
        border-radius: 5px;
        margin-top: 10px;
        padding-top: 8px;
        background-color: #ffffff;
        color: #333;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        subcontrol-position: top left;
        padding: 2px 8px;
        color: #555;
        background-color: #f0f0f0;
        border-radius: 3px;
    }
    QLineEdit, QTextEdit {
        border: 1px solid #c8c8c8;
        border-radius: 3px;
        padding: 4px 6px;
        background-color: #ffffff;
        color: #2d2d2d;
    }
    QLineEdit:focus, QTextEdit:focus {
        border: 1px solid #888;
    }
    QLineEdit[readOnly="true"] {
        background-color: #f7f7f7;
        color: #555;
    }
    QTabWidget::pane {
        border: 1px solid #d0d0d0;
        background: #ffffff;
    }
    QTabBar::tab {
        background: #e8e8e8;
        border: 1px solid #d0d0d0;
        padding: 7px 18px;
        margin-right: 2px;
        border-top-left-radius: 4px;
        border-top-right-radius: 4px;
        color: #555;
    }
    QTabBar::tab:selected {
        background: #1F497D;
        color: white;
    }
    QTabBar::tab:hover:!selected {
        background: #dde3eb;
    }
    QPushButton {
        border: 1px solid #b0b0b0;
        border-radius: 4px;
        padding: 5px 14px;
        background-color: #f0f0f0;
        color: #333;
    }
    QPushButton:hover  { background-color: #e2e2e2; }
    QPushButton:pressed { background-color: #d5d5d5; }
    QTableWidget {
        gridline-color: #e0e0e0;
        border: 1px solid #d0d0d0;
        background: white;
        alternate-background-color: #fafafa;
    }
    QHeaderView::section {
        background-color: #1F497D;
        color: white;
        padding: 5px;
        font-weight: bold;
        border: none;
    }
    QListWidget {
        border: 1px solid #d0d0d0;
        background: white;
        alternate-background-color: #fafafa;
    }
    QListWidget::item:selected {
        background-color: #1F497D;
        color: white;
    }
    QProgressBar {
        border: 1px solid #b0b0b0;
        border-radius: 3px;
        background: #f0f0f0;
        text-align: center;
    }
    QProgressBar::chunk { background-color: #1F497D; border-radius: 3px; }
    QScrollArea { border: none; }
    QSplitter::handle { background-color: #d0d0d0; }
    """

    def __init__(self):
        super().__init__()
        self.setWindowTitle(self.APP_NAME)
        self.setMinimumSize(1180, 820)

        self.setStyleSheet(self.APP_STYLESHEET)

        self.settings       = QSettings("NIPT", "ReportGeneratorV4")
        self.all_patients   = []          # list of dicts (raw from excel)
        self.batch_patients = []          # list of dicts (editable, normalised)
        self._preview_worker = None
        self._batch_current  = -1        # index of selected batch patient
        self._preview_tmp = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "_nipt_preview.pdf")
        self._batch_preview_tmp = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "_nipt_batch_preview.pdf")

        self._init_ui()
        self._load_settings()
        self.statusBar().showMessage(f"Ready  •  {self.APP_NAME}")

    # ─────────────────────────────────────────────────────────────────────────
    # UI Shell
    # ─────────────────────────────────────────────────────────────────────────

    def _init_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        vbox = QVBoxLayout(root)
        vbox.setContentsMargins(8, 8, 8, 6)
        vbox.setSpacing(6)

        # Header
        hdr = QLabel(self.APP_NAME)
        hdr.setFont(QFont("Segoe UI", 16, QFont.Weight.Bold))
        hdr.setStyleSheet(
            "color: white; background-color: #1F497D; "
            "border-radius: 5px; padding: 6px 14px;")
        vbox.addWidget(hdr)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_manual_tab(),  "Manual Entry")
        self.tabs.addTab(self._build_bulk_tab(),    "Batch Upload")
        self.tabs.addTab(self._build_compare_tab(), "Report Comparison")
        self.tabs.addTab(self._build_guide_tab(),   "User Guide")

        # Set icons
        self.tabs.setTabIcon(0, self.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon))
        self.tabs.setTabIcon(1, self.style().standardIcon(QStyle.StandardPixmap.SP_FileDialogListView))
        self.tabs.setTabIcon(2, self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload))
        self.tabs.setTabIcon(3, self.style().standardIcon(QStyle.StandardPixmap.SP_MessageBoxInformation))
        vbox.addWidget(self.tabs)


        # Shared footer
        foot = QGroupBox("Export Settings")
        fl   = QHBoxLayout(foot)
        self.cb_pdf      = QCheckBox("PDF");      self.cb_pdf.setChecked(True)
        self.cb_docx     = QCheckBox("DOCX");     self.cb_docx.setChecked(True)
        self.cb_branding = QCheckBox("With Branding Logo"); self.cb_branding.setChecked(True)
        self.cb_branding.stateChanged.connect(self._schedule_preview)
        self.out_dir_edit = QLineEdit()
        self.out_dir_edit.setPlaceholderText("Select output folder …")
        btn_br = QPushButton("Browse")
        btn_br.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon))
        btn_br.clicked.connect(self._browse_out)

        for w in [self.cb_pdf, self.cb_docx, self.cb_branding]:
            fl.addWidget(w)
        fl.addSpacing(20); fl.addWidget(QLabel("Output Folder:"))
        fl.addWidget(self.out_dir_edit, 1); fl.addWidget(btn_br)
        vbox.addWidget(foot)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        vbox.addWidget(self.progress_bar)

    # ─────────────────────────────────────────────────────────────────────────
    # Manual Entry Tab  (left form | right QPdfView)
    # ─────────────────────────────────────────────────────────────────────────

    def _build_manual_tab(self):
        tab    = QWidget()
        layout = QHBoxLayout(tab)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        layout.addWidget(splitter)

        # LEFT: scrollable form
        left   = QWidget()
        left_v = QVBoxLayout(left)
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        form_w = QWidget(); form_v = QVBoxLayout(form_w); form_v.setSpacing(8)

        p_grp  = QGroupBox("Patient Information")
        p_form = QFormLayout(p_grp)
        p_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        self.m = {}
        # Extract 'dob' and 'dob_type' manually to set up combo box
        self.m_dob_type = QComboBox()
        self.m_dob_type.addItems(["Date of Birth", "Age"])
        self.m_dob_type.setStyleSheet("QComboBox { border: none; font-weight: bold; background: transparent; }")
        self.m_dob_type.currentTextChanged.connect(self._schedule_preview)
        
        self.m["dob"] = QLineEdit()
        self.m["dob"].textChanged.connect(self._schedule_preview)
        p_form.addRow(self.m_dob_type, self.m["dob"])

        for key, label in [
            ("name",            "Patient Name"),
            ("pin",             "PIN / MRN"),
            ("ga",              "Gestational Age"),
            ("sample_id",       "Sample ID"),
            ("collection_date", "Collection Date"),
            ("received_date",   "Receipt Date"),
            ("preg_status",     "Pregnancy Status"),
            ("preg_type",       "Pregnancy Type"),
            ("clinician",       "Referring Clinician"),
            ("hospital",        "Hospital / Clinic"),
            ("specimen",        "Specimen Type"),
            ("indication",      "Clinical Indication"),
        ]:
            self.m[key] = QLineEdit()
            self.m[key].textChanged.connect(self._schedule_preview)
            p_form.addRow(label + ":", self.m[key])
        form_v.addWidget(p_grp)

        z_grp = QGroupBox("Laboratory Z-Scores")
        z_v   = QVBoxLayout(z_grp)
        ff_r  = QHBoxLayout()
        self.m["ff"] = QLineEdit("0.0")
        self.m["ff"].textChanged.connect(self._schedule_preview)
        ff_r.addWidget(QLabel("Fetal DNA Fraction (%):"))
        ff_r.addWidget(self.m["ff"])
        z_v.addLayout(ff_r)
        z_grid = QHBoxLayout()
        c1 = QFormLayout(); c2 = QFormLayout()
        c1.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        c2.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        for i in range(1, 12):
            self.m[f"chr{i}"] = QLineEdit("0.00")
            self.m[f"chr{i}"].textChanged.connect(self._schedule_preview)
            c1.addRow(f"Chr {i}:", self.m[f"chr{i}"])
        for i in range(12, 23):
            self.m[f"chr{i}"] = QLineEdit("0.00")
            self.m[f"chr{i}"].textChanged.connect(self._schedule_preview)
            c2.addRow(f"Chr {i}:", self.m[f"chr{i}"])
        self.m["chrX"] = QLineEdit("0.00")
        self.m["chrX"].textChanged.connect(self._schedule_preview)
        c2.addRow("Chr X:", self.m["chrX"])
        z_grid.addLayout(c1); z_grid.addLayout(c2)
        z_v.addLayout(z_grid)
        form_v.addWidget(z_grp)
        scroll.setWidget(form_w)
        left_v.addWidget(scroll)

        btn_grp = QGroupBox("Actions")
        btn_h   = QHBoxLayout(btn_grp)
        save_draft = QPushButton("Save Draft"); save_draft.clicked.connect(self._save_draft)
        save_draft.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogSaveButton))
        load_draft = QPushButton("Load Draft"); load_draft.clicked.connect(self._load_draft)
        load_draft.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton))
        gen_btn    = QPushButton("Generate Report"); gen_btn.clicked.connect(self._save_manual)
        gen_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogApplyButton))
        gen_btn.setStyleSheet(
            "background-color:#1F497D;color:white;font-weight:bold;padding:7px 18px;")
        clear_btn  = QPushButton("Clear Form"); clear_btn.clicked.connect(self._clear_form)
        clear_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogResetButton))

        for b in [save_draft, load_draft, gen_btn, clear_btn]:
            btn_h.addWidget(b)
        btn_h.addStretch()
        left_v.addWidget(btn_grp)

        splitter.addWidget(left)

        # RIGHT: PDF preview
        right   = QGroupBox("Report Preview")
        right_v = QVBoxLayout(right)
        if HAS_PDF_VIEW:
            self.pdf_doc  = QPdfDocument(self)
            self.pdf_view = QPdfView(self)
            self.pdf_view.setDocument(self.pdf_doc)
            self.pdf_view.setPageMode(QPdfView.PageMode.MultiPage)
        else:
            self.pdf_doc  = None
            self.pdf_view = QTextBrowser()
            self.pdf_view.setHtml(
                "<div style='padding:40px;text-align:center;color:#888'>"
                "<p>Edit any field — preview auto-refreshes after 1 second.</p></div>")
        right_v.addWidget(self.pdf_view)

        self.preview_status = QLabel("")
        self.preview_status.setStyleSheet("color:#777;font-size:11px;padding:2px;")
        right_v.addWidget(self.preview_status)

        refresh = QPushButton("Refresh Preview")
        refresh.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload))
        refresh.clicked.connect(self._start_preview)

        right_v.addWidget(refresh)
        splitter.addWidget(right)
        splitter.setSizes([560, 700])

        # Debounce timer (1 s like PGT-A)
        self._preview_timer = QTimer()
        self._preview_timer.setSingleShot(True)
        self._preview_timer.setInterval(1000)
        self._preview_timer.timeout.connect(self._start_preview)

        return tab

    # ─────────────────────────────────────────────────────────────────────────
    # Batch Upload Tab  (list | inline editor | preview)  — mirrors PGT-A exactly
    # ─────────────────────────────────────────────────────────────────────────

    def _build_bulk_tab(self):
        tab    = QWidget()
        layout = QVBoxLayout(tab)

        # 1. File + output selector
        top = QHBoxLayout()
        file_grp = QGroupBox("1. Select Excel File")
        file_h   = QHBoxLayout(file_grp)
        self.excel_lbl = QLabel("No file selected")
        self.excel_lbl.setStyleSheet("padding:4px;border:1px solid #ccc;background:white;")
        btn_xls = QPushButton("Browse"); btn_xls.clicked.connect(self._browse_excel)
        file_h.addWidget(self.excel_lbl, 1); file_h.addWidget(btn_xls)
        top.addWidget(file_grp, 1)

        out_grp  = QGroupBox("2. Set Output Folder")
        out_h    = QHBoxLayout(out_grp)
        self.bulk_out_edit = QLineEdit()
        self.bulk_out_edit.setPlaceholderText("Select output folder …")
        self.bulk_out_edit.textChanged.connect(self._sync_out_dir_to_global)
        btn_out_br = QPushButton("Browse")
        btn_out_br.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon))
        btn_out_br.clicked.connect(self._browse_out)
        out_h.addWidget(self.bulk_out_edit, 1); out_h.addWidget(btn_out_br)
        layout.addLayout(top)
        layout.addWidget(out_grp)

        # 2. Content area: list | editor | preview
        content = QGroupBox("3. Review and Edit Patients")
        c_h     = QHBoxLayout(content)

        # LEFT: patient list
        left_panel = QWidget()
        left_v     = QVBoxLayout(left_panel)
        left_v.addWidget(QLabel("Patients:"))
        self.batch_search = QLineEdit()
        self.batch_search.setPlaceholderText("Search by name …")
        self.batch_search.setClearButtonEnabled(True)

        self.batch_search.textChanged.connect(self._filter_batch_list)
        left_v.addWidget(self.batch_search)

        self.batch_list = QListWidget()
        self.batch_list.setAlternatingRowColors(True)
        self.batch_list.currentItemChanged.connect(self._on_batch_select)
        left_v.addWidget(self.batch_list)

        draft_row = QHBoxLayout()
        save_draft_btn = QPushButton("Save All Draft"); save_draft_btn.clicked.connect(self._save_bulk_draft)
        load_draft_btn = QPushButton("Load Draft");     load_draft_btn.clicked.connect(self._load_bulk_draft)
        draft_row.addWidget(save_draft_btn); draft_row.addWidget(load_draft_btn)
        left_v.addLayout(draft_row)

        gen_all = QPushButton("Generate All Reports")
        gen_all.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogApplyButton))
        gen_all.setStyleSheet(
            "background:#1F497D;color:white;font-weight:bold;padding:10px;margin-top:4px;")
        gen_all.clicked.connect(self._run_batch)

        left_v.addWidget(gen_all)
        c_h.addWidget(left_panel)

        # RIGHT: split editor | preview
        editor_preview = QSplitter(Qt.Orientation.Horizontal)

        # Editor panel (scrollable, all fields)
        editor_w = QWidget()
        editor_v = QVBoxLayout(editor_w)
        editor_v.addWidget(QLabel("Patient Editor:"))

        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        self.batch_editor_container = QWidget()
        self.batch_editor_layout    = QVBoxLayout(self.batch_editor_container)
        self.batch_editor_placeholder = QLabel(
            "Select a patient from the list on the left")
        self.batch_editor_placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.batch_editor_placeholder.setStyleSheet("color:#888;font-style:italic;padding:40px;")
        self.batch_editor_layout.addWidget(self.batch_editor_placeholder)
        scroll.setWidget(self.batch_editor_container)
        editor_v.addWidget(scroll)
        editor_preview.addWidget(editor_w)

        # Batch preview panel
        prev_grp = QGroupBox("Batch Report Preview")
        prev_v   = QVBoxLayout(prev_grp)
        if HAS_PDF_VIEW:
            self.batch_pdf_doc  = QPdfDocument(self)
            self.batch_pdf_view = QPdfView(self)
            self.batch_pdf_view.setDocument(self.batch_pdf_doc)
            self.batch_pdf_view.setPageMode(QPdfView.PageMode.MultiPage)
        else:
            self.batch_pdf_doc  = None
            self.batch_pdf_view = QTextBrowser()
            self.batch_pdf_view.setHtml(
                "<div style='padding:30px;color:#888;text-align:center'>"
                "<p>Edit fields to see preview here</p></div>")
        prev_v.addWidget(self.batch_pdf_view)
        prev_refresh = QPushButton("Refresh Preview")
        prev_refresh.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_BrowserReload))
        prev_refresh.clicked.connect(self._update_batch_preview)

        prev_v.addWidget(prev_refresh)
        editor_preview.addWidget(prev_grp)
        editor_preview.setSizes([500, 500])

        c_h.addWidget(editor_preview)
        c_h.setStretch(0, 1); c_h.setStretch(1, 4)
        layout.addWidget(content)

        # Batch preview timer
        self._batch_preview_timer = QTimer()
        self._batch_preview_timer.setSingleShot(True)
        self._batch_preview_timer.setInterval(1000)
        self._batch_preview_timer.timeout.connect(self._update_batch_preview)

        # Dict to hold current editor widgets
        self.be = {}   # batch editor field widgets: key→QLineEdit

        return tab

    # ─────────────────────────────────────────────────────────────────────────
    # Compare Tab
    # ─────────────────────────────────────────────────────────────────────────

    def _build_compare_tab(self):
        tab    = QWidget()
        layout = QVBoxLayout(tab)

        mode_g = QGroupBox("Comparison Mode")
        mode_h = QHBoxLayout(mode_g)
        self.dir_mode  = QRadioButton("Directory Mode (Bulk)"); self.dir_mode.setChecked(True)
        self.file_mode = QRadioButton("Individual File Mode")
        self.dir_mode.toggled.connect(self._toggle_compare_mode)
        mode_h.addWidget(self.dir_mode); mode_h.addWidget(self.file_mode); mode_h.addStretch()
        layout.addWidget(mode_g)

        # Directory mode
        self.dir_grp = QGroupBox("Select Report Directories")
        dv = QVBoxLayout(self.dir_grp)
        m_r = QHBoxLayout()
        self.manual_dir_lbl = QLabel("No directory selected")
        self.manual_dir_lbl.setStyleSheet("padding:5px;border:1px solid #ccc;background:white;")
        m_b = QPushButton("Manual Folder"); m_b.clicked.connect(self._browse_manual_dir)
        m_r.addWidget(QLabel("Manual:  ")); m_r.addWidget(self.manual_dir_lbl, 1); m_r.addWidget(m_b)
        dv.addLayout(m_r)
        a_r = QHBoxLayout()
        self.auto_dir_lbl = QLabel("No directory selected")
        self.auto_dir_lbl.setStyleSheet("padding:5px;border:1px solid #ccc;background:white;")
        a_b = QPushButton("Automated Folder"); a_b.clicked.connect(self._browse_auto_dir)
        a_r.addWidget(QLabel("Automated:")); a_r.addWidget(self.auto_dir_lbl, 1); a_r.addWidget(a_b)
        dv.addLayout(a_r)
        layout.addWidget(self.dir_grp)

        # File mode
        self.file_grp = QGroupBox("Select Individual Files")
        fv = QVBoxLayout(self.file_grp)
        self.file_grp.setVisible(False)
        mf_r = QHBoxLayout()
        self.manual_file_lbl = QLabel("No file selected")
        self.manual_file_lbl.setStyleSheet("padding:5px;border:1px solid #ccc;background:white;")
        mf_b = QPushButton("Manual PDF"); mf_b.clicked.connect(self._browse_manual_file)
        mf_r.addWidget(QLabel("Manual:  ")); mf_r.addWidget(self.manual_file_lbl, 1); mf_r.addWidget(mf_b)
        fv.addLayout(mf_r)
        af_r = QHBoxLayout()
        self.auto_file_lbl = QLabel("No file selected")
        self.auto_file_lbl.setStyleSheet("padding:5px;border:1px solid #ccc;background:white;")
        af_b = QPushButton("Automated PDF"); af_b.clicked.connect(self._browse_auto_file)
        af_r.addWidget(QLabel("Automated:")); af_r.addWidget(self.auto_file_lbl, 1); af_r.addWidget(af_b)
        fv.addLayout(af_r)
        layout.addWidget(self.file_grp)

        run_b = QPushButton("Run Comparison")
        run_b.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_CommandLink))
        run_b.setStyleSheet("background:#1F497D;color:white;font-weight:bold;padding:10px;")
        run_b.clicked.connect(self._run_compare)

        layout.addWidget(run_b)

        res_grp = QGroupBox("2. Comparison Results")
        res_v   = QVBoxLayout(res_grp)
        self.cmp_results = QTextBrowser()
        self.cmp_results.setHtml(
            "<p style='color:#999;text-align:center;padding:40px'>"
            "Results will appear here after running comparison …</p>")
        res_v.addWidget(self.cmp_results)
        layout.addWidget(res_grp)

        self._manual_dir = self._auto_dir = None
        curr_dir = os.path.dirname(os.path.abspath(__file__))
        def_m = os.path.join(curr_dir, "Comparison", "Manual")
        def_a = os.path.join(curr_dir, "Comparison", "Automated")
        if os.path.exists(def_m): self.manual_dir_lbl.setText(def_m)
        if os.path.exists(def_a): self.auto_dir_lbl.setText(def_a)
        return tab

    # ─────────────────────────────────────────────────────────────────────────
    # User Guide Tab
    # ─────────────────────────────────────────────────────────────────────────

    def _build_guide_tab(self):
        tab = QWidget(); lay = QVBoxLayout(tab)
        b   = QTextBrowser()
        b.setHtml("""
        <style>
          body{font-family:'Segoe UI',Arial,sans-serif;padding:10px}
          h2{color:#1F497D;border-bottom:2px solid #ccc;padding-bottom:6px}
          h3{color:#444;margin-top:16px}
          ul{margin-left:18px}
          code{background:#f0f0f0;padding:1px 4px;border-radius:2px;font-family:monospace}
        </style>
        <h2>NIPT Report Generator – User Guide</h2>

        <h3>Manual Entry Tab</h3>
        <p>Fill in patient demographics and Z-scores. The <b>right pane</b> shows a live PDF
        preview that auto-refreshes <b>1 second</b> after you stop typing — just like PGT-A.</p>
        <ul>
          <li><b>Save Draft</b> — save all fields to a JSON file for later</li>
          <li><b>Load Draft</b> — restore a saved state (also triggers a preview refresh)</li>
          <li><b>Generate Report</b> — export PDF and/or DOCX to your output folder</li>
          <li><b>Refresh Preview</b> — manually trigger a preview regeneration</li>
        </ul>

        <h3>Batch Upload Tab</h3>
        <p>Select an <code>.xlsx</code> file. All patients load into the <b>left list</b>.
        Click any name to open the <b>full inline editor</b> in the centre — you can edit all
        demographics and Z-scores without leaving this tab. The <b>right pane</b>
        shows a live preview for the selected patient.</p>
        <ul>
          <li><b>Sheet1</b> — demographics (Patient Name, Sample ID, GA …)</li>
          <li><b>Sheet2</b> (optional) — Z-scores, joined on <i>Sample ID</i></li>
          <li><b>Save All Draft</b> — save all edited batch data as JSON</li>
          <li><b>Generate All Reports</b> — export every patient in the batch</li>
        </ul>

        <h3>Report Comparison Tab</h3>
        <p>Compare <b>Manual</b> reports (clinician-signed) vs <b>Automated</b> reports
        (pipeline-generated). In <i>Directory Mode</i> files are matched by filename;
        in <i>File Mode</i> pick one PDF from each source.</p>

        <h3>Export Settings</h3>
        <ul>
          <li><b>PDF / DOCX</b> — choose which formats to export</li>
          <li><b>With Branding Logo</b> — toggle Anderson Diagnostics letterhead</li>
          <li><b>Output Folder</b> — all reports are saved here</li>
        </ul>
        """)
        lay.addWidget(b)
        return tab

    # ─────────────────────────────────────────────────────────────────────────
    # Settings
    # ─────────────────────────────────────────────────────────────────────────

    def _load_settings(self):
        path = self.settings.value("output_dir", os.path.expanduser("~"))
        self.out_dir_edit.setText(path)
        if hasattr(self, 'bulk_out_edit'): self.bulk_out_edit.setText(path)
        self.cb_pdf.setChecked(self.settings.value("do_pdf",   "true") == "true")
        self.cb_docx.setChecked(self.settings.value("do_docx",  "true") == "true")
        self.cb_branding.setChecked(self.settings.value("branding","true") == "true")

    def _save_settings(self):
        self.settings.setValue("output_dir", self.out_dir_edit.text())
        self.settings.setValue("do_pdf",   str(self.cb_pdf.isChecked()).lower())
        self.settings.setValue("do_docx",  str(self.cb_docx.isChecked()).lower())
        self.settings.setValue("branding", str(self.cb_branding.isChecked()).lower())

    def _browse_out(self):
        d = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if d:
            self.out_dir_edit.setText(d)
            if hasattr(self, 'bulk_out_edit'): self.bulk_out_edit.setText(d)
            self._save_settings()

    def _sync_out_dir_to_global(self, text):
        if self.out_dir_edit.text() != text:
            self.out_dir_edit.setText(text)

    # ─────────────────────────────────────────────────────────────────────────
    # Manual live preview (debounced)
    # ─────────────────────────────────────────────────────────────────────────

    def _schedule_preview(self):
        self._preview_timer.start()

    def _collect_manual(self):
        p = {k: v.text() for k, v in self.m.items()
             if k != "ff" and not k.startswith("chr")}
        p["dob_type"] = self.m_dob_type.currentText()
        z = {}
        for i in range(1, 23):
            try:   z[f"chr{i}"] = float(self.m[f"chr{i}"].text() or 0)
            except: z[f"chr{i}"] = 0.0
        try:   z["chrX"]           = float(self.m["chrX"].text()  or 0)
        except: z["chrX"]          = 0.0
        try:   z["fetal_fraction"] = float(self.m["ff"].text()    or 0)
        except: z["fetal_fraction"]= 0.0
        return p, z

    def _start_preview(self):
        if self._preview_worker and self._preview_worker.isRunning():
            return
        p, z = self._collect_manual()
        self.preview_status.setText("Generating preview …")
        self._preview_worker = PreviewWorker(

            p, z, self._preview_tmp,
            show_logo=self.cb_branding.isChecked())
        self._preview_worker.finished.connect(self._on_preview_ready)
        self._preview_worker.error.connect(
            lambda e: self.preview_status.setText(f"❌ {e[:80]}"))
        self._preview_worker.start()

    def _on_preview_ready(self, path):
        self.preview_status.setText("Preview updated")

        if HAS_PDF_VIEW and self.pdf_doc:
            self.pdf_doc.close(); self.pdf_doc.load(path)
            self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitInView)
        else:
            if sys.platform == "win32":    os.startfile(path)
            elif sys.platform == "darwin": subprocess.run(["open", path])
            else:                          subprocess.run(["xdg-open", path])

    # ─────────────────────────────────────────────────────────────────────────
    # Manual actions
    # ─────────────────────────────────────────────────────────────────────────

    def _save_manual(self):
        from nipt_template       import NIPTReportTemplate
        from nipt_docx_generator import NIPTDocxGenerator
        p, z = self._collect_manual()
        out  = self.out_dir_edit.text()
        if not out:
            QMessageBox.warning(self, "No Output Folder", "Set an output folder first."); return
        self._save_settings()
        branding = self.cb_branding.isChecked()
        suffix = "with_logo" if branding else "without_logo"
        base = f"Report_{p.get('name','').replace(' ','_')}_{p.get('sample_id','')}_{suffix}"
        try:
            if self.cb_pdf.isChecked():
                NIPTReportTemplate(os.path.join(out, base+".pdf")).generate(
                    p, z, with_logo=branding)
            if self.cb_docx.isChecked():
                NIPTDocxGenerator(os.path.join(out, base+".docx")).generate(
                    p, z, with_logo=branding)
            QMessageBox.information(self, "Saved", f"Exported: {base}")
        except Exception as e:
            QMessageBox.critical(self, "Export Error", str(e))

    def _save_draft(self):
        p, z = self._collect_manual()
        name = p.get("name","Draft").replace(" ","_")
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Draft", f"Draft_{name}.json", "JSON (*.json)")
        if path:
            try:
                with open(path,"w") as f:
                    json.dump({"patient_details":p,"z_scores":z}, f, indent=4)
                QMessageBox.information(self, "Saved", "Draft saved.")
            except Exception as e: QMessageBox.critical(self, "Error", str(e))

    def _load_draft(self):
        path, _ = QFileDialog.getOpenFileName(self, "Load Draft", "", "JSON (*.json)")
        if not path: return
        try:
            with open(path) as f: d = json.load(f)
            p = d.get("patient_details",{}); z = d.get("z_scores",{})
            for k, v in p.items():
                if k in self.m: self.m[k].setText(str(v))
            if "dob_type" in p:
                idx = self.m_dob_type.findText(str(p["dob_type"]))
                if idx >= 0: self.m_dob_type.setCurrentIndex(idx)
            if "fetal_fraction" in z: self.m["ff"].setText(str(z["fetal_fraction"]))
            for k, v in z.items():
                if k in self.m: self.m[k].setText(str(v))
            self._start_preview()
        except Exception as e: QMessageBox.critical(self, "Error", str(e))

    def _clear_form(self):
        for k, w in self.m.items():
            w.setText("0.00" if k.startswith("chr") or k == "ff" else "")
        self.m_dob_type.setCurrentIndex(0)

    # ─────────────────────────────────────────────────────────────────────────
    # Batch: load Excel
    # ─────────────────────────────────────────────────────────────────────────

    def _browse_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel (*.xlsx *.xls)")
        if path:
            self.excel_lbl.setText(os.path.basename(path))
            self._load_excel(path)

    def _load_excel(self, path):
        import pandas as pd
        try:
            xls   = pd.ExcelFile(path)
            sname = "Sheet1" if "Sheet1" in xls.sheet_names else xls.sheet_names[0]
            df    = pd.read_excel(xls, sname)
            df.columns = [str(c).strip() for c in df.columns]

            if "Sheet2" in xls.sheet_names:
                df_z = pd.read_excel(xls,"Sheet2")
                df_z.columns = [str(c).strip() for c in df_z.columns]
                id1 = next((c for c in ["Sample ID","Sample Name"] if c in df.columns),   None)
                id2 = next((c for c in ["Sample ID","Sample"]      if c in df_z.columns), None)
                if id1 and id2:
                    df = pd.merge(df, df_z, left_on=id1, right_on=id2, how="left")

            raw = df.to_dict("records")
            # Normalise to internal keys
            self.batch_patients = []
            for p in raw:
                ff = p.get("Fetal DNA", p.get("FF", 0))
                try:   ff = float(ff)
                except: ff = 0.0
                ff_pct = ff * 100 if ff < 1 else ff

                bdp = {
                    "name":            str(p.get("Patient Name",         p.get("Sample Name",""))),
                    "pin":             str(p.get("PIN",                   p.get("Sample Name",""))),
                    "dob":             str(p.get("Date of birth/Age",     p.get("Age",""))).replace(" 00:00:00","").split(" /")[0],
                    "ga":              str(p.get("Gestational Age",       "")),
                    "sample_id":       str(p.get("Sample ID",            p.get("Sample Name",""))),
                    "collection_date": str(p.get("Collection date",       p.get("Col Date",""))).replace(" 00:00:00",""),
                    "received_date":   str(p.get("Received date",         p.get("Rec Date",""))).replace(" 00:00:00",""),
                    "preg_status":     str(p.get("Pregnancy status",      p.get("Status",""))),
                    "preg_type":       str(p.get("Pregnancy type",        "")),
                    "clinician":       str(p.get("Referring Clinician",   p.get("Ref Doctor",""))),
                    "hospital":        str(p.get("Hospital",              "")),
                    "indication":      str(p.get("Indication",            "")),
                    "specimen":        str(p.get("Specimen",              "")),
                    "ff":              str(ff_pct),
                }
                for i in range(1, 23):
                    try:   bdp[f"chr{i}"] = f"{float(p.get(f'chr{i}', 0) or 0):.2f}"
                    except: bdp[f"chr{i}"] = "0.00"
                try:   bdp["chrX"] = f"{float(p.get('chrX', 0) or 0):.2f}"
                except: bdp["chrX"] = "0.00"

                self.batch_patients.append(bdp)

            self._rebuild_batch_list()
            self.statusBar().showMessage(
                f"Loaded {len(self.batch_patients)} patients from {os.path.basename(path)}")
        except Exception as e:
            QMessageBox.critical(self, "Load Error", str(e))

    def _rebuild_batch_list(self, filter_text=""):
        self.batch_list.clear()
        for i, p in enumerate(self.batch_patients):
            name = p.get("name","Unknown")
            if filter_text and filter_text.lower() not in name.lower():
                continue
            item = QListWidgetItem(f"{i+1}. {name}")
            item.setData(Qt.ItemDataRole.UserRole, i)
            self.batch_list.addItem(item)

    def _filter_batch_list(self, text):
        self._rebuild_batch_list(filter_text=text)

    # ─────────────────────────────────────────────────────────────────────────
    # Batch: inline editor
    # ─────────────────────────────────────────────────────────────────────────

    def _on_batch_select(self, current, previous):
        if current is None: return
        idx = current.data(Qt.ItemDataRole.UserRole)
        if idx is None: return
        # Save previous edits first
        if self._batch_current >= 0:
            self._save_batch_editor_to_data(self._batch_current)
        self._batch_current = idx
        self._load_batch_editor(idx)

    def _load_batch_editor(self, idx):
        p = self.batch_patients[idx]

        # Clear old editor
        while self.batch_editor_layout.count():
            child = self.batch_editor_layout.takeAt(0)
            if child.widget(): child.widget().deleteLater()
        self.be = {}

        # Patient info group
        p_grp  = QGroupBox(f"Patient Information — {p.get('name','')}")
        p_form = QFormLayout(p_grp)
        p_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        # Explicitly handle dob_type and dob
        dob_layout = QHBoxLayout()
        self.be_dob_type = QComboBox()
        self.be_dob_type.addItems(["Date of Birth", "Age"])
        self.be_dob_type.setStyleSheet("QComboBox { border: none; font-weight: bold; background: transparent; }")
        
        # Set current value if present
        current_type = p.get("dob_type", "Date of Birth")
        idx = self.be_dob_type.findText(current_type)
        if idx >= 0: self.be_dob_type.setCurrentIndex(idx)
        self.be_dob_type.currentTextChanged.connect(self._schedule_batch_preview)
        
        dob_ed = QLineEdit(p.get("dob", ""))
        dob_ed.textChanged.connect(self._schedule_batch_preview)
        self.be["dob"] = dob_ed
        
        p_form.addRow(self.be_dob_type, dob_ed)

        for key, label in [
            ("name",            "Patient Name"),
            ("pin",             "PIN / MRN"),
            ("ga",              "Gestational Age"),
            ("sample_id",       "Sample ID"),
            ("collection_date", "Collection Date"),
            ("received_date",   "Receipt Date"),
            ("preg_status",     "Pregnancy Status"),
            ("preg_type",       "Pregnancy Type"),
            ("clinician",       "Referring Clinician"),
            ("hospital",        "Hospital / Clinic"),
            ("specimen",        "Specimen Type"),
            ("indication",      "Clinical Indication"),
        ]:
            ed = QLineEdit(p.get(key, ""))
            ed.textChanged.connect(self._schedule_batch_preview)
            self.be[key] = ed
            p_form.addRow(label + ":", ed)
        self.batch_editor_layout.addWidget(p_grp)

        # Z-scores group
        z_grp = QGroupBox("Laboratory Z-Scores")
        z_v   = QVBoxLayout(z_grp)
        ff_r  = QHBoxLayout()
        ff_ed = QLineEdit(p.get("ff","0.0"))
        ff_ed.textChanged.connect(self._schedule_batch_preview)
        self.be["ff"] = ff_ed
        ff_r.addWidget(QLabel("Fetal DNA Fraction (%):"))
        ff_r.addWidget(ff_ed)
        z_v.addLayout(ff_r)

        z_grid = QHBoxLayout()
        zc1 = QFormLayout(); zc2 = QFormLayout()
        zc1.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        zc2.setLabelAlignment(Qt.AlignmentFlag.AlignRight)
        for i in range(1, 12):
            e = QLineEdit(p.get(f"chr{i}","0.00"))
            e.textChanged.connect(self._schedule_batch_preview)
            self.be[f"chr{i}"] = e
            zc1.addRow(f"Chr {i}:", e)
        for i in range(12, 23):
            e = QLineEdit(p.get(f"chr{i}","0.00"))
            e.textChanged.connect(self._schedule_batch_preview)
            self.be[f"chr{i}"] = e
            zc2.addRow(f"Chr {i}:", e)
        exw = QLineEdit(p.get("chrX","0.00"))
        exw.textChanged.connect(self._schedule_batch_preview)
        self.be["chrX"] = exw
        zc2.addRow("Chr X:", exw)
        z_grid.addLayout(zc1); z_grid.addLayout(zc2)
        z_v.addLayout(z_grid)
        self.batch_editor_layout.addWidget(z_grp)

        # Per-patient actions row
        p_act_h = QHBoxLayout()
        draft_btn = QPushButton("Save Individual Draft")
        draft_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogSaveButton))
        draft_btn.clicked.connect(lambda: self._save_individual_batch_draft(idx))
        
        gen_btn = QPushButton("Generate Report for This Patient")
        gen_btn.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogApplyButton))
        gen_btn.setStyleSheet("background:#1F497D;color:white;font-weight:bold;padding:8px;")
        gen_btn.clicked.connect(lambda: self._generate_single_batch(idx))
        
        p_act_h.addWidget(draft_btn); p_act_h.addWidget(gen_btn)
        self.batch_editor_layout.addLayout(p_act_h)
        self.batch_editor_layout.addStretch()

        self._update_batch_preview()

    def _save_batch_editor_to_data(self, idx):
        if not self.be: return
        for key, ed in self.be.items():
            self.batch_patients[idx][key] = ed.text()
        if hasattr(self, 'be_dob_type'):
            self.batch_patients[idx]["dob_type"] = self.be_dob_type.currentText()

    def _schedule_batch_preview(self):
        self._batch_preview_timer.start()

    def _collect_batch_editor(self):
        if not self.be: return None, None
        p = {k: v.text() for k, v in self.be.items()
             if k != "ff" and not k.startswith("chr")}
        if hasattr(self, 'be_dob_type'):
            p["dob_type"] = self.be_dob_type.currentText()
        else:
            p["dob_type"] = "Date of Birth"
        z = {}
        for i in range(1, 23):
            key = f"chr{i}"
            try:   z[key] = float(self.be[key].text() or 0)
            except: z[key] = 0.0
        try:   z["chrX"]           = float(self.be["chrX"].text() or 0)
        except: z["chrX"]          = 0.0
        try:   z["fetal_fraction"] = float(self.be["ff"].text()   or 0)
        except: z["fetal_fraction"]= 0.0
        return p, z

    def _update_batch_preview(self):
        p, z = self._collect_batch_editor()
        if p is None: return
        worker = PreviewWorker(p, z, self._batch_preview_tmp,
                               show_logo=self.cb_branding.isChecked())
        def _on_batch_preview_ready(path):
            if HAS_PDF_VIEW and self.batch_pdf_doc:
                self.batch_pdf_doc.close(); self.batch_pdf_doc.load(path)
                self.batch_pdf_view.setZoomMode(QPdfView.ZoomMode.FitInView)
        worker.finished.connect(_on_batch_preview_ready)
        worker.start()
        self._batch_preview_worker = worker  # keep reference alive

    def _save_individual_batch_draft(self, idx):
        self._save_batch_editor_to_data(idx)
        p = self.batch_patients[idx]
        name = p.get("name","Draft").replace(" ","_")
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Individual Draft", f"Draft_{name}.json", "JSON (*.json)")
        if path:
            try:
                # Reconstruct z_scores like _collect_batch_editor
                z = {f"chr{i}": float(p.get(f"chr{i}", 0) or 0) for i in range(1, 23)}
                z["chrX"] = float(p.get("chrX", 0) or 0)
                z["fetal_fraction"] = float(p.get("ff", 0) or 0)
                
                patient_info = {k: v for k, v in p.items() if k not in z and k != "ff"}
                
                with open(path, "w") as f:
                    json.dump({"patient_details": patient_info, "z_scores": z}, f, indent=4)
                QMessageBox.information(self, "Saved", "Individual draft saved.")
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def _generate_single_batch(self, idx):
        from nipt_template       import NIPTReportTemplate
        from nipt_docx_generator import NIPTDocxGenerator
        self._save_batch_editor_to_data(idx)
        p_d = self.batch_patients[idx]
        out = self.out_dir_edit.text()
        if not out:
            QMessageBox.warning(self, "No Output", "Set an output folder first."); return
        try:
            z = {f"chr{i}": float(p_d.get(f"chr{i}",0) or 0) for i in range(1,23)}
            z["chrX"]           = float(p_d.get("chrX",0) or 0)
            z["fetal_fraction"] = float(p_d.get("ff",0)   or 0)
            p_info = {k: p_d.get(k,"") for k in [
                "name","pin","dob","dob_type","ga","sample_id","collection_date",
                "received_date","preg_status","preg_type","clinician",
                "hospital","indication","specimen"]}
            branding = self.cb_branding.isChecked()
            suffix = "with_logo" if branding else "without_logo"
            base = f"Report_{p_info['name'].replace(' ','_')}_{p_info['sample_id']}_{suffix}"
            if self.cb_pdf.isChecked():
                NIPTReportTemplate(os.path.join(out, base+".pdf")).generate(
                    p_info, z, with_logo=branding)
            if self.cb_docx.isChecked():
                NIPTDocxGenerator(os.path.join(out, base+".docx")).generate(
                    p_info, z, with_logo=branding)
            QMessageBox.information(self, "Done", f"Saved {base}")
        except Exception as e: QMessageBox.critical(self, "Error", str(e))

    def _save_bulk_draft(self):
        if self._batch_current >= 0:
            self._save_batch_editor_to_data(self._batch_current)
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Bulk Draft", "BulkDraft.json", "JSON (*.json)")
        if path:
            with open(path,"w") as f: json.dump(self.batch_patients, f, indent=4)
            QMessageBox.information(self, "Saved", "Bulk draft saved.")

    def _load_bulk_draft(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Load Bulk Draft", "", "JSON (*.json)")
        if not path: return
        with open(path) as f: self.batch_patients = json.load(f)
        self._rebuild_batch_list()

    def _run_batch(self):
        if self._batch_current >= 0:
            self._save_batch_editor_to_data(self._batch_current)
        if not self.batch_patients:
            QMessageBox.warning(self, "No Data", "Load an Excel file first."); return
        out = self.out_dir_edit.text()
        if not out:
            QMessageBox.warning(self, "No Output", "Set an output folder."); return
        self._save_settings()
        self.progress_bar.setVisible(True); self.progress_bar.setValue(0)
        self.batch_worker = BatchWorker(
            self.batch_patients, out,
            self.cb_pdf.isChecked(), self.cb_docx.isChecked(),
            self.cb_branding.isChecked())
        self.batch_worker.progress.connect(
            lambda v, s: (self.progress_bar.setValue(v), self.statusBar().showMessage(s)))
        self.batch_worker.finished.connect(
            lambda c, t: (self.progress_bar.setVisible(False),
                          QMessageBox.information(self, "Done", f"Processed {c}/{t} cases.")))
        self.batch_worker.error.connect(
            lambda s: QMessageBox.warning(self, "Warning", s))
        self.batch_worker.start()

    # ─────────────────────────────────────────────────────────────────────────
    # Compare Logic (Demographics Based)
    # ─────────────────────────────────────────────────────────────────────────

    def _toggle_compare_mode(self, checked):
        self.dir_grp.setVisible(checked); self.file_grp.setVisible(not checked)

    def _browse_manual_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Manual Reports Folder")
        if d: self.manual_dir_lbl.setText(d); self._manual_dir = d

    def _browse_auto_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Automated Reports Folder")
        if d: self.auto_dir_lbl.setText(d); self._auto_dir = d

    def _browse_manual_file(self):
        p, _ = QFileDialog.getOpenFileName(self, "Manual PDF", "", "PDF (*.pdf)")
        if p: self.manual_file_lbl.setText(os.path.basename(p)); self._manual_file = p

    def _browse_auto_file(self):
        p, _ = QFileDialog.getOpenFileName(self, "Automated PDF", "", "PDF (*.pdf)")
        if p: self.auto_file_lbl.setText(os.path.basename(p)); self._auto_file = p

    def _extract_pdf_demographics(self, path):
        """Helper to extract key fields from NIPT PDF using PyPDF2."""
        import PyPDF2
        fields = {
            "Patient Name": "Not found",
            "PIN":          "Not found",
            "Sample ID":    "Not found",
            "GA":           "Not found",
            "Result":       "Not found",
            "FF %":         "Not found"
        }
        try:
            with open(path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                text = ""
                for page in reader.pages:
                    text += page.extract_text() + "\n"
                
                # Simple extraction based on nipt_template labels
                import re
                patterns = {
                    "Patient Name": r"Patient name\s*:\s*(.*?)(?=\s*Specimen|$)",
                    "PIN":          r"PIN\s*:\s*(.*?)(?=\s|$|\n)",
                    "Sample ID":    r"Sample Number\s*:\s*(.*?)(?=\s|$|\n)",
                    "GA":           r"Gestational Age\s*:\s*(.*?)(?=\s|$|\n)",
                    "Result":       r"Screen result\s*:\s*(Low risk|High risk|Inconclusive)",
                    "FF %":         r"Fetal fraction\s*:\s*([\d\.]+)%"
                }
                for k, p in patterns.items():
                    m = re.search(p, text, re.IGNORECASE)
                    if m: fields[k] = m.group(1).strip()
        except Exception as e:
            print(f"Extraction error: {e}")
        return fields

    def _run_compare(self):
        rows = []
        if self.dir_mode.isChecked():
            m_dir = self.manual_dir_lbl.text()
            a_dir = self.auto_dir_lbl.text()
            if not os.path.isdir(m_dir) or not os.path.isdir(a_dir):
                QMessageBox.warning(self, "Missing Folders",
                    "Select both a Manual and an Automated folder."); return
            mf = {f for f in os.listdir(m_dir) if f.lower().endswith(".pdf")}
            af = {f for f in os.listdir(a_dir) if f.lower().endswith(".pdf")}
            only_m = mf - af; only_a = af - mf; matched = mf & af
            rows.append(f"<h2 style='color:#1F497D'>Directory Comparison Report</h2>"
                f"<p><b>Manual:</b> {m_dir}<br><b>Automated:</b> {a_dir}</p>"
                f"<p>Total: {len(mf|af)} files | Matched: {len(matched)} | "
                f"Manual only: {len(only_m)} | Auto only: {len(only_a)}</p><hr>")
            if only_m:
                rows.append("<h3 style='color:#d97706'>Only in Manual:</h3><ul>")
                for f in sorted(only_m): rows.append(f"<li>{f}</li>")
                rows.append("</ul>")
            if only_a:
                rows.append("<h3 style='color:#7c3aed'>Only in Automated:</h3><ul>")
                for f in sorted(only_a): rows.append(f"<li>{f}</li>")
                rows.append("</ul>")
            if matched:
                rows.append(f"<h3 style='color:#16a34a'>Matched Pairs ({len(matched)}):</h3><ul>")
                for f in sorted(matched):
                    # We could do bulk demographics here too, but for speed just names for now
                    rows.append(f"<li><b>{f}</b> – Matched by name</li>")
                rows.append("</ul>")
        else:
            mf_path = getattr(self, "_manual_file", None)
            af_path = getattr(self, "_auto_file",   None)
            if not mf_path or not af_path:
                QMessageBox.warning(self, "Missing Files",
                    "Select both a Manual and an Automated PDF."); return
            
            m_fields = self._extract_pdf_demographics(mf_path)
            a_fields = self._extract_pdf_demographics(af_path)

            rows = [f"<h2 style='color:#1F497D'>Demographic Detail Comparison</h2>",
                    "<table border='1' cellpadding='6' "
                    "style='border-collapse:collapse;width:100%'>",
                    "<tr style='background:#f0f0f0'><th>Field</th>"
                    "<th>Manual PDF</th><th>Automated PDF</th></tr>"]
            
            for field in m_fields.keys():
                mv = m_fields[field]
                av = a_fields[field]
                color = "#000" if mv == av else "#dc2626"
                rows.append(f"<tr><td><b>{field}</b></td>"
                            f"<td>{mv}</td>"
                            f"<td style='color:{color}'>{av}</td></tr>")
            
            rows.append("</table>")
            
            # File size as secondary info
            ms = os.path.getsize(mf_path); as_ = os.path.getsize(af_path)
            rows.append(f"<p style='color:#777;font-size:11px;margin-top:10px'>"
                        f"Technical: Manual ({ms//1024}KB) vs Automated ({as_//1024}KB)</p>")
            
        self.cmp_results.setHtml("".join(rows))
        self.statusBar().showMessage("Comparison complete.")



# =============================================================================
if __name__ == "__main__":
    app    = QApplication(sys.argv)
    window = NIPTApp()
    window.show()
    sys.exit(app.exec())
