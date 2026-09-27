"""
MMU App-2 Main Window UI
Provides three dedicated input fields:
1. Question Screenshot & Gemini API Extractor
2. Diagram / Visual Image for Dataset
3. Answer (Text / Quick Options)
Generates Excel (with embedded thumbnails), CSV, JSONL, and images/ directory.
"""

import os
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Optional

from PIL import Image, ImageGrab, ImageTk

try:
    from core.dataset_manager import DatasetManager
    from core.gemini_client import GeminiClient
    from core.models import sanitize_source_name
except ImportError:
    from ..core.dataset_manager import DatasetManager
    from ..core.gemini_client import GeminiClient
    from ..core.models import sanitize_source_name


class App2MainWindow:
    """
    Main user interface for MMU App-2.
    """

    def __init__(self, root: tk.Tk, manager: DatasetManager, gemini_client: GeminiClient, prompt_startup: bool = True):
        self.root = root
        self.manager = manager
        self.gemini = gemini_client

        # State for active question being created
        self.question_screenshot: Optional[Image.Image] = None
        self.diagram_image: Optional[Image.Image] = None
        self._qs_photo: Optional[ImageTk.PhotoImage] = None
        self._diag_photo: Optional[ImageTk.PhotoImage] = None

        self.auto_extract_var = tk.BooleanVar(value=True)
        self.auto_furnish_text_var = tk.BooleanVar(value=True)
        self.q_mode_var = tk.StringVar(value="both")
        self.is_extracting: bool = False
        self._status_clear_job: Optional[str] = None

        # Fonts and styling
        self._setup_fonts()
        self._setup_styles()

        # Window setup
        self.root.title("MMU Dataset Builder v2 — Gemini Multimodal MCQ Creator")
        self.root.geometry("1180x820")
        self.root.minsize(700, 500)
        self.root.resizable(True, True)

        # Build UI layout
        self._create_widgets()
        self._bind_shortcuts()
        self._update_stats_display()

        if prompt_startup:
            self.root.after(150, self.prompt_startup_choice)

    def _setup_fonts(self) -> None:
        """Determines best fonts for Bengali and English UI."""
        import tkinter.font as tkfont
        available = tkfont.families()

        if "Nirmala UI" in available:
            self.bangla_font = "Nirmala UI"
        elif "Vrinda" in available:
            self.bangla_font = "Vrinda"
        elif "Kalpurush" in available:
            self.bangla_font = "Kalpurush"
        elif "Segoe UI" in available:
            self.bangla_font = "Segoe UI"
        else:
            self.bangla_font = "Arial"

        self.ui_font = "Segoe UI" if "Segoe UI" in available else "Arial"

        self.font_h1 = (self.ui_font, 13, "bold")
        self.font_h2 = (self.ui_font, 11, "bold")
        self.font_body = (self.ui_font, 9)
        self.font_bold = (self.ui_font, 9, "bold")
        self.font_question = (self.bangla_font, 12, "bold")
        self.font_answer = (self.bangla_font, 11, "bold")

    def _setup_styles(self) -> None:
        """Configures visual color tokens and ttk styles."""
        self.color_bg = "#F8FAFC"
        self.color_card = "#FFFFFF"
        self.color_border = "#E2E8F0"
        self.color_primary = "#2563EB"
        self.color_accent = "#7C3AED"
        self.color_text = "#0F172A"
        self.color_text_muted = "#64748B"

        self.root.configure(bg=self.color_bg)
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

    def _create_widgets(self) -> None:
        """Constructs all visual components."""
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        # =============================================================
        # 1. TOP APP BAR (RESPONSIVE)
        # =============================================================
        self.header_bar = tk.Frame(self.root, bg="#1E293B", padx=14, pady=6)
        self.header_bar.grid(row=0, column=0, sticky="ew")

        # Title & Subtitle container
        self.tbox = tk.Frame(self.header_bar, bg="#1E293B")

        lbl_title = tk.Label(
            self.tbox,
            text="MMU Dataset Builder v2",
            font=self.font_h1,
            fg="#FFFFFF",
            bg="#1E293B",
        )
        lbl_title.pack(anchor="w")

        init_ds_text = f"Dataset: {self.manager.excel_file.name}" if (self.manager.excel_file and self.manager.excel_file.is_file()) else "No dataset loaded"
        self.lbl_dataset_info = tk.Label(
            self.tbox,
            text=init_ds_text,
            font=(self.ui_font, 8),
            fg="#94A3B8",
            bg="#1E293B",
        )
        self.lbl_dataset_info.pack(anchor="w")

        # Top Controls Container
        self.ctrl_box = tk.Frame(self.header_bar, bg="#1E293B")

        self.btn_api_key = tk.Button(
            self.ctrl_box,
            text="Gemini API Key",
            font=self.font_bold,
            bg="#334155",
            fg="#FFFFFF",
            activebackground="#475569",
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            padx=8,
            pady=3,
            command=self.configure_api_key_dialog,
        )

        self.btn_export = tk.Button(
            self.ctrl_box,
            text="Export (ZIP)",
            font=self.font_bold,
            bg="#059669",
            fg="#FFFFFF",
            activebackground="#047857",
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            padx=8,
            pady=3,
            command=self.export_dataset_zip,
        )

        self.btn_new = tk.Button(
            self.ctrl_box,
            text="+ New Dataset",
            font=self.font_bold,
            bg="#2563EB",
            fg="#FFFFFF",
            activebackground="#1D4ED8",
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            padx=8,
            pady=3,
            command=self.create_new_dataset_dialog,
        )

        self.btn_import = tk.Button(
            self.ctrl_box,
            text="Import Dataset",
            font=self.font_bold,
            bg="#3B82F6",
            fg="#FFFFFF",
            activebackground="#2563EB",
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            padx=8,
            pady=3,
            command=self.import_dataset_file,
        )

        self.btn_img_dir = tk.Button(
            self.ctrl_box,
            text="Open Images",
            font=self.font_body,
            bg="#334155",
            fg="#E2E8F0",
            relief="flat",
            cursor="hand2",
            padx=7,
            pady=3,
            command=self.open_images_folder,
        )

        self.btn_folder = tk.Button(
            self.ctrl_box,
            text="Open Folder",
            font=self.font_body,
            bg="#334155",
            fg="#E2E8F0",
            relief="flat",
            cursor="hand2",
            padx=7,
            pady=3,
            command=self.open_dataset_folder,
        )

        self._all_header_buttons = [
            self.btn_new,
            self.btn_import,
            self.btn_export,
            self.btn_api_key,
            self.btn_img_dir,
            self.btn_folder,
        ]

        self._header_layout_mode: Optional[str] = None
        self._apply_header_layout(1180)
        self.root.bind("<Configure>", self._on_window_configure)

        # =============================================================
        # 2. METADATA CONFIGURATION STRIP
        # =============================================================
        meta_strip = tk.Frame(self.root, bg="#F1F5F9", padx=16, pady=8, highlightbackground="#CBD5E1", highlightthickness=1)
        meta_strip.grid(row=1, column=0, sticky="ew")

        # Source / Prefix
        tk.Label(meta_strip, text="Source Name:", font=self.font_bold, bg="#F1F5F9", fg=self.color_text).pack(side="left", padx=(0, 4))
        self.var_source = tk.StringVar(value=self.manager.current_source_name)
        self.ent_source = tk.Entry(meta_strip, textvariable=self.var_source, width=16, font=self.font_body)
        self.ent_source.pack(side="left", padx=(0, 16))
        self.var_source.trace_add("write", lambda *args: self._update_id_preview())

        # Next Generated ID Badge
        next_idx = len(self.manager.entries) + 1
        raw_src = self.manager.current_source_name.strip()
        init_id_badge = f"Next ID: {sanitize_source_name(raw_src)}_{next_idx:03d}" if raw_src else "Next ID: (set source)"
        self.lbl_id_badge = tk.Label(
            meta_strip,
            text=init_id_badge,
            font=(self.ui_font, 9, "bold"),
            bg="#DBEAFE",
            fg="#1E40AF",
            padx=10,
            pady=3,
        )
        self.lbl_id_badge.pack(side="left", padx=(0, 16))

        # Dataset count badge (right side)
        self.lbl_count_badge = tk.Label(
            meta_strip,
            text=f"Questions in Dataset: {self.manager.total_count}",
            font=self.font_bold,
            bg="#F1F5F9",
            fg="#475569",
        )
        self.lbl_count_badge.pack(side="right")

        # =============================================================
        # 3. MAIN WORKSPACE (THREE INPUT FIELDS)
        # =============================================================
        body_container = tk.Frame(self.root, bg=self.color_bg, padx=14, pady=10)
        body_container.grid(row=2, column=0, sticky="nsew")

        # 2 Columns: Left Column (Field 1: Question Screenshot & Text)
        #            Right Column (Field 2: Diagram Image & Field 3: Answer)
        body_container.columnconfigure(0, weight=5, minsize=300)
        body_container.columnconfigure(1, weight=5, minsize=300)
        body_container.rowconfigure(0, weight=1)

        # -------------------------------------------------------------
        # LEFT CARD: FIELD 1 — QUESTION SCREENSHOT & RAW TEXT GEMINI FURNISHER
        # -------------------------------------------------------------
        card_q = tk.Frame(body_container, bg=self.color_card, highlightbackground=self.color_border, highlightthickness=1, padx=14, pady=12)
        card_q.grid(row=0, column=0, sticky="nsew", padx=(0, 7))
        card_q.columnconfigure(0, weight=1)
        card_q.rowconfigure(4, weight=1)  # Final text expands

        # Title Row + Mode Selector
        q_head = tk.Frame(card_q, bg=self.color_card)
        q_head.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        tk.Label(
            q_head,
            text="Question Input & Gemini Furnisher",
            font=self.font_h2,
            bg=self.color_card,
            fg=self.color_text,
        ).pack(side="left")

        # Mode Tab Buttons
        mode_bar = tk.Frame(q_head, bg=self.color_card)
        mode_bar.pack(side="right")

        self.btn_tab_screenshot = tk.Button(
            mode_bar,
            text="Screenshot",
            font=self.font_bold,
            padx=8,
            pady=2,
            relief="flat",
            cursor="hand2",
            command=lambda: self.set_q_mode("screenshot"),
        )
        self.btn_tab_screenshot.pack(side="left", padx=2)

        self.btn_tab_text = tk.Button(
            mode_bar,
            text="Raw / Bijoy Text",
            font=self.font_bold,
            padx=8,
            pady=2,
            relief="flat",
            cursor="hand2",
            command=lambda: self.set_q_mode("text"),
        )
        self.btn_tab_text.pack(side="left", padx=2)

        self.btn_tab_both = tk.Button(
            mode_bar,
            text="Both",
            font=self.font_bold,
            padx=8,
            pady=2,
            relief="flat",
            cursor="hand2",
            command=lambda: self.set_q_mode("both"),
        )
        self.btn_tab_both.pack(side="left", padx=2)

        # Container for Input Zones
        self.frame_input_zone = tk.Frame(card_q, bg=self.color_card)
        self.frame_input_zone.grid(row=1, column=0, sticky="ew", pady=(0, 4))

        # --- A. FRAME FOR SCREENSHOT OCR ---
        self.frame_mode_screenshot = tk.Frame(self.frame_input_zone, bg=self.color_card)

        ss_head = tk.Frame(self.frame_mode_screenshot, bg=self.color_card)
        ss_head.pack(fill="x", pady=(0, 4))
        tk.Label(
            ss_head,
            text="Question Screenshot:",
            font=self.font_bold,
            fg=self.color_text_muted,
            bg=self.color_card,
        ).pack(side="left")

        tk.Button(
            ss_head,
            text="Browse",
            font=self.font_body,
            bg="#F1F5F9",
            fg="#334155",
            relief="flat",
            cursor="hand2",
            padx=6,
            pady=1,
            command=self.browse_question_screenshot,
        ).pack(side="right")

        tk.Button(
            ss_head,
            text="Paste (Ctrl+V)",
            font=self.font_body,
            bg="#EFF6FF",
            fg="#1D4ED8",
            relief="flat",
            cursor="hand2",
            padx=6,
            pady=1,
            command=self.paste_question_screenshot,
        ).pack(side="right", padx=(0, 4))

        self.canvas_qs_frame = tk.Frame(self.frame_mode_screenshot, bg="#F8FAFC", highlightbackground="#CBD5E1", highlightthickness=1, height=110)
        self.canvas_qs_frame.pack(fill="x", expand=False, pady=(0, 4))
        self.canvas_qs_frame.pack_propagate(False)

        self.canvas_qs = tk.Canvas(self.canvas_qs_frame, bg="#F8FAFC", highlightthickness=0, cursor="hand2")
        self.canvas_qs.pack(fill="both", expand=True)
        self.canvas_qs.bind("<Button-1>", lambda e: self.paste_question_screenshot())
        self._render_qs_placeholder()

        extract_bar = tk.Frame(self.frame_mode_screenshot, bg=self.color_card)
        extract_bar.pack(fill="x", pady=(0, 4))

        self.btn_extract = tk.Button(
            extract_bar,
            text="Extract Screenshot with Gemini",
            font=self.font_bold,
            bg="#7C3AED",
            fg="#FFFFFF",
            activebackground="#6D28D9",
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            padx=10,
            pady=3,
            command=self.start_gemini_extraction,
        )
        self.btn_extract.pack(side="left")

        chk_auto = tk.Checkbutton(
            extract_bar,
            text="Auto-extract on paste",
            variable=self.auto_extract_var,
            font=self.font_body,
            bg=self.color_card,
            activebackground=self.color_card,
            cursor="hand2",
        )
        chk_auto.pack(side="left", padx=(10, 0))

        # --- B. FRAME FOR RAW / BIJOY / GARBLED TEXT ---
        self.frame_mode_text = tk.Frame(self.frame_input_zone, bg=self.color_card)

        txt_head = tk.Frame(self.frame_mode_text, bg=self.color_card)
        txt_head.pack(fill="x", pady=(0, 4))

        tk.Label(
            txt_head,
            text="Raw / Bijoy / Garbled Question Text:",
            font=self.font_bold,
            fg=self.color_text_muted,
            bg=self.color_card,
        ).pack(side="left")

        tk.Button(
            txt_head,
            text="Clear",
            font=self.font_body,
            bg="#FEF2F2",
            fg="#B91C1C",
            relief="flat",
            cursor="hand2",
            padx=6,
            pady=1,
            command=self.clear_raw_text,
        ).pack(side="right")

        tk.Button(
            txt_head,
            text="Paste (Ctrl+V)",
            font=self.font_body,
            bg="#EFF6FF",
            fg="#1D4ED8",
            relief="flat",
            cursor="hand2",
            padx=6,
            pady=1,
            command=self.paste_raw_text,
        ).pack(side="right", padx=(0, 4))

        raw_box_frame = tk.Frame(self.frame_mode_text, bg="#FFFFFF", highlightbackground="#CBD5E1", highlightthickness=1)
        raw_box_frame.pack(fill="x", expand=False, pady=(0, 4))

        self.txt_raw_question = tk.Text(
            raw_box_frame,
            font=self.font_body,
            fg="#0F172A",
            bg="#FFFFFF",
            wrap="word",
            relief="flat",
            padx=8,
            pady=6,
            height=4,
            undo=True,
            insertbackground="#0F172A",
        )
        raw_scroll = ttk.Scrollbar(raw_box_frame, orient="vertical", command=self.txt_raw_question.yview)
        self.txt_raw_question.configure(yscrollcommand=raw_scroll.set)
        self.txt_raw_question.pack(side="left", fill="both", expand=True)
        raw_scroll.pack(side="right", fill="y")
        self.txt_raw_question.bind("<Control-Return>", lambda e: (self.start_text_furnishing(), "break")[1])
        self.txt_raw_question.bind("<Control-KP_Enter>", lambda e: (self.start_text_furnishing(), "break")[1])

        raw_action_bar = tk.Frame(self.frame_mode_text, bg=self.color_card)
        raw_action_bar.pack(fill="x", pady=(0, 4))

        self.btn_furnish_text = tk.Button(
            raw_action_bar,
            text="Furnish Text with Gemini",
            font=self.font_bold,
            bg="#7C3AED",
            fg="#FFFFFF",
            activebackground="#6D28D9",
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            padx=10,
            pady=3,
            command=self.start_text_furnishing,
        )
        self.btn_furnish_text.pack(side="left")

        chk_auto_furnish = tk.Checkbutton(
            raw_action_bar,
            text="Auto-furnish on paste",
            variable=self.auto_furnish_text_var,
            font=self.font_body,
            bg=self.color_card,
            activebackground=self.color_card,
            cursor="hand2",
        )
        chk_auto_furnish.pack(side="left", padx=(10, 0))

        tk.Label(
            raw_action_bar,
            text="(Ctrl+Enter to furnish)",
            font=(self.ui_font, 8),
            fg="#64748B",
            bg=self.color_card,
        ).pack(side="left", padx=(6, 0))

        # --- C. SHARED STATUS BAR ---
        status_bar = tk.Frame(card_q, bg=self.color_card)
        status_bar.grid(row=2, column=0, sticky="ew", pady=(2, 2))

        self.lbl_extract_status = tk.Label(
            status_bar,
            text="",
            font=(self.ui_font, 8, "bold"),
            fg="#059669",
            bg=self.color_card,
        )
        self.lbl_extract_status.pack(side="left")

        # --- D. FINAL FURNISHED QUESTION STRING (OUTPUT) ---
        lbl_qtext_header = tk.Label(
            card_q,
            text="Final Furnished Question String (Editable, Ctrl+Z to Undo):",
            font=self.font_bold,
            fg=self.color_text_muted,
            bg=self.color_card,
        )
        lbl_qtext_header.grid(row=3, column=0, sticky="nw", pady=(2, 2))

        # Extracted Question Text Area
        q_frame = tk.Frame(card_q, bg="#FFFFFF", highlightbackground="#CBD5E1", highlightthickness=1)
        q_frame.grid(row=4, column=0, sticky="nsew", pady=(0, 4))
        q_frame.columnconfigure(0, weight=1)
        q_frame.rowconfigure(0, weight=1)

        self.txt_question = tk.Text(
            q_frame,
            font=self.font_question,
            fg="#0F172A",
            bg="#FFFFFF",
            wrap="word",
            relief="flat",
            padx=10,
            pady=8,
            insertbackground="#0F172A",
            selectbackground="#DBEAFE",
            height=7,
            undo=True,
            maxundo=100,
            autoseparators=True,
        )
        q_scroll = ttk.Scrollbar(q_frame, orient="vertical", command=self.txt_question.yview)
        self.txt_question.configure(yscrollcommand=q_scroll.set)
        self.txt_question.grid(row=0, column=0, sticky="nsew")
        q_scroll.grid(row=0, column=1, sticky="ns")

        # Explicit undo/redo bindings
        self.txt_question.bind("<Control-z>", self._on_question_undo)
        self.txt_question.bind("<Control-Z>", self._on_question_undo)
        self.txt_question.bind("<Control-y>", self._on_question_redo)
        self.txt_question.bind("<Control-Y>", self._on_question_redo)

        # -------------------------------------------------------------
        # RIGHT CARD: FIELD 2 — DIAGRAM IMAGE & FIELD 3 — ANSWER
        # -------------------------------------------------------------
        card_right = tk.Frame(body_container, bg=self.color_card, highlightbackground=self.color_border, highlightthickness=1, padx=14, pady=12)
        card_right.grid(row=0, column=1, sticky="nsew", padx=(7, 0))
        card_right.columnconfigure(0, weight=1)
        card_right.rowconfigure(1, weight=1)  # Diagram canvas expands

        # Header for Diagram / Visual Image
        diag_head = tk.Frame(card_right, bg=self.color_card)
        diag_head.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        tk.Label(
            diag_head,
            text="Diagram / Visual Image",
            font=self.font_h2,
            bg=self.color_card,
            fg=self.color_text,
        ).pack(side="left")

        btn_clear_diag = tk.Button(
            diag_head,
            text="Clear",
            font=self.font_body,
            bg="#FEF2F2",
            fg="#B91C1C",
            relief="flat",
            cursor="hand2",
            padx=6,
            pady=1,
            command=self.clear_diagram_image,
        )
        btn_clear_diag.pack(side="right", padx=(3, 0))

        btn_browse_diag = tk.Button(
            diag_head,
            text="Browse",
            font=self.font_body,
            bg="#F1F5F9",
            fg="#334155",
            relief="flat",
            cursor="hand2",
            padx=6,
            pady=1,
            command=self.browse_diagram_image,
        )
        btn_browse_diag.pack(side="right", padx=(3, 0))

        btn_paste_diag = tk.Button(
            diag_head,
            text="Paste",
            font=self.font_body,
            bg="#EFF6FF",
            fg="#1D4ED8",
            relief="flat",
            cursor="hand2",
            padx=6,
            pady=1,
            command=self.paste_diagram_image,
        )
        btn_paste_diag.pack(side="right")

        # Diagram Preview Canvas
        self.canvas_diag_frame = tk.Frame(card_right, bg="#F8FAFC", highlightbackground="#CBD5E1", highlightthickness=1)
        self.canvas_diag_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        self.canvas_diag = tk.Canvas(self.canvas_diag_frame, bg="#F8FAFC", highlightthickness=0, cursor="hand2")
        self.canvas_diag.pack(fill="both", expand=True)
        self.canvas_diag.bind("<Button-1>", lambda e: self.paste_diagram_image())
        self.canvas_diag.bind("<Configure>", lambda e: self._render_diag_image() if self.diagram_image else self._render_diag_placeholder())
        self._render_diag_placeholder()

        # FIELD 3: ANSWER SECTION
        ans_card = tk.Frame(card_right, bg="#F8FAFC", highlightbackground="#E2E8F0", highlightthickness=1, padx=12, pady=10)
        ans_card.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        ans_card.columnconfigure(0, weight=1)

        # 1. Answer Label positioned above options
        tk.Label(
            ans_card,
            text="Answer:",
            font=self.font_h2,
            bg="#F8FAFC",
            fg=self.color_text,
        ).grid(row=0, column=0, sticky="w", pady=(0, 6))

        # 2. Quick select buttons (ক, খ, গ, ঘ) in full-width row
        quick_btn_box = tk.Frame(ans_card, bg="#F8FAFC")
        quick_btn_box.grid(row=1, column=0, sticky="ew")
        for col_idx in range(4):
            quick_btn_box.columnconfigure(col_idx, weight=1)

        for col_idx, opt in enumerate(["ক", "খ", "গ", "ঘ"]):
            b = tk.Button(
                quick_btn_box,
                text=opt,
                font=self.font_answer,
                bg="#E0F2FE",
                fg="#0369A1",
                activebackground="#BAE6FD",
                relief="flat",
                cursor="hand2",
                padx=6,
                pady=3,
                command=lambda val=opt: self.set_quick_answer(val),
            )
            b.grid(row=0, column=col_idx, sticky="ew", padx=2)

        # 3. Free-form answer input entry
        self.var_answer = tk.StringVar(value="")
        self.ent_answer = tk.Entry(
            ans_card,
            textvariable=self.var_answer,
            font=self.font_answer,
            bg="#FFFFFF",
            relief="solid",
            bd=1,
        )
        self.ent_answer.grid(row=2, column=0, sticky="ew", pady=(8, 0))

        # Primary Submission Button: Add to Dataset
        self.btn_submit = tk.Button(
            card_right,
            text="Save & Add Question to Dataset (Ctrl+Enter)",
            font=(self.ui_font, 10, "bold"),
            bg="#059669",
            fg="#FFFFFF",
            activebackground="#047857",
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            pady=8,
            command=self.submit_question,
        )
        self.btn_submit.grid(row=3, column=0, sticky="ew")

        # =============================================================
        # 4. BOTTOM STATUS BAR
        # =============================================================
        status_bar = tk.Frame(self.root, bg="#FFFFFF", highlightbackground="#E2E8F0", highlightthickness=1, padx=14, pady=6)
        status_bar.grid(row=3, column=0, sticky="ew")

        self.lbl_status = tk.Label(
            status_bar,
            text="Ready. Paste question screenshot (Ctrl+V) to extract with Gemini.",
            font=self.font_body,
            fg="#334155",
            bg="#FFFFFF",
        )
        self.lbl_status.pack(side="left")

        self.lbl_key_indicator = tk.Label(
            status_bar,
            text="",
            font=(self.ui_font, 8, "bold"),
            bg="#FFFFFF",
        )
        self.lbl_key_indicator.pack(side="right")
        self._update_api_key_badge()
        self.set_q_mode("both")

    def _bind_shortcuts(self) -> None:
        """Binds global keyboard shortcuts."""
        # Main submit shortcut
        self.root.bind("<Control-Return>", lambda e: self.submit_question())
        self.root.bind("<Control-s>", lambda e: (self.submit_question(), "break")[1])
        self.root.bind("<Control-S>", lambda e: (self.submit_question(), "break")[1])

        # New, Import & Export shortcuts
        self.root.bind("<Control-n>", lambda e: self.create_new_dataset_dialog())
        self.root.bind("<Control-N>", lambda e: self.create_new_dataset_dialog())
        self.root.bind("<Control-o>", lambda e: self.import_dataset_file())
        self.root.bind("<Control-O>", lambda e: self.import_dataset_file())
        self.root.bind("<Control-e>", lambda e: self.export_dataset_zip())
        self.root.bind("<Control-E>", lambda e: self.export_dataset_zip())

        # Clipboard paste shortcuts
        self.root.bind("<Control-v>", lambda e: self._on_global_paste())
        self.root.bind("<Control-V>", lambda e: self._on_global_paste())

    def _on_global_paste(self) -> None:
        """Intelligent global paste routing based on focus and clipboard content."""
        focused = self.root.focus_get()
        # If user is explicitly focused on text area or entry, allow normal text paste
        if isinstance(focused, (tk.Text, tk.Entry)):
            return

        # Check if clipboard has image
        img = self._get_image_from_clipboard()
        if img is not None:
            if self.q_mode_var.get() == "text":
                self.set_q_mode("screenshot")
            self.paste_question_screenshot()
            return

        # Otherwise check if clipboard has text
        try:
            txt = self.root.clipboard_get()
            if txt and txt.strip():
                if self.q_mode_var.get() == "screenshot":
                    self.set_q_mode("text")
                self.paste_raw_text()
                return
        except Exception:
            pass

        self.paste_question_screenshot()

    # =============================================================
    # MODE SWITCHING & RAW TEXT HANDLERS
    # =============================================================
    def set_q_mode(self, mode: str) -> None:
        """Switches between screenshot, raw text, or both views."""
        self.q_mode_var.set(mode)
        # Update tab button styling
        for btn, key in [
            (self.btn_tab_screenshot, "screenshot"),
            (self.btn_tab_text, "text"),
            (self.btn_tab_both, "both"),
        ]:
            if key == mode:
                btn.config(bg="#2563EB", fg="#FFFFFF")
            else:
                btn.config(bg="#F1F5F9", fg="#334155")

        if mode == "screenshot":
            self.frame_mode_screenshot.pack(fill="x", expand=False)
            self.frame_mode_text.pack_forget()
            self.canvas_qs_frame.config(height=160)
            if self.question_screenshot:
                self._render_qs_image()
            else:
                self._render_qs_placeholder()
        elif mode == "text":
            self.frame_mode_screenshot.pack_forget()
            self.frame_mode_text.pack(fill="x", expand=False)
            self.txt_raw_question.config(height=6)
            self.txt_raw_question.focus_set()
        else:  # both
            self.frame_mode_screenshot.pack(fill="x", expand=False)
            self.frame_mode_text.pack(fill="x", expand=False, pady=(6, 0))
            self.canvas_qs_frame.config(height=100)
            self.txt_raw_question.config(height=4)
            if self.question_screenshot:
                self._render_qs_image()
            else:
                self._render_qs_placeholder()

    def paste_raw_text(self) -> None:
        """Pastes raw question text from clipboard into txt_raw_question."""
        try:
            txt = self.root.clipboard_get()
        except Exception:
            self.show_toast("Clipboard is empty or does not contain text.", is_error=True)
            return

        if not txt or not txt.strip():
            self.show_toast("Clipboard does not contain text.", is_error=True)
            return

        self.txt_raw_question.delete("1.0", tk.END)
        self.txt_raw_question.insert("1.0", txt.strip())
        self.show_toast(f"Pasted raw question text ({len(txt)} chars).")

        if self.auto_furnish_text_var.get():
            self.start_text_furnishing()

    def clear_raw_text(self) -> None:
        """Clears raw question text box."""
        self.txt_raw_question.delete("1.0", tk.END)

    def start_text_furnishing(self) -> None:
        """Starts Gemini raw text conversion & furnishing in a background thread."""
        if self.is_extracting:
            return

        raw_text = self.txt_raw_question.get("1.0", tk.END).strip()
        if not raw_text:
            self.show_toast("Please enter or paste raw question text first.", is_error=True)
            return

        if not self.gemini.is_configured():
            ans = messagebox.askyesno(
                "Gemini API Key Required",
                "You have not configured your Gemini API Key yet.\n\nWould you like to enter it now?",
                parent=self.root,
            )
            if ans:
                self.configure_api_key_dialog()
            return

        self.is_extracting = True
        self.btn_furnish_text.config(state="disabled", text="Furnishing with Gemini...")
        self.lbl_extract_status.config(text="Calling Gemini API to furnish text...", fg="#D97706")
        self.show_toast("Furnishing raw question text with Gemini API...")

        thread = threading.Thread(target=self._run_text_furnishing_worker, args=(raw_text,), daemon=True)
        thread.start()

    def _run_text_furnishing_worker(self, raw_text: str) -> None:
        """Worker thread function for text furnishing."""
        try:
            success, result = self.gemini.furnish_question_text(raw_text)
        except Exception as e:
            success, result = False, str(e)

        self.root.after(0, lambda: self._on_text_furnishing_complete(success, result))

    def _on_text_furnishing_complete(self, success: bool, result: str) -> None:
        """Called on main UI thread when Gemini text furnishing completes."""
        self.is_extracting = False
        self.btn_furnish_text.config(state="normal", text="Furnish Text with Gemini")

        if success:
            model_used = self.gemini.last_used_model or self.gemini.primary_model
            print(f"[Console] Furnished question using model: '{model_used}'\n", flush=True)
            self.txt_question.delete("1.0", tk.END)
            self.txt_question.insert("1.0", result)
            self.txt_question.edit_reset()
            self.lbl_extract_status.config(text=f"✓ Furnished with {model_used}!", fg="#059669")
            self.show_toast(f"✓ Question furnished using {model_used}!", is_success=True)
        else:
            self.lbl_extract_status.config(text="Furnishing failed", fg="#DC2626")
            messagebox.showerror("Gemini Furnishing Error", result)
            self.show_toast("Gemini furnishing failed.", is_error=True)

    # =============================================================
    # QUESTION SCREENSHOT HANDLERS
    # =============================================================
    def paste_question_screenshot(self) -> None:
        """Fetches screenshot from clipboard for question extraction."""
        img = self._get_image_from_clipboard()
        if img is None:
            self.show_toast("Clipboard does not contain an image. Please copy screenshot first.", is_error=True)
            return

        self.question_screenshot = img
        self._render_qs_image()
        self.show_toast(f"Pasted question screenshot ({img.width}×{img.height} px).")

        # Auto-extract if enabled
        if self.auto_extract_var.get():
            self.start_gemini_extraction()

    def browse_question_screenshot(self) -> None:
        """Opens file dialog to select question screenshot."""
        file_path = filedialog.askopenfilename(
            parent=self.root,
            title="Select Question Screenshot Image",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.webp;*.bmp"), ("All Files", "*.*")],
        )
        if not file_path:
            return
        try:
            img = Image.open(file_path)
            img.load()
            self.question_screenshot = img
            self._render_qs_image()
            if self.auto_extract_var.get():
                self.start_gemini_extraction()
        except Exception as e:
            messagebox.showerror("Error Opening Image", str(e))

    def _render_qs_placeholder(self) -> None:
        """Draws placeholder for question screenshot canvas."""
        self.canvas_qs.delete("all")
        self.canvas_qs.update_idletasks()
        w = max(100, self.canvas_qs.winfo_width())
        h = max(80, self.canvas_qs.winfo_height())
        cx, cy = w // 2, h // 2

        self.canvas_qs.create_rectangle(15, 10, w - 15, h - 10, dash=(4, 4), outline="#94A3B8", fill="#F8FAFC")
        self.canvas_qs.create_text(cx, cy - 10, text="Click or Paste Question Screenshot (Ctrl+V)", font=self.font_bold, fill="#334155")
        self.canvas_qs.create_text(cx, cy + 12, text="Sends screenshot to Gemini API to extract question string", font=(self.ui_font, 8), fill="#64748B")

    def _render_qs_image(self) -> None:
        """Renders scaled question screenshot on canvas."""
        if self.question_screenshot is None:
            self._render_qs_placeholder()
            return

        self.canvas_qs.delete("all")
        self.canvas_qs.update_idletasks()
        w = max(100, self.canvas_qs.winfo_width())
        h = max(80, self.canvas_qs.winfo_height())

        orig_w, orig_h = self.question_screenshot.size
        scale = min((w - 20) / orig_w, (h - 20) / orig_h, 1.0)
        dw = max(1, int(orig_w * scale))
        dh = max(1, int(orig_h * scale))

        thumb = self.question_screenshot.copy().resize((dw, dh), Image.Resampling.LANCZOS)
        self._qs_photo = ImageTk.PhotoImage(thumb)

        cx, cy = w // 2, h // 2
        self.canvas_qs.create_rectangle(cx - dw // 2 - 1, cy - dh // 2 - 1, cx + dw // 2 + 1, cy + dh // 2 + 1, outline="#CBD5E1", fill="#FFFFFF")
        self.canvas_qs.create_image(cx, cy, image=self._qs_photo)

    # =============================================================
    # GEMINI EXTRACTION
    # =============================================================
    def start_gemini_extraction(self) -> None:
        """Starts Gemini API OCR in a background thread."""
        if self.question_screenshot is None:
            self.show_toast("Please paste or load a question screenshot first.", is_error=True)
            return

        if not self.gemini.is_configured():
            res = messagebox.askyesno(
                "Gemini API Key Required",
                "You have not configured your Gemini API Key yet.\n\nWould you like to enter it now?",
            )
            if res:
                self.configure_api_key_dialog()
            return

        if self.is_extracting:
            return

        self.is_extracting = True
        self.btn_extract.config(state="disabled", text="Extracting with Gemini...")
        self.lbl_extract_status.config(text="Calling Gemini API...", fg="#D97706")
        self.show_toast("Sending question screenshot to Gemini API...")

        # Run extraction in background thread
        thread = threading.Thread(target=self._run_extraction_worker, daemon=True)
        thread.start()

    def _run_extraction_worker(self) -> None:
        """Worker thread function."""
        try:
            success, result = self.gemini.extract_question(self.question_screenshot)
        except Exception as e:
            success, result = False, str(e)

        # Dispatch back to main UI thread
        self.root.after(0, lambda: self._on_extraction_complete(success, result))

    def _on_extraction_complete(self, success: bool, result: str) -> None:
        """Called on main thread when Gemini responds."""
        self.is_extracting = False
        self.btn_extract.config(state="normal", text="Extract Screenshot with Gemini")

        if success:
            model_used = self.gemini.last_used_model or self.gemini.primary_model
            print(f"[Console] Extracted question using model: '{model_used}'\n", flush=True)
            self.txt_question.delete("1.0", tk.END)
            self.txt_question.insert("1.0", result)
            self.txt_question.edit_reset()
            self.lbl_extract_status.config(text=f"✓ Extracted with {model_used}!", fg="#059669")
            self.show_toast(f"✓ Question extracted using {model_used}!", is_success=True)
        else:
            self.lbl_extract_status.config(text="Extraction failed", fg="#DC2626")
            messagebox.showerror("Gemini Extraction Error", result)
            self.show_toast("Gemini extraction failed.", is_error=True)

    # =============================================================
    # DIAGRAM / VISUAL IMAGE HANDLERS
    # =============================================================
    def paste_diagram_image(self) -> None:
        """Fetches diagram visual from clipboard."""
        img = self._get_image_from_clipboard()
        if img is None:
            self.show_toast("Clipboard does not contain an image.", is_error=True)
            return

        self.diagram_image = img
        self._render_diag_image()
        self.show_toast(f"Pasted diagram visual ({img.width}×{img.height} px).")

    def browse_diagram_image(self) -> None:
        """Selects diagram visual via file dialog."""
        file_path = filedialog.askopenfilename(
            parent=self.root,
            title="Select Diagram Visual Image",
            filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.webp;*.bmp"), ("All Files", "*.*")],
        )
        if not file_path:
            return
        try:
            with Image.open(file_path) as img:
                self.diagram_image = img.copy()
            self._render_diag_image()
            self.show_toast(f"Loaded visual: {Path(file_path).name}")
        except Exception as e:
            messagebox.showerror("Error Opening Image", str(e))

    def clear_diagram_image(self) -> None:
        """Clears diagram visual."""
        self.diagram_image = None
        self._render_diag_placeholder()
        self.show_toast("Cleared diagram image.")

    def _render_diag_placeholder(self) -> None:
        """Renders diagram canvas placeholder."""
        self.canvas_diag.delete("all")
        self.canvas_diag.update_idletasks()
        w = max(100, self.canvas_diag.winfo_width())
        h = max(80, self.canvas_diag.winfo_height())
        cx, cy = w // 2, h // 2

        self.canvas_diag.create_rectangle(15, 10, w - 15, h - 10, dash=(4, 4), outline="#94A3B8", fill="#F8FAFC")
        self.canvas_diag.create_text(cx, cy - 10, text="Click to Paste Visual Diagram", font=self.font_bold, fill="#334155")
        self.canvas_diag.create_text(cx, cy + 12, text="This visual will be saved into images/ and embedded in Excel", font=(self.ui_font, 8), fill="#64748B")

    def _render_diag_image(self) -> None:
        """Renders scaled diagram visual on canvas."""
        if self.diagram_image is None:
            self._render_diag_placeholder()
            return

        self.canvas_diag.delete("all")
        self.canvas_diag.update_idletasks()
        w = max(100, self.canvas_diag.winfo_width())
        h = max(80, self.canvas_diag.winfo_height())

        orig_w, orig_h = self.diagram_image.size
        scale = min((w - 20) / orig_w, (h - 20) / orig_h, 1.0)
        dw = max(1, int(orig_w * scale))
        dh = max(1, int(orig_h * scale))

        thumb = self.diagram_image.copy().resize((dw, dh), Image.Resampling.LANCZOS)
        self._diag_photo = ImageTk.PhotoImage(thumb)

        cx, cy = w // 2, h // 2
        self.canvas_diag.create_rectangle(cx - dw // 2 - 1, cy - dh // 2 - 1, cx + dw // 2 + 1, cy + dh // 2 + 1, outline="#CBD5E1", fill="#FFFFFF")
        self.canvas_diag.create_image(cx, cy, image=self._diag_photo)

    # =============================================================
    # ANSWER HANDLERS
    # =============================================================
    def set_quick_answer(self, option_letter: str) -> None:
        """Sets answer from quick buttons."""
        self.var_answer.set(option_letter)
        self.show_toast(f"Selected answer: {option_letter}")

    # =============================================================
    # SUBMIT AND DATASET GENERATION
    # =============================================================
    def submit_question(self) -> None:
        """
        Validates inputs, saves visual to images/ folder (if present),
        updates Excel workbook, and resets fields for the next question.
        """
        q_text = self.txt_question.get("1.0", "end-1c").strip()
        ans = self.var_answer.get().strip()
        has_image = self.diagram_image is not None

        # Check if all fields are empty
        if not q_text and not has_image and not ans:
            messagebox.showwarning(
                "All Fields Empty",
                "Question text, diagram image, and answer are all empty.\n\nPlease enter or paste the content before adding.",
            )
            return

        # Check if any field (question, image, or answer) is empty and warn with confirmation
        missing = []
        if not q_text:
            missing.append("Question Text (Field 1)")
        if not has_image:
            missing.append("Diagram Image (Field 2)")
        if not ans:
            missing.append("Answer (Field 3)")

        if missing:
            missing_items = "\n".join(f"  • {item}" for item in missing)
            proceed = messagebox.askyesno(
                "Warning: Empty Field(s) Detected",
                f"The following field(s) are empty:\n\n{missing_items}\n\nDo you want to proceed and save anyway?",
                icon="warning",
                default="no",
            )
            if not proceed:
                self.show_toast("Addition cancelled. Please fill in the missing fields.", is_error=True)
                return

        source = sanitize_source_name(self.var_source.get() or "dataset")

        try:
            item = self.manager.add_question(
                source=source,
                question_text=q_text,
                answer_text=ans,
                diagram_image=self.diagram_image,
                overwrite_image=True,
                overwrite_existing_entry=True,
            )

            # Reset inputs for next question
            self.reset_form_inputs()

            self._update_stats_display()
            self.show_toast(f"✓ Saved Question {item.id} to {self.manager.excel_file.name}!", is_success=True)

        except PermissionError as pe:
            messagebox.showerror(
                "Excel File Locked",
                f"Cannot write to {self.manager.excel_file.name}:\n\n{pe}\n\nPlease close the file in Excel and try again.",
            )
        except Exception as e:
            messagebox.showerror("Error Saving Question", str(e))

    # =============================================================
    # CONFIGURATION & DIALOGS
    # =============================================================
    def configure_api_key_dialog(self) -> None:
        """Opens modal dialog to set or update Gemini API Key."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Configure Gemini API Key")
        dlg.geometry("520x220")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        frm = tk.Frame(dlg, padx=20, pady=16)
        frm.pack(fill="both", expand=True)

        tk.Label(frm, text="Google Gemini API Key", font=self.font_h2).pack(anchor="w", pady=(0, 4))
        tk.Label(
            frm,
            text="Get your free API key at: https://aistudio.google.com/app/apikey",
            font=(self.ui_font, 8),
            fg="#2563EB",
        ).pack(anchor="w", pady=(0, 10))

        ent_key = tk.Entry(frm, width=54, font=(self.ui_font, 10), show="•")
        current_k = self.gemini.api_key
        if current_k:
            ent_key.insert(0, current_k)
        ent_key.pack(fill="x", pady=(0, 6))

        # Show / Hide toggle & Model Selection
        opt_row = tk.Frame(frm)
        opt_row.pack(fill="x", pady=(0, 14))

        show_var = tk.BooleanVar(value=False)

        def toggle_show():
            ent_key.config(show="" if show_var.get() else "•")

        chk_show = tk.Checkbutton(opt_row, text="Show Key", variable=show_var, command=toggle_show)
        chk_show.pack(side="left")

        # Model Selection dropdown
        tk.Label(opt_row, text="Model:", font=self.font_bold).pack(side="left", padx=(20, 4))
        var_model = tk.StringVar(value=self.gemini.primary_model)
        combo_model = ttk.Combobox(
            opt_row,
            textvariable=var_model,
            values=[
                "gemini-3.5-flash-lite",
                "gemini-flash-lite-latest",
                "gemini-2.5-flash-lite",
                "gemini-3.8-flash",
            ],
            state="readonly",
            width=24,
        )
        combo_model.pack(side="left")

        btn_box = tk.Frame(frm)
        btn_box.pack(fill="x")

        def save():
            new_key = ent_key.get().strip()
            if not new_key:
                messagebox.showwarning("Empty Key", "Please provide a valid Gemini API Key.", parent=dlg)
                return
            chosen_model = var_model.get().strip() or "gemini-3.5-flash-lite"
            self.gemini.save_api_key(new_key, model=chosen_model)
            self._update_api_key_badge()
            self.show_toast(f"✓ Gemini API Key saved ({chosen_model})!", is_success=True)
            dlg.destroy()

        tk.Button(btn_box, text="Save Key", font=self.font_bold, bg="#2563EB", fg="#FFFFFF", padx=12, pady=4, relief="flat", command=save).pack(side="right")
        tk.Button(btn_box, text="Cancel", font=self.font_body, padx=8, pady=4, relief="flat", command=dlg.destroy).pack(side="right", padx=(0, 8))

    def reset_form_inputs(self) -> None:
        """Resets all input fields, screenshots, and diagrams for a clean question form."""
        self.question_screenshot = None
        self.diagram_image = None
        self._render_qs_placeholder()
        self._render_diag_placeholder()
        self.txt_question.delete("1.0", tk.END)
        self.txt_raw_question.delete("1.0", tk.END)
        self.var_answer.set("")
        self.lbl_extract_status.config(text="")

    def create_new_dataset(self, source_name: str = "", with_timestamp: bool = True) -> Path:
        """Initializes a brand new, empty dataset with the given source name."""
        clean_src = sanitize_source_name(source_name or "dataset")
        dest_file = self.manager.create_new_dataset(clean_src, with_timestamp=with_timestamp)
        self.var_source.set(self.manager.current_source_name)
        self.reset_form_inputs()
        self._update_stats_display()
        self.show_toast(f"✓ Created new dataset: {dest_file.name}", is_success=True)
        return dest_file

    def create_new_dataset_dialog(self) -> None:
        """Prompts user to create a new dataset anytime with minimal clean UI."""
        dlg = tk.Toplevel(self.root)
        dlg.title("Create New Dataset")
        dlg.transient(self.root)
        dlg.grab_set()

        frm = tk.Frame(dlg, padx=24, pady=20, bg="#FFFFFF")
        frm.pack(fill="both", expand=True)

        tk.Label(frm, text="Create New Dataset", font=self.font_h2, fg="#0F172A", bg="#FFFFFF").pack(anchor="w", pady=(0, 12))

        row = tk.Frame(frm, bg="#FFFFFF")
        row.pack(fill="x", pady=(0, 10))
        tk.Label(row, text="Source Name:", font=self.font_bold, bg="#FFFFFF").pack(side="left", padx=(0, 8))
        var_name = tk.StringVar(value="")
        ent_name = tk.Entry(row, textvariable=var_name, font=self.font_body)
        ent_name.pack(side="left", fill="x", expand=True)
        ent_name.focus_set()

        var_ts = tk.BooleanVar(value=True)
        chk_ts = tk.Checkbutton(
            frm,
            text="Add timestamp to filename (recommended)",
            variable=var_ts,
            font=self.font_body,
            bg="#FFFFFF",
            activebackground="#FFFFFF",
        )
        chk_ts.pack(anchor="w", pady=(0, 16))

        btn_box = tk.Frame(frm, bg="#FFFFFF")
        btn_box.pack(fill="x", pady=(4, 0))

        def on_create():
            name = var_name.get().strip()
            if not name:
                messagebox.showwarning("Name Required", "Please enter a source name for the dataset.", parent=dlg)
                ent_name.focus_set()
                return
            self.create_new_dataset(source_name=name, with_timestamp=var_ts.get())
            dlg.destroy()

        ent_name.bind("<Return>", lambda e: on_create())
        dlg.bind("<Escape>", lambda e: dlg.destroy())

        tk.Button(
            btn_box,
            text="Create",
            font=self.font_bold,
            bg="#2563EB",
            fg="#FFFFFF",
            activebackground="#1D4ED8",
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            padx=16,
            pady=5,
            command=on_create,
        ).pack(side="right")

        tk.Button(
            btn_box,
            text="Cancel",
            font=self.font_body,
            padx=10,
            pady=5,
            relief="flat",
            command=dlg.destroy,
        ).pack(side="right", padx=(0, 8))

        # Measure content and center with ample DPI margin
        dlg.update_idletasks()
        req_w = max(420, dlg.winfo_reqwidth() + 50)
        req_h = max(260, dlg.winfo_reqheight() + 30)
        try:
            rx = self.root.winfo_x()
            ry = self.root.winfo_y()
            rw = self.root.winfo_width()
            rh = self.root.winfo_height()
            x = rx + max(0, (rw - req_w) // 2)
            y = ry + max(0, (rh - req_h) // 2)
            dlg.geometry(f"{req_w}x{req_h}+{x}+{y}")
        except Exception:
            dlg.geometry(f"{req_w}x{req_h}")
        dlg.minsize(380, 240)

    def prompt_startup_choice(self) -> None:
        """
        Minimal startup message box with 2 clean options:
        1. Create New Dataset
        2. Import Dataset
        """
        dlg = tk.Toplevel(self.root)
        dlg.title("Get Started")
        dlg.transient(self.root)
        dlg.grab_set()

        box = tk.Frame(dlg, padx=26, pady=22, bg="#FFFFFF")
        box.pack(fill="both", expand=True)

        tk.Label(
            box,
            text="Get Started",
            font=self.font_h1,
            fg="#0F172A",
            bg="#FFFFFF",
        ).pack(anchor="w", pady=(0, 4))

        tk.Label(
            box,
            text="Choose an option to begin:",
            font=self.font_body,
            fg="#64748B",
            bg="#FFFFFF",
        ).pack(anchor="w", pady=(0, 16))

        def do_create():
            dlg.destroy()
            self.create_new_dataset_dialog()

        def do_import():
            dlg.withdraw()
            imported = self.import_dataset_file()
            if imported:
                dlg.destroy()
            else:
                dlg.deiconify()

        btn_create = tk.Button(
            box,
            text="+ Create New Dataset",
            font=self.font_bold,
            bg="#2563EB",
            fg="#FFFFFF",
            activebackground="#1D4ED8",
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            pady=8,
            command=do_create,
        )
        btn_create.pack(fill="x", pady=(0, 10))

        btn_import = tk.Button(
            box,
            text="Import Dataset (XLSX / CSV)",
            font=self.font_bold,
            bg="#059669",
            fg="#FFFFFF",
            activebackground="#047857",
            activeforeground="#FFFFFF",
            relief="flat",
            cursor="hand2",
            pady=8,
            command=do_import,
        )
        btn_import.pack(fill="x", pady=(0, 8))

        # Dynamically measure content with ample DPI margin and center over root
        dlg.update_idletasks()
        req_w = max(420, dlg.winfo_reqwidth() + 50)
        req_h = max(260, dlg.winfo_reqheight() + 30)
        try:
            rx = self.root.winfo_x()
            ry = self.root.winfo_y()
            rw = self.root.winfo_width()
            rh = self.root.winfo_height()
            x = rx + max(0, (rw - req_w) // 2)
            y = ry + max(0, (rh - req_h) // 2)
            dlg.geometry(f"{req_w}x{req_h}+{x}+{y}")
        except Exception:
            dlg.geometry(f"{req_w}x{req_h}")
        dlg.minsize(380, 240)

    def import_dataset_file(self) -> bool:
        """
        Prompts user to select a .csv or .xlsx file, copies it into the active data/
        directory with a timestamped name ({stem}_{timestamp}{ext}), and switches to working on it.
        Returns True if successful, False otherwise.
        """
        fpath = filedialog.askopenfilename(
            parent=self.root,
            title="Import Dataset File (Excel or CSV)",
            filetypes=[("Dataset Files (*.xlsx;*.csv)", "*.xlsx;*.csv"), ("All Files", "*.*")],
        )
        if not fpath:
            return False

        src_path = Path(fpath).resolve()
        if not src_path.is_file():
            messagebox.showerror("File Not Found", f"Cannot find file:\n{src_path}")
            return False

        try:
            import shutil
            from datetime import datetime

            orig_stem = sanitize_source_name(src_path.stem)
            ext = src_path.suffix.lower()
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            dest_filename = f"{orig_stem}_{timestamp}{ext}"
            dest_path = self.manager.data_dir / dest_filename

            # Ensure data_dir exists and copy file
            self.manager.data_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(str(src_path), str(dest_path))

            # Also check if the source file had an images/ folder next to it; copy any images if present
            src_images_dir = src_path.parent / "images"
            if src_images_dir.is_dir() and src_images_dir != self.manager.images_dir:
                self.manager.images_dir.mkdir(parents=True, exist_ok=True)
                for img_f in src_images_dir.iterdir():
                    if img_f.is_file():
                        target_img = self.manager.images_dir / img_f.name
                        if not target_img.exists():
                            shutil.copy2(str(img_f), str(target_img))

            # Switch to working on the newly imported file
            count = self.manager.load_dataset_file(dest_path)
            self.var_source.set(self.manager.current_source_name)
            self.reset_form_inputs()
            self._update_stats_display()

            self.show_toast(f"✓ Imported {count} questions! Working on {dest_filename}", is_success=True)
            messagebox.showinfo(
                "Dataset Imported",
                f"File imported and copied to:\n{dest_filename}\n\nQuestions loaded: {count}\nActive source: {self.manager.current_source_name}",
            )
            return True
        except Exception as e:
            messagebox.showerror("Error Importing Dataset", str(e))
            return False

    def export_dataset_zip(self) -> None:
        """
        Exports CSV, JSONL, HTML, and bundles everything into
        a portable sourcename_timestamp.zip package.
        """
        if self.manager.total_count == 0:
            messagebox.showwarning("Empty Dataset", "There are no questions in the current dataset to export.")
            return

        try:
            source_name = self.var_source.get() or "dataset"
            zip_path = self.manager.export_zip_package(source_name=source_name)

            self.show_toast(f"✓ Exported ZIP package: {zip_path.name}", is_success=True)
            open_folder = messagebox.askyesno(
                "Export Complete",
                f"Dataset package successfully generated:\n\n{zip_path.name}\n\nLocation: {zip_path.parent}\n\nWould you like to open the exports folder?",
                icon="info",
            )
            if open_folder:
                self._reveal_in_explorer(self.manager.exports_dir)
        except Exception as e:
            messagebox.showerror("Error Exporting Package", str(e))

    def open_dataset_folder(self) -> None:
        """Reveals active dataset directory in Windows File Explorer."""
        self._reveal_in_explorer(self.manager.data_dir)

    def open_images_folder(self) -> None:
        """Reveals images/ folder in Windows File Explorer."""
        self.manager.images_dir.mkdir(parents=True, exist_ok=True)
        self._reveal_in_explorer(self.manager.images_dir)

    def _reveal_in_explorer(self, folder: Path) -> None:
        """Opens path in native OS explorer."""
        try:
            if sys.platform == "win32":
                os.startfile(str(folder))
            else:
                import subprocess
                subprocess.Popen(["xdg-open", str(folder)])
        except Exception as e:
            self.show_toast(f"Could not open folder: {e}", is_error=True)

    # =============================================================
    # HELPERS
    # =============================================================
    def _update_id_preview(self) -> None:
        """Updates the live ID preview badge."""
        try:
            raw_src = self.var_source.get().strip()
            next_idx = len(self.manager.entries) + 1
            if raw_src:
                src = sanitize_source_name(raw_src)
                self.lbl_id_badge.config(text=f"Next ID: {src}_{next_idx:03d}")
            else:
                self.lbl_id_badge.config(text="Next ID: (set source)")
        except Exception:
            pass

    def _apply_header_layout(self, width: Optional[int] = None) -> None:
        """Adapts top bar layout dynamically so all buttons remain fully visible at any window width and DPI scaling."""
        if width is None:
            width = self.root.winfo_width()
        if width <= 1:
            width = 1180

        # Calculate exact needed widths based on current widget metrics & DPI
        pad = 36
        t_w = self.tbox.winfo_reqwidth()
        b_w = sum(b.winfo_reqwidth() + 4 for b in self._all_header_buttons)
        avail = width - pad

        if avail >= t_w + b_w + 20:
            target_mode = "wide"
        elif avail >= b_w:
            target_mode = "stacked_single"
        else:
            target_mode = "stacked_grid"

        if target_mode == self._header_layout_mode:
            return

        for b in self._all_header_buttons:
            b.grid_forget()

        if target_mode == "wide":
            self.tbox.grid(row=0, column=0, sticky="w")
            self.ctrl_box.grid(row=0, column=1, sticky="e")
            self.header_bar.columnconfigure(0, weight=1)
            self.header_bar.columnconfigure(1, weight=0)
            for i, b in enumerate(self._all_header_buttons):
                b.grid(row=0, column=i, padx=2, pady=1)
        elif target_mode == "stacked_single":
            self.tbox.grid(row=0, column=0, sticky="w")
            self.ctrl_box.grid(row=1, column=0, sticky="w", pady=(3, 0))
            self.header_bar.columnconfigure(0, weight=1)
            self.header_bar.columnconfigure(1, weight=0)
            for i, b in enumerate(self._all_header_buttons):
                b.grid(row=0, column=i, padx=2, pady=1)
        else:  # stacked_grid
            self.tbox.grid(row=0, column=0, sticky="w")
            self.ctrl_box.grid(row=1, column=0, sticky="w", pady=(3, 0))
            self.header_bar.columnconfigure(0, weight=1)
            self.header_bar.columnconfigure(1, weight=0)
            # Row 0: Primary dataset management
            self.btn_new.grid(row=0, column=0, padx=2, pady=1, sticky="w")
            self.btn_import.grid(row=0, column=1, padx=2, pady=1, sticky="w")
            self.btn_export.grid(row=0, column=2, padx=2, pady=1, sticky="w")
            # Row 1: Key & folder shortcuts
            self.btn_api_key.grid(row=1, column=0, padx=2, pady=1, sticky="w")
            self.btn_img_dir.grid(row=1, column=1, padx=2, pady=1, sticky="w")
            self.btn_folder.grid(row=1, column=2, padx=2, pady=1, sticky="w")

        self._header_layout_mode = target_mode

    def _on_window_configure(self, event) -> None:
        """Handles window resize to dynamically adapt top header bar."""
        if event.widget == self.root:
            self._apply_header_layout(event.width)

    def _update_stats_display(self) -> None:
        """Updates header info and count badge."""
        if self.manager.excel_file and self.manager.excel_file.is_file():
            self.lbl_dataset_info.config(text=f"Dataset: {self.manager.excel_file.name}")
        elif self.manager.current_source_name:
            self.lbl_dataset_info.config(text=f"Dataset: {self.manager.current_source_name}.xlsx")
        else:
            self.lbl_dataset_info.config(text="No dataset loaded")

        self.lbl_count_badge.config(text=f"Questions in Dataset: {self.manager.total_count}")
        self._update_id_preview()
        self._apply_header_layout()

    def _update_api_key_badge(self) -> None:
        """Updates the top button indicator for Gemini API key."""
        model_name = self.gemini.primary_model
        if self.gemini.is_configured():
            self.btn_api_key.config(text=f"✓ {model_name}", bg="#059669")
            self.lbl_key_indicator.config(text=f"{model_name} Ready", fg="#059669")
        else:
            self.btn_api_key.config(text="Gemini API Key", bg="#D97706")
            self.lbl_key_indicator.config(text=f"{model_name} (Key needed)", fg="#D97706")
        self._apply_header_layout()

    def _get_image_from_clipboard(self) -> Optional[Image.Image]:
        """Safely fetches image from system clipboard."""
        try:
            clip = ImageGrab.grabclipboard()
            if isinstance(clip, Image.Image):
                return clip
            if isinstance(clip, list) and clip and os.path.isfile(clip[0]):
                img = Image.open(clip[0])
                img.load()
                return img
        except Exception:
            pass
        return None

    def show_toast(self, message: str, is_error: bool = False, is_success: bool = False) -> None:
        """Displays status message in bottom bar."""
        color = "#DC2626" if is_error else ("#059669" if is_success else "#2563EB")
        self.lbl_status.config(text=message, fg=color)

        if self._status_clear_job:
            self.root.after_cancel(self._status_clear_job)
        self._status_clear_job = self.root.after(
            5000,
            lambda: self.lbl_status.config(
                text="Ready. Paste question screenshot (Ctrl+V) to extract with Gemini.",
                fg="#334155",
            ),
        )

    def _on_question_undo(self, event=None):
        """Performs undo on the question text widget."""
        try:
            self.txt_question.edit_undo()
        except tk.TclError:
            pass
        return "break"

    def _on_question_redo(self, event=None):
        """Performs redo on the question text widget."""
        try:
            self.txt_question.edit_redo()
        except tk.TclError:
            pass
        return "break"

