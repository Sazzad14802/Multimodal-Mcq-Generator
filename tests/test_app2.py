"""
MMU App-2 Unit and Integration Tests
Validates models, GeminiClient, DatasetManager, and UI flow.
"""

import os
import shutil
import sys
import tempfile
import unittest
from unittest.mock import patch
from pathlib import Path
import tkinter as tk
from PIL import Image
import openpyxl

# Add app-2 to sys.path
app2_dir = Path(__file__).resolve().parent.parent
if str(app2_dir) not in sys.path:
    sys.path.insert(0, str(app2_dir))

from core.models import QuestionItem, generate_unique_image_filename, sanitize_source_name
from core.gemini_client import GeminiClient
from core.dataset_manager import DatasetManager
from ui.main_window import App2MainWindow


class TestApp2Suite(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.data_dir = Path(self.temp_dir) / "data"
        self.manager = DatasetManager(data_dir=self.data_dir)
        self.gemini = GeminiClient(base_dir=Path(self.temp_dir), api_key="dummy_key_for_testing")

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_naming_and_sanitization(self):
        """Verify unique filename generation and source name sanitization."""
        self.assertEqual(sanitize_source_name("phy1-5.pdf"), "phy1-5")
        self.assertEqual(sanitize_source_name("my_test file #2"), "my_test_file_2")

        images_dir = self.data_dir / "images"
        name1 = generate_unique_image_filename(images_dir, "phy1-5", index=1)
        self.assertEqual(name1, "phy1-5_001.png")

        # Create dummy file to force collision
        (images_dir / name1).touch()
        name2 = generate_unique_image_filename(images_dir, "phy1-5", index=1)
        self.assertEqual(name2, "phy1-5_001_2.png")

        # Verify backward-compatible page/qno fallback if provided
        name_legacy = generate_unique_image_filename(images_dir, "phy1-5", page=2, question_no=14)
        self.assertEqual(name_legacy, "phy1-5_p02_q14.png")

    def test_gemini_client_config(self):
        """Verify GeminiClient key persistence and base64 conversion."""
        self.assertTrue(self.gemini.is_configured())

        test_img = Image.new("RGB", (20, 20), color=(255, 0, 0))
        b64 = self.gemini.image_to_base64_png(test_img)
        self.assertIsInstance(b64, str)
        self.assertTrue(len(b64) > 10)

        # Test saving new key
        self.gemini.save_api_key("new_test_key_1234567890")
        loaded_key = self.gemini.load_saved_api_key()
        self.assertEqual(loaded_key, "new_test_key_1234567890")

    def test_add_question_and_excel_generation(self):
        """Verify add_question creates images and biology.xlsx (5 cols), with export helpers."""
        diagram = Image.new("RGB", (120, 80), color=(100, 150, 200))
        item = self.manager.add_question(
            source="biology",
            question_text="দণ্ডটির dx অংশের ভর কোনটি?\n\nক) ML⁻¹ dx\nখ) LM⁻¹ dx",
            answer_text="ক",
            diagram_image=diagram,
        )

        self.assertEqual(item.id, "biology_001")
        self.assertEqual(self.manager.total_count, 1)

        # Check that image was saved
        saved_img = self.data_dir / item.image_path
        self.assertTrue(saved_img.is_file(), f"Image file {saved_img} should exist on disk")

        # Check Excel file
        self.assertTrue(self.manager.excel_file.is_file())
        wb = openpyxl.load_workbook(str(self.manager.excel_file), data_only=True)
        try:
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            self.assertEqual(len(rows), 2)  # Header + 1 question
            self.assertEqual(rows[0], ("id", "question", "answer", "image", "image_path"))
            self.assertEqual(rows[1][0], "biology_001")
            self.assertIn("দণ্ডটির dx অংশের ভর কোনটি?", rows[1][1])
            self.assertEqual(rows[1][2], "ক")
            self.assertEqual(rows[1][4], "images/biology_001.png")
            # Verify embedded drawing exists
            self.assertEqual(len(ws._images), 1)
        finally:
            wb.close()

        # Check CSV export on demand
        self.manager.export_csv()
        self.assertTrue(self.manager.csv_file.is_file())
        with open(self.manager.csv_file, "r", encoding="utf-8-sig") as f:
            csv_content = f.read()
            self.assertIn("biology_001", csv_content)

        # Check JSONL export on demand
        self.manager.export_jsonl()
        self.assertTrue(self.manager.jsonl_file.is_file())
        with open(self.manager.jsonl_file, "r", encoding="utf-8") as f:
            jsonl_content = f.read()
            self.assertIn("biology_001", jsonl_content)

    def test_ui_initialization_and_undo(self):
        """Verify UI starts cleanly without biology auto-filled and text box undo works."""
        root = tk.Tk()
        root.withdraw()
        try:
            app = App2MainWindow(root, self.manager, self.gemini, prompt_startup=False)
            # Verify source starts empty (not auto-filled with biology)
            self.assertEqual(app.var_source.get(), "")
            self.assertEqual(app.lbl_id_badge.cget("text"), "Next ID: (set source)")

            # When source is set, ID badge updates
            app.var_source.set("physics")
            self.assertIn("Next ID: physics_001", app.lbl_id_badge.cget("text"))

            # Test text box undo
            app.txt_question.insert("1.0", "মূল প্রশ্ন")
            app.txt_question.edit_reset()
            app.txt_question.insert("end", " - পরিবর্তন")
            self.assertIn("পরিবর্তন", app.txt_question.get("1.0", "end-1c"))

            app._on_question_undo()
            self.assertEqual(app.txt_question.get("1.0", "end-1c"), "মূল প্রশ্ন")

            # Test raw text input box and mode switching
            self.assertTrue(hasattr(app, "txt_raw_question"))
            app.set_q_mode("text")
            self.assertEqual(app.q_mode_var.get(), "text")
            app.txt_raw_question.insert("1.0", "GKwU †evgviæ wegvb...")
            self.assertIn("†evgviæ", app.txt_raw_question.get("1.0", "end-1c"))

            app.set_q_mode("screenshot")
            self.assertEqual(app.q_mode_var.get(), "screenshot")

            app.set_q_mode("both")
            self.assertEqual(app.q_mode_var.get(), "both")
        finally:
            root.destroy()

    def test_default_model_configuration(self):
        """Verify gemini-3.5-flash-lite is the default model."""
        client = GeminiClient(base_dir=Path(self.temp_dir))
        self.assertEqual(client.primary_model, "gemini-3.5-flash-lite")
        self.assertIn("gemini-flash-lite-latest", client.fallback_models)

    def test_empty_field_warnings(self):
        """Verify warnings trigger when question, image, or answer is empty."""
        root = tk.Tk()
        root.withdraw()
        try:
            app = App2MainWindow(root, self.manager, self.gemini, prompt_startup=False)

            # Case 1: All fields empty -> showwarning called
            with patch("tkinter.messagebox.showwarning") as mock_warn:
                app.submit_question()
                mock_warn.assert_called_once()
                self.assertEqual(self.manager.total_count, 0)

            # Case 2: Only question provided (image & answer empty) -> askyesno called
            app.txt_question.insert("1.0", "প্রশ্ন ১: ত্বরণের একক কোনটি?")
            with patch("tkinter.messagebox.askyesno", return_value=False) as mock_ask:
                app.submit_question()
                mock_ask.assert_called_once()
                call_args = mock_ask.call_args[0]
                self.assertIn("Warning: Empty Field(s) Detected", call_args[0])
                self.assertIn("Diagram Image", call_args[1])
                self.assertIn("Answer", call_args[1])
                # Cancelled so total_count remains 0
                self.assertEqual(self.manager.total_count, 0)

            # Case 3: If user confirms 'Yes', it saves despite missing image/answer
            with patch("tkinter.messagebox.askyesno", return_value=True) as mock_ask:
                app.submit_question()
                mock_ask.assert_called_once()
                self.assertEqual(self.manager.total_count, 1)
                saved_item = self.manager.entries[0]
                self.assertEqual(saved_item.question, "প্রশ্ন ১: ত্বরণের একক কোনটি?")
                self.assertEqual(saved_item.answer, "")
        finally:
            root.destroy()

    def test_manual_visual_flow_no_auto_pull(self):
        """Verify diagram visual is strictly manual and does not auto-pull from disk."""
        # Create dummy image in images/ directory
        img1 = Image.new("RGB", (60, 60), color=(10, 20, 30))
        img1.save(str(self.manager.images_dir / "biology_001.png"))

        root = tk.Tk()
        root.withdraw()
        try:
            app = App2MainWindow(root, self.manager, self.gemini, prompt_startup=False)

            # 1. diagram_image MUST be None on startup (no auto-pull)
            self.assertIsNone(app.diagram_image)
            self.assertFalse(hasattr(app, "auto_load_img_var"))
            self.assertFalse(hasattr(app, "current_auto_image_name"))

            # 2. Enter question, answer, and source name without diagram image
            app.var_source.set("biology")
            app.txt_question.insert("1.0", "বায়োলজি প্রশ্ন ১: মাইটোকনড্রিয়ার কাজ কী?")
            app.var_answer.set("খ")
            with patch("tkinter.messagebox.askyesno", return_value=True):
                app.submit_question()

            # 3. Verify entry was saved with no image
            self.assertEqual(self.manager.total_count, 1)
            item1 = self.manager.entries[0]
            self.assertEqual(item1.id, "biology_001")
            self.assertEqual(item1.image_path, "")
            self.assertIsNone(app.diagram_image)
        finally:
            root.destroy()

    def test_export_html_and_zip_package(self):
        """Verify HTML export and ZIP packaging with sourcename_timestamp."""
        import zipfile

        # Add question with diagram
        diag = Image.new("RGB", (100, 100), color=(50, 120, 200))
        self.manager.add_question(
            source="biology",
            question_text="কোষ প্রাচীরের মূল উপাদান কী?\nক) সেলুলোজ\nখ) প্রোটিন",
            answer_text="ক",
            diagram_image=diag,
        )
        # Add question without diagram
        self.manager.add_question(
            source="biology",
            question_text="প্লাস্টিড কয় প্রকার?\nক) ২\nখ) ৩",
            answer_text="খ",
            diagram_image=None,
        )

        self.assertEqual(self.manager.total_count, 2)

        # 1. Test HTML Export
        html_file = self.manager.export_html()
        self.assertTrue(html_file.is_file())
        with open(html_file, "r", encoding="utf-8") as f:
            content = f.read()
            self.assertIn("কোষ প্রাচীরের মূল উপাদান কী?", content)
            self.assertIn("images/biology_001.png", content)
            self.assertIn("প্লাস্টিড কয় প্রকার?", content)
            self.assertIn("সঠিক উত্তর: ক", content)
            self.assertIn("সঠিক উত্তর: খ", content)

        # 2. Test ZIP Package Export
        zip_path = self.manager.export_zip_package(source_name="biology")
        self.assertTrue(zip_path.is_file())
        self.assertTrue(zip_path.name.startswith("biology_"))
        self.assertTrue(zip_path.name.endswith(".zip"))

        # Verify ZIP contents
        with zipfile.ZipFile(str(zip_path), "r") as zf:
            namelist = zf.namelist()
            self.assertIn("biology.xlsx", namelist)
            self.assertIn("biology.csv", namelist)
            self.assertIn("biology.jsonl", namelist)
            self.assertIn("biology.html", namelist)
            self.assertIn("images/biology_001.png", namelist)

    def test_import_dataset_file_and_ui_buttons(self):
        """Verify dataset import workflow (copy to data/ with timestamp) and UI buttons."""
        import csv

        # Create an external CSV file with sample data
        ext_dir = Path(self.temp_dir) / "external_import"
        ext_dir.mkdir()
        ext_img_dir = ext_dir / "images"
        ext_img_dir.mkdir()

        ext_img = Image.new("RGB", (40, 40), color=(200, 100, 50))
        ext_img.save(str(ext_img_dir / "chem_001.png"))

        ext_csv = ext_dir / "chemistry_mcq.csv"
        with open(ext_csv, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["id", "question", "answer", "image_path"])
            writer.writerow(["chem_001", "পানির রাসায়নিক সংকেত কোনটি?", "H2O", "images/chem_001.png"])

        root = tk.Tk()
        root.withdraw()
        try:
            app = App2MainWindow(root, self.manager, self.gemini, prompt_startup=False)

            # Check that import and export buttons exist in UI
            self.assertTrue(hasattr(app, "btn_import"))
            self.assertTrue(hasattr(app, "btn_export"))
            self.assertEqual(app.btn_import.cget("text"), "Import Dataset")
            self.assertEqual(app.btn_export.cget("text"), "Export (ZIP)")

            # Test import execution via filedialog mock
            with patch("tkinter.filedialog.askopenfilename", return_value=str(ext_csv)):
                with patch("tkinter.messagebox.showinfo"):
                    app.import_dataset_file()

            # Verify dataset was loaded and active file is timestamped
            self.assertEqual(self.manager.total_count, 1)
            self.assertEqual(self.manager.entries[0].id, "chem_001")
            self.assertEqual(self.manager.entries[0].answer, "H2O")
            self.assertTrue(self.manager.excel_file.name.startswith("chemistry_mcq_"))
            self.assertTrue(self.manager.excel_file.name.endswith(".xlsx"))
            # Verify image was copied over to data/images/
            self.assertTrue((self.manager.images_dir / "chem_001.png").is_file())
            # Verify clean source name
            self.assertEqual(self.manager.current_source_name, "chemistry_mcq")
        finally:
            root.destroy()

    def test_no_auto_load_on_init_and_create_new_dataset(self):
        """Verify DatasetManager does NOT auto-load existing files on init, and create_new_dataset works."""
        # 1. Place an existing dataset on disk
        existing_xlsx = self.data_dir / "biology.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["id", "question", "answer", "image", "image_path"])
        ws.append(["biology_001", "পূর্ববর্তী প্রশ্ন", "ক", "", ""])
        wb.save(str(existing_xlsx))
        wb.close()

        # 2. Instantiate new manager
        fresh_manager = DatasetManager(data_dir=self.data_dir)
        # MUST NOT auto-load!
        self.assertEqual(fresh_manager.total_count, 0)
        self.assertEqual(len(fresh_manager.entries), 0)

        # 3. Test create_new_dataset with timestamp
        new_file = fresh_manager.create_new_dataset("physics", with_timestamp=True)
        self.assertTrue(new_file.is_file())
        self.assertTrue(new_file.name.startswith("physics_"))
        self.assertEqual(fresh_manager.total_count, 0)
        self.assertEqual(fresh_manager.current_source_name, "physics")

    def test_ui_new_dataset_and_startup_prompt_availability(self):
        """Verify UI has + New Dataset button and startup choice prompt method."""
        root = tk.Tk()
        root.withdraw()
        try:
            app = App2MainWindow(root, self.manager, self.gemini, prompt_startup=False)
            self.assertTrue(hasattr(app, "btn_new"))
            self.assertEqual(app.btn_new.cget("text"), "+ New Dataset")
            self.assertTrue(hasattr(app, "prompt_startup_choice"))
            self.assertTrue(hasattr(app, "create_new_dataset_dialog"))

            # Test creating a new dataset through the UI method
            dest = app.create_new_dataset("math_algebra", with_timestamp=True)
            self.assertTrue(dest.is_file())
            self.assertEqual(app.var_source.get(), "math_algebra")
            self.assertEqual(self.manager.total_count, 0)
            self.assertIn("math_algebra_001", app.lbl_id_badge.cget("text"))
        finally:
            root.destroy()

    def test_responsive_header_no_clipping_on_resize(self):
        """Verify all top bar buttons adapt layout and stay completely visible across widths."""
        root = tk.Tk()
        try:
            app = App2MainWindow(root, self.manager, self.gemini, prompt_startup=False)
            self.assertEqual(len(app._all_header_buttons), 6)

            for w in [1600, 1200, 1000, 850, 750, 700]:
                root.geometry(f"{w}x700")
                root.update()
                app._apply_header_layout(w)
                root.update()

                # Verify mode is one of the valid responsive tiers
                self.assertIn(app._header_layout_mode, ["wide", "stacked_single", "stacked_grid"])

                # Verify all buttons are gridded and visible
                for btn in app._all_header_buttons:
                    self.assertGreater(btn.winfo_width(), 0)
                    self.assertGreater(btn.winfo_height(), 0)
                    btn_right = btn.winfo_x() + btn.winfo_width() + app.ctrl_box.winfo_x()
                    self.assertLessEqual(btn_right, root.winfo_width() + 10)
        finally:
            root.destroy()


if __name__ == "__main__":
    unittest.main()


