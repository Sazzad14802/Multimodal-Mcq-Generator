"""
MMU App-2 Dataset Manager
Synchronizes question data into dataset.xlsx (with embedded thumbnails),
dataset.csv (UTF-8 BOM), and dataset.jsonl, while saving diagram visuals to images/.
"""

import csv
from datetime import datetime
import html
import json
from pathlib import Path
import re
import shutil
from typing import List, Optional, Dict, Any
import zipfile
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image

from .models import QuestionItem, generate_unique_image_filename, sanitize_source_name


class DatasetManager:
    """
    Manages Excel synchronization, image file storage, CSV, and JSONL persistence.
    """

    EXCEL_HEADERS = [
        "id",
        "question",
        "answer",
        "image",
        "image_path",
    ]

    CSV_HEADERS = [
        "id",
        "question",
        "answer",
        "image_path",
    ]

    def __init__(self, data_dir: Optional[Path] = None):
        if data_dir is None:
            self.data_dir = Path.cwd() / "data"
        else:
            self.data_dir = Path(data_dir).resolve()

        self.images_dir = self.data_dir / "images"
        self.exports_dir = self.data_dir / "exports"
        self.excel_file = self.data_dir / "dataset.xlsx"
        self.csv_file = self.data_dir / "dataset.csv"
        self.jsonl_file = self.data_dir / "dataset.jsonl"

        # Ensure directories exist immediately
        self.images_dir.mkdir(parents=True, exist_ok=True)
        self.exports_dir.mkdir(parents=True, exist_ok=True)

        self.current_source_name = ""
        self.entries: List[QuestionItem] = []
        # Note: Datasets are no longer auto-loaded on init.
        # Use create_new_dataset() or load_dataset_file() explicitly.

    def create_new_dataset(self, source_name: str = "", with_timestamp: bool = True) -> Path:
        """
        Initializes a fresh, empty dataset with the specified source name.
        Creates an empty Excel file with standard headers immediately.
        """
        clean_src = sanitize_source_name(source_name or "dataset")
        self.current_source_name = clean_src
        self.entries = []

        if with_timestamp:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            stem = f"{clean_src}_{ts}"
        else:
            stem = clean_src

        self.excel_file = self.data_dir / f"{stem}.xlsx"
        self.csv_file = self.data_dir / f"{stem}.csv"
        self.jsonl_file = self.data_dir / f"{stem}.jsonl"

        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.save_excel()
        return self.excel_file

    def set_dataset_directory(self, target_dir: Path) -> None:
        """Points active dataset to another directory, creating images/ folder."""
        self.data_dir = Path(target_dir).resolve()
        self.images_dir = self.data_dir / "images"
        self.exports_dir = self.data_dir / "exports"
        self.excel_file = self.data_dir / "dataset.xlsx"
        self.csv_file = self.data_dir / "dataset.csv"
        self.jsonl_file = self.data_dir / "dataset.jsonl"

        self.images_dir.mkdir(parents=True, exist_ok=True)
        self.exports_dir.mkdir(parents=True, exist_ok=True)

    def load_dataset_file(self, file_path: Path) -> int:
        """Loads an existing .xlsx or .csv dataset file and sets directory."""
        file_path = Path(file_path).resolve()
        if not file_path.is_file():
            raise FileNotFoundError(f"File not found: {file_path}")

        ext = file_path.suffix.lower()
        if ext not in [".xlsx", ".csv"]:
            raise ValueError("Supported dataset files must be .xlsx or .csv")

        # Update target directory and ensure images/ exists
        self.set_dataset_directory(file_path.parent)
        target_stem = file_path.stem

        # Strip timestamp suffix like _20260927_234012 if present to determine clean source name
        clean_stem = re.sub(r"_\d{8}_\d{6}$", "", target_stem)
        if clean_stem.lower() != "dataset":
            self.current_source_name = sanitize_source_name(clean_stem)
        else:
            self.current_source_name = "dataset"

        self.excel_file = self.data_dir / f"{target_stem}.xlsx"
        self.csv_file = self.data_dir / f"{target_stem}.csv"
        self.jsonl_file = self.data_dir / f"{target_stem}.jsonl"

        if ext == ".xlsx":
            loaded = self._read_excel(file_path, default_source=self.current_source_name)
        else:
            loaded = self._read_csv(file_path, default_source=self.current_source_name)

        self.entries = loaded

        # Auto-generate xlsx if loaded from csv
        if ext == ".csv" and not self.excel_file.is_file():
            try:
                self.save_excel()
            except PermissionError:
                pass

        return len(self.entries)

    def _read_excel(self, path: Path, default_source: str) -> List[QuestionItem]:
        wb = openpyxl.load_workbook(str(path), data_only=True)
        try:
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                return []
            headers = [str(c).strip().lower() if c else "" for c in rows[0]]
            col_map = {name: i for i, name in enumerate(headers)}

            id_i = col_map.get("id")
            page_i = col_map.get("page")
            qno_i = col_map.get("question_no") or col_map.get("qno")
            q_i = col_map.get("question") or col_map.get("q")
            ans_i = col_map.get("answer") or col_map.get("ans") or col_map.get("answer_option")
            img_i = col_map.get("image_path") or col_map.get("image")

            items: List[QuestionItem] = []
            for row in rows[1:]:
                if not row or all(v is None for v in row):
                    continue
                page = int(row[page_i]) if page_i is not None and row[page_i] is not None and str(row[page_i]).strip() else None
                qno = int(row[qno_i]) if qno_i is not None and row[qno_i] is not None and str(row[qno_i]).strip() else None
                q_text = str(row[q_i] or "") if q_i is not None else ""
                ans_text = str(row[ans_i] or "") if ans_i is not None else ""
                img_path = str(row[img_i] or "") if img_i is not None else ""

                entry_id = str(row[id_i] or "") if id_i is not None and row[id_i] else f"{default_source}_{len(items)+1:03d}"

                items.append(
                    QuestionItem(
                        id=entry_id,
                        source=default_source,
                        question=q_text,
                        answer=ans_text,
                        image_path=img_path,
                        page=page,
                        question_no=qno,
                    )
                )
            return items
        finally:
            wb.close()

    def _read_csv(self, path: Path, default_source: str) -> List[QuestionItem]:
        with open(path, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            rows = list(reader)
        if not rows:
            return []
        headers = [str(c).strip().lower() for c in rows[0]]
        col_map = {name: i for i, name in enumerate(headers)}

        id_i = col_map.get("id")
        page_i = col_map.get("page")
        qno_i = col_map.get("question_no") or col_map.get("qno")
        q_i = col_map.get("question") or col_map.get("q")
        ans_i = col_map.get("answer") or col_map.get("ans")
        img_i = col_map.get("image_path") or col_map.get("image")

        items: List[QuestionItem] = []
        for row in rows[1:]:
            if not row or all(not str(v).strip() for v in row):
                continue
            page = int(row[page_i]) if page_i is not None and row[page_i] and str(row[page_i]).strip() else None
            qno = int(row[qno_i]) if qno_i is not None and row[qno_i] and str(row[qno_i]).strip() else None
            q_text = str(row[q_i] or "") if q_i is not None else ""
            ans_text = str(row[ans_i] or "") if ans_i is not None else ""
            img_path = str(row[img_i] or "") if img_i is not None else ""

            entry_id = str(row[id_i] or "") if id_i is not None and row[id_i] else f"{default_source}_{len(items)+1:03d}"

            items.append(
                QuestionItem(
                    id=entry_id,
                    source=default_source,
                    question=q_text,
                    answer=ans_text,
                    image_path=img_path,
                    page=page,
                    question_no=qno,
                )
            )
        return items

    def add_question(
        self,
        source: str,
        question_text: str,
        answer_text: str,
        diagram_image: Optional[Image.Image] = None,
        image_ext: str = ".png",
        page: Optional[int] = None,
        question_no: Optional[int] = None,
        image_filename: Optional[str] = None,
        overwrite_image: bool = True,
        overwrite_existing_entry: bool = True,
    ) -> QuestionItem:
        """
        Creates and adds a new question item, saves diagram image to images/ (overwriting if specified),
        and synchronizes dataset Excel file immediately.
        """
        self.images_dir.mkdir(parents=True, exist_ok=True)
        clean_src = sanitize_source_name(source or self.current_source_name or "biology")
        img_rel_path = ""
        next_idx = len(self.entries) + 1

        # Determine target image filename
        if image_filename:
            unique_name = Path(image_filename).name
        elif overwrite_image:
            ext = image_ext if image_ext.startswith(".") else f".{image_ext}"
            unique_name = f"{clean_src}_{next_idx:03d}{ext.lower()}"
        else:
            unique_name = generate_unique_image_filename(
                self.images_dir,
                source=clean_src,
                ext=image_ext,
                index=next_idx,
                page=page,
                question_no=question_no,
                allow_overwrite=False,
            )

        save_path = self.images_dir / unique_name

        # Save diagram image if provided (overwriting existing file cleanly)
        if diagram_image is not None:
            diagram_image.save(str(save_path))
            img_rel_path = f"images/{unique_name}"
        elif image_filename and save_path.is_file():
            img_rel_path = f"images/{unique_name}"
        else:
            img_rel_path = ""

        # Determine entry ID
        if image_filename:
            entry_id = Path(image_filename).stem
        elif page is not None and question_no is not None:
            entry_id = f"{clean_src}_p{int(page):02d}_q{int(question_no):02d}"
        else:
            entry_id = f"{clean_src}_{next_idx:03d}"

        # Check if ID already exists in entries
        existing_idx = None
        for i, e in enumerate(self.entries):
            if e.id == entry_id:
                existing_idx = i
                break

        if existing_idx is not None and not overwrite_existing_entry:
            suf = 2
            candidate_id = f"{entry_id}_{suf}"
            existing_ids = {e.id for e in self.entries}
            while candidate_id in existing_ids:
                suf += 1
                candidate_id = f"{entry_id}_{suf}"
            entry_id = candidate_id

        item = QuestionItem(
            id=entry_id,
            source=clean_src,
            question=question_text.strip(),
            answer=answer_text.strip(),
            image_path=img_rel_path,
            page=page,
            question_no=question_no,
        )

        if existing_idx is not None and overwrite_existing_entry:
            self.entries[existing_idx] = item
        else:
            self.entries.append(item)

        # Synchronize Excel workbook (fast)
        self.save_excel()

        return item

    def delete_question(self, index: int) -> Optional[QuestionItem]:
        """Deletes question by index, removes its image from disk, and syncs Excel."""
        if not (0 <= index < len(self.entries)):
            return None

        entry = self.entries.pop(index)
        if entry.has_image_file(self.data_dir):
            try:
                img_path = entry.get_image_file_path(self.data_dir)
                if img_path.is_file():
                    img_path.unlink()
            except Exception:
                pass

        try:
            self.save_excel()
        except PermissionError:
            pass

        return entry

    def update_question(self, index: int, question_text: str, answer_text: str) -> bool:
        """Updates text of an existing question and persists changes to Excel."""
        if not (0 <= index < len(self.entries)):
            return False

        entry = self.entries[index]
        entry.question = question_text.strip()
        entry.answer = answer_text.strip()

        try:
            self.save_excel()
        except PermissionError:
            pass
        return True

    def save_excel(self, output_path: Optional[Path] = None) -> Path:
        """
        Generates standard 7-column Excel workbook with embedded thumbnails,
        dynamic row heights, wrap text, and styling.
        """
        dest = output_path or self.excel_file
        dest.parent.mkdir(parents=True, exist_ok=True)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "MMU Dataset"

        ws.append(self.EXCEL_HEADERS)

        # Styling
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        center_align = Alignment(horizontal="center", vertical="center")
        left_align_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

        thin_border = Border(
            left=Side(style="thin", color="D1D5DB"),
            right=Side(style="thin", color="D1D5DB"),
            top=Side(style="thin", color="D1D5DB"),
            bottom=Side(style="thin", color="D1D5DB"),
        )

        for col_i in range(1, len(self.EXCEL_HEADERS) + 1):
            cell = ws.cell(row=1, column=col_i)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border

        ws.row_dimensions[1].height = 28

        col_widths = {
            "A": 20,  # id
            "B": 54,  # question
            "C": 12,  # answer
            "D": 32,  # image
            "E": 30,  # image_path
        }
        for col_let, w in col_widths.items():
            ws.column_dimensions[col_let].width = w

        data_font = Font(name="Segoe UI", size=10)
        bangla_font = Font(name="Kalpurush", size=10)

        for idx, entry in enumerate(self.entries):
            row_idx = idx + 2
            has_img = entry.has_image_file(self.data_dir)
            img_rel = entry.standard_relative_path if has_img else ""

            row_data = [
                entry.id,
                entry.question,
                entry.answer,
                "",  # placeholder for image thumbnail
                img_rel,
            ]
            ws.append(row_data)

            for col_i in range(1, len(self.EXCEL_HEADERS) + 1):
                c = ws.cell(row=row_idx, column=col_i)
                c.border = thin_border
                c.font = bangla_font if col_i in [2, 3] else data_font
                if col_i == 2:
                    c.alignment = left_align_wrap
                elif col_i in [1, 5]:
                    c.alignment = Alignment(horizontal="left", vertical="center")
                else:
                    c.alignment = center_align

            num_lines = max(1, entry.question.count("\n") + 1)
            base_row_h = max(28, num_lines * 18)

            if has_img:
                img_file = entry.get_image_file_path(self.data_dir)
                try:
                    with Image.open(str(img_file)) as p_img:
                        ow, oh = p_img.size
                    max_w, max_h = 190, 100
                    scale = min(max_w / ow, max_h / oh)
                    dw = max(1, int(ow * scale))
                    dh = max(1, int(oh * scale))

                    ex_img = OpenpyxlImage(str(img_file))
                    ex_img.width = dw
                    ex_img.height = dh
                    ws.add_image(ex_img, f"D{row_idx}")

                    ws.row_dimensions[row_idx].height = max(base_row_h, int(dh * 0.75 + 14))
                except Exception:
                    ws.row_dimensions[row_idx].height = base_row_h
            else:
                ws.row_dimensions[row_idx].height = base_row_h

        ws.freeze_panes = "A2"
        wb.save(str(dest))
        return dest

    def export_csv(self, output_path: Optional[Path] = None) -> Path:
        """Exports dataset to UTF-8 BOM CSV."""
        dest = output_path or self.csv_file
        dest.parent.mkdir(parents=True, exist_ok=True)

        with open(dest, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(self.CSV_HEADERS)
            for entry in self.entries:
                has_img = entry.has_image_file(self.data_dir)
                img_p = entry.standard_relative_path if has_img else ""
                writer.writerow([
                    entry.id,
                    entry.question,
                    entry.answer,
                    img_p,
                ])
        return dest

    def export_jsonl(self, output_path: Optional[Path] = None) -> Path:
        """Exports dataset to JSONL."""
        dest = output_path or self.jsonl_file
        dest.parent.mkdir(parents=True, exist_ok=True)

        with open(dest, "w", encoding="utf-8") as f:
            for entry in self.entries:
                has_img = entry.has_image_file(self.data_dir)
                d = {
                    "id": entry.id,
                    "source": entry.source,
                    "question": entry.question,
                    "answer": entry.answer,
                    "image_path": entry.standard_relative_path if has_img else "",
                }
                f.write(json.dumps(d, ensure_ascii=False) + "\n")
        return dest

    def export_html(self, output_path: Optional[Path] = None, title: Optional[str] = None) -> Path:
        """
        Exports dataset to a standalone, responsive, clean HTML preview file.
        Includes questions, Bengali options, answer badges, and diagram visuals.
        """
        src_name = sanitize_source_name(self.current_source_name or "dataset")
        dest = output_path or (self.exports_dir / f"{src_name}.html")
        dest.parent.mkdir(parents=True, exist_ok=True)

        doc_title = title or f"{src_name} - Multimodal MCQ Dataset"

        cards_html = []
        for idx, entry in enumerate(self.entries, start=1):
            has_img = entry.has_image_file(self.data_dir)
            img_html = ""
            if has_img:
                img_path = entry.get_image_file_path(self.data_dir)
                rel_img_src = f"images/{img_path.name}"
                img_html = f'''
                    <div class="q-image-box">
                        <a href="{rel_img_src}" target="_blank">
                            <img src="{rel_img_src}" alt="Diagram for {html.escape(entry.id)}" loading="lazy">
                        </a>
                        <div class="q-image-label">{html.escape(img_path.name)}</div>
                    </div>'''

            ans_html = ""
            if entry.answer:
                ans_html = f'<div class="ans-badge">সঠিক উত্তর: {html.escape(entry.answer)}</div>'

            q_text_escaped = html.escape(entry.question)

            card = f'''
            <article class="question-card" id="{html.escape(entry.id)}">
                <div class="q-header">
                    <span class="q-id">#{idx:03d} &bull; {html.escape(entry.id)}</span>
                    {ans_html}
                </div>
                <div class="q-body">
                    <div class="q-text">{q_text_escaped}</div>
                    {img_html}
                </div>
            </article>'''
            cards_html.append(card)

        all_cards = "\n".join(cards_html)

        html_content = f'''<!DOCTYPE html>
<html lang="bn">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{html.escape(doc_title)}</title>
    <style>
        :root {{
            --primary: #2563EB;
            --bg: #F8FAFC;
            --card: #FFFFFF;
            --border: #E2E8F0;
            --text: #0F172A;
            --muted: #64748B;
            --success: #059669;
            --success-bg: #ECFDF5;
        }}
        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{
            font-family: 'Kalpurush', 'Nirmala UI', 'Vrinda', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: var(--bg);
            color: var(--text);
            line-height: 1.6;
            padding: 24px;
        }}
        .container {{ max-width: 960px; margin: 0 auto; }}
        header {{
            background: #1E293B;
            color: #FFFFFF;
            padding: 24px 28px;
            border-radius: 12px;
            margin-bottom: 24px;
            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
        }}
        header h1 {{ font-size: 24px; font-weight: 700; margin-bottom: 8px; }}
        header .meta {{ color: #94A3B8; font-size: 14px; display: flex; flex-wrap: wrap; gap: 16px; margin-top: 8px; }}
        .stats-badge {{
            display: inline-block;
            background: #2563EB;
            color: #FFFFFF;
            font-size: 13px;
            font-weight: 600;
            padding: 4px 10px;
            border-radius: 6px;
            margin-top: 12px;
        }}
        .question-card {{
            background: var(--card);
            border: 1px solid var(--border);
            border-radius: 10px;
            padding: 20px;
            margin-bottom: 18px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            page-break-inside: avoid;
        }}
        .q-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid var(--border);
            padding-bottom: 10px;
            margin-bottom: 14px;
        }}
        .q-id {{
            font-weight: 700;
            color: var(--primary);
            font-size: 14px;
            background: #EFF6FF;
            padding: 3px 8px;
            border-radius: 4px;
        }}
        .ans-badge {{
            background: var(--success-bg);
            color: var(--success);
            font-weight: 700;
            font-size: 14px;
            padding: 4px 12px;
            border-radius: 20px;
            border: 1px solid #A7F3D0;
        }}
        .q-body {{
            display: flex;
            gap: 20px;
            flex-wrap: wrap;
        }}
        .q-text {{
            flex: 1 1 500px;
            white-space: pre-wrap;
            font-size: 15px;
            line-height: 1.7;
        }}
        .q-image-box {{
            flex: 0 0 280px;
            max-width: 320px;
            text-align: center;
        }}
        .q-image-box img {{
            max-width: 100%;
            height: auto;
            border: 1px solid var(--border);
            border-radius: 6px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            background: #FFFFFF;
        }}
        .q-image-label {{
            font-size: 11px;
            color: var(--muted);
            margin-top: 4px;
        }}
        @media print {{
            body {{ padding: 0; background: #FFF; }}
            header {{ background: #333; color: #FFF; }}
            .question-card {{ border: 1px solid #CCC; box-shadow: none; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>{html.escape(doc_title)}</h1>
            <div class="meta">
                <span>📁 Source: <strong>{html.escape(src_name)}</strong></span>
                <span>📊 Total Questions: <strong>{len(self.entries)}</strong></span>
            </div>
            <div class="stats-badge">Generated by MMU Dataset Builder v2</div>
        </header>
        <main>
{all_cards}
        </main>
    </div>
</body>
</html>'''

        with open(dest, "w", encoding="utf-8") as f:
            f.write(html_content)

        return dest

    def export_zip_package(
        self,
        source_name: Optional[str] = None,
        output_dir: Optional[Path] = None,
    ) -> Path:
        """
        Generates CSV, JSONL, HTML, and current Excel workbook,
        then bundles everything along with referenced images into:
        exports/{sourcename}_{timestamp}.zip
        """
        src = sanitize_source_name(source_name or self.current_source_name or "dataset")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_filename = f"{src}_{timestamp}.zip"

        dest_dir = output_dir or self.exports_dir
        dest_dir.mkdir(parents=True, exist_ok=True)
        zip_path = dest_dir / zip_filename

        # Save excel first
        self.save_excel()

        temp_export_dir = dest_dir / f"_tmp_{src}_{timestamp}"
        temp_export_dir.mkdir(parents=True, exist_ok=True)

        try:
            csv_path = self.export_csv(temp_export_dir / f"{src}.csv")
            jsonl_path = self.export_jsonl(temp_export_dir / f"{src}.jsonl")
            html_path = self.export_html(temp_export_dir / f"{src}.html")

            with zipfile.ZipFile(str(zip_path), "w", compression=zipfile.ZIP_DEFLATED) as zf:
                # Add Excel workbook
                if self.excel_file.is_file():
                    zf.write(str(self.excel_file), arcname=f"{src}.xlsx")

                # Add CSV, JSONL, HTML
                zf.write(str(csv_path), arcname=f"{src}.csv")
                zf.write(str(jsonl_path), arcname=f"{src}.jsonl")
                zf.write(str(html_path), arcname=f"{src}.html")

                # Add referenced images
                added_images = set()
                for entry in self.entries:
                    if entry.has_image_file(self.data_dir):
                        img_file = entry.get_image_file_path(self.data_dir)
                        if img_file.name not in added_images:
                            zf.write(str(img_file), arcname=f"images/{img_file.name}")
                            added_images.add(img_file.name)

            return zip_path
        finally:
            shutil.rmtree(temp_export_dir, ignore_errors=True)

    @property
    def total_count(self) -> int:
        return len(self.entries)
