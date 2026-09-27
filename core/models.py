"""
MMU App-2 Data Models
Dataclass definitions and helpers for question items.
"""

from dataclasses import dataclass
from pathlib import Path
import re
from typing import Dict, Any, Optional


def sanitize_source_name(name: str) -> str:
    """
    Sanitizes a string to be safely used as a filename and ID prefix.
    Replaces non-alphanumeric/hyphen/underscore chars with underscores.
    """
    clean = Path(name).stem
    clean = re.sub(r"[^\w\-]", "_", clean)
    clean = re.sub(r"_+", "_", clean).strip("_")
    return clean or "source"


def generate_unique_image_filename(
    images_dir: Path,
    source: str,
    ext: str = ".png",
    index: Optional[int] = None,
    page: Optional[int] = None,
    question_no: Optional[int] = None,
    allow_overwrite: bool = False,
) -> str:
    """
    Generates unique image filename format:
    <source_name>_<index:03d>.<ext>
    If file already exists, appends _2, _3 without overwriting, unless allow_overwrite is True.
    """
    clean_src = sanitize_source_name(source)
    if not ext.startswith("."):
        ext = f".{ext}"
    ext = ext.lower()

    if page is not None and question_no is not None:
        base_name = f"{clean_src}_p{int(page):02d}_q{int(question_no):02d}"
    elif index is not None:
        base_name = f"{clean_src}_{int(index):03d}"
    else:
        base_name = f"{clean_src}_001"

    candidate = f"{base_name}{ext}"
    if allow_overwrite or not (images_dir / candidate).exists():
        return candidate

    suffix = 2
    while (images_dir / f"{base_name}_{suffix}{ext}").exists():
        suffix += 1

    return f"{base_name}_{suffix}{ext}"


@dataclass
class QuestionItem:
    """Represents a single multimodal MCQ question item."""
    id: str
    source: str
    question: str
    answer: str = ""
    image_path: str = ""
    page: Optional[int] = None
    question_no: Optional[int] = None

    @property
    def image_filename(self) -> str:
        """Returns image filename."""
        if self.image_path:
            return Path(self.image_path).name
        clean_src = sanitize_source_name(self.source)
        if self.page is not None and self.question_no is not None:
            return f"{clean_src}_p{int(self.page):02d}_q{int(self.question_no):02d}.png"
        return f"{self.id}.png"

    @property
    def standard_relative_path(self) -> str:
        """Returns relative path: images/<filename>"""
        return f"images/{self.image_filename}"

    def get_image_file_path(self, data_dir: Path) -> Path:
        """Returns absolute path to the image file, ensuring images directory exists."""
        images_dir = data_dir / "images"
        images_dir.mkdir(parents=True, exist_ok=True)

        if self.image_path:
            direct_path = data_dir / self.image_path
            if direct_path.is_file() and direct_path.stat().st_size > 0:
                return direct_path
            by_name = images_dir / Path(self.image_path).name
            if by_name.is_file() and by_name.stat().st_size > 0:
                return by_name

        candidate = images_dir / self.image_filename
        if candidate.is_file() and candidate.stat().st_size > 0:
            return candidate

        return candidate

    def has_image_file(self, data_dir: Path) -> bool:
        """Checks if the image file exists on disk."""
        path = self.get_image_file_path(data_dir)
        return path.is_file() and path.stat().st_size > 0

    def to_dict(self) -> Dict[str, Any]:
        """Converts to dictionary representation."""
        d = {
            "id": self.id,
            "source": self.source,
            "question": self.question,
            "answer": self.answer,
            "image_path": self.image_path,
        }
        if self.page is not None:
            d["page"] = self.page
        if self.question_no is not None:
            d["question_no"] = self.question_no
        return d

    @classmethod
    def from_dict(cls, data: Dict[str, Any], default_source: str = "source") -> "QuestionItem":
        """Instantiates QuestionItem from dictionary."""
        source = str(data.get("source") or default_source)
        clean_src = sanitize_source_name(source)
        page = int(data["page"]) if data.get("page") is not None and str(data["page"]).strip() else None
        qno = int(data["question_no"]) if data.get("question_no") is not None and str(data["question_no"]).strip() else None

        entry_id = str(data.get("id") or "").strip()
        if not entry_id:
            if page is not None and qno is not None:
                entry_id = f"{clean_src}_p{page:02d}_q{qno:02d}"
            else:
                entry_id = f"{clean_src}_001"

        ans = str(data.get("answer") or data.get("answer_option") or "").strip()
        img = str(data.get("image_path") or data.get("image") or "").strip()

        return cls(
            id=entry_id,
            source=source,
            question=str(data.get("question") or ""),
            answer=ans,
            image_path=img,
            page=page,
            question_no=qno,
        )
