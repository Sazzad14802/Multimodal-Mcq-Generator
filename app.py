"""
MMU App-2: Gemini-Powered Multimodal MCQ Dataset Builder
Launcher Script.
"""

import sys
from pathlib import Path
import tkinter as tk

# Ensure app-2 directory is in sys.path
app2_dir = Path(__file__).resolve().parent
if str(app2_dir) not in sys.path:
    sys.path.insert(0, str(app2_dir))

from core.dataset_manager import DatasetManager
from core.gemini_client import GeminiClient
from ui.main_window import App2MainWindow


def set_high_dpi_awareness():
    """Enables Windows High-DPI awareness for crisp text and rendering."""
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass


def main():
    set_high_dpi_awareness()

    root = tk.Tk()

    # Initialize data manager with default app-2/data directory
    data_dir = app2_dir / "data"
    manager = DatasetManager(data_dir=data_dir)

    # Initialize Gemini client with local config storage
    gemini_client = GeminiClient(base_dir=app2_dir)

    app = App2MainWindow(root, manager, gemini_client)
    root.mainloop()


if __name__ == "__main__":
    main()
