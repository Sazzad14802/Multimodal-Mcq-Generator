"""
MMU App-2 Gemini API Client
Direct REST integration for extracting Bengali MCQ question strings from screenshots.
Zero external library dependencies (uses Python standard library urllib & json).
"""

import base64
import io
import json
import os
from pathlib import Path
from typing import Optional, Tuple
import urllib.error
import urllib.request
from PIL import Image


DEFAULT_VISION_PROMPT = """Extract the MCQ from the image.

LANGUAGE AND TEXT PRESERVATION:
- Preserve the original language of every word.
- Bengali text must remain Bengali.
- English text must remain English.
- Do NOT translate Bengali ↔ English.
- Preserve English capitalization exactly as shown.
- Preserve numbers as they appear.

MATHEMATICAL NOTATION:
- Do NOT use LaTeX.
- Do NOT use $ delimiters.
- Do NOT output LaTeX commands such as \\vec, \\frac, \\sqrt, \\alpha, \\beta, etc.
- Convert mathematical notation into visible Unicode symbols wherever possible.
- Use √ for square roots.
- Use actual mathematical symbols such as ×, ÷, ±, ≤, ≥, ≠, θ, α, β, π, etc.
- Represent vectors using a visible vector notation such as A⃗.
- Represent fractions as a/b when a single Unicode fraction is not practical.
- Preserve the mathematical meaning exactly.

MCQ STRUCTURE:
- Extract the question separately.
- Extract each option separately.
- If textual options are present, format as:
<Question text>

ক) <option 1 text>
খ) <option 2 text>
গ) <option 3 text>
ঘ) <option 4 text>
- Do not include "ক)", "খ)", "গ)", or "ঘ)" inside the option text.
- CRITICAL: If the answer options themselves are DIAGRAMS, GRAPHS, FIGURES, or IMAGES (and not plain text/equations), DO NOT append the options to the question string. Only return the question text itself.
- Do not solve the question.
- Do not add explanations.
- Do not add or remove information.
- Preserve the original text as closely as possible.
- Do NOT output markdown code blocks (no ```).

Return only the structured output.
"""

DEFAULT_TEXT_FURNISH_PROMPT = """Extract the MCQ from the text.
If the input text is raw, garbled, or legacy ANSI/Bijoy encoded (e.g. 'GKwU †evgviæ wegvb f‚wg...'), decode it into accurate standard Bengali Unicode.

LANGUAGE AND TEXT PRESERVATION:
- Preserve the original language of every word.
- Bengali text must remain Bengali.
- English text must remain English.
- Do NOT translate Bengali ↔ English.
- Preserve English capitalization exactly as shown.
- Preserve numbers as they appear.

MATHEMATICAL NOTATION:
- Do NOT use LaTeX.
- Do NOT use $ delimiters.
- Do NOT output LaTeX commands such as \\vec, \\frac, \\sqrt, \\alpha, \\beta, etc.
- Convert mathematical notation into visible Unicode symbols wherever possible.
- Use √ for square roots.
- Use actual mathematical symbols such as ×, ÷, ±, ≤, ≥, ≠, θ, α, β, π, etc.
- Represent vectors using a visible vector notation such as A⃗.
- Represent fractions as a/b when a single Unicode fraction is not practical.
- Preserve the mathematical meaning exactly.

MCQ STRUCTURE:
- Extract the question separately.
- Extract each option separately.
- If textual options are present, format as:
<Question text>

ক) <option 1 text>
খ) <option 2 text>
গ) <option 3 text>
ঘ) <option 4 text>
- Do not include "ক)", "খ)", "গ)", or "ঘ)" inside the option text.
- If no options are present in the input, or if the options are diagrams/figures, do NOT invent or append options; output only the clean question text.
- Do not solve the question.
- Do not add explanations.
- Do not add or remove information.
- Preserve the original text as closely as possible.
- Do NOT output markdown code blocks (no ```).

Return only the structured output.
"""

# Backward compatibility alias
DEFAULT_PROMPT = DEFAULT_VISION_PROMPT


class GeminiClient:
    """
    Handles communication with Google's Gemini API for multimodal question extraction
    and raw/Bijoy text conversion & furnishing.
    """

    CONFIG_FILE_NAME = ".gemini_config.json"

    def __init__(self, base_dir: Optional[Path] = None, api_key: Optional[str] = None):
        self.base_dir = Path(base_dir).resolve() if base_dir else Path.cwd()
        self.config_path = self.base_dir / self.CONFIG_FILE_NAME
        self.api_key: str = api_key or self.load_saved_api_key()
        self.primary_model: str = "gemini-3.5-flash-lite"
        self.fallback_models = [
            "gemini-flash-lite-latest",
            "gemini-2.5-flash-lite",
            "gemini-3.8-flash",
            "gemini-flash-latest",
        ]
        self.last_used_model: Optional[str] = None
        self._load_saved_model()

    def _load_saved_model(self) -> None:
        """Loads saved model preference if present."""
        if self.config_path.is_file():
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    saved_model = data.get("model", "").strip()
                    if saved_model:
                        self.primary_model = saved_model
            except Exception:
                pass

    def load_saved_api_key(self) -> str:
        """Loads API key from config file or environment variable."""
        # 1. Check environment variable
        env_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
        if env_key and env_key.strip():
            return env_key.strip()

        # 2. Check local config file
        if self.config_path.is_file():
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    saved = data.get("api_key", "").strip()
                    if saved:
                        return saved
            except Exception:
                pass

        return ""

    def save_api_key(self, key: str, model: Optional[str] = None) -> None:
        """Persists API key and selected model to local config file."""
        self.api_key = key.strip()
        if model:
            self.primary_model = model.strip()

        data = {}
        if self.config_path.is_file():
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                data = {}

        data["api_key"] = self.api_key
        data["model"] = self.primary_model
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)

    def is_configured(self) -> bool:
        """Returns True if a non-empty API key is present."""
        return bool(self.api_key and len(self.api_key.strip()) > 10)

    def image_to_base64_png(self, image: Image.Image) -> str:
        """Converts PIL Image to base64-encoded PNG string."""
        buf = io.BytesIO()
        # Convert RGBA/P to RGB if needed or preserve RGBA for PNG
        if image.mode not in ["RGB", "RGBA"]:
            image = image.convert("RGB")
        image.save(buf, format="PNG")
        buf.seek(0)
        return base64.b64encode(buf.read()).decode("utf-8")

    def _send_gemini_request(
        self,
        payload: dict,
        operation_name: str = "OCR fetch",
        timeout: int = 30,
    ) -> Tuple[bool, str]:
        """
        Sends request payload to Gemini REST API across primary model and fallbacks,
        logging execution and returning (success, text_or_error).
        """
        if not self.is_configured():
            return (
                False,
                "Gemini API Key is not configured. Please enter your API key using the '🔑 API Key' button above.",
            )

        json_bytes = json.dumps(payload).encode("utf-8")
        models_to_try = [self.primary_model] + [m for m in self.fallback_models if m != self.primary_model]
        last_error = ""

        print(f"\n==================================================", flush=True)
        print(f"[Gemini API] Initiating {operation_name}...", flush=True)
        print(f"[Gemini API] Primary model requested: '{self.primary_model}'", flush=True)

        for attempt, model in enumerate(models_to_try, start=1):
            print(f"[Gemini API] Attempt {attempt}/{len(models_to_try)}: Fetching from model '{model}'...", flush=True)
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={self.api_key}"
            req = urllib.request.Request(
                url,
                data=json_bytes,
                headers={"Content-Type": "application/json"},
                method="POST",
            )

            try:
                with urllib.request.urlopen(req, timeout=timeout) as response:
                    res_body = response.read().decode("utf-8")
                    data = json.loads(res_body)

                    # Parse response text
                    candidates = data.get("candidates", [])
                    if candidates and "content" in candidates[0]:
                        parts = candidates[0]["content"].get("parts", [])
                        if parts and "text" in parts[0]:
                            text = parts[0]["text"].strip()
                            # Clean any residual markdown backticks if Gemini added them
                            if text.startswith("```"):
                                lines = text.splitlines()
                                if lines and lines[0].startswith("```"):
                                    lines = lines[1:]
                                if lines and lines[-1].startswith("```"):
                                    lines = lines[:-1]
                                text = "\n".join(lines).strip()
                            self.last_used_model = model
                            print(f"[Gemini API] SUCCESS! Response successfully fetched using model: '{model}'", flush=True)
                            print(f"==================================================\n", flush=True)
                            return (True, text)

                    print(f"[Gemini API] Warning: Model '{model}' returned empty or blocked content response.", flush=True)
                    return (False, "Gemini returned empty or blocked content response.")

            except urllib.error.HTTPError as he:
                err_text = ""
                try:
                    err_json = json.loads(he.read().decode("utf-8"))
                    err_msg = err_json.get("error", {}).get("message", "")
                    err_text = f"HTTP {he.code}: {err_msg}"
                except Exception:
                    err_text = f"HTTP {he.code}: {he.reason}"

                last_error = err_text
                print(f"[Gemini API] Model '{model}' failed: {err_text}", flush=True)
                # If 404 model not found / deprecated, try next fallback model
                if he.code == 404:
                    print(f"[Gemini API] Model '{model}' not accessible on this account/endpoint. Trying fallback...", flush=True)
                    continue
                # For auth or quota errors, break early
                break

            except urllib.error.URLError as ue:
                print(f"[Gemini API] Network connection failed for model '{model}': {ue.reason}", flush=True)
                return (False, f"Network connection failed: {ue.reason}")
            except Exception as ex:
                print(f"[Gemini API] Request failed for model '{model}': {ex}", flush=True)
                return (False, f"Request failed: {ex}")

        print(f"[Gemini API] FAILED all attempts. Last error: {last_error}", flush=True)
        print(f"==================================================\n", flush=True)
        return (False, f"Gemini extraction failed: {last_error}")

    def extract_question(
        self,
        image: Image.Image,
        custom_prompt: Optional[str] = None,
        timeout: int = 30,
    ) -> Tuple[bool, str]:
        """
        Sends the question screenshot to Gemini API and returns (success, extracted_text_or_error).
        """
        prompt = custom_prompt or DEFAULT_VISION_PROMPT
        b64_data = self.image_to_base64_png(image)

        # Payload structure for Gemini REST API
        payload = {
            "contents": [
                {
                    "parts": [
                        {"text": prompt},
                        {
                            "inline_data": {
                                "mime_type": "image/png",
                                "data": b64_data,
                            }
                        },
                    ]
                }
            ],
            "generationConfig": {
                "temperature": 0.1,
                "topP": 0.95,
                "maxOutputTokens": 1024,
            },
        }

        return self._send_gemini_request(payload, operation_name="Image Screenshot OCR", timeout=timeout)

    def furnish_question_text(
        self,
        raw_text: str,
        custom_prompt: Optional[str] = None,
        timeout: int = 30,
    ) -> Tuple[bool, str]:
        """
        Sends raw, garbled, or legacy ANSI/Bijoy question text to Gemini API to convert & furnish
        into clean standard Unicode Bengali with structured MCQ options.
        """
        if not raw_text or not raw_text.strip():
            return (False, "Input text is empty. Please enter or paste question text.")

        prompt = custom_prompt or DEFAULT_TEXT_FURNISH_PROMPT
        full_text_input = f"{prompt}\n\nRAW QUESTION TEXT TO FURNISH:\n\"\"\"\n{raw_text.strip()}\n\"\"\""

        payload = {
            "contents": [
                {
                    "parts": [
                        {"text": full_text_input}
                    ]
                }
            ],
            "generationConfig": {
                "temperature": 0.1,
                "topP": 0.95,
                "maxOutputTokens": 1024,
            },
        }

        return self._send_gemini_request(payload, operation_name="Raw Text Furnishing", timeout=timeout)
