import streamlit as st

# Build marker used in Auto Detect cache keys so code updates cannot reuse
# stale detected-field selections from an older engine version.
AUTO_DETECT_ENGINE_VERSION = "2026-09-13-COMPARISON-ENGINE-EVIDENCE-LINKED-VISUAL-14"
import pandas as pd
import fitz
import re
import unicodedata
import io
from pathlib import Path
from io import BytesIO

from PIL import Image, ImageDraw
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage

def _apply_tool_css():
    # =========================================================
    # DARK UI
    # =========================================================

    st.markdown(
        """
        <style>

        html,
        body,
        [data-testid="stAppViewContainer"],
        [data-testid="stApp"],
        .stApp,
        .main,
        [data-testid="stMain"] {
            background-color: #0e1117 !important;
            color: #ffffff !important;
        }

        [data-testid="stHeader"] {
            background-color: #0e1117 !important;
        }

        .stApp,
        .stApp p,
        .stApp label,
        .stApp span,
        .stApp div {
            color: #ffffff;
        }

        .main-title {
            color: #ffffff !important;
            font-size: 34px;
            font-weight: 700;
            margin-top: 5px;
            margin-bottom: 4px;
        }

        .sub-title {
            color: #b8c0cc !important;
            font-size: 15px;
            margin-bottom: 30px;
        }

        .section-title {
            color: #ffffff !important;
            font-size: 20px;
            font-weight: 700;
            margin-top: 12px;
            margin-bottom: 10px;
        }

        [data-testid="stFileUploader"] {
            background-color: #161b22 !important;
            border: 1px solid #4b5563 !important;
            border-radius: 12px !important;
            padding: 8px !important;
        }

        [data-testid="stFileUploaderDropzone"] {
            background-color: #161b22 !important;
            border: 1px solid #4b5563 !important;
            border-radius: 10px !important;
        }

        [data-testid="stFileUploaderDropzoneInstructions"] {
            color: #ffffff !important;
        }

        [data-testid="stFileUploaderDropzoneInstructions"] span {
            color: #ffffff !important;
        }

        [data-testid="stFileUploader"] button {
            background-color: #111827 !important;
            color: #ffffff !important;
            border: 1px solid #6b7280 !important;
            border-radius: 8px !important;
        }

        [data-testid="stFileUploader"] button:hover {
            background-color: #1f2937 !important;
            color: #ffffff !important;
        }

        [data-baseweb="select"] > div {
            background-color: #161b22 !important;
            color: #ffffff !important;
            border: 1px solid #4b5563 !important;
            border-radius: 10px !important;
        }

        [data-baseweb="select"] input {
            color: #ffffff !important;
        }

        [data-baseweb="select"] span {
            color: #ffffff !important;
        }

        [data-baseweb="popover"] {
            background-color: #161b22 !important;
        }

        [role="option"] {
            background-color: #161b22 !important;
            color: #ffffff !important;
        }

        [role="option"]:hover {
            background-color: #263241 !important;
        }

        [data-baseweb="tag"] {
            background-color: #2563eb !important;
            color: #ffffff !important;
        }

        [data-baseweb="tag"] span {
            color: #ffffff !important;
        }

        div.stButton > button {
            background-color: #2196F3 !important;
            color: #ffffff !important;
            border: 2px solid #000000 !important;
            border-radius: 12px !important;
            font-size: 18px !important;
            font-weight: 700 !important;
            height: 54px !important;
            width: 100% !important;
            box-shadow: none !important;
        }

        div.stButton > button:hover {
            background-color: #1976D2 !important;
            color: #ffffff !important;
            border: 2px solid #000000 !important;
        }

        div.stDownloadButton > button {
            background-color: #1f2937 !important;
            color: #ffffff !important;
            border: 1px solid #6b7280 !important;
            border-radius: 10px !important;
            font-weight: 600 !important;
        }

        div.stDownloadButton > button:hover {
            background-color: #374151 !important;
            color: #ffffff !important;
        }

        [data-testid="stDataFrame"] {
            border: 1px solid #374151 !important;
            border-radius: 10px !important;
        }

        [data-testid="stMetric"] {
            background-color: #161b22 !important;
            border: 1px solid #374151 !important;
            border-radius: 10px !important;
            padding: 12px !important;
        }

        [data-testid="stMetricLabel"] {
            color: #b8c0cc !important;
        }

        [data-testid="stMetricValue"] {
            color: #ffffff !important;
        }

        hr {
            border-color: #30363d !important;
        }

        .stCaption {
            color: #9ca3af !important;
        }

        [data-testid="stAlert"] {
            border-radius: 10px !important;
        }

        [data-testid="stSpinner"] {
            color: #ffffff !important;
        }

        </style>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# OPTIMIZED COMPARISON ENGINE
# =========================================================

import pandas as pd
import re
import unicodedata
import io
from PIL import Image


# =========================================================
# BASIC DATA HELPERS
# =========================================================

def is_blank_value(value):
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    return str(value).strip() == ""


def normalize_text(text):
    if is_blank_value(text):
        return ""

    value = unicodedata.normalize("NFKC", str(text))
    value = value.casefold()

    value = value.replace("\u200b", "").replace("\ufeff", "")
    value = value.replace("’", "'").replace("`", "'")
    value = value.replace("–", "-").replace("—", "-")
    value = value.replace("\r", " ").replace("\n", " ")

    # Common PDF bullet extraction artefact.
    value = re.sub(r"(^|\s)n(?=\s)", " ", value)

    # Treat punctuation/separators as spacing differences.
    value = re.sub(r"[,.;:|/\\]+", " ", value)
    value = re.sub(r"-+", " ", value)
    value = re.sub(r"[^\w%#'\s]", " ", value, flags=re.UNICODE)

    # Apostrophe differences should not create a mismatch.
    value = value.replace("'", "")
    value = re.sub(r"\s+", " ", value).strip()
    return value


def normalize_symbol_text(text):
    """Normalize a custom-font symbol keystroke without changing its code point.

    The symbol itself is the data. Custom font workflows may use Unicode
    private-use characters or ordinary keystrokes mapped to custom glyphs.
    NFKC/case-fold normalization can change or erase the very character we are
    trying to validate, so this helper removes only whitespace and invisible
    zero-width markers.
    """
    if is_blank_value(text):
        return ""

    value = str(text)
    value = value.replace("\u200b", "").replace("\ufeff", "")
    value = re.sub(r"\s+", "", value)
    return value


def _symbol_contains_private_or_nontext_glyph(text):
    """Return True when a value contains a glyph/code point safer to match by exact substring."""
    if is_blank_value(text):
        return False

    for char in str(text):
        code = ord(char)
        category = unicodedata.category(char)
        # Basic/supplementary private-use areas plus Unicode symbol/math/mark
        # categories are typical for custom-font glyph keys.
        if (0xE000 <= code <= 0xF8FF) or (0xF0000 <= code <= 0xFFFFD) or (
            category.startswith("S") or category.startswith("M")
        ):
            return True
    return False


def _symbol_text_matches(expected, actual):
    """Exact custom-font keystroke matching without unsafe fuzzy matching.

    The symbol itself is the data. The complete normalized character sequence
    is checked first. For custom/private-use glyphs, an exact contained sequence
    is safe. For ordinary ASCII keys used by a custom font (for example 'A'),
    matching is token-based against the ORIGINAL text so that 'A' cannot match
    the 'A' inside 'MADE'.
    """
    expected_raw = normalize_symbol_text(expected)
    actual_text = str(actual or "")
    actual_raw = normalize_symbol_text(actual_text)

    if not expected_raw or not actual_raw:
        return False

    if expected_raw == actual_raw:
        return True

    if _symbol_contains_private_or_nontext_glyph(expected_raw):
        return expected_raw in actual_raw

    # Important: use the original whitespace/punctuation boundaries here.
    # normalize_symbol_text removes spaces, which would turn 'A V' into 'AV'
    # and incorrectly prevent the custom key 'A' from matching as its own unit.
    units = [
        normalize_symbol_text(unit)
        for unit in re.split(r"[^A-Za-z0-9_]+", actual_text)
        if unit
    ]
    return expected_raw in units


def compact_text(text):
    return normalize_text(text).replace(" ", "")


def tokenize(text):
    value = normalize_text(text)
    return value.split() if value else []


def normalize_numeric(value):
    if is_blank_value(value):
        return None
    text = normalize_text(value)
    # Keep only a clean integer/decimal when the value is numeric.
    match = re.fullmatch(r"[-+]?\d+(?:\.\d+)?", text)
    if not match:
        return None
    try:
        number = float(match.group(0))
        if number.is_integer():
            return str(int(number))
        return str(number).rstrip("0").rstrip(".")
    except Exception:
        return None


def get_available_fields(df):
    """Only columns containing at least one populated cell are selectable."""
    result = []
    for column in df.columns:
        if df[column].map(lambda value: not is_blank_value(value)).any():
            result.append(str(column))
    return result


def load_excel(file):
    file.seek(0)
    df = pd.read_excel(file, header=0)
    df.columns = [str(column).strip() for column in df.columns]
    return df


# =========================================================
# FIELD CLASSIFICATION
# =========================================================

ADMIN_FIELD_PATTERNS = (
    r"^sr$",
    r"^sr\.?\s*no\.?$",
    r"^serial",
    r"^job\s*(?:no|number)?$",
    r"^order\s*(?:no|number|date)?$",
    r"^ticket\s*(?:no|number)?$",
    r"^created",
    r"^modified",
    r"^timestamp",
)


def get_field_region(field_name):
    original = str(field_name).casefold()
    normalized = normalize_text(field_name)
    compact = re.sub(r"[^a-z0-9]", "", original)

    if "_en" in original or compact.endswith("en") or "english" in normalized:
        return "EN"
    if "_fr" in original or compact.endswith("fr") or "french" in normalized or "canada" in normalized:
        return "FR"
    if "_sp" in original or compact.endswith("sp") or "spanish" in normalized or "espanol" in normalized:
        return "SP"

    if "_ca" in original or compact.endswith("ca"):
        if (
            compact.startswith("wc")
            or compact.startswith("fib")
            or compact.startswith("min")
            or "care" in compact
            or "wash" in compact
            or "fiber" in compact
            or "fibre" in compact
            or "content" in compact
            or "composition" in compact
            or "coo" in compact
        ):
            return "FR"

    return ""


def get_field_type(field_name):
    field = normalize_text(field_name)
    compact = field.replace(" ", "").replace("_", "").replace("-", "")

    # Explicit sequence fields such as OSZ1, OSZ2 ...
    if re.fullmatch(r"osz\d+", compact):
        return "OSZ"

    if "symbol" in compact or compact in {"caremark", "caresymbol", "washsymbol"}:
        return "SYMBOL"

    # Barcode / GTIN family fields.  Keep this ahead of IDENTIFIER so UPC/EAN/GTIN
    # columns get the barcode-specific normalization rather than identifier-prefix
    # logic.  This does not affect ordinary item/style/supplier identifiers.
    if (
        "barcode" in compact
        or "barcodenumber" in compact
        or compact in {"upc", "upca", "ean", "ean8", "ean13", "ean14", "gtin", "gtin8", "gtin12", "gtin13", "gtin14", "jan", "isbn", "itf"}
        or re.search(r"(?:^|(?:_|-|\s))(?:upc|ean|gtin|jan|isbn|itf)(?:\d+)?(?:$|(?:_|-|\s))", str(field_name).casefold())
    ):
        return "BARCODE"

    if (
        compact == "rn"
        or "rnno" in compact
        or "rnnumber" in compact
        or "registrationnumber" in compact
        or "companyrn" in compact
        or compact.startswith("rn")
    ):
        return "RN"

    # Canadian registration number fields are distinct from RN but use the
    # same structured-number evidence. Keep CA_Number from being classified as
    # generic text.
    if compact in {"ca", "canumber", "caregistrationnumber", "registrationcanumber"} or "canumber" in compact:
        return "CA"

    if "productionmark" in compact or "prodmark" in compact or compact == "production":
        return "PRODUCTION_MARK"

    if compact == "iso" or compact.startswith("iso"):
        return "IDENTIFIER"

    if (
        "sku" in compact
        or "itemcode" in compact
        or "itemnumber" in compact
        or "itemno" in compact
        or "stylecode" in compact
        or compact == "style"
        or "productcode" in compact
        or "supwsp" in compact
        or "supplier" in compact
        or "vendorid" in compact
        or "vendorcode" in compact
    ):
        return "IDENTIFIER"

    if "batch" in compact or "lotnumber" in compact or "lotno" in compact or compact == "lot":
        return "BATCH"

    if (
        "quantity" in compact
        or compact == "qty"
        or "units" in compact
        or "pieces" in compact
        or compact == "pcs"
    ):
        return "QUANTITY"

    # Country of Origin / Made-In field families.
    # Recognize semantic schema conventions (not individual job columns), so
    # fields such as MIN_EN / MIN_SP are handled without a hardcoded list.
    if (
        "coo" in compact
        or "countryoforigin" in compact
        or "countryorigin" in compact
        or "madein" in compact
        or compact == "origin"
        or compact.startswith("min")
    ):
        return "COO"

    # Fiber / Fabric / Content field families.
    # Common production schemas use FIB_* as shorthand for visible fiber
    # composition. This is intentionally a semantic family rule, not a list of
    # individual field names.
    if (
        "fiber" in compact
        or "fibre" in compact
        or "fabric" in compact
        or "content" in compact
        or "composition" in compact
        or "compodsc" in compact
        or "lhcompodsc" in compact
        or "fabrication" in compact
        or "material" in compact
        or compact.startswith("fib")
    ):
        return "CONTENT"

    # Care/wash instruction field families. WC_* is a common shorthand for
    # wash-care text.
    if (
        "care" in compact
        or "wash" in compact
        or "washing" in compact
        or "laundry" in compact
        or "instruction" in compact
        or compact.startswith("wc")
    ):
        return "CARE"

    if (
        "size" in compact
        or "sizeline" in compact
        or "alpha" in compact
        or "waist" in compact
        or "inseam" in compact
        or compact == "fit"
        or re.fullmatch(r"s\d+", compact)
    ):
        return "SIZE"

    if "brand" in compact:
        return "BRAND"

    if "color" in compact or "colour" in compact:
        return "COLOR"

    if "gender" in compact:
        return "GENDER"

    if "attribute" in compact or "technology" in compact or "feature" in compact:
        return "ATTRIBUTE"

    return "GENERAL"


def is_admin_field(field_name):
    normalized = normalize_text(field_name)
    return any(re.match(pattern, normalized) for pattern in ADMIN_FIELD_PATTERNS)


# =========================================================
# PDF/OCR EXTRACTION
# =========================================================

def _usable_text(text):
    if not text or not str(text).strip():
        return False
    alnum = re.sub(r"[^\w%#]", "", str(text), flags=re.UNICODE)
    return len(alnum) >= 6


def _text_quality_score(text):
    """Score extracted artwork text so OCR can be preferred over weak PDF text layers."""
    text = str(text or "")
    if not text.strip():
        return 0

    alnum = len(re.findall(r"[A-Za-z0-9]", text))
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    useful_lines = sum(1 for line in lines if len(re.sub(r"[^A-Za-z0-9%#]", "", line)) >= 2)
    numeric_runs = len(re.findall(r"(?<![A-Za-z0-9])\d+(?![A-Za-z0-9])", text))

    return alnum + useful_lines * 8 + numeric_runs * 3


def _ocr_image_with_data(image):
    """
    Multi-pass OCR for small artwork text.

    Returns:
        primary_text, primary_word_boxes, language, supplemental_text

    primary_text is reconstructed from the strongest OCR pass using physical
    word coordinates so the visual reading order is preserved. supplemental_text
    contains additional unique lines from weaker passes and is used only as
    secondary evidence for Auto Detect.
    """

    try:
        import pytesseract
        from pytesseract import Output
    except ImportError as exc:
        raise RuntimeError(
            "OCR support is not installed. Add pytesseract to requirements.txt."
        ) from exc

    try:
        from PIL import ImageOps, ImageEnhance, ImageFilter
    except Exception:
        ImageOps = ImageEnhance = ImageFilter = None

    work_image = image.convert("RGB")
    if work_image.width < 1600:
        scale = 1600 / max(1, work_image.width)
        work_image = work_image.resize(
            (int(work_image.width * scale), int(work_image.height * scale)),
            Image.Resampling.LANCZOS
        )

    variants = [("color", work_image)]
    if ImageOps is not None:
        gray = ImageOps.grayscale(work_image)
        gray = ImageOps.autocontrast(gray)
        if ImageEnhance is not None:
            gray = ImageEnhance.Contrast(gray).enhance(1.25)
        if ImageFilter is not None:
            gray = gray.filter(ImageFilter.SHARPEN)
        variants.append(("gray", gray))

    # Keep the OCR workload controlled: 3 primary passes.
    requested_passes = [
        ("color", 11, "eng"),
        ("gray", 11, "eng"),
        ("color", 6, "eng"),
    ]

    errors = []
    results = []

    for variant_name, psm, lang in requested_passes:
        variant_image = dict(variants).get(variant_name, work_image)
        try:
            data = pytesseract.image_to_data(
                variant_image,
                lang=lang,
                output_type=Output.DICT,
                config=f"--psm {psm}"
            )

            words = []
            grouped_text = {}
            conf_values = []

            for i, raw_text in enumerate(data.get("text", [])):
                word = str(raw_text or "").strip()
                if not word:
                    continue
                try:
                    conf = float(
                        data.get("conf", ["-1"] * len(data.get("text", [])))[i]
                    )
                except Exception:
                    conf = -1.0

                item = {
                    "text": word,
                    "left": int(data.get("left", [0])[i]),
                    "top": int(data.get("top", [0])[i]),
                    "width": int(data.get("width", [0])[i]),
                    "height": int(data.get("height", [0])[i]),
                    "conf": conf,
                    "block_num": int(data.get("block_num", [0])[i]),
                    "par_num": int(data.get("par_num", [0])[i]),
                    "line_num": int(data.get("line_num", [0])[i]),
                }
                words.append(item)
                if conf >= 0:
                    conf_values.append(conf)
                line_key = (
                    item["block_num"],
                    item["par_num"],
                    item["line_num"]
                )
                grouped_text.setdefault(line_key, []).append(item)

            text_lines = []
            for _key, line_words in sorted(grouped_text.items(), key=lambda pair: pair[0]):
                line_words.sort(key=lambda item: (item.get("top", 0), item.get("left", 0)))
                text_lines.append(" ".join(item["text"] for item in line_words))

            text = "\n".join(text_lines)
            if not _usable_text(text):
                continue

            avg_conf = sum(conf_values) / len(conf_values) if conf_values else 0

            # Prefer a clean reading order over a noisy OCR pass that happens
            # to contain more total characters. Artwork often contains logos,
            # barcode fragments and isolated symbols that inflate raw length.
            non_empty_lines = [line.strip() for line in text.splitlines() if line.strip()]
            junk_lines = sum(
                1
                for line in non_empty_lines
                if len(re.sub(r"[^A-Za-z0-9%#]", "", line)) <= 1
            )
            quality = (
                _text_quality_score(text)
                + (avg_conf * 0.20)
                - (junk_lines * 18)
            )

            results.append({
                "text": text,
                "words": words,
                "lang": lang,
                "quality": quality,
            })
        except Exception as exc:
            errors.append(f"{variant_name}/psm{psm}/{lang}: {exc}")

    if not results:
        if errors:
            raise RuntimeError(
                "OCR could not run. Tesseract may be missing. Details: "
                + " | ".join(errors[:4])
            )
        return "", [], "", ""

    best = max(results, key=lambda item: item["quality"])

    # Rebuild the strongest OCR pass from physical coordinates. Tesseract's
    # block/line numbering can occasionally reverse adjacent words on artwork;
    # geometry is a safer source of visual reading order.
    ordered_words = [
        word for word in best["words"]
        if str(word.get("text", "")).strip()
    ]

    primary_lines = []
    if ordered_words:
        heights = [max(1, int(word.get("height", 1))) for word in ordered_words]
        median_height = float(sorted(heights)[len(heights) // 2]) if heights else 20.0
        line_tolerance = max(10.0, median_height * 0.65)

        line_groups = []
        for word in sorted(
            ordered_words,
            key=lambda item: (
                float(item.get("top", 0)) + float(item.get("height", 0)) / 2.0,
                float(item.get("left", 0)),
            )
        ):
            center_y = (
                float(word.get("top", 0))
                + float(word.get("height", 0)) / 2.0
            )

            best_group = None
            best_distance = None
            for group in line_groups:
                distance = abs(center_y - group["center_y"])
                if distance <= line_tolerance and (
                    best_distance is None or distance < best_distance
                ):
                    best_group = group
                    best_distance = distance

            if best_group is None:
                line_groups.append({
                    "center_y": center_y,
                    "words": [word],
                })
            else:
                best_group["words"].append(word)
                best_group["center_y"] = sum(
                    float(w.get("top", 0)) + float(w.get("height", 0)) / 2.0
                    for w in best_group["words"]
                ) / len(best_group["words"])

        line_groups.sort(key=lambda group: group["center_y"])

        for group in line_groups:
            group["words"].sort(key=lambda item: float(item.get("left", 0)))
            line = " ".join(
                str(word.get("text", "")).strip()
                for word in group["words"]
                if str(word.get("text", "")).strip()
            )
            line = re.sub(r"\s+", " ", line).strip()
            if line:
                primary_lines.append(line)

    if not primary_lines:
        primary_lines = [
            re.sub(r"\s+", " ", line).strip()
            for line in best["text"].splitlines()
            if line.strip()
        ]

    seen = {normalize_text(line) for line in primary_lines if normalize_text(line)}

    # We do not append weaker OCR lines to the primary comparison text because
    # that can introduce incorrect duplicate/alternate readings. Instead,
    # return them separately for Auto Detect's secondary evidence.
    supplemental_lines = []
    for result in sorted(results, key=lambda item: item["quality"], reverse=True):
        if result is best:
            continue
        for line in result["text"].splitlines():
            clean = re.sub(r"\s+", " ", line).strip()
            key = normalize_text(clean)
            if key and key not in seen and key not in {
                normalize_text(existing) for existing in supplemental_lines
            }:
                supplemental_lines.append(clean)

    return (
        "\n".join(primary_lines),
        best["words"],
        best["lang"],
        "\n".join(supplemental_lines),
    )


def _ocr_image(image):
    text, _words, _lang, _supplemental = _ocr_image_with_data(image)
    return text


def _render_pdf_page(page):
    pixmap = page.get_pixmap(
        matrix=fitz.Matrix(6.0, 6.0),
        alpha=False
    )
    return Image.open(
        io.BytesIO(pixmap.tobytes("png"))
    ).convert("RGB")


def _image_to_png_bytes(image):
    output = BytesIO()
    image.save(output, format="PNG")
    return output.getvalue()


def get_output_page_count(file):
    name = str(getattr(file, "name", "")).casefold()

    if name.endswith(".pdf"):
        file.seek(0)
        data = file.read()
        document = fitz.open(stream=data, filetype="pdf")
        count = len(document)
        document.close()
        file.seek(0)
        return count

    if name.endswith((".jpg", ".jpeg", ".png")):
        return 1

    return 0


def extract_output_pages(file):
    """
    Extract artwork pages with OCR as the primary source.

    Every page also stores the rendered full-page image and OCR word boxes so
    the exact artwork can later be displayed with highlight overlays.
    """
    name = str(getattr(file, "name", "")).casefold()

    if name.endswith(".pdf"):
        file.seek(0)
        data = file.read()
        if not data:
            raise ValueError("The Output Artwork PDF is empty.")

        document = fitz.open(stream=data, filetype="pdf")
        pages = []

        try:
            for page_number, page in enumerate(document, start=1):
                direct_text = page.get_text("text") or ""
                image = _render_pdf_page(page)

                # Keep original PDF word coordinates as a visual fallback for
                # custom-font symbols. OCR may not recognize the glyph at all.
                direct_words = []
                try:
                    pdf_words = page.get_text("words") or []
                    sx = image.width / max(1.0, float(page.rect.width))
                    sy = image.height / max(1.0, float(page.rect.height))
                    for word in pdf_words:
                        if len(word) < 5:
                            continue
                        x0, y0, x1, y1, word_text = word[:5]
                        word_text = str(word_text or "")
                        if not word_text.strip():
                            continue
                        direct_words.append({
                            "text": word_text,
                            "left": int(round(float(x0) * sx)),
                            "top": int(round(float(y0) * sy)),
                            "width": int(round((float(x1) - float(x0)) * sx)),
                            "height": int(round((float(y1) - float(y0)) * sy)),
                        })
                except Exception:
                    direct_words = []
                ocr_text = ""
                ocr_words = []
                ocr_error = None
                ocr_lang = ""

                try:
                    (
                        ocr_text,
                        ocr_words,
                        ocr_lang,
                        ocr_alt_text,
                    ) = _ocr_image_with_data(image)
                except Exception as exc:
                    ocr_error = exc

                ocr_scale = 1600 / max(1, image.width) if image.width < 1600 else 1.0

                if _usable_text(ocr_text):
                    text = ocr_text
                    source_type = "ocr"
                elif _usable_text(direct_text):
                    text = direct_text
                    source_type = "pdf_text"
                else:
                    detail = str(ocr_error) if ocr_error else "no usable text"
                    raise RuntimeError(
                        f"Page {page_number}: no readable artwork text was found. {detail}"
                    )

                pages.append({
                    "page": page_number,
                    "text": str(text),
                    "source_type": source_type,
                    "direct_text": str(direct_text or ""),
                    "ocr_text": str(ocr_text or ""),
                    "ocr_alt_text": str(ocr_alt_text or ""),
                    "ocr_words": ocr_words,
                    "direct_words": direct_words,
                    "ocr_lang": ocr_lang,
                    "ocr_scale_x": ocr_scale,
                    "ocr_scale_y": ocr_scale,
                    "image_bytes": _image_to_png_bytes(image),
                    "image_width": image.width,
                    "image_height": image.height,
                })
        finally:
            document.close()

        file.seek(0)
        return pages

    if name.endswith((".jpg", ".jpeg", ".png")):
        file.seek(0)
        image = Image.open(file).convert("RGB")
        (
            ocr_text,
            ocr_words,
            ocr_lang,
            ocr_alt_text,
        ) = _ocr_image_with_data(image)
        if not _usable_text(ocr_text):
            raise RuntimeError("No readable artwork text was detected in the image.")
        ocr_scale = 1600 / max(1, image.width) if image.width < 1600 else 1.0
        file.seek(0)
        return [{
            "page": 1,
            "text": str(ocr_text),
            "source_type": "ocr",
            "direct_text": "",
            "ocr_text": str(ocr_text),
            "ocr_alt_text": str(ocr_alt_text),
            "ocr_words": ocr_words,
            "direct_words": [],
            "ocr_lang": ocr_lang,
            "ocr_scale_x": ocr_scale,
            "ocr_scale_y": ocr_scale,
            "image_bytes": _image_to_png_bytes(image),
            "image_width": image.width,
            "image_height": image.height,
        }]

    raise ValueError(
        "Unsupported output format. Please upload PDF, JPG, JPEG, or PNG."
    )


# =========================================================
# VISUAL ARTWORK HIGHLIGHTING
# =========================================================

# Presentation-only palette.  The comparison engine never uses these colors.
# The same field receives the same color in the artwork, table swatch, and
# Excel report.
FIELD_VISUAL_COLORS = [
    "#2563EB",  # blue
    "#0F766E",  # teal
    "#7C3AED",  # violet
    "#DB2777",  # pink
    "#EA580C",  # orange
    "#CA8A04",  # gold
    "#0891B2",  # cyan
    "#C026D3",  # fuchsia
    "#65A30D",  # lime
    "#DC2626",  # red
    "#4F46E5",  # indigo
    "#B45309",  # amber
    "#15803D",  # dark green
    "#0284C7",  # sky blue
    "#BE123C",  # rose
    "#4338CA",  # dark indigo
    "#16A34A",  # green
    "#9D174D",  # deep pink
]


def get_field_visual_colors(fields):
    """Return a stable FIELD -> HEX map based only on displayed field order."""
    colors = {}
    for index, field in enumerate([str(x) for x in (fields or [])]):
        colors[field] = FIELD_VISUAL_COLORS[index % len(FIELD_VISUAL_COLORS)]
    return colors


def add_visual_column(report, selected_fields):
    """Add a presentation-only VISUAL swatch column without changing the report."""
    if report is None:
        return report

    display = report.copy()
    colors = get_field_visual_colors(selected_fields)
    # Streamlit dataframe receives a styling layer for the swatch column.
    # Keep the underlying value empty so Excel can use a true solid-color cell
    # instead of a text glyph that can make every swatch look similar.
    swatches = ["" for _field in display.get("FIELD", []).tolist()]

    if "VISUAL" in display.columns:
        display["VISUAL"] = swatches
    else:
        insert_at = display.columns.get_loc("STATUS") if "STATUS" in display.columns else len(display.columns)
        display.insert(insert_at, "VISUAL", swatches)
    return display


def _visual_norm(text):
    if text is None:
        return ""
    value = unicodedata.normalize("NFKC", str(text)).casefold()
    value = value.replace("%", "")
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def _visual_compact(text):
    return re.sub(r"[^a-z0-9]", "", _visual_norm(text))


def _page_ocr_scale(page):
    """Return OCR-coordinate -> stored-image coordinate scale factors."""
    try:
        sx = float(page.get("ocr_scale_x", 1.0) or 1.0)
        sy = float(page.get("ocr_scale_y", 1.0) or 1.0)
    except Exception:
        sx = sy = 1.0
    if sx <= 0:
        sx = 1.0
    if sy <= 0:
        sy = 1.0
    return sx, sy


def _scaled_word(word, sx, sy):
    """Convert OCR coordinates back to the original stored artwork image."""
    return {
        **word,
        "left": int(round(float(word.get("left", 0)) / sx)),
        "top": int(round(float(word.get("top", 0)) / sy)),
        "width": max(1, int(round(float(word.get("width", 1)) / sx))),
        "height": max(1, int(round(float(word.get("height", 1)) / sy))),
    }


def _visual_group_words(page):
    """Return OCR words grouped by physical line in stored-image coordinates."""
    sx, sy = _page_ocr_scale(page)
    grouped = {}

    for raw_word in page.get("ocr_words", []) or []:
        if not isinstance(raw_word, dict):
            continue
        if not str(raw_word.get("text", "")).strip():
            continue
        word = _scaled_word(raw_word, sx, sy)
        key = (
            raw_word.get("block_num", 0),
            raw_word.get("par_num", 0),
            raw_word.get("line_num", 0),
        )
        grouped.setdefault(key, []).append(word)

    groups = list(grouped.values())
    for group in groups:
        group.sort(key=lambda item: (float(item.get("left", 0)), float(item.get("top", 0))))

    groups.sort(
        key=lambda group: (
            min(float(word.get("top", 0)) for word in group),
            min(float(word.get("left", 0)) for word in group),
        )
    )
    return groups


def _boxes_from_words(words):
    if not words:
        return []
    left = min(int(word.get("left", 0)) for word in words)
    top = min(int(word.get("top", 0)) for word in words)
    right = max(int(word.get("left", 0)) + int(word.get("width", 0)) for word in words)
    bottom = max(int(word.get("top", 0)) + int(word.get("height", 0)) for word in words)
    return [(left, top, right, bottom)] if right > left and bottom > top else []


def _find_visual_symbol_boxes(page, target):
    """Locate custom-font symbol keystrokes using OCR/direct PDF word boxes.

    Presentation-only. The comparison decision has already been made by the
    SYMBOL checker. This function only finds the visual location of that exact
    keystroke.
    """
    expected = normalize_symbol_text(target)
    if not expected:
        return []

    def raw_matches(expected_value, actual_value):
        return _symbol_text_matches(expected_value, actual_value)

    # Prefer the original PDF text layer because it preserves custom-font
    # keystrokes even when OCR does not recognize them.
    for word in page.get("direct_words", []) or []:
        if raw_matches(expected, word.get("text", "")):
            return _boxes_from_words([word])

    # OCR fallback when the symbol survived OCR as a usable token.
    for word in page.get("ocr_words", []) or []:
        if raw_matches(expected, word.get("text", "")):
            return _boxes_from_words([word])

    return []


def _find_visual_boxes(page, target, field_name=""):
    """
    Presentation-only lookup of the actual OCR words corresponding to the
    comparison result's PDF OUTPUT value.

    IMPORTANT: OCR boxes are converted back from the enlarged OCR image to the
    original stored artwork image before drawing.  This is the key protection
    against the offset caused by the OCR preprocessing resize.
    """
    target_norm = _visual_norm(target)
    target_compact = _visual_compact(target)
    if not target_norm or target_norm in {"not found", "-", "—"}:
        return []

    groups = _visual_group_words(page)
    if not groups:
        return []

    # 1) Exact contiguous OCR-word sequence.
    for group in groups:
        normalized = [_visual_norm(word.get("text", "")) for word in group]
        for start in range(len(group)):
            accumulated = []
            selected = []
            for end in range(start, len(group)):
                token = normalized[end]
                if not token:
                    continue
                accumulated.append(token)
                selected.append(group[end])
                joined = " ".join(accumulated).strip()

                if joined == target_norm or _visual_compact(joined) == target_compact:
                    return _boxes_from_words(selected)

                # Prevent a search from walking through unrelated later words.
                if len(_visual_compact(joined)) > len(target_compact) + 8:
                    break

    # 2) Numeric component inside a combined token such as 44-12.
    if re.fullmatch(r"\d+", target_norm):
        for group in groups:
            for word in group:
                raw = str(word.get("text", "")).strip()
                match = re.search(r"(\d+)[-/](\d+)", raw)
                if not match:
                    continue

                first, second = match.group(1), match.group(2)
                if target_norm not in {first, second}:
                    continue

                left = int(word.get("left", 0))
                top = int(word.get("top", 0))
                width = int(word.get("width", 0))
                height = int(word.get("height", 0))
                full = f"{first}-{second}"
                total_chars = max(1, len(full))

                if target_norm == first:
                    end_x = left + max(1, int(round(width * len(first) / total_chars)))
                    return [(left, top, min(left + width, end_x), top + height)]

                start_x = left + max(1, int(round(width * (len(first) + 1) / total_chars)))
                return [(min(left + width - 1, start_x), top, left + width, top + height)]

    # 3) Controlled cross-line exact compact sequence.
    ordered = [word for group in groups for word in group]
    for start in range(len(ordered)):
        compact = ""
        selected = []
        for end in range(start, min(len(ordered), start + 16)):
            token = _visual_compact(ordered[end].get("text", ""))
            if not token:
                continue
            compact += token
            selected.append(ordered[end])
            if compact == target_compact:
                return _boxes_from_words(selected)
            if len(compact) > len(target_compact) + 8:
                break

    return []


def _hex_rgb(hex_color):
    value = str(hex_color).lstrip("#")
    if len(value) != 6:
        return (37, 99, 235)
    try:
        return tuple(int(value[index:index + 2], 16) for index in (0, 2, 4))
    except Exception:
        return (37, 99, 235)


def _load_visual_fonts():
    try:
        from PIL import ImageFont
        candidates = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/dejavu/DejaVuSans.ttf",
        ]
        bold_candidates = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf",
        ]
        regular_path = next((p for p in candidates if Path(p).exists()), None)
        bold_path = next((p for p in bold_candidates if Path(p).exists()), None)
        if regular_path and bold_path:
            return (
                ImageFont.truetype(regular_path, 22),
                ImageFont.truetype(bold_path, 22),
            )
    except Exception:
        pass
    return None, None


def _text_size(draw, text, font):
    if font is None:
        return max(20, len(str(text)) * 11), 20
    try:
        bbox = draw.textbbox((0, 0), text, font=font)
        return bbox[2] - bbox[0], bbox[3] - bbox[1]
    except Exception:
        return max(20, len(str(text)) * 11), 20


def _rect_intersects(a, b, pad=8):
    return not (
        a[2] + pad < b[0]
        or a[0] - pad > b[2]
        or a[3] + pad < b[1]
        or a[1] - pad > b[3]
    )


def _place_label_above_or_below(box, label_w, label_h, image_w, image_h, occupied):
    left, top, right, bottom = box
    gap = 10
    x = max(6, min(image_w - label_w - 6, int((left + right - label_w) / 2)))

    candidates = [
        (x, top - label_h - gap),
        (x, bottom + gap),
        (max(6, left - label_w - gap), int((top + bottom - label_h) / 2)),
        (min(image_w - label_w - 6, right + gap), int((top + bottom - label_h) / 2)),
    ]

    for candidate in candidates:
        cx, cy = candidate
        if cx < 6 or cy < 6 or cx + label_w > image_w - 6 or cy + label_h > image_h - 6:
            continue
        rect = (cx, cy, cx + label_w, cy + label_h)
        if not any(_rect_intersects(rect, existing, pad=6) for existing in occupied):
            return candidate

    return (
        max(6, min(image_w - label_w - 6, x)),
        max(6, min(image_h - label_h - 6, top - label_h - gap)),
    )


def _visual_group_text(group):
    return re.sub(
        r"\s+",
        " ",
        " ".join(
            str(word.get("text", "")).strip()
            for word in group
            if str(word.get("text", "")).strip()
        )
    ).strip()


def _visual_word_center(word):
    left = float(word.get("left", 0) or 0)
    top = float(word.get("top", 0) or 0)
    width = float(word.get("width", 0) or 0)
    height = float(word.get("height", 0) or 0)
    return left + width / 2.0, top + height / 2.0


def _visual_direct_groups(page):
    """Build line groups from the PDF text layer using rendered-image coordinates."""
    words = []
    for word in page.get("direct_words", []) or []:
        if not isinstance(word, dict):
            continue
        if not str(word.get("text", "")).strip():
            continue
        words.append({
            **word,
            "left": int(word.get("left", 0) or 0),
            "top": int(word.get("top", 0) or 0),
            "width": max(1, int(word.get("width", 1) or 1)),
            "height": max(1, int(word.get("height", 1) or 1)),
        })

    if not words:
        return []

    heights = sorted(max(1, int(w.get("height", 1))) for w in words)
    median_height = float(heights[len(heights) // 2]) if heights else 16.0
    tolerance = max(3.0, median_height * 0.55)

    groups = []
    for word in sorted(words, key=lambda item: (_visual_word_center(item)[1], item["left"])):
        _, cy = _visual_word_center(word)
        best_group = None
        best_distance = None
        for group in groups:
            distance = abs(cy - group["center_y"])
            if distance <= tolerance and (best_distance is None or distance < best_distance):
                best_group = group
                best_distance = distance
        if best_group is None:
            groups.append({"center_y": cy, "words": [word]})
        else:
            best_group["words"].append(word)
            best_group["center_y"] = sum(_visual_word_center(w)[1] for w in best_group["words"]) / len(best_group["words"])

    result = [group["words"] for group in groups]
    for group in result:
        group.sort(key=lambda item: item["left"])
    result.sort(key=lambda group: (min(w["top"] for w in group), min(w["left"] for w in group)))
    return result


def _visual_all_groups(page):
    """Prefer OCR geometry, then supplement missing lines from the PDF text layer."""
    ocr_groups = _visual_group_words(page)
    direct_groups = _visual_direct_groups(page)

    if not ocr_groups:
        return direct_groups
    if not direct_groups:
        return ocr_groups

    # OCR is normally better for visual reading order, but a PDF text layer can
    # contain custom-font or tiny-text words that OCR completely misses. Keep OCR
    # as the primary geometry and add only clearly distinct PDF lines.
    merged = list(ocr_groups)
    existing_lines = [_visual_norm(_visual_group_text(group)) for group in merged]
    for direct_group in direct_groups:
        direct_text = _visual_norm(_visual_group_text(direct_group))
        if not direct_text:
            continue
        if any(direct_text == line or _visual_compact(direct_text) == _visual_compact(line) for line in existing_lines if line):
            continue
        # Only add a direct group if its center is not almost identical to a very
        # different OCR group. This prevents duplicate text boxes on the same line.
        _, dcy = _visual_word_center(direct_group[0])
        too_close = False
        for ocr_group in merged:
            _, ocy = _visual_word_center(ocr_group[0])
            if abs(dcy - ocy) <= max(5.0, float(max(1, direct_group[0].get("height", 1))) * 0.7):
                too_close = True
                break
        if not too_close:
            merged.append(direct_group)
            existing_lines.append(direct_text)

    merged.sort(key=lambda group: (min(float(w.get("top", 0)) for w in group), min(float(w.get("left", 0)) for w in group)))
    return merged


def _visual_match_line_score(target_line, actual_line):
    """Similarity score used only to align comparison text to OCR/PDF line geometry."""
    target_norm = _visual_norm(target_line)
    actual_norm = _visual_norm(actual_line)
    if not target_norm or not actual_norm:
        return 0.0
    if target_norm == actual_norm or _visual_compact(target_norm) == _visual_compact(actual_norm):
        return 1.0
    from difflib import SequenceMatcher
    token_ratio = SequenceMatcher(
        None,
        target_norm.split(),
        actual_norm.split(),
        autojunk=False,
    ).ratio()
    compact_ratio = SequenceMatcher(
        None,
        _visual_compact(target_norm),
        _visual_compact(actual_norm),
        autojunk=False,
    ).ratio()
    return max(token_ratio, compact_ratio * 0.98)


def _visual_find_text_occurrences(page, target, groups=None, min_score=0.80):
    """Find one or more physically grounded occurrences of a text/value."""
    target = str(target or "").strip()
    if not target or normalize_text(target) in {"not found", "-", "—"}:
        return []

    groups = groups or _visual_all_groups(page)
    if not groups:
        return []

    target_lines = [re.sub(r"\s+", " ", str(line).strip()) for line in target.splitlines() if str(line).strip()]
    target_norm = _visual_norm(target)
    target_compact = _visual_compact(target)

    # Exact physical line match.
    exact_groups = []
    used_start = 0
    for target_line in target_lines:
        best_idx = None
        for idx in range(used_start, len(groups)):
            line_text = _visual_group_text(groups[idx])
            if not line_text:
                continue
            if _visual_norm(target_line) == _visual_norm(line_text) or _visual_compact(target_line) == _visual_compact(line_text):
                best_idx = idx
                break
        if best_idx is None:
            exact_groups = []
            break
        exact_groups.append(best_idx)
        used_start = best_idx + 1

    if exact_groups and len(exact_groups) == len(target_lines):
        return [_boxes_from_words(groups[idx])[0] for idx in exact_groups if _boxes_from_words(groups[idx])]

    # Exact contiguous words within each physical line.
    for group in groups:
        normalized = [_visual_norm(word.get("text", "")) for word in group]
        compact_tokens = [_visual_compact(word.get("text", "")) for word in group]
        for start in range(len(group)):
            joined_parts = []
            chosen = []
            compact_joined = ""
            for end in range(start, len(group)):
                token = normalized[end]
                compact_token = compact_tokens[end]
                if not token and not compact_token:
                    continue
                joined_parts.append(token)
                compact_joined += compact_token
                chosen.append(group[end])
                joined = " ".join(x for x in joined_parts if x).strip()
                if joined == target_norm or compact_joined == target_compact:
                    box = _boxes_from_words(chosen)
                    if box:
                        return box
                if len(compact_joined) > len(target_compact) + 10:
                    break

    # Fuzzy monotonic line alignment for OCR line wrapping.
    if len(target_lines) > 1:
        chosen = []
        search_from = 0
        for target_line in target_lines:
            best_idx = None
            best_score = 0.0
            for idx in range(search_from, min(len(groups), search_from + 18)):
                actual_line = _visual_group_text(groups[idx])
                score = _visual_match_line_score(target_line, actual_line)
                if score > best_score:
                    best_score = score
                    best_idx = idx
            if best_idx is None or best_score < min_score:
                continue
            chosen.append(best_idx)
            search_from = best_idx + 1
        if chosen and len(chosen) >= max(1, int(len(target_lines) * 0.60)):
            result = []
            for idx in chosen:
                result.extend(_boxes_from_words(groups[idx]))
            if result:
                return result

    # Token-level fallback. Prefer longer tokens to avoid selecting a lone 'S',
    # 'M', 'L', '18', etc. unless that is genuinely all the evidence available.
    tokens = sorted(
        [token for token in tokenize(target) if len(token) >= 2 or token.isdigit()],
        key=lambda item: (-len(item), item),
    )
    for token in tokens:
        token_compact = _visual_compact(token)
        if not token_compact:
            continue
        matches = []
        for group in groups:
            for word in group:
                word_norm = _visual_norm(word.get("text", ""))
                word_compact = _visual_compact(word.get("text", ""))
                if word_norm == _visual_norm(token) or word_compact == token_compact:
                    matches.append(word)
        if matches:
            return _boxes_from_words([matches[0]])

    return []


def _visual_field_markers(field_name):
    """Strong region anchors for field-specific visual evidence."""
    compact = re.sub(r"[^a-z0-9]", "", str(field_name).casefold())
    field_type = get_field_type(field_name)
    region = get_field_region(field_name)

    if field_type == "CONTENT":
        if "fibspmexico" in compact:
            return ["cr/ec/gt/pa/sv : cuerpo", "cr/ec/gt/pa/sv:", "cr/ec/gt/pa/sv", "cuerpo"], [
                "mx :", "mx cuerpo", "machine wash", "laver", "lavar", "made in", "rn"
            ]
        if "fibmexico" in compact:
            return ["mx : cuerpo", "mx cuerpo", "mx :", "mx:"], [
                "cr/ec/gt/pa/sv", "machine wash", "laver", "lavar", "made in", "rn"
            ]
        if "fibca" in compact:
            return ["ca : extérieur", "ca : exterieur", "ca extérieur", "ca exterieur", "extérieur", "exterieur"], [
                "mx :", "cr/ec/gt/pa/sv", "machine wash", "laver", "lavar", "made in", "rn"
            ]
        if "fibsp" in compact:
            return ["cr/ec/gt/pa/sv : cuerpo", "cr/ec/gt/pa/sv:", "cr/ec/gt/pa/sv", "cuerpo"], [
                "mx :", "machine wash", "laver", "lavar", "made in", "rn"
            ]
        if "fiben" in compact:
            return ["us : shell", "us shell", "shell:", "shell"], [
                "ca :", "mx :", "cr/ec/gt/pa/sv", "machine wash", "laver", "lavar", "made in", "rn"
            ]
        return ["shell", "content"], ["machine wash", "laver", "lavar", "made in", "rn"]

    if field_type == "CARE":
        if region == "FR":
            return ["laver à la machine", "laver a la machine", "laver"], ["machine wash", "lavar", "made in", "rn"]
        if region == "SP":
            return ["lavar a máquina", "lavar a maquina", "lavar"], ["machine wash", "laver", "made in", "rn"]
        return ["machine wash", "wash"], ["laver", "lavar", "made in", "rn"]

    if field_type == "COO":
        if region == "FR":
            return ["fabrique en"], ["made in", "hecho en", "rn", "ca"]
        if region == "SP":
            return ["hecho en"], ["made in", "fabrique en", "rn", "ca"]
        return ["made in"], ["fabrique en", "hecho en", "rn", "ca"]

    if field_type in {"RN", "CA"}:
        return ["rn", "ca"], ["machine wash", "laver", "lavar", "made in"]

    return [], []


def _visual_find_direct_scalar_box(page, target):
    """Search the original PDF text layer directly for a short scalar value."""
    target = str(target or "").strip()
    target_norm = _visual_norm(target)
    target_compact = _visual_compact(target)
    if not target_norm:
        return []

    words = []
    for word in page.get("direct_words", []) or []:
        if not str(word.get("text", "")).strip():
            continue
        words.append(word)

    # Whole-word/whole-token match first.
    for word in words:
        word_norm = _visual_norm(word.get("text", ""))
        word_compact = _visual_compact(word.get("text", ""))
        if word_norm == target_norm or word_compact == target_compact:
            return _boxes_from_words([word])

    # A registered-symbol or punctuation suffix may make the PDF word longer
    # than the compared identifier. Compact substring matching is safe for
    # identifiers of 3+ characters.
    if len(target_compact) >= 3:
        for word in words:
            word_compact = _visual_compact(word.get("text", ""))
            if target_compact in word_compact:
                return _boxes_from_words([word])

    return []


def _visual_region_groups(page, field_name):
    """Return the physically relevant OCR/PDF line groups for a semantic field."""
    groups = _visual_all_groups(page)
    if not groups:
        return []

    starts, stops = _visual_field_markers(field_name)
    starts = [_visual_norm(x) for x in starts if _visual_norm(x)]
    stops = [_visual_norm(x) for x in stops if _visual_norm(x)]
    if not starts:
        return []

    start_idx = None
    for idx, group in enumerate(groups):
        line_text = _visual_norm(_visual_group_text(group))
        if line_text and any(marker in line_text for marker in starts):
            start_idx = idx
            break
    if start_idx is None:
        return []

    field_type = get_field_type(field_name)
    selected = []
    max_lines = 10 if field_type == "CONTENT" else 22 if field_type == "CARE" else 3

    for idx in range(start_idx, min(len(groups), start_idx + max_lines)):
        group = groups[idx]
        line_text = _visual_norm(_visual_group_text(group))
        if not line_text:
            continue

        if idx > start_idx:
            # A region marker on the same line is handled below; a marker on a
            # new line ends the current semantic region before that line.
            if any(stop in line_text for stop in stops):
                break

        group_to_use = list(group)

        # When multiple regional values share one physical OCR line, clip the
        # group before the next region marker instead of highlighting the entire
        # combined line. This is critical for CA/MX/CR-EC-GT-PA-SV content.
        if idx >= start_idx and stops:
            for wi, word in enumerate(group_to_use):
                word_norm = _visual_norm(word.get("text", ""))
                if not word_norm:
                    continue
                if any(stop == word_norm or stop in word_norm for stop in stops):
                    if idx == start_idx and wi == 0:
                        group_to_use = []
                    else:
                        group_to_use = group_to_use[:wi]
                    break

        if not group_to_use:
            break

        if field_type == "CONTENT" and idx > start_idx:
            raw_line = _visual_group_text(group_to_use)
            has_content = bool(re.search(r"\d{1,3}\s*%", raw_line)) or any(
                material in _visual_norm(raw_line)
                for material in (
                    "polyester", "spandex", "elastane", "cotton", "nylon", "rayon",
                    "viscose", "acrylic", "linen", "wool", "elastodiene", "polyamide"
                )
            )
            if not has_content:
                if idx == start_idx + 1 and selected:
                    selected.append(group_to_use)
                    continue
                break

        selected.append(group_to_use)

        if field_type == "CONTENT" and len(selected) >= 8:
            break
        if field_type not in {"CONTENT", "CARE"} and len(selected) >= 3:
            break

    return selected


def _visual_region_boxes(page, field_name, actual_value=""):
    """Find a semantic region when exact text occurrence is insufficient."""
    groups = _visual_region_groups(page, field_name)
    if not groups:
        return []
    boxes = []
    for group in groups:
        boxes.extend(_boxes_from_words(group))
    return boxes



def _visual_size_like_text(text):
    norm = _visual_norm(text)
    if not norm:
        return False
    patterns = [
        r"\b(?:xxxs|xxs|xs|s|m|l|xl|xxl|xxxl|2xl|3xl|4xl|5xl)\b",
        r"\b(?:tp|p|g|gg|tg|ttg|ech|ch|eg|ee|eeg|petite)\b",
        r"\b\d{1,3}\s*[-/]\s*\d{1,3}\b",
        r"\b\d{1,3}\s*\(\s*\d{1,3}(?:\s*[-/]\s*\d{1,3})?\s*\)",
    ]
    return any(re.search(pattern, norm, re.IGNORECASE) for pattern in patterns)


def _visual_raw_size_like_text(text):
    raw = str(text or "").casefold()
    normalized = _visual_norm(raw)
    patterns = [
        r"\b(?:xxxs|xxs|xs|s|m|l|xl|xxl|xxxl|2xl|3xl|4xl|5xl)\b",
        r"\b(?:tp|p|g|gg|tg|ttg|ech|ch|eg|ee|eeg|petite)\b",
        r"\b\d{1,3}\s*[/\-]\s*\d{1,3}\b",
        r"\b\d{1,3}\s*\(\s*\d{1,3}(?:\s*[-/]\s*\d{1,3})?\s*\)",
        r"\b\d{1,3}\s*[-]\s*\d{1,3}\b",
    ]
    return any(re.search(pattern, raw, re.IGNORECASE) for pattern in patterns) or _visual_size_like_text(normalized)


def _visual_size_sequence_blocks(page):
    """Group physical size rows into actual size blocks rather than individual lines."""
    groups = _visual_all_groups(page)
    if not groups:
        return []

    candidates = []
    for idx, group in enumerate(groups):
        text = _visual_group_text(group)
        if not _visual_raw_size_like_text(text):
            continue
        box_list = _boxes_from_words(group)
        if not box_list:
            continue
        left, top, right, bottom = box_list[0]
        candidates.append({
            "group_index": idx,
            "group": group,
            "text": text,
            "box": (left, top, right, bottom),
            "cx": (left + right) / 2.0,
            "cy": (top + bottom) / 2.0,
        })

    if len(candidates) < 2:
        return []

    blocks = []
    # Size blocks in the current artworks are primarily vertical columns. Group
    # rows with similar X and ordinary line spacing; a larger vertical gap starts
    # the next logical size block.
    for item in sorted(candidates, key=lambda x: (x["cx"], x["cy"])):
        placed = False
        for block in blocks:
            avg_x = sum(x["cx"] for x in block) / len(block)
            last = block[-1]
            gap = item["cy"] - last["cy"]
            typical_heights = [max(1, b["box"][3] - b["box"][1]) for b in block]
            typical_h = sorted(typical_heights)[len(typical_heights) // 2]
            x_tolerance = max(24.0, typical_h * 2.5)
            max_gap = max(46.0, typical_h * 1.95)
            if abs(item["cx"] - avg_x) <= x_tolerance and -5.0 <= gap <= max_gap:
                block.append(item)
                block.sort(key=lambda x: x["cy"])
                placed = True
                break
        if not placed:
            blocks.append([item])

    # Some columns can be split into multiple blocks because OCR drops the
    # regional numeric line. Merge neighboring blocks when the gap is small and
    # they share the same horizontal column.
    merged = []
    for block in sorted(blocks, key=lambda b: (sum(x["cx"] for x in b) / len(b), min(x["cy"] for x in b))):
        if not merged:
            merged.append(block)
            continue
        prev = merged[-1]
        prev_x = sum(x["cx"] for x in prev) / len(prev)
        curr_x = sum(x["cx"] for x in block) / len(block)
        prev_bottom = max(x["box"][3] for x in prev)
        curr_top = min(x["box"][1] for x in block)
        gap = curr_top - prev_bottom
        if abs(prev_x - curr_x) <= 34 and gap <= 48:
            prev.extend(block)
            prev.sort(key=lambda x: x["cy"])
        else:
            merged.append(block)

    result = []
    for block in merged:
        if not block:
            continue
        # Only keep plausible size blocks; long unrelated text strings should
        # not become a sequence merely because they contain an 'S' or a number.
        alpha_size_count = sum(
            bool(re.search(r"\b(?:xxs|xs|s|m|l|xl|xxl|xxxl|tp|tg|ttg|ech|ch|eg|eeg|p|g)\b", _visual_norm(item["text"])))
            for item in block
        )
        if alpha_size_count >= 1 or len(block) >= 2:
            result.append(block)

    # Sort blocks from top-to-bottom for the dominant size column. If several
    # columns exist, prioritize the longest/plausible size column.
    result.sort(key=lambda block: (min(item["cy"] for item in block), min(item["cx"] for item in block)))
    return result


def _find_visual_osz_box(page, field_name, actual_value=""):
    """Locate OSZ/OS_Size_N by size-block position, not by arbitrary number position."""
    compact = re.sub(r"[^a-z0-9]", "", str(field_name).casefold())
    match = re.search(r"(?:osz|ossize)(\d+)$", compact)
    if not match:
        return []
    index = int(match.group(1))
    if index <= 0:
        return []

    blocks = _visual_size_sequence_blocks(page)
    if not blocks or index > len(blocks):
        return []

    actual_norm = _visual_norm(actual_value)
    actual_compact = _visual_compact(actual_value)

    # Prefer a block whose combined text substantially overlaps the compared
    # output. For a mismatch, this still chooses the correct physical size block.
    def block_score(block):
        text = " ".join(item["text"] for item in block)
        norm = _visual_norm(text)
        score = 0.0
        if actual_norm:
            if actual_norm == norm:
                score += 50
            elif actual_compact and actual_compact in _visual_compact(norm):
                score += 35
            # Token overlap is useful when punctuation/line wrapping differs.
            at = set(tokenize(actual_value))
            bt = set(tokenize(text))
            if at and bt:
                score += 20.0 * len(at & bt) / max(1, len(at))
        score += min(12.0, len(block) * 2.0)
        return score

    # The report's OS_Size_N numbering is the semantic sequence order. First try
    # the exact ordinal block because that is the strongest source of identity.
    chosen_block = blocks[index - 1]

    # If the ordinal block looks implausible but another block strongly matches
    # the actual output, use the stronger semantic match.
    ordinal_score = block_score(chosen_block)
    best_idx = max(
        range(len(blocks)),
        key=lambda idx: block_score(blocks[idx])
    )
    best_score = block_score(blocks[best_idx])
    if best_score >= ordinal_score + 25:
        chosen_block = blocks[best_idx]

    boxes = []
    for item in chosen_block:
        boxes.extend(_boxes_from_words(item["group"]))
    return boxes



def _visual_text_similarity(a, b):
    """Return a robust token/compact similarity for visual evidence matching."""
    a = str(a or "").strip()
    b = str(b or "").strip()
    if not a or not b:
        return 0.0
    if _visual_norm(a) == _visual_norm(b) or _visual_compact(a) == _visual_compact(b):
        return 1.0
    from difflib import SequenceMatcher
    token_ratio = SequenceMatcher(
        None,
        _visual_norm(a).split(),
        _visual_norm(b).split(),
        autojunk=False,
    ).ratio()
    compact_ratio = SequenceMatcher(
        None,
        _visual_compact(a),
        _visual_compact(b),
        autojunk=False,
    ).ratio()
    return max(token_ratio, compact_ratio * 0.98)


def _visual_find_multiline_block_boxes(page, target, groups=None, min_score=0.48, max_lines=10):
    """
    Locate a multi-line field as a physical block.

    This deliberately works from the actual text returned by the comparison
    engine (usually PDF text or OCR output) rather than trying to rediscover a
    generic phrase.  It is tolerant of OCR punctuation/line-break differences
    and can therefore map FAIL values such as an OS_Size block where only one
    number differs.
    """
    target = str(target or "").strip()
    if not target:
        return []
    groups = groups or _visual_all_groups(page)
    if not groups:
        return []

    # Preserve explicit source newlines when they exist.  When they do not,
    # derive soft target lines from common size-row patterns.
    raw_lines = [re.sub(r"\s+", " ", line).strip() for line in target.splitlines() if str(line).strip()]
    if len(raw_lines) <= 1:
        compact_target = _visual_compact(target)
        # Common artwork size rows.  This lets a collapsed comparison result
        # still resolve to four physical size lines.
        extracted = re.findall(
            r"(?:\b(?:XXXS|XXS|XS|S|M|L|XL|XXL|XXXL|TP|P|G|GG|TG|TTG|ECH|CH|EG|EE|EEG)\b\s*\([^)]*\)|"
            r"\b\d{1,3}\s*/\s*\d{1,3}\s*[-–]\s*\d{1,3}(?:\s*[-–]\s*\d{1,3})?\b)",
            target,
            flags=re.IGNORECASE,
        )
        if len(extracted) >= 2:
            raw_lines = extracted
        else:
            raw_lines = [target]

    # Build candidate windows.  Size and long regional blocks rarely exceed ten
    # physical OCR lines, so this remains lightweight even on large pages.
    target_count = len(raw_lines)
    candidate_sizes = range(
        max(1, target_count - 2),
        min(max_lines, target_count + 2) + 1,
    )

    best = None
    group_texts = [_visual_group_text(g) for g in groups]

    for window_size in candidate_sizes:
        if window_size > len(groups):
            continue
        for start in range(0, len(groups) - window_size + 1):
            window = groups[start:start + window_size]
            texts = group_texts[start:start + window_size]

            # For explicit multi-line targets, align target line -> one physical
            # group. For collapsed targets, allow a token coverage score.
            if len(raw_lines) > 1:
                scores = []
                search_pos = 0
                for target_line in raw_lines:
                    local_best = 0.0
                    local_idx = None
                    for idx in range(search_pos, len(texts)):
                        score = _visual_text_similarity(target_line, texts[idx])
                        if score > local_best:
                            local_best = score
                            local_idx = idx
                    if local_idx is None:
                        scores.append(0.0)
                        continue
                    scores.append(local_best)
                    search_pos = local_idx + 1
                if not scores:
                    continue
                line_coverage = sum(1 for score in scores if score >= 0.42) / max(1, len(scores))
                avg_score = sum(scores) / len(scores)
                # Strongly reward candidates where the expected number of rows
                # is actually present. This prevents a random nearby paragraph
                # from winning on one long shared word.
                score = avg_score * 0.72 + line_coverage * 0.28
            else:
                joined = " ".join(texts)
                score = _visual_text_similarity(target, joined)

            # Size-shaped windows receive a small bonus when most groups look
            # like size rows. This is especially useful when the page has many
            # unrelated numeric lines nearby.
            size_like = sum(1 for text in texts if _visual_raw_size_like_text(text))
            if size_like >= max(2, min(4, len(raw_lines))):
                score += 0.05

            # Prefer tighter windows over oversized paragraphs.
            score -= max(0, window_size - max(1, target_count)) * 0.015

            if best is None or score > best[0]:
                best = (score, start, window_size)

    if best is None or best[0] < min_score:
        return []

    _, start, window_size = best
    boxes = []
    for group in groups[start:start + window_size]:
        boxes.extend(_boxes_from_words(group))
    return boxes


def _visual_field_markers(field_name):
    """Field-specific visual anchors with explicit regional boundaries."""
    compact = re.sub(r"[^a-z0-9]", "", str(field_name).casefold())
    field_type = get_field_type(field_name)
    region = get_field_region(field_name)

    if field_type == "CONTENT":
        # These are intentionally specific. Generic anchors such as "cuerpo" or
        # "lavar" caused several Spanish fields to point at the same first match.
        if "fibspmexico" in compact:
            return ["mx :", "mx:", "cr/ec/gt/pa/sv :", "cr/ec/gt/pa/sv:", "cuerpo"], [
                "machine wash", "laver", "lavar a", "made in", "hecho en", "fabrique en", "rn ", "ca :"
            ]
        if "fibmexico" in compact:
            return ["mx :", "mx:", "mx cuerpo"], [
                "cr/ec/gt/pa/sv :", "cr/ec/gt/pa/sv:", "machine wash", "laver", "lavar a", "made in", "hecho en", "rn ", "ca :"
            ]
        if "fibsp" in compact:
            return ["cr/ec/gt/pa/sv :", "cr/ec/gt/pa/sv:", "cr/ec/gt/pa/sv"], [
                "machine wash", "laver", "lavar a", "made in", "hecho en", "fabrique en", "rn ", "ca :"
            ]
        if "fibca" in compact:
            return ["ca : exterieur", "ca : extérieur", "ca exterieur", "ca extérieur", "exterieur", "extérieur"], [
                "mx :", "cr/ec/gt/pa/sv :", "cr/ec/gt/pa/sv:", "machine wash", "laver", "lavar a", "made in", "rn "
            ]
        if "fiben" in compact:
            return ["us : shell", "us shell", "shell:"], [
                "ca :", "mx :", "cr/ec/gt/pa/sv :", "machine wash", "laver", "lavar a", "made in", "rn "
            ]
        return ["shell", "exterieur", "cuerpo", "content"], ["machine wash", "laver", "lavar a", "made in", "rn "]

    if field_type == "CARE":
        if region == "FR":
            return ["ca : laver", "ca: laver", "laver à la machine", "laver a la machine"], [
                "mx/pa/sv :", "mx/pa/sv:", "cr/ec/gt/pa/sv :", "cr/ec/gt/pa/sv", "made in", "hecho en", "rn "
            ]
        if region == "SP":
            return ["mx/pa/sv : lavar", "mx/pa/sv: lavar", "cr/ec/gt/mx/pa/sv : lavar", "lavar a maquina", "lavar a máquina"], [
                "actual other sizes", "made in", "hecho en", "rn ", "ca :"
            ]
        return ["machine wash", "machine wash cold", "wash"], [
            "ca : laver", "ca: laver", "laver", "lavar", "made in", "hecho en", "rn "
        ]

    if field_type == "COO":
        if region == "FR":
            return ["fabrique en"], ["made in", "hecho en", "rn", "ca"]
        if region == "SP":
            return ["hecho en"], ["made in", "fabrique en", "rn", "ca"]
        return ["made in"], ["fabrique en", "hecho en", "rn", "ca"]

    if field_type in {"RN", "CA"}:
        return ["rn ", "ca ", "rn", "ca"], ["machine wash", "laver", "lavar", "made in", "hecho en"]

    return [], []


def _visual_region_groups(page, field_name):
    """Return only the semantic artwork region belonging to one field family."""
    groups = _visual_all_groups(page)
    if not groups:
        return []

    starts, stops = _visual_field_markers(field_name)
    starts = [_visual_norm(x) for x in starts if _visual_norm(x)]
    stops = [_visual_norm(x) for x in stops if _visual_norm(x)]
    if not starts:
        return []

    field_type = get_field_type(field_name)
    candidates = []

    # Find every possible anchor; later fields such as FIB_Mexico/FIB_SP need
    # occurrence-aware selection instead of always taking the first "cuerpo".
    for idx, group in enumerate(groups):
        line_text = _visual_norm(_visual_group_text(group))
        if not line_text:
            continue
        if any(marker in line_text for marker in starts):
            candidates.append(idx)

    if not candidates:
        return []

    def collect_from(start_idx):
        selected = []
        anchor_group = groups[start_idx]
        anchor_words = _boxes_from_words(anchor_group)
        anchor_left = anchor_words[0][0] if anchor_words else 0
        for idx in range(start_idx, min(len(groups), start_idx + (10 if field_type == "CONTENT" else 18 if field_type == "CARE" else 3))):
            group = groups[idx]
            line_text = _visual_norm(_visual_group_text(group))
            if not line_text:
                continue

            # CONTENT and CARE are usually columnar. Keep the region in the same
            # physical text column as its anchor so an adjacent IMPORTED BY block
            # cannot be accidentally absorbed into the highlight.
            group_boxes = _boxes_from_words(group)
            group_left = group_boxes[0][0] if group_boxes else anchor_left
            if field_type in {"CONTENT", "CARE"} and idx > start_idx:
                if abs(group_left - anchor_left) > 85:
                    continue

            if idx > start_idx and any(stop in line_text for stop in stops):
                break

            selected.append(group)

            if field_type == "CONTENT":
                has_pct = bool(re.search(r"\b\d{1,3}\s*%", line_text))
                has_material = any(
                    material in line_text
                    for material in (
                        "polyester", "spandex", "elastane", "elasthanne", "cotton", "nylon",
                        "rayon", "viscose", "acrylic", "linen", "wool", "elastodiene",
                        "polyamide", "poliester", "poliéster", "elastano", "elastán"
                    )
                )
                if idx > start_idx and not (has_pct or has_material):
                    if len(selected) > 1:
                        selected.pop()
                        break
            elif field_type == "CARE":
                if idx > start_idx and any(stop in line_text for stop in stops):
                    break

        return selected

    # Strongly anchored fields: choose the first candidate that contains the
    # compared semantic family. For repeated Spanish regions, use all matching
    # candidate blocks for FIB_SP_Mexico, but one block for FIB_Mexico/FIB_SP.
    compact = re.sub(r"[^a-z0-9]", "", str(field_name).casefold())
    if "fibspmexico" in compact:
        combined = []
        for start_idx in candidates[:2]:
            block = collect_from(start_idx)
            if block:
                combined.extend(block)
        return combined

    for start_idx in candidates:
        block = collect_from(start_idx)
        if block:
            return block
    return []


def _visual_region_boxes(page, field_name, actual_value=""):
    """Find a semantic region when exact text occurrence is insufficient."""
    groups = _visual_region_groups(page, field_name)
    if not groups:
        return []
    boxes = []
    for group in groups:
        boxes.extend(_boxes_from_words(group))
    return boxes


def _visual_find_size_field_boxes(page, actual_value, expected_value="", field_name=""):
    """Resolve Main_Size/OS_Size_N to one coherent physical size block."""
    groups = _visual_group_words(page)
    if not groups:
        return []

    size_items = []
    for idx, group in enumerate(groups):
        text = _visual_group_text(group)
        if not _visual_raw_size_like_text(text):
            continue
        boxes = _boxes_from_words(group)
        if not boxes:
            continue
        box = boxes[0]
        size_items.append({
            "index": idx,
            "text": text,
            "group": group,
            "box": box,
            "cx": (box[0] + box[2]) / 2.0,
            "cy": (box[1] + box[3]) / 2.0,
        })

    if len(size_items) < 2:
        return []

    # Identify the dominant size column by X-position. This is much safer than
    # treating every numeric-looking line on the page as part of one OSZ block.
    clusters = []
    for item in sorted(size_items, key=lambda x: x["cx"]):
        placed = False
        for cluster in clusters:
            avg_x = sum(x["cx"] for x in cluster) / len(cluster)
            if abs(item["cx"] - avg_x) <= 95:
                cluster.append(item)
                placed = True
                break
        if not placed:
            clusters.append([item])
    dominant = max(clusters, key=lambda c: (len(c), max(x["cy"] for x in c) - min(x["cy"] for x in c)))
    dominant.sort(key=lambda x: x["cy"])

    # Remove occasional non-size fragments that happen to land in the same X
    # band. Valid size rows normally contain a recognized alpha size token or a
    # regional numeric-size pattern.
    size_token_re = re.compile(
        r"^\s*(?:xxxs|xxs|xs|s|m|l|xl|xxl|xxxl|tp|p|g|gg|tg|ttg|ech|ch|eg|ee|eeg)\b",
        re.IGNORECASE,
    )
    numeric_size_re = re.compile(r"^\s*\d{1,3}\s*/\s*\d{1,3}\s*[-–]\s*\d", re.IGNORECASE)
    cleaned = [
        item for item in dominant
        if size_token_re.search(item["text"]) or numeric_size_re.search(item["text"])
    ]
    if len(cleaned) >= 2:
        dominant = cleaned

    # Split the dominant column at the noticeably larger vertical whitespace
    # between logical size blocks. Within a block, rows are tightly packed;
    # between blocks the gap is visibly larger. This avoids treating TG/EG as
    # new blocks just because they are themselves alphabetic size labels.
    blocks = []
    current_block = []
    previous = None
    raw_gaps = []
    for item in dominant:
        if previous is not None:
            raw_gaps.append(max(0.0, item["box"][1] - previous["box"][3]))
        previous = item
    positive_gaps = sorted(g for g in raw_gaps if g >= 0)
    median_gap = positive_gaps[len(positive_gaps) // 2] if positive_gaps else 5.0
    split_gap = max(12.0, median_gap * 1.55)

    previous = None
    for item in dominant:
        if previous is not None:
            gap = item["box"][1] - previous["box"][3]
            if gap > split_gap and current_block:
                blocks.append(current_block)
                current_block = []
        current_block.append(item)
        previous = item
    if current_block:
        blocks.append(current_block)

    blocks = [block for block in blocks if block]
    if not blocks:
        return []

    compact_field = re.sub(r"[^a-z0-9]", "", str(field_name).casefold())
    if compact_field == "mainsize":
        block_index = 0
    else:
        match = re.search(r"(?:osz|ossize)(\d+)$", compact_field)
        block_index = int(match.group(1)) - 1 if match else 0

    if block_index < 0 or block_index >= len(blocks):
        return []

    # If the selected block has a strong textual match to the actual output,
    # use it. Otherwise retain the semantic ordinal because OS_Size_N is itself
    # the field's identity.
    chosen = blocks[block_index]
    target = str(actual_value or expected_value or "").strip()
    block_text = " ".join(item["text"] for item in chosen)
    if target and _visual_text_similarity(target, block_text) < 0.38:
        # A single missing/merged line can shift a block boundary. Search for a
        # nearby block with a materially stronger match without abandoning the
        # ordinal field identity on weak evidence.
        ranked = sorted(
            ((
                _visual_text_similarity(target, " ".join(item["text"] for item in block)),
                idx,
                block,
            ) for idx, block in enumerate(blocks)),
            reverse=True,
        )
        if ranked and ranked[0][0] >= 0.62:
            chosen = ranked[0][2]

    boxes = []
    for item in chosen:
        boxes.extend(_boxes_from_words(item["group"]))
    return boxes


def _visual_field_uses_block_mapping(field_name, expected_value="", actual_value=""):
    field_type = get_field_type(field_name)
    if field_type in {"CARE", "CONTENT"}:
        return True
    if field_type in {"ATTRIBUTE", "GENERAL", "BRAND"}:
        sample = str(actual_value or expected_value or "").strip()
        return len(tokenize(sample)) >= 5 or len(sample) >= 45
    return False


def _visual_find_field_boxes(page, field_name, expected_value, actual_value, status):
    """
    Evidence-linked visual mapper.

    The validation result remains authoritative.  This function only answers:
    "Which physical pixels correspond to the evidence that the validator used?"
    It deliberately avoids a generic first-match search for structured fields.
    """
    field_type = get_field_type(field_name)
    actual_value = str(actual_value or "").strip()
    expected_value = str(expected_value or "").strip()

    if not actual_value or actual_value.casefold() in {"not found", "-", "—"}:
        if status in {"PASS", "FAIL"} and field_type in {"CONTENT", "CARE", "COO"}:
            return _visual_region_boxes(page, field_name, actual_value)
        return []

    compact_field = re.sub(r"[^a-z0-9]", "", str(field_name).casefold())
    is_osz = bool(re.search(r"(?:osz|ossize)\d+$", compact_field))

    # ------------------------------------------------------------------
    # 1. Multi-line structured blocks. This is the highest-confidence path
    #    for Main_Size and OS_Size_N and fixes the old "first size column"
    #    behaviour.
    # ------------------------------------------------------------------
    if is_osz:
        if compact_field.startswith("osz") and not compact_field.startswith("ossize"):
            boxes = _visual_find_osz_field_boxes(
                page,
                field_name,
                actual_value=actual_value,
                expected_value=expected_value,
            )
            if boxes:
                return boxes
        else:
            boxes = _visual_find_size_field_boxes(
                page,
                actual_value,
                expected_value,
                field_name=field_name,
            )
            if boxes:
                return boxes
    elif field_type == "SIZE":
        boxes = _visual_find_size_field_boxes(
            page,
            actual_value,
            expected_value,
            field_name=field_name,
        )
        if boxes:
            return boxes

    # ------------------------------------------------------------------
    # 2. CONTENT / CARE use semantic regional anchors first. Exact phrase
    #    search is only a fallback because several language fields may contain
    #    identical material/care text.
    # ------------------------------------------------------------------
    if field_type in {"CONTENT", "CARE"}:
        region_groups = _visual_region_groups(page, field_name)
        if region_groups:
            # For long semantic fields the region itself is the evidence.
            # Returning the complete region is more trustworthy than trying to
            # rediscover one repeated phrase inside a multilingual block.
            region_boxes = []
            for group in region_groups:
                region_boxes.extend(_boxes_from_words(group))
            if region_boxes:
                return region_boxes

    # ------------------------------------------------------------------
    # 3. COO / RN / CA: use the exact scalar first, but search the semantic
    #    line if the number is embedded in RN/CA combined text.
    # ------------------------------------------------------------------
    if field_type in {"COO", "RN", "CA"}:
        if field_type in {"RN", "CA"}:
            direct = _visual_find_direct_scalar_box(page, actual_value)
            if direct:
                return direct
        exact = _visual_find_text_occurrences(page, actual_value, min_score=0.74)
        if exact:
            return exact
        region_boxes = _visual_region_boxes(page, field_name, actual_value)
        if region_boxes:
            return region_boxes

    # ------------------------------------------------------------------
    # 4. Symbols / identifiers are safest against the PDF text layer.
    # ------------------------------------------------------------------
    if field_type == "SYMBOL":
        boxes = _find_visual_symbol_boxes(page, actual_value)
        if not boxes:
            boxes = _find_visual_symbol_boxes(page, expected_value)
        if boxes:
            return boxes

    if field_type == "IDENTIFIER":
        boxes = _visual_find_direct_scalar_box(page, actual_value)
        if boxes:
            return boxes

    # ------------------------------------------------------------------
    # 5. General scalar/text fields.
    # ------------------------------------------------------------------
    boxes = _visual_find_text_occurrences(page, actual_value, min_score=0.80)
    if boxes:
        return boxes

    # ------------------------------------------------------------------
    # 6. FAIL-specific token fallback.  This is deliberately last so an
    #    isolated "S", "18" or other tiny token cannot hijack a structured
    #    field's highlight.
    # ------------------------------------------------------------------
    if status == "FAIL":
        expected_tokens = tokenize(expected_value)
        actual_tokens = tokenize(actual_value)
        try:
            from difflib import SequenceMatcher
            matcher = SequenceMatcher(None, expected_tokens, actual_tokens, autojunk=False)
            for tag, _a1, _a2, b1, b2 in matcher.get_opcodes():
                if tag not in {"replace", "insert"}:
                    continue
                fragment = " ".join(actual_tokens[b1:b2]).strip()
                if not fragment or len(fragment) < 2:
                    continue
                boxes = _visual_find_text_occurrences(page, fragment, min_score=0.80)
                if boxes:
                    return boxes
        except Exception:
            pass

    for token in sorted(tokenize(actual_value), key=lambda x: (-len(x), x)):
        if len(token) < 2 and len(tokenize(actual_value)) > 1:
            continue
        boxes = _visual_find_text_occurrences(page, token, min_score=0.82)
        if boxes:
            return boxes

    return []

def _merge_nearby_boxes(boxes, gap=8):
    """Merge boxes that overlap or are nearly adjacent on the same line."""
    if not boxes:
        return []
    cleaned = []
    for box in boxes:
        try:
            l, t, r, b = [int(v) for v in box]
        except Exception:
            continue
        if r > l and b > t:
            cleaned.append((l, t, r, b))
    if not cleaned:
        return []

    cleaned.sort(key=lambda b: (b[1], b[0]))
    merged = []
    for box in cleaned:
        if not merged:
            merged.append(box)
            continue
        ml, mt, mr, mb = merged[-1]
        l, t, r, b = box
        vertical_overlap = min(mb, b) - max(mt, t)
        horizontal_gap = max(0, max(l, ml) - min(r, mr))
        if vertical_overlap > -gap and horizontal_gap <= gap:
            merged[-1] = (min(ml, l), min(mt, t), max(mr, r), max(mb, b))
        else:
            merged.append(box)
    return merged


def _visual_marker_position(box, radius, image_w, image_h, occupied):
    """Place numbered marker just outside the highlight without covering text."""
    x1, y1, x2, y2 = box
    candidates = [
        (x1 - radius - 4, y1 - radius - 4),
        (x2 + radius + 4, y1 - radius - 4),
        (x1 - radius - 4, y2 + radius + 4),
        (x2 + radius + 4, y2 + radius + 4),
    ]
    for cx, cy in candidates:
        cx = max(radius + 3, min(image_w - radius - 3, cx))
        cy = max(radius + 3, min(image_h - radius - 3, cy))
        rect = (cx - radius, cy - radius, cx + radius, cy + radius)
        if not any(_rect_intersects(rect, existing, pad=4) for existing in occupied):
            return cx, cy
    return max(radius + 3, min(image_w - radius - 3, x1)), max(radius + 3, min(image_h - radius - 3, y1))


def build_highlighted_page_image(page, page_report, field_colors=None):
    """Render clean, evidence-driven QC highlights on the original artwork."""
    image_bytes = page.get("image_bytes")
    if not image_bytes:
        return None

    base = Image.open(BytesIO(image_bytes)).convert("RGBA")
    if page_report is None or page_report.empty:
        return base.convert("RGB")

    overlay = Image.new("RGBA", base.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    field_colors = field_colors or {}

    field_order = []
    for _, row in page_report.iterrows():
        field = str(row.get("FIELD", "")).strip()
        if field and field not in field_order:
            field_order.append(field)
    number_by_field = {field: i + 1 for i, field in enumerate(field_order)}

    marker_occupied = []
    annotated = []
    min_dim = min(base.width, base.height)
    pad = max(3, int(min_dim * 0.0022))
    radius = max(13, int(min_dim * 0.010))

    for _, row in page_report.iterrows():
        status = str(row.get("STATUS", "")).strip().upper()
        if status not in {"PASS", "FAIL"}:
            continue

        field_name = str(row.get("FIELD", "")).strip()
        if not field_name:
            continue

        expected_value = str(row.get("ORDER FORM DATA", "") or "").strip()
        actual_value = str(row.get("PDF OUTPUT", "") or "").strip()
        boxes = _visual_find_field_boxes(
            page,
            field_name,
            expected_value,
            actual_value,
            status,
        )
        boxes = _merge_nearby_boxes(boxes, gap=max(6, pad * 2))
        if not boxes:
            continue

        color = field_colors.get(field_name, FIELD_VISUAL_COLORS[0])
        rgb = _hex_rgb(color)

        for left, top, right, bottom in boxes:
            left = max(0, left - pad)
            top = max(0, top - pad)
            right = min(base.width - 1, right + pad)
            bottom = min(base.height - 1, bottom + pad)
            if right <= left or bottom <= top:
                continue

            if status == "FAIL":
                # Strong enough to notice, but still translucent so the artwork
                # remains readable underneath.
                draw.rounded_rectangle(
                    (left, top, right, bottom),
                    radius=max(4, pad),
                    fill=(218, 54, 51, 92),
                    outline=rgb + (255,),
                    width=max(2, int(min_dim * 0.0014)),
                )
                draw.rounded_rectangle(
                    (max(0, left - 2), max(0, top - 2), min(base.width - 1, right + 2), min(base.height - 1, bottom + 2)),
                    radius=max(4, pad + 1),
                    outline=(218, 54, 51, 235),
                    width=max(2, int(min_dim * 0.0009)),
                )
            else:
                draw.rounded_rectangle(
                    (left, top, right, bottom),
                    radius=max(4, pad),
                    fill=rgb + (54,),
                    outline=rgb + (230,),
                    width=max(2, int(min_dim * 0.0011)),
                )

        # One small number marker per field; no large labels are painted over the artwork.
        anchor_box = (
            min(box[0] for box in boxes),
            min(box[1] for box in boxes),
            max(box[2] for box in boxes),
            max(box[3] for box in boxes),
        )
        cx, cy = _visual_marker_position(
            anchor_box,
            radius,
            base.width,
            base.height,
            marker_occupied,
        )
        marker_rect = (cx - radius, cy - radius, cx + radius, cy + radius)
        marker_occupied.append(marker_rect)

        marker_color = (218, 54, 51) if status == "FAIL" else rgb
        draw.ellipse(
            marker_rect,
            fill=(255, 255, 255, 245),
            outline=marker_color + (255,),
            width=max(2, int(min_dim * 0.0011)),
        )

        label = str(number_by_field.get(field_name, ""))
        font = _load_visual_fonts()[1] or _load_visual_fonts()[0]
        tw, th = _text_size(draw, label, font)
        draw.text(
            (int(cx - tw / 2), int(cy - th / 2 - 1)),
            label,
            fill=marker_color + (255,),
            font=font,
        )

        annotated.append({
            "field": field_name,
            "status": status,
            "number": number_by_field.get(field_name, 0),
            "boxes": boxes,
        })

    result = Image.alpha_composite(base, overlay).convert("RGB")
    # Keep the annotation metadata available to the UI without putting labels on
    # the artwork itself. This function intentionally returns only an image so it
    # remains compatible with the existing Excel-report API.
    return result

def clean_pdf_line(line):
    if not line:
        return ""
    line = str(line).replace("\u200b", "").replace("\ufeff", "").strip()
    line = re.sub(r"^\s*n\s+(?=[A-Za-z])", "", line)
    return line.strip()


# =========================================================
# ATOMIC PAGE MODEL
# =========================================================

def build_page_lines(page_text, product_type):
    """
    Create atomic lines. We do NOT create overlapping blocks and then consume
    blocks, because consuming an overlapping block can accidentally consume
    unrelated values. Instead, every decision is tied to actual line IDs.
    """
    raw_lines = str(page_text or "").splitlines()
    lines = []

    panel_pattern = re.compile(
        r"^\s*(?:panel\s*)?(\d{1,3})\s*$",
        re.IGNORECASE
    )

    for raw_index, raw in enumerate(raw_lines):
        line = clean_pdf_line(raw)
        if not line:
            continue
        if len(line) > 1500:
            continue

        # PFL: remove only explicit panel labels. Do not discard standalone
        # numeric lines because those may be genuine OSZ/size data.
        if product_type == "PFL" and re.match(r"^\s*panel\s*[-#: ]?\d{1,3}\s*$", line, re.IGNORECASE):
            continue

        lines.append({
            "line_id": len(lines),
            "raw_index": raw_index,
            "text": line,
            "norm": normalize_text(line)
        })

    return lines


def _select_comparison_text(page):
    """Choose the cleanest comparison text while preserving OCR as fallback.

    For editable PDFs, the PDF text layer often preserves reading order and
    structured multi-line content much better than OCR. OCR remains the fallback
    for scanned/non-editable artwork. This is a comparison-source decision only;
    visual OCR boxes are still retained separately for annotations.
    """
    direct = str(page.get("direct_text", "") or "").strip()
    ocr = str(page.get("ocr_text", "") or "").strip()

    if direct:
        direct_quality = _text_quality_score(direct)
        ocr_quality = _text_quality_score(ocr) if ocr else 0
        direct_alnum = len(re.findall(r"[A-Za-z0-9]", direct))

        # Prefer a meaningful PDF text layer unless it is dramatically weaker
        # than OCR. The 0.55 guard protects genuinely bad/partial text layers.
        if direct_alnum >= 20 and (not ocr or direct_quality >= max(80, ocr_quality * 0.55)):
            return direct, "pdf_text"

    if ocr:
        return ocr, "ocr"

    if direct:
        return direct, "pdf_text"

    return str(page.get("text", "") or ""), str(page.get("source_type", "pdf_text"))


def build_page_state(page, product_type):
    if not isinstance(page, dict):
        raise TypeError(
            f"Output page data is malformed. Expected a page dictionary, got {type(page).__name__}."
        )

    comparison_text, comparison_source = _select_comparison_text(page)
    lines = build_page_lines(
        comparison_text,
        product_type
    )

    # Defensive validation: every line must be a dictionary with the fields
    # consumed by the matching engine. This prevents the vague
    # "list indices must be integers or slices, not str" failure.
    invalid_lines = [
        index for index, line in enumerate(lines)
        if not isinstance(line, dict)
        or "line_id" not in line
        or "text" not in line
        or "norm" not in line
    ]
    if invalid_lines:
        raise TypeError(
            f"Output page {page.get('page', '?')} contains malformed OCR/text lines: {invalid_lines[:5]}"
        )

    # A custom-font symbol may be perfectly preserved in the PDF text layer
    # while OCR converts the glyph into an unrelated character. Keep an
    # auxiliary raw direct-text line list for SYMBOL validation only.
    direct_text = str(page.get("direct_text", "") or "")
    direct_lines = []
    for raw_index, raw in enumerate(direct_text.splitlines()):
        clean = str(raw or "").replace("\u200b", "").replace("\ufeff", "").strip()
        if not clean:
            continue
        direct_lines.append({
            "line_id": 1000000 + raw_index,
            "raw_index": raw_index,
            "text": clean,
            "norm": normalize_text(clean),
        })

    return {
        "page": page.get("page"),
        "source_type": comparison_source,
        "comparison_text": comparison_text,
        "lines": lines,
        "direct_lines": direct_lines,
        "consumed": set(),
        "consumed_spans": {},
        # Keep the existing line-based OSZ runs intact.  The additional
        # coordinate-aware candidates are used only as a fallback when OCR
        # formatting has mixed multiple standalone numbers into one line.
        "osz_runs": extract_standalone_numeric_runs_from_lines(lines),
        # Keep a second, immutable numeric sequence from the PDF text layer.
        # This is a targeted OSZ fallback only; it does not alter normal field
        # matching or consumption. It protects clean PDF text when OCR has
        # merged some sequence values into neighbouring lines.
        "osz_direct_runs": extract_standalone_numeric_runs_from_lines(
            build_page_lines(page.get("direct_text", ""), product_type)
        ),
        "osz_sequence_candidates": extract_osz_sequence_candidates(page),
    }


def extract_standalone_numeric_runs_from_lines(lines):
    runs = []
    current = []

    for line in lines:
        if re.fullmatch(r"[-+]?\d+", line["norm"]):
            current.append(line)
        else:
            if current:
                runs.append(current)
                current = []

    if current:
        runs.append(current)

    return [run for run in runs if run]


def _osz_numeric_words(page):
    """Return standalone integer OCR tokens in stored artwork-image coordinates."""
    result = []
    sx, sy = _page_ocr_scale(page)
    for raw_word in page.get("ocr_words", []) or []:
        if not isinstance(raw_word, dict):
            continue
        raw = str(raw_word.get("text", "")).strip()
        if not re.fullmatch(r"\d+", raw):
            continue
        word = _scaled_word(raw_word, sx, sy)
        if word["width"] <= 0 or word["height"] <= 0:
            continue
        left = int(word["left"])
        top = int(word["top"])
        width = int(word["width"])
        height = int(word["height"])
        result.append({
            "value": raw,
            "left": left,
            "top": top,
            "width": width,
            "height": height,
            "center_x": left + width / 2.0,
            "center_y": top + height / 2.0,
        })
    return result


def _visual_find_osz_field_boxes(page, field_name, actual_value="", expected_value=""):
    """Map standalone OSZ/OSZ_N values to their exact numeric token boxes."""
    compact = re.sub(r"[^a-z0-9]", "", str(field_name).casefold())
    match = re.search(r"(?:osz)(\d+)$", compact)
    if not match:
        return []
    index = int(match.group(1))
    if index <= 0:
        return []

    candidates = extract_osz_sequence_candidates(page)
    if not candidates:
        return []

    expected_num = normalize_numeric(expected_value)
    actual_num = normalize_numeric(actual_value)

    ranked = []
    for candidate in candidates:
        items = candidate.get("items", [])
        if len(items) < index:
            continue
        item = items[index - 1]
        value = normalize_numeric(item.get("value", ""))
        score = float(candidate.get("score", 0.0))
        if actual_num is not None and value == actual_num:
            score += 45.0
        if expected_num is not None and value == expected_num:
            score += 35.0
        # Prefer a genuinely long sequence over short random numeric runs.
        score += min(35.0, len(items) * 2.5)
        ranked.append((score, item))

    if not ranked:
        return []

    ranked.sort(key=lambda item: item[0], reverse=True)
    item = ranked[0][1]
    return [_boxes_from_words([item])[0]] if _boxes_from_words([item]) else []


def _score_osz_sequence(items, orientation):
    if len(items) < 3:
        return -1.0

    if orientation == "vertical":
        ordered = sorted(items, key=lambda x: (x["center_y"], x["center_x"]))
        axis = [item["center_y"] for item in ordered]
        cross = [item["center_x"] for item in ordered]
    else:
        ordered = sorted(items, key=lambda x: (x["center_x"], x["center_y"]))
        axis = [item["center_x"] for item in ordered]
        cross = [item["center_y"] for item in ordered]

    gaps = [axis[i + 1] - axis[i] for i in range(len(axis) - 1)]
    positive_gaps = [g for g in gaps if g > 0]
    if not positive_gaps:
        return -1.0

    median_gap = sorted(positive_gaps)[len(positive_gaps) // 2]
    if median_gap <= 0:
        return -1.0

    cross_spread = max(cross) - min(cross)
    # OSZ lists can have some OCR jitter, but should remain visually aligned.
    alignment_penalty = cross_spread / max(1.0, median_gap)

    regularity = sum(
        1.0
        for gap in positive_gaps
        if 0.45 * median_gap <= gap <= 1.80 * median_gap
    ) / len(positive_gaps)

    # Prefer longer, well-aligned runs.  A minimum regularity prevents random
    # page numbers or price fragments from becoming an OSZ sequence.
    if regularity < 0.60 or alignment_penalty > 0.55:
        return -1.0

    return len(ordered) * 10.0 + regularity * 10.0 - alignment_penalty * 5.0


def extract_osz_sequence_candidates(page):
    """Build dynamic OSZ candidates from artwork geometry, independent of field count."""
    words = _osz_numeric_words(page)
    if len(words) < 3:
        return []

    candidates = []
    for orientation in ("vertical", "horizontal"):
        # Build loose spatial groups around a common axis.
        if orientation == "vertical":
            words_sorted = sorted(words, key=lambda x: x["center_x"])
            axis_key = "center_x"
        else:
            words_sorted = sorted(words, key=lambda x: x["center_y"])
            axis_key = "center_y"

        groups = []
        for word in words_sorted:
            added = False
            for group in groups:
                reference = sum(item[axis_key] for item in group) / len(group)
                tolerance = max(18.0, min(90.0, max(word["width"], word["height"]) * 1.6))
                if abs(word[axis_key] - reference) <= tolerance:
                    group.append(word)
                    added = True
                    break
            if not added:
                groups.append([word])

        for group in groups:
            score = _score_osz_sequence(group, orientation)
            if score < 0:
                continue
            ordered = (
                sorted(group, key=lambda x: (x["center_y"], x["center_x"]))
                if orientation == "vertical"
                else sorted(group, key=lambda x: (x["center_x"], x["center_y"]))
            )
            # Deduplicate identical sequences.
            signature = tuple((item["value"], round(item["center_x"] / 5), round(item["center_y"] / 5)) for item in ordered)
            if any(existing["signature"] == signature for existing in candidates):
                continue
            candidates.append({
                "orientation": orientation,
                "items": ordered,
                "score": score,
                "signature": signature,
            })

    candidates.sort(key=lambda item: item["score"], reverse=True)
    return candidates[:12]


def _osz_candidate_matches_expected(candidate, index, expected_num):
    items = candidate.get("items", [])
    if len(items) < index:
        return False
    return normalize_numeric(items[index - 1].get("value")) == expected_num


def line_is_available(line, state):
    # A line remains available when only part of it has already been assigned.
    # This is required for combined output lines containing multiple fields.
    return line["line_id"] not in state.get("consumed", set())


def consume_lines(state, lines):
    for line in lines:
        state.setdefault("consumed", set()).add(line["line_id"])


def consume_match(state, match_info):
    """Safely consume either a multi-line match or one matched span."""
    if not isinstance(match_info, dict):
        return

    matched_lines = match_info.get("lines")
    if matched_lines:
        if isinstance(matched_lines, dict):
            matched_lines = [matched_lines]
        valid_lines = [
            line for line in matched_lines
            if isinstance(line, dict) and "line_id" in line
        ]
        if valid_lines:
            consume_lines(state, valid_lines)
            return

    line = match_info.get("line")
    start = match_info.get("start")
    end = match_info.get("end")
    if isinstance(line, dict) and start is not None and end is not None:
        consume_span(state, line, int(start), int(end))


def join_lines(lines):
    return " ".join(
        line["text"]
        for line in lines
    ).strip()


# =========================================================
# SAFE EXACT MATCHING
# =========================================================

def _used_spans(state, line_id):
    return state.setdefault("consumed_spans", {}).get(line_id, [])


def _span_overlaps(a_start, a_end, used_spans):
    return any(a_start < end and a_end > start for start, end in used_spans)


def consume_span(state, line, start, end):
    state.setdefault("consumed_spans", {}).setdefault(line["line_id"], []).append((start, end))


def _normalized_match_positions(expected_norm, actual_norm, field_type):
    if not expected_norm or not actual_norm:
        return []

    # Structured scalar/numeric values must match as whole tokens.
    if normalize_numeric(expected_norm) is not None:
        return [
            (m.start(), m.end())
            for m in re.finditer(
                rf"(?<![A-Za-z0-9]){re.escape(expected_norm)}(?![A-Za-z0-9])",
                actual_norm
            )
        ]

    # Identifiers: normal whole-token match. The asymmetric prefix rule is
    # handled separately by find_identifier_match().
    if field_type == "IDENTIFIER":
        return [
            (m.start(), m.end())
            for m in re.finditer(
                rf"(?<![A-Za-z0-9]){re.escape(expected_norm)}(?![A-Za-z0-9])",
                actual_norm
            )
        ]

    # A single alphanumeric token must not match inside another word.
    if len(expected_norm.split()) == 1 and re.fullmatch(r"[A-Za-z0-9%#]+", expected_norm):
        return [
            (m.start(), m.end())
            for m in re.finditer(
                rf"(?<![A-Za-z0-9]){re.escape(expected_norm)}(?![A-Za-z0-9])",
                actual_norm
            )
        ]

    # Multi-word text can occur inside a combined artwork line.
    return [
        (m.start(), m.end())
        for m in re.finditer(
            re.escape(expected_norm),
            actual_norm
        )
    ]


def safe_exact_match(expected, actual, field_name):
    field_type = get_field_type(field_name)

    if field_type == "SYMBOL":
        exp = normalize_symbol_text(expected)
        act = normalize_symbol_text(actual)
        return bool(exp and exp == act)

    expected_numeric = normalize_numeric(expected)
    if expected_numeric is not None:
        actual_numbers = re.findall(
            r"(?<![A-Za-z0-9])[-+]?\d+(?:\.\d+)?(?![A-Za-z0-9])",
            normalize_text(actual)
        )
        return expected_numeric in {
            normalize_numeric(number)
            for number in actual_numbers
        }

    exp = normalize_text(expected)
    act = normalize_text(actual)

    if not exp or not act:
        return False

    if exp == act:
        return True

    positions = _normalized_match_positions(exp, act, field_type)
    return bool(positions)


def find_identifier_match(expected, state):
    expected_norm = normalize_text(expected)
    if not expected_norm:
        return None

    token_pattern = re.compile(r"[A-Za-z0-9][A-Za-z0-9_\-/]*")

    # First: exact identifier token.
    for line in state["lines"]:
        if not line_is_available(line, state):
            continue

        actual_norm = line["norm"]

        for match in token_pattern.finditer(actual_norm):
            token = match.group(0)
            if token.casefold() != expected_norm.casefold():
                continue

            if _span_overlaps(match.start(), match.end(), _used_spans(state, line["line_id"])):
                continue

            return {
                "kind": "PASS",
                "line": line,
                "start": match.start(),
                "end": match.end(),
                "actual": token,
                "difference": "—",
                "match_type": "IDENTIFIER_EXACT"
            }

    # Second: the Order Form has the base code and the PDF has an additional
    # suffix/static portion. Base -> longer output is intentionally PASS.
    for line in state["lines"]:
        if not line_is_available(line, state):
            continue

        actual_norm = line["norm"]

        for match in token_pattern.finditer(actual_norm):
            token = match.group(0)

            if not token.casefold().startswith(expected_norm.casefold()):
                continue

            if len(token) <= len(expected_norm):
                continue

            if _span_overlaps(match.start(), match.end(), _used_spans(state, line["line_id"])):
                continue

            return {
                "kind": "PASS",
                "line": line,
                "start": match.start(),
                "end": match.start() + len(expected_norm),
                "actual": token,
                "difference": "Base identifier matched; PDF contains additional suffix/static characters.",
                "match_type": "IDENTIFIER_BASE_PLUS_SUFFIX"
            }

    # Third: when both sides are extended identifiers that share a meaningful
    # code prefix but differ, report a genuine FAIL. Example:
    # Order Form USX690 vs PDF USX609.
    for line in state["lines"]:
        if not line_is_available(line, state):
            continue

        actual_norm = line["norm"]
        for match in token_pattern.finditer(actual_norm):
            token = match.group(0)
            if token.casefold() == expected_norm.casefold():
                continue
            common = 0
            for left, right in zip(expected_norm.casefold(), token.casefold()):
                if left != right:
                    break
                common += 1
            if common < 3:
                continue
            if len(token) < 2:
                continue
            if _span_overlaps(match.start(), match.end(), _used_spans(state, line["line_id"])):
                continue

            return {
                "kind": "FAIL",
                "line": line,
                "start": match.start(),
                "end": match.end(),
                "actual": token,
                "difference": (
                    f"Identifier mismatch: Order Form has '{expected}', "
                    f"but PDF has '{token}'."
                ),
                "match_type": "IDENTIFIER_PREFIX_MISMATCH"
            }

    # Fourth: reverse condition — PDF has only a shorter base and Order Form
    # expects the longer identifier. This must NOT pass.
    for line in state["lines"]:
        if not line_is_available(line, state):
            continue

        actual_norm = line["norm"]

        for match in token_pattern.finditer(actual_norm):
            token = match.group(0)

            if len(token) < 2:
                continue

            if not expected_norm.casefold().startswith(token.casefold()):
                continue

            if token.casefold() == expected_norm.casefold():
                continue

            if _span_overlaps(match.start(), match.end(), _used_spans(state, line["line_id"])):
                continue

            return {
                "kind": "FAIL",
                "line": line,
                "start": match.start(),
                "end": match.end(),
                "actual": token,
                "difference": (
                    f"Identifier incomplete: Order Form has '{expected}', "
                    f"but PDF has '{token}'."
                ),
                "match_type": "IDENTIFIER_SHORTER_OUTPUT"
            }

    return None


def find_exact_lines(expected, field_name, state, max_window=8):
    """Find an exact value while allowing multiple independent fields on one PDF line."""
    lines = state["lines"]
    available = [line for line in lines if line_is_available(line, state)]
    if not available:
        return None

    field_type = get_field_type(field_name)

    if field_type == "IDENTIFIER":
        identifier = find_identifier_match(expected, state)
        if identifier and identifier["kind"] == "PASS":
            return [identifier["line"]], identifier
        return None

    numeric_expected = normalize_numeric(expected)
    if numeric_expected is not None:
        for line in available:
            positions = _normalized_match_positions(
                numeric_expected,
                line["norm"],
                field_type
            )
            for start, end in positions:
                if not _span_overlaps(start, end, _used_spans(state, line["line_id"])):
                    return [line], {
                        "kind": "PASS",
                        "line": line,
                        "start": start,
                        "end": end,
                        "actual": numeric_expected,
                        "difference": "—",
                        "match_type": "NUMERIC_TOKEN"
                    }
        return None

    preferred_single_line = field_type in {
        "SYMBOL", "RN", "OSZ", "SIZE", "COLOR", "GENDER",
        "BATCH", "QUANTITY", "BRAND", "ATTRIBUTE", "GENERAL"
    }
    search_window = 1 if preferred_single_line else max_window

    # First search individual lines. This is essential for combined output such
    # as 'SF8334 67 YZP': each value can be consumed independently.
    for line in available:
        positions = _normalized_match_positions(
            normalize_text(expected),
            line["norm"],
            field_type
        )
        for start, end in positions:
            if not _span_overlaps(start, end, _used_spans(state, line["line_id"])):
                return [line], {
                    "kind": "PASS",
                    "line": line,
                    "start": start,
                    "end": end,
                    "actual": line["norm"][start:end],
                    "difference": "—",
                    "match_type": "EXACT_COMBINED_LINE"
                }

    if search_window <= 1:
        return None

    # Multi-line exact match for wrapped care/content/general text.
    for window_size in range(2, search_window + 1):
        for start_index in range(0, len(available) - window_size + 1):
            candidate = available[start_index:start_index + window_size]
            ids = [line["line_id"] for line in candidate]
            if ids != list(range(ids[0], ids[-1] + 1)):
                continue

            text = join_lines(candidate)
            if safe_exact_match(expected, text, field_name):
                return candidate, {
                    "kind": "PASS",
                    "lines": candidate,
                    "difference": "—",
                    "match_type": "EXACT_MULTI_LINE"
                }

    return None

# =========================================================
# STRUCTURED EXTRACTORS
# =========================================================

def extract_coo_value(text):
    normalized = normalize_text(text)
    if not normalized:
        return None

    patterns = [
        r"\bmade\s+in\s+([a-z][a-z\s\-]*)",
        r"\bfabrique\s+en\s+([a-z][a-z\s\-]*)",
        r"\bhecho\s+en\s+([a-z][a-z\s\-]*)",
    ]

    for pattern in patterns:
        match = re.search(pattern, normalized)
        if not match:
            continue
        full = match.group(0).strip()
        full = re.split(
            r"\b(?:rn|ca|sku|size|color|colour|wash|machine)\b",
            full,
            maxsplit=1
        )[0].strip()
        return full

    return None


def coo_language(text):
    normalized = normalize_text(text)
    if "made in" in normalized:
        return "EN"
    if "fabrique en" in normalized:
        return "FR"
    if "hecho en" in normalized:
        return "SP"
    return ""


def extract_rn_ca_components(text):
    """Extract RN and CA numbers independently from a combined artwork line."""
    normalized = normalize_text(text)
    if not normalized:
        return {"rn": None, "ca": None}

    result = {"rn": None, "ca": None}

    rn_match = re.search(
        r"\brn\s*[#:.-]?\s*([0-9][0-9a-z\-/]*)",
        normalized,
        re.IGNORECASE
    )
    if rn_match:
        result["rn"] = rn_match.group(1)

    ca_match = re.search(
        r"\bca\s*[#:.-]?\s*([0-9][0-9a-z\-/]*)",
        normalized,
        re.IGNORECASE
    )
    if ca_match:
        result["ca"] = ca_match.group(1)

    return result


def extract_rn_value(text):
    """Backward-compatible RN extractor; return RN first, then CA only if no RN exists."""
    parts = extract_rn_ca_components(text)
    return parts.get("rn") or parts.get("ca")


def extract_identifier_value(text):
    normalized = normalize_text(text)
    patterns = [
        r"\bsku\s*[:#-]?\s*([a-z0-9][a-z0-9_\-/]*)",
        r"\bitem\s*(?:code|no|number)\s*[:#-]?\s*([a-z0-9][a-z0-9_\-/]*)",
        r"\bstyle\s*(?:code|no|number)?\s*[:#-]?\s*([a-z0-9][a-z0-9_\-/]*)",
    ]
    for pattern in patterns:
        match = re.search(pattern, normalized, re.IGNORECASE)
        if match:
            return match.group(1)
    return None


def extract_size_value(text):
    normalized = normalize_text(text)
    match = re.search(
        r"\bsize\s*[:#-]?\s*([a-z0-9][a-z0-9\s\-/]*)",
        normalized,
        re.IGNORECASE
    )
    if not match:
        return None
    value = re.split(
        r"\b(?:rn|ca|made|color|colour|sku|style)\b",
        match.group(1),
        maxsplit=1
    )[0].strip()
    return value or None


def extract_color_value(text):
    normalized = normalize_text(text)
    match = re.search(
        r"\b(?:color|colour)\s*[:#-]?\s*([a-z][a-z\s\-/]*)",
        normalized,
        re.IGNORECASE
    )
    if not match:
        return None
    value = re.split(
        r"\b(?:size|rn|ca|made|country|sku|style)\b",
        match.group(1),
        maxsplit=1
    )[0].strip()
    return value or None


def extract_gender_value(text):
    normalized = normalize_text(text)
    for value in [
        "boys", "girls", "women", "men", "unisex",
        "boy", "girl", "woman", "man"
    ]:
        if re.search(rf"\b{re.escape(value)}\b", normalized):
            return value
    return None


def _content_material_key(material):
    """Normalize a material name for comparison while preserving semantic words."""
    value = _auto_strip_accents(material) if "_auto_strip_accents" in globals() else unicodedata.normalize("NFKD", str(material or "")).casefold()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def extract_content_values(text):
    """Extract percentage/material pairs while preserving source order.

    Works with: 60% cotton, 60cotton, POLY-ESTER, SPANDEX/ELASTANE, and
    component labels such as SHELL/LINER/DOUBLURE/FORRO without treating those
    labels as part of the material.
    """
    normalized = normalize_text(text)
    if not normalized:
        return []

    stop_labels = {
        "shell", "liner", "lining", "body", "cuerpo", "forro",
        "exterior", "extrieur", "exterieur", "doublure",
        "made", "hecho", "fabrique", "rn", "ca", "mx",
        "cr", "ec", "gt", "pa", "sv", "us", "size", "color", "colour"
    }

    values = []

    # Standard percentage forms. For each percentage, take the first material
    # phrase immediately following it, stopping at the next component label or
    # next percentage.
    pct_matches = list(re.finditer(r"(?<![a-z0-9])(\d{1,3}(?:\.\d+)?)\s*%?", normalized))
    for index, match in enumerate(pct_matches):
        pct = match.group(1)
        tail_start = match.end()
        tail_end = pct_matches[index + 1].start() if index + 1 < len(pct_matches) else len(normalized)
        tail = normalized[tail_start:tail_end].strip()
        if not tail:
            continue

        # Remove common separators and stop when a structural/component label
        # begins.
        tail = re.sub(r"^[^a-z0-9]+", "", tail)
        tokens = re.findall(r"[a-z][a-z0-9]*(?:[-/][a-z][a-z0-9]*)?", tail, flags=re.IGNORECASE)
        material_tokens = []
        for token in tokens[:6]:
            compact = re.sub(r"[^a-z0-9]", "", token.casefold())
            if not compact:
                continue
            if compact in stop_labels:
                break
            material_tokens.append(token)

        # Keep normal textile materials concise while allowing combinations such
        # as SPANDEX/ELASTANE. Component labels stop the capture.
        if not material_tokens:
            continue

        material = " ".join(material_tokens[:2]).strip()
        if material:
            values.append(f"{pct}% {material}")

    # OCR sometimes collapses 60% cotton to 60cotton. Add only attached-number
    # matches that were not already captured.
    for match in re.finditer(r"(?<![a-z0-9])(\d{1,3})(?=[a-z])([a-z][a-z0-9]*(?:[-/][a-z][a-z0-9]*)?)", normalized):
        pct = match.group(1)
        material = match.group(2)
        candidate = f"{pct}% {material}"
        if not any(_content_material_key(candidate) == _content_material_key(existing) for existing in values):
            values.append(candidate)

    return values


def normalize_composition(values):
    # Preserve source order. Ordering is presentation/semantic information;
    # sorting it hides which exact component differs.
    return [
        normalize_text(value)
        for value in values
        if normalize_text(value)
    ]


def analyze_content_difference(expected, actual):
    """Return the actual Content defect in QC-friendly language."""
    expected_values = extract_content_values(expected)
    actual_values = extract_content_values(actual)

    if not expected_values or not actual_values:
        return f"Expected: {expected} | Found: {actual}"

    issues = []
    from difflib import SequenceMatcher

    max_len = max(len(expected_values), len(actual_values))

    for index in range(max_len):
        exp = expected_values[index] if index < len(expected_values) else None
        act = actual_values[index] if index < len(actual_values) else None

        if exp is None:
            issues.append(f"Extra content in PDF: {act}")
            continue

        if act is None:
            issues.append(f"Missing from PDF: {exp}")
            continue

        exp_match = re.fullmatch(
            r"(?P<pct>\d+(?:\.\d+)?)%\s*(?P<material>.+)",
            normalize_text(exp),
            re.IGNORECASE
        )
        act_match = re.fullmatch(
            r"(?P<pct>\d+(?:\.\d+)?)%\s*(?P<material>.+)",
            normalize_text(act),
            re.IGNORECASE
        )

        if not exp_match or not act_match:
            if normalize_text(exp) != normalize_text(act):
                issues.append(f"Content mismatch: {exp} → {act}")
            continue

        exp_pct = exp_match.group("pct")
        act_pct = act_match.group("pct")
        exp_material = exp_match.group("material").strip()
        act_material = act_match.group("material").strip()

        if exp_pct != act_pct:
            issues.append(
                f"Percentage mismatch: Order Form {exp_pct}% → PDF {act_pct}% ({exp_material.upper()})"
            )
            continue

        if normalize_text(exp_material) == normalize_text(act_material):
            continue

        similarity = SequenceMatcher(
            None,
            normalize_text(exp_material),
            normalize_text(act_material),
            autojunk=False
        ).ratio()

        if similarity >= 0.70:
            issues.append(
                f'Spelling mistake: Order Form "{exp_material.upper()}" → PDF "{act_material.upper()}"'
            )
        else:
            issues.append(
                f'Material mismatch: Order Form "{exp_material.upper()}" → PDF "{act_material.upper()}"'
            )

    return "; ".join(issues) if issues else "Content differs."


def composition_matches(expected, actual):
    expected_values = normalize_composition(
        extract_content_values(expected)
    )
    actual_values = normalize_composition(
        extract_content_values(actual)
    )
    return bool(expected_values and expected_values == actual_values)


# =========================================================
# FIELD-SPECIFIC REGIONS
# =========================================================

FIELD_ANCHORS = {
    "COO": ["made in", "fabrique en", "hecho en"],
    "CARE": ["machine wash", "wash", "laver", "laver", "lavar", "bleach", "dry clean"],
    "CONTENT": ["%", "shell", "liner", "lining", "body", "fiber", "fibre", "content", "composition"],
    "RN": ["rn", "ca"],
    "IDENTIFIER": ["sku", "item code", "item no", "item number", "style", "style code"],
    "SIZE": ["size"],
    "COLOR": ["color", "colour"],
    "GENDER": ["boys", "girls", "women", "men", "unisex"],
    "BATCH": ["batch", "lot"],
    "QUANTITY": ["quantity", "qty", "units", "pcs"],
    "BRAND": ["brand"],
    "ATTRIBUTE": ["attribute", "technology", "feature"],
    "SYMBOL": [],
    "OSZ": [],
    "GENERAL": [],
}


def line_relevant_for_field(line_text, field_name):
    field_type = get_field_type(field_name)
    normalized = normalize_text(line_text)

    if not normalized:
        return False

    if field_type == "SYMBOL":
        return False

    if field_type == "OSZ":
        # OSZ values are resolved from a sequence, not generic relevance.
        return False

    for anchor in FIELD_ANCHORS.get(field_type, []):
        anchor_norm = normalize_text(anchor)
        if anchor_norm and anchor_norm in normalized:
            return True

    region = get_field_region(field_name)

    region_markers = {
        "EN": ["made in", "machine wash", "shell", "liner", "polyester", "bleach"],
        "FR": ["fabrique en", "laver", "extérieur", "doublure", "polyester", "sans chlore"],
        "SP": ["hecho en", "lavar", "forro", "poliester", "cloro"],
    }

    for marker in region_markers.get(region, []):
        if normalize_text(marker) in normalized:
            return True

    return False


def collect_region_from_anchor(state, field_name):
    """Collect a meaningful field region without crossing another major field."""
    lines = state["lines"]
    field_type = get_field_type(field_name)

    start_idx = None

    for idx, line in enumerate(lines):
        if not line_is_available(line, state):
            continue
        if line_relevant_for_field(line["text"], field_name):
            start_idx = idx
            break

    if start_idx is None:
        return None

    region = []

    stop_types = {
        "COO": ("RN", "CONTENT", "CARE"),
        "CONTENT": ("COO", "RN", "CARE"),
        "CARE": ("COO", "CONTENT", "RN"),
    }

    stop_patterns = {
        "COO": [r"\b(?:rn|ca)\s*[#:.-]?\s*\w+"],
        "CONTENT": [r"\b(?:made in|rn|ca)\b"],
        "CARE": [r"\b(?:made in|rn|ca)\b"],
    }

    for idx in range(start_idx, len(lines)):
        line = lines[idx]

        if not line_is_available(line, state):
            break

        text = line["text"]
        norm = line["norm"]

        if idx > start_idx:
            if field_type in stop_types:
                if any(
                    any(marker in norm for marker in FIELD_ANCHORS.get(stop_type, []))
                    for stop_type in stop_types[field_type]
                ):
                    break

            if field_type in stop_patterns:
                if any(re.search(pattern, norm) for pattern in stop_patterns[field_type]):
                    break

        region.append(line)

        # Avoid swallowing an entire page.
        if len(region) >= 20:
            break

    return region or None


# =========================================================
# OSZ SEQUENCE DETECTION
# =========================================================

def get_osz_index(field_name):
    compact = normalize_text(field_name).replace(" ", "")
    match = re.fullmatch(r"osz(\d+)", compact)
    return int(match.group(1)) if match else None


def extract_standalone_numeric_runs(state):
    runs = []
    current = []

    for line in state["lines"]:
        if not line_is_available(line, state):
            if current:
                runs.append(current)
                current = []
            continue

        norm = normalize_text(line["text"])

        # Standalone integer only. This deliberately excludes RN# 55285,
        # dates, dimensions, item codes and mixed text.
        if re.fullmatch(r"[-+]?\d+", norm):
            current.append(line)
        else:
            if current:
                runs.append(current)
                current = []

    if current:
        runs.append(current)

    return [run for run in runs if run]


def find_osz_value(
    expected,
    field_name,
    state,
    osz_group_size=1
):
    expected_num = normalize_numeric(expected)
    index = get_osz_index(field_name)

    if expected_num is None or index is None:
        return None

    # Explicit OSZ label, when available.
    label_pattern = re.compile(
        rf"\bosz\s*{index}\s*[:#=-]?\s*(\d+)\b",
        re.IGNORECASE
    )

    for line in state["lines"]:
        if not line_is_available(line, state):
            continue
        match = label_pattern.search(line["text"])
        if not match:
            continue

        actual = normalize_numeric(match.group(1))
        consume_lines(state, [line])
        if actual == expected_num:
            return {
                "status": "PASS",
                "pdf": line["text"],
                "difference": "—",
                "match_type": "OSZ_LABEL"
            }
        return {
            "status": "FAIL",
            "pdf": line["text"],
            "difference": f"Expected: {expected} | Found: {match.group(1)}",
            "match_type": "OSZ_LABEL_MISMATCH"
        }

    # Coordinate-aware OSZ mapping is preferred because OCR can merge a
    # standalone number with an unrelated neighbouring line (for example
    # turning a clean sequence 6 / 8 into a line such as "0 8").
    candidates = [
        candidate
        for candidate in state.get("osz_sequence_candidates", [])
        if len(candidate.get("items", [])) >= index
    ]

    if candidates:
        # First choose a candidate whose index already agrees with the Excel
        # value. This prevents an unrelated numeric run elsewhere on the page
        # from being selected merely because it is long.
        matching_candidates = [
            candidate
            for candidate in candidates
            if _osz_candidate_matches_expected(candidate, index, expected_num)
        ]
        chosen = max(
            matching_candidates or candidates,
            key=lambda candidate: candidate.get("score", 0.0)
        )

        item = chosen["items"][index - 1]
        actual = normalize_numeric(item.get("value"))
        actual_text = str(item.get("value", "")).strip() or "Not found"

        if actual == expected_num:
            return {
                "status": "PASS",
                "pdf": actual_text,
                "difference": "—",
                "match_type": "OSZ_SEQUENCE_GEOMETRY"
            }

        if actual is not None:
            return {
                "status": "FAIL",
                "pdf": actual_text,
                "difference": f"Expected: {expected} | Found: {actual_text}",
                "match_type": "OSZ_SEQUENCE_GEOMETRY_MISMATCH"
            }

    # Original line-based sequence logic remains as the final fallback.
    # This preserves the behaviour for PDFs whose text layer already exposes
    # one clean standalone number per line.
    runs = [run for run in state.get("osz_runs", []) if len(run) >= index]
    if not runs:
        return None

    runs.sort(key=lambda run: (-len(run), run[0]["line_id"]))
    run = runs[0]
    line = run[index - 1]

    if not line_is_available(line, state):
        return None

    actual = normalize_numeric(line["text"])

    if actual == expected_num:
        consume_lines(state, [line])
        return {
            "status": "PASS",
            "pdf": line["text"],
            "difference": "—",
            "match_type": "OSZ_SEQUENCE"
        }

    if actual is not None:
        consume_lines(state, [line])
        return {
            "status": "FAIL",
            "pdf": line["text"],
            "difference": f"Expected: {expected} | Found: {line['text']}",
            "match_type": "OSZ_SEQUENCE_MISMATCH"
        }

    return None


# =========================================================
# DIFFERENCE DESCRIPTION
# =========================================================

def describe_text_difference(expected, actual):
    """Explanation only; never used as a PASS/FAIL decision."""
    expected_tokens = tokenize(expected)
    actual_tokens = tokenize(actual)

    if not expected_tokens:
        return "Expected value is blank."
    if not actual_tokens:
        return "Expected value is missing from the output."

    from difflib import SequenceMatcher

    matcher = SequenceMatcher(
        None,
        expected_tokens,
        actual_tokens,
        autojunk=False
    )

    differences = []

    for tag, a1, a2, b1, b2 in matcher.get_opcodes():
        if tag == "equal":
            continue

        left = " ".join(expected_tokens[a1:a2]).strip()
        right = " ".join(actual_tokens[b1:b2]).strip()

        if tag == "replace":
            differences.append(f"{left} → {right}")
        elif tag == "delete":
            differences.append(f"Missing: {left}")
        elif tag == "insert":
            differences.append(f"Extra: {right}")

    return "; ".join(differences[:8]) or "Content differs."


def _field_casefold_compact(text):
    return re.sub(r"[^a-z0-9]", "", _auto_strip_accents(text))


def _care_start_markers(region):
    return {
        "EN": ["machine wash", "wash cold"],
        "FR": ["laver a la machine", "laver a machine", "laver à la machine", "laver"],
        "SP": ["lavar a maquina", "lavar a máquina", "lavar a maquina con", "lavar"],
        "": ["machine wash", "laver a la machine", "lavar a maquina", "lavar"],
    }.get(region, ["machine wash", "laver a la machine", "lavar a maquina"])


def _care_region_stop_markers(region):
    common = ["made in", "fabrique en", "hecho en", "rn", "ca :", "mx :", "actual other sizes"]
    if region == "EN":
        return common + ["@ ca", "ca :"]
    if region == "FR":
        return common + ["@ cr/ec", "mx/pa/sv :", "lavar a maquina"]
    if region == "SP":
        return common + ["actual other sizes", "rn"]
    return common


def find_care_region(state, field_name):
    """Find the language-specific care block using strong region starts."""
    region = get_field_region(field_name)
    lines = state["lines"]
    start_patterns = {
        "EN": ["machine wash"],
        "FR": ["laver à la machine", "laver a la machine"],
        "SP": ["lavar a maquina", "lavar a máquina"],
        "": ["machine wash", "laver à la machine", "laver a la machine", "lavar a maquina", "lavar a máquina"],
    }
    starts = [normalize_text(x) for x in start_patterns.get(region, start_patterns[""]) if normalize_text(x)]

    start = None
    for idx, line in enumerate(lines):
        if not line_is_available(line, state):
            continue
        norm = normalize_text(line["text"])
        if any(marker in norm for marker in starts):
            start = idx
            break

    if start is None:
        return None

    region_lines = []
    for idx in range(start, len(lines)):
        line = lines[idx]
        if not line_is_available(line, state):
            break
        norm = normalize_text(line["text"])

        if idx > start:
            # Any explicit COO/content/RN marker begins another artwork region.
            if extract_coo_value(line["text"]) or extract_rn_ca_components(line["text"]).get("rn"):
                break
            if extract_content_values(line["text"]):
                # A composition region after the care text belongs elsewhere.
                break
            if region == "EN" and ("laver à la machine" in norm or "laver a la machine" in norm or "lavar a maquina" in norm):
                break
            if region == "FR" and ("lavar a maquina" in norm or "lavar a máquina" in norm):
                break
            if region == "SP" and ("machine wash" in norm or "laver à la machine" in norm or "laver a la machine" in norm):
                break
            if "actual other sizes" in norm:
                break
            if len(re.findall(r"[A-Za-zÀ-ÿ]", line["text"])) == 0:
                break

        region_lines.append(line)
        if len(region_lines) >= 20:
            break

    return region_lines or None


def _content_region_markers(field_name):
    compact = normalize_text(field_name).replace(" ", "").replace("_", "").replace("-", "")
    if "fib_en" in compact or compact.endswith("en"):
        return ["us : shell", "us shell", "shell:", "shell"]
    if "fib_ca" in compact or compact.endswith("ca"):
        return ["ca : extérieur", "ca : exterieur", "ca extérieur", "ca exterieur", "exterieur", "extérieur"]
    if "fibspmexico" in compact:
        return ["cr/ec/gt/pa/sv : cuerpo", "cr/ec/gt/pa/sv :", "cr/ec/gt/pa/sv"]
    if "fibmexico" in compact:
        return ["mx : cuerpo", "mx cuerpo", "mx :"]
    if "fibsp" in compact or compact.endswith("sp"):
        return ["mx : cuerpo", "mx cuerpo", "cr/ec/gt/pa/sv : cuerpo", "cuerpo"]
    return ["shell", "exterieur", "cuerpo", "content"]


def _content_region_stop_markers(field_name):
    return [
        "ca : exterieur", "ca : extérieur", "mx :", "cr/ec/gt/pa/sv :",
        "machine wash", "laver a la machine", "laver à la machine",
        "made in", "hecho en", "fabrique en", "rn ", "actual other sizes"
    ]


def find_content_region(state, field_name):
    """Locate a complete composition block without crossing into another region."""
    lines = state["lines"]
    starts = [normalize_text(x) for x in _content_region_markers(field_name) if normalize_text(x)]
    stops = [normalize_text(x) for x in _content_region_stop_markers(field_name) if normalize_text(x)]

    start = None
    for idx, line in enumerate(lines):
        if not line_is_available(line, state):
            continue
        norm = normalize_text(line["text"])
        if any(marker in norm for marker in starts):
            start = idx
            break
    if start is None and not normalize_text(field_name).replace(" ", "").replace("_", "").replace("-", "").startswith("fib"):
        # Generic CONTENT fields (e.g. 'Content') may not have a schema-specific
        # marker. Anchor on the first composition-shaped line.
        for idx, line in enumerate(lines):
            if not line_is_available(line, state):
                continue
            if re.search(r"\d{1,3}\s*%", line.get("text", "")):
                start = idx
                break

    if start is None:
        return None

    selected = []
    for idx in range(start, len(lines)):
        line = lines[idx]
        if not line_is_available(line, state):
            break

        norm = normalize_text(line["text"])
        is_stop = idx > start and any(marker in norm for marker in stops)

        if is_stop:
            current_text = join_lines(selected) if selected else ""
            current_values = extract_content_values(current_text)
            compact_field = normalize_text(field_name).replace(" ", "").replace("_", "").replace("-", "")
            has_complete_tail = any(re.match(r"100\s*%?\s+.+", normalize_text(v)) for v in current_values)
            if not compact_field.startswith("fib") and current_values:
                break
            # If the current block already contains a complete 100% component,
            # do not absorb the next region. Otherwise include one boundary line
            # because OCR/PDF extraction can split the material name across it.
            if has_complete_tail:
                break
            selected.append(line)
            break

        selected.append(line)

        joined = join_lines(selected)
        current_values = extract_content_values(joined)
        has_complete_100 = any(
            re.match(r"100\s*%?\s+.+", normalize_text(v))
            for v in current_values
        )
        if has_complete_100 and idx > start and not re.search(r"\d{1,3}\s*%", line["text"]):
            break

        if len(selected) >= 8:
            break

    return selected or None



def _composition_signature(values):
    signature = []
    for value in values:
        match = re.match(r"(\d+(?:\.\d+)?)%?\s*(.+)", normalize_text(value))
        if not match:
            signature.append(("", _content_material_key(value)))
            continue
        pct = match.group(1)
        material = _content_material_key(match.group(2))
        signature.append((pct, material))
    return signature


def _describe_composition_difference(expected, actual):
    exp = _composition_signature(extract_content_values(expected))
    act = _composition_signature(extract_content_values(actual))
    if exp == act:
        return "—"
    return f"Expected: {expected} | Found: {actual}"


def _size_block_score(expected, actual):
    from difflib import SequenceMatcher
    exp = normalize_text(expected)
    act = normalize_text(actual)
    if not exp or not act:
        return 0.0
    if exp == act or _field_casefold_compact(exp) == _field_casefold_compact(act):
        return 1.0
    token_ratio = SequenceMatcher(None, exp.split(), act.split(), autojunk=False).ratio()
    compact_ratio = SequenceMatcher(None, _field_casefold_compact(exp), _field_casefold_compact(act), autojunk=False).ratio()
    return max(token_ratio, compact_ratio * 0.98)


def find_size_block_match(expected, field_name, state):
    """Find a multi-line size block, returning PASS or a best-evidence FAIL."""
    expected = str(expected or "").strip()
    if not expected:
        return None

    available = [line for line in state["lines"] if line_is_available(line, state)]
    if not available:
        return None

    expected_line_count = max(1, len([x for x in expected.splitlines() if x.strip()]))
    min_window = max(1, expected_line_count - 1)
    max_window = min(8, expected_line_count + 2)

    exact_candidates = []
    scored = []
    for size in range(min_window, max_window + 1):
        for start in range(0, len(available) - size + 1):
            candidate = available[start:start + size]
            ids = [line["line_id"] for line in candidate]
            if ids != list(range(ids[0], ids[-1] + 1)):
                continue
            text = join_lines(candidate)
            score = _size_block_score(expected, text)
            scored.append((score, candidate, text))
            if score >= 0.999:
                exact_candidates.append((candidate, text))

    if exact_candidates:
        # If the same exact size block exists twice, consume the first available
        # occurrence; subsequent duplicate fields can use the next occurrence.
        candidate, text = exact_candidates[0]
        return {
            "status": "PASS",
            "lines": candidate,
            "pdf": text,
            "difference": "—",
            "match_type": "SIZE_BLOCK_EXACT"
        }

    if not scored:
        return None

    # Favor candidates that contain a size-like token from the expected block.
    expected_markers = [
        token.casefold() for token in re.findall(r"[A-Za-z]{1,4}", expected)
        if token.casefold() not in {"size"}
    ]
    ranked = []
    for score, candidate, text in scored:
        norm = normalize_text(text)
        marker_hits = sum(1 for token in expected_markers if normalize_text(token) in norm)
        digit_hits = len(set(re.findall(r"\d+", normalize_text(expected))) & set(re.findall(r"\d+", norm)))
        ranked.append((score + marker_hits * 0.03 + digit_hits * 0.015, score, candidate, text))

    ranked.sort(key=lambda item: item[0], reverse=True)
    best = ranked[0]
    if best[1] >= 0.55 and (len(expected_markers) <= 1 or best[0] >= 0.62):
        return {
            "status": "FAIL",
            "lines": best[2],
            "pdf": best[3],
            "difference": f"Expected: {expected} | Found: {best[3]}",
            "match_type": "SIZE_BLOCK_MISMATCH"
        }

    return None


# =========================================================
# FIELD CHECKERS
# =========================================================




# =========================================================
# BARCODE / UPC / EAN / GTIN SUPPORT
# =========================================================

BARCODE_CONFUSION_MAP = str.maketrans({
    # Common OCR confusions in human-readable barcode digits.
    "O": "0", "Q": "0", "D": "0",
    "I": "1", "L": "1", "T": "1",
    "Z": "2",
    "S": "5",
    "G": "6",
    "B": "8",
})


def normalize_barcode_digits(value, allow_ocr_confusions=False):
    """Return the digit sequence represented by a barcode value."""
    if is_blank_value(value):
        return ""

    raw = str(value).strip().upper()
    if allow_ocr_confusions:
        raw = raw.translate(BARCODE_CONFUSION_MAP)
    return re.sub(r"[^0-9]", "", raw)


def barcode_family(digits):
    """Return the common retail barcode family from digit length."""
    length = len(str(digits or ""))
    return {
        8: "EAN-8",
        12: "UPC-A",
        13: "EAN-13",
        14: "GTIN-14",
    }.get(length, "BARCODE")


def barcode_check_digit_valid(digits):
    """Validate UPC/EAN/GTIN check digit when the length is supported."""
    digits = normalize_barcode_digits(digits)
    if len(digits) not in {8, 12, 13, 14}:
        return False

    body = digits[:-1]
    provided = int(digits[-1])
    total = 0

    # GTIN checksum: weighting depends on parity from the right.
    # Starting from the rightmost body digit, weights alternate 3, 1, 3, 1...
    reversed_body = list(reversed(body))
    for index, char in enumerate(reversed_body):
        weight = 3 if index % 2 == 0 else 1
        total += int(char) * weight

    calculated = (10 - (total % 10)) % 10
    return calculated == provided


def barcode_gtin14(digits):
    """Canonical GTIN-14 representation for 8/12/13/14-digit GTINs."""
    digits = normalize_barcode_digits(digits)
    if len(digits) in {8, 12, 13, 14}:
        return digits.zfill(14)
    return ""


def barcode_equivalent(expected_digits, actual_digits):
    """Exact barcode identity with UPC-A/EAN/GTIN zero-padding equivalence."""
    expected_digits = normalize_barcode_digits(expected_digits)
    actual_digits = normalize_barcode_digits(actual_digits)

    if not expected_digits or not actual_digits:
        return False

    if expected_digits == actual_digits:
        return True

    # UPC-A is the 12-digit form of the same GTIN represented as EAN-13 with a
    # leading zero. GTIN-14 normalization also safely handles leading-zero forms
    # of EAN-8/UPC-A/EAN-13.
    expected_gtin14 = barcode_gtin14(expected_digits)
    actual_gtin14 = barcode_gtin14(actual_digits)
    return bool(expected_gtin14 and actual_gtin14 and expected_gtin14 == actual_gtin14)


def _barcode_chunks_from_line(line_text):
    """Extract barcode-like numeric strings while preserving grouping."""
    text = str(line_text or "")
    candidates = []

    # Plain numeric groups separated by spaces/hyphens/dots/slashes/parentheses.
    # The total normalized length is checked later against the expected barcode.
    for match in re.finditer(
        r"(?<![A-Za-z0-9])[0-9][0-9\s.\-_/()]*[0-9](?![A-Za-z0-9])",
        text,
    ):
        raw = match.group(0)
        digits = normalize_barcode_digits(raw)
        if 8 <= len(digits) <= 14:
            candidates.append({
                "digits": digits,
                "raw": raw.strip(),
                "start": match.start(),
                "end": match.end(),
                "source": "digits",
            })

    # OCR-tolerant version: allow only a tightly controlled set of common
    # digit/letter confusions, and only when the final normalized length is
    # barcode-like. This is never used to create a fuzzy PASS against arbitrary
    # text; it is simply a second candidate representation.
    for match in re.finditer(
        r"(?<![A-Za-z0-9])[0-9OQDILTZSBG][0-9OQDILTZSBG\s.\-_/()]*[0-9OQDILTZSBG](?![A-Za-z0-9])",
        text.upper(),
    ):
        raw = match.group(0)
        digits = normalize_barcode_digits(raw, allow_ocr_confusions=True)
        if 8 <= len(digits) <= 14 and any(ch.isalpha() for ch in raw):
            candidates.append({
                "digits": digits,
                "raw": raw.strip(),
                "start": match.start(),
                "end": match.end(),
                "source": "ocr_confusion",
            })

    # Deduplicate identical digit candidates from the two extraction paths.
    unique = []
    seen = set()
    for item in candidates:
        key = (item["digits"], item["start"], item["end"])
        if key in seen:
            continue
        seen.add(key)
        unique.append(item)
    return unique


def _barcode_candidates_from_state(state):
    """Collect barcode-like candidates from comparison text and raw PDF text."""
    candidates = []
    seen = set()

    pools = []
    for line in state.get("lines", []):
        pools.append(("comparison", line.get("text", "")))
    for line in state.get("direct_lines", []):
        pools.append(("direct_pdf", line.get("text", "")))

    for source_kind, text in pools:
        for item in _barcode_chunks_from_line(text):
            key = (item["digits"], item["raw"], source_kind)
            if key in seen:
                continue
            seen.add(key)
            copied = dict(item)
            copied["source_kind"] = source_kind
            copied["line_text"] = str(text or "").strip()
            candidates.append(copied)

    return candidates


def find_barcode_match(expected, state):
    """Robust retail barcode matching without fuzzy/partial acceptance."""
    expected_digits = normalize_barcode_digits(expected)
    if len(expected_digits) not in {8, 12, 13, 14}:
        return None

    candidates = _barcode_candidates_from_state(state)
    if not candidates:
        return None

    # First, exact / GTIN-equivalent match. Prefer exact source text over OCR
    # confusion candidates when both represent the same number.
    exact = [
        candidate for candidate in candidates
        if barcode_equivalent(expected_digits, candidate["digits"])
    ]
    if exact:
        exact.sort(
            key=lambda c: (
                c.get("source") == "ocr_confusion",
                c.get("source_kind") != "direct_pdf",
                len(c.get("raw", "")),
            )
        )
        best = exact[0]
        return {
            "status": "PASS",
            "pdf": best["raw"],
            "difference": "—",
            "match_type": f"BARCODE_{barcode_family(expected_digits)}",
            "barcode_digits": best["digits"],
        }

    # If no exact match exists, look for a strong same-family candidate in a
    # line that is explicitly barcode-labelled. This permits a genuine FAIL to
    # be shown instead of falsely reporting NOT FOUND, while avoiding random
    # page numbers/OSZ values becoming barcode failures.
    labelled_candidates = []
    for candidate in candidates:
        line = normalize_text(candidate.get("line_text", ""))
        if any(marker in line for marker in (
            "barcode", "upc", "ean", "gtin", "jan", "isbn", "itf"
        )):
            if len(candidate["digits"]) in {8, 12, 13, 14}:
                labelled_candidates.append(candidate)

    if labelled_candidates:
        # Prefer same length; otherwise prefer GTIN-compatible retail families.
        same_length = [
            candidate for candidate in labelled_candidates
            if len(candidate["digits"]) == len(expected_digits)
        ]
        pool = same_length or labelled_candidates
        pool.sort(
            key=lambda c: (
                abs(len(c["digits"]) - len(expected_digits)),
                c.get("source") == "ocr_confusion",
            )
        )
        best = pool[0]
        return {
            "status": "FAIL",
            "pdf": best["raw"],
            "difference": (
                f"Barcode mismatch: Order Form has '{normalize_barcode_digits(expected)}' "
                f"but PDF has '{best['digits']}'."
            ),
            "match_type": "BARCODE_MISMATCH",
            "barcode_digits": best["digits"],
        }

    # Otherwise do not manufacture a FAIL merely because the page contains
    # numbers. Return NOT FOUND until barcode-specific evidence exists.
    return None


def check_field(
    expected,
    field_name,
    state,
    osz_group_size=1
):
    """Deterministic field validation. PASS is always attempted first."""

    if is_blank_value(expected):
        return {
            "status": "SKIP",
            "pdf": "—",
            "difference": "Blank Order Form value — field ignored.",
            "match_type": "BLANK"
        }

    expected = str(expected).strip()
    field_type = get_field_type(field_name)

    # -----------------------------------------------------
    # BARCODE / UPC / EAN / GTIN
    # -----------------------------------------------------
    if field_type == "BARCODE":
        barcode_result = find_barcode_match(expected, state)
        if barcode_result:
            # BARCODE matching is independent of the normal line-consumption
            # rules. We consume a matching line only when we have a concrete
            # barcode candidate, while leaving other numbers on the page alone.
            target_digits = barcode_result.get("barcode_digits", "")
            for line in state["lines"]:
                if not line_is_available(line, state):
                    continue
                if normalize_barcode_digits(line.get("text", "")) == target_digits:
                    consume_lines(state, [line])
                    break
            return barcode_result

        return {
            "status": "NOT FOUND",
            "pdf": "Not found",
            "difference": "Barcode/UPC/EAN value was not detected reliably.",
            "match_type": "BARCODE_NOT_FOUND",
        }

    # -----------------------------------------------------
    # OSZ sequence
    # -----------------------------------------------------
    if field_type == "OSZ":
        result = find_osz_value(
            expected,
            field_name,
            state,
            osz_group_size=osz_group_size
        )
        if result:
            return result

        return {
            "status": "NOT FOUND",
            "pdf": "Not found",
            "difference": "OSZ sequence/value was not detected.",
            "match_type": "NOT_FOUND"
        }

    # -----------------------------------------------------
    # Symbol: exact custom-font keystroke match.
    # -----------------------------------------------------
    if field_type == "SYMBOL":
        expected_symbol = normalize_symbol_text(expected)

        if not expected_symbol:
            return {
                "status": "NOT FOUND",
                "pdf": "Not found",
                "difference": "Symbol/value was not detected exactly.",
                "match_type": "NOT_FOUND"
            }

        # PASS 1: current OCR/comparison lines. The expected symbol can share
        # a line with other content, so do not require the whole line to equal
        # the symbol. _symbol_text_matches still enforces exact code-point/token
        # matching and never uses fuzzy similarity.
        for line in state["lines"]:
            if not line_is_available(line, state):
                continue

            if _symbol_text_matches(expected_symbol, line["text"]):
                actual_raw = normalize_symbol_text(line["text"])
                # Consume only the line/span used by the symbol check. This keeps
                # other fields on the same OCR line independently usable.
                if expected_symbol in actual_raw:
                    start_pos = actual_raw.find(expected_symbol)
                    consume_span(
                        state,
                        line,
                        start_pos,
                        start_pos + len(expected_symbol)
                    )
                else:
                    consume_lines(state, [line])

                return {
                    "status": "PASS",
                    "pdf": line["text"],
                    "difference": "—",
                    "match_type": "SYMBOL_KEYSTROKE_EXACT"
                }

        # PASS 2: raw PDF text layer. This is critical for custom fonts because
        # OCR may replace a private-use glyph even though the original PDF text
        # contains the exact keystroke. This fallback is isolated to SYMBOL.
        for line in state.get("direct_lines", []):
            if not line_is_available(line, state):
                continue

            if _symbol_text_matches(expected_symbol, line["text"]):
                actual_raw = normalize_symbol_text(line["text"])
                if expected_symbol in actual_raw:
                    start_pos = actual_raw.find(expected_symbol)
                    consume_span(
                        state,
                        line,
                        start_pos,
                        start_pos + len(expected_symbol)
                    )
                else:
                    consume_lines(state, [line])

                return {
                    "status": "PASS",
                    "pdf": line["text"],
                    "difference": "—",
                    "match_type": "SYMBOL_DIRECT_PDF_EXACT"
                }

        return {
            "status": "NOT FOUND",
            "pdf": "Not found",
            "difference": "Symbol/value was not detected exactly.",
            "match_type": "NOT_FOUND"
        }

    # -----------------------------------------------------
    # CONTENT: structured composition region
    # -----------------------------------------------------
    if field_type == "CONTENT":
        expected_values = extract_content_values(expected)
        region = find_content_region(state, field_name)
        if region:
            actual_text = join_lines(region)
            actual_values = extract_content_values(actual_text)
            if expected_values and actual_values:
                if _composition_signature(expected_values) == _composition_signature(actual_values):
                    return {
                        "status": "PASS",
                        "pdf": actual_text,
                        "difference": "—",
                        "match_type": "CONTENT_REGION_EXACT"
                    }

                # A related composition region with meaningful components is a
                # real FAIL, not NOT FOUND.
                expected_materials = {_content_material_key(x) for x in expected_values}
                actual_materials = {_content_material_key(x) for x in actual_values}
                overlap = expected_materials & actual_materials
                if overlap or len(actual_values) >= 1:
                    return {
                        "status": "FAIL",
                        "pdf": actual_text,
                        "difference": _describe_composition_difference(expected, actual_text),
                        "match_type": "CONTENT_REGION_MISMATCH"
                    }

        # Fallback: search broader contiguous windows. This protects raster/OCR
        # documents where the component heading itself is not recognized.
        available = [line for line in state["lines"] if line_is_available(line, state)]
        max_window = min(8, len(available))
        for size in range(1, max_window + 1):
            for start_index in range(0, len(available) - size + 1):
                candidate = available[start_index:start_index + size]
                ids = [line["line_id"] for line in candidate]
                if ids != list(range(ids[0], ids[-1] + 1)):
                    continue
                actual_text = join_lines(candidate)
                actual_values = extract_content_values(actual_text)
                if actual_values and expected_values:
                    if _composition_signature(expected_values) == _composition_signature(actual_values):
                            return {
                            "status": "PASS",
                            "pdf": actual_text,
                            "difference": "—",
                            "match_type": "CONTENT_EXACT"
                        }

        return {
            "status": "NOT FOUND",
            "pdf": "Not found",
            "difference": "Expected composition was not detected in the relevant artwork region.",
            "match_type": "NOT_FOUND"
        }

    # -----------------------------------------------------
    # RN / CA / combined RNCA
    # -----------------------------------------------------
    if field_type in {"RN", "CA"}:
        compact_field = normalize_text(field_name).replace(" ", "").replace("_", "").replace("-", "")
        expected_parts = extract_rn_ca_components(expected)
        is_combined = compact_field in {"rnca", "rncanumber"} or (
            field_type == "RN" and expected_parts.get("rn") and expected_parts.get("ca")
        )

        if is_combined:
            exp_rn = expected_parts.get("rn")
            exp_ca = expected_parts.get("ca")
            if not exp_rn or not exp_ca:
                # Try to extract two numeric components directly from the value.
                nums = re.findall(r"\d{4,}", normalize_text(expected))
                if len(nums) >= 2:
                    exp_rn, exp_ca = nums[0], nums[1]

            for line in state["lines"]:
                if not line_is_available(line, state):
                    continue
                parts = extract_rn_ca_components(line["text"])
                if not parts.get("rn") and not parts.get("ca"):
                    continue
                rn_match = exp_rn and parts.get("rn") == str(exp_rn)
                ca_match = exp_ca and parts.get("ca") == str(exp_ca)
                if rn_match and ca_match:
                    # Consume only the two numeric spans so MIN/COO and other
                    # content on the same line remain independently available.
                    norm_line = line["norm"]
                    spans = []
                    for token in (parts.get("rn"), parts.get("ca")):
                        if token:
                            pos = norm_line.find(str(token))
                            if pos >= 0:
                                spans.append((pos, pos + len(str(token))))
                    for a, b in spans:
                        consume_span(state, line, a, b)
                    return {
                        "status": "PASS",
                        "pdf": line["text"],
                        "difference": "—",
                        "match_type": "RN_CA_COMBINED_EXACT"
                    }

                # Same structured line but wrong number = genuine FAIL.
                if (exp_rn and parts.get("rn")) or (exp_ca and parts.get("ca")):
                    return {
                        "status": "FAIL",
                        "pdf": line["text"],
                        "difference": f"Expected: {expected} | Found: {line['text']}",
                        "match_type": "RN_CA_COMBINED_MISMATCH"
                    }

            return {
                "status": "NOT FOUND",
                "pdf": "Not found",
                "difference": "Combined RN/CA value was not detected.",
                "match_type": "NOT_FOUND"
            }

        wanted_key = "rn" if field_type == "RN" else "ca"
        expected_component = expected_parts.get(wanted_key)
        if expected_component is None:
            digits = re.sub(r"\D", "", normalize_text(expected))
            expected_component = digits or str(expected).strip()

        for line in state["lines"]:
            if not line_is_available(line, state):
                continue
            parts = extract_rn_ca_components(line["text"])
            actual_component = parts.get(wanted_key)
            if actual_component is None:
                continue

            if str(actual_component) == str(expected_component):
                pos = line["norm"].find(str(actual_component))
                if pos >= 0:
                    consume_span(state, line, pos, pos + len(str(actual_component)))
                return {
                    "status": "PASS",
                    "pdf": line["text"],
                    "difference": "—",
                    "match_type": f"{wanted_key.upper()}_EXACT_SPAN"
                }

            return {
                "status": "FAIL",
                "pdf": line["text"],
                "difference": f"Expected: {expected} | Found: {wanted_key.upper()} {actual_component}",
                "match_type": f"{wanted_key.upper()}_MISMATCH"
            }

        return {
            "status": "NOT FOUND",
            "pdf": "Not found",
            "difference": f"Expected {wanted_key.upper()} value was not detected.",
            "match_type": "NOT_FOUND"
        }

    # -----------------------------------------------------
    # COO
    # -----------------------------------------------------
    if field_type == "COO":
        expected_coo = extract_coo_value(expected)
        expected_target = normalize_text(
            expected_coo if expected_coo else expected
        )
        expected_region = get_field_region(field_name)

        # Prefer the requested language. Do not let English COO satisfy French
        # COO merely because the country happens to be the same.
        candidates = []
        for line in state["lines"]:
            if not line_is_available(line, state):
                continue
            actual_coo = extract_coo_value(line["text"])
            if not actual_coo:
                continue
            region = coo_language(line["text"])
            candidates.append((line, actual_coo, region))

        preferred = [
            item for item in candidates
            if expected_region and item[2] == expected_region
        ]

        # When no language suffix is supplied, any COO language may be used.
        search_candidates = preferred if preferred else (
            candidates if not expected_region else []
        )

        for line, actual_coo, _region in search_candidates:
            if normalize_text(actual_coo) == expected_target:
                consume_lines(state, [line])
                return {
                    "status": "PASS",
                    "pdf": line["text"],
                    "difference": "—",
                    "match_type": "COO_EXACT"
                }

        # A same-language different COO is a genuine FAIL.
        for line, actual_coo, _region in search_candidates:
            consume_lines(state, [line])
            return {
                "status": "FAIL",
                "pdf": line["text"],
                "difference": (
                    f"Expected: {expected} | "
                    f"Found: {actual_coo}"
                ),
                "match_type": "COO_MISMATCH"
            }

        return {
            "status": "NOT FOUND",
            "pdf": "Not found",
            "difference": (
                "Expected COO was not detected in the requested language/region."
            ),
            "match_type": "NOT_FOUND"
        }

    # -----------------------------------------------------
    # GENDER
    # -----------------------------------------------------
    if field_type == "GENDER":
        expected_gender = extract_gender_value(expected)
        expected_target = normalize_text(
            expected_gender if expected_gender else expected
        )

        for line in state["lines"]:
            if not line_is_available(line, state):
                continue
            actual_gender = extract_gender_value(line["text"])
            if not actual_gender:
                continue

            if normalize_text(actual_gender) == expected_target:
                consume_lines(state, [line])
                return {
                    "status": "PASS",
                    "pdf": line["text"],
                    "difference": "—",
                    "match_type": "GENDER_EXACT"
                }

            consume_lines(state, [line])
            return {
                "status": "FAIL",
                "pdf": line["text"],
                "difference": (
                    f"Expected: {expected} | "
                    f"Found: {actual_gender}"
                ),
                "match_type": "GENDER_MISMATCH"
            }

        return {
            "status": "NOT FOUND",
            "pdf": "Not found",
            "difference": "Gender value was not detected.",
            "match_type": "NOT_FOUND"
        }

    # -----------------------------------------------------
    # SIZE
    # -----------------------------------------------------
    if field_type == "SIZE":
        expected_size = normalize_text(expected)
        expected_size = re.sub(
            r"^size\s*[:#-]?\s*",
            "",
            expected_size
        ).strip()

        # Multi-line/structured size fields must be treated as a block. This is
        # essential for Main_Size and OS_Size_* schemas where several regional
        # labels belong to one artwork size entry.
        if "\n" in str(expected) or len(str(expected_size).split()) >= 2:
            block = find_size_block_match(expected, field_name, state)
            if block:
                consume_lines(state, block["lines"])
                return {
                    "status": block["status"],
                    "pdf": block["pdf"],
                    "difference": block["difference"],
                    "match_type": block["match_type"]
                }

        # Prefer labeled single-size blocks.
        for line in state["lines"]:
            if not line_is_available(line, state):
                continue
            actual_size = extract_size_value(line["text"])
            if actual_size is None:
                continue

            if normalize_text(actual_size) == expected_size:
                consume_lines(state, [line])
                return {
                    "status": "PASS",
                    "pdf": line["text"],
                    "difference": "—",
                    "match_type": "SIZE_EXACT"
                }

            consume_lines(state, [line])
            return {
                "status": "FAIL",
                "pdf": line["text"],
                "difference": (
                    f"Expected: {expected} | "
                    f"Found: {actual_size}"
                ),
                "match_type": "SIZE_MISMATCH"
            }

        exact = find_exact_lines(
            expected,
            field_name,
            state,
            max_window=1
        )
        if exact:
            lines, match_info = exact
            consume_match(state, match_info)
            pdf_value = match_info.get("actual") or join_lines(lines)
            return {
                "status": "PASS",
                "pdf": pdf_value,
                "difference": "—",
                "match_type": match_info.get("match_type", "EXACT")
            }

        return {
            "status": "NOT FOUND",
            "pdf": "Not found",
            "difference": "Size value/block was not detected.",
            "match_type": "NOT_FOUND"
        }

    # -----------------------------------------------------
    # PRODUCTION MARK
    # -----------------------------------------------------
    if field_type == "PRODUCTION_MARK":
        expected_raw = str(expected).strip()
        for line in state["lines"]:
            if not line_is_available(line, state):
                continue
            tokens = [t for t in re.split(r"[^A-Za-z0-9]+", line["text"]) if t]
            for token in tokens:
                if token.casefold() == expected_raw.casefold():
                    pos = line["norm"].find(normalize_text(token))
                    if pos >= 0:
                        consume_span(state, line, pos, pos + len(normalize_text(token)))
                    return {
                        "status": "PASS",
                        "pdf": line["text"],
                        "difference": "—",
                        "match_type": "PRODUCTION_MARK_EXACT"
                    }
        return {
            "status": "NOT FOUND",
            "pdf": "Not found",
            "difference": "Production mark was not detected.",
            "match_type": "NOT_FOUND"
        }

    # -----------------------------------------------------
    # COLOR
    # -----------------------------------------------------
    if field_type == "COLOR":
        expected_color = normalize_text(expected)
        expected_color = re.sub(
            r"^(?:color|colour)\s*[:#-]?\s*",
            "",
            expected_color
        ).strip()

        for line in state["lines"]:
            if not line_is_available(line, state):
                continue
            actual_color = extract_color_value(line["text"])
            if actual_color is None:
                continue

            if normalize_text(actual_color) == expected_color:
                consume_lines(state, [line])
                return {
                    "status": "PASS",
                    "pdf": line["text"],
                    "difference": "—",
                    "match_type": "COLOR_EXACT"
                }

            consume_lines(state, [line])
            return {
                "status": "FAIL",
                "pdf": line["text"],
                "difference": (
                    f"Expected: {expected} | "
                    f"Found: {actual_color}"
                ),
                "match_type": "COLOR_MISMATCH"
            }

        return _generic_exact_field(expected, field_name, state)

    # -----------------------------------------------------
    # CARE
    # -----------------------------------------------------
    if field_type == "CARE":
        expected_care_norm = normalize_text(expected)
        comparison_text = normalize_text(state.get("comparison_text", ""))
        # A complete expected care phrase can span several extracted lines.
        # Validate it against the page-level text first, then use the regional
        # block only for the displayed PDF value/visual location.
        if expected_care_norm and expected_care_norm in comparison_text:
            care_region = find_care_region(state, field_name)
            actual_display = join_lines(care_region) if care_region else str(expected)
            return {
                "status": "PASS",
                "pdf": actual_display,
                "difference": "—",
                "match_type": "CARE_PAGE_TEXT_EXACT"
            }

        care_region = find_care_region(state, field_name)

        if care_region:
            actual_text = join_lines(care_region)
            expected_norm = normalize_text(expected)
            actual_norm = normalize_text(actual_text)

            if expected_norm == actual_norm or expected_norm in actual_norm:
                consume_lines(state, care_region)
                return {
                    "status": "PASS",
                    "pdf": actual_text,
                    "difference": "—",
                    "match_type": "CARE_EXACT_REGION"
                }

            expected_tokens = set(tokenize(expected))
            actual_tokens = set(tokenize(actual_text))
            common = expected_tokens & actual_tokens

            if expected_tokens and len(common) >= max(3, int(len(expected_tokens) * 0.55)):
                consume_lines(state, care_region)
                return {
                    "status": "FAIL",
                    "pdf": actual_text,
                    "difference": describe_text_difference(expected, actual_text),
                    "match_type": "CARE_MISMATCH"
                }

        return {
            "status": "NOT FOUND",
            "pdf": "Not found",
            "difference": "Care instruction was not detected in the relevant artwork region.",
            "match_type": "NOT_FOUND"
        }

    # -----------------------------------------------------
    # IDENTIFIER
    # -----------------------------------------------------
    if field_type == "IDENTIFIER":
        identifier = find_identifier_match(expected, state)
        if identifier:
            consume_span(
                state,
                identifier["line"],
                identifier["start"],
                identifier["end"]
            )
            return {
                "status": identifier["kind"],
                "pdf": identifier["actual"],
                "difference": identifier["difference"],
                "match_type": identifier["match_type"]
            }

        return {
            "status": "NOT FOUND",
            "pdf": "Not found",
            "difference": "Expected identifier was not detected.",
            "match_type": "NOT_FOUND"
        }

    # -----------------------------------------------------
    # GENERAL / BRAND / ATTRIBUTE / BATCH / QUANTITY
    # -----------------------------------------------------
    return _generic_exact_field(
        expected,
        field_name,
        state
    )


def _generic_exact_field(expected, field_name, state):
    exact = find_exact_lines(
        expected,
        field_name,
        state,
        max_window=8
    )

    if exact:
        lines, match_info = exact
        consume_match(state, match_info)

        if match_info.get("actual"):
            pdf_value = match_info["actual"]
        else:
            pdf_value = join_lines(lines)

        return {
            "status": "PASS",
            "pdf": pdf_value,
            "difference": "—",
            "match_type": match_info.get("match_type", "EXACT")
        }

    # IMPORTANT: unknown/general fields do not generate invented FAILs.
    return {
        "status": "NOT FOUND",
        "pdf": "Not found",
        "difference": "Expected value was not detected.",
        "match_type": "NOT_FOUND"
    }


# =========================================================
# FIELD MATCHING ORDER
# =========================================================

FIELD_PRIORITY = {
    "OSZ": 10,
    "BARCODE": 15,
    "RN": 20,
    "CA": 20,
    "IDENTIFIER": 20,
    "COO": 30,
    "CONTENT": 40,
    "CARE": 50,
    "GENDER": 60,
    "COLOR": 60,
    "SIZE": 60,
    "BATCH": 60,
    "QUANTITY": 60,
    "SYMBOL": 70,
    "BRAND": 80,
    "ATTRIBUTE": 90,
    "GENERAL": 100,
}


def order_fields_for_matching(fields):
    indexed = list(enumerate(fields))

    return [
        field
        for _index, field in sorted(
            indexed,
            key=lambda pair: (
                FIELD_PRIORITY.get(
                    get_field_type(pair[1]),
                    100
                ),
                pair[0]
            )
        )
    ]


# =========================================================
# BUILD REPORT
# =========================================================

def build_report(
    df,
    pdf_pages,
    selected_fields,
    product_type,
    page_row_mapping=None
):
    """
    Full validation report.

    Matching order is optimized for specificity, but report output is returned
    in the original selected-field order per page.
    """

    results = []
    field_no = 1

    # Pre-compute OSZ group sizes per row. They are used only to make sequence
    # mapping safe when more than one OSZ field exists.
    osz_fields = [
        field for field in selected_fields
        if get_field_type(field) == "OSZ"
    ]
    osz_group_size = len(osz_fields)

    for page_index, page in enumerate(pdf_pages):

        page_number = int(page.get("page", page_index + 1))

        if page_row_mapping is not None and page_number not in page_row_mapping:
            continue

        if page_row_mapping and page_number in page_row_mapping:
            excel_index = int(page_row_mapping[page_number])
        else:
            excel_index = page_index

        if excel_index >= len(df):
            for field in selected_fields:
                results.append({
                    "FIELD NO": field_no,
                    "PDF PAGE": page_number,
                    "EXCEL ROW": "N/A",
                    "FIELD": field,
                    "ORDER FORM DATA": "No Excel row",
                    "PDF OUTPUT": "No corresponding Order Form row",
                    "STATUS": "NOT FOUND",
                    "DIFFERENCE": "No corresponding Excel row."
                })
                field_no += 1
            continue

        row = df.iloc[excel_index]
        state = build_page_state(
            page,
            product_type
        )

        # Match in smart specificity order so a precise RN/OSZ/COO matcher gets
        # the relevant artwork before a broad textual field can consume it.
        matching_fields = order_fields_for_matching(
            selected_fields
        )

        page_results = {}

        for field in matching_fields:

            value = "" if is_blank_value(row[field]) else str(row[field]).strip()

            result = check_field(
                value,
                field,
                state,
                osz_group_size=osz_group_size
            )

            page_results[field] = {
                "field": field,
                "value": value,
                "result": result
            }

        # Return results in the user's original field order, not matching order.
        for field in selected_fields:
            item = page_results[field]
            result = item["result"]

            results.append({
                "FIELD NO": field_no,
                "PDF PAGE": page_number,
                "EXCEL ROW": excel_index + 2,
                "FIELD": field,
                "ORDER FORM DATA": item["value"],
                "PDF OUTPUT": result["pdf"],
                "STATUS": result["status"],
                "DIFFERENCE": result["difference"]
            })

            field_no += 1

    return pd.DataFrame(results)


# =========================================================
# AUTO DETECT
# =========================================================

def _auto_detect_page_text(page):
    """Return normalized OCR/direct text for one artwork page."""
    if not page:
        return ""

    values = []
    # Prefer a clean PDF text layer when it exists, while retaining OCR/alternate
    # text as supplementary evidence for scanned or partially editable artwork.
    for key in ("direct_text", "ocr_text", "ocr_alt_text", "text"):
        value = page.get(key, "")
        if value and str(value).strip():
            values.append(str(value))

    return "\n".join(values)


def _auto_detect_lines(page):
    text = _auto_detect_page_text(page)
    return [line for line in (normalize_text(x) for x in str(text).splitlines()) if line]


def _auto_number_evidence(expected, text):
    """Exact numeric evidence with boundaries; never accepts a substring."""
    numeric = normalize_numeric(expected)
    if numeric is None:
        return False
    return bool(
        re.search(
            rf"(?<![A-Za-z0-9]){re.escape(numeric)}(?![A-Za-z0-9])",
            normalize_text(text)
        )
    )


def _auto_identifier_evidence(expected, field_name, text):
    """Asymmetric identifier evidence used for item/style/supplier codes."""
    expected_compact = re.sub(r"[^a-z0-9]", "", normalize_text(expected))
    field_compact = (
        normalize_text(field_name)
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
    )

    # Supplier/vendor IDs such as USX are intentionally allowed at 3 chars
    # because the artwork may contain an appended static suffix (USX609).
    minimum_length = 3 if any(
        key in field_compact
        for key in ("supwsp", "supplier", "vendorid", "vendorcode")
    ) else 4

    if not expected_compact or len(expected_compact) < minimum_length:
        return False

    tokens = re.findall(r"[a-z0-9]+", normalize_text(text))
    for token in tokens:
        if token == expected_compact:
            return True
        if len(expected_compact) >= 5 and token.startswith(expected_compact):
            return True
        if len(expected_compact) == 3 and token.startswith(expected_compact):
            return True

    return False


def _auto_material_evidence(expected, text):
    """Strong evidence for canonical visible composition-description fields."""
    expected_parts = extract_content_values(expected)
    if not expected_parts:
        return False

    norm_text = normalize_text(text)
    compact_text_value = re.sub(r"[^a-z0-9%]", "", norm_text)

    matches = 0
    for part in expected_parts:
        part_norm = normalize_text(part)
        material = re.sub(r"[^a-z]", "", part_norm.casefold())
        pct_match = re.search(r"(\d+(?:\.\d+)?)\s*%?", part_norm)

        if not material or len(material) < 3:
            continue

        if material in compact_text_value:
            if pct_match and re.search(
                rf"{re.escape(pct_match.group(1))}\s*%?\s*{re.escape(material)}",
                norm_text,
            ):
                matches += 1
            elif re.search(rf"\b{re.escape(material)}\b", norm_text):
                matches += 1

    required = 2 if len(expected_parts) >= 2 else 1
    has_composition_shape = bool(
        re.search(r"\d+(?:\.\d+)?\s*%?", norm_text)
        and re.search(r"[a-z]{3,}", norm_text)
    )
    return matches >= required and has_composition_shape


def _auto_strip_accents(text):
    """Return a comparison-safe lowercase string with accents removed."""
    value = unicodedata.normalize("NFKD", str(text or "")).casefold()
    return "".join(char for char in value if not unicodedata.combining(char))


def _auto_dynamic_composition_evidence(expected, field_name, text):
    """
    Detect visible fiber/composition data without depending on one Excel schema.

    The Order Form column name identifies the semantic family, while this helper
    proves the value is actually represented in the artwork. It handles
    multilingual accents, split words such as POLY-ESTER, and multiple
    percentage/material pairs.
    """
    expected_raw = str(expected or "").strip()
    if not expected_raw:
        return False

    field_compact = (
        normalize_text(field_name)
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
    )
    if not any(token in field_compact for token in (
        "fib", "fiber", "fibre", "fabric", "content",
        "composition", "compodsc", "lhcompodsc", "fabrication", "material"
    )):
        return False

    expected_clean = _auto_strip_accents(expected_raw)
    text_clean = _auto_strip_accents(text or "")

    # Normalize separators but retain words/digits for pair matching.
    expected_clean = re.sub(r"[^a-z0-9%]+", " ", expected_clean)
    text_clean = re.sub(r"[^a-z0-9%]+", " ", text_clean)

    # Capture percentage + material tokens. This is intentionally permissive
    # about punctuation because artwork may render the same composition across
    # different lines or with hyphenation.
    pairs = re.findall(
        r"(\d+(?:\.\d+)?)\s*%?\s*([a-z][a-z0-9]*(?:\s+[a-z][a-z0-9]*)?)",
        expected_clean,
        flags=re.IGNORECASE,
    )

    # Keep only plausible material words and de-duplicate while preserving order.
    candidates = []
    excluded = {
        "shell", "liner", "lining", "body", "cuerpo", "forro",
        "extrieur", "exterieur", "doublure", "made", "in", "hecho"
    }
    for pct, material_group in pairs:
        material = re.sub(r"\s+", " ", material_group).strip()
        material_words = material.split()
        if material_words and material_words[-1] in excluded:
            material_words = material_words[:-1]
        if not material_words:
            continue
        material = " ".join(material_words)
        if len(re.sub(r"[^a-z]", "", material)) < 3:
            continue
        item = (pct, material)
        if item not in candidates:
            candidates.append(item)

    # If the field has composition-shaped data, require the key pairs to be
    # present. With 2+ expected pairs, two independent matches are enough to
    # establish high-confidence artwork evidence.
    matched = 0
    for pct, material in candidates:
        material_compact = re.sub(r"[^a-z]", "", material)
        pct_pattern = re.escape(pct)
        # Permit spaces/hyphens between a material's letters (POLY-ESTER).
        material_letters = [re.escape(ch) for ch in material_compact]
        flexible_material = r"\s*[-–—]?\s*".join(material_letters)
        if re.search(
            rf"{pct_pattern}\s*%?\s*(?:[a-z0-9]+\s*)?{flexible_material}",
            text_clean,
            flags=re.IGNORECASE,
        ):
            matched += 1
            continue

        # Fall back to material-only evidence when OCR/PDF extraction breaks
        # the percentage association but preserves the material itself.
        if re.search(rf"\b{flexible_material}\b", text_clean, flags=re.IGNORECASE):
            matched += 1

    if candidates:
        required = min(2, len(candidates))
        if matched >= required:
            return True

        # Catch legitimate variable-field FAILs, such as POLYSTER in the Order
        # Form versus POLYESTER in artwork. We require composition-shaped artwork
        # and multiple material similarities rather than a generic fuzzy match.
        from rapidfuzz import fuzz
        material_expected = [
            re.sub(r"[^a-z]", "", material)
            for _pct, material in candidates
        ]
        actual_tokens = [
            re.sub(r"[^a-z]", "", token)
            for token in re.findall(r"[a-z]+", text_clean, flags=re.IGNORECASE)
            if len(re.sub(r"[^a-z]", "", token)) >= 4
        ]
        fuzzy_hits = 0
        for expected_material in material_expected:
            if not expected_material:
                continue
            if any(
                fuzz.ratio(expected_material, actual_material) >= 84
                for actual_material in actual_tokens
            ):
                fuzzy_hits += 1
        return fuzzy_hits >= min(2, len(material_expected))

    # Non-percentage content/material fields can still auto-detect if their
    # actual phrase is directly represented. This remains conservative.
    target_compact = re.sub(r"[^a-z0-9]", "", expected_clean)
    text_compact = re.sub(r"[^a-z0-9]", "", text_clean)
    return bool(target_compact) and len(target_compact) >= 6 and target_compact in text_compact


def _auto_multiline_size_evidence(expected, text):
    """Detect a structured size block even when OCR line breaks differ."""
    expected_raw = str(expected or "").strip()
    if not expected_raw:
        return False

    # Strong exact/compact match first.
    expected_norm = normalize_text(expected_raw)
    actual_norm = normalize_text(text or "")
    expected_compact = re.sub(r"[^a-z0-9]", "", expected_norm)
    actual_compact = re.sub(r"[^a-z0-9]", "", actual_norm)
    if expected_compact and expected_compact in actual_compact:
        return True

    # Compare meaningful size tokens from each line/block. Require multiple
    # independent anchors so a random number on artwork cannot trigger a match.
    expected_tokens = [
        token for token in re.findall(r"[a-z]+|\d+(?:[-/]\d+)*", _auto_strip_accents(expected_raw))
        if len(token) >= 1
    ]
    if not expected_tokens:
        return False

    actual_clean = _auto_strip_accents(text or "")
    hits = 0
    considered = []
    for token in expected_tokens:
        compact = re.sub(r"[^a-z0-9]", "", token)
        if not compact:
            continue
        # Ignore very common single-letter multilingual size labels unless they
        # appear with their neighboring numeric size range.
        if len(compact) == 1 and compact.isalpha():
            continue
        considered.append(compact)
        if compact in re.sub(r"[^a-z0-9]", "", actual_clean):
            hits += 1

    if len(considered) >= 4:
        return hits >= max(3, int(len(considered) * 0.55))
    if len(considered) >= 2:
        return hits >= len(considered)
    return False


def _auto_coo_evidence(expected, field_name, text):
    """
    High-confidence COO Auto Detect evidence.

    Auto Detect must find both PASS and FAIL cases. Therefore an exact value is
    strong evidence, but for COO we also accept a matching language/COO marker
    with a different country so a wrong country can still be detected as a field
    that belongs in the artwork.
    """
    from rapidfuzz import fuzz

    expected_coo = extract_coo_value(expected)
    target = normalize_text(expected_coo if expected_coo else expected)
    region = get_field_region(field_name)
    norm_text = normalize_text(text)

    # One/two-letter language/origin codes are internal codes, not visible COO.
    if re.fullmatch(r"[a-z]{1,2}", target):
        target = ""

    target_compact = re.sub(r"[^a-z0-9]", "", target)
    if target_compact and len(target_compact) >= 4:
        for line in _auto_detect_lines({"ocr_text": text}):
            line_compact = re.sub(r"[^a-z0-9]", "", line)
            if target_compact == line_compact:
                return True
            if target_compact and target_compact in line_compact:
                return True
            if fuzz.ratio(target_compact, line_compact) >= 92:
                return True

    # Region-specific structural evidence lets Auto Detect catch a wrong COO
    # value rather than dropping the field simply because its expected country
    # was not found.
    markers = {
        "EN": ("made in", "country of origin"),
        "FR": ("fabrique en", "fabriqué en"),
        "SP": ("hecho en",),
        "": ("made in", "fabrique en", "fabriqué en", "hecho en"),
    }
    for marker in markers.get(region, markers[""]):
        if marker in norm_text:
            # Require at least one alphabetic/numeric location token after the
            # marker so a stray heading does not count as COO evidence.
            after = norm_text.split(marker, 1)[1].strip()
            if re.search(r"[a-z]{3,}", after):
                return True

    return False


def _auto_semantic_type(field_name):
    """
    Auto Detect-only semantic classifier.

    This is deliberately separate from get_field_type() so expanding Auto Detect
    does not alter the established manual comparison engine or its matching rules.
    It recognizes schema families such as MIN_*, FIB_* and WC_* rather than
    hardcoding individual job columns.
    """
    raw = str(field_name or "").casefold().strip()
    compact = normalize_text(field_name).replace(" ", "").replace("_", "").replace("-", "")

    if (
        "coo" in compact
        or "countryoforigin" in compact
        or "countryorigin" in compact
        or "madein" in compact
        or compact == "origin"
        or re.match(r"^min(?:_|$)", raw)
        or compact in {"minen", "minsp", "minfr", "minenglish", "minspanish", "minfrench"}
    ):
        return "COO"

    if (
        "fiber" in compact
        or "fibre" in compact
        or "fabric" in compact
        or "content" in compact
        or "composition" in compact
        or "compodsc" in compact
        or "lhcompodsc" in compact
        or "fabrication" in compact
        or "material" in compact
        or compact.startswith("fib")
    ):
        return "CONTENT"

    if (
        "care" in compact
        or "wash" in compact
        or "washing" in compact
        or "laundry" in compact
        or "instruction" in compact
        or compact.startswith("wc")
    ):
        return "CARE"

    return None


def _auto_generic_field_allowed(field_name):
    """
    Allow only known artwork-variable semantics among GENERAL technical fields.
    This prevents operational/database columns from being auto-selected just
    because they are populated.
    """
    compact = (
        normalize_text(field_name)
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
    )

    patterns = (
        "stylewofinish",
        "cdstyle",
        "cdfinishing",
        "finishing",
        "cdimport",
        "import",
        "designstyle",
        "lblstyle",
        "antfamily",
        "family",
        "compodsc",
        "lhcompodsc",
    )
    return any(token in compact for token in patterns)


def _auto_content_field_allowed(field_name):
    """
    Allow semantic CONTENT families through Auto Detect.

    The actual artwork-evidence test remains responsible for deciding whether
    the populated value is really present in the PDF. This function therefore
    must not whitelist individual customer/job column names.
    """
    compact = (
        normalize_text(field_name)
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
    )

    semantic_tokens = (
        "fiber", "fibre", "fabric", "content", "composition",
        "compodsc", "lhcompodsc", "fabrication", "material", "fib"
    )
    return any(token in compact for token in semantic_tokens)


def _auto_text_evidence(expected, field_name, text):
    """Conservative Auto Detect evidence check."""
    from rapidfuzz import fuzz

    field_type = get_field_type(field_name)

    # Custom-font symbols must be checked before normalize_text() because
    # normalization can change/remove the actual keystroke/codepoint.
    if field_type == "SYMBOL":
        raw_expected = str(expected or "")
        if not raw_expected.strip():
            return False
        return any(
            _symbol_text_matches(raw_expected, line)
            for line in str(text or "").splitlines()
        )

    expected_norm = normalize_text(expected)
    if not expected_norm:
        return False

    if field_type == "GENERAL":
        field_type = _auto_semantic_type(field_name) or field_type
    compact_expected = re.sub(r"[^a-z0-9]", "", expected_norm)
    norm_text = normalize_text(text)

    if field_type == "BARCODE":
        expected_digits = normalize_barcode_digits(expected)
        if len(expected_digits) not in {8, 12, 13, 14}:
            return False
        for line in str(text or "").splitlines():
            for candidate in _barcode_chunks_from_line(line):
                if barcode_equivalent(expected_digits, candidate["digits"]):
                    return True
        return False

    if field_type == "IDENTIFIER":
        return _auto_identifier_evidence(expected, field_name, norm_text)

    if field_type == "RN":
        compact_field = normalize_text(field_name).replace(" ", "").replace("_", "").replace("-", "")
        if compact_field == "rnca":
            parts = extract_rn_ca_components(expected)
            rn = parts.get("rn")
            ca = parts.get("ca")
            if rn and ca:
                return bool(re.search(rf"\brn\s*[#:\-./ ]*{re.escape(rn)}\b", norm_text)) and bool(
                    re.search(rf"\bca\s*[#:\-./ ]*{re.escape(ca)}\b", norm_text)
                )
        digits = re.sub(r"\D", "", expected_norm)
        if not digits or len(digits) < 4:
            return False
        return bool(re.search(rf"\brn\s*[#:\-./ ]*{re.escape(digits)}\b", norm_text)) or bool(
            re.search(rf"(?<!\d){re.escape(digits)}(?!\d)", norm_text)
        )

    if field_type == "CA":
        digits = re.sub(r"\D", "", expected_norm)
        if not digits or len(digits) < 4:
            return False
        return bool(re.search(rf"\bca\s*[#:\-./ ]*{re.escape(digits)}\b", norm_text))

    if field_type == "COO":
        return _auto_coo_evidence(expected, field_name, norm_text)

    if field_type == "CONTENT":
        return _auto_content_field_allowed(field_name) and _auto_dynamic_composition_evidence(
            expected, field_name, text
        )

    if field_type == "CARE":
        tokens = [t for t in re.findall(r"[a-z]+", expected_norm) if len(t) >= 4]
        if len(tokens) >= 3:
            unique = set(tokens)
            hits = sum(1 for token in unique if token in norm_text)
            if hits >= max(3, int(len(unique) * 0.35)):
                return True

        # If wording is wrong, Auto Detect still needs to recognize the relevant
        # language care block so the field can be selected and subsequently FAIL.
        region = get_field_region(field_name)
        care_markers = {
            "EN": ("machine wash", "wash", "bleach", "tumble dry", "cool iron", "dry clean"),
            "FR": ("laver", "blanchiment", "secher", "secher", "repasser", "nettoyer a sec"),
            "SP": ("lavar", "blanqueador", "cloro", "secar", "planchar", "limpiar en seco"),
            "": ("machine wash", "laver", "lavar", "bleach", "blanchiment", "blanqueador"),
        }
        marker_set = care_markers.get(region, care_markers[""])
        marker_hits = sum(1 for marker in marker_set if marker in norm_text)
        return marker_hits >= 1

    if field_type == "SIZE":
        return _auto_multiline_size_evidence(expected, text)

    if field_type == "SYMBOL":
        # Exact custom-font keystroke matching. This deliberately avoids fuzzy
        # matching because the symbol/codepoint itself is the data.
        for line in str(text or "").splitlines():
            if _symbol_text_matches(expected_norm, line):
                return True
        return False

    if field_type == "PRODUCTION_MARK":
        expected_raw = str(expected or "").strip()
        if not expected_raw:
            return False
        # Exact token match first; this supports one-letter marks such as V.
        for line in str(text or "").splitlines():
            tokens = [t for t in re.split(r"[^A-Za-z0-9]+", line) if t]
            if any(token.casefold() == expected_raw.casefold() for token in tokens):
                return True
        return False

    if field_type == "COLOR":
        if expected_norm in norm_text:
            return True
        lines = _auto_detect_lines({"ocr_text": text})
        if not lines or len(expected_norm) < 4:
            return False
        return max(fuzz.ratio(expected_norm, line) for line in lines) >= 86

    if field_type == "GENDER":
        aliases = {
            "men": ("men", "men's", "mens", "male"),
            "women": ("women", "women's", "womens", "female"),
            "boys": ("boys", "boy", "boy's", "boys'"),
            "girls": ("girls", "girl", "girl's", "girls'"),
        }
        key = expected_norm.replace("'", "")
        for base, variants in aliases.items():
            if key == base or key.rstrip("s") == base.rstrip("s"):
                return any(normalize_text(v) in norm_text for v in variants)
        return expected_norm in norm_text

    if field_type == "BRAND":
        if compact_expected and compact_expected in re.sub(r"[^a-z0-9]", "", norm_text):
            return True
        lines = _auto_detect_lines({"ocr_text": text})
        return bool(lines) and len(expected_norm) >= 4 and max(
            fuzz.ratio(expected_norm, line) for line in lines
        ) >= 88

    if field_type == "ATTRIBUTE":
        tokens = [t for t in re.findall(r"[a-z0-9]+", expected_norm) if len(t) >= 3]
        if not tokens:
            return False
        best = 0
        for line in _auto_detect_lines({"ocr_text": text}):
            line_tokens = set(re.findall(r"[a-z0-9]+", line))
            overlap = len(set(tokens) & line_tokens) / max(1, len(set(tokens)))
            best = max(best, overlap)
        return best >= 0.60

    if field_type in {"QUANTITY", "BATCH"}:
        # Quantities/lots are deliberately manual because artwork contains
        # many unrelated numbers and barcode data.
        return False

    if field_type == "OSZ":
        return _auto_number_evidence(expected, norm_text)

    # GENERAL: only selected semantic technical fields are eligible.
    if not _auto_generic_field_allowed(field_name):
        return False

    # Never auto-detect a one-letter technical code such as MADE_IN=F.
    if re.fullmatch(r"[a-z]{1,2}", expected_norm):
        return False

    if normalize_numeric(expected) is not None:
        return _auto_number_evidence(expected_norm, norm_text)

    lines = _auto_detect_lines({"ocr_text": text})
    if not lines:
        return False

    if compact_expected and compact_expected in re.sub(r"[^a-z0-9]", "", norm_text):
        return True

    if len(expected_norm) >= 6:
        return max(fuzz.ratio(expected_norm, line) for line in lines) >= 90

    return False


def _auto_candidate_priority(field_name):
    """Lower number = preferred canonical source column when duplicate values exist."""
    compact = (
        normalize_text(field_name)
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
    )

    if compact.startswith("compodsc"):
        return 0
    if compact.startswith("lhcompodsc"):
        return 10
    return 20


def _auto_field_signature(field_name, value):
    """Semantic duplicate signature so equivalent columns do not flood Auto Detect."""
    field_type = get_field_type(field_name)

    if field_type == "CONTENT":
        parts = extract_content_values(value)
        if parts:
            parsed = []
            for part in parts:
                match = re.match(
                    r"(\d+(?:\.\d+)?)%?\s*(.+)",
                    normalize_text(part)
                )
                if match:
                    parsed.append(
                        (
                            match.group(1),
                            normalize_text(match.group(2))
                        )
                    )
                else:
                    parsed.append(("", normalize_text(part)))
            return (field_type, tuple(parsed))

    return (field_type, normalize_text(value))


def auto_detect_fields(
    df,
    output_pages,
    product_type,
    page_row_mapping=None
):
    """
    Controlled Auto Detect.

    A field is selected only when:
      1. it is populated in the mapped Order Form row,
      2. its column belongs to an allowed artwork-variable family, and
      3. its actual value has strong evidence in the mapped artwork.

    Population alone is never enough.
    """
    available_fields = get_available_fields(df)
    allowed_types = {
        "IDENTIFIER",
        "BARCODE",
        "BRAND",
        "GENDER",
        "SIZE",
        "COLOR",
        "COO",
        "CONTENT",
        "CARE",
        "ATTRIBUTE",
        "RN",
        "CA",
        "OSZ",
        "SYMBOL",
        "PRODUCTION_MARK",
        "SYMBOL",
        "GENERAL",
    }

    candidates = []
    for field in available_fields:
        if is_admin_field(field):
            continue

        field_type = get_field_type(field)
        auto_type = _auto_semantic_type(field)

        if field_type not in allowed_types and not auto_type:
            continue

        # GENERAL operational/database columns remain protected. A GENERAL field
        # is allowed through Auto Detect only when the separate semantic family
        # classifier recognizes it or the legacy generic whitelist recognizes it.
        if field_type == "GENERAL" and not auto_type and not _auto_generic_field_allowed(field):
            continue

        compact = (
            normalize_text(field)
            .replace(" ", "")
            .replace("_", "")
            .replace("-", "")
        )

        # Internal/support codes are useful metadata but are not normally visible
        # artwork fields. Keep them out of Auto Detect without changing manual
        # field selection or the underlying comparison engine.
        if compact in {"carecode", "caresuffixcode"}:
            continue

        # Never auto-select translated/internal material columns.
        if any(token in compact for token in (
            "p1mat",
            "multi",
            "translation",
            "greek",
            "arabic",
            "turkish",
            "indonesia",
            "matfull",
        )):
            if not compact.startswith("compodsc"):
                continue

        candidates.append(field)

    if not candidates or not output_pages:
        return []

    rows_to_check = []
    for page in output_pages:
        page_number = int(page.get("page", 1))

        if page_row_mapping is not None:
            if page_number not in page_row_mapping:
                continue
            row_index = int(page_row_mapping[page_number])
        elif len(df) == 1:
            row_index = 0
        else:
            row_index = page_number - 1

        if 0 <= row_index < len(df):
            rows_to_check.append((page, df.iloc[row_index]))

    if not rows_to_check:
        return []

    # Prefer canonical composition descriptions before equivalent helper columns.
    candidates = sorted(
        enumerate(candidates),
        key=lambda item: (_auto_candidate_priority(item[1]), item[0])
    )
    candidates = [field for _idx, field in candidates]

    detected = []

    for field in candidates:
        found = False

        for page, row in rows_to_check:
            value = row.get(field, "")
            if is_blank_value(value):
                continue

            if _auto_text_evidence(
                value,
                field,
                _auto_detect_page_text(page)
            ):
                found = True
                break

        if not found:
            continue

        # IMPORTANT: do not collapse different Order Form columns just because
        # their values happen to be identical. Regional/language fields such as
        # FIB_SP, FIB_Mexico and FIB_SP_Mexico may intentionally carry the same
        # text today while still being separate selectable artwork fields.
        detected.append(field)

    # Return fields in their original Excel order.
    original_order = {str(column): idx for idx, column in enumerate(df.columns)}
    detected.sort(key=lambda field: original_order.get(field, 10**9))
    return detected


# =========================================================
# STATUS COLORS
# =========================================================

def style_status(value):
    if value == "PASS":
        return (
            "background-color: #238636;"
            "color: white;"
            "font-weight: bold;"
        )

    if value == "FAIL":
        return (
            "background-color: #da3633;"
            "color: white;"
            "font-weight: bold;"
        )

    if value == "NOT FOUND":
        return (
            "background-color: #9e6a03;"
            "color: white;"
            "font-weight: bold;"
        )

    if value == "SKIP":
        return (
            "background-color: #555555;"
            "color: white;"
            "font-weight: bold;"
        )

    return ""



def create_excel_report(report, product_type, comparison_method, selected_fields, visual_pages=None):
    """Create a professional Excel QC report, including full artwork pages with highlights."""

    output = BytesIO()
    wb = Workbook()

    NAVY = "1F4E78"
    LIGHT_BLUE = "D9EAF7"
    GREEN = "238636"
    RED = "DA3633"
    AMBER = "9E6A03"
    GREY = "555555"
    WHITE = "FFFFFF"
    LIGHT_BORDER = "D9E1F2"

    pass_count = int((report["STATUS"] == "PASS").sum()) if not report.empty else 0
    fail_count = int((report["STATUS"] == "FAIL").sum()) if not report.empty else 0
    not_found_count = int((report["STATUS"] == "NOT FOUND").sum()) if not report.empty else 0
    skip_count = int((report["STATUS"] == "SKIP").sum()) if not report.empty else 0
    total_checks = len(report)

    ws = wb.active
    ws.title = "Summary"
    ws.merge_cells("A1:F2")
    ws["A1"] = "PDF PROOFREADING QC REPORT"
    ws["A1"].font = Font(bold=True, size=20, color=WHITE)
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    summary = [
        ("Comparison Method", comparison_method or "—"),
        ("Product Type", product_type or "—"),
        ("Selected / Detected Fields", len(selected_fields or [])),
        ("Total Field Checks", total_checks),
        ("PASS", pass_count),
        ("FAIL", fail_count),
        ("NOT FOUND", not_found_count),
        ("IGNORED / SKIP", skip_count),
    ]

    for r, (label, value) in enumerate(summary, start=4):
        ws.cell(r, 1, label)
        ws.cell(r, 2, value)
        ws.cell(r, 1).font = Font(bold=True)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=LIGHT_BLUE)
        ws.cell(r, 1).alignment = Alignment(vertical="center")
        ws.cell(r, 2).alignment = Alignment(vertical="center", wrap_text=True)

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 28
    ws.freeze_panes = "A4"

    finding_row = 14
    ws.merge_cells(start_row=finding_row, start_column=1, end_row=finding_row, end_column=6)
    ws.cell(finding_row, 1, "QC NOTES")
    ws.cell(finding_row, 1).font = Font(bold=True, color=WHITE)
    ws.cell(finding_row, 1).fill = PatternFill("solid", fgColor=NAVY)

    notes = []
    if fail_count:
        notes.append(f"{fail_count} FAIL result(s) require review.")
    if not_found_count:
        notes.append(f"{not_found_count} field(s) could not be reliably located.")
    if skip_count:
        notes.append(f"{skip_count} blank/ignored field check(s) were skipped.")
    if not notes:
        notes.append("All checked fields passed.")
    for idx, note in enumerate(notes, start=finding_row + 1):
        ws.merge_cells(start_row=idx, start_column=1, end_row=idx, end_column=6)
        ws.cell(idx, 1, "• " + note)
        ws.cell(idx, 1).alignment = Alignment(wrap_text=True, vertical="top")

    comparison = wb.create_sheet("Field Comparison")
    display_report = add_visual_column(report, selected_fields)
    headers = list(display_report.columns)
    header_fill = PatternFill("solid", fgColor=NAVY)
    for col_idx, header in enumerate(headers, start=1):
        cell = comparison.cell(1, col_idx, header)
        cell.font = Font(bold=True, color=WHITE)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    if "VISUAL" in headers:
        visual_header = comparison.cell(1, headers.index("VISUAL") + 1)
        visual_header.comment = None

    for row_idx, row in enumerate(display_report.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row, start=1):
            cell = comparison.cell(row_idx, col_idx, value)
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    status_col = None
    for idx, header in enumerate(headers, start=1):
        if header == "STATUS":
            status_col = idx
            break
    if status_col:
        for row_idx in range(2, comparison.max_row + 1):
            cell = comparison.cell(row_idx, status_col)
            status = str(cell.value or "")
            if status == "PASS":
                cell.fill = PatternFill("solid", fgColor=GREEN)
                cell.font = Font(color=WHITE, bold=True)
            elif status == "FAIL":
                cell.fill = PatternFill("solid", fgColor=RED)
                cell.font = Font(color=WHITE, bold=True)
            elif status == "NOT FOUND":
                cell.fill = PatternFill("solid", fgColor=AMBER)
                cell.font = Font(color=WHITE, bold=True)
            elif status == "SKIP":
                cell.fill = PatternFill("solid", fgColor=GREY)
                cell.font = Font(color=WHITE, bold=True)

    visual_col = headers.index("VISUAL") + 1 if "VISUAL" in headers else None
    if visual_col:
        field_col = headers.index("FIELD") + 1
        field_colors = get_field_visual_colors(selected_fields)
        for row_idx in range(2, comparison.max_row + 1):
            field_name = str(comparison.cell(row_idx, field_col).value or "")
            color = field_colors.get(field_name, "6B7280").lstrip("#").upper()
            cell = comparison.cell(row_idx, visual_col)
            cell.value = ""
            cell.fill = PatternFill("solid", fgColor=color)
            cell.font = Font(color=WHITE, bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")

    width_map = {
        "FIELD NO": 12,
        "PDF PAGE": 12,
        "EXCEL ROW": 12,
        "FIELD": 24,
        "ORDER FORM DATA": 42,
        "PDF OUTPUT": 52,
        "STATUS": 16,
        "VISUAL": 12,
        "DIFFERENCE": 58,
        "MATCH TYPE": 22,
    }
    for col_idx, header in enumerate(headers, start=1):
        comparison.column_dimensions[get_column_letter(col_idx)].width = width_map.get(header, 20)
    comparison.row_dimensions[1].height = 32
    for row_idx in range(2, comparison.max_row + 1):
        comparison.row_dimensions[row_idx].height = 55
    comparison.freeze_panes = "A2"
    comparison.auto_filter.ref = comparison.dimensions

    thin = Side(style="thin", color=LIGHT_BORDER)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in comparison.iter_rows():
        for cell in row:
            cell.border = border

    if visual_pages:
        visual = wb.create_sheet("Artwork Visual Validation")
        visual.column_dimensions["A"].width = 22
        visual.column_dimensions["B"].width = 24
        visual.column_dimensions["C"].width = 24
        visual["A1"] = "ARTWORK VISUAL VALIDATION"
        visual["A1"].font = Font(bold=True, size=16, color=WHITE)
        visual["A1"].fill = PatternFill("solid", fgColor=NAVY)
        visual.merge_cells("A1:C2")
        visual["A1"].alignment = Alignment(horizontal="center", vertical="center")

        row_cursor = 4
        for page in visual_pages:
            page_num = page.get("page")
            page_report = report[report["PDF PAGE"] == page_num] if "PDF PAGE" in report.columns else report.iloc[0:0]
            field_colors = get_field_visual_colors(selected_fields)
            highlighted = build_highlighted_page_image(
                page,
                page_report,
                field_colors=field_colors,
            )
            if highlighted is None:
                continue
            visual.cell(row_cursor, 1, f"Artwork Page {page_num}")
            visual.cell(row_cursor, 1).font = Font(bold=True, size=13, color=WHITE)
            visual.cell(row_cursor, 1).fill = PatternFill("solid", fgColor=NAVY)
            visual.merge_cells(start_row=row_cursor, start_column=1, end_row=row_cursor, end_column=3)
            row_cursor += 1
            image_data = _visual_image_bytes(highlighted)
            if image_data:
                img = XLImage(BytesIO(image_data))
                img.width = min(560, highlighted.width)
                img.height = int(highlighted.height * (img.width / highlighted.width))
                anchor = f"A{row_cursor}"
                visual.add_image(img, anchor)
                row_height_count = max(20, int(img.height / 1.35))
                for r in range(row_cursor, row_cursor + max(1, int(row_height_count / 15))):
                    visual.row_dimensions[r].height = 15
                row_cursor += max(35, int(img.height / 14))
            row_cursor += 2

    wb.save(output)
    output.seek(0)
    return output.getvalue()


def main():
    """Render Tool 1: Order Form → Output Check."""

    _apply_tool_css()

    # =========================================================
    # TOOL 1 SESSION STATE
    # =========================================================
    if "of_reset_id" not in st.session_state:
        st.session_state["of_reset_id"] = 0

    if "of_report" not in st.session_state:
        st.session_state["of_report"] = None

    if "of_visual_pages" not in st.session_state:
        st.session_state["of_visual_pages"] = None

    if "of_auto_detected_fields" not in st.session_state:
        st.session_state["of_auto_detected_fields"] = []

    if "of_auto_detect_key" not in st.session_state:
        st.session_state["of_auto_detect_key"] = None

    if "of_auto_output_pages" not in st.session_state:
        st.session_state["of_auto_output_pages"] = None

    # A code update must never leave an old comparison report visible in the
    # same Streamlit session. This is presentation/session hygiene only; it
    # does not modify any comparison decision.
    if st.session_state.get("of_report_build_version") != AUTO_DETECT_ENGINE_VERSION:
        st.session_state["of_report"] = None
        st.session_state["of_visual_pages"] = None
        st.session_state["of_report_selected_fields"] = []
        st.session_state["of_report_product_type"] = None
        st.session_state["of_report_comparison_method"] = None
        st.session_state["of_report_build_version"] = AUTO_DETECT_ENGINE_VERSION

    # =========================================================
    # TITLE
    # =========================================================
    st.markdown(
        '<div class="main-title">🔍 PDF Proofreader</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="sub-title">'
        'Compare selected variable Order Form fields against PDF artwork.'
        '</div>',
        unsafe_allow_html=True
    )

    # =========================================================
    # TOOL NAVIGATION
    # =========================================================
    nav_left, nav_right = st.columns([1, 1])

    with nav_left:
        if st.button("← HOME", key="of_back_home", width="stretch"):
            st.session_state["selected_tool"] = None
            st.session_state["of_report"] = None
            st.session_state["of_visual_pages"] = None
            st.session_state["of_auto_detected_fields"] = []
            st.session_state["of_auto_detect_key"] = None
            st.session_state["of_auto_output_pages"] = None
            st.session_state["of_report_build_version"] = None
            st.rerun()

    with nav_right:
        if st.button(
            "🆕 NEW START",
            key="of_new_start",
            width="stretch"
        ):
            st.session_state["of_reset_id"] += 1
            st.session_state["of_report"] = None
            st.session_state["of_visual_pages"] = None
            st.session_state["of_auto_detected_fields"] = []
            st.session_state["of_auto_detect_key"] = None
            st.session_state["of_auto_output_pages"] = None
            st.session_state["of_report_build_version"] = None
            st.rerun()

    # =========================================================
    # PRODUCT TYPE
    # =========================================================
    product_type = st.selectbox(
        "Select Product Type",
        options=[
            "----- SELECT -----",
            "PFL",
            "HTL",
            "Other"
        ],
        index=0,
        key=f"of_product_type_{st.session_state['of_reset_id']}",
        help=(
            "PFL = panelled artwork where variable data may continue "
            "across panels. HTL / Other = standard continuous-data comparison."
        )
    )

    if product_type == "----- SELECT -----":
        st.info(
            "Please select a Product Type before starting the comparison."
        )

    # =========================================================
    # UPLOAD AREA
    # =========================================================
    left_column, right_column = st.columns(2)

    with left_column:
        st.markdown(
            '<div class="section-title">📊 Order Form</div>',
            unsafe_allow_html=True
        )

        excel_file = st.file_uploader(
            "Upload Excel Order Form",
            type=["xlsx", "xls"],
            key=f"excel_upload_{st.session_state['of_reset_id']}"
        )

    with right_column:
        st.markdown(
            '<div class="section-title">📄 Output Artwork</div>',
            unsafe_allow_html=True
        )

        output_file = st.file_uploader(
            "Upload Output Artwork",
            type=["pdf", "jpg", "jpeg", "png"],
            key=f"output_upload_{st.session_state['of_reset_id']}"
        )

    # =========================================================
    # LOAD ORDER FORM
    # =========================================================
    df = None

    if excel_file:
        try:
            df = load_excel(excel_file)
        except Exception as error:
            st.error(
                f"Unable to read the Excel Order Form: {error}"
            )
            return

    # =========================================================
    # FILE INFORMATION + PAGE SELECTION + ROW MAPPING
    # =========================================================
    output_page_count = 0
    selected_pdf_pages = []
    page_row_mapping = {}
    mapping_ready = False

    if excel_file and output_file:
        try:
            output_page_count = get_output_page_count(output_file)
        except Exception as error:
            st.error(f"Unable to determine Output Artwork page count: {error}")
            output_page_count = 0

        st.markdown(
            '<div class="section-title">📌 File Information</div>',
            unsafe_allow_html=True
        )

        info1, info2, info3 = st.columns(3)
        with info1:
            st.metric("Excel Data Rows", len(df))
        with info2:
            st.metric("Output Pages", output_page_count)
        with info3:
            extension = str(output_file.name).split(".")[-1].upper()
            st.metric("Output Type", extension)

        if output_page_count == 1:
            selected_pdf_pages = [1]
            st.caption("One Output page detected. Select which Order Form data row this page should use.")
        elif output_page_count > 1:
            page_mode = st.radio(
                "Artwork page selection",
                options=["All Pages", "Specific Page(s)"],
                horizontal=True,
                key=f"page_mode_{st.session_state['of_reset_id']}"
            )

            if page_mode == "All Pages":
                selected_pdf_pages = list(range(1, output_page_count + 1))
            else:
                selected_pdf_pages = st.multiselect(
                    "Select Output Page(s)",
                    options=list(range(1, output_page_count + 1)),
                    placeholder="Type or select page number(s)...",
                    key=f"selected_pdf_pages_{st.session_state['of_reset_id']}"
                )
                selected_pdf_pages = sorted(int(page) for page in selected_pdf_pages)

            st.caption(
                "Page numbers are the actual PDF page numbers. They are never renumbered after selection."
            )

        # -----------------------------------------------------
        # PAGE → ORDER FORM ROW MAPPING
        # -----------------------------------------------------
        if selected_pdf_pages:
            st.markdown(
                '<div class="section-title">🔗 Page → Order Form Row Mapping</div>',
                unsafe_allow_html=True
            )

            if len(df) == 1:
                for page_number in selected_pdf_pages:
                    page_row_mapping[int(page_number)] = 0
                st.success(
                    "✅ Excel contains one data row. Every selected PDF page will use Data Row 1 (Excel Row 2)."
                )
                mapping_ready = True

            else:
                st.caption(
                    "Each selected PDF page is mapped to an Order Form data row. "
                    "The default follows Page N → Data Row N, but you can change it."
                )

                mapping_labels = [
                    f"Data Row {i + 1}  (Excel Row {i + 2})"
                    for i in range(len(df))
                ]

                all_mapped = True
                for page_number in selected_pdf_pages:
                    natural_index = int(page_number) - 1
                    default_index = natural_index if natural_index < len(df) else None

                    selected_row_label = st.selectbox(
                        f"PDF Page {page_number} → Order Form Data Row",
                        options=mapping_labels,
                        index=default_index,
                        placeholder="Select an Order Form data row...",
                        key=(
                            f"page_row_map_{page_number}_"
                            f"{st.session_state['of_reset_id']}"
                        )
                    )

                    if selected_row_label:
                        selected_row_index = mapping_labels.index(selected_row_label)
                        page_row_mapping[int(page_number)] = selected_row_index
                    else:
                        all_mapped = False

                mapping_ready = bool(selected_pdf_pages) and all_mapped

                if mapping_ready:
                    st.success("✅ Page-to-row mapping is ready for comparison.")
                else:
                    st.info(
                        "Please select an Order Form data row for every selected PDF page before comparison."
                    )

        # =========================================================
        # COMPARISON METHOD
        # =========================================================
        comparison_method = None
        selected_fields = []

        st.divider()
        st.markdown(
            '<div class="section-title">⚙️ Comparison Method</div>',
            unsafe_allow_html=True
        )
        st.caption(
            "Choose how the Order Form data should be matched to the Output."
        )

        comparison_method = st.radio(
            "Comparison Method",
            options=["Auto Detect", "Select Fields"],
            index=None,
            horizontal=True,
            key=f"comparison_method_{st.session_state['of_reset_id']}"
        )

        available_fields = get_available_fields(df)

        if comparison_method == "Auto Detect":
            auto_key = (
                AUTO_DETECT_ENGINE_VERSION,
                str(getattr(excel_file, "name", "")),
                int(getattr(excel_file, "size", 0)),
                str(getattr(output_file, "name", "")),
                int(getattr(output_file, "size", 0)),
                product_type,
                tuple(sorted(page_row_mapping.items()))
            )

            auto_widget_key = f"auto_selected_fields_{st.session_state['of_reset_id']}"

            if not mapping_ready:
                st.info(
                    "Complete the Page → Order Form Row Mapping first. Auto Detect will then read only the mapped artwork page(s)."
                )
                detected_fields = []
            else:
                if st.session_state.get("of_auto_detect_key") != auto_key:
                    try:
                        with st.spinner("Auto Detect is reading the mapped artwork page(s) with OCR..."):
                            auto_pages = extract_output_pages(output_file)
                            detected_fields = auto_detect_fields(
                                df,
                                auto_pages,
                                product_type,
                                page_row_mapping=page_row_mapping
                            )
                        st.session_state["of_auto_detected_fields"] = detected_fields
                        st.session_state["of_auto_output_pages"] = auto_pages
                        st.session_state["of_auto_detect_key"] = auto_key
                        # Reset the editable Auto Detect selection only when the
                        # files/mapping actually change. Manual edits then remain
                        # untouched on subsequent Streamlit reruns.
                        st.session_state[auto_widget_key] = list(detected_fields)
                    except Exception as error:
                        st.error(
                            f"Unable to run Auto Detect: {type(error).__name__}: {error}"
                        )
                        with st.expander("Technical error details", expanded=False):
                            import traceback
                            st.code(traceback.format_exc())

                detected_fields = st.session_state.get("of_auto_detected_fields", [])

            st.markdown(
                '<div class="section-title">🤖 Auto Detected Fields</div>',
                unsafe_allow_html=True
            )
            st.caption(
                "Only fields with strong evidence in the mapped artwork are pre-selected. "
                "You can remove a detected field or add another populated field before comparison."
            )

            default_auto = [field for field in detected_fields if field in available_fields]
            # Session state is populated above only when Auto Detect input changes.
            # This keeps the field list editable without Streamlit overwriting the
            # user's add/remove choices on every rerun.
            selected_fields = st.multiselect(
                "Review detected fields",
                options=available_fields,
                placeholder="Type to search or add a populated field...",
                label_visibility="collapsed",
                key=auto_widget_key
            )

            st.caption(
                f"Auto Detect engine: {AUTO_DETECT_ENGINE_VERSION} • "
                f"{len(selected_fields)} field(s) currently selected"
            )

            if selected_fields:
                st.caption("Selected fields: " + ", ".join(selected_fields))
            else:
                st.warning(
                    "Auto Detect did not find high-confidence populated fields in the mapped artwork. "
                    "You can add fields directly from the dropdown."
                )

        elif comparison_method == "Select Fields":
            st.markdown(
                '<div class="section-title">Select Variable Fields to Validate</div>',
                unsafe_allow_html=True
            )
            st.caption(
                "Only populated Order Form fields are shown. Search directly inside the dropdown by typing the field name."
            )

            previous = st.session_state.get(
                f"selected_fields_{st.session_state['of_reset_id']}",
                []
            )
            previous = [field for field in previous if field in available_fields]

            selected_fields = st.multiselect(
                "Select the fields from your Order Form",
                options=available_fields,
                default=previous,
                placeholder="Type to search fields...",
                label_visibility="collapsed",
                key=f"selected_fields_{st.session_state['of_reset_id']}"
            )

            if selected_fields:
                preview_rows = []
                for field in selected_fields:
                    values = []
                    for value in df[field].tolist():
                        if is_blank_value(value):
                            continue
                        values.append(str(value).strip())
                    preview_rows.append({
                        "Excel Field": field,
                        "Values": len(values),
                        "Preview": " | ".join(values[:3])
                    })

                with st.expander("🔎 Preview Selected Fields"):
                    st.dataframe(
                        pd.DataFrame(preview_rows),
                        width="stretch",
                        hide_index=True
                    )
            else:
                st.info(
                    "No fields are currently selected. Click the field dropdown and type to search for a populated field."
                )

        # =========================================================
        # COMPARE BUTTON
        # =========================================================
        st.markdown("<br>", unsafe_allow_html=True)

        compare_ready = (
            comparison_method is not None
            and bool(selected_fields)
            and product_type != "----- SELECT -----"
            and mapping_ready
        )

        if st.button(
            "🔍  COMPARE & PROOFREAD",
            width="stretch",
            key=f"of_compare_button_{st.session_state['of_reset_id']}",
            disabled=not compare_ready
        ):
            try:
                with st.spinner("Reading output and preparing comparison..."):
                    if (
                        comparison_method == "Auto Detect"
                        and st.session_state.get("of_auto_output_pages")
                    ):
                        output_pages = st.session_state["of_auto_output_pages"]
                    else:
                        output_pages = extract_output_pages(output_file)

                    if not output_pages:
                        raise ValueError("No readable output pages were detected.")

                    if not selected_fields:
                        raise ValueError("No fields are selected for comparison.")

                    st.session_state["of_report"] = build_report(
                        df,
                        output_pages,
                        selected_fields,
                        product_type,
                        page_row_mapping=page_row_mapping
                    )
                    st.session_state["of_report_selected_fields"] = selected_fields
                    st.session_state["of_report_product_type"] = product_type
                    st.session_state["of_report_comparison_method"] = comparison_method
                    st.session_state["of_report_build_version"] = AUTO_DETECT_ENGINE_VERSION
                    st.session_state["of_visual_pages"] = [
                        page for page in output_pages
                        if int(page.get("page", 0)) in selected_pdf_pages
                    ]

            except Exception as error:
                import traceback
                st.error(
                    f"Unable to process the Output Artwork: {type(error).__name__}: {error}"
                )
                with st.expander("Technical error details", expanded=True):
                    st.code(traceback.format_exc())
    # =========================================================
    # SAVED REPORT
    # =========================================================
    report = st.session_state.get("of_report")

    if report is not None:
        st.divider()

        st.markdown(
            '<div class="section-title">QC Report</div>',
            unsafe_allow_html=True
        )

        report_product_type = st.session_state.get(
            "of_report_product_type",
            product_type
        )

        report_method = st.session_state.get(
            "of_report_comparison_method",
            "Select Fields"
        )

        report_fields = st.session_state.get(
            "of_report_selected_fields",
            []
        )

        st.caption(
            f"Comparison method: {report_method} • "
            f"Product type: {report_product_type}"
        )

        if report_method == "Auto Detect" and report_fields:
            st.caption(
                "Auto-detected fields: " + ", ".join(report_fields)
            )

        pass_count = int(
            (report["STATUS"] == "PASS").sum()
        )

        fail_count = int(
            (report["STATUS"] == "FAIL").sum()
        )

        not_found_count = int(
            (report["STATUS"] == "NOT FOUND").sum()
        )

        skip_count = int(
            (report["STATUS"] == "SKIP").sum()
        )

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric("PASS", pass_count)

        with col2:
            st.metric("FAIL", fail_count)

        with col3:
            st.metric("NOT FOUND", not_found_count)

        with col4:
            st.metric("IGNORED", skip_count)

        display_report = add_visual_column(report, report_fields)
        field_colors = get_field_visual_colors(report_fields)

        styled_report = (
            display_report
            .style
            .map(
                style_status,
                subset=["STATUS"]
            )
        )

        visual_styles = pd.DataFrame(
            "",
            index=display_report.index,
            columns=display_report.columns,
        )
        for row_index, field_name in display_report["FIELD"].items():
            if field_name in field_colors:
                visual_styles.loc[row_index, "VISUAL"] = (
                    f"background-color: {field_colors[field_name]};"
                    "color: white;"
                    "font-size: 17px;"
                    "font-weight: bold;"
                    "text-align: center;"
                )

        styled_report = styled_report.apply(
            lambda _df: visual_styles,
            axis=None,
        )

        st.dataframe(
            styled_report,
            width="stretch",
            hide_index=True
        )

        st.divider()

        # =========================================================
        # VISUAL ARTWORK COMPARISON
        # =========================================================
        visual_pages = st.session_state.get("of_visual_pages") or []
        if visual_pages:
            st.markdown(
                '<div class="section-title">🖼️ Visual Artwork Comparison</div>',
                unsafe_allow_html=True
            )
            st.caption(
                "Visual evidence is mapped from the same comparison result used for PASS/FAIL. "
                "Only numbered markers are placed on the artwork; field details remain outside the artwork."
            )

            page_numbers = [int(page.get("page", 0)) for page in visual_pages]
            if len(page_numbers) > 1:
                selected_visual_page = st.selectbox(
                    "Artwork Page",
                    options=page_numbers,
                    index=0,
                    format_func=lambda value: f"Artwork Page {value}",
                    key=f"of_visual_page_selector_{st.session_state['of_reset_id']}",
                )
            else:
                selected_visual_page = page_numbers[0]

            selected_visual_page_data = next(
                (page for page in visual_pages if int(page.get("page", 0)) == int(selected_visual_page)),
                visual_pages[0],
            )
            page_report = report[report["PDF PAGE"] == selected_visual_page] if "PDF PAGE" in report.columns else report.iloc[0:0]

            visual_zoom = st.slider(
                "Visual scale",
                min_value=50,
                max_value=150,
                value=90,
                step=5,
                format="%d%%",
                key=f"of_visual_zoom_{st.session_state['of_reset_id']}",
            )

            highlighted = build_highlighted_page_image(
                selected_visual_page_data,
                page_report,
                field_colors=field_colors,
            )

            visual_col, findings_col = st.columns([2.25, 1], gap="large")

            with visual_col:
                if highlighted is not None:
                    display_image = highlighted
                    if visual_zoom != 100:
                        new_width = max(200, int(display_image.width * visual_zoom / 100.0))
                        new_height = max(200, int(display_image.height * visual_zoom / 100.0))
                        display_image = display_image.resize(
                            (new_width, new_height),
                            Image.Resampling.LANCZOS,
                        )
                    st.image(display_image, width="stretch")

                    st.caption(
                        "Numbered markers correspond to the Field / Status entries in the Visual Findings panel."
                    )
                else:
                    st.warning("Visual evidence could not be rendered for this artwork page.")

            with findings_col:
                st.markdown("### Visual Findings")

                page_checks = []
                for _, row in page_report.iterrows():
                    status = str(row.get("STATUS", "")).strip().upper()
                    if status not in {"PASS", "FAIL"}:
                        continue
                    field = str(row.get("FIELD", "")).strip()
                    if not field:
                        continue

                    expected = str(row.get("ORDER FORM DATA", "") or "").strip()
                    actual = str(row.get("PDF OUTPUT", "") or "").strip()
                    evidence = _visual_find_field_boxes(
                        selected_visual_page_data,
                        field,
                        expected,
                        actual,
                        status,
                    )
                    page_checks.append({
                        "field": field,
                        "status": status,
                        "number": next((i + 1 for i, f in enumerate(report_fields) if f == field), ""),
                        "expected": expected,
                        "actual": actual,
                        "difference": str(row.get("DIFFERENCE", "") or "").strip(),
                        "has_visual": bool(evidence),
                    })

                if not page_checks:
                    st.info("No PASS/FAIL checks are associated with this page.")
                else:
                    # FAIL first, then PASS, while retaining field order within each status.
                    page_checks.sort(key=lambda item: (0 if item["status"] == "FAIL" else 1, item["number"]))

                    for item in page_checks:
                        field = item["field"]
                        status = item["status"]
                        rgb = field_colors.get(field, FIELD_VISUAL_COLORS[0])
                        status_label = "FAIL" if status == "FAIL" else "PASS"
                        marker = "🔴" if status == "FAIL" else "🟢"
                        evidence_label = "visual mapped" if item["has_visual"] else "visual location unavailable"

                        with st.container(border=True):
                            st.markdown(
                                f"**{item['number']}. {field}**  {marker} **{status_label}**",
                                unsafe_allow_html=False,
                            )
                            st.caption(evidence_label)

                            if status == "FAIL":
                                st.markdown(
                                    f"**Expected:** {item['expected'] or '—'}\n\n"
                                    f"**Found:** {item['actual'] or '—'}\n\n"
                                    f"**Difference:** {item['difference'] or '—'}"
                                )
                            else:
                                st.markdown(
                                    f"**Matched:** {item['actual'] or item['expected'] or '—'}"
                                )

                    mapped_count = sum(1 for item in page_checks if item["has_visual"])
                    st.caption(
                        f"Visual evidence coverage: {mapped_count}/{len(page_checks)} PASS/FAIL field(s) mapped on this page."
                    )

            # Compact legend below the viewer.
            legend_items = []
            for field in report_fields:
                if field in set(page_report.get("FIELD", [])):
                    color = field_colors.get(field, FIELD_VISUAL_COLORS[0])
                    legend_items.append(
                        f'<span style="display:inline-flex;align-items:center;margin-right:12px;">'
                        f'<span style="display:inline-block;width:10px;height:10px;border-radius:50%;'
                        f'background:{color};margin-right:5px;"></span>{field}</span>'
                    )
            if legend_items:
                st.markdown(
                    "<div style='margin-top:8px;line-height:1.8;'>" + "".join(legend_items) + "</div>",
                    unsafe_allow_html=True,
                )

        with st.expander("ℹ️ How this validation works"):
            st.write(
                """
                **Variable-data validation**

                Only the fields selected from the Order Form are treated as variable artwork data.

                **OCR validation**

                Artwork pages are rendered as images and OCR is used as the primary extraction source for non-editable artwork.

                **Combined output lines**

                Multiple Order Form fields can be matched independently when the artwork prints them on one line.

                **Page mapping**

                Each selected PDF page is mapped to a specific Order Form data row.

                By default, Page N → Data Row N (Excel Row N+1), but the mapping can be changed manually.

                When the Excel contains only one data row, every selected PDF page uses Data Row 1 (Excel Row 2).

                **PFL mode**

                Panel-numbered artwork is treated as a continuous stream so selected variable data can continue from one panel into the next panel.

                **Mismatch detection**

                If the selected Order Form value is present in the PDF → PASS.

                If the expected value is absent but a relevant alternative value is detected → FAIL.

                If an Order Form field is blank, that field is not required and is ignored.
                """
            )

        excel_data = create_excel_report(
            report=report,
            product_type=report_product_type,
            comparison_method=report_method,
            selected_fields=report_fields,
            visual_pages=visual_pages
        )

        st.download_button(
            label="⬇️ Download Excel QC Report",
            data=excel_data,
            file_name="PDF_Proofreading_QC_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
            key=f"of_download_excel_qc_report_{st.session_state['of_reset_id']}"
        )

    # =========================================================
    # INITIAL INSTRUCTIONS
    # =========================================================
    if not excel_file:
        st.caption("Upload an Order Form to begin.")
    elif not output_file:
        st.caption("Upload the Output Artwork to continue.")
    elif product_type == "----- SELECT -----":
        st.caption("Select the Product Type to continue.")
    elif comparison_method is None:
        st.caption("Select a Comparison Method to continue.")
    elif comparison_method == "Select Fields" and not selected_fields:
        st.caption("Select the variable fields you want to validate.")
