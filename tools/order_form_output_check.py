import streamlit as st
import pandas as pd
import fitz
import re
import unicodedata
import io
from io import BytesIO

from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

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
    if is_blank_value(text):
        return ""
    value = unicodedata.normalize("NFKC", str(text))
    value = value.replace("\u200b", "").replace("\ufeff", "")
    value = re.sub(r"\s+", "", value)
    return value.casefold()


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
    normalized = normalize_text(field_name).replace(" ", "")

    if (
        "_en" in original
        or normalized.endswith("en")
        or "english" in normalized
    ):
        return "EN"

    if (
        "_fr" in original
        or normalized.endswith("fr")
        or "french" in normalized
        or "canada" in normalized
    ):
        return "FR"

    if (
        "_sp" in original
        or normalized.endswith("sp")
        or "spanish" in normalized
        or "espanol" in normalized
    ):
        return "SP"

    return ""


def get_field_type(field_name):
    field = normalize_text(field_name)
    compact = field.replace(" ", "").replace("_", "").replace("-", "")

    # Explicit sequence fields such as OSZ1, OSZ2 ...
    if re.fullmatch(r"osz\d+", compact):
        return "OSZ"

    if "symbol" in compact or compact in {"caremark", "caresymbol", "washsymbol"}:
        return "SYMBOL"

    if (
        compact == "rn"
        or "rnno" in compact
        or "rnnumber" in compact
        or "registrationnumber" in compact
        or "companyrn" in compact
        or compact.startswith("rn")
    ):
        return "RN"

    if (
        "sku" in compact
        or "itemcode" in compact
        or "itemnumber" in compact
        or "itemno" in compact
        or "stylecode" in compact
        or compact == "style"
        or "productcode" in compact
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

    if (
        "coo" in compact
        or "countryoforigin" in compact
        or "countryorigin" in compact
        or "madein" in compact
        or compact == "origin"
    ):
        return "COO"

    if (
        "fiber" in compact
        or "fibre" in compact
        or "fabric" in compact
        or "content" in compact
        or "composition" in compact
        or "fabrication" in compact
        or "material" in compact
    ):
        return "CONTENT"

    if (
        "care" in compact
        or "wash" in compact
        or "washing" in compact
        or "laundry" in compact
        or "instruction" in compact
    ):
        return "CARE"

    if (
        "size" in compact
        or "sizeline" in compact
        or "alpha" in compact
        or "waist" in compact
        or "inseam" in compact
        or compact == "fit"
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


def _ocr_image(image):
    try:
        import pytesseract
    except ImportError as exc:
        raise RuntimeError(
            "OCR support is not installed. Add pytesseract to requirements.txt."
        ) from exc

    try:
        return pytesseract.image_to_string(
            image,
            lang="eng+fra+spa"
        )
    except Exception as exc:
        raise RuntimeError(
            "OCR could not run. Make sure Tesseract OCR and its language data "
            "are installed in the Streamlit environment."
        ) from exc


def _render_pdf_page(page):
    pixmap = page.get_pixmap(
        matrix=fitz.Matrix(2.5, 2.5),
        alpha=False
    )
    return Image.open(
        io.BytesIO(pixmap.tobytes("png"))
    ).convert("RGB")


def get_output_page_count(file):
    name = str(getattr(file, "name", "")).casefold()

    if name.endswith(".pdf"):
        file.seek(0)
        data = file.read()
        document = fitz.open(stream=data, filetype="pdf")
        count = len(document)
        document.close()
        return count

    if name.endswith((".jpg", ".jpeg", ".png")):
        return 1

    return 0


def extract_output_pages(file):
    name = str(getattr(file, "name", "")).casefold()

    if name.endswith(".pdf"):
        file.seek(0)
        data = file.read()
        document = fitz.open(stream=data, filetype="pdf")
        pages = []

        for page_number, page in enumerate(document, start=1):
            direct_text = page.get_text("text") or ""

            if _usable_text(direct_text):
                text = direct_text
                source_type = "pdf_text"
            else:
                text = _ocr_image(_render_pdf_page(page))
                source_type = "ocr"

            pages.append({
                "page": page_number,
                "text": text,
                "source_type": source_type
            })

        document.close()
        return pages

    if name.endswith((".jpg", ".jpeg", ".png")):
        file.seek(0)
        image = Image.open(file).convert("RGB")
        return [{
            "page": 1,
            "text": _ocr_image(image),
            "source_type": "ocr"
        }]

    raise ValueError(
        "Unsupported output format. Please upload PDF, JPG, JPEG, or PNG."
    )


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


def build_page_state(page, product_type):
    lines = build_page_lines(
        page.get("text", ""),
        product_type
    )
    return {
        "page": page.get("page"),
        "source_type": page.get("source_type", "pdf_text"),
        "lines": lines,
        "consumed": set(),
        "osz_runs": extract_standalone_numeric_runs_from_lines(lines),
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


def line_is_available(line, state):
    return line["line_id"] not in state["consumed"]


def consume_lines(state, lines):
    for line in lines:
        state["consumed"].add(line["line_id"])


def join_lines(lines):
    return " ".join(
        line["text"]
        for line in lines
    ).strip()


# =========================================================
# SAFE EXACT MATCHING
# =========================================================

def normalized_contains(expected, actual):
    expected_norm = normalize_text(expected)
    actual_norm = normalize_text(actual)

    if not expected_norm or not actual_norm:
        return False

    # IMPORTANT: for multi-word text, substring is acceptable because artwork
    # may include a label around it. For a single numeric token, this function
    # must NOT be used because 2 must never match inside RN# 55285.
    if len(expected_norm.split()) <= 1 and normalize_numeric(expected_norm):
        return False

    return expected_norm in actual_norm


def safe_exact_match(expected, actual, field_name):
    field_type = get_field_type(field_name)

    if field_type == "SYMBOL":
        exp = normalize_symbol_text(expected)
        act = normalize_symbol_text(actual)
        return bool(exp and exp == act)

    # Numeric value: exact numeric token only.
    numeric_expected = normalize_numeric(expected)
    if numeric_expected is not None:
        actual_norm = normalize_text(actual)
        actual_numbers = re.findall(
            r"(?<![A-Za-z0-9])[-+]?\d+(?:\.\d+)?(?![A-Za-z0-9])",
            actual_norm
        )
        return numeric_expected in {
            normalize_numeric(number)
            for number in actual_numbers
        }

    exp = normalize_text(expected)
    act = normalize_text(actual)

    if not exp or not act:
        return False

    # Full normalized text / token sequence.
    if exp == act:
        return True

    if exp in act:
        return True

    # Space-independent comparison is only for longer textual values.
    compact_exp = compact_text(expected)
    compact_act = compact_text(actual)

    return bool(
        len(exp.split()) > 1
        and compact_exp
        and compact_exp in compact_act
    )


def find_exact_lines(expected, field_name, state, max_window=8):
    """Find the smallest unused contiguous line span containing the expected value."""
    lines = state["lines"]
    available = [line for line in lines if line_is_available(line, state)]
    if not available:
        return None

    field_type = get_field_type(field_name)
    numeric_expected = normalize_numeric(expected)

    # Numeric scalar values must never be matched as substrings in mixed text.
    # This is the critical protection against 2/4/6 matching inside RN# 55285.
    if numeric_expected is not None:
        search_window = 1
        for line in available:
            if field_type in {"OSZ", "RN"}:
                continue
            if field_type not in {"QUANTITY", "BATCH"} and normalize_text(line["text"]) == numeric_expected:
                return [line]

        return None

    preferred_single_line = field_type in {
        "SYMBOL", "RN", "IDENTIFIER", "OSZ", "SIZE", "COLOR",
        "GENDER", "BATCH", "QUANTITY"
    }
    search_window = 1 if preferred_single_line else max_window

    for window_size in range(1, search_window + 1):
        for start in range(0, len(available) - window_size + 1):
            candidate = available[start:start + window_size]
            ids = [line["line_id"] for line in candidate]

            if ids != list(range(ids[0], ids[-1] + 1)):
                continue

            text = join_lines(candidate)

            if safe_exact_match(expected, text, field_name):
                return candidate

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


def extract_rn_value(text):
    normalized = normalize_text(text)
    if not normalized:
        return None

    match = re.search(
        r"\b(?:rn|ca)\s*[#:.-]?\s*([0-9][0-9a-z\-\/]*)",
        normalized,
        re.IGNORECASE
    )
    return match.group(1) if match else None


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


def extract_content_values(text):
    normalized = normalize_text(text)
    if not normalized:
        return []

    # Stop material capture at another percentage sign or common section label.
    pattern = re.compile(
        r"(?P<pct>\d{1,3}(?:\.\d+)?)\s*%\s*"
        r"(?P<material>[a-z][a-z0-9\s\-]*?)(?=\s+\d{1,3}(?:\.\d+)?\s*%|$)",
        re.IGNORECASE
    )

    values = []
    for match in pattern.finditer(normalized):
        material = match.group("material").strip()
        material = re.sub(
            r"\b(?:shell|liner|lining|body|rn|ca|made|size|color|colour)\b.*$",
            "",
            material,
        ).strip()
        if material:
            values.append(
                f"{match.group('pct')}% {material}"
            )
    return values


def normalize_composition(values):
    return sorted(
        normalize_text(value)
        for value in values
        if normalize_text(value)
    )


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

    # Stable sequence mapping. The run is captured before any OSZ field is
    # consumed, so OSZ2 always remains the second original sequence value.
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


# =========================================================
# FIELD CHECKERS
# =========================================================

def find_care_region(state, field_name):
    region = get_field_region(field_name)
    marker_sets = {
        "EN": ["machine wash", "wash", "bleach", "dry clean", "tumble dry", "cool iron"],
        "FR": ["laver", "blanchiment", "nettoyage", "sécher", "repasser"],
        "SP": ["lavar", "cloro", "secadora", "plancha", "limpieza en seco"],
        "": ["machine wash", "wash", "bleach", "dry clean", "laver", "lavar"],
    }
    markers = [normalize_text(m) for m in marker_sets.get(region, marker_sets[""]) if normalize_text(m)]

    lines = state["lines"]
    start = None
    for idx, line in enumerate(lines):
        if not line_is_available(line, state):
            continue
        if any(marker in line["norm"] for marker in markers):
            start = idx
            break

    if start is None:
        return None

    region_lines = []
    for idx in range(start, len(lines)):
        line = lines[idx]
        if not line_is_available(line, state):
            break

        if idx > start:
            if extract_coo_value(line["text"]) or extract_rn_value(line["text"]):
                break
            if extract_content_values(line["text"]):
                break

            # A non-text/symbol line, an isolated number, or a new obvious
            # technical line ends the care region. This is essential so symbols,
            # RN and OSZ data are not swallowed by the care matcher.
            ascii_letters = len(re.findall(r"[A-Za-z]", line["text"]))
            if ascii_letters == 0:
                break

            if re.fullmatch(r"\s*[-+]?\d+(?:\.\d+)?\s*", line["text"]):
                break

        region_lines.append(line)
        if len(region_lines) >= 20:
            break

    return region_lines or None


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
    # Symbol: exact symbol/text only. Never substring-match.
    # -----------------------------------------------------
    if field_type == "SYMBOL":
        for line in state["lines"]:
            if not line_is_available(line, state):
                continue

            if normalize_symbol_text(expected) == normalize_symbol_text(line["text"]):
                consume_lines(state, [line])
                return {
                    "status": "PASS",
                    "pdf": line["text"],
                    "difference": "—",
                    "match_type": "SYMBOL_EXACT"
                }

        return {
            "status": "NOT FOUND",
            "pdf": "Not found",
            "difference": "Symbol/value was not detected exactly.",
            "match_type": "NOT_FOUND"
        }

    # -----------------------------------------------------
    # CONTENT: parse the composition before any generic text check.
    # -----------------------------------------------------
    if field_type == "CONTENT":
        expected_values = normalize_composition(
            extract_content_values(expected)
        )

        if expected_values:
            # Search contiguous windows, but only over unconsumed lines.
            available = [
                line for line in state["lines"]
                if line_is_available(line, state)
            ]

            max_window = min(8, len(available))

            for size in range(1, max_window + 1):
                for start in range(0, len(available) - size + 1):
                    candidate = available[start:start + size]
                    ids = [line["line_id"] for line in candidate]
                    if ids != list(range(ids[0], ids[-1] + 1)):
                        continue

                    text = join_lines(candidate)
                    actual_values = normalize_composition(
                        extract_content_values(text)
                    )

                    if expected_values == actual_values:
                        consume_lines(state, candidate)
                        return {
                            "status": "PASS",
                            "pdf": text,
                            "difference": "—",
                            "match_type": "CONTENT_EXACT"
                        }

            # If we have a relevant composition region with overlapping
            # components, report the actual composition as FAIL.
            for line in state["lines"]:
                if not line_is_available(line, state):
                    continue
                text = line["text"]
                actual_values = normalize_composition(
                    extract_content_values(text)
                )
                if not actual_values:
                    continue
                if set(expected_values) & set(actual_values):
                    consume_lines(state, [line])
                    return {
                        "status": "FAIL",
                        "pdf": text,
                        "difference": (
                            f"Expected: {expected} | "
                            f"Found: {' | '.join(actual_values)}"
                        ),
                        "match_type": "CONTENT_MISMATCH"
                    }

            return {
                "status": "NOT FOUND",
                "pdf": "Not found",
                "difference": "Expected composition was not detected.",
                "match_type": "NOT_FOUND"
            }

    # -----------------------------------------------------
    # RN
    # -----------------------------------------------------
    if field_type == "RN":
        expected_rn = extract_rn_value(expected)
        expected_target = normalize_text(
            expected_rn if expected_rn else expected
        )

        for line in state["lines"]:
            if not line_is_available(line, state):
                continue

            actual_rn = extract_rn_value(line["text"])
            if actual_rn is None:
                continue

            if normalize_text(actual_rn) == expected_target:
                consume_lines(state, [line])
                return {
                    "status": "PASS",
                    "pdf": line["text"],
                    "difference": "—",
                    "match_type": "RN_EXACT"
                }

            consume_lines(state, [line])
            return {
                "status": "FAIL",
                "pdf": line["text"],
                "difference": (
                    f"Expected: {expected} | "
                    f"Found: RN# {actual_rn}"
                ),
                "match_type": "RN_MISMATCH"
            }

        return {
            "status": "NOT FOUND",
            "pdf": "Not found",
            "difference": "RN value was not detected.",
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

        # Prefer labeled size blocks.
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

        # Then exact text, useful when artwork prints only the value.
        exact = find_exact_lines(
            expected,
            field_name,
            state,
            max_window=1
        )
        if exact:
            consume_lines(state, exact)
            return {
                "status": "PASS",
                "pdf": join_lines(exact),
                "difference": "—",
                "match_type": "EXACT"
            }

        return {
            "status": "NOT FOUND",
            "pdf": "Not found",
            "difference": "Size value was not detected.",
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
    # GENERAL/BRAND/ATTRIBUTE/IDENTIFIER/BATCH/QUANTITY
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
        consume_lines(state, exact)
        return {
            "status": "PASS",
            "pdf": join_lines(exact),
            "difference": "—",
            "match_type": "EXACT"
        }

    # IMPORTANT: for unknown fields we do not manufacture a FAIL. This avoids
    # the historical problem where unrelated PDF text was interpreted as a mismatch.
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
    "RN": 20,
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
    product_type
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

        excel_index = page_index

        if excel_index >= len(df):
            for field in selected_fields:
                results.append({
                    "FIELD NO": field_no,
                    "PDF PAGE": page["page"],
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
                "PDF PAGE": page["page"],
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

def auto_detect_fields(
    df,
    output_pages,
    product_type
):
    """
    Auto Detect is now a REPORT-DRIVEN operation.

    It considers every populated non-administrative field and runs the same
    validation engine used by manual Select Fields.

    Therefore Auto Detect can detect:
        PASS -> expected data found
        FAIL -> relevant field-specific wrong data found

    NOT FOUND is not enough evidence to call a field an artwork field.
    """

    available_fields = get_available_fields(df)

    candidates = [
        field for field in available_fields
        if not is_admin_field(field)
    ]

    if not candidates:
        return []

    probe = build_report(
        df,
        output_pages,
        order_fields_for_matching(candidates),
        product_type
    )

    if probe.empty:
        return []

    detected = []

    for field in candidates:
        field_rows = probe[
            probe["FIELD"] == field
        ]

        if field_rows.empty:
            continue

        if field_rows["STATUS"].isin(
            ["PASS", "FAIL"]
        ).any():
            detected.append(field)

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



def create_excel_report(report, product_type, comparison_method, selected_fields):
    """Create a professional Excel QC report with Summary and Field Comparison sheets."""

    output = BytesIO()
    wb = Workbook()

    # ---------------------------------------------------------
    # COLORS
    # ---------------------------------------------------------
    NAVY = "1F4E78"
    LIGHT_BLUE = "D9EAF7"
    GREEN = "238636"
    RED = "DA3633"
    AMBER = "9E6A03"
    GREY = "555555"
    WHITE = "FFFFFF"
    LIGHT_BORDER = "D9E1F2"

    # ---------------------------------------------------------
    # COUNTS
    # ---------------------------------------------------------
    pass_count = int((report["STATUS"] == "PASS").sum()) if not report.empty else 0
    fail_count = int((report["STATUS"] == "FAIL").sum()) if not report.empty else 0
    not_found_count = int((report["STATUS"] == "NOT FOUND").sum()) if not report.empty else 0
    skip_count = int((report["STATUS"] == "SKIP").sum()) if not report.empty else 0
    total_checks = len(report)

    if fail_count > 0:
        overall = "FAIL"
    elif not_found_count > 0:
        overall = "REVIEW"
    else:
        overall = "PASS"

    # =========================================================
    # SUMMARY
    # =========================================================
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
        ("Overall Result", overall),
    ]

    for r, (label, value) in enumerate(summary, start=4):
        ws.cell(r, 1, label)
        ws.cell(r, 2, value)
        ws.cell(r, 1).font = Font(bold=True)
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=LIGHT_BLUE)
        ws.cell(r, 1).alignment = Alignment(vertical="center")
        ws.cell(r, 2).alignment = Alignment(vertical="center", wrap_text=True)

    result_cell = ws["B12"]
    result_color = GREEN if overall == "PASS" else RED if overall == "FAIL" else AMBER
    result_cell.fill = PatternFill("solid", fgColor=result_color)
    result_cell.font = Font(bold=True, color=WHITE)
    result_cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 28
    ws.freeze_panes = "A4"

    # Small finding section
    finding_row = 15
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

    # =========================================================
    # FIELD COMPARISON
    # =========================================================
    comparison = wb.create_sheet("Field Comparison")

    headers = list(report.columns)
    header_fill = PatternFill("solid", fgColor=NAVY)

    for col_idx, header in enumerate(headers, start=1):
        cell = comparison.cell(1, col_idx, header)
        cell.font = Font(bold=True, color=WHITE)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_idx, row in enumerate(report.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row, start=1):
            cell = comparison.cell(row_idx, col_idx, value)
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    # Status colors
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

    # Table
    if comparison.max_row >= 2 and comparison.max_column >= 1:
        ref = f"A1:{get_column_letter(comparison.max_column)}{comparison.max_row}"
        table = Table(displayName="QCFieldComparison", ref=ref)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        comparison.add_table(table)

    # Widths
    width_map = {
        "FIELD NO": 12,
        "PDF PAGE": 12,
        "EXCEL ROW": 12,
        "FIELD": 24,
        "ORDER FORM DATA": 42,
        "PDF OUTPUT": 52,
        "STATUS": 16,
        "DIFFERENCE": 58,
        "MATCH TYPE": 22,
    }

    for col_idx, header in enumerate(headers, start=1):
        letter = get_column_letter(col_idx)
        comparison.column_dimensions[letter].width = width_map.get(header, 20)

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
            st.rerun()

    with nav_right:
        if st.button(
            "🆕 NEW START",
            key="of_new_start",
            width="stretch"
        ):
            st.session_state["of_reset_id"] += 1
            st.session_state["of_report"] = None
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
    # COMPARISON METHOD
    # =========================================================
    comparison_method = None
    selected_fields = []

    if excel_file and output_file:
        st.divider()

        st.markdown(
            '<div class="section-title">⚙️ Comparison Method</div>',
            unsafe_allow_html=True
        )

        st.caption(
            "Choose how the Order Form data should be matched to the Output. "
            "Nothing is selected automatically."
        )

        comparison_method = st.radio(
            "Comparison Method",
            options=["Auto Detect", "Select Fields"],
            index=None,
            horizontal=True,
            key=f"comparison_method_{st.session_state['of_reset_id']}"
        )

        if comparison_method == "Select Fields":
            st.markdown(
                '<div class="section-title">'
                'Select Variable Fields to Validate'
                '</div>',
                unsafe_allow_html=True
            )

            st.caption(
                "Only populated Order Form fields are shown. Only the fields selected below will participate in the comparison."
            )

            available_fields = get_available_fields(df)

            selected_fields = st.multiselect(
                "Select the fields from your Order Form",
                options=available_fields,
                default=[],
                label_visibility="collapsed",
                key=f"selected_fields_{st.session_state['of_reset_id']}"
            )

            if selected_fields:
                preview_rows = []

                for field in selected_fields:
                    values = []

                    for value in df[field].tolist():
                        if pd.isna(value):
                            continue

                        value = str(value).strip()

                        if value:
                            values.append(value)

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
                    "Select at least one Order Form field to continue."
                )

    # =========================================================
    # FILE INFORMATION
    # Page count is metadata only. Output text/OCR is NOT extracted here.
    # =========================================================
    if excel_file and output_file:
        try:
            output_page_count = get_output_page_count(output_file)
        except Exception as error:
            output_page_count = 0
            st.warning(
                f"Unable to determine output page count: {error}"
            )

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

        if output_page_count and len(df) != output_page_count:
            st.warning(
                "⚠️ Excel row count and output page count do not have the same "
                "count. The existing mapping will still use Output Page 1 → "
                "Excel Row 2, Output Page 2 → Excel Row 3, and so on."
            )
        elif output_page_count:
            st.success("✅ Excel rows and output pages match.")

    # =========================================================
    # COMPARE BUTTON
    # Extraction and comparison start ONLY after COMPARE.
    # =========================================================
    if excel_file and output_file:
        st.markdown("<br>", unsafe_allow_html=True)

        compare_ready = (
            comparison_method is not None
            and (
                comparison_method == "Auto Detect"
                or bool(selected_fields)
            )
            and product_type != "----- SELECT -----"
        )

        if st.button(
            "🔍  COMPARE & PROOFREAD",
            width="stretch",
            key=f"of_compare_button_{st.session_state['of_reset_id']}",
            disabled=not compare_ready
        ):
            try:
                with st.spinner(
                    "Reading output and preparing comparison..."
                ):
                    # =====================================================
                    # EXTRACTION LAYER
                    # =====================================================
                    output_pages = extract_output_pages(output_file)

                    if not output_pages:
                        raise ValueError(
                            "No readable output pages were detected."
                        )

                    # =====================================================
                    # COMPARISON LAYER
                    # =====================================================
                    if comparison_method == "Auto Detect":
                        detected_fields = auto_detect_fields(
                            df,
                            output_pages,
                            product_type
                        )

                        if not detected_fields:
                            st.session_state["of_report"] = pd.DataFrame([
                                {
                                    "FIELD NO": 1,
                                    "PDF PAGE": "—",
                                    "EXCEL ROW": "—",
                                    "FIELD": "Auto Detect",
                                    "ORDER FORM DATA": "—",
                                    "PDF OUTPUT": "—",
                                    "STATUS": "NOT FOUND",
                                    "DIFFERENCE": (
                                        "No relevant Order Form fields could be "
                                        "reliably associated with the output."
                                    )
                                }
                            ])
                        else:
                            st.session_state["of_report"] = build_report(
                                df,
                                output_pages,
                                detected_fields,
                                product_type
                            )

                        st.session_state["of_report_selected_fields"] = (
                            detected_fields
                        )

                    else:
                        st.session_state["of_report"] = build_report(
                            df,
                            output_pages,
                            selected_fields,
                            product_type
                        )

                        st.session_state["of_report_selected_fields"] = (
                            selected_fields
                        )

                    st.session_state["of_report_product_type"] = product_type
                    st.session_state["of_report_comparison_method"] = (
                        comparison_method
                    )

            except Exception as error:
                st.error(
                    f"Unable to process the Output Artwork: {error}"
                )

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

        styled_report = (
            report
            .style
            .map(
                style_status,
                subset=["STATUS"]
            )
        )

        st.dataframe(
            styled_report,
            width="stretch",
            hide_index=True
        )

        st.divider()

        if fail_count > 0:
            st.error(
                f"❌ FAIL — {fail_count} variable-data mismatch(es) detected."
            )
        elif not_found_count > 0:
            st.warning(
                f"⚠️ REVIEW — {not_found_count} selected variable "
                f"field(s) could not be located."
            )
        else:
            st.success(
                "✅ PASS — All selected variable fields matched the PDF artwork."
            )

        with st.expander("ℹ️ How this validation works"):
            st.write(
                """
                **Variable-data validation**

                Only the fields selected from the Order Form are treated as
                variable artwork data.

                **Static PDF content is ignored.**

                PDF bullets/keystrokes such as `n`, regional prefixes,
                addresses, phone numbers and other unselected static artwork
                content do not create failures.

                **Page mapping**

                PDF Page 1 → Excel Row 2

                PDF Page 2 → Excel Row 3

                PDF Page 3 → Excel Row 4

                and so on.

                **PFL mode**

                Panel-numbered artwork is treated as a continuous stream so
                selected variable data can continue from one panel into the
                next panel.

                **Mismatch detection**

                If the selected Order Form value is present in the PDF → PASS.

                If the expected value is absent but a relevant alternative
                value is detected → FAIL.

                If an Order Form field is blank, that field is not required
                and is ignored.
                """
            )

        excel_data = create_excel_report(
            report=report,
            product_type=report_product_type,
            comparison_method=report_method,
            selected_fields=report_fields
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
