import streamlit as st
import pandas as pd
import fitz
import re
import unicodedata
import io
import hashlib
import textwrap
from io import BytesIO

from PIL import Image, ImageDraw, ImageOps, ImageEnhance, ImageFilter, ImageFont
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


def _tesseract_data_pass(image, lang="eng", config="--psm 6", offset_x=0, offset_y=0, scale_back=1.0):
    """Run one OCR pass and return normalized word records in original-image coordinates."""
    import pytesseract
    from pytesseract import Output

    data = pytesseract.image_to_data(
        image,
        lang=lang,
        output_type=Output.DICT,
        config=config,
    )

    words = []
    texts = data.get("text", [])
    count = len(texts)

    for i in range(count):
        raw_text = str(texts[i] or "").strip()
        if not raw_text:
            continue

        try:
            conf = float(data.get("conf", ["-1"] * count)[i])
        except Exception:
            conf = -1.0

        try:
            left = int(data.get("left", [0] * count)[i])
            top = int(data.get("top", [0] * count)[i])
            width = int(data.get("width", [0] * count)[i])
            height = int(data.get("height", [0] * count)[i])
        except Exception:
            continue

        # A resized crop is mapped back to the original full-page coordinates.
        if scale_back != 1.0:
            left = int(left / scale_back)
            top = int(top / scale_back)
            width = max(1, int(width / scale_back))
            height = max(1, int(height / scale_back))

        words.append({
            "text": raw_text,
            "left": int(left + offset_x),
            "top": int(top + offset_y),
            "width": width,
            "height": height,
            "conf": conf,
            "block_num": int(data.get("block_num", [0] * count)[i]),
            "par_num": int(data.get("par_num", [0] * count)[i]),
            "line_num": int(data.get("line_num", [0] * count)[i]),
        })

    return words


def _deduplicate_ocr_words(words):
    """Merge duplicate detections from overlapping OCR passes."""
    best = {}

    for word in words:
        text_norm = _visual_norm(word.get("text", ""))
        if not text_norm:
            continue

        left = int(word.get("left", 0))
        top = int(word.get("top", 0))
        width = max(1, int(word.get("width", 1)))
        height = max(1, int(word.get("height", 1)))

        # Spatial bucketing allows tiny coordinate differences between OCR passes.
        key = (
            text_norm,
            round(left / 12),
            round(top / 12),
        )

        current = best.get(key)
        if current is None or float(word.get("conf", -1)) > float(current.get("conf", -1)):
            best[key] = word

    result = list(best.values())
    result.sort(key=lambda item: (
        int(item.get("top", 0)),
        int(item.get("left", 0)),
    ))
    return result


def _ocr_words_to_text(words):
    """Build readable OCR lines from word boxes using their physical Y positions."""
    if not words:
        return ""

    ordered = sorted(
        words,
        key=lambda item: (
            int(item.get("top", 0)),
            int(item.get("left", 0)),
        ),
    )

    lines = []
    current = []
    current_y = None
    current_height = 0

    for word in ordered:
        top = int(word.get("top", 0))
        height = max(1, int(word.get("height", 1)))
        center_y = top + height / 2

        if current_y is None:
            current = [word]
            current_y = center_y
            current_height = height
            continue

        tolerance = max(8, int(max(current_height, height) * 0.65))
        if abs(center_y - current_y) <= tolerance:
            current.append(word)
            current_y = (current_y * (len(current) - 1) + center_y) / len(current)
            current_height = max(current_height, height)
        else:
            current.sort(key=lambda item: int(item.get("left", 0)))
            lines.append(" ".join(str(item.get("text", "")) for item in current).strip())
            current = [word]
            current_y = center_y
            current_height = height

    if current:
        current.sort(key=lambda item: int(item.get("left", 0)))
        lines.append(" ".join(str(item.get("text", "")) for item in current).strip())

    return "\n".join(line for line in lines if line)


def _prepare_ocr_variant(image, mode):
    """Create OCR-friendly variants without changing the final coordinate system."""
    base = image.convert("RGB")

    if mode == "gray":
        return ImageOps.grayscale(base).convert("RGB")

    if mode == "contrast":
        gray = ImageOps.grayscale(base)
        gray = ImageEnhance.Contrast(gray).enhance(1.8)
        gray = ImageEnhance.Sharpness(gray).enhance(1.4)
        return gray.convert("RGB")

    if mode == "sharp":
        return base.filter(ImageFilter.SHARPEN)

    return base


def _ocr_image_with_data(image):
    """
    OCR 2.0 pipeline for artwork QC.

    The page is processed using several OCR views plus overlapping enlarged
    horizontal crops. Crop OCR is mapped back to the original page coordinates,
    allowing the comparison engine to locate tiny artwork text while preserving
    the complete original artwork for visual highlighting.
    """
    try:
        import pytesseract
    except ImportError as exc:
        raise RuntimeError(
            "OCR support is not installed. Add pytesseract to requirements.txt."
        ) from exc

    errors = []
    all_words = []

    # We deliberately try English first. If that works, optional language packs
    # are not required for ordinary apparel artwork.
    languages = ["eng"]
    selected_lang = "eng"

    # Full-page passes. The rendered page is already high resolution, so these
    # preserve the complete layout and provide a strong baseline.
    full_passes = [
        ("normal", "--psm 6"),
        ("normal", "--psm 11"),
        ("contrast", "--psm 11"),
    ]

    for mode, config in full_passes:
        try:
            variant = _prepare_ocr_variant(image, mode)
            all_words.extend(
                _tesseract_data_pass(
                    variant,
                    lang="eng",
                    config=config,
                )
            )
        except Exception as exc:
            errors.append(f"full/{mode}/{config}: {exc}")

    # Region OCR is important for wide artwork labels. The small variable-data
    # block often becomes much easier for Tesseract when surrounding graphics
    # are removed. We therefore OCR four overlapping quadrants plus a broad
    # lower-page strip. All crops are enlarged before OCR and then mapped back
    # into the original full-page coordinate system.
    width, height = image.size
    x_mid = int(width * 0.50)
    y_mid = int(height * 0.52)
    x_overlap = int(width * 0.12)
    y_overlap = int(height * 0.12)

    regions = [
        # left, top, right, bottom
        (0, 0, min(width, x_mid + x_overlap), min(height, y_mid + y_overlap)),
        (max(0, x_mid - x_overlap), 0, width, min(height, y_mid + y_overlap)),
        (0, max(0, y_mid - y_overlap), min(width, x_mid + x_overlap), height),
        (max(0, x_mid - x_overlap), max(0, y_mid - y_overlap), width, height),
        (0, int(height * 0.45), width, height),
    ]

    for region_index, (left, top, right, bottom) in enumerate(regions):
        if right <= left or bottom <= top:
            continue

        crop = image.crop((left, top, right, bottom))

        scale = 1.65
        enlarged = crop.resize(
            (
                max(1, int(crop.width * scale)),
                max(1, int(crop.height * scale)),
            ),
            Image.Resampling.LANCZOS,
        )

        for mode, config in (
            ("normal", "--psm 11"),
            ("contrast", "--psm 11"),
        ):
            try:
                variant = _prepare_ocr_variant(enlarged, mode)
                all_words.extend(
                    _tesseract_data_pass(
                        variant,
                        lang="eng",
                        config=config,
                        offset_x=left,
                        offset_y=top,
                        scale_back=scale,
                    )
                )
            except Exception as exc:
                errors.append(
                    f"region/{region_index}/{mode}/{config}: {exc}"
                )

    words = _deduplicate_ocr_words(all_words)

    # If English gave us nothing useful, try the broader language set once.
    text = _ocr_words_to_text(words)
    if not _usable_text(text):
        try:
            broader_words = []
            for mode, config in (
                ("normal", "--psm 6"),
                ("contrast", "--psm 11"),
            ):
                variant = _prepare_ocr_variant(image, mode)
                broader_words.extend(
                    _tesseract_data_pass(
                        variant,
                        lang="eng+fra+spa",
                        config=config,
                    )
                )

            words = _deduplicate_ocr_words(broader_words)
            text = _ocr_words_to_text(words)
            if text:
                selected_lang = "eng+fra+spa"
        except Exception as exc:
            errors.append(f"eng+fra+spa: {exc}")

    if not text:
        if errors:
            raise RuntimeError(
                "OCR could not produce usable text. Details: "
                + " | ".join(errors[:10])
            )
        return "", [], selected_lang

    return text, words, selected_lang


def _ocr_image(image):
    text, _words, _lang = _ocr_image_with_data(image)
    return text


def _render_pdf_page(page):
    # Higher internal resolution is intentional. The displayed visual image is
    # scaled down separately in Streamlit, so this improves OCR without making
    # the screen image unnecessarily large.
    pixmap = page.get_pixmap(
        matrix=fitz.Matrix(8.0, 8.0),
        alpha=False,
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
                ocr_text = ""
                ocr_words = []
                ocr_error = None
                ocr_lang = ""

                try:
                    ocr_text, ocr_words, ocr_lang = _ocr_image_with_data(image)
                except Exception as exc:
                    ocr_error = exc

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
                    "ocr_words": ocr_words,
                    "ocr_lang": ocr_lang,
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
        ocr_text, ocr_words, ocr_lang = _ocr_image_with_data(image)
        if not _usable_text(ocr_text):
            raise RuntimeError("No readable artwork text was detected in the image.")
        file.seek(0)
        return [{
            "page": 1,
            "text": str(ocr_text),
            "source_type": "ocr",
            "direct_text": "",
            "ocr_text": str(ocr_text),
            "ocr_words": ocr_words,
            "ocr_lang": ocr_lang,
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

# =========================================================
# VISUAL ARTWORK HIGHLIGHTING
# =========================================================

VISUAL_FIELD_PALETTE = [
    (30, 136, 229),
    (156, 39, 176),
    (0, 150, 136),
    (255, 152, 0),
    (76, 175, 80),
    (233, 30, 99),
    (121, 85, 72),
    (63, 81, 181),
    (0, 188, 212),
    (255, 193, 7),
    (244, 67, 54),
    (67, 160, 71),
    (94, 53, 177),
    (3, 169, 244),
    (255, 87, 34),
    (0, 121, 107),
    (117, 117, 117),
    (205, 92, 92),
    (46, 125, 50),
    (123, 31, 162),
    (2, 119, 189),
    (239, 108, 0),
    (0, 105, 92),
    (173, 20, 87),
]


def _field_visual_color(field_name):
    """Give every field a deterministic, reusable annotation color."""
    digest = hashlib.md5(
        str(field_name).encode("utf-8", errors="ignore")
    ).hexdigest()
    index = int(digest[:8], 16) % len(VISUAL_FIELD_PALETTE)
    return VISUAL_FIELD_PALETTE[index]


def _load_visual_font(size):
    size = max(12, int(size))
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
    ]

    for font_path in candidates:
        try:
            return ImageFont.truetype(font_path, size=size)
        except Exception:
            continue

    return ImageFont.load_default()


def _wrap_annotation(text, max_chars=30):
    return "\n".join(
        textwrap.wrap(
            str(text),
            width=max_chars,
            break_long_words=False,
            break_on_hyphens=False,
        )
    )


def _visual_norm(text):
    if text is None:
        return ""
    value = unicodedata.normalize("NFKC", str(text)).casefold()
    value = value.replace("%", "")
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def _word_text_norm(text):
    value = _visual_norm(text)
    return value.replace(" ", "")


def _find_word_boxes_for_target(page, target, field_name=""):
    """
    Find OCR rectangles for a report value.

    The OCR passes use several segmentation modes, so Tesseract's internal
    block/line IDs are not reliable after the results are merged. Visual
    matching therefore uses the physical coordinates of OCR words instead.
    """
    target_norm = _visual_norm(target)
    if not target_norm or target_norm in {"not found", "—", "-"}:
        return []

    words = [
        w for w in page.get("ocr_words", [])
        if str(w.get("text", "")).strip()
    ]
    if not words:
        return []

    def word_center_y(word):
        return int(word.get("top", 0)) + max(
            1, int(word.get("height", 1))
        ) / 2

    ordered = sorted(
        words,
        key=lambda item: (
            word_center_y(item),
            int(item.get("left", 0)),
        ),
    )

    # Build physical OCR lines from Y positions rather than OCR block IDs.
    physical_lines = []
    current = []
    current_y = None
    current_height = 0

    for word in ordered:
        center_y = word_center_y(word)
        height = max(1, int(word.get("height", 1)))

        if current_y is None:
            current = [word]
            current_y = center_y
            current_height = height
            continue

        tolerance = max(
            5,
            int(max(current_height, height) * 0.85),
        )

        if abs(center_y - current_y) <= tolerance:
            current.append(word)
            current_y = (
                current_y * (len(current) - 1) + center_y
            ) / len(current)
            current_height = max(current_height, height)
        else:
            current.sort(
                key=lambda item: int(item.get("left", 0))
            )
            physical_lines.append(current)
            current = [word]
            current_y = center_y
            current_height = height

    if current:
        current.sort(
            key=lambda item: int(item.get("left", 0))
        )
        physical_lines.append(current)

    def build_box(group):
        if not group:
            return None

        left = min(int(w.get("left", 0)) for w in group)
        top = min(int(w.get("top", 0)) for w in group)
        right = max(
            int(w.get("left", 0)) + int(w.get("width", 0))
            for w in group
        )
        bottom = max(
            int(w.get("top", 0)) + int(w.get("height", 0))
            for w in group
        )

        if right <= left or bottom <= top:
            return None

        return (left, top, right, bottom)

    # ---------------------------------------------------------
    # 1. Exact contiguous match on physical OCR lines.
    # ---------------------------------------------------------
    for line_words in physical_lines:
        normalized_words = [
            _visual_norm(word["text"])
            for word in line_words
        ]

        for start in range(len(line_words)):
            joined = ""

            for end in range(start, len(line_words)):
                token = normalized_words[end]
                if not token:
                    continue

                joined = (joined + " " + token).strip()

                if (
                    joined == target_norm
                    or joined.replace(" ", "")
                    == target_norm.replace(" ", "")
                ):
                    box = build_box(line_words[start:end + 1])
                    if box:
                        return [box]

                if len(joined.replace(" ", "")) > len(
                    target_norm.replace(" ", "")
                ) + 8:
                    break

    # ---------------------------------------------------------
    # 2. Exact match across neighbouring physical lines.
    # ---------------------------------------------------------
    compact_target = target_norm.replace(" ", "")

    for start_line in range(len(physical_lines)):
        selected = []
        combined = ""

        for end_line in range(
            start_line,
            min(len(physical_lines), start_line + 4),
        ):
            for word in physical_lines[end_line]:
                token = _word_text_norm(word["text"])
                if not token:
                    continue

                combined += token
                selected.append(word)

                if combined == compact_target:
                    box = build_box(selected)
                    if box:
                        return [box]

                if len(combined) > len(compact_target) + 8:
                    break

            if len(combined) > len(compact_target) + 8:
                break

    # ---------------------------------------------------------
    # 3. Token-based fallback.
    #
    # Useful when OCR inserts/reorders words in dense artwork while
    # still recognizing the important words individually.
    # ---------------------------------------------------------
    from difflib import SequenceMatcher

    target_tokens = [
        token for token in target_norm.split()
        if (
            len(token) >= 3
            or token.isdigit()
        )
    ]

    # Do not use extremely generic tokens by themselves.
    if field_name:
        field_type = get_field_type(field_name)
    else:
        field_type = ""

    stop_tokens = {
        "the", "and", "with", "from", "made", "only",
        "wash", "not", "dry", "do",
    }

    token_candidates = []

    for target_token in target_tokens:
        candidates = []

        for word in ordered:
            word_norm = _visual_norm(word["text"])
            compact_word = word_norm.replace(" ", "")

            if not word_norm:
                continue

            exact = (
                word_norm == target_token
                or compact_word == target_token.replace(" ", "")
            )

            if exact:
                similarity = 1.0
            elif len(target_token) >= 4:
                similarity = SequenceMatcher(
                    None,
                    target_token,
                    compact_word,
                    autojunk=False,
                ).ratio()
            else:
                similarity = 0.0

            if similarity >= 0.78:
                candidates.append((similarity, word))

        if candidates:
            candidates.sort(
                key=lambda pair: (
                    -pair[0],
                    int(pair[1].get("top", 0)),
                    int(pair[1].get("left", 0)),
                )
            )
            token_candidates.append(
                (target_token, candidates[:12])
            )

    if not token_candidates:
        return []

    # Find a spatial cluster that contains the largest number of expected
    # tokens. This prevents a word being highlighted from an unrelated area.
    clusters = []

    for target_token, candidates in token_candidates:
        for similarity, word in candidates:
            cy = word_center_y(word)
            cx = (
                int(word.get("left", 0))
                + max(1, int(word.get("width", 1))) / 2
            )

            clusters.append({
                "target": target_token,
                "similarity": similarity,
                "word": word,
                "cx": cx,
                "cy": cy,
            })

    if not clusters:
        return []

    best_cluster = None

    for anchor in clusters:
        members = []

        for candidate in clusters:
            # A single field can legitimately cover a compact multi-line block.
            distance_x = abs(candidate["cx"] - anchor["cx"])
            distance_y = abs(candidate["cy"] - anchor["cy"])

            y_limit = 120 if field_type == "CARE" else 58

            if (
                distance_x <= max(260, page.get("image_width", 1000) * 0.18)
                and distance_y <= y_limit
            ):
                members.append(candidate)

        unique_targets = {}
        for member in members:
            token = member["target"]
            previous = unique_targets.get(token)

            if previous is None or member["similarity"] > previous["similarity"]:
                unique_targets[token] = member

        score = (
            len(unique_targets),
            sum(
                member["similarity"]
                for member in unique_targets.values()
            ),
        )

        if best_cluster is None or score > best_cluster["score"]:
            best_cluster = {
                "score": score,
                "members": list(unique_targets.values()),
            }

    if not best_cluster:
        return []

    expected_count = max(1, len(target_tokens))

    # A field with multiple meaningful expected tokens should not be mapped to
    # a random single OCR word unless that field itself is one token.
    if (
        expected_count >= 2
        and best_cluster["score"][0] < min(2, expected_count)
    ):
        return []

    selected_words = [
        member["word"]
        for member in best_cluster["members"]
    ]

    box = build_box(selected_words)
    return [box] if box else []


def build_highlighted_page_image(page, page_report):
    """
    Return the complete artwork page with field-specific colored highlights.

    Every independently compared field gets its own deterministic color.
    The annotation reads:
        <Excel field name> • PASS
    or:
        <Excel field name> • FAIL
    """
    image_bytes = page.get("image_bytes")
    if not image_bytes:
        return None

    base = Image.open(
        BytesIO(image_bytes)
    ).convert("RGBA")

    overlay = Image.new(
        "RGBA",
        base.size,
        (0, 0, 0, 0),
    )
    draw = ImageDraw.Draw(overlay)

    if page_report is not None and not page_report.empty:
        rows = page_report[
            page_report["STATUS"].isin(["PASS", "FAIL"])
        ].copy()

        # Draw each field independently. Sorting by FIELD makes the visual
        # result deterministic across reruns.
        sort_columns = ["FIELD"]
        if "FIELD NO" in rows.columns:
            sort_columns.append("FIELD NO")

        rows = rows.sort_values(
            by=sort_columns,
            kind="stable",
        )

        font_size = max(
            14,
            min(
                34,
                int(min(base.size) * 0.018),
            ),
        )
        font = _load_visual_font(font_size)

        for _, row in rows.iterrows():
            field_name = str(row.get("FIELD", "") or "").strip()
            status = str(row.get("STATUS", "") or "").strip()
            target = str(row.get("PDF OUTPUT", "") or "").strip()

            if not field_name or status not in {"PASS", "FAIL"}:
                continue

            boxes = _find_word_boxes_for_target(
                page,
                target,
                field_name,
            )
            if not boxes:
                continue

            rgb = _field_visual_color(field_name)

            fill = (
                int(rgb[0]),
                int(rgb[1]),
                int(rgb[2]),
                58 if status == "PASS" else 78,
            )
            outline = (
                int(rgb[0]),
                int(rgb[1]),
                int(rgb[2]),
                235,
            )

            label_text = _wrap_annotation(
                f"{field_name} • {status}",
                max_chars=34,
            )

            label_bbox = draw.multiline_textbbox(
                (0, 0),
                label_text,
                font=font,
                spacing=2,
            )
            label_width = label_bbox[2] - label_bbox[0] + 14
            label_height = label_bbox[3] - label_bbox[1] + 10

            for box_index, (left, top, right, bottom) in enumerate(boxes):
                pad = max(
                    4,
                    int(min(base.size) * 0.0035),
                )

                left = max(0, left - pad)
                top = max(0, top - pad)
                right = min(base.width - 1, right + pad)
                bottom = min(base.height - 1, bottom + pad)

                draw.rounded_rectangle(
                    (left, top, right, bottom),
                    radius=max(4, pad),
                    fill=fill,
                    outline=outline,
                    width=max(2, pad // 2),
                )

                # Put the field annotation immediately above the matched text.
                label_left = left
                label_top = top - label_height - 4

                if label_top < 0:
                    label_top = min(
                        base.height - label_height,
                        bottom + 4,
                    )

                label_left = min(
                    max(0, label_left),
                    max(0, base.width - label_width),
                )

                label_right = min(
                    base.width,
                    label_left + label_width,
                )
                label_bottom = min(
                    base.height,
                    label_top + label_height,
                )

                # Slightly translucent white background helps the annotation
                # remain readable without obscuring the complete artwork.
                draw.rounded_rectangle(
                    (
                        label_left,
                        label_top,
                        label_right,
                        label_bottom,
                    ),
                    radius=5,
                    fill=(
                        255,
                        255,
                        255,
                        235,
                    ),
                    outline=outline,
                    width=2,
                )

                draw.multiline_text(
                    (
                        label_left + 7,
                        label_top + 4,
                    ),
                    label_text,
                    fill=(
                        rgb[0],
                        rgb[1],
                        rgb[2],
                        255,
                    ),
                    font=font,
                    spacing=2,
                )

    return Image.alpha_composite(
        base,
        overlay,
    ).convert("RGB")


def _visual_image_bytes(image):
    if image is None:
        return None
    return _image_to_png_bytes(image)

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
    if not isinstance(page, dict):
        raise TypeError(
            f"Output page data is malformed. Expected a page dictionary, got {type(page).__name__}."
        )

    lines = build_page_lines(
        page.get("text", ""),
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

    return {
        "page": page.get("page"),
        "source_type": page.get("source_type", "pdf_text"),
        "lines": lines,
        "consumed": set(),
        "consumed_spans": {},
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

    # Support both normal PDF text ('60% cotton') and OCR/PDF extraction that
    # collapses the percent sign ('60cotton'). Material is captured until the
    # next percentage/material component or end of the content run.
    pattern = re.compile(
        r"(?P<pct>\d{1,3}(?:\.\d+)?)\s*%?\s*"
        r"(?P<material>[a-z][a-z0-9-]*(?:\s+[a-z][a-z0-9-]*){0,3})"
        r"(?=\s+\d{1,3}(?:\.\d+)?\s*%?|$)",
        re.IGNORECASE
    )

    values = []
    for match in pattern.finditer(normalized):
        material = match.group("material").strip()
        if not material:
            continue

        # Prevent accidental capture of common section/identifier words.
        material = re.split(
            r"\b(?:shell|liner|lining|body|rn|ca|made|size|color|colour)\b.*$",
            material,
            maxsplit=1
        )[0].strip()

        if material:
            values.append(
                f"{match.group('pct')}% {material}"
            )

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

        # Preserve the original PDF page number when only specific pages
        # were selected. Page 5 must still map to Data Row 5.
        try:
            original_page_index = int(page.get("page", page_index + 1)) - 1
        except Exception:
            original_page_index = page_index

        if page_row_mapping is not None:
            excel_index = page_row_mapping.get(original_page_index)
        else:
            excel_index = original_page_index

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
    product_type,
    page_row_mapping=None
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
        product_type,
        page_row_mapping=page_row_mapping
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
            highlighted = build_highlighted_page_image(page, page_report)
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
                img.width = min(900, highlighted.width)
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
    # OUTPUT PAGE SELECTION + PAGE → ORDER FORM ROW MAPPING
    # =========================================================
    output_page_count = 0
    mapping_complete = True
    page_row_mapping = {}
    selected_page_numbers = []
    selected_page_indices = []

    if excel_file and output_file:
        try:
            output_page_count = get_output_page_count(output_file)
        except Exception as error:
            st.warning(f"Unable to determine output page count: {error}")

        if output_page_count > 1:
            st.markdown(
                '<div class="section-title">📄 Artwork Page Selection</div>',
                unsafe_allow_html=True
            )
            st.caption(
                "Choose all artwork pages or select specific page number(s). "
                "Specific page selection is useful when checking one page from a large multi-page PDF."
            )

            page_mode = st.radio(
                "Artwork Pages to Compare",
                options=["All Pages", "Specific Page(s)"],
                horizontal=True,
                key=f"of_page_mode_{st.session_state['of_reset_id']}"
            )

            if page_mode == "Specific Page(s)":
                page_options = list(range(1, output_page_count + 1))
                previous_pages = st.session_state.get("of_selected_page_numbers", [])
                if not previous_pages:
                    previous_pages = [1]
                selected_page_numbers = st.multiselect(
                    "Select Artwork Page Number(s)",
                    options=page_options,
                    default=[p for p in previous_pages if p in page_options],
                    placeholder="Select page number(s)...",
                    key=f"of_selected_pages_{st.session_state['of_reset_id']}"
                )
                st.session_state["of_selected_page_numbers"] = selected_page_numbers
            else:
                selected_page_numbers = list(range(1, output_page_count + 1))
                st.session_state["of_selected_page_numbers"] = selected_page_numbers
        elif output_page_count == 1:
            selected_page_numbers = [1]
            st.session_state["of_selected_page_numbers"] = [1]

        selected_page_indices = [p - 1 for p in selected_page_numbers]

        # Mapping rules:
        # 1 data row  -> every selected page uses Data Row 1 / Excel Row 2.
        # Multiple rows -> Page N automatically maps to Data Row N whenever
        # that row exists. Only pages beyond the available rows need manual
        # row selection.
        if selected_page_indices:
            if len(df) == 1:
                for page_index in selected_page_indices:
                    page_row_mapping[page_index] = 0
            else:
                unmatched_pages = []
                for page_index in selected_page_indices:
                    if page_index < len(df):
                        page_row_mapping[page_index] = page_index
                    else:
                        unmatched_pages.append(page_index)

                if unmatched_pages:
                    st.markdown(
                        '<div class="section-title">🔗 Additional Page → Order Form Row Mapping</div>',
                        unsafe_allow_html=True
                    )
                    st.caption(
                        "These selected artwork pages do not have an automatic matching Order Form row. "
                        "Choose the correct data row below."
                    )

                    row_choices = []
                    for data_index in range(len(df)):
                        row = df.iloc[data_index]
                        preview_parts = []
                        preferred = []
                        for col in df.columns:
                            compact = normalize_text(col).replace(" ", "").replace("_", "").replace("-", "")
                            if any(token in compact for token in (
                                "itemcode", "itemnumber", "stylecode", "cdstyle",
                                "size", "color", "colour", "gender", "productgender",
                                "cdimport", "rn"
                            )):
                                preferred.append(col)
                        ordered_cols = preferred + [col for col in df.columns if col not in preferred]
                        seen_preview = set()
                        for col in ordered_cols:
                            value = row[col]
                            if is_blank_value(value):
                                continue
                            value_text = str(value).strip()
                            if not value_text:
                                continue
                            compact_col = normalize_text(col).replace(" ", "")
                            if compact_col in seen_preview:
                                continue
                            seen_preview.add(compact_col)
                            preview_parts.append(f"{col}: {value_text}")
                            if len(preview_parts) >= 3:
                                break
                        preview = " | ".join(preview_parts)
                        label = f"Data Row {data_index + 1} (Excel Row {data_index + 2})"
                        if preview:
                            label += f" — {preview}"
                        row_choices.append((data_index, label))

                    option_labels = ["— SELECT ORDER FORM ROW —"] + [label for _idx, label in row_choices]
                    label_to_index = {label: idx for idx, label in row_choices}

                    for page_index in unmatched_pages:
                        existing = st.session_state["of_page_row_mapping"].get(page_index)
                        default_pos = 0
                        if existing is not None:
                            for pos, (data_index, _label) in enumerate(row_choices, start=1):
                                if data_index == existing:
                                    default_pos = pos
                                    break
                        selected_label = st.selectbox(
                            f"Artwork Page {page_index + 1}",
                            options=option_labels,
                            index=default_pos,
                            key=f"of_page_row_select_{st.session_state['of_reset_id']}_{page_index}",
                            help="Data Row 1 corresponds to Excel Row 2; Data Row 2 corresponds to Excel Row 3; and so on."
                        )
                        if selected_label == "— SELECT ORDER FORM ROW —":
                            mapping_complete = False
                        else:
                            page_row_mapping[page_index] = label_to_index[selected_label]

            st.session_state["of_page_row_mapping"] = page_row_mapping

            if len(df) == 1 and selected_page_numbers:
                st.success(
                    "✅ Single-data-row Order Form detected. Every selected artwork page will be compared against Data Row 1 (Excel Row 2)."
                )
            elif mapping_complete and selected_page_numbers:
                mapping_preview = [
                    f"Page {page_index + 1} → Data Row {page_row_mapping[page_index] + 1} (Excel Row {page_row_mapping[page_index] + 2})"
                    for page_index in selected_page_indices
                    if page_index in page_row_mapping
                ]
                st.success("✅ Page-to-row mapping ready. " + " • ".join(mapping_preview))
            elif selected_page_numbers:
                st.warning("Please complete the row selection for every artwork page shown above.")
        else:
            mapping_complete = False
            if output_page_count > 1:
                st.info("Select at least one artwork page to compare.")

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

        # =========================================================
        # AUTO DETECT / MANUAL FIELD SELECTION
        # =========================================================
        available_fields = get_available_fields(df)

        if comparison_method == "Auto Detect":
            auto_key = (
                str(getattr(excel_file, "name", "")),
                int(getattr(excel_file, "size", 0)),
                str(getattr(output_file, "name", "")),
                int(getattr(output_file, "size", 0)),
                product_type,
                tuple(selected_page_numbers),
                tuple(sorted(page_row_mapping.items())),
            )

            if st.session_state.get("of_auto_detect_key") != auto_key:
                try:
                    with st.spinner("Auto Detect is reading the artwork with OCR..."):
                        auto_all_pages = extract_output_pages(output_file)
                        auto_pages = [
                            page for page in auto_all_pages
                            if int(page.get("page", 0)) in set(selected_page_numbers)
                        ]
                        detected_fields = auto_detect_fields(
                            df,
                            auto_pages,
                            product_type,
                            page_row_mapping=page_row_mapping
                        )
                    st.session_state["of_auto_detected_fields"] = detected_fields
                    st.session_state["of_auto_output_pages"] = auto_pages
                    st.session_state["of_auto_detect_key"] = auto_key
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
                "Auto Detect proposes populated fields that can be associated with the artwork. "
                "Review the list, remove fields, or add another populated field before comparison."
            )

            default_auto = [field for field in detected_fields if field in available_fields]
            selected_fields = st.multiselect(
                "Review detected fields",
                options=available_fields,
                default=default_auto,
                placeholder="Type to search fields...",
                label_visibility="collapsed",
                key=f"auto_selected_fields_{st.session_state['of_reset_id']}"
            )

            if selected_fields:
                st.caption("Selected fields: " + ", ".join(selected_fields))
            else:
                st.info("Auto Detect did not select any populated fields. Add fields manually from the list above.")

        else:
            st.markdown(
                '<div class="section-title">Select Variable Fields to Validate</div>',
                unsafe_allow_html=True
            )
            st.caption(
                "Only populated Order Form fields are shown. Open the dropdown and type to search by Excel column name or field terminology."
            )

            # Streamlit's multiselect has built-in live search.
            # The user can click the dropdown and start typing immediately;
            # matching fields are filtered as the text is entered, without
            # requiring a separate search box or pressing Enter.
            previous = st.session_state.get(
                f"selected_fields_{st.session_state['of_reset_id']}", []
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
                st.info("Select at least one Order Form field to continue.")

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
            st.info(
                f"ℹ️ The Order Form contains {len(df)} data row(s) while the "
                f"artwork contains {output_page_count} page(s). This is allowed. "
                "The current comparison maps Output Page 1 → Excel Row 2, "
                "Output Page 2 → Excel Row 3, and so on."
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
            and bool(selected_page_numbers)
            and mapping_complete
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
                    # Reuse the OCR pages already prepared for Auto Detect.
                    if (
                        comparison_method == "Auto Detect"
                        and st.session_state.get("of_auto_output_pages")
                    ):
                        output_pages = st.session_state["of_auto_output_pages"]
                    else:
                        all_output_pages = extract_output_pages(output_file)
                    output_pages = [
                        page for page in all_output_pages
                        if int(page.get("page", 0)) in set(selected_page_numbers)
                    ]

                    if not output_pages:
                        raise ValueError(
                            "No readable output pages were detected."
                        )

                    if not selected_fields:
                        raise ValueError(
                            "No fields are selected for comparison."
                        )

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
                    st.session_state["of_visual_pages"] = output_pages
                    st.session_state["of_report_page_row_mapping"] = dict(page_row_mapping)
                    st.session_state["of_report_selected_pages"] = list(selected_page_numbers)

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

        report_selected_pages = st.session_state.get("of_report_selected_pages", [])
        if report_selected_pages:
            st.caption("Artwork pages compared: " + ", ".join(str(p) for p in report_selected_pages))

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
                "Each compared field is highlighted in a different color and "
                "annotated directly on the full artwork as FIELD • PASS/FAIL."
            )

            # Compact color legend for the field-specific annotations.
            legend_items = []
            for field in sorted(set(
                str(value)
                for value in report_fields
                if str(value).strip()
            )):
                rgb = _field_visual_color(field)
                legend_items.append(
                    f'<span style="display:inline-flex;align-items:center;'
                    f'margin-right:14px;margin-bottom:6px;">'
                    f'<span style="width:12px;height:12px;border-radius:3px;'
                    f'background:rgb({rgb[0]},{rgb[1]},{rgb[2]});'
                    f'display:inline-block;margin-right:5px;"></span>'
                    f'{field}</span>'
                )

            if legend_items:
                st.markdown(
                    "".join(legend_items),
                    unsafe_allow_html=True
                )

            for page in visual_pages:
                page_num = page.get("page")
                page_report = report[report["PDF PAGE"] == page_num] if "PDF PAGE" in report.columns else report.iloc[0:0]
                highlighted = build_highlighted_page_image(page, page_report)
                if highlighted is not None:
                    st.markdown(f"**Artwork Page {page_num}**")
                    st.image(highlighted, width=750)

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

                PDF Page 1 → Excel Row 2

                PDF Page 2 → Excel Row 3

                PDF Page 3 → Excel Row 4

                and so on.

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
