import streamlit as st
import pandas as pd
import fitz
import re
import unicodedata
import io

from PIL import Image


# =========================================================
# CSS
# =========================================================

def _apply_tool_css():

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

        </style>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# TEXT NORMALIZATION
# =========================================================

def normalize_text(text):

    if text is None:
        return ""

    if pd.isna(text):
        return ""

    text = str(text)

    text = unicodedata.normalize("NFKC", text)

    text = text.lower()

    text = text.replace("’", "'")
    text = text.replace("`", "'")
    text = text.replace("–", "-")
    text = text.replace("—", "-")

    text = text.replace("\n", " ")
    text = text.replace("\r", " ")

    # PDF bullet artefact
    text = re.sub(
        r"(^|\s)n(?=\s)",
        " ",
        text
    )

    # Normalize separators
    text = re.sub(
        r"[,.;:|/\\]+",
        " ",
        text
    )

    text = re.sub(
        r"-+",
        " ",
        text
    )

    text = re.sub(
        r"[^\w%#'\s]",
        " ",
        text,
        flags=re.UNICODE
    )

    # Apostrophes do not affect validation
    text = text.replace("'", "")

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


def compact_text(text):

    return normalize_text(text).replace(" ", "")


def tokenize(text):

    value = normalize_text(text)

    if not value:
        return []

    return value.split()


# =========================================================
# FIELD TYPE
# =========================================================

def get_field_type(field_name):

    field = normalize_text(field_name)

    compact = field.replace(" ", "")

    # IDENTIFIERS
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

    # BATCH / LOT
    if (
        "batch" in compact
        or "lotnumber" in compact
        or "lotno" in compact
        or compact == "lot"
    ):
        return "BATCH"

    # QUANTITY
    if (
        "quantity" in compact
        or compact == "qty"
        or "units" in compact
        or "pieces" in compact
        or compact == "pcs"
    ):
        return "QUANTITY"

    # COO
    if (
        "coo" in compact
        or "countryoforigin" in compact
        or "countryorigin" in compact
        or "madein" in compact
        or compact == "origin"
    ):
        return "COO"

    # CONTENT
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

    # CARE
    if (
        "care" in compact
        or "wash" in compact
        or "washing" in compact
        or "laundry" in compact
        or "instruction" in compact
    ):
        return "CARE"

    # SIZE
    if (
        "size" in compact
        or "sizeline" in compact
        or "alpha" in compact
        or "waist" in compact
        or "inseam" in compact
        or compact == "fit"
    ):
        return "SIZE"

    # RN
    if (
        compact == "rn"
        or "registrationnumber" in compact
        or "companyrn" in compact
    ):
        return "RN"

    # BRAND
    if "brand" in compact:
        return "BRAND"

    # COLOR
    if (
        "color" in compact
        or "colour" in compact
    ):
        return "COLOR"

    # GENDER
    if "gender" in compact:
        return "GENDER"

    # ATTRIBUTE
    if (
        "attribute" in compact
        or "technology" in compact
        or "feature" in compact
    ):
        return "ATTRIBUTE"

    return "GENERAL"


# =========================================================
# FIELD LANGUAGE
# =========================================================

def get_field_region(field_name):

    original = str(field_name).lower()
    field = normalize_text(field_name)
    compact = field.replace(" ", "")

    if (
        "_en" in original
        or compact.endswith("en")
        or "english" in compact
    ):
        return "EN"

    if (
        "_fr" in original
        or compact.endswith("fr")
        or "french" in compact
        or "canada" in compact
    ):
        return "FR"

    if (
        "_sp" in original
        or compact.endswith("sp")
        or "spanish" in compact
        or "espanol" in compact
        or "span" in compact
    ):
        return "SP"

    return ""


# =========================================================
# LOAD EXCEL
# =========================================================

def load_excel(file):

    file.seek(0)

    df = pd.read_excel(
        file,
        header=0
    )

    df.columns = [
        str(column).strip()
        for column in df.columns
    ]

    return df


# =========================================================
# OUTPUT EXTRACTION
# =========================================================

def _usable_text(text):

    if not text:
        return False

    value = str(text).strip()

    if not value:
        return False

    alnum = re.sub(
        r"[^\w%#]",
        "",
        value,
        flags=re.UNICODE
    )

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
            "OCR could not run. Make sure Tesseract OCR and its language "
            "data are installed in the Streamlit environment."
        ) from exc


def _render_pdf_page(page):

    pixmap = page.get_pixmap(
        matrix=fitz.Matrix(2.5, 2.5),
        alpha=False
    )

    image = Image.open(
        io.BytesIO(
            pixmap.tobytes("png")
        )
    ).convert("RGB")

    return image


def get_output_page_count(file):

    name = str(
        getattr(file, "name", "")
    ).lower()

    if name.endswith(".pdf"):

        file.seek(0)

        pdf_bytes = file.read()

        document = fitz.open(
            stream=pdf_bytes,
            filetype="pdf"
        )

        count = len(document)

        document.close()

        return count

    if name.endswith(
        (".jpg", ".jpeg", ".png")
    ):
        return 1

    return 0


def extract_output_pages(file):

    name = str(
        getattr(file, "name", "")
    ).lower()

    if name.endswith(".pdf"):

        file.seek(0)

        pdf_bytes = file.read()

        document = fitz.open(
            stream=pdf_bytes,
            filetype="pdf"
        )

        pages = []

        for page_number, page in enumerate(document):

            direct_text = page.get_text("text") or ""

            if _usable_text(direct_text):

                text = direct_text
                source_type = "pdf_text"

            else:

                image = _render_pdf_page(page)

                text = _ocr_image(image)

                source_type = "ocr"

            pages.append(
                {
                    "page": page_number + 1,
                    "text": text,
                    "source_type": source_type
                }
            )

        document.close()

        return pages

    if name.endswith(
        (".jpg", ".jpeg", ".png")
    ):

        file.seek(0)

        image = Image.open(file).convert("RGB")

        text = _ocr_image(image)

        return [
            {
                "page": 1,
                "text": text,
                "source_type": "ocr"
            }
        ]

    raise ValueError(
        "Unsupported output format. Please upload PDF, JPG, JPEG, or PNG."
    )


def load_pdf(file):
    return extract_output_pages(file)


# =========================================================
# PDF CLEANING
# =========================================================

def clean_pdf_line(line):

    if not line:
        return ""

    line = str(line).strip()

    # Remove PDF bullet artefact "n"
    line = re.sub(
        r"^\s*n\s+(?=[A-Za-z])",
        "",
        line
    )

    return line.strip()


# =========================================================
# PDF BLOCK CREATION
# =========================================================

def create_pdf_blocks(page_text):

    if not page_text:
        return []

    raw_lines = page_text.splitlines()

    lines = []

    for raw_line in raw_lines:

        line = clean_pdf_line(raw_line)

        if not line:
            continue

        if len(line) > 1500:
            continue

        lines.append(line)

    if not lines:
        return []

    blocks = []

    # Individual lines
    for line in lines:
        blocks.append(line)

    # Adjacent blocks
    maximum = min(
        8,
        len(lines)
    )

    for size in range(
        2,
        maximum + 1
    ):

        for start in range(
            len(lines) - size + 1
        ):

            block = " ".join(
                lines[
                    start:start + size
                ]
            )

            if block:
                blocks.append(block)

    # Remove duplicates
    unique = []

    seen = set()

    for block in blocks:

        normalized = normalize_text(block)

        if not normalized:
            continue

        if normalized in seen:
            continue

        seen.add(normalized)

        unique.append(block)

    return unique


# =========================================================
# PFL BLOCK CREATION
# =========================================================

def get_pfl_panel_blocks(page_text):

    if not page_text:
        return []

    raw_lines = page_text.splitlines()

    cleaned = []

    panel_pattern = re.compile(
        r"^\s*(?:panel\s*)?(\d{1,3})\s*$",
        re.IGNORECASE
    )

    for raw in raw_lines:

        line = clean_pdf_line(raw)

        if not line:
            continue

        if len(line) > 1500:
            continue

        if panel_pattern.match(line):
            continue

        if re.match(
            r"^\s*panel\s*[-#:]?\s*\d{1,3}\s*$",
            line,
            re.IGNORECASE
        ):
            continue

        cleaned.append(line)

    if not cleaned:
        return []

    blocks = list(cleaned)

    maximum = min(
        24,
        len(cleaned)
    )

    for size in range(
        2,
        maximum + 1
    ):

        for start in range(
            len(cleaned) - size + 1
        ):

            block = " ".join(
                cleaned[
                    start:start + size
                ]
            )

            if block:
                blocks.append(block)

    unique = []

    seen = set()

    for block in blocks:

        normalized = normalize_text(block)

        if not normalized:
            continue

        if normalized in seen:
            continue

        seen.add(normalized)

        unique.append(block)

    return unique


def get_comparison_blocks(
    page_text,
    product_type
):

    if product_type == "PFL":

        return get_pfl_panel_blocks(
            page_text
        )

    return create_pdf_blocks(
        page_text
    )


# =========================================================
# BASIC MATCHING
# =========================================================

def exact_match(
    expected,
    actual
):

    expected_normalized = normalize_text(
        expected
    )

    actual_normalized = normalize_text(
        actual
    )

    if not expected_normalized:
        return False

    if not actual_normalized:
        return False

    # Normal normalized match
    if expected_normalized in actual_normalized:
        return True

    # Space-independent match
    expected_compact = compact_text(
        expected
    )

    actual_compact = compact_text(
        actual
    )

    if (
        expected_compact
        and
        expected_compact in actual_compact
    ):
        return True

    return False


def find_exact_value(
    expected,
    pdf_blocks
):

    for block in pdf_blocks:

        if exact_match(
            expected,
            block
        ):

            return block

    return None


# =========================================================
# FIELD-SPECIFIC EXTRACTION
#
# IMPORTANT:
#
# We deliberately do NOT use generic fuzzy comparison here.
#
# Each field type gets its own logic.
# =========================================================

def extract_coo_value(text):

    if not text:
        return None

    normalized = normalize_text(text)

    patterns = [

        r"\bmade\s+in\s+([a-z][a-z\s\-]+)",

        r"\bfabrique\s+en\s+([a-z][a-z\s\-]+)",

        r"\bhecho\s+en\s+([a-z][a-z\s\-]+)",

    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            normalized,
            re.IGNORECASE
        )

        if match:

            value = match.group(0).strip()

            value = re.split(
                r"\b(?:rn|ca|sku|size|color|colour)\b",
                value,
                maxsplit=1
            )[0].strip()

            return value

    return None


def extract_content_values(text):

    if not text:
        return []

    normalized = normalize_text(text)

    results = []

    # Examples:
    # 100% cotton
    # 60% cotton 40% polyester
    # shell 100% cotton

    pattern = re.compile(
        r"(\d{1,3}(?:\.\d+)?)\s*%\s*"
        r"([a-z][a-z0-9\s\-]*)",
        re.IGNORECASE
    )

    for match in pattern.finditer(normalized):

        percentage = match.group(1)

        material = match.group(2).strip()

        # Stop material before another percentage / common field
        material = re.split(
            r"\b(?:shell|liner|body|lining|exclusive|of|rn|ca)\b",
            material,
            maxsplit=1
        )[0].strip()

        material = re.sub(
            r"\s+",
            " ",
            material
        )

        if material:

            results.append(
                f"{percentage}% {material}"
            )

    return results


def extract_gender_value(text):

    if not text:
        return None

    normalized = normalize_text(text)

    gender_values = [
        "boys",
        "girls",
        "women",
        "men",
        "unisex",
        "boy",
        "girl",
        "woman",
        "man"
    ]

    for value in gender_values:

        if re.search(
            rf"\b{re.escape(value)}\b",
            normalized
        ):

            return value

    return None


def extract_color_value(text):

    if not text:
        return None

    normalized = normalize_text(text)

    patterns = [
        r"\bcolor\s+([a-z][a-z\s\-]*)",
        r"\bcolour\s+([a-z][a-z\s\-]*)"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            normalized,
            re.IGNORECASE
        )

        if match:

            value = match.group(1).strip()

            value = re.split(
                r"\b(?:size|rn|ca|made|country)\b",
                value,
                maxsplit=1
            )[0].strip()

            if value:
                return value

    return None


def extract_size_value(text):

    if not text:
        return None

    normalized = normalize_text(text)

    patterns = [

        r"\bsize\s*[:\-]?\s*([a-z0-9][a-z0-9\s\-\/]*)",

    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            normalized,
            re.IGNORECASE
        )

        if match:

            value = match.group(1).strip()

            value = re.split(
                r"\b(?:rn|ca|made|color|colour)\b",
                value,
                maxsplit=1
            )[0].strip()

            if value:
                return value

    return None


def extract_identifier_value(text):

    if not text:
        return None

    normalized = normalize_text(text)

    patterns = [

        r"\bsku\s*[:#\-]?\s*([a-z0-9][a-z0-9\-_\/]*)",

        r"\bitem\s*(?:code|no|number)\s*[:#\-]?\s*"
        r"([a-z0-9][a-z0-9\-_\/]*)",

        r"\bstyle\s*(?:code|no|number)?\s*[:#\-]?\s*"
        r"([a-z0-9][a-z0-9\-_\/]*)",

    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            normalized,
            re.IGNORECASE
        )

        if match:

            return match.group(1).strip()

    return None


# =========================================================
# FIELD ANCHORS
# =========================================================

FIELD_ANCHORS = {

    "COO": [
        "made in",
        "fabrique en",
        "hecho en"
    ],

    "CONTENT": [
        "%",
        "shell",
        "liner",
        "lining",
        "body",
        "fiber",
        "fibre",
        "content",
        "composition"
    ],

    "CARE": [
        "wash",
        "laver",
        "lavar",
        "bleach",
        "blanchiment",
        "dry clean",
        "detergent"
    ],

    "SIZE": [
        "size"
    ],

    "IDENTIFIER": [
        "sku",
        "item code",
        "item no",
        "style",
        "style code"
    ],

    "BATCH": [
        "batch",
        "lot"
    ],

    "QUANTITY": [
        "quantity",
        "qty",
        "units",
        "pcs"
    ],

    "RN": [
        "rn",
        "ca"
    ],

    "COLOR": [
        "color",
        "colour"
    ],

    "BRAND": [
        "brand"
    ],

    "GENDER": [
        "boys",
        "girls",
        "women",
        "men",
        "unisex"
    ],

    "ATTRIBUTE": [
        "attribute",
        "technology",
        "feature"
    ],

    "GENERAL": []
}


# =========================================================
# RELEVANT BLOCK CHECK
# =========================================================

def is_relevant_block(
    block,
    field_type,
    field_name
):

    normalized = normalize_text(
        block
    )

    if not normalized:
        return False

    anchors = FIELD_ANCHORS.get(
        field_type,
        []
    )

    for anchor in anchors:

        anchor_normalized = normalize_text(
            anchor
        )

        if (
            anchor_normalized
            and
            anchor_normalized in normalized
        ):

            return True

    # Language clues
    region = get_field_region(
        field_name
    )

    if region == "EN":

        markers = [
            "made in",
            "machine wash",
            "shell",
            "liner",
            "cotton",
            "polyester",
            "bleach"
        ]

        for marker in markers:

            if normalize_text(marker) in normalized:
                return True

    elif region == "FR":

        markers = [
            "fabrique en",
            "laver",
            "extérieur",
            "doublure",
            "polyester",
            "sans chlore"
        ]

        for marker in markers:

            if normalize_text(marker) in normalized:
                return True

    elif region == "SP":

        markers = [
            "hecho en",
            "lavar",
            "forro",
            "poliester",
            "cloro"
        ]

        for marker in markers:

            if normalize_text(marker) in normalized:
                return True

    return False


# =========================================================
# FIELD-SPECIFIC MISMATCH SEARCH
# =========================================================

def search_field_mismatch(
    expected,
    pdf_blocks,
    field_name
):

    field_type = get_field_type(
        field_name
    )

    expected_normalized = normalize_text(
        expected
    )

    # =====================================================
    # COO
    # =====================================================

    if field_type == "COO":

        expected_coo = extract_coo_value(
            expected
        )

        # If Excel says only a country, e.g. INDIA
        if expected_coo is None:

            if expected_normalized:
                expected_country = expected_normalized
            else:
                return None

        else:

            expected_country = expected_coo

        for block in pdf_blocks:

            actual_coo = extract_coo_value(
                block
            )

            if not actual_coo:
                continue

            actual_normalized = normalize_text(
                actual_coo
            )

            if (
                expected_normalized in actual_normalized
                or
                expected_country in actual_normalized
            ):

                continue

            return {
                "status": "FAIL",
                "pdf": block,
                "difference": (
                    f"Expected: {expected} | "
                    f"Found: {actual_coo}"
                )
            }

        return None

    # =====================================================
    # CONTENT
    # =====================================================

    if field_type == "CONTENT":

        expected_content = extract_content_values(
            expected
        )

        # If Excel value doesn't contain percentage,
        # don't try to invent a content comparison.
        if not expected_content:
            return None

        for block in pdf_blocks:

            actual_content = extract_content_values(
                block
            )

            if not actual_content:
                continue

            expected_normalized_values = [
                normalize_text(value)
                for value in expected_content
            ]

            actual_normalized_values = [
                normalize_text(value)
                for value in actual_content
            ]

            # If any complete expected composition exists,
            # it is not a mismatch.
            all_found = all(
                value in actual_normalized_values
                for value in expected_normalized_values
            )

            if all_found:
                continue

            return {
                "status": "FAIL",
                "pdf": block,
                "difference": (
                    f"Expected: {expected} | "
                    f"Found: {' | '.join(actual_content)}"
                )
            }

        return None

    # =====================================================
    # GENDER
    # =====================================================

    if field_type == "GENDER":

        expected_gender = extract_gender_value(
            expected
        )

        if not expected_gender:
            expected_gender = expected_normalized

        for block in pdf_blocks:

            actual_gender = extract_gender_value(
                block
            )

            if not actual_gender:
                continue

            if (
                normalize_text(expected_gender)
                ==
                normalize_text(actual_gender)
            ):
                continue

            return {
                "status": "FAIL",
                "pdf": block,
                "difference": (
                    f"Expected: {expected} | "
                    f"Found: {actual_gender}"
                )
            }

        return None

    # =====================================================
    # COLOR
    # =====================================================

    if field_type == "COLOR":

        expected_color = normalize_text(
            expected
        )

        # Remove optional "color:"
        expected_color = re.sub(
            r"^color\s*:?\s*",
            "",
            expected_color
        )

        expected_color = re.sub(
            r"^colour\s*:?\s*",
            "",
            expected_color
        )

        for block in pdf_blocks:

            actual_color = extract_color_value(
                block
            )

            if not actual_color:
                continue

            if (
                expected_color
                in
                normalize_text(actual_color)
            ):

                continue

            return {
                "status": "FAIL",
                "pdf": block,
                "difference": (
                    f"Expected: {expected} | "
                    f"Found: {actual_color}"
                )
            }

        return None

    # =====================================================
    # SIZE
    # =====================================================

    if field_type == "SIZE":

        expected_size = normalize_text(
            expected
        )

        expected_size = re.sub(
            r"^size\s*:?\s*",
            "",
            expected_size
        )

        for block in pdf_blocks:

            actual_size = extract_size_value(
                block
            )

            if not actual_size:
                continue

            if (
                expected_size
                in
                normalize_text(actual_size)
            ):

                continue

            return {
                "status": "FAIL",
                "pdf": block,
                "difference": (
                    f"Expected: {expected} | "
                    f"Found: {actual_size}"
                )
            }

        return None

    # =====================================================
    # IDENTIFIER
    # =====================================================

    if field_type == "IDENTIFIER":

        expected_identifier = normalize_text(
            expected
        )

        expected_identifier = re.sub(
            r"^(sku|item code|item no|style code|style)\s*[:#\-]?\s*",
            "",
            expected_identifier
        )

        for block in pdf_blocks:

            actual_identifier = extract_identifier_value(
                block
            )

            if not actual_identifier:
                continue

            if (
                expected_identifier
                ==
                normalize_text(actual_identifier)
            ):

                continue

            return {
                "status": "FAIL",
                "pdf": block,
                "difference": (
                    f"Expected: {expected} | "
                    f"Found: {actual_identifier}"
                )
            }

        return None

    # =====================================================
    # CARE
    #
    # IMPORTANT:
    #
    # We don't perform generic fuzzy comparison.
    #
    # We only consider a relevant care block and look
    # for meaningful expected words.
    # =====================================================

    if field_type == "CARE":

        expected_tokens = tokenize(
            expected
        )

        if not expected_tokens:
            return None

        important_tokens = [
            token
            for token in expected_tokens
            if len(token) >= 4
        ]

        if not important_tokens:
            return None

        for block in pdf_blocks:

            if not is_relevant_block(
                block,
                field_type,
                field_name
            ):
                continue

            actual_normalized = normalize_text(
                block
            )

            # Exact expected value should already have
            # been handled, so here we only search for
            # meaningful overlap.
            common = [
                token
                for token in important_tokens
                if token in actual_normalized
            ]

            # A relevant care block with enough expected
            # vocabulary is a genuine mismatch candidate.
            if len(common) >= max(
                1,
                int(len(important_tokens) * 0.45)
            ):

                return {
                    "status": "FAIL",
                    "pdf": block,
                    "difference": (
                        f"Expected: {expected} | "
                        f"Found: {block}"
                    )
                }

        return None

    # =====================================================
    # BATCH
    # =====================================================

    if field_type == "BATCH":

        expected_value = normalize_text(
            expected
        )

        for block in pdf_blocks:

            normalized = normalize_text(
                block
            )

            match = re.search(
                r"\b(?:batch|lot)\s*[:#\-]?\s*"
                r"([a-z0-9\-_\/]+)",
                normalized
            )

            if not match:
                continue

            actual_value = match.group(1)

            if actual_value != expected_value:
                return {
                    "status": "FAIL",
                    "pdf": block,
                    "difference": (
                        f"Expected: {expected} | "
                        f"Found: {actual_value}"
                    )
                }

        return None

    # =====================================================
    # QUANTITY
    # =====================================================

    if field_type == "QUANTITY":

        expected_value = normalize_text(
            expected
        )

        for block in pdf_blocks:

            normalized = normalize_text(
                block
            )

            match = re.search(
                r"\b(?:quantity|qty|units|pcs)\s*"
                r"[:#\-]?\s*(\d+)",
                normalized
            )

            if not match:
                continue

            actual_value = match.group(1)

            if actual_value != expected_value:
                return {
                    "status": "FAIL",
                    "pdf": block,
                    "difference": (
                        f"Expected: {expected} | "
                        f"Found: {actual_value}"
                    )
                }

        return None

    # =====================================================
    # GENERAL
    #
    # NO FUZZY MATCHING.
    #
    # We deliberately don't guess a mismatch for an unknown
    # field. This prevents unrelated artwork text becoming
    # a false FAIL.
    # =====================================================

    return None


# =========================================================
# CHECK ONE FIELD
# =========================================================

def check_field(
    expected,
    pdf_blocks,
    field_name
):

    if expected is None:
        return {
            "status": "SKIP",
            "pdf": "—",
            "difference": "Blank Order Form value."
        }

    if pd.isna(expected):
        return {
            "status": "SKIP",
            "pdf": "—",
            "difference": "Blank Order Form value."
        }

    expected = str(
        expected
    ).strip()

    if not expected:
        return {
            "status": "SKIP",
            "pdf": "—",
            "difference": "Blank Order Form value."
        }

    # =====================================================
    # 1. EXACT MATCH = PASS
    # =====================================================

    exact_block = find_exact_value(
        expected,
        pdf_blocks
    )

    if exact_block:

        return {
            "status": "PASS",
            "pdf": exact_block,
            "difference": "—"
        }

    # =====================================================
    # 2. FIELD-SPECIFIC WRONG VALUE = FAIL
    # =====================================================

    mismatch = search_field_mismatch(
        expected,
        pdf_blocks,
        field_name
    )

    if mismatch:
        return mismatch

    # =====================================================
    # 3. EXPECTED VALUE NOT DETECTED
    # =====================================================

    return {
        "status": "NOT FOUND",
        "pdf": "Not found",
        "difference": (
            "Expected value was not detected in the "
            "relevant artwork content."
        )
    }


# =========================================================
# FIELD AVAILABILITY
#
# IMPORTANT:
#
# Only fields with at least one non-empty cell appear
# in Select Fields.
# =========================================================

def get_available_fields(df):

    available = []

    for column in df.columns:

        series = df[column]

        has_value = series.apply(
            lambda value:
                not pd.isna(value)
                and
                str(value).strip() != ""
        ).any()

        if has_value:
            available.append(
                str(column)
            )

    return available


# =========================================================
# AUTO DETECT CANDIDATE FIELD
# =========================================================

AUTO_IGNORED_FIELD_WORDS = {
    "sr",
    "serial",
    "serialno",
    "serialnumber",
    "jobno",
    "jobnumber",
    "orderno",
    "ordernumber",
    "orderdate",
    "ticket",
    "ticketno",
    "createddate",
    "modifieddate"
}


def _auto_field_is_candidate(
    field_name
):

    normalized = normalize_text(
        field_name
    )

    compact = normalized.replace(
        " ",
        ""
    ).replace(
        "_",
        ""
    ).replace(
        "-",
        ""
    )

    if not compact:
        return False

    if compact in AUTO_IGNORED_FIELD_WORDS:
        return False

    if get_field_type(field_name) != "GENERAL":
        return True

    keywords = [
        "sku",
        "item",
        "style",
        "batch",
        "lot",
        "qty",
        "quantity",
        "code",
        "attribute",
        "color",
        "colour",
        "size",
        "gender",
        "content",
        "care",
        "coo",
        "origin",
        "brand"
    ]

    return any(
        keyword in compact
        for keyword in keywords
    )


# =========================================================
# AUTO DETECT
#
# IMPORTANT CHANGE:
#
# Auto Detect now checks each populated field against the
# corresponding artwork row.
#
# It recognizes:
#
# PASS -> expected value exists
# FAIL -> field-specific wrong value exists
#
# It does NOT use generic fuzzy matching.
# =========================================================

def auto_detect_fields(
    df,
    output_pages,
    product_type
):

    candidates = [
        str(column)
        for column in df.columns
        if _auto_field_is_candidate(
            str(column)
        )
    ]

    detected = []

    page_blocks = [
        get_comparison_blocks(
            page.get("text", ""),
            product_type
        )
        for page in output_pages
    ]

    for field in candidates:

        field_detected = False

        for page_index, blocks in enumerate(
            page_blocks
        ):

            if page_index >= len(df):
                break

            value = df.iloc[
                page_index
            ][field]

            if pd.isna(value):
                continue

            value = str(
                value
            ).strip()

            if not value:
                continue

            # ---------------------------------------------
            # PASS evidence
            # ---------------------------------------------

            exact_block = find_exact_value(
                value,
                blocks
            )

            if exact_block:

                field_detected = True
                break

            # ---------------------------------------------
            # FAIL evidence
            # ---------------------------------------------

            mismatch = search_field_mismatch(
                value,
                blocks,
                field
            )

            if mismatch:

                field_detected = True
                break

        if field_detected:
            detected.append(field)

    return detected


# =========================================================
# BUILD REPORT
#
# PDF Page 1 -> Excel Row 2
# PDF Page 2 -> Excel Row 3
# etc.
# =========================================================

def build_report(
    df,
    pdf_pages,
    selected_fields,
    product_type
):

    results = []

    field_no = 1

    for page_index, page in enumerate(
        pdf_pages
    ):

        excel_index = page_index

        # -------------------------------------------------
        # No corresponding Excel row
        # -------------------------------------------------

        if excel_index >= len(df):

            for field in selected_fields:

                results.append(
                    {
                        "FIELD NO": field_no,
                        "PDF PAGE": page["page"],
                        "EXCEL ROW": "N/A",
                        "FIELD": field,
                        "ORDER FORM DATA": "No Excel row",
                        "PDF OUTPUT": "No corresponding Order Form row",
                        "STATUS": "NOT FOUND",
                        "DIFFERENCE": "No corresponding Excel row."
                    }
                )

                field_no += 1

            continue

        row = df.iloc[
            excel_index
        ]

        pdf_blocks = get_comparison_blocks(
            page["text"],
            product_type
        )

        for field in selected_fields:

            value = row[field]

            if pd.isna(value):
                value = ""
            else:
                value = str(
                    value
                ).strip()

            # -------------------------------------------------
            # Blank cell = ignored
            # -------------------------------------------------

            if not value:

                results.append(
                    {
                        "FIELD NO": field_no,
                        "PDF PAGE": page["page"],
                        "EXCEL ROW": excel_index + 2,
                        "FIELD": field,
                        "ORDER FORM DATA": "",
                        "PDF OUTPUT": "—",
                        "STATUS": "SKIP",
                        "DIFFERENCE": (
                            "Blank Order Form value — "
                            "PDF content ignored."
                        )
                    }
                )

                field_no += 1

                continue

            result = check_field(
                value,
                pdf_blocks,
                field
            )

            results.append(
                {
                    "FIELD NO": field_no,
                    "PDF PAGE": page["page"],
                    "EXCEL ROW": excel_index + 2,
                    "FIELD": field,
                    "ORDER FORM DATA": value,
                    "PDF OUTPUT": result["pdf"],
                    "STATUS": result["status"],
                    "DIFFERENCE": result["difference"]
                }
            )

            field_no += 1

    return pd.DataFrame(
        results
    )


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


# =========================================================
# MAIN
# =========================================================

def main():

    _apply_tool_css()

    # =====================================================
    # SESSION STATE
    # =====================================================

    if "of_reset_id" not in st.session_state:
        st.session_state["of_reset_id"] = 0

    if "of_report" not in st.session_state:
        st.session_state["of_report"] = None

    # =====================================================
    # TITLE
    # =====================================================

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

    # =====================================================
    # NAVIGATION
    # =====================================================

    nav_left, nav_right = st.columns(
        [1, 1]
    )

    with nav_left:

        if st.button(
            "← HOME",
            key="of_back_home",
            width="stretch"
        ):

            st.session_state[
                "selected_tool"
            ] = None

            st.session_state[
                "of_report"
            ] = None

            st.rerun()

    with nav_right:

        if st.button(
            "🆕 NEW START",
            key="of_new_start",
            width="stretch"
        ):

            st.session_state[
                "of_reset_id"
            ] += 1

            st.session_state[
                "of_report"
            ] = None

            st.rerun()

    # =====================================================
    # PRODUCT TYPE
    # =====================================================

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

    # =====================================================
    # UPLOAD AREA
    # =====================================================

    left_column, right_column = st.columns(
        2
    )

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
            type=[
                "pdf",
                "jpg",
                "jpeg",
                "png"
            ],
            key=f"output_upload_{st.session_state['of_reset_id']}"
        )

    # =====================================================
    # LOAD EXCEL
    # =====================================================

    df = None

    if excel_file:

        try:

            df = load_excel(
                excel_file
            )

        except Exception as error:

            st.error(
                f"Unable to read the Excel Order Form: {error}"
            )

            return

    # =====================================================
    # COMPARISON METHOD
    # =====================================================

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
            options=[
                "Auto Detect",
                "Select Fields"
            ],
            index=None,
            horizontal=True,
            key=f"comparison_method_{st.session_state['of_reset_id']}"
        )

        # =================================================
        # SELECT FIELDS
        # =================================================

        if comparison_method == "Select Fields":

            st.markdown(
                '<div class="section-title">'
                'Select Variable Fields to Validate'
                '</div>',
                unsafe_allow_html=True
            )

            st.caption(
                "Only fields containing Order Form data are shown below."
            )

            # ---------------------------------------------
            # IMPORTANT:
            # EMPTY COLUMNS ARE REMOVED
            # ---------------------------------------------

            available_fields = get_available_fields(
                df
            )

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

                        value = str(
                            value
                        ).strip()

                        if value:
                            values.append(value)

                    preview_rows.append(
                        {
                            "Excel Field": field,
                            "Values": len(values),
                            "Preview": " | ".join(
                                values[:3]
                            )
                        }
                    )

                with st.expander(
                    "🔎 Preview Selected Fields"
                ):

                    st.dataframe(
                        pd.DataFrame(
                            preview_rows
                        ),
                        width="stretch",
                        hide_index=True
                    )

            else:

                if available_fields:

                    st.info(
                        "Select at least one Order Form field to continue."
                    )

                else:

                    st.warning(
                        "No populated Order Form fields were found."
                    )

    # =====================================================
    # FILE INFORMATION
    # =====================================================

    if excel_file and output_file:

        try:

            output_page_count = get_output_page_count(
                output_file
            )

        except Exception as error:

            output_page_count = 0

            st.warning(
                f"Unable to determine output page count: {error}"
            )

        st.markdown(
            '<div class="section-title">📌 File Information</div>',
            unsafe_allow_html=True
        )

        info1, info2, info3 = st.columns(
            3
        )

        with info1:

            st.metric(
                "Excel Data Rows",
                len(df)
            )

        with info2:

            st.metric(
                "Output Pages",
                output_page_count
            )

        with info3:

            extension = str(
                output_file.name
            ).split(".")[-1].upper()

            st.metric(
                "Output Type",
                extension
            )

        if (
            output_page_count
            and
            len(df) != output_page_count
        ):

            st.warning(
                "⚠️ Excel row count and output page count do not have the same "
                "count. The existing mapping will still use Output Page 1 → "
                "Excel Row 2, Output Page 2 → Excel Row 3, and so on."
            )

        elif output_page_count:

            st.success(
                "✅ Excel rows and output pages match."
            )

    # =====================================================
    # COMPARE BUTTON
    # =====================================================

    if excel_file and output_file:

        st.markdown(
            "<br>",
            unsafe_allow_html=True
        )

        compare_ready = (
            comparison_method is not None
            and
            (
                comparison_method == "Auto Detect"
                or
                bool(selected_fields)
            )
            and
            product_type != "----- SELECT -----"
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

                    # =====================================
                    # EXTRACTION
                    # =====================================

                    output_pages = extract_output_pages(
                        output_file
                    )

                    if not output_pages:

                        raise ValueError(
                            "No readable output pages were detected."
                        )

                    # =====================================
                    # AUTO DETECT
                    # =====================================

                    if comparison_method == "Auto Detect":

                        detected_fields = auto_detect_fields(
                            df,
                            output_pages,
                            product_type
                        )

                        if not detected_fields:

                            st.session_state[
                                "of_report"
                            ] = pd.DataFrame(
                                [
                                    {
                                        "FIELD NO": 1,
                                        "PDF PAGE": "—",
                                        "EXCEL ROW": "—",
                                        "FIELD": "Auto Detect",
                                        "ORDER FORM DATA": "—",
                                        "PDF OUTPUT": "—",
                                        "STATUS": "NOT FOUND",
                                        "DIFFERENCE": (
                                            "No populated artwork-related "
                                            "Order Form fields were detected."
                                        )
                                    }
                                ]
                            )

                        else:

                            st.session_state[
                                "of_report"
                            ] = build_report(
                                df,
                                output_pages,
                                detected_fields,
                                product_type
                            )

                        st.session_state[
                            "of_report_selected_fields"
                        ] = detected_fields

                    # =====================================
                    # SELECTED FIELDS
                    # =====================================

                    else:

                        st.session_state[
                            "of_report"
                        ] = build_report(
                            df,
                            output_pages,
                            selected_fields,
                            product_type
                        )

                        st.session_state[
                            "of_report_selected_fields"
                        ] = selected_fields

                    st.session_state[
                        "of_report_product_type"
                    ] = product_type

                    st.session_state[
                        "of_report_comparison_method"
                    ] = comparison_method

            except Exception as error:

                st.error(
                    f"Unable to process the Output Artwork: {error}"
                )

    # =====================================================
    # REPORT
    # =====================================================

    report = st.session_state.get(
        "of_report"
    )

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

        if (
            report_method == "Auto Detect"
            and
            report_fields
        ):

            st.caption(
                "Auto-detected fields: "
                +
                ", ".join(
                    report_fields
                )
            )

        # =================================================
        # COUNTS
        # =================================================

        pass_count = int(
            (
                report["STATUS"]
                ==
                "PASS"
            ).sum()
        )

        fail_count = int(
            (
                report["STATUS"]
                ==
                "FAIL"
            ).sum()
        )

        not_found_count = int(
            (
                report["STATUS"]
                ==
                "NOT FOUND"
            ).sum()
        )

        skip_count = int(
            (
                report["STATUS"]
                ==
                "SKIP"
            ).sum()
        )

        col1, col2, col3, col4 = st.columns(
            4
        )

        with col1:

            st.metric(
                "PASS",
                pass_count
            )

        with col2:

            st.metric(
                "FAIL",
                fail_count
            )

        with col3:

            st.metric(
                "NOT FOUND",
                not_found_count
            )

        with col4:

            st.metric(
                "IGNORED",
                skip_count
            )

        # =================================================
        # REPORT TABLE
        # =================================================

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

        # =================================================
        # CONCLUSION
        # =================================================

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

        # =================================================
        # HOW IT WORKS
        # =================================================

        with st.expander(
            "ℹ️ How this validation works"
        ):

            st.write(
                """
                **Variable-data validation**

                Only populated fields selected from the Order Form are
                treated as variable artwork data.

                Empty Order Form fields are not shown in the field-selection
                list and are not required for validation.

                **PASS**

                If the expected Order Form value is detected in the relevant
                artwork content, the result is PASS.

                **FAIL**

                If the expected value is not present but a field-specific
                alternative value can be identified, the result is FAIL.

                **NOT FOUND**

                If the expected value cannot be found and there is not enough
                field-specific evidence to identify an alternative, the result
                is NOT FOUND.

                **Important**

                The validator does not compare the entire PDF against the
                entire Excel value using generic fuzzy similarity.

                Different field types use different validation logic so that
                unrelated artwork text does not create false failures.

                **Page mapping**

                PDF Page 1 → Excel Row 2

                PDF Page 2 → Excel Row 3

                PDF Page 3 → Excel Row 4

                and so on.

                **PFL mode**

                Panel-numbered artwork is treated as a continuous stream so
                selected variable data can continue from one panel into the
                next panel.

                **Blank Order Form data**

                A blank Order Form value is ignored and does not generate a
                PASS, FAIL, or NOT FOUND result.
                """
            )

        # =================================================
        # DOWNLOAD
        # =================================================

        csv_data = (
            report
            .to_csv(
                index=False
            )
            .encode(
                "utf-8-sig"
            )
        )

        st.download_button(
            label="⬇️ Download QC Report",
            data=csv_data,
            file_name="PDF_Proofreading_QC_Report.csv",
            mime="text/csv",
            width="stretch",
            key=f"of_download_qc_report_{st.session_state['of_reset_id']}"
        )

    # =====================================================
    # INITIAL INSTRUCTIONS
    # =====================================================

    if not excel_file:

        st.caption(
            "Upload an Order Form to begin."
        )

    elif not output_file:

        st.caption(
            "Upload the Output Artwork to continue."
        )

    elif product_type == "----- SELECT -----":

        st.caption(
            "Select the Product Type to continue."
        )

    elif comparison_method is None:

        st.caption(
            "Select a Comparison Method to continue."
        )

    elif (
        comparison_method == "Select Fields"
        and
        not selected_fields
    ):

        st.caption(
            "Select the variable fields you want to validate."
        )


# =========================================================
# RUN
# =========================================================

if __name__ == "__main__":
    main()
