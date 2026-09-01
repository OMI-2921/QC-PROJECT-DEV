import streamlit as st
import pandas as pd
import fitz
import re
import unicodedata
import io

from PIL import Image
from rapidfuzz import fuzz
from difflib import SequenceMatcher

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

def normalize_text(text):

    if text is None:
        return ""

    text = str(text)

    text = unicodedata.normalize(
        "NFKC",
        text
    )

    text = text.lower()

    # -----------------------------------------------------
    # Normalize common PDF characters
    # -----------------------------------------------------

    text = text.replace("’", "'")
    text = text.replace("`", "'")
    text = text.replace("–", "-")
    text = text.replace("—", "-")

    # -----------------------------------------------------
    # PDF line breaks
    # -----------------------------------------------------

    text = text.replace("\n", " ")
    text = text.replace("\r", " ")

    # -----------------------------------------------------
    # IMPORTANT:
    #
    # The PDF may extract bullets as "n".
    # We DO NOT want that "n" to become real content.
    # -----------------------------------------------------

    text = re.sub(
        r"(^|\s)n(?=\s)",
        " ",
        text
    )

    # -----------------------------------------------------
    # Separators
    # -----------------------------------------------------

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

    # -----------------------------------------------------
    # Remove remaining punctuation
    # -----------------------------------------------------

    text = re.sub(
        r"[^\w%#'\s]",
        " ",
        text
    )

    # -----------------------------------------------------
    # Apostrophe differences ignored
    # -----------------------------------------------------

    text = text.replace(
        "'",
        ""
    )

    # -----------------------------------------------------
    # Multiple spaces
    # -----------------------------------------------------

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


def compact_text(text):

    return normalize_text(
        text
    ).replace(
        " ",
        ""
    )


def tokenize(text):

    value = normalize_text(
        text
    )

    if not value:
        return []

    return value.split()


# =========================================================
# FIELD TYPE DETECTION
#
# This is NOT deciding what to check.
# The user-selected field is always checked.
#
# This only helps us locate the correct variable
# information inside the PDF.
# =========================================================

def get_field_type(field_name):

    field = normalize_text(
        field_name
    )

    compact = field.replace(
        " ",
        ""
    )

    # Common variable identifiers / production data
    # These are useful in Auto Detect mode.
    if (
        "sku" in compact
        or "itemcode" in compact
        or "itemnumber" in compact
        or "itemno" in compact
        or "stylecode" in compact
        or "style" == compact
    ):
        return "IDENTIFIER"

    if (
        "batch" in compact
        or "lotnumber" in compact
        or "lotno" in compact
        or "lot" == compact
    ):
        return "BATCH"

    if (
        "quantity" in compact
        or "qty" == compact
        or "units" in compact
        or "pieces" in compact
        or "pcs" == compact
    ):
        return "QUANTITY"

    # Country of Origin

    if (
        "coo" in compact
        or "countryoforigin" in compact
        or "countryorigin" in compact
        or "madein" in compact
        or "origin" in compact
    ):

        return "COO"


    # Fiber / Fabric / Content

    if (
        "fiber" in compact
        or "fibre" in compact
        or "fabric" in compact
        or "content" in compact
        or "composition" in compact
        or "fabrication" in compact
    ):

        return "CONTENT"


    # Care / Washing

    if (
        "care" in compact
        or "wash" in compact
        or "washing" in compact
        or "laundry" in compact
        or "instruction" in compact
    ):

        return "CARE"


    # Size

    if (
        "size" in compact
        or "sizeline" in compact
        or "alpha" in compact
        or "waist" in compact
        or "inseam" in compact
        or "fit" in compact
    ):

        return "SIZE"


    # RN

    if (
        compact == "rn"
        or "registrationnumber" in compact
        or "companyrn" in compact
    ):

        return "RN"


    # Brand

    if (
        "brand" in compact
    ):

        return "BRAND"


    # Color

    if (
        "color" in compact
        or "colour" in compact
    ):

        return "COLOR"


    # Gender

    if (
        "gender" in compact
    ):

        return "GENDER"


    # Attribute

    if (
        "attribute" in compact
        or "technology" in compact
        or "feature" in compact
        or "description" in compact
    ):

        return "ATTRIBUTE"


    return "GENERAL"


# =========================================================
# FIELD LANGUAGE / REGION DETECTION
# =========================================================

def get_field_region(field_name):

    field = normalize_text(
        field_name
    )

    compact = field.replace(
        " ",
        ""
    )

    # English

    if (
        "_en" in str(field_name).lower()
        or compact.endswith("en")
        or "english" in compact
    ):

        return "EN"


    # French / Canada

    if (
        "_fr" in str(field_name).lower()
        or compact.endswith("fr")
        or "french" in compact
        or "canada" in compact
    ):

        return "FR"


    # Spanish

    if (
        "_sp" in str(field_name).lower()
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
# OUTPUT EXTRACTION LAYER
#
# Extraction is deliberately separate from comparison.
# The comparison engine always receives the same structure:
#     [{"page": 1, "text": "..."}, ...]
#
# Readable PDF  -> direct PDF text extraction
# Scanned PDF   -> OCR only when the page has no usable text
# JPG / PNG     -> OCR
# =========================================================

def _usable_text(text):
    """Return True when a PDF page appears to have a usable text layer."""

    if not text:
        return False

    value = str(text).strip()
    if not value:
        return False

    # Ignore pages whose text layer contains only a few stray characters.
    alnum = re.sub(r"[^\w%#]", "", value, flags=re.UNICODE)
    return len(alnum) >= 6


def _ocr_image(image):
    """OCR one PIL image. OCR stays completely behind the scenes."""

    try:
        import pytesseract
    except ImportError as exc:
        raise RuntimeError(
            "OCR support is not installed. Add pytesseract to requirements.txt."
        ) from exc

    try:
        # English + French + Spanish cover the multilingual artwork commonly
        # encountered by this validator. Tesseract uses whichever installed
        # language data is available.
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
    """Render a PDF page at a useful OCR resolution."""

    pixmap = page.get_pixmap(
        matrix=fitz.Matrix(2.5, 2.5),
        alpha=False
    )

    image = Image.open(
        io.BytesIO(pixmap.tobytes("png"))
    ).convert("RGB")

    return image


def get_output_page_count(file):
    """Get page count without extracting/comparing output text."""

    name = str(getattr(file, "name", "")).lower()

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

    if name.endswith((".jpg", ".jpeg", ".png")):
        return 1

    return 0


def extract_output_pages(file):
    """
    Extract output into the common comparison representation.

    This function is called ONLY after the user presses COMPARE.
    """

    name = str(getattr(file, "name", "")).lower()

    # ---------------------------------------------------------
    # READABLE PDF / SCANNED PDF
    # ---------------------------------------------------------
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
                # Normal/readable PDF: use the existing direct extraction path.
                text = direct_text
                source_type = "pdf_text"
            else:
                # No usable text layer: render and OCR this page only.
                image = _render_pdf_page(page)
                text = _ocr_image(image)
                source_type = "ocr"

            pages.append({
                "page": page_number + 1,
                "text": text,
                "source_type": source_type
            })

        document.close()
        return pages

    # ---------------------------------------------------------
    # JPG / JPEG / PNG
    # ---------------------------------------------------------
    if name.endswith((".jpg", ".jpeg", ".png")):

        file.seek(0)
        image = Image.open(file).convert("RGB")

        text = _ocr_image(image)

        return [{
            "page": 1,
            "text": text,
            "source_type": "ocr"
        }]

    raise ValueError(
        "Unsupported output format. Please upload PDF, JPG, JPEG, or PNG."
    )


# Backward-compatible name for any internal code that still calls load_pdf.
# It now supports the same PDF behavior as the new extraction layer.
def load_pdf(file):
    return extract_output_pages(file)


# =========================================================
# CLEAN PDF LINE
# =========================================================

def clean_pdf_line(line):

    if not line:
        return ""

    line = str(line).strip()

    # -----------------------------------------------------
    # Remove PDF bullet/keystroke "n"
    #
    # Examples:
    #
    # n US :
    # n CA :
    # n MX :
    #
    # becomes:
    #
    # US :
    # CA :
    # MX :
    # -----------------------------------------------------

    line = re.sub(
        r"^\s*n\s+(?=[A-Za-z])",
        "",
        line
    )

    return line.strip()


# =========================================================
# CREATE SMART PDF BLOCKS
# =========================================================

def create_pdf_blocks(page_text):

    if not page_text:
        return []

    raw_lines = page_text.splitlines()

    lines = []

    for raw_line in raw_lines:

        line = clean_pdf_line(
            raw_line
        )

        if not line:
            continue

        if len(line) > 1500:
            continue

        lines.append(
            line
        )


    if not lines:
        return []


    blocks = []


    # -----------------------------------------------------
    # INDIVIDUAL LINES
    #
    # Important because many variable fields are contained
    # inside one PDF line.
    # -----------------------------------------------------

    for line in lines:

        blocks.append(
            line
        )


    # -----------------------------------------------------
    # ADJACENT LINE BLOCKS
    #
    # Used for wrapped care/content text.
    # -----------------------------------------------------

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

                blocks.append(
                    block
                )


    # -----------------------------------------------------
    # Full page is deliberately NOT used as the primary
    # comparison block.
    #
    # This prevents unrelated static/regional information
    # from being mixed with variable data.
    # -----------------------------------------------------


    # -----------------------------------------------------
    # Remove duplicates
    # -----------------------------------------------------

    unique = []

    seen = set()

    for block in blocks:

        normalized = normalize_text(
            block
        )

        if not normalized:
            continue

        if normalized in seen:
            continue

        seen.add(
            normalized
        )

        unique.append(
            block
        )

    return unique


# =========================================================
# EXACT MATCH
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


    # Normal match

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


# =========================================================
# EXPECTED VALUE SEARCH
# =========================================================

def search_exact_value(
    expected,
    pdf_blocks
):

    for block in pdf_blocks:

        if exact_match(
            expected,
            block
        ):

            return {
                "status": "PASS",
                "pdf": block,
                "difference": "—",
                "score": 100
            }

    return None


# =========================================================
# FIELD ANCHORS
#
# These help us locate the relevant variable value.
# They prevent static information from being treated as
# a mismatch.
# =========================================================

FIELD_ANCHORS = {

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

    "COO": [
        "made in",
        "hecho en",
        "fabrique en",
        "madein"
    ],

    "CONTENT": [
        "shell",
        "liner",
        "body",
        "fabric",
        "fiber",
        "fibre",
        "content",
        "composition",
        "exterior",
        "extérieur",
        "forro",
        "doublure"
    ],

    "CARE": [
        "machine wash",
        "wash",
        "lavar",
        "laver",
        "dry clean",
        "bleach",
        "blanchiment",
        "detergent",
        "detergente"
    ],

    "RN": [
        "rn",
        "ca"
    ],

    "SIZE": [
        "size"
    ],

    "COLOR": [
        "color",
        "colour"
    ],

    "BRAND": [
        "brand"
    ],

    "GENDER": [
        "girls",
        "boys",
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
# CHECK IF BLOCK IS RELEVANT TO FIELD
# =========================================================

def is_relevant_block(
    block,
    field_type,
    field_name,
    expected
):

    normalized_block = normalize_text(
        block
    )

    if not normalized_block:
        return False


    # -----------------------------------------------------
    # Field-specific anchors
    # -----------------------------------------------------

    anchors = FIELD_ANCHORS.get(
        field_type,
        []
    )


    for anchor in anchors:

        anchor_norm = normalize_text(
            anchor
        )

        if (
            anchor_norm
            and
            anchor_norm in normalized_block
        ):

            return True


    # -----------------------------------------------------
    # Language-specific clues
    #
    # Only used when the selected field name explicitly
    # indicates a language.
    # -----------------------------------------------------

    region = get_field_region(
        field_name
    )


    if region == "EN":

        english_markers = [
            "made in",
            "machine wash",
            "shell",
            "liner",
            "polyester",
            "span dex",
            "bleach"
        ]

        for marker in english_markers:

            if normalize_text(marker) in normalized_block:

                return True


    if region == "FR":

        french_markers = [
            "laver",
            "machine",
            "extérieur",
            "doublure",
            "polyester",
            "élasthanne",
            "sans chlore"
        ]

        for marker in french_markers:

            if normalize_text(marker) in normalized_block:

                return True


    if region == "SP":

        spanish_markers = [
            "lavar",
            "máquina",
            "maquina",
            "cuerpo",
            "forro",
            "poliéster",
            "poliester",
            "cloro",
            "hecho en"
        ]

        for marker in spanish_markers:

            if normalize_text(marker) in normalized_block:

                return True


    return False


# =========================================================
# TOKEN OVERLAP
# =========================================================

def token_overlap(
    expected,
    actual
):

    expected_tokens = set(
        tokenize(expected)
    )

    actual_tokens = set(
        tokenize(actual)
    )

    if not expected_tokens:
        return 0

    common = (
        expected_tokens
        &
        actual_tokens
    )

    return (
        len(common)
        /
        len(expected_tokens)
    )


# =========================================================
# DIFFERENCE
# =========================================================

def get_difference(
    expected,
    actual
):

    expected_tokens = tokenize(
        expected
    )

    actual_tokens = tokenize(
        actual
    )

    if not expected_tokens:
        return "Content differs."

    if not actual_tokens:
        return "Expected value is missing from PDF."

    matcher = SequenceMatcher(
        None,
        expected_tokens,
        actual_tokens
    )

    differences = []

    for tag, a1, a2, b1, b2 in matcher.get_opcodes():

        if tag == "equal":
            continue


        expected_part = " ".join(
            expected_tokens[a1:a2]
        )

        actual_part = " ".join(
            actual_tokens[b1:b2]
        )


        if tag == "replace":

            differences.append(
                f"{expected_part} → {actual_part}"
            )


        elif tag == "delete":

            differences.append(
                f"Missing: {expected_part}"
            )


        elif tag == "insert":

            # IMPORTANT:
            #
            # Extra static PDF content should NOT become
            # a failure automatically.
            #
            # Therefore inserted words are only shown when
            # they occur inside a relevant candidate.

            differences.append(
                f"Extra: {actual_part}"
            )


    if not differences:

        return "Content differs."

    return "; ".join(
        differences[:8]
    )


# =========================================================
# FIND PROBABLE VARIABLE MISMATCH
#
# This is the most important new function.
#
# We DO NOT compare the Order Form value against the
# entire PDF.
#
# We only consider PDF blocks relevant to the selected
# variable field.
# =========================================================

def search_probable_mismatch(
    expected,
    pdf_blocks,
    field_name
):

    expected_normalized = normalize_text(
        expected
    )

    expected_tokens = tokenize(
        expected
    )

    if not expected_tokens:
        return None


    field_type = get_field_type(
        field_name
    )


    candidates = []


    # -----------------------------------------------------
    # First: field-relevant blocks only
    # -----------------------------------------------------

    for block in pdf_blocks:

        if is_relevant_block(
            block,
            field_type,
            field_name,
            expected
        ):

            candidates.append(
                block
            )


    # -----------------------------------------------------
    # If no relevant block exists:
    #
    # DO NOT start comparing against the entire PDF.
    #
    # This is what prevents static data from creating
    # false failures.
    # -----------------------------------------------------

    if not candidates:

        return None


    best = None


    for block in candidates:

        actual_normalized = normalize_text(
            block
        )

        actual_tokens = tokenize(
            block
        )

        if not actual_tokens:
            continue


        # -------------------------------------------------
        # Similarity
        # -------------------------------------------------

        ratio = fuzz.ratio(
            expected_normalized,
            actual_normalized
        )

        partial = fuzz.partial_ratio(
            expected_normalized,
            actual_normalized
        )

        token_ratio = fuzz.token_set_ratio(
            expected_normalized,
            actual_normalized
        )

        overlap = token_overlap(
            expected,
            block
        )


        # -------------------------------------------------
        # Score
        # -------------------------------------------------

        score = (
            ratio * 0.35
            +
            partial * 0.20
            +
            token_ratio * 0.25
            +
            overlap * 100 * 0.20
        )


        # -------------------------------------------------
        # Determine whether this is actually a useful
        # mismatch candidate.
        # -------------------------------------------------

        expected_word_count = len(
            expected_tokens
        )


        if field_type == "COO":

            # COO is special.
            #
            # If the PDF has "MADE IN VIETNAM" while the
            # Order Form says "MADE IN CHINA", this should
            # be a mismatch.
            #
            # "MADE IN" gives us the strong contextual anchor.

            expected_has_made_in = (
                "made in"
                in
                expected_normalized
            )

            actual_has_made_in = (
                "made in"
                in
                actual_normalized
            )


            if (
                expected_has_made_in
                and
                actual_has_made_in
            ):

                # Extract everything after "made in"

                expected_country = re.sub(
                    r"^.*?made in\s+",
                    "",
                    expected_normalized
                ).strip()

                actual_country = re.sub(
                    r"^.*?made in\s+",
                    "",
                    actual_normalized
                ).strip()


                # Country comparison

                if (
                    expected_country
                    and
                    actual_country
                ):

                    country_similarity = fuzz.ratio(
                        expected_country,
                        actual_country
                    )


                    if (
                        country_similarity < 95
                    ):

                        return {
                            "status": "FAIL",
                            "pdf": block,
                            "difference": get_difference(
                                expected,
                                block
                            ),
                            "score": 100
                        }


        # -------------------------------------------------
        # General mismatch rules
        # -------------------------------------------------

        if expected_word_count <= 2:

            acceptable = (
                score >= 72
                and
                overlap >= 0.35
            )

        elif expected_word_count <= 5:

            acceptable = (
                score >= 68
                and
                overlap >= 0.35
            )

        else:

            acceptable = (
                score >= 65
                and
                overlap >= 0.30
            )


        if not acceptable:
            continue


        # -------------------------------------------------
        # Keep best candidate
        # -------------------------------------------------

        if (
            best is None
            or
            score > best["score"]
        ):

            best = {
                "status": "FAIL",
                "pdf": block,
                "difference": get_difference(
                    expected,
                    block
                ),
                "score": score
            }


    return best


# =========================================================
# CHECK ONE VARIABLE FIELD
# =========================================================

def check_field(
    expected,
    pdf_blocks,
    field_name
):

    # -----------------------------------------------------
    # Blank Order Form value
    #
    # IMPORTANT:
    #
    # If the variable field has no value in the Order Form,
    # it is NOT required.
    # -----------------------------------------------------

    if (
        expected is None
        or
        str(expected).strip() == ""
    ):

        return {
            "status": "SKIP",
            "pdf": "—",
            "difference": "No variable data in Order Form."
        }


    expected = str(
        expected
    ).strip()


    # -----------------------------------------------------
    # 1. EXACT VARIABLE VALUE
    # -----------------------------------------------------

    exact = search_exact_value(
        expected,
        pdf_blocks
    )

    if exact:

        return exact


    # -----------------------------------------------------
    # 2. SEARCH FOR A RELEVANT DIFFERENT VALUE
    #
    # Example:
    #
    # Order Form:
    # MADE IN CHINA
    #
    # PDF:
    # MADE IN VIETNAM
    #
    # Result:
    # FAIL
    # -----------------------------------------------------

    probable = search_probable_mismatch(
        expected,
        pdf_blocks,
        field_name
    )

    if probable:

        return probable


    # -----------------------------------------------------
    # 3. NOT FOUND
    #
    # This means:
    #
    # - The expected variable value is absent
    # - AND we did not find a sufficiently relevant
    #   alternative variable value
    #
    # We do NOT call unrelated static PDF information
    # a FAIL.
    # -----------------------------------------------------

    return {
        "status": "NOT FOUND",
        "pdf": "Not found in relevant PDF area",
        "difference": "Selected variable value was not detected."
    }


# =========================================================
# PFL PANEL SUPPORT
# =========================================================

def get_pfl_panel_blocks(page_text):
    """
    PFL mode treats panel-numbered artwork as a continuous stream.
    Panel labels such as 1, 2, 3 or PANEL 1 / PANEL 2 are removed
    from the comparison content. The extracted lines are then grouped
    into larger blocks so text split at a panel boundary can be found.
    """
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

        # Ignore standalone panel sequence numbers.
        if panel_pattern.match(line):
            continue

        # Ignore common textual panel labels.
        if re.match(r"^\s*panel\s*[-#:]?\s*\d{1,3}\s*$", line, re.IGNORECASE):
            continue

        cleaned.append(line)

    if not cleaned:
        return []

    blocks = list(cleaned)

    # Larger windows are the important part for PFL because a value
    # may begin near the end of one panel and continue into the next.
    maximum = min(24, len(cleaned))

    for size in range(2, maximum + 1):
        for start in range(len(cleaned) - size + 1):
            block = " ".join(cleaned[start:start + size])
            if block:
                blocks.append(block)

    unique = []
    seen = set()

    for block in blocks:
        normalized = normalize_text(block)
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        unique.append(block)

    return unique


def get_comparison_blocks(page_text, product_type):
    if product_type == "PFL":
        return get_pfl_panel_blocks(page_text)
    return create_pdf_blocks(page_text)


# =========================================================
# AUTO DETECT MODE
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


def _auto_field_is_candidate(field_name):
    """Keep Auto Detect focused on fields that can reasonably appear in artwork."""

    compact = normalize_text(field_name).replace(" ", "").replace("_", "").replace("-", "")

    if not compact:
        return False

    if compact in AUTO_IGNORED_FIELD_WORDS:
        return False

    # Explicitly recognized variable/business fields are candidates.
    if get_field_type(field_name) != "GENERAL":
        return True

    # Common artwork variable names that are not covered by field typing.
    keywords = (
        "sku", "item", "style", "batch", "lot", "qty", "quantity",
        "number", "code", "description", "attribute", "value"
    )

    return any(word in compact for word in keywords)


def auto_detect_fields(df, output_pages, product_type):
    """
    Identify Order Form columns that have evidence of corresponding artwork data.

    We do not blindly select every Excel column. A field must have at least one
    nonblank value and evidence in the output through exact matching or a strong
    field-aware candidate. This keeps administrative Excel columns out of the QC.
    """

    candidates = [
        str(column)
        for column in df.columns
        if _auto_field_is_candidate(str(column))
    ]

    detected = []

    # Build page blocks once. No comparison happens before COMPARE because this
    # function itself is called only from the Compare button handler.
    page_blocks = [
        get_comparison_blocks(page.get("text", ""), product_type)
        for page in output_pages
    ]

    for field in candidates:
        found_evidence = False

        for page_index, blocks in enumerate(page_blocks):
            if page_index >= len(df):
                break

            value = df.iloc[page_index][field]

            if pd.isna(value) or not str(value).strip():
                continue

            value = str(value).strip()

            # Exact value is the strongest evidence that this column belongs
            # to the output artwork.
            if search_exact_value(value, blocks):
                found_evidence = True
                break

            # Strong field-aware mismatch evidence means the value was likely
            # intended for artwork even if the printed value is wrong.
            probable = search_probable_mismatch(
                value,
                blocks,
                field
            )

            if probable:
                found_evidence = True
                break

        if found_evidence:
            detected.append(field)

    return detected


# =========================================================
# BUILD REPORT
#
# PAGE 1 → EXCEL ROW 2
# PAGE 2 → EXCEL ROW 3
# PAGE 3 → EXCEL ROW 4
# =========================================================

def build_report(
    df,
    pdf_pages,
    selected_fields,
    product_type
):

    results = []

    field_no = 1


    # -----------------------------------------------------
    # Process PDF pages.
    #
    # PDF page index 0 = Excel row 2
    # PDF page index 1 = Excel row 3
    # etc.
    # -----------------------------------------------------

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


        # -------------------------------------------------
        # Get corresponding Excel row
        # -------------------------------------------------

        row = df.iloc[
            excel_index
        ]


        # -------------------------------------------------
        # Extract smart PDF blocks
        # -------------------------------------------------

        pdf_blocks = get_comparison_blocks(
            page["text"],
            product_type
        )


        # -------------------------------------------------
        # Compare ONLY user-selected fields
        # -------------------------------------------------

        for field in selected_fields:

            value = row[field]


            if pd.isna(value):

                value = ""

            else:

                value = str(
                    value
                ).strip()


            # -------------------------------------------------
            # Blank variable field = NOT REQUIRED
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
                        "DIFFERENCE": "Blank Order Form value — PDF content ignored."
                    }
                )

                field_no += 1

                continue


            # -------------------------------------------------
            # Compare variable field
            # -------------------------------------------------

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
                "Only the fields selected below will participate in the comparison."
            )

            selected_fields = st.multiselect(
                "Select the fields from your Order Form",
                options=[str(column) for column in df.columns],
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

        csv_data = (
            report
            .to_csv(index=False)
            .encode("utf-8-sig")
        )

        st.download_button(
            label="⬇️ Download QC Report",
            data=csv_data,
            file_name="PDF_Proofreading_QC_Report.csv",
            mime="text/csv",
            width="stretch",
            key=f"of_download_qc_report_{st.session_state['of_reset_id']}"
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
