import streamlit as st
import pandas as pd
import fitz
import re
import unicodedata
from rapidfuzz import fuzz
from difflib import SequenceMatcher


# =========================================================
# SESSION STATE
# =========================================================

if "of_product_type" not in st.session_state:
    st.session_state["of_product_type"] = "Other"

if "of_selected_fields" not in st.session_state:
    st.session_state["of_selected_fields"] = []

if "of_result" not in st.session_state:
    st.session_state["of_result"] = None


# =========================================================
# TOOL CSS
# =========================================================

st.markdown(
    """
    <style>

    .tool-page-title {
        font-size: 34px;
        font-weight: 800;
        margin-bottom: 4px;
        color: #ffffff;
    }

    .tool-page-subtitle {
        color: #aeb8c7;
        font-size: 15px;
        margin-bottom: 28px;
    }

    .tool-section-title {
        font-size: 19px;
        font-weight: 750;
        color: #ffffff;
        margin-top: 10px;
        margin-bottom: 10px;
    }

    .field-info-card {
        padding: 14px 16px;
        border-radius: 14px;
        background: rgba(30, 41, 59, 0.72);
        border: 1px solid rgba(148, 163, 184, 0.20);
        margin-bottom: 12px;
    }

    .field-info-title {
        color: #93c5fd;
        font-size: 13px;
        font-weight: 700;
        letter-spacing: 0.7px;
    }

    .field-info-value {
        color: #ffffff;
        font-size: 15px;
        margin-top: 4px;
    }

    .result-pass {
        color: #4ade80;
        font-weight: 800;
    }

    .result-fail {
        color: #f87171;
        font-weight: 800;
    }

    .result-warning {
        color: #fbbf24;
        font-weight: 800;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# TITLE
# =========================================================

def render_title():

    st.markdown(
        '<div class="tool-page-title">'
        '🔍 Order Form → Output Check'
        '</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="tool-page-subtitle">'
        'Select variable Order Form fields and compare them '
        'against the final PDF artwork.'
        '</div>',
        unsafe_allow_html=True
    )


# =========================================================
# NORMALIZATION
# =========================================================

def normalize_text(text):

    if text is None:
        return ""

    text = str(text)

    text = unicodedata.normalize(
        "NFKC",
        text
    )

    text = text.lower()

    text = text.replace("’", "'")
    text = text.replace("`", "'")
    text = text.replace("–", "-")
    text = text.replace("—", "-")

    text = text.replace("\n", " ")
    text = text.replace("\r", " ")

    # PDF bullet extraction cleanup.
    text = re.sub(
        r"(^|\s)n(?=\s)",
        " ",
        text
    )

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
        text
    )

    text = text.replace(
        "'",
        ""
    )

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
# FIELD TYPE
# =========================================================

def get_field_type(field_name):

    field = normalize_text(
        field_name
    )

    compact = field.replace(
        " ",
        ""
    )

    if (
        "coo" in compact
        or "countryoforigin" in compact
        or "countryorigin" in compact
        or "madein" in compact
        or "origin" in compact
    ):
        return "COO"

    if (
        "fiber" in compact
        or "fibre" in compact
        or "fabric" in compact
        or "content" in compact
        or "composition" in compact
        or "fabrication" in compact
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
        or "fit" in compact
    ):
        return "SIZE"

    if (
        compact == "rn"
        or "registrationnumber" in compact
        or "companyrn" in compact
    ):
        return "RN"

    if "brand" in compact:
        return "BRAND"

    if (
        "color" in compact
        or "colour" in compact
    ):
        return "COLOR"

    if "gender" in compact:
        return "GENDER"

    if (
        "attribute" in compact
        or "technology" in compact
        or "feature" in compact
        or "description" in compact
    ):
        return "ATTRIBUTE"

    return "GENERAL"


# =========================================================
# FIELD REGION
# =========================================================

def get_field_region(field_name):

    field = normalize_text(
        field_name
    )

    compact = field.replace(
        " ",
        ""
    )

    original = str(
        field_name
    ).lower()

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
# EXCEL
# =========================================================

def load_excel(file):

    file.seek(0)

    df = pd.read_excel(
        file,
        header=0
    )

    df = df.dropna(
        axis=0,
        how="all"
    )

    df = df.dropna(
        axis=1,
        how="all"
    )

    df.columns = [
        str(column).strip()
        for column in df.columns
    ]

    return df


# =========================================================
# PDF
# =========================================================

def load_pdf(file):

    file.seek(0)

    pdf_bytes = file.read()

    document = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    pages = []

    for page_number, page in enumerate(
        document
    ):

        text = page.get_text(
            "text"
        )

        pages.append(
            {
                "page": page_number + 1,
                "text": text
            }
        )

    document.close()

    return pages


# =========================================================
# PDF LINE CLEANING
# =========================================================

def clean_pdf_line(line):

    if not line:
        return ""

    line = str(line).strip()

    line = re.sub(
        r"^\s*n\s+(?=[A-Za-z])",
        "",
        line
    )

    return line.strip()


# =========================================================
# STANDARD PDF BLOCKS
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

    # Individual lines.
    for line in lines:
        blocks.append(
            line
        )

    # Adjacent lines for wrapped content.
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
# PFL PANEL LOGIC
# =========================================================

def _panel_number_from_line(line):

    if not line:
        return None

    value = str(
        line
    ).strip()

    match = re.fullmatch(
        r"(?:panel\s*(?:no\.?|number|#)?\s*[-:]?\s*)?(\d{1,3})",
        value,
        flags=re.IGNORECASE
    )

    if not match:
        return None

    number = int(
        match.group(1)
    )

    if 1 <= number <= 99:
        return number

    return None


def split_pfl_panels(page_text):

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

    markers = []

    for index, line in enumerate(lines):

        number = _panel_number_from_line(
            line
        )

        if number is not None:

            markers.append(
                {
                    "index": index,
                    "number": number
                }
            )

    distinct_numbers = []

    for marker in markers:

        if marker["number"] not in distinct_numbers:

            distinct_numbers.append(
                marker["number"]
            )

    if (
        len(distinct_numbers) < 2
        or
        1 not in distinct_numbers
    ):
        return lines

    if len(distinct_numbers) != len(markers):
        return lines

    first_marker_index = markers[0]["index"]

    number_is_above = (
        first_marker_index <= 1
    )

    segments = []

    if number_is_above:

        for position, marker in enumerate(
            markers
        ):

            start = marker["index"] + 1

            if position + 1 < len(markers):

                end = markers[
                    position + 1
                ]["index"]

            else:

                end = len(lines)

            content = lines[
                start:end
            ]

            if content:

                segments.append(
                    {
                        "number": marker["number"],
                        "lines": content
                    }
                )

    else:

        start = 0

        for marker in markers:

            end = marker["index"]

            content = lines[
                start:end
            ]

            if content:

                segments.append(
                    {
                        "number": marker["number"],
                        "lines": content
                    }
                )

            start = (
                marker["index"] + 1
            )

        if start < len(lines):

            segments.append(
                {
                    "number": 999999,
                    "lines": lines[start:]
                }
            )

    if not segments:
        return lines

    segments.sort(
        key=lambda item: item["number"]
    )

    ordered_lines = []

    for segment in segments:

        ordered_lines.extend(
            segment["lines"]
        )

    return ordered_lines


def create_pfl_pdf_blocks(page_text):

    ordered_lines = split_pfl_panels(
        page_text
    )

    if not ordered_lines:
        return []

    blocks = []

    for line in ordered_lines:

        blocks.append(
            line
        )

    maximum = min(
        20,
        len(ordered_lines)
    )

    for size in range(
        2,
        maximum + 1
    ):

        for start in range(
            len(ordered_lines) - size + 1
        ):

            block = " ".join(
                ordered_lines[
                    start:start + size
                ]
            )

            if block:

                blocks.append(
                    block
                )

    complete_sequence = " ".join(
        ordered_lines
    )

    if complete_sequence:

        blocks.append(
            complete_sequence
        )

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

    if expected_normalized in actual_normalized:
        return True

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
# =========================================================

FIELD_ANCHORS = {

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
# RELEVANCE
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
            "bleach"
        ]

        for marker in english_markers:

            if normalize_text(
                marker
            ) in normalized_block:

                return True

    if region == "FR":

        french_markers = [
            "laver",
            "machine",
            "extérieur",
            "doublure",
            "polyester",
            "elasthanne",
            "sans chlore"
        ]

        for marker in french_markers:

            if normalize_text(
                marker
            ) in normalized_block:

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

            if normalize_text(
                marker
            ) in normalized_block:

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

            differences.append(
                f"Extra: {actual_part}"
            )

    if not differences:
        return "Content differs."

    return "; ".join(
        differences[:8]
    )


# =========================================================
# PROBABLE MISMATCH
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

        score = (
            ratio * 0.35
            +
            partial * 0.20
            +
            token_ratio * 0.25
            +
            overlap * 100 * 0.20
        )

        expected_word_count = len(
            expected_tokens
        )

        if field_type == "COO":

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

                if (
                    expected_country
                    and
                    actual_country
                ):

                    country_similarity = fuzz.ratio(
                        expected_country,
                        actual_country
                    )

                    if country_similarity < 95:

                        return {
                            "status": "FAIL",
                            "pdf": block,
                            "difference": get_difference(
                                expected,
                                block
                            ),
                            "score": 100
                        }

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
# CHECK FIELD
# =========================================================

def check_field(
    expected,
    pdf_blocks,
    field_name
):

    if (
        expected is None
        or
        str(expected).strip() == ""
    ):

        return {
            "status": "SKIP",
            "pdf": "—",
            "difference":
                "No variable data in Order Form."
        }

    expected = str(
        expected
    ).strip()

    exact = search_exact_value(
        expected,
        pdf_blocks
    )

    if exact:
        return exact

    probable = search_probable_mismatch(
        expected,
        pdf_blocks,
        field_name
    )

    if probable:
        return probable

    return {
        "status": "NOT FOUND",
        "pdf": "Not found in relevant PDF area",
        "difference":
            "Selected variable value was not detected."
    }


# =========================================================
# REPORT
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

        if excel_index >= len(df):

            for field in selected_fields:

                results.append(
                    {
                        "FIELD NO": field_no,
                        "PDF PAGE": page["page"],
                        "EXCEL ROW": "N/A",
                        "FIELD": field,
                        "ORDER FORM DATA": "No Excel row",
                        "PDF OUTPUT":
                            "No corresponding Order Form row",
                        "STATUS": "NOT FOUND",
                        "DIFFERENCE":
                            "No corresponding Excel row."
                    }
                )

                field_no += 1

            continue

        row = df.iloc[
            excel_index
        ]

        if product_type == "PFL":

            pdf_blocks = create_pfl_pdf_blocks(
                page["text"]
            )

        else:

            pdf_blocks = create_pdf_blocks(
                page["text"]
            )

        for field in selected_fields:

            value = row[field]

            if pd.isna(value):

                value = ""

            else:

                value = str(
                    value
                ).strip()

            if not value:

                results.append(
                    {
                        "FIELD NO": field_no,
                        "PDF PAGE": page["page"],
                        "EXCEL ROW":
                            excel_index + 2,
                        "FIELD": field,
                        "ORDER FORM DATA": "",
                        "PDF OUTPUT": "—",
                        "STATUS": "SKIP",
                        "DIFFERENCE":
                            "Blank Order Form value — "
                            "PDF content ignored."
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
                    "EXCEL ROW":
                        excel_index + 2,
                    "FIELD": field,
                    "ORDER FORM DATA": value,
                    "PDF OUTPUT": result["pdf"],
                    "STATUS": result["status"],
                    "DIFFERENCE":
                        result["difference"]
                }
            )

            field_no += 1

    return pd.DataFrame(
        results
    )


# =========================================================
# STATUS STYLING
# =========================================================

def style_status(value):

    if value == "PASS":

        return (
            "background-color:#238636;"
            "color:white;"
            "font-weight:bold;"
        )

    if value == "FAIL":

        return (
            "background-color:#da3633;"
            "color:white;"
            "font-weight:bold;"
        )

    if value == "NOT FOUND":

        return (
            "background-color:#9e6a03;"
            "color:white;"
            "font-weight:bold;"
        )

    if value == "SKIP":

        return (
            "background-color:#555555;"
            "color:white;"
            "font-weight:bold;"
        )

    return ""


# =========================================================
# MAIN
# =========================================================

def main():

    render_title()

    # ======================================================
    # NEW START
    # ======================================================

    top_left, top_right = st.columns(
        [7, 1]
    )

    with top_right:

        if st.button(
            "↻ NEW START",
            key="of_new_start",
            width="stretch"
        ):

            st.session_state[
                "of_product_type"
            ] = "Other"

            st.session_state[
                "of_selected_fields"
            ] = []

            st.session_state[
                "of_result"
            ] = None

            st.rerun()

    # ======================================================
    # PRODUCT TYPE
    # ======================================================

    st.markdown(
        '<div class="tool-section-title">'
        '🏷️ Product Type'
        '</div>',
        unsafe_allow_html=True
    )

    product_types = [
        "Other",
        "HTL",
        "PFL"
    ]

    current_type = st.session_state.get(
        "of_product_type",
        "Other"
    )

    if current_type not in product_types:
        current_type = "Other"

    product_type = st.selectbox(
        "Product Type",
        product_types,
        index=product_types.index(
            current_type
        ),
        key="of_product_type"
    )

    if product_type == "PFL":

        st.info(
            "PFL mode enabled — panel sequence and "
            "cross-panel continuation logic will be used."
        )

    else:

        st.caption(
            "Standard comparison mode."
        )

    # ======================================================
    # FILE UPLOADS
    # ======================================================

    st.markdown(
        '<div class="tool-section-title">'
        '📂 Upload Files'
        '</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(
        2
    )

    with col1:

        excel_file = st.file_uploader(
            "📊 Order Form Excel",
            type=[
                "xlsx",
                "xls"
            ],
            key="of_excel_upload"
        )

    with col2:

        pdf_file = st.file_uploader(
            "📄 Output Artwork PDF",
            type=[
                "pdf"
            ],
            key="of_pdf_upload"
        )

    # ======================================================
    # READ EXCEL
    # ======================================================

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

        if df.empty:

            st.error(
                "The uploaded Excel does not contain usable data."
            )

            return

        # ==================================================
        # EXCEL INFORMATION
        # ==================================================

        st.markdown(
            '<div class="tool-section-title">'
            '📌 Order Form Fields'
            '</div>',
            unsafe_allow_html=True
        )

        st.caption(
            "Select ONLY the Excel fields that should be "
            "validated against the PDF. Nothing is selected "
            "automatically."
        )

        excel_columns = [
            str(column)
            for column in df.columns
        ]

        selected_fields = st.multiselect(
            "Select fields to validate",
            options=excel_columns,
            default=st.session_state.get(
                "of_selected_fields",
                []
            ),
            key="of_selected_fields"
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
                        values.append(
                            value
                        )

                preview_rows.append(
                    {
                        "Excel Field": field,
                        "Values": len(values),
                        "Preview":
                            " | ".join(
                                values[:3]
                            )
                    }
                )

            preview_df = pd.DataFrame(
                preview_rows
            )

            with st.expander(
                "🔎 Preview Selected Fields"
            ):

                st.dataframe(
                    preview_df,
                    width="stretch",
                    hide_index=True
                )

        else:

            st.info(
                "Select at least one Excel field to continue."
            )

    else:

        selected_fields = []

    # ======================================================
    # REQUIRE BOTH FILES
    # ======================================================

    if not excel_file:

        return

    if not pdf_file:

        st.info(
            "Upload the Output Artwork PDF to continue."
        )

        return

    if not selected_fields:

        return

    # ======================================================
    # READ PDF
    # ======================================================

    try:

        pdf_pages = load_pdf(
            pdf_file
        )

    except Exception as error:

        st.error(
            f"Unable to read the PDF: {error}"
        )

        return

    if not pdf_pages:

        st.error(
            "No pages could be read from the PDF."
        )

        return

    # ======================================================
    # FILE INFORMATION
    # ======================================================

    st.markdown(
        '<div class="tool-section-title">'
        '📌 Validation Setup'
        '</div>',
        unsafe_allow_html=True
    )

    info1, info2, info3 = st.columns(
        3
    )

    with info1:

        st.metric(
            "Excel Rows",
            len(df)
        )

    with info2:

        st.metric(
            "PDF Pages",
            len(pdf_pages)
        )

    with info3:

        st.metric(
            "Selected Fields",
            len(selected_fields)
        )

    if len(df) != len(pdf_pages):

        st.warning(
            "Excel row count and PDF page count do not match. "
            "The existing mapping will still be used: "
            "PDF Page 1 → Excel Row 2, PDF Page 2 → Excel Row 3, "
            "and so on."
        )

    else:

        st.success(
            "Excel row count and PDF page count match."
        )

    # ======================================================
    # COMPARE
    # ======================================================

    st.markdown(
        '<div class="tool-section-title">'
        '🚀 Run Validation'
        '</div>',
        unsafe_allow_html=True
    )

    compare_clicked = st.button(
        "🔍  COMPARE & PROOFREAD",
        key="of_compare",
        type="primary",
        width="stretch"
    )

    if not compare_clicked:

        return

    # ======================================================
    # RUN
    # ======================================================

    with st.spinner(
        "Checking selected variable artwork data..."
    ):

        report = build_report(
            df,
            pdf_pages,
            selected_fields,
            product_type
        )

    st.session_state[
        "of_result"
    ] = report

    # ======================================================
    # REPORT
    # ======================================================

    st.divider()

    st.markdown(
        '<div class="tool-section-title">'
        '📋 QC Report'
        '</div>',
        unsafe_allow_html=True
    )

    if report.empty:

        st.warning(
            "No validation results were generated."
        )

        return

    pass_count = int(
        (
            report["STATUS"] == "PASS"
        ).sum()
    )

    fail_count = int(
        (
            report["STATUS"] == "FAIL"
        ).sum()
    )

    not_found_count = int(
        (
            report["STATUS"] == "NOT FOUND"
        ).sum()
    )

    skip_count = int(
        (
            report["STATUS"] == "SKIP"
        ).sum()
    )

    # ======================================================
    # SUMMARY
    # ======================================================

    c1, c2, c3, c4 = st.columns(
        4
    )

    with c1:

        st.metric(
            "PASS",
            pass_count
        )

    with c2:

        st.metric(
            "FAIL",
            fail_count
        )

    with c3:

        st.metric(
            "NOT FOUND",
            not_found_count
        )

    with c4:

        st.metric(
            "IGNORED",
            skip_count
        )

    # ======================================================
    # TABLE
    # ======================================================

    styled_report = (
        report
        .style
        .map(
            style_status,
            subset=[
                "STATUS"
            ]
        )
    )

    st.dataframe(
        styled_report,
        width="stretch",
        hide_index=True
    )

    # ======================================================
    # CONCLUSION
    # ======================================================

    st.divider()

    if fail_count > 0:

        st.error(
            f"❌ FAIL — {fail_count} variable-data "
            f"mismatch(es) detected."
        )

    elif not_found_count > 0:

        st.warning(
            f"⚠️ REVIEW — {not_found_count} selected "
            f"variable field(s) could not be located."
        )

    else:

        st.success(
            "✅ PASS — All selected variable fields "
            "matched the PDF artwork."
        )

    # ======================================================
    # DIFFERENCE DETAILS
    # ======================================================

    failures = report[
        report["STATUS"].isin(
            [
                "FAIL",
                "NOT FOUND"
            ]
        )
    ]

    if not failures.empty:

        st.markdown(
            "### 🔎 Difference Details"
        )

        for _, result in failures.iterrows():

            field_name = result[
                "FIELD"
            ]

            status = result[
                "STATUS"
            ]

            if status == "FAIL":

                title = (
                    f"❌ {field_name} "
                    f"— Page {result['PDF PAGE']}"
                )

            else:

                title = (
                    f"⚠️ {field_name} "
                    f"— Page {result['PDF PAGE']}"
                )

            with st.expander(
                title
            ):

                left, right = st.columns(
                    2
                )

                with left:

                    st.markdown(
                        "**Order Form Data**"
                    )

                    st.code(
                        str(
                            result[
                                "ORDER FORM DATA"
                            ]
                        )
                    )

                with right:

                    st.markdown(
                        "**PDF Output**"
                    )

                    st.code(
                        str(
                            result[
                                "PDF OUTPUT"
                            ]
                        )
                    )

                st.markdown(
                    "**Difference**"
                )

                difference = str(
                    result[
                        "DIFFERENCE"
                    ]
                )

                if (
                    "→" in difference
                ):

                    st.error(
                        difference
                    )

                else:

                    st.warning(
                        difference
                    )

    # ======================================================
    # LOGIC EXPLANATION
    # ======================================================

    with st.expander(
        "ℹ️ How validation works"
    ):

        st.write(
            """
            **Variable-data validation**

            Only fields selected from the Order Form are
            treated as variable artwork data. Unselected
            static artwork text is not used to generate
            mismatches.

            **Page mapping**

            PDF Page 1 → Excel Row 2

            PDF Page 2 → Excel Row 3

            PDF Page 3 → Excel Row 4

            and so on.

            **Blank fields**

            If the selected Order Form field is blank,
            that field is ignored.

            **PFL**

            PFL mode detects panel numbering when possible
            and allows variable data to continue across
            panel boundaries.

            **Difference detection**

            When an expected value is found exactly,
            the result is PASS.

            When a relevant alternative value is found,
            the result is FAIL and the report shows the
            difference.

            Example:

            Order Form:
            MADE IN CHINA

            PDF:
            MADE IN VIETNAM

            Difference:
            CHINA → VIETNAM
            """
        )

    # ======================================================
    # DOWNLOAD REPORT
    # ======================================================

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
        file_name=(
            "Order_Form_Output_QC_Report.csv"
        ),
        mime="text/csv",
        width="stretch"
    )
