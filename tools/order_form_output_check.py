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
    .tool-page-title { font-size:34px; font-weight:800; margin-bottom:4px; color:#ffffff; }
    .tool-page-subtitle { color:#aeb8c7; font-size:15px; margin-bottom:28px; }
    .tool-section-title { font-size:19px; font-weight:750; color:#ffffff; margin-top:10px; margin-bottom:10px; }
    .field-info-card { padding:14px 16px; border-radius:14px; background:rgba(30,41,59,.72); border:1px solid rgba(148,163,184,.20); margin-bottom:12px; }
    .field-info-title { color:#93c5fd; font-size:13px; font-weight:700; letter-spacing:.7px; }
    .field-info-value { color:#ffffff; font-size:15px; margin-top:4px; }
    .result-pass { color:#4ade80; font-weight:800; }
    .result-fail { color:#f87171; font-weight:800; }
    .result-warning { color:#fbbf24; font-weight:800; }
    </style>
    """,
    unsafe_allow_html=True,
)


def render_title():
    st.markdown('<div class="tool-page-title">🔍 Order Form → Output Check</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="tool-page-subtitle">Select variable Order Form fields and compare them against the final PDF artwork.</div>',
        unsafe_allow_html=True,
    )


# =========================================================
# NORMALIZATION
# =========================================================

def normalize_text(text):
    if text is None:
        return ""
    text = unicodedata.normalize("NFKC", str(text))
    text = text.lower()
    text = text.replace("’", "'").replace("`", "'").replace("–", "-").replace("—", "-")
    text = text.replace("\n", " ").replace("\r", " ")
    text = re.sub(r"(^|\s)n(?=\s)", " ", text)
    text = re.sub(r"\s*:\s*-\s*", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def tokenize(text):
    value = normalize_text(text)
    return value.split() if value else []


def normalize_strict_text(text):
    if text is None:
        return ""
    text = unicodedata.normalize("NFKC", str(text))
    text = text.replace("\r", " ").replace("\n", " ")
    text = text.replace("’", "'").replace("`", "'")
    text = text.replace("–", "-").replace("—", "-")
    text = re.sub(r"\s*:\s*-\s*", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


# =========================================================
# STATIC / MARKET PREFIX HANDLING
# =========================================================

MARKET_PREFIXES = [
    "US", "CA", "MX", "CR/EC/GT/PA/SV", "CR/EC/GT/MX/PA/SV",
    "USA", "CANADA", "MEXICO"
]

MARKET_PATTERN = re.compile(
    r"(?:^|\s)(?:n\s*)?(US|CA|MX|CR\s*/\s*EC\s*/\s*GT(?:\s*/\s*MX)?\s*/\s*PA\s*/\s*SV|USA|CANADA|MEXICO)\s*:\s*",
    re.IGNORECASE,
)


def canonical_prefix(value):
    return re.sub(r"\s+", "", str(value or "").upper())


def clean_pdf_artifact_prefix(text):
    text = str(text or "")
    text = re.sub(r"^\s*n\s+", "", text, flags=re.IGNORECASE)
    return text.strip()


def remove_leading_market_prefix(text):
    text = clean_pdf_artifact_prefix(text)
    return re.sub(
        r"^(?:US|CA|MX|CR/EC/GT/PA/SV|CR/EC/GT/MX/PA/SV|USA|CANADA|MEXICO)\s*[:\-]\s*",
        "",
        text,
        flags=re.IGNORECASE,
    ).strip()


# =========================================================
# FIELD TYPE / REGION
# =========================================================

def get_field_type(field_name):
    field = normalize_text(field_name)
    compact = field.replace(" ", "")

    if (
        "coo" in compact or "countryoforigin" in compact or
        "countryorigin" in compact or "madein" in compact or
        compact.startswith("min_") or compact.startswith("min") or
        "origin" in compact
    ):
        return "COO"

    if (
        "fiber" in compact or "fibre" in compact or "fib_" in compact or
        compact.startswith("fib") or "fabric" in compact or
        "content" in compact or "composition" in compact or "fabrication" in compact
    ):
        return "CONTENT"

    if (
        "care" in compact or "wash" in compact or "washing" in compact or
        "laundry" in compact or "instruction" in compact or compact.startswith("wc")
    ):
        return "CARE"

    if (
        "size" in compact or "sizeline" in compact or "alpha" in compact or
        "waist" in compact or "inseam" in compact or "fit" in compact or
        compact.startswith("os_") or compact.startswith("os")
    ):
        return "SIZE"

    if (
        compact in {"rn", "rnca", "rnnumber", "rnca_number"} or
        "registrationnumber" in compact or "companyrn" in compact or
        compact.startswith("rn")
    ):
        return "RN"

    if "brand" in compact:
        return "BRAND"
    if "color" in compact or "colour" in compact:
        return "COLOR"
    if "gender" in compact:
        return "GENDER"
    if "attribute" in compact or "technology" in compact or "feature" in compact or "description" in compact:
        return "ATTRIBUTE"
    return "GENERAL"


def get_field_region(field_name):
    original = str(field_name or "").lower()
    compact = normalize_text(field_name).replace(" ", "")

    if "mexico" in compact and ("_sp" in original or compact.endswith("sp") or "spanish" in compact):
        return "SP_MX"
    if "mexico" in compact and ("_sp" in original or compact.endswith("sp") or "spanish" in compact):
        return "SP_MX"
    if "mexico" in compact:
        return "MX"
    if "canada" in compact:
        return "CA"
    if "_en" in original or compact.endswith("en") or "english" in compact:
        return "EN"
    if "_fr" in original or compact.endswith("fr") or "french" in compact:
        return "FR"
    if "_sp" in original or compact.endswith("sp") or "spanish" in compact or "espanol" in compact:
        return "SP"
    return ""


FIELD_ANCHORS = {
    "COO": ["made in", "hecho en", "fabrique en"],
    "CONTENT": ["shell", "liner", "body", "fabric", "fiber", "fibre", "content", "composition", "exterior", "extérieur", "forro", "doublure", "cuerpo"],
    "CARE": ["machine wash", "wash", "lavar", "laver", "dry clean", "bleach", "blanchiment", "detergent", "detergente"],
    "RN": ["rn", "ca", "registration number"],
    "SIZE": ["size", "oz", "ounce", "waist", "inseam"],
    "COLOR": ["color", "colour"],
    "BRAND": ["brand"],
    "GENDER": ["girls", "boys", "women", "men", "unisex"],
    "ATTRIBUTE": ["attribute", "technology", "feature"],
    "GENERAL": [],
}


# =========================================================
# EXCEL / PDF
# =========================================================

def load_excel(file):
    file.seek(0)
    df = pd.read_excel(file, header=0)
    df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_pdf(file):
    file.seek(0)
    document = fitz.open(stream=file.read(), filetype="pdf")
    pages = []
    for page_number, page in enumerate(document):
        pages.append({"page": page_number + 1, "text": page.get_text("text")})
    document.close()
    return pages


# =========================================================
# PFL LINE PREPARATION
# =========================================================

def clean_pdf_line(line):
    if not line:
        return ""
    line = str(line).strip()
    line = re.sub(r"^\s*n\s+(?=[A-Za-z])", "", line)
    return line.strip()


def panel_number_from_line(line):
    if not line:
        return None
    match = re.fullmatch(
        r"(?:panel\s*(?:no\.?|number|#)?\s*[-:]?\s*)?(\d{1,3})",
        str(line).strip(), flags=re.IGNORECASE
    )
    if not match:
        return None
    number = int(match.group(1))
    return number if 1 <= number <= 99 else None


def reorder_pfl_lines(page_text):
    lines = [clean_pdf_line(x) for x in str(page_text or "").splitlines()]
    lines = [x for x in lines if x and len(x) <= 1500]
    if not lines:
        return []

    markers = []
    for i, line in enumerate(lines):
        n = panel_number_from_line(line)
        if n is not None:
            markers.append({"index": i, "number": n})

    distinct = list(dict.fromkeys(m["number"] for m in markers))
    if len(distinct) < 2 or 1 not in distinct or len(distinct) != len(markers):
        return lines

    # Use detected panel order when reliable.
    segments = []
    above = markers[0]["index"] <= 1
    if above:
        for pos, marker in enumerate(markers):
            start = marker["index"] + 1
            end = markers[pos + 1]["index"] if pos + 1 < len(markers) else len(lines)
            if start < end:
                segments.append((marker["number"], lines[start:end]))
    else:
        start = 0
        for marker in markers:
            end = marker["index"]
            if start < end:
                segments.append((marker["number"], lines[start:end]))
            start = marker["index"] + 1
        if start < len(lines):
            segments.append((999999, lines[start:]))

    if not segments:
        return lines
    segments.sort(key=lambda x: x[0])
    out = []
    for _, segment in segments:
        out.extend(segment)
    return out


def create_standard_blocks(page_text):
    lines = [clean_pdf_line(x) for x in str(page_text or "").splitlines()]
    lines = [x for x in lines if x and len(x) <= 1500]
    blocks = []
    for i, line in enumerate(lines):
        blocks.append({"text": line, "start": i, "end": i + 1})
    for size in range(2, min(8, len(lines)) + 1):
        for i in range(len(lines) - size + 1):
            blocks.append({"text": " ".join(lines[i:i + size]), "start": i, "end": i + size})
    return blocks


def split_market_segments(lines):
    """Split the PDF stream into true market-labelled segments.

    Handles prefixes split by PDF extraction, e.g.:
        CR/EC/GT/
        MX/PA/SV :
    """
    work_lines = list(lines)
    merged = []
    i = 0
    while i < len(work_lines):
        current = work_lines[i]
        if i + 1 < len(work_lines):
            a = re.sub(r"\s+", "", current).upper()
            b = re.sub(r"\s+", "", work_lines[i + 1]).upper()
            if a.endswith("CR/EC/GT/") and b.startswith("MX/PA/SV"):
                current = current.rstrip() + " " + work_lines[i + 1].lstrip()
                i += 1
        merged.append(current)
        i += 1
    lines = merged

    segments = []
    current_prefix = ""
    current_parts = []
    current_start = None
    segment_no = 0

    def flush(end_line):
        nonlocal current_prefix, current_parts, current_start, segment_no
        if current_parts:
            text = " ".join(current_parts).strip()
            text = re.split(r"\b(?:ACTUAL\s+OTHER\s+SIZES|1/1)\b", text, flags=re.IGNORECASE)[0].strip()
            if text:
                segment_no += 1
                segments.append({
                    "prefix": current_prefix,
                    "text": text,
                    "start": current_start if current_start is not None else 0,
                    "end": end_line,
                    "regional": bool(current_prefix),
                    "source_id": f"SEG:{segment_no}",
                })
        current_prefix = ""
        current_parts = []
        current_start = None

    for line_index, raw in enumerate(lines):
        work = clean_pdf_line(raw)
        if not work:
            continue

        matches = list(MARKET_PATTERN.finditer(work))
        if not matches:
            if current_parts:
                current_parts.append(work)
            continue

        cursor = 0
        for match in matches:
            before = work[cursor:match.start()].strip()
            if before:
                if current_parts:
                    current_parts.append(before)
                else:
                    segments.append({
                        "prefix": "",
                        "text": before,
                        "start": line_index,
                        "end": line_index + 1,
                        "regional": False,
                        "source_id": f"TXT:{line_index}:{len(segments)}",
                    })

            flush(line_index + 1)
            current_prefix = match.group(1).upper().replace(" ", "")
            current_start = line_index
            current_parts = []
            cursor = match.end()

        tail = work[cursor:].strip()
        if tail:
            current_parts.append(tail)

    flush(len(lines))
    return [x for x in segments if x["text"]]


def field_scoped_text(text, field_type):
    text = str(text or "").strip()
    if not text:
        return text

    if field_type == "CONTENT":
        starts = [
            "shell", "exterior", "extérieur", "cuerpo", "body", "liner", "forro", "doublure"
        ]
        positions = [text.lower().find(x) for x in starts if text.lower().find(x) >= 0]
        if positions:
            text = text[min(positions):]
        stop_terms = [
            "machine wash", "laver", "lavar", "e1s6x", "rn ", "made in", "hecho en"
        ]
        lower = text.lower()
        stops = [lower.find(x) for x in stop_terms if lower.find(x) > 0]
        if stops:
            text = text[:min(stops)]

    elif field_type == "CARE":
        starts = ["machine wash", "laver", "lavar", "wash"]
        positions = [text.lower().find(x) for x in starts if text.lower().find(x) >= 0]
        if positions:
            text = text[min(positions):]

    text = re.split(r"\b(?:ACTUAL\s+OTHER\s+SIZES|1/1)\b", text, flags=re.IGNORECASE)[0]
    return text.strip()


def create_pfl_blocks(page_text):
    lines = reorder_pfl_lines(page_text)
    if not lines:
        return []

    blocks = []
    field_types = ["CONTENT", "CARE"]

    # Regional segments are the strongest PFL candidates. Keep the segment
    # itself as the source so two different markets on the same PDF line can
    # still be used independently.
    for seg in split_market_segments(lines):
        for field_type in field_types:
            scoped = field_scoped_text(seg["text"], field_type)
            if not scoped:
                continue
            if not any(normalize_text(a) in normalize_text(scoped) for a in FIELD_ANCHORS[field_type]):
                continue
            blocks.append({
                "prefix": seg["prefix"],
                "text": scoped,
                "start": seg["start"],
                "end": seg["end"],
                "regional": bool(seg.get("prefix")),
                "source_id": f"{seg['source_id']}:{field_type}",
            })

    # Individual lines and short continuation windows are used for size fields
    # and for unlabelled continuation fragments in PFL.
    max_window = min(10, len(lines))
    for size in range(1, max_window + 1):
        for start_i in range(len(lines) - size + 1):
            text = " ".join(lines[start_i:start_i + size]).strip()
            if text:
                blocks.append({
                    "prefix": "",
                    "text": text,
                    "start": start_i,
                    "end": start_i + size,
                    "regional": False,
                    "source_id": f"WIN:{start_i}:{start_i + size}",
                })

    return blocks


# =========================================================
# FIELD / BLOCK COMPATIBILITY
# =========================================================

def contains_rn(text):
    n = normalize_text(text)
    return bool(
        re.search(r"\brn\s*\d{3,}\b", n) or
        re.search(r"\bca\s*\d{3,}\b", n) or
        re.search(r"\bregistration\s+number\s*\d+", n)
    )


def contains_size_signal(text):
    n = normalize_text(text)
    return bool(
        re.search(r"\b(?:xxxs|xxs|xs|s|m|l|xl|xxl|xxxl|tp|ttg|eeg|tg|eg|ch|p|g|ech)\b", n) or
        re.search(r"\bsize\b", n) or
        re.search(r"\b\d+(?:\.\d+)?\s*oz\b", n) or
        re.search(r"\b\d{2,3}/\d{2,3}\b", n)
    )


def contains_coo(text):
    n = normalize_text(text)
    return any(x in n for x in ["made in", "hecho en", "fabrique en"])


def candidate_market_ok(field_name, block):
    field_type = get_field_type(field_name)
    region = get_field_region(field_name)
    prefix = canonical_prefix(block.get("prefix", ""))

    if field_type not in {"CONTENT", "CARE"} or not region:
        return True
    if region == "EN":
        return prefix in {"US", ""}
    if region in {"FR", "CA"}:
        return prefix in {"CA", ""}
    if region == "MX":
        return prefix in {"MX", ""}
    if region == "SP_MX":
        return prefix in {"CR/EC/GT/MX/PA/SV", "CR/EC/GT/PA/SV", "MX", ""}
    if region == "SP":
        return prefix in {"CR/EC/GT/MX/PA/SV", "CR/EC/GT/PA/SV", "MX", "SP", ""}
    return True


def preferred_market_pool(field_name, pdf_blocks):
    field_type = get_field_type(field_name)
    region = get_field_region(field_name)
    if field_type not in {"CONTENT", "CARE"} or not region:
        return pdf_blocks

    exact = []
    for block in pdf_blocks:
        prefix = canonical_prefix(block.get("prefix", ""))
        if region == "EN" and prefix == "US":
            exact.append(block)
        elif region in {"FR", "CA"} and prefix == "CA":
            exact.append(block)
        elif region == "MX" and prefix == "MX":
            exact.append(block)
        elif region == "SP_MX" and prefix in {"CR/EC/GT/MX/PA/SV", "CR/EC/GT/PA/SV"}:
            exact.append(block)
        elif region == "SP" and prefix in {"CR/EC/GT/MX/PA/SV", "CR/EC/GT/PA/SV", "MX"}:
            exact.append(block)
    return exact if exact else pdf_blocks


def field_anchor_hit(field_name, text):
    field_type = get_field_type(field_name)
    n = normalize_text(text)
    return any(normalize_text(a) in n for a in FIELD_ANCHORS.get(field_type, []))


def block_allowed_for_field(field_name, block):
    field_type = get_field_type(field_name)
    text = block["text"]
    n = normalize_text(text)

    if not n:
        return False

    if not candidate_market_ok(field_name, block):
        return False

    if field_type == "SIZE":
        return not contains_rn(text) and contains_size_signal(text)

    if field_type == "RN":
        return contains_rn(text)

    if field_type == "COO":
        return contains_coo(text)

    if field_type == "CONTENT":
        if contains_rn(text) or contains_coo(text):
            return False
        return any(normalize_text(a) in n for a in FIELD_ANCHORS["CONTENT"])

    if field_type == "CARE":
        if contains_rn(text) or contains_coo(text):
            return False
        return any(normalize_text(a) in n for a in FIELD_ANCHORS["CARE"])

    return True


# =========================================================
# MATCH SCORING
# =========================================================

def token_coverage(expected, actual):
    et = set(tokenize(expected))
    at = set(tokenize(actual))
    if not et:
        return 0.0
    return len(et & at) / len(et)


def ordered_token_coverage(expected, actual):
    et = tokenize(expected)
    at = tokenize(actual)
    if not et or not at:
        return 0.0
    pos = 0
    hits = 0
    for token in et:
        found = False
        for j in range(pos, len(at)):
            if token == at[j]:
                hits += 1
                pos = j + 1
                found = True
                break
        if not found:
            continue
    return hits / len(et)


def score_block(expected, block, field_name):
    actual = remove_leading_market_prefix(block["text"])
    en = normalize_text(expected)
    an = normalize_text(actual)
    coverage = token_coverage(expected, actual)
    ordered = ordered_token_coverage(expected, actual)
    ratio = fuzz.ratio(en, an) if en and an else 0
    partial = fuzz.partial_ratio(en, an) if en and an else 0

    field_type = get_field_type(field_name)
    anchor_bonus = 0
    for anchor in FIELD_ANCHORS.get(field_type, []):
        if normalize_text(anchor) in an:
            anchor_bonus = 12
            break

    region_bonus = 0
    region = get_field_region(field_name)
    prefix = block.get("prefix", "").upper()
    if region == "EN" and prefix == "US":
        region_bonus = 18
    elif region == "FR" and prefix == "CA":
        region_bonus = 18
    elif region == "MX" and prefix in {"MX", "CR/EC/GT/MX/PA/SV", "CR/EC/GT/PA/SV"}:
        region_bonus = 18
    elif region == "SP" and prefix in {"CR/EC/GT/MX/PA/SV", "CR/EC/GT/PA/SV", "MX"}:
        region_bonus = 14

    # Coverage is intentionally dominant. This prevents a short fragment
    # such as "SECAR..." from beating the full WC_SP field.
    score = (
        coverage * 60
        + ordered * 18
        + ratio * 0.10
        + partial * 0.06
        + anchor_bonus
        + region_bonus
    )

    # Prefer candidates that are not dramatically larger than the field.
    expected_len = max(1, len(tokenize(expected)))
    actual_len = max(1, len(tokenize(actual)))
    size_ratio = min(expected_len, actual_len) / max(expected_len, actual_len)
    score += size_ratio * 8

    return score, coverage, ordered


# =========================================================
# CANDIDATE SELECTION
# =========================================================

def choose_pfl_candidate(expected, pdf_blocks, field_name, used_sources):
    field_type = get_field_type(field_name)
    region = get_field_region(field_name)

    def available(block):
        if not block_allowed_for_field(field_name, block):
            return False
        source_id = block.get("source_id", "")
        return source_id not in used_sources

    preferred = preferred_market_pool(field_name, pdf_blocks)
    base = [b for b in preferred if available(b)]
    if not base:
        base = [b for b in pdf_blocks if available(b)]

    # For explicit regional content/care fields, prefer the actual regional
    # segment. Generic windows are only used for continuation.
    if field_type in {"CONTENT", "CARE"} and region:
        regional = [b for b in base if b.get("regional") and candidate_market_ok(field_name, b)]
        if regional:
            base = regional

    scored = []
    for block in base:
        score, coverage, ordered = score_block(expected, block, field_name)
        scored.append((score, coverage, ordered, block))
    scored.sort(key=lambda x: (x[1], x[2], x[0]), reverse=True)

    best = None
    if scored:
        score, coverage, ordered, block = scored[0]
        best = {"blocks": [block], "score": score, "coverage": coverage, "ordered": ordered}

    # Continuation: combine the best regional/seed block with one strong
    # unlabelled block elsewhere on the PFL page. This handles WC_SP where the
    # LAVAR portion and the SECAR portion are separated by the panel layout.
    if field_type in {"CONTENT", "CARE"} and scored:
        best_actual = remove_leading_market_prefix(scored[0][3]["text"])
        expected_len = max(1, len(tokenize(expected)))
        actual_len = max(1, len(tokenize(best_actual)))
        best_length_ratio = min(expected_len, actual_len) / max(expected_len, actual_len)
        needs_continuation = (
            scored[0][1] < 0.70
            or scored[0][2] < 0.70
            or best_length_ratio < 0.80
        )
    else:
        needs_continuation = False

    if field_type in {"CONTENT", "CARE"} and needs_continuation:
        seeds = [x[3] for x in scored[:12]]
        continuation_pool = []
        for b in pdf_blocks:
            if not available(b):
                continue
            if b.get("regional"):
                # A second labelled market is not a continuation of this field.
                continue
            if not field_anchor_hit(field_name, b["text"]):
                continue
            continuation_pool.append(b)
        # Keep search targeted; rank continuations by their independent coverage.
        continuation_pool.sort(
            key=lambda b: token_coverage(expected, remove_leading_market_prefix(b["text"])),
            reverse=True,
        )
        continuation_pool = continuation_pool[:20]

        for first in seeds:
            first_payload = remove_leading_market_prefix(first["text"])
            first_cov = token_coverage(expected, first_payload)
            for second in continuation_pool:
                if second.get("source_id") == first.get("source_id"):
                    continue
                first_ids = {first.get("source_id", "")}
                second_id = second.get("source_id", "")
                if second_id in first_ids:
                    continue

                combined_text = (first_payload + " " + remove_leading_market_prefix(second["text"])).strip()
                combined = {
                    "prefix": first.get("prefix", ""),
                    "text": combined_text,
                    "start": min(first["start"], second["start"]),
                    "end": max(first["end"], second["end"]),
                    "regional": bool(first.get("regional")),
                }
                cscore, coverage, ordered = score_block(expected, combined, field_name)
                second_cov = token_coverage(expected, second["text"])
                if coverage < max(first_cov, second_cov) + 0.15:
                    continue
                cscore += coverage * 20
                if best is None or cscore > best["score"]:
                    best = {
                        "blocks": [first, second],
                        "score": cscore,
                        "coverage": coverage,
                        "ordered": ordered,
                    }

    return best


# =========================================================
# STRICT PFL DIFFERENCE
# =========================================================

def strict_compare(expected, actual):
    expected_strict = normalize_strict_text(expected)
    actual_strict = normalize_strict_text(actual)

    if expected_strict == actual_strict:
        return True, "—"

    # Token-level output keeps punctuation attached, so differences such as
    # "POLY-ESTER" vs "POLYESTER" remain visible.
    exp = expected_strict.split()
    act = actual_strict.split()
    matcher = SequenceMatcher(None, exp, act)
    differences = []

    for tag, a1, a2, b1, b2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        e = " ".join(exp[a1:a2])
        a = " ".join(act[b1:b2])
        if tag == "replace":
            differences.append(f"{e} → {a}")
        elif tag == "delete":
            differences.append(f"Missing from PDF: {e}")
        elif tag == "insert":
            differences.append(f"Extra in PDF: {a}")

    return False, "; ".join(differences[:12]) or "Text differs."


# =========================================================
# FIELD COMPARISON
# =========================================================

def check_field(expected, pdf_blocks, field_name, product_type, used_sources):
    if expected is None or str(expected).strip() == "":
        return {"status": "SKIP", "pdf": "—", "difference": "No variable data in Order Form."}

    expected = str(expected).strip()

    if product_type == "PFL":
        chosen = choose_pfl_candidate(expected, pdf_blocks, field_name, used_sources)
        if chosen is None:
            return {"status": "NOT FOUND", "pdf": "Not found in relevant PDF area", "difference": "Selected variable value was not detected."}

        blocks = chosen["blocks"]
        actual = " ".join(remove_leading_market_prefix(b["text"]) for b in blocks).strip()

        # If the Order Form explicitly contains the market prefix, retain it.
        expected_strict = normalize_strict_text(expected)
        expected_has_prefix = bool(re.match(
            r"^(?:US|CA|MX|CR/EC/GT/PA/SV|CR/EC/GT/MX/PA/SV)\s*[:\-]",
            expected_strict,
            flags=re.IGNORECASE,
        ))
        if expected_has_prefix:
            actual = " ".join(b["text"] for b in blocks).strip()

        passed, diff = strict_compare(expected, actual)

        # Semantic normalization is used only to recognize :- style layout
        # artifacts. The final strict comparison still catches real changes.
        if not passed:
            en = normalize_text(expected)
            an = normalize_text(actual)
            if en and en == an:
                passed = True
                diff = "—"

        for b in blocks:
            if b.get("source_id"):
                used_sources.add(b["source_id"])

        return {
            "status": "PASS" if passed else "FAIL",
            "pdf": actual,
            "difference": "—" if passed else diff,
        }

    allowed = [b for b in pdf_blocks if block_allowed_for_field(field_name, b)]
    if not allowed:
        allowed = pdf_blocks

    best = None
    for block in allowed:
        actual = remove_leading_market_prefix(block["text"])
        if normalize_text(expected) in normalize_text(actual):
            candidate = {"score": 100, "block": block}
        else:
            sc, cov, _ = score_block(expected, block, field_name)
            candidate = {"score": sc, "coverage": cov, "block": block}
        if best is None or candidate["score"] > best["score"]:
            best = candidate

    if best is None:
        return {"status": "NOT FOUND", "pdf": "Not found in relevant PDF area", "difference": "Selected variable value was not detected."}

    actual = best["block"]["text"]
    actual_clean = remove_leading_market_prefix(actual)
    if normalize_text(expected) == normalize_text(actual_clean) or normalize_text(expected) in normalize_text(actual_clean):
        return {"status": "PASS", "pdf": actual, "difference": "—"}
    return {"status": "FAIL", "pdf": actual, "difference": get_difference(expected, actual_clean)}


# =========================================================
# STANDARD DIFFERENCE
# =========================================================

def get_difference(expected, actual):
    exp = tokenize(expected)
    act = tokenize(remove_leading_market_prefix(actual))
    matcher = SequenceMatcher(None, exp, act)
    differences = []
    for tag, a1, a2, b1, b2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        e = " ".join(exp[a1:a2])
        a = " ".join(act[b1:b2])
        if tag == "replace":
            differences.append(f"{e} → {a}")
        elif tag == "delete":
            differences.append(f"Missing: {e}")
        elif tag == "insert":
            differences.append(f"Extra: {a}")
    return "; ".join(differences[:12]) or "Content differs."


# =========================================================
# REPORT
# =========================================================

def build_report(df, pdf_pages, selected_fields, product_type):
    results = []
    field_no = 1

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
                    "DIFFERENCE": "No corresponding Excel row.",
                })
                field_no += 1
            continue

        row = df.iloc[excel_index]
        pdf_blocks = create_pfl_blocks(page["text"]) if product_type == "PFL" else create_standard_blocks(page["text"])
        used_sources = set()

        for field in selected_fields:
            value = "" if pd.isna(row[field]) else str(row[field]).strip()

            if not value:
                results.append({
                    "FIELD NO": field_no,
                    "PDF PAGE": page["page"],
                    "EXCEL ROW": excel_index + 2,
                    "FIELD": field,
                    "ORDER FORM DATA": "",
                    "PDF OUTPUT": "—",
                    "STATUS": "SKIP",
                    "DIFFERENCE": "Blank Order Form value — PDF content ignored.",
                })
                field_no += 1
                continue

            result = check_field(value, pdf_blocks, field, product_type, used_sources)
            results.append({
                "FIELD NO": field_no,
                "PDF PAGE": page["page"],
                "EXCEL ROW": excel_index + 2,
                "FIELD": field,
                "ORDER FORM DATA": value,
                "PDF OUTPUT": result["pdf"],
                "STATUS": result["status"],
                "DIFFERENCE": result["difference"],
            })
            field_no += 1

    return pd.DataFrame(results)


def style_status(value):
    if value == "PASS":
        return "background-color:#238636;color:white;font-weight:bold;"
    if value == "FAIL":
        return "background-color:#da3633;color:white;font-weight:bold;"
    if value == "NOT FOUND":
        return "background-color:#9e6a03;color:white;font-weight:bold;"
    if value == "SKIP":
        return "background-color:#555555;color:white;font-weight:bold;"
    return ""


# =========================================================
# MAIN
# =========================================================

def main():
    render_title()

    top_left, top_right = st.columns([7, 1])
    with top_right:
        if st.button("↻ NEW START", key="of_new_start", width="stretch"):
            st.session_state["of_product_type"] = "Other"
            st.session_state["of_selected_fields"] = []
            st.session_state["of_result"] = None
            st.rerun()

    st.markdown('<div class="tool-section-title">🏷️ Product Type</div>', unsafe_allow_html=True)
    product_types = ["Other", "HTL", "PFL"]
    current_type = st.session_state.get("of_product_type", "Other")
    if current_type not in product_types:
        current_type = "Other"
    product_type = st.selectbox(
        "Product Type",
        product_types,
        index=product_types.index(current_type),
        key="of_product_type",
    )

    if product_type == "PFL":
        st.info("PFL mode enabled — panel sequence, continuation and strict text comparison logic will be used.")
    else:
        st.caption("Standard comparison mode.")

    st.markdown('<div class="tool-section-title">📂 Upload Files</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    with col1:
        excel_file = st.file_uploader("📊 Order Form Excel", type=["xlsx", "xls"], key="of_excel_upload")
    with col2:
        pdf_file = st.file_uploader("📄 Output Artwork PDF", type=["pdf"], key="of_pdf_upload")

    df = None
    if excel_file:
        try:
            df = load_excel(excel_file)
        except Exception as error:
            st.error(f"Unable to read the Excel Order Form: {error}")
            return

        if df.empty:
            st.error("The uploaded Excel does not contain usable data.")
            return

        st.markdown('<div class="tool-section-title">📌 Order Form Fields</div>', unsafe_allow_html=True)
        st.caption("Select ONLY the Excel fields that should be validated against the PDF. Nothing is selected automatically.")

        excel_columns = [str(column) for column in df.columns]
        selected_fields = st.multiselect(
            "Select fields to validate",
            options=excel_columns,
            default=st.session_state.get("of_selected_fields", []),
            key="of_selected_fields",
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
                    "Preview": " | ".join(values[:3]),
                })
            with st.expander("🔎 Preview Selected Fields"):
                st.dataframe(pd.DataFrame(preview_rows), width="stretch", hide_index=True)
        else:
            st.info("Select at least one Excel field to continue.")
    else:
        selected_fields = []

    if not excel_file:
        return
    if not pdf_file:
        st.info("Upload the Output Artwork PDF to continue.")
        return
    if not selected_fields:
        return

    try:
        pdf_pages = load_pdf(pdf_file)
    except Exception as error:
        st.error(f"Unable to read the PDF: {error}")
        return

    if not pdf_pages:
        st.error("No pages could be read from the PDF.")
        return

    st.markdown('<div class="tool-section-title">📌 Validation Setup</div>', unsafe_allow_html=True)
    info1, info2, info3 = st.columns(3)
    with info1:
        st.metric("Excel Rows", len(df))
    with info2:
        st.metric("PDF Pages", len(pdf_pages))
    with info3:
        st.metric("Selected Fields", len(selected_fields))

    if len(df) != len(pdf_pages):
        st.warning(
            "Excel row count and PDF page count do not match. The existing mapping will still be used: "
            "PDF Page 1 → Excel Row 2, PDF Page 2 → Excel Row 3, and so on."
        )
    else:
        st.success("Excel row count and PDF page count match.")

    st.markdown('<div class="tool-section-title">🚀 Run Validation</div>', unsafe_allow_html=True)
    compare_clicked = st.button("🔍  COMPARE & PROOFREAD", key="of_compare", type="primary", width="stretch")
    if not compare_clicked:
        return

    with st.spinner("Checking selected variable artwork data..."):
        report = build_report(df, pdf_pages, selected_fields, product_type)

    st.session_state["of_result"] = report

    st.divider()
    st.markdown('<div class="tool-section-title">📋 QC Report</div>', unsafe_allow_html=True)

    if report.empty:
        st.warning("No validation results were generated.")
        return

    pass_count = int((report["STATUS"] == "PASS").sum())
    fail_count = int((report["STATUS"] == "FAIL").sum())
    not_found_count = int((report["STATUS"] == "NOT FOUND").sum())
    skip_count = int((report["STATUS"] == "SKIP").sum())

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("PASS", pass_count)
    with c2:
        st.metric("FAIL", fail_count)
    with c3:
        st.metric("NOT FOUND", not_found_count)
    with c4:
        st.metric("IGNORED", skip_count)

    styled_report = report.style.map(style_status, subset=["STATUS"])
    st.dataframe(styled_report, width="stretch", hide_index=True)

    st.divider()
    if fail_count > 0:
        st.error(f"❌ FAIL — {fail_count} variable-data mismatch(es) detected.")
    elif not_found_count > 0:
        st.warning(f"⚠️ REVIEW — {not_found_count} selected variable field(s) could not be located.")
    else:
        st.success("✅ PASS — All selected variable fields matched the PDF artwork.")

    failures = report[report["STATUS"].isin(["FAIL", "NOT FOUND"])]
    if not failures.empty:
        st.markdown("### 🔎 Difference Details")
        for _, result in failures.iterrows():
            field_name = result["FIELD"]
            title = f"❌ {field_name} — Page {result['PDF PAGE']}" if result["STATUS"] == "FAIL" else f"⚠️ {field_name} — Page {result['PDF PAGE']}"
            with st.expander(title):
                left, right = st.columns(2)
                with left:
                    st.markdown("**Order Form Data**")
                    st.code(str(result["ORDER FORM DATA"]))
                with right:
                    st.markdown("**PDF Output**")
                    st.code(str(result["PDF OUTPUT"]))
                st.markdown("**Difference**")
                difference = str(result["DIFFERENCE"])
                if "→" in difference or "Extra" in difference or "Missing" in difference:
                    st.error(difference)
                else:
                    st.warning(difference)

    with st.expander("ℹ️ How validation works"):
        st.write(
            """
            **Variable-data validation**

            Only fields selected from the Order Form are treated as variable artwork data.

            **PFL matching**

            PFL first identifies the most relevant market/field region, then evaluates the full
            continuation window instead of selecting a short fragment solely because it is similar.

            **PFL strict comparison**

            After the correct region is found, the original text is compared for case, punctuation,
            numbers, wording, missing content and extra content. PDF line breaks are joined, but the
            actual text and punctuation are not silently removed.

            **Static market prefixes**

            Prefixes such as US:, CA:, MX: and regional market labels are treated as structural when
            they are not part of the selected Order Form value. If the selected field itself contains
            such a prefix, it remains part of the comparison.

            **One-time use**

            A matched PFL source range is locked so the same source region is not repeatedly assigned
            to multiple selected fields.

            **Page mapping**

            PDF Page 1 → Excel Row 2
            PDF Page 2 → Excel Row 3
            PDF Page 3 → Excel Row 4
            and so on.
            """
        )

    csv_data = report.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label="⬇️ Download QC Report",
        data=csv_data,
        file_name="Order_Form_Output_QC_Report.csv",
        mime="text/csv",
        width="stretch",
    )


if __name__ == "__main__":
    main()
