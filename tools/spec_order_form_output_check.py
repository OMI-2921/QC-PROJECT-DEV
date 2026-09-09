import io
import math
import re
import unicodedata
from collections import defaultdict
from copy import deepcopy
from pathlib import Path

import pandas as pd
import streamlit as st
import fitz

from PIL import Image, ImageDraw, ImageFont

# Reuse the proven Order Form -> Output engine.  Tool 3 adds an ORG Spec layer
# around that engine rather than replacing the existing field validation logic.
try:
    from .order_form_output_check import (
        _apply_tool_css as _apply_base_css,
        check_field,
        extract_output_pages,
        get_available_fields,
        get_field_type,
        get_field_region,
        is_admin_field,
        is_blank_value,
        load_excel,
        normalize_text,
        order_fields_for_matching,
        build_page_state,
        find_identifier_match,
    )
except Exception:
    from order_form_output_check import (
        _apply_tool_css as _apply_base_css,
        check_field,
        extract_output_pages,
        get_available_fields,
        get_field_type,
        get_field_region,
        is_admin_field,
        is_blank_value,
        load_excel,
        normalize_text,
        order_fields_for_matching,
        build_page_state,
        find_identifier_match,
    )


TOOL3_VERSION = "2026-09-10-V2-FULL-INTERACTIVE-DASHBOARD"

# -----------------------------------------------------------------------------
# Visual palette.  The screenshot/UI requested by the user is a restrained
# dark enterprise QC application.  These colors are deliberately stable so
# the same meaning is preserved in the artwork, table, and legend.
# -----------------------------------------------------------------------------
STATIC_YELLOW = (250, 204, 21, 120)
VARIABLE_BLUE = (59, 130, 246, 135)
PASS_GREEN = (34, 197, 94, 135)
FAIL_RED = (239, 68, 68, 145)
REVIEW_ORANGE = (249, 115, 22, 145)
UNACCOUNTED_PURPLE = (168, 85, 247, 135)
BORDER_DARK = "#263247"
CARD_BG = "#111a2a"
CARD_BG_2 = "#0d1625"
PAGE_BG = "#07111f"
MUTED = "#94a3b8"
TEXT = "#f8fafc"
BLUE = "#2196f3"

FIELD_COLORS = {
    "CARE": (56, 189, 248, 125),
    "CONTENT": (168, 85, 247, 125),
    "COO": (20, 184, 166, 125),
    "RN": (245, 158, 11, 125),
    "IDENTIFIER": (244, 63, 94, 125),
    "SIZE": (99, 102, 241, 125),
    "COLOR": (236, 72, 153, 125),
    "GENDER": (14, 165, 233, 125),
    "BRAND": (132, 204, 22, 125),
    "ATTRIBUTE": (234, 179, 8, 125),
    "QUANTITY": (20, 184, 166, 125),
    "BATCH": (249, 115, 22, 125),
    "BARCODE": (6, 182, 212, 125),
    "OSZ": (139, 92, 246, 125),
    "SYMBOL": (217, 70, 239, 125),
    "GENERAL": (100, 116, 139, 125),
}


def _tool3_css():
    # Use the same base styling as Tool 1 first, then layer the Tool 3 layout.
    try:
        _apply_base_css()
    except Exception:
        pass

    st.markdown(
        """
        <style>
        html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"],
        .stApp, [data-testid="stMain"] {
            background: #07111f !important;
            color: #f8fafc !important;
        }
        [data-testid="stHeader"] { background: #07111f !important; }
        .block-container { padding-top: 1.1rem !important; max-width: 1540px !important; }

        .t3-header {
            display: flex; align-items: center; gap: 16px;
            padding: 2px 2px 14px 2px;
        }
        .t3-logo {
            width: 42px; height: 42px; border-radius: 12px;
            background: linear-gradient(135deg, #0ea5e9, #2563eb);
            display: flex; align-items: center; justify-content: center;
            font-size: 22px; box-shadow: 0 0 18px rgba(37,99,235,.25);
        }
        .t3-title { font-size: 27px; font-weight: 800; line-height: 1.08; color: #f8fafc; }
        .t3-subtitle { font-size: 13px; color: #8fa1b8; margin-top: 4px; }
        .t3-rule { height: 34px; width: 1px; background: #334155; margin: 0 2px; }

        .t3-side {
            background: #091424; border: 1px solid #1e2c40; border-radius: 12px;
            padding: 8px; min-height: 590px;
        }
        .t3-nav {
            display: flex; align-items: center; gap: 10px; padding: 10px 12px;
            color: #aebdd0; border-radius: 9px; margin-bottom: 5px; font-size: 13px;
        }
        .t3-nav.active { background: #0d2c53; color: #7cc9ff; }
        .t3-nav .ico { width: 18px; text-align:center; }

        .t3-upload-card, .t3-panel, .t3-summary-card, .t3-legend-card {
            background: #0d1829; border: 1px solid #22334a; border-radius: 10px;
        }
        .t3-upload-card { padding: 14px; min-height: 142px; }
        .t3-card-title { font-weight: 750; color: #f4f7fb; font-size: 14px; }
        .t3-card-sub { color: #8395aa; font-size: 11px; margin-top: 2px; }
        .t3-fileline { font-size: 11px; color: #c7d3e2; margin-top: 8px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        .t3-ok { color: #22c55e; font-size: 16px; margin-left: 5px; }

        [data-testid="stFileUploader"] {
            background: #0b1524 !important; border: 1px solid #263a53 !important;
            border-radius: 8px !important; padding: 4px !important;
        }
        [data-testid="stFileUploaderDropzone"] {
            background: #0b1524 !important; border: 1px dashed #38506f !important;
            border-radius: 7px !important;
        }
        [data-testid="stFileUploaderDropzoneInstructions"] { font-size: 11px !important; }
        [data-testid="stFileUploaderDropzoneInstructions"] span { color: #dbe7f4 !important; }

        [data-baseweb="select"] > div {
            background: #0b1524 !important; border: 1px solid #2a3d56 !important;
            color: #f8fafc !important; border-radius: 8px !important;
        }
        [data-baseweb="select"] span, [data-baseweb="select"] input { color: #f8fafc !important; }

        .t3-section-head {
            font-size: 14px; font-weight: 750; color: #f8fafc; padding: 11px 14px;
            border-bottom: 1px solid #1d2b3f;
        }
        .t3-tabs {
            display:flex; gap: 7px; padding: 10px 12px; border-bottom: 1px solid #1d2b3f;
        }
        .t3-tab {
            color:#9cadbf; background:#0a1320; border:1px solid #26374d;
            border-radius:7px; padding:7px 10px; font-size:11px;
        }
        .t3-tab.active { color:#eaf5ff; border-color:#1888ff; background:#0b2440; }

        .t3-badge {
            display:inline-block; padding:4px 9px; border-radius:999px; font-size:10px;
            font-weight:800; letter-spacing:.02em;
        }
        .t3-badge.pass { background:#155e34; color:#bbf7d0; }
        .t3-badge.fail { background:#7f1d1d; color:#fecaca; }
        .t3-badge.static { background:#6b5200; color:#fef08a; }
        .t3-badge.variable { background:#124f9b; color:#bfdbfe; }
        .t3-badge.review { background:#7c2d12; color:#fed7aa; }
        .t3-badge.missing { background:#581c87; color:#e9d5ff; }

        .t3-small { color:#8da0b5; font-size:10px; }
        .t3-metric { font-size: 22px; font-weight: 800; color:#f8fafc; }
        .t3-metric-label { font-size:10px; color:#8496aa; }

        .t3-summary-row { display:flex; justify-content:space-between; align-items:center; padding:8px 0; border-bottom:1px solid #1a2637; font-size:11px; }
        .t3-summary-row:last-child { border-bottom:none; }
        .dot { width:9px; height:9px; border-radius:50%; display:inline-block; margin-right:7px; }
        .dot.static { background:#facc15; }
        .dot.variable { background:#3b82f6; }
        .dot.issue { background:#f43f5e; }
        .dot.review { background:#f97316; }

        .t3-zoom { font-size:11px; color:#9fb0c4; text-align:center; }
        .t3-note { font-size:10px; color:#8395aa; line-height:1.5; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _norm_casefold(text):
    value = unicodedata.normalize("NFKC", str(text or "")).strip()
    value = re.sub(r"\s+", " ", value)
    return value.casefold()


def _compact(text):
    return re.sub(r"[^a-z0-9]+", "", _norm_casefold(text))


def _letters_only(text):
    return re.sub(r"[^A-Za-z]", "", unicodedata.normalize("NFKC", str(text or "")))


def _case_signature(text):
    sig = []
    for ch in unicodedata.normalize("NFKC", str(text or "")):
        if ch.isalpha():
            sig.append("U" if ch.isupper() else "L")
    return sig


def _case_pattern_matches(reference, actual):
    """Compare letter-case shape while ignoring numbers, punctuation and spaces."""
    ref_letters = _letters_only(reference)
    actual_letters = _letters_only(actual)
    if not ref_letters or not actual_letters:
        return True

    # For changed variable content (e.g. Polyester replacing Cotton), compare
    # aggregate word-level case conventions rather than letter-by-letter text.
    ref_words = re.findall(r"[A-Za-z]+", str(reference or ""))
    actual_words = re.findall(r"[A-Za-z]+", str(actual or ""))
    if not ref_words or not actual_words:
        return True

    def word_case(w):
        if w.isupper():
            return "UPPER"
        if w.islower():
            return "LOWER"
        if len(w) > 1 and w[0].isupper() and w[1:].islower():
            return "TITLE"
        return "MIXED"

    ref_modes = [word_case(w) for w in ref_words]
    actual_modes = [word_case(w) for w in actual_words]

    # Prefer the ORG convention for the corresponding first/last words.
    target = ref_modes[0]
    checks = [actual_modes[0]]
    if len(ref_modes) > 1 and len(actual_modes) > 1:
        target = ref_modes[-1]
        checks.append(actual_modes[-1])

    return all(mode == target for mode in checks)


def _case_issue(reference, actual, product_type):
    if product_type == "PFL":
        # PFL is governed by the Order Form. This function is not used as a
        # decision override there; it is retained for UI notes.
        return False
    return not _case_pattern_matches(reference, actual)


def _safe_float(v, default=0.0):
    try:
        return float(v)
    except Exception:
        return default


def _word_bbox(words):
    if not words:
        return None
    left = min(int(w.get("left", 0)) for w in words)
    top = min(int(w.get("top", 0)) for w in words)
    right = max(int(w.get("left", 0)) + int(w.get("width", 0)) for w in words)
    bottom = max(int(w.get("top", 0)) + int(w.get("height", 0)) for w in words)
    if right <= left or bottom <= top:
        return None
    return (left, top, right, bottom)


def _make_visual_blocks(page):
    """Turn OCR words into line/block objects with text, bbox, and normalized anchors."""
    words = [w for w in page.get("ocr_words", []) if str(w.get("text", "")).strip()]
    grouped = defaultdict(list)
    for word in words:
        key = (
            word.get("block_num", 0),
            word.get("par_num", 0),
            word.get("line_num", 0),
        )
        grouped[key].append(word)

    blocks = []
    for key, items in grouped.items():
        items = sorted(items, key=lambda x: (int(x.get("left", 0)), int(x.get("top", 0))))
        bbox = _word_bbox(items)
        if not bbox:
            continue
        text = " ".join(str(x.get("text", "")).strip() for x in items).strip()
        if not text:
            continue
        left, top, right, bottom = bbox
        blocks.append({
            "key": key,
            "text": text,
            "norm": _norm_casefold(text),
            "compact": _compact(text),
            "bbox": bbox,
            "cx": (left + right) / 2,
            "cy": (top + bottom) / 2,
            "width": max(1, right - left),
            "height": max(1, bottom - top),
            "words": items,
        })

    blocks.sort(key=lambda b: (b["bbox"][1], b["bbox"][0]))
    return blocks


def _bbox_metrics(a, b, page_w, page_h):
    acx, acy = a["cx"], a["cy"]
    bcx, bcy = b["cx"], b["cy"]
    dx = abs(acx - bcx) / max(1, page_w)
    dy = abs(acy - bcy) / max(1, page_h)

    a_left, a_top, a_right, a_bottom = a["bbox"]
    b_left, b_top, b_right, b_bottom = b["bbox"]
    a_w = max(1, a_right - a_left)
    a_h = max(1, a_bottom - a_top)
    b_w = max(1, b_right - b_left)
    b_h = max(1, b_bottom - b_top)

    width_ratio = b_w / a_w
    height_ratio = b_h / a_h
    scale_score = min(width_ratio, 1 / width_ratio) * min(height_ratio, 1 / height_ratio)

    # Alignment evidence: compare center and left/right anchors.  This makes
    # center-aligned labels less likely to be treated as shifted left/right.
    center_shift = abs(acx - bcx) / max(1, page_w)
    left_shift = abs(a_left - b_left) / max(1, page_w)
    right_shift = abs(a_right - b_right) / max(1, page_w)
    best_anchor = min(center_shift, left_shift, right_shift)

    diagonal = math.hypot(page_w, page_h)
    center_diag = math.hypot(acx - bcx, acy - bcy) / max(1, diagonal)

    return {
        "dx": dx,
        "dy": dy,
        "center_diag": center_diag,
        "scale_score": scale_score,
        "best_anchor": best_anchor,
        "width_ratio": width_ratio,
        "height_ratio": height_ratio,
    }


def _static_match_score(org_block, out_block, page_w, page_h):
    """Score identical text at the same physical position."""
    if org_block["norm"] != out_block["norm"] and org_block["compact"] != out_block["compact"]:
        return None

    metrics = _bbox_metrics(org_block, out_block, page_w, page_h)

    # The user asked for tolerance for small movement/shrink, not broad fuzzy
    # placement.  Position must stay close and the physical footprint should
    # remain reasonably comparable.
    if metrics["center_diag"] > 0.045:
        return None
    if metrics["best_anchor"] > 0.035:
        return None
    if metrics["width_ratio"] < 0.60 or metrics["width_ratio"] > 1.70:
        return None
    if metrics["height_ratio"] < 0.55 or metrics["height_ratio"] > 1.80:
        return None

    score = (
        (1.0 - min(1.0, metrics["center_diag"] / 0.045)) * 0.55
        + metrics["scale_score"] * 0.30
        + (1.0 - min(1.0, metrics["best_anchor"] / 0.035)) * 0.15
    )
    return max(0.0, min(1.0, score))


def compare_static_blocks(org_page, out_page):
    """Return ORG->Output exact static matches based on text + position."""
    org_blocks = _make_visual_blocks(org_page)
    out_blocks = _make_visual_blocks(out_page)
    page_w = max(int(out_page.get("image_width", 1)), 1)
    page_h = max(int(out_page.get("image_height", 1)), 1)

    candidates = []
    for oi, org in enumerate(org_blocks):
        if len(org["compact"]) < 2:
            continue
        # Do not let empty or purely punctuation blocks become static.
        if not re.search(r"[A-Za-z0-9]", org["text"]):
            continue
        for ui, out in enumerate(out_blocks):
            score = _static_match_score(org, out, page_w, page_h)
            if score is not None:
                candidates.append((score, oi, ui))

    candidates.sort(reverse=True)
    used_org = set()
    used_out = set()
    matches = []

    for score, oi, ui in candidates:
        if oi in used_org or ui in used_out:
            continue
        used_org.add(oi)
        used_out.add(ui)
        matches.append({
            "org": org_blocks[oi],
            "output": out_blocks[ui],
            "score": score,
            "classification": "STATIC",
            "status": "PASS",
        })

    return matches


def _all_visual_blocks(page):
    return _make_visual_blocks(page)


def _nearest_org_block(output_block, org_page, excluded=None):
    excluded = excluded or set()
    blocks = _all_visual_blocks(org_page)
    page_w = max(int(org_page.get("image_width", 1)), 1)
    page_h = max(int(org_page.get("image_height", 1)), 1)
    best = None
    best_score = 10**9
    for idx, block in enumerate(blocks):
        if idx in excluded:
            continue
        metrics = _bbox_metrics(output_block, block, page_w, page_h)
        score = metrics["center_diag"] + metrics["best_anchor"] * 0.5
        if score < best_score:
            best_score = score
            best = block
    if best is None or best_score > 0.10:
        return None
    return best


def _find_blocks_for_text(page, target):
    target_norm = _norm_casefold(target)
    target_compact = _compact(target)
    if not target_norm or target_norm in {"not found", "—", "-"}:
        return []
    blocks = _all_visual_blocks(page)
    exact = []
    containing = []

    for block in blocks:
        if block["norm"] == target_norm or block["compact"] == target_compact:
            exact.append(block)
            continue
        if target_compact and (target_compact in block["compact"] or block["compact"] in target_compact):
            containing.append(block)

    return exact or containing


def _field_color(field_name):
    return FIELD_COLORS.get(get_field_type(field_name), FIELD_COLORS["GENERAL"])


def _clean_display(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def _field_type_label(field_name):
    field_type = get_field_type(field_name)
    region = get_field_region(field_name)
    return f"{field_type}{' • ' + region if region else ''}"


def _format_field_result(field, expected, result, out_page, org_page, product_type):
    actual = result.get("pdf", "")
    actual_text = "" if str(actual).strip() in {"Not found", "—", "-"} else str(actual).strip()

    out_blocks = _find_blocks_for_text(out_page, actual_text)
    output_block = out_blocks[0] if out_blocks else None
    org_block = _nearest_org_block(output_block, org_page) if output_block else None

    status = result.get("status", "NOT FOUND")
    review_case = False
    case_note = ""

    if output_block and org_block and product_type in {"HTL", "Other"} and status == "PASS":
        review_case = _case_issue(org_block["text"], actual_text, product_type)
        if review_case:
            case_note = "Data matches Order Form; confirm Output letter case against ORG Spec."

    classification = "VARIABLE" if status in {"PASS", "FAIL"} else "MISSING / UNACCOUNTED"
    final_status = status
    if review_case and status == "PASS":
        final_status = "REVIEW"

    return {
        "field": field,
        "field_type": _field_type_label(field),
        "expected": expected,
        "actual": actual_text or "Not found",
        "status": final_status,
        "base_status": status,
        "classification": classification,
        "difference": case_note or result.get("difference", "—"),
        "match_type": result.get("match_type", ""),
        "output_block": output_block,
        "org_reference_block": org_block,
    }


def _new_output_state(page, product_type):
    return build_page_state(page, product_type)


def _run_order_form_validation(df, output_pages, selected_fields, product_type, page_row_mapping=None):
    """Run the existing field engine per mapped Output page without altering it."""
    results = []
    field_no = 1
    fields_for_match = order_fields_for_matching(selected_fields)
    osz_group_size = sum(1 for f in selected_fields if get_field_type(f) == "OSZ")

    for page_index, output_page in enumerate(output_pages):
        page_no = int(output_page.get("page", page_index + 1))
        if page_row_mapping and page_no in page_row_mapping:
            row_idx = int(page_row_mapping[page_no])
        else:
            row_idx = min(page_index, len(df) - 1)

        if row_idx < 0 or row_idx >= len(df):
            for field in selected_fields:
                results.append({
                    "FIELD NO": field_no,
                    "PDF PAGE": page_no,
                    "EXCEL ROW": "N/A",
                    "FIELD": field,
                    "ORDER FORM DATA": "No corresponding Order Form row",
                    "result": {"status": "NOT FOUND", "pdf": "No corresponding Order Form row", "difference": "No corresponding Excel row."},
                })
                field_no += 1
            continue

        row = df.iloc[row_idx]
        state = _new_output_state(output_page, product_type)
        page_results = {}

        for field in fields_for_match:
            expected = "" if is_blank_value(row[field]) else str(row[field]).strip()
            page_results[field] = check_field(expected, field, state, osz_group_size=osz_group_size)

        for field in selected_fields:
            expected = "" if is_blank_value(row[field]) else str(row[field]).strip()
            page_results.setdefault(field, {"status": "SKIP", "pdf": "—", "difference": "Blank Order Form value — field ignored."})
            enriched = _format_field_result(
                field, expected, page_results[field], output_page,
                output_pages[page_index].get("_org_page", output_page), product_type
            )
            enriched.update({
                "FIELD NO": field_no,
                "PDF PAGE": page_no,
                "EXCEL ROW": row_idx + 2,
            })
            results.append(enriched)
            field_no += 1

    return results


def _build_static_rows(org_pages, output_pages):
    rows = []
    limit = min(len(org_pages), len(output_pages))
    for idx in range(limit):
        org = org_pages[idx]
        out = output_pages[idx]
        matches = compare_static_blocks(org, out)
        for match in matches:
            rows.append({
                "PDF PAGE": out.get("page", idx + 1),
                "Element / Field": "Static Element",
                "ORG Spec": match["org"]["text"],
                "Output": match["output"]["text"],
                "Order Form": "—",
                "Status": "STATIC",
                "Notes": "Text and position match within allowed tolerance.",
                "output_block": match["output"],
                "org_block": match["org"],
            })
    return rows


def _prepare_pages(org_pages, output_pages):
    for idx, out in enumerate(output_pages):
        if idx < len(org_pages):
            out["_org_page"] = org_pages[idx]
        else:
            out["_org_page"] = None


def _selectable_fields(df):
    fields = get_available_fields(df)
    return [f for f in fields if not is_admin_field(f)]


def _dynamic_auto_detect(df, output_pages, org_pages, product_type):
    """Detect fields that produce meaningful PASS/FAIL evidence, plus visible case REVIEW."""
    candidates = _selectable_fields(df)
    if not candidates:
        return []

    detected = []
    # Probe with the same engine, but keep each candidate isolated so one field
    # cannot consume text that belongs to another field.
    for field in candidates:
        found = False
        for page_idx, out_page in enumerate(output_pages):
            row_idx = min(page_idx, len(df) - 1)
            if row_idx < 0 or row_idx >= len(df):
                continue
            expected = "" if is_blank_value(df.iloc[row_idx][field]) else str(df.iloc[row_idx][field]).strip()
            if not expected:
                continue
            state = _new_output_state(out_page, product_type)
            result = check_field(expected, field, state, osz_group_size=1)
            if result.get("status") in {"PASS", "FAIL"}:
                found = True
                break
        if found:
            detected.append(field)
    return detected


def _page_row_mapping_ui(df, pages, reset_id):
    """Compact page->row selector.  Defaults remain intuitive and editable."""
    if not pages:
        return {}

    with st.expander("Page → Order Form Data Mapping", expanded=(len(pages) > 1), icon="🧭"):
        st.caption("Default: page 1 → Excel row 2, page 2 → Excel row 3, etc. You can override any page.")
        mapping = {}
        labels = [f"Excel Row {i + 2}" for i in range(len(df))]
        options = list(range(len(df)))
        cols = st.columns(min(4, max(1, len(pages))))
        for idx, page in enumerate(pages):
            page_no = int(page.get("page", idx + 1))
            default_idx = min(idx, len(df) - 1)
            with cols[idx % len(cols)]:
                choice = st.selectbox(
                    f"PDF Page {page_no}",
                    options=options,
                    index=default_idx if default_idx in options else 0,
                    format_func=lambda x: labels[x],
                    key=f"t3_row_{reset_id}_{page_no}",
                )
                mapping[page_no] = choice
        return mapping


def _resize_for_ui(image, max_width=660):
    if image is None:
        return None
    if image.width <= max_width:
        return image
    scale = max_width / image.width
    return image.resize((int(image.width * scale), int(image.height * scale)), Image.Resampling.LANCZOS)


def _draw_label(draw, xy, label, fill):
    x, y = xy
    try:
        font = ImageFont.load_default()
    except Exception:
        font = None
    pad_x = 5
    pad_y = 3
    bbox = draw.textbbox((0, 0), label, font=font)
    w = bbox[2] - bbox[0] + pad_x * 2
    h = bbox[3] - bbox[1] + pad_y * 2
    draw.rounded_rectangle([x, max(0, y - h), x + w, y], radius=4, fill=fill[:3] + (220,))
    draw.text((x + pad_x, max(0, y - h + pad_y - 1)), label, fill=(255, 255, 255, 255), font=font)


def _paint_box(image, bbox, fill, label=None, outline=(255, 255, 255, 235), width=3):
    left, top, right, bottom = [int(v) for v in bbox]
    overlay = Image.new("RGBA", image.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    draw.rounded_rectangle([left, top, right, bottom], radius=4, fill=fill, outline=outline, width=width)
    if label:
        _draw_label(draw, (left, max(top, 26)), label, outline)
    return Image.alpha_composite(image.convert("RGBA"), overlay)


def _build_annotated_output(output_page, static_matches, variable_rows, page_idx):
    raw = output_page.get("image_bytes")
    if not raw:
        return None
    image = Image.open(io.BytesIO(raw)).convert("RGBA")

    # Static elements first.
    for match in static_matches:
        block = match["output"]
        image = _paint_box(
            image,
            block["bbox"],
            STATIC_YELLOW,
            "STATIC",
            outline=(250, 204, 21, 245),
            width=3,
        )

    # Variable/data layer second.  FAIL and REVIEW deliberately override the
    # normal field color so QC attention is visually unambiguous.
    for row in variable_rows:
        block = row.get("output_block")
        if not block:
            continue
        status = row.get("status")
        if status == "FAIL":
            fill, outline, label = FAIL_RED, (239, 68, 68, 250), f"{row['field']} • FAIL"
        elif status == "REVIEW":
            fill, outline, label = REVIEW_ORANGE, (249, 115, 22, 250), f"{row['field']} • REVIEW"
        elif status == "PASS":
            fill, outline, label = _field_color(row["field"]), _field_color(row["field"]), f"{row['field']} • PASS"
        else:
            continue
        image = _paint_box(image, block["bbox"], fill, label, outline=outline, width=3)

    return image.convert("RGB")


def _render_static_side_by_side(org_page, output_page, static_matches, variable_rows=None):
    org_raw = org_page.get("image_bytes") if org_page else None
    out_img = _build_annotated_output(output_page, static_matches, variable_rows or [], output_page.get("page", 1))
    org_img = Image.open(io.BytesIO(org_raw)).convert("RGB") if org_raw else None
    if org_img is not None:
        org_img = _resize_for_ui(org_img, 620)
    if out_img is not None:
        out_img = _resize_for_ui(out_img, 620)
    return org_img, out_img


def _html_badge(value):
    text = str(value or "")
    cls = "pass"
    if text == "STATIC":
        cls = "static"
    elif text == "VARIABLE":
        cls = "variable"
    elif text == "REVIEW":
        cls = "review"
    elif "MISSING" in text:
        cls = "missing"
    elif text == "FAIL" or text == "ISSUE":
        cls = "fail"
    return f'<span class="t3-badge {cls}">{text}</span>'


def _summary_metrics(static_rows, variable_rows):
    static_count = len(static_rows)
    variable_count = sum(1 for r in variable_rows if r.get("status") in {"PASS", "FAIL", "REVIEW"})
    issue_count = sum(1 for r in variable_rows if r.get("status") == "FAIL")
    review_count = sum(1 for r in variable_rows if r.get("status") == "REVIEW")
    missing_count = sum(1 for r in variable_rows if r.get("status") == "NOT FOUND")
    overall = "FAIL" if issue_count else ("REVIEW" if review_count else "PASS")
    return overall, static_count, variable_count, issue_count, review_count, missing_count


def _build_dataframe(static_rows, variable_rows):
    rows = []
    for r in static_rows:
        rows.append({
            "PDF PAGE": r["PDF PAGE"],
            "Element / Field": "Static Element",
            "Type": "STATIC",
            "ORG Spec": r["ORG Spec"],
            "Output": r["Output"],
            "Order Form": "—",
            "Status": "STATIC",
            "Notes": r["Notes"],
        })
    for r in variable_rows:
        rows.append({
            "PDF PAGE": r["PDF PAGE"],
            "Element / Field": r["field"],
            "Type": "VARIABLE",
            "ORG Spec": r["org_reference_block"]["text"] if r.get("org_reference_block") else "Not matched",
            "Output": r["actual"],
            "Order Form": r["expected"],
            "Status": r["status"],
            "Notes": r["difference"],
        })
    return pd.DataFrame(rows)


def _to_excel_bytes(summary, comparison_df, annotated_images, mapping, product_type):
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"

    header_fill = PatternFill("solid", fgColor="12233A")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="24364D")

    ws["A1"] = "ORG SPEC + ORDER FORM + OUTPUT QC"
    ws["A1"].font = Font(size=18, bold=True, color="FFFFFF")
    ws["A2"] = "Product Type"
    ws["B2"] = product_type
    ws["A3"] = "Overall Result"
    ws["B3"] = summary[0]
    ws["A4"] = "Static Elements"
    ws["B4"] = summary[1]
    ws["A5"] = "Variable Elements"
    ws["B5"] = summary[2]
    ws["A6"] = "Issues"
    ws["B6"] = summary[3]
    ws["A7"] = "Manual Review"
    ws["B7"] = summary[4]
    ws["A8"] = "Not Found / Unaccounted"
    ws["B8"] = summary[5]

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font

    detail = wb.create_sheet("Field Comparison")
    if not comparison_df.empty:
        for col_idx, col in enumerate(comparison_df.columns, start=1):
            c = detail.cell(1, col_idx, col)
            c.fill = header_fill
            c.font = header_font
            c.border = Border(bottom=thin)
        for row_idx, row in enumerate(comparison_df.itertuples(index=False), start=2):
            for col_idx, value in enumerate(row, start=1):
                detail.cell(row_idx, col_idx, value)

        status_col = list(comparison_df.columns).index("Status") + 1
        for r in range(2, detail.max_row + 1):
            status = str(detail.cell(r, status_col).value or "")
            if status == "PASS":
                detail.cell(r, status_col).fill = PatternFill("solid", fgColor="198754")
            elif status == "FAIL":
                detail.cell(r, status_col).fill = PatternFill("solid", fgColor="DC3545")
            elif status == "STATIC":
                detail.cell(r, status_col).fill = PatternFill("solid", fgColor="C79B00")
            elif status == "REVIEW":
                detail.cell(r, status_col).fill = PatternFill("solid", fgColor="E67E22")

    visual = wb.create_sheet("Artwork Visual Validation")
    visual["A1"] = "Page Visual Validation"
    visual["A1"].font = Font(size=16, bold=True)
    visual["A3"] = "Visual status"
    visual["B3"] = "Yellow = Static | Field colors = Order Form data | Red = Fail | Orange = Review"

    row_cursor = 5
    for page_no, image in sorted(annotated_images.items()):
        visual.cell(row_cursor, 1, f"PDF Page {page_no}").font = Font(bold=True)
        row_cursor += 1
        if image is not None:
            tmp = io.BytesIO()
            image.save(tmp, format="PNG")
            tmp.seek(0)
            img = XLImage(tmp)
            img.width = min(560, image.width)
            img.height = int(image.height * (img.width / image.width))
            visual.add_image(img, f"A{row_cursor}")
            row_cursor += max(35, int(img.height / 14))
            row_cursor += 2

    for sheet in wb.worksheets:
        for col in range(1, min(sheet.max_column, 12) + 1):
            sheet.column_dimensions[get_column_letter(col)].width = 18 if col != 1 else 24
        for row in sheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = 20

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def _table_html(df):
    if df.empty:
        return "<div class='t3-note'>No comparison rows available.</div>"
    headers = list(df.columns)
    parts = ["<div style='overflow-x:auto'><table style='width:100%; border-collapse:collapse; font-size:11px'>"]
    parts.append("<thead><tr>")
    for h in headers:
        parts.append(f"<th style='text-align:left;padding:8px;border-bottom:1px solid #26374d;color:#93a6bc'>{h}</th>")
    parts.append("</tr></thead><tbody>")
    for _, row in df.iterrows():
        parts.append("<tr>")
        for h in headers:
            value = row[h]
            if h == "Status":
                value = _html_badge(value)
            else:
                value = str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            parts.append(f"<td style='padding:8px;border-bottom:1px solid #1d2b3f;color:#d7e0eb;vertical-align:top'>{value}</td>")
        parts.append("</tr>")
    parts.append("</tbody></table></div>")
    return "".join(parts)



# -----------------------------------------------------------------------------
# INTERACTIVE DASHBOARD HELPERS
# -----------------------------------------------------------------------------

def _status_chip(text, kind=None):
    value = str(text or "")
    key = (kind or value).upper()
    cls = {
        "PASS": "pass", "FAIL": "fail", "STATIC": "static", "VARIABLE": "variable",
        "REVIEW": "review", "MISSING / UNACCOUNTED": "missing", "NOT FOUND": "missing",
        "ISSUE": "fail",
    }.get(key, "review")
    return f'<span class="t3-badge {cls}">{value}</span>'


def _escape_html(value):
    return (str(value or "")
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))


def _result_icon(overall):
    if overall == "PASS":
        return "✓"
    if overall == "REVIEW":
        return "!"
    return "×"


def _build_page_structural_stats(org_page, out_page, static_matches, variable_rows):
    out_blocks = _all_visual_blocks(out_page) if out_page else []
    org_blocks = _all_visual_blocks(org_page) if org_page else []
    static_n = len(static_matches)
    variable_n = sum(1 for r in variable_rows if r.get("classification") == "VARIABLE")
    fail_n = sum(1 for r in variable_rows if r.get("status") == "FAIL")
    review_n = sum(1 for r in variable_rows if r.get("status") == "REVIEW")
    return {
        "org_text_blocks": len(org_blocks),
        "output_text_blocks": len(out_blocks),
        "static": static_n,
        "variable": variable_n,
        "fail": fail_n,
        "review": review_n,
    }


def _render_reference_card(df, row_idx, product_type, selected_fields, page_no):
    if row_idx is None or row_idx < 0 or row_idx >= len(df):
        return "<div class='t3-note'>No Order Form row mapped to this page.</div>"
    row = df.iloc[row_idx]
    rows = []
    for field in selected_fields[:14]:
        value = _clean_display(row[field])
        if value:
            rows.append(
                "<div class='t3-reference-row'>"
                f"<div class='t3-reference-label'>{_escape_html(field)}</div>"
                f"<div class='t3-reference-value'>{_escape_html(value)}</div>"
                "</div>"
            )
    if not rows:
        rows.append("<div class='t3-note'>No populated selected fields for this row.</div>")
    return (
        f"<div class='t3-reference-head'><span>Page {page_no}</span>"
        f"<span class='t3-reference-product'>{_escape_html(product_type)}</span></div>"
        + "".join(rows)
    )


def _render_right_summary(overall, static_count, variable_count, issue_count, review_count, missing_count):
    result_class = "pass" if overall == "PASS" else ("review" if overall == "REVIEW" else "fail")
    return f"""
    <div class="t3-result-card">
        <div class="t3-card-title">Result Summary</div>
        <div class="t3-result-main">
            <div class="t3-result-circle {result_class}">{_result_icon(overall)}</div>
            <div><div class="t3-result-word {result_class}">{overall}</div>
            <div class="t3-note">All selected checks completed.</div></div>
        </div>
        <div class="t3-summary-row"><span><i class="dot static"></i>Static Elements</span><b>{static_count}</b></div>
        <div class="t3-summary-row"><span><i class="dot variable"></i>Variable Elements</span><b>{variable_count}</b></div>
        <div class="t3-summary-row"><span><i class="dot issue"></i>Issues Found</span><b>{issue_count}</b></div>
        <div class="t3-summary-row"><span><i class="dot review"></i>Requires Manual Review</span><b>{review_count}</b></div>
        <div class="t3-summary-row"><span><i class="dot missing"></i>Unaccounted / Not Found</span><b>{missing_count}</b></div>
    </div>
    """


def _render_highlight_legend():
    return """
    <div class="t3-legend-card">
        <div class="t3-card-title">Key Highlights</div>
        <div class="t3-legend-grid">
            <div><span class="t3-swatch static"></span>Static (ORG ↔ Output)</div>
            <div><span class="t3-swatch variable"></span>Order Form Data (Variable)</div>
            <div><span class="t3-swatch pass"></span>Match / Pass</div>
            <div><span class="t3-swatch issue"></span>Mismatch / Fail</div>
            <div><span class="t3-swatch review"></span>Needs Review</div>
        </div>
    </div>
    """


def _make_top_css():
    return """
    <style>
    .t3-main-grid { margin-top: 2px; }
    .t3-upload-row { display:grid; grid-template-columns: 1fr 1fr 1fr 1.05fr; gap: 14px; }
    .t3-upload-wrap { padding: 2px; }
    .t3-upload-card { padding: 15px 15px 11px 15px !important; min-height: 126px; }
    .t3-file-title { display:flex; align-items:center; gap:10px; margin-bottom:9px; }
    .t3-file-icon { width:36px; height:36px; border-radius:8px; background:#0b223c; border:1px solid #2a4564; display:flex; align-items:center; justify-content:center; font-size:20px; }
    .t3-file-state { margin-left:auto; color:#22c55e; font-size:18px; }
    .t3-product-card { padding:15px; min-height:126px; }
    .t3-product-label { color:#cbd6e4; font-size:11px; margin-bottom:7px; font-weight:700; }
    .t3-run-area { margin-top:10px; }
    .t3-panel-shell { background:#0c1727; border:1px solid #21334a; border-radius:10px; overflow:hidden; }
    .t3-panel-toolbar { display:flex; justify-content:space-between; align-items:center; padding:10px 12px; border-bottom:1px solid #1e2c3f; }
    .t3-toolbar-title { font-size:12px; font-weight:800; color:#e8eef6; }
    .t3-toolbar-meta { font-size:10px; color:#8194aa; }
    .t3-ref-card { background:#0b1626; border:1px solid #1f3046; border-radius:9px; padding:10px; }
    .t3-reference-head { display:flex; justify-content:space-between; gap:8px; padding-bottom:8px; border-bottom:1px solid #1d2b3e; margin-bottom:4px; color:#eff5fb; font-size:11px; font-weight:800; }
    .t3-reference-product { color:#7cc9ff; }
    .t3-reference-row { padding:7px 0; border-bottom:1px solid #182638; }
    .t3-reference-row:last-child { border-bottom:none; }
    .t3-reference-label { color:#7f94ab; font-size:9px; margin-bottom:3px; }
    .t3-reference-value { color:#d8e4f1; font-size:11px; line-height:1.3; }
    .t3-visual-toolbar { display:flex; gap:7px; align-items:center; flex-wrap:wrap; padding:10px 12px; border-bottom:1px solid #1d2b3f; }
    .t3-page-pill { background:#091424; border:1px solid #263b54; color:#c5d4e4; padding:6px 9px; border-radius:7px; font-size:10px; }
    .t3-image-label { font-size:11px; font-weight:800; color:#dfe8f2; margin:6px 0; }
    .t3-page-stat { display:grid; grid-template-columns:repeat(4,1fr); gap:8px; margin-top:10px; }
    .t3-mini-stat { background:#0b1625; border:1px solid #1e2d41; border-radius:8px; padding:8px; }
    .t3-mini-stat b { font-size:15px; color:#f6f9fc; display:block; }
    .t3-mini-stat span { color:#7f93aa; font-size:9px; }
    .t3-result-card, .t3-legend-card { background:#0d1829; border:1px solid #22334a; border-radius:10px; padding:13px; margin-bottom:10px; }
    .t3-result-main { display:flex; align-items:center; gap:12px; margin:13px 0; }
    .t3-result-circle { width:48px; height:48px; border-radius:50%; display:flex; align-items:center; justify-content:center; font-weight:900; font-size:28px; }
    .t3-result-circle.pass { background:#2bd675; color:#062716; }
    .t3-result-circle.review { background:#f97316; color:#3b1605; }
    .t3-result-circle.fail { background:#ef4444; color:#320606; }
    .t3-result-word { font-size:22px; font-weight:900; }
    .t3-result-word.pass { color:#35e27f; }
    .t3-result-word.review { color:#fb923c; }
    .t3-result-word.fail { color:#fb6060; }
    .t3-legend-grid { display:grid; grid-template-columns:1fr; gap:9px; margin-top:10px; color:#c5d3e1; font-size:10px; }
    .t3-swatch { width:12px; height:12px; border-radius:3px; display:inline-block; margin-right:7px; vertical-align:-2px; }
    .t3-swatch.static { background:#facc15; } .t3-swatch.variable { background:#3b82f6; } .t3-swatch.pass { background:#22c55e; } .t3-swatch.issue { background:#ef4444; } .t3-swatch.review { background:#f97316; }
    .t3-swatch.missing, .dot.missing { background:#a855f7; }
    .t3-action-row { display:flex; gap:8px; align-items:center; }
    .t3-sidebar-note { color:#71859c; font-size:9px; padding:12px 9px 4px 9px; line-height:1.5; }
    .t3-control-label { color:#92a6bb; font-size:9px; text-transform:uppercase; letter-spacing:.05em; font-weight:700; }
    .t3-empty { border:1px dashed #29405b; border-radius:10px; padding:28px; color:#758aa0; text-align:center; font-size:11px; }
    .t3-dataframe-title { font-size:12px; font-weight:800; color:#e9f0f8; margin-bottom:8px; }
    .t3-issue-item { background:#0b1625; border:1px solid #25384f; border-radius:8px; padding:9px; margin-bottom:7px; }
    .t3-issue-head { display:flex; justify-content:space-between; gap:10px; align-items:center; }
    .t3-issue-field { color:#eaf2fb; font-size:11px; font-weight:800; }
    .t3-issue-detail { color:#8ea2b8; font-size:10px; line-height:1.4; margin-top:5px; }
    @media (max-width: 1100px) { .t3-upload-row { grid-template-columns:1fr 1fr; } }
    </style>
    """


def _render_detail_table(comparison_df, page_filter=None, status_filter=None, type_filter=None):
    df = comparison_df.copy()
    if page_filter and page_filter != "All":
        try:
            df = df[df["PDF PAGE"].astype(str) == str(page_filter)]
        except Exception:
            pass
    if status_filter and status_filter != "All":
        df = df[df["Status"].astype(str) == status_filter]
    if type_filter and type_filter != "All":
        df = df[df["Type"].astype(str) == type_filter]
    return df


def _interactive_comparison_table(df, key_prefix):
    if df.empty:
        st.markdown("<div class='t3-empty'>No rows match the current filters.</div>", unsafe_allow_html=True)
        return
    show = df[[c for c in ["PDF PAGE", "Element / Field", "Type", "ORG Spec", "Output", "Order Form", "Status", "Notes"] if c in df.columns]].copy()
    st.dataframe(
        show,
        width="stretch",
        hide_index=True,
        height=min(520, 155 + max(1, len(show)) * 44),
        column_config={
            "PDF PAGE": st.column_config.NumberColumn("Page", width="small"),
            "Element / Field": st.column_config.TextColumn("Element / Field", width="medium"),
            "Type": st.column_config.TextColumn("Type", width="small"),
            "ORG Spec": st.column_config.TextColumn("ORG Spec", width="large"),
            "Output": st.column_config.TextColumn("Output", width="large"),
            "Order Form": st.column_config.TextColumn("Order Form", width="large"),
            "Status": st.column_config.TextColumn("Status", width="small"),
            "Notes": st.column_config.TextColumn("Notes", width="large"),
        },
        key=f"{key_prefix}_df"
    )


def _render_issues(variable_results, static_rows, reset_id):
    issues = []
    for row in variable_results:
        if row.get("status") in {"FAIL", "REVIEW", "NOT FOUND"}:
            issues.append(row)
    if not issues:
        st.markdown("<div class='t3-empty'>No issues or review items were generated.</div>", unsafe_allow_html=True)
        return
    selected_page = st.selectbox("Issue page", ["All"] + sorted({str(r.get("PDF PAGE")) for r in issues}), key=f"t3_issue_page_{reset_id}")
    selected = [r for r in issues if selected_page == "All" or str(r.get("PDF PAGE")) == selected_page]
    for row in selected:
        kind = row.get("status", "REVIEW")
        st.markdown(
            f"<div class='t3-issue-item'><div class='t3-issue-head'><div class='t3-issue-field'>{_escape_html(row.get('field'))} • Page {row.get('PDF PAGE')}</div>{_status_chip(kind)}</div>"
            f"<div class='t3-issue-detail'><b>Order Form:</b> {_escape_html(row.get('expected','—'))}<br><b>Output:</b> {_escape_html(row.get('actual','—'))}<br><b>Finding:</b> {_escape_html(row.get('difference','—'))}</div></div>",
            unsafe_allow_html=True,
        )


def _render_summary_panel(static_rows, variable_results, product_type, output_pages, org_pages):
    overall, static_count, variable_count, issue_count, review_count, missing_count = _summary_metrics(static_rows, variable_results)
    st.markdown(_render_right_summary(overall, static_count, variable_count, issue_count, review_count, missing_count), unsafe_allow_html=True)
    st.markdown("<div class='t3-legend-card'><div class='t3-card-title'>Validation Breakdown</div>", unsafe_allow_html=True)
    total = max(1, static_count + variable_count + missing_count)
    metrics = [("Static", static_count), ("Variable", variable_count), ("Issues", issue_count), ("Review", review_count), ("Unaccounted", missing_count)]
    for label, value in metrics:
        pct = min(100, int(round((value / total) * 100))) if total else 0
        st.markdown(f"<div style='margin:9px 0'><div style='display:flex;justify-content:space-between;font-size:10px;color:#91a4b9'><span>{label}</span><span>{value}</span></div><div style='height:6px;background:#172538;border-radius:6px;overflow:hidden;margin-top:4px'><div style='height:6px;width:{pct}%;background:#2674b8;border-radius:6px'></div></div></div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown(_render_highlight_legend(), unsafe_allow_html=True)
    st.markdown("<div class='t3-legend-card'><div class='t3-card-title'>QC Interpretation</div><div class='t3-note' style='margin-top:8px'><b>Static</b> = ORG and Output contain the same text in approximately the same position.<br><br><b>Variable</b> = Output data validated from the Order Form using Tool 1 logic.<br><br><b>Review</b> = Data may be correct, but presentation/case or another structural point needs confirmation.<br><br><b>Unaccounted</b> = no reliable ORG/Order Form explanation was found.</div></div>", unsafe_allow_html=True)



def _default_mapping(df, output_pages):
    if df.empty:
        return {}
    return {int(page.get("page", idx + 1)): min(idx, len(df) - 1) for idx, page in enumerate(output_pages)}


def main():
    _tool3_css()
    st.markdown(_make_top_css(), unsafe_allow_html=True)

    if "t3_reset_id" not in st.session_state:
        st.session_state["t3_reset_id"] = 0
    if "t3_result" not in st.session_state:
        st.session_state["t3_result"] = None
    if "t3_selected_fields" not in st.session_state:
        st.session_state["t3_selected_fields"] = None
    if "t3_active_section" not in st.session_state:
        st.session_state["t3_active_section"] = "Comparison View"
    if "t3_history" not in st.session_state:
        st.session_state["t3_history"] = []
    if "t3_zoom" not in st.session_state:
        st.session_state["t3_zoom"] = 100
    if st.session_state.get("t3_version") != TOOL3_VERSION:
        st.session_state["t3_result"] = None
        st.session_state["t3_selected_fields"] = None
        st.session_state["t3_active_section"] = "Comparison View"
        st.session_state["t3_version"] = TOOL3_VERSION

    reset_id = st.session_state["t3_reset_id"]

    # =============================== HEADER ===============================
    st.markdown(
        """
        <div class="t3-header">
            <div class="t3-logo">✓</div>
            <div class="t3-rule"></div>
            <div>
                <div class="t3-title">ORG Spec + Order Form + Output Check</div>
                <div class="t3-subtitle">Verify artwork accuracy, data, positioning and presentation</div>
            </div>
            <div style="margin-left:auto;color:#a8bfd7;font-size:20px">⚙</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Outer application shell: left navigation, main workspace, live results rail.
    nav_col, workspace_col, rail_col = st.columns([0.88, 4.45, 1.35], gap="small")

    # ============================= LEFT NAV ==============================
    with nav_col:
        st.markdown('<div class="t3-side">', unsafe_allow_html=True)
        active = st.session_state.get("t3_active_section", "Comparison View")
        if st.button("⌂  Home", width="stretch", type="primary" if active == "Comparison View" else "secondary", key=f"t3_nav_home_{reset_id}"):
            st.session_state["t3_active_section"] = "Comparison View"
            st.rerun()
        if st.button("＋  New Check", width="stretch", key=f"t3_nav_new_{reset_id}"):
            st.session_state["t3_reset_id"] += 1
            st.session_state["t3_result"] = None
            st.session_state["t3_selected_fields"] = None
            st.session_state["t3_active_section"] = "Comparison View"
            st.rerun()
        if st.button("↶  History", width="stretch", key=f"t3_nav_history_{reset_id}"):
            st.session_state["t3_active_section"] = "History"
            st.rerun()
        if st.button("⚙  Settings", width="stretch", key=f"t3_nav_settings_{reset_id}"):
            st.session_state["t3_active_section"] = "Settings"
            st.rerun()
        st.markdown("<div class='t3-sidebar-note'>Tool 3 reads ORG as the structural reference, Order Form as the variable-data source, and Output as the final artwork.</div></div>", unsafe_allow_html=True)

    # =========================== MAIN WORKSPACE ==========================
    with workspace_col:
        # ---------------------------- Upload row -------------------------
        upload_cols = st.columns(4, gap="small")
        with upload_cols[0]:
            st.markdown('<div class="t3-upload-card"><div class="t3-file-title"><div class="t3-file-icon">▤</div><div><div class="t3-card-title">1. Upload Order Form</div><div class="t3-card-sub">Excel / CSV</div></div></div>', unsafe_allow_html=True)
            order_file = st.file_uploader("Order Form", type=["xlsx", "xls", "csv"], label_visibility="collapsed", key=f"t3_order_{reset_id}")
            if order_file:
                st.markdown(f"<div class='t3-fileline'>{_escape_html(order_file.name)} <span class='t3-file-state'>✓</span></div>", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with upload_cols[1]:
            st.markdown('<div class="t3-upload-card"><div class="t3-file-title"><div class="t3-file-icon">▤</div><div><div class="t3-card-title">2. Upload ORG Spec</div><div class="t3-card-sub">PDF / Image</div></div></div>', unsafe_allow_html=True)
            org_file = st.file_uploader("ORG Spec", type=["pdf", "jpg", "jpeg", "png"], label_visibility="collapsed", key=f"t3_org_{reset_id}")
            if org_file:
                st.markdown(f"<div class='t3-fileline'>{_escape_html(org_file.name)} <span class='t3-file-state'>✓</span></div>", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with upload_cols[2]:
            st.markdown('<div class="t3-upload-card"><div class="t3-file-title"><div class="t3-file-icon">▤</div><div><div class="t3-card-title">3. Upload Output</div><div class="t3-card-sub">PDF / Image</div></div></div>', unsafe_allow_html=True)
            output_file = st.file_uploader("Output", type=["pdf", "jpg", "jpeg", "png"], label_visibility="collapsed", key=f"t3_output_{reset_id}")
            if output_file:
                st.markdown(f"<div class='t3-fileline'>{_escape_html(output_file.name)} <span class='t3-file-state'>✓</span></div>", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with upload_cols[3]:
            st.markdown('<div class="t3-product-card"><div class="t3-product-label">Product Type</div>', unsafe_allow_html=True)
            product_type = st.selectbox("Product Type", ["----- SELECT -----", "PFL", "HTL", "Other"], index=0, key=f"t3_product_{reset_id}")
            ready = bool(order_file and org_file and output_file and product_type != "----- SELECT -----")
            run_clicked = st.button("▶  Run Check", width="stretch", type="primary", disabled=not ready, key=f"t3_run_{reset_id}")
            st.markdown('</div>', unsafe_allow_html=True)

        # ---------------------- Run / load comparison --------------------
        if run_clicked:
            try:
                with st.spinner("Reading Order Form, ORG Spec and Output..."):
                    if str(getattr(order_file, "name", "")).casefold().endswith(".csv"):
                        df = pd.read_csv(order_file)
                    else:
                        df = load_excel(order_file)
                    if df.empty:
                        raise ValueError("The uploaded Order Form does not contain any data rows.")
                    org_pages = extract_output_pages(org_file)
                    output_pages = extract_output_pages(output_file)
                    if not org_pages or not output_pages:
                        raise ValueError("Both ORG Spec and Output must contain at least one readable page.")
                    fields = _selectable_fields(df)
                    if not fields:
                        raise ValueError("No usable fields were found in the Order Form.")
                    mapping = _default_mapping(df, output_pages)
                    _prepare_pages(org_pages, output_pages)
                    detected = _dynamic_auto_detect(df, output_pages, org_pages, product_type)
                    static_rows = _build_static_rows(org_pages, output_pages)
                    st.session_state["t3_selected_fields"] = detected
                    st.session_state["t3_active_section"] = "Comparison View"
                    st.session_state["t3_result"] = {
                        "df": df, "org_pages": org_pages, "output_pages": output_pages,
                        "mapping": mapping, "all_fields": fields,
                        "detected_fields": detected, "static_rows": static_rows,
                        "product_type": product_type,
                    }
                    st.session_state["t3_history"].append({
                        "order": getattr(order_file, "name", "Order Form"),
                        "org": getattr(org_file, "name", "ORG Spec"),
                        "output": getattr(output_file, "name", "Output"),
                        "product": product_type,
                    })
                st.success("Initial ORG → Output structural scan completed.")
                st.rerun()
            except Exception as exc:
                st.error(f"QC could not be started: {exc}")

        # --------------------- History / Settings views ------------------
        active = st.session_state.get("t3_active_section", "Comparison View")
        if active == "History":
            st.markdown('<div class="t3-panel-shell"><div class="t3-panel-toolbar"><span class="t3-toolbar-title">Check History</span><span class="t3-toolbar-meta">Current Streamlit session</span></div></div>', unsafe_allow_html=True)
            history = st.session_state.get("t3_history", [])
            if not history:
                st.markdown("<div class='t3-empty' style='margin-top:12px'>No checks have been run in this session.</div>", unsafe_allow_html=True)
            else:
                for i, item in enumerate(reversed(history[-10:]), 1):
                    st.markdown(f"<div class='t3-issue-item'><div class='t3-issue-head'><div class='t3-issue-field'>Check {i} • {_escape_html(item['product'])}</div><span class='t3-badge pass'>COMPLETED</span></div><div class='t3-issue-detail'>Order Form: {_escape_html(item['order'])}<br>ORG: {_escape_html(item['org'])}<br>Output: {_escape_html(item['output'])}</div></div>", unsafe_allow_html=True)
            # Keep the right rail visible as application chrome.
        elif active == "Settings":
            st.markdown('<div class="t3-panel-shell"><div class="t3-panel-toolbar"><span class="t3-toolbar-title">Settings</span><span class="t3-toolbar-meta">Display controls</span></div></div>', unsafe_allow_html=True)
            st.checkbox("Show page structural statistics", value=True, key=f"t3_setting_stats_{reset_id}")
            st.checkbox("Show ORG reference beside variable findings", value=True, key=f"t3_setting_org_{reset_id}")
            st.select_slider("Viewer image width", options=[520, 580, 620, 680], value=620, key=f"t3_width_{reset_id}")
            st.info("These settings control presentation only. The existing Order Form → Output comparison engine remains unchanged.")
        else:
            result = st.session_state.get("t3_result")
            if not result:
                st.markdown("<div class='t3-empty' style='margin-top:12px'>Upload the three inputs, select Product Type, then click <b>Run Check</b>.</div>", unsafe_allow_html=True)
            else:
                df = result["df"]
                org_pages = result["org_pages"]
                output_pages = result["output_pages"]
                product_type = result["product_type"]
                mapping = result["mapping"].copy()

                # ---------------- Mapping + field selection --------------
                map_col, field_col, product_col = st.columns([1.25, 2.65, 1.0], gap="small")
                with map_col:
                    with st.expander("🧭 Page Mapping", expanded=len(output_pages) > 1):
                        st.caption("Default: page 1 → Excel row 2, page 2 → Excel row 3, etc.")
                        options = list(range(len(df)))
                        labels = [f"Excel Row {i+2}" for i in options]
                        for idx, page in enumerate(output_pages):
                            page_no = int(page.get("page", idx+1))
                            default = mapping.get(page_no, min(idx, len(df)-1))
                            mapping[page_no] = st.selectbox(f"PDF Page {page_no}", options, index=max(0,min(int(default),len(df)-1)), format_func=lambda x, labels=labels: labels[x], key=f"t3_map_{reset_id}_{page_no}")
                        result["mapping"] = mapping
                with field_col:
                    available = result["all_fields"]
                    saved = st.session_state.get("t3_selected_fields")
                    defaults = [f for f in (saved if saved is not None else result.get("detected_fields", [])) if f in available]
                    selected_fields = st.multiselect("Field Comparison", available, default=defaults, key=f"t3_fields_{reset_id}", help="Type inside the box to search. Blank Order Form fields are excluded.")
                    st.session_state["t3_selected_fields"] = selected_fields
                with product_col:
                    st.markdown("<div class='t3-control-label'>Product</div>", unsafe_allow_html=True)
                    st.markdown(f"<div style='font-size:16px;font-weight:850;color:#f6f9fc;margin-top:4px'>{_escape_html(product_type)}</div>", unsafe_allow_html=True)
                    if st.button("Auto Detect", width="stretch", key=f"t3_auto_{reset_id}"):
                        detected = _dynamic_auto_detect(df, output_pages, org_pages, product_type)
                        result["detected_fields"] = detected
                        st.session_state["t3_selected_fields"] = detected
                        st.rerun()

                # ------------------ Validation calculation ----------------
                _prepare_pages(org_pages, output_pages)
                variable_results = []
                by_page_blocks = defaultdict(list)
                fields_for_match = order_fields_for_matching(selected_fields) if selected_fields else []
                osz_group_size = sum(1 for f in selected_fields if get_field_type(f) == "OSZ")
                for page_idx, out_page in enumerate(output_pages):
                    page_no = int(out_page.get("page", page_idx + 1))
                    row_idx = mapping.get(page_no, min(page_idx, len(df)-1))
                    if row_idx < 0 or row_idx >= len(df):
                        continue
                    row = df.iloc[row_idx]
                    state = _new_output_state(out_page, product_type)
                    page_temp = {}
                    for field in fields_for_match:
                        expected = "" if is_blank_value(row[field]) else str(row[field]).strip()
                        page_temp[field] = check_field(expected, field, state, osz_group_size=osz_group_size)
                    org_page = org_pages[page_idx] if page_idx < len(org_pages) else None
                    for field in selected_fields:
                        expected = "" if is_blank_value(row[field]) else str(row[field]).strip()
                        base = page_temp.get(field, {"status":"SKIP","pdf":"—","difference":"Blank Order Form value — field ignored."})
                        enriched = _format_field_result(field, expected, base, out_page, org_page, product_type)
                        enriched.update({"PDF PAGE": page_no, "EXCEL ROW": row_idx + 2})
                        variable_results.append(enriched)
                        by_page_blocks[page_no].append(enriched)

                static_rows = _build_static_rows(org_pages, output_pages)
                overall, static_count, variable_count, issue_count, review_count, missing_count = _summary_metrics(static_rows, variable_results)

                # ---------------------- Top tabs --------------------------
                top_tabs = st.tabs(["👁  Comparison View", "☷  Details", "▥  Summary"])

                # =================== COMPARISON VIEW ======================
                with top_tabs[0]:
                    controls = st.columns([1.18,1.32,1.0,1.0])
                    page_options = [int(p.get("page",i+1)) for i,p in enumerate(output_pages)]
                    with controls[0]:
                        selected_page = st.selectbox("Output Page", page_options, key=f"t3_view_page_{reset_id}")
                    with controls[1]:
                        view_mode = st.selectbox("Viewer", ["ORG + Output", "Output Only", "ORG Only"], key=f"t3_view_mode_{reset_id}")
                    with controls[2]:
                        st.markdown("<div class='t3-control-label' style='margin-bottom:6px'>Zoom</div>", unsafe_allow_html=True)
                        zc = st.columns(3)
                        with zc[0]:
                            if st.button("−", key=f"t3_zminus_{reset_id}", width="stretch"):
                                st.session_state["t3_zoom"] = max(60, st.session_state["t3_zoom"] - 10); st.rerun()
                        with zc[1]: st.markdown(f"<div class='t3-page-pill'>{st.session_state['t3_zoom']}%</div>", unsafe_allow_html=True)
                        with zc[2]:
                            if st.button("+", key=f"t3_zplus_{reset_id}", width="stretch"):
                                st.session_state["t3_zoom"] = min(150, st.session_state["t3_zoom"] + 10); st.rerun()
                    with controls[3]:
                        st.markdown("<div class='t3-control-label' style='margin-bottom:6px'>Product</div>", unsafe_allow_html=True)
                        st.markdown(f"<div class='t3-page-pill'>{_escape_html(product_type)}</div>", unsafe_allow_html=True)

                    idx = page_options.index(selected_page)
                    output_page = output_pages[idx]
                    org_page = org_pages[idx] if idx < len(org_pages) else None
                    static_matches = compare_static_blocks(org_page, output_page) if org_page else []
                    page_vars = by_page_blocks.get(selected_page, [])
                    stats = _build_page_structural_stats(org_page, output_page, static_matches, page_vars)
                    org_img, out_img = _render_static_side_by_side(org_page, output_page, static_matches, page_vars)
                    width = int(st.session_state.get(f"t3_width_{reset_id}", 620))
                    # Zoom is visual scaling, not a comparison change.
                    width = max(420, int(width * st.session_state.get("t3_zoom", 100) / 100))
                    org_img = _resize_for_ui(org_img, width) if org_img else None
                    out_img = _resize_for_ui(out_img, width) if out_img else None

                    if view_mode == "ORG + Output":
                        cols = st.columns([1.25,1.25,1.0], gap="small")
                        with cols[0]:
                            st.markdown("<div class='t3-image-label'>ORG Spec</div>", unsafe_allow_html=True)
                            if org_img is not None: st.image(org_img, width="stretch")
                        with cols[1]:
                            st.markdown("<div class='t3-image-label'>Output</div>", unsafe_allow_html=True)
                            if out_img is not None: st.image(out_img, width="stretch")
                        with cols[2]:
                            st.markdown("<div class='t3-image-label'>Order Form (Reference)</div>", unsafe_allow_html=True)
                            row_idx = mapping.get(selected_page, min(idx, len(df)-1))
                            st.markdown(f"<div class='t3-ref-card'>{_render_reference_card(df,row_idx,product_type,selected_fields,selected_page)}</div>", unsafe_allow_html=True)
                    elif view_mode == "Output Only":
                        cols = st.columns([1.55,1.0], gap="small")
                        with cols[0]:
                            st.markdown("<div class='t3-image-label'>Output</div>", unsafe_allow_html=True)
                            if out_img is not None: st.image(out_img, width="stretch")
                        with cols[1]:
                            st.markdown("<div class='t3-image-label'>Order Form (Reference)</div>", unsafe_allow_html=True)
                            row_idx = mapping.get(selected_page, min(idx, len(df)-1))
                            st.markdown(f"<div class='t3-ref-card'>{_render_reference_card(df,row_idx,product_type,selected_fields,selected_page)}</div>", unsafe_allow_html=True)
                    else:
                        cols = st.columns([1.55,1.0], gap="small")
                        with cols[0]:
                            st.markdown("<div class='t3-image-label'>ORG Spec</div>", unsafe_allow_html=True)
                            if org_img is not None: st.image(org_img, width="stretch")
                        with cols[1]:
                            st.markdown("<div class='t3-image-label'>Order Form (Reference)</div>", unsafe_allow_html=True)
                            row_idx = mapping.get(selected_page, min(idx, len(df)-1))
                            st.markdown(f"<div class='t3-ref-card'>{_render_reference_card(df,row_idx,product_type,selected_fields,selected_page)}</div>", unsafe_allow_html=True)

                    if st.session_state.get(f"t3_setting_stats_{reset_id}", True):
                        s_cols = st.columns(4)
                        for col,(n,lbl) in zip(s_cols,[(stats['static'],'Static'),(stats['variable'],'Variable'),(stats['fail'],'Fail'),(stats['review'],'Review')]):
                            with col: st.markdown(f"<div class='t3-mini-stat'><b>{n}</b><span>{lbl}</span></div>", unsafe_allow_html=True)

                    lower_tabs = st.tabs(["Findings", "Field Comparison", "Issues"])
                    with lower_tabs[0]:
                        page_rows = []
                        for m in static_matches:
                            page_rows.append({"Element / Field":"Static Element","Type":"STATIC","ORG Spec":m["org"]["text"],"Output":m["output"]["text"],"Status":"STATIC","Notes":"Text and position match within allowed tolerance."})
                        for r in page_vars:
                            page_rows.append({"Element / Field":r["field"],"Type":r["classification"],"ORG Spec":r["org_reference_block"]["text"] if r.get("org_reference_block") else "Not matched","Output":r["actual"],"Status":r["status"],"Notes":r["difference"]})
                        _interactive_comparison_table(pd.DataFrame(page_rows), f"t3_page_findings_{reset_id}")
                    with lower_tabs[1]:
                        comp_df = _build_dataframe(static_rows, variable_results)
                        filter_cols = st.columns(3)
                        p_opts = ["All"] + sorted({str(x) for x in comp_df["PDF PAGE"].unique()}) if not comp_df.empty else ["All"]
                        s_opts = ["All"] + sorted({str(x) for x in comp_df["Status"].unique()}) if not comp_df.empty else ["All"]
                        t_opts = ["All"] + sorted({str(x) for x in comp_df["Type"].unique()}) if not comp_df.empty else ["All"]
                        with filter_cols[0]: pf = st.selectbox("Page", p_opts, key=f"t3_pf_{reset_id}")
                        with filter_cols[1]: sf = st.selectbox("Status", s_opts, key=f"t3_sf_{reset_id}")
                        with filter_cols[2]: tf = st.selectbox("Type", t_opts, key=f"t3_tf_{reset_id}")
                        _interactive_comparison_table(_render_detail_table(comp_df,pf,sf,tf), f"t3_comp_{reset_id}")
                    with lower_tabs[2]:
                        _render_issues(variable_results, static_rows, reset_id)

                # ========================= DETAILS ========================
                with top_tabs[1]:
                    detail_pages = [int(p.get("page",i+1)) for i,p in enumerate(output_pages)]
                    detail_page = st.selectbox("Page to inspect", detail_pages, key=f"t3_detail_page_{reset_id}")
                    di = detail_pages.index(detail_page)
                    orgp = org_pages[di] if di < len(org_pages) else None
                    outp = output_pages[di]
                    sm = compare_static_blocks(orgp,outp) if orgp else []
                    vr = by_page_blocks.get(detail_page,[])
                    ds = _build_page_structural_stats(orgp,outp,sm,vr)
                    dcols = st.columns(4)
                    for col,(n,lbl) in zip(dcols,[(ds['org_text_blocks'],'ORG text blocks'),(ds['output_text_blocks'],'Output text blocks'),(ds['static'],'Static matches'),(len(vr),'Selected field checks')]):
                        with col: st.markdown(f"<div class='t3-mini-stat'><b>{n}</b><span>{lbl}</span></div>", unsafe_allow_html=True)
                    st.markdown("### Static Elements")
                    if sm:
                        for m in sm:
                            metrics = _bbox_metrics(m['org'],m['output'],max(1,int(outp.get('image_width',1))),max(1,int(outp.get('image_height',1))))
                            st.markdown(f"<div class='t3-issue-item'><div class='t3-issue-head'><div class='t3-issue-field'>{_escape_html(m['org']['text'])}</div>{_status_chip('STATIC')}</div><div class='t3-issue-detail'>Position shift X: {metrics['dx']:.3f} • Y: {metrics['dy']:.3f} • Scale score: {metrics['scale_score']:.2f} • Match confidence: {m['score']:.2f}</div></div>", unsafe_allow_html=True)
                    else:
                        st.markdown("<div class='t3-empty'>No static elements matched within the current position tolerance.</div>", unsafe_allow_html=True)
                    st.markdown("### Variable Field Detail")
                    if vr:
                        for r in vr:
                            org_text = r['org_reference_block']['text'] if r.get('org_reference_block') else 'Not matched'
                            st.markdown(f"<div class='t3-issue-item'><div class='t3-issue-head'><div class='t3-issue-field'>{_escape_html(r['field'])}</div>{_status_chip(r['status'])}</div><div class='t3-issue-detail'><b>Type:</b> {_escape_html(r['field_type'])}<br><b>ORG:</b> {_escape_html(org_text)}<br><b>Order Form:</b> {_escape_html(r['expected'])}<br><b>Output:</b> {_escape_html(r['actual'])}<br><b>Finding:</b> {_escape_html(r['difference'])}</div></div>", unsafe_allow_html=True)
                    else:
                        st.markdown("<div class='t3-empty'>No variable fields selected for this page.</div>", unsafe_allow_html=True)

                # ========================= SUMMARY ========================
                with top_tabs[2]:
                    _render_summary_panel(static_rows, variable_results, product_type, output_pages, org_pages)
                    st.markdown("### Downloadable QC Report")
                    annotated_images = {}
                    for i,outp in enumerate(output_pages):
                        page_no = int(outp.get('page',i+1))
                        orgp = org_pages[i] if i < len(org_pages) else None
                        sm = compare_static_blocks(orgp,outp) if orgp else []
                        annotated_images[page_no] = _build_annotated_output(outp,sm,by_page_blocks.get(page_no,[]),i)
                    comp_df = _build_dataframe(static_rows, variable_results)
                    excel_bytes = _to_excel_bytes((overall,static_count,variable_count,issue_count,review_count,missing_count),comp_df,annotated_images,mapping,product_type)
                    st.download_button("⬇  Download QC Excel Report",data=excel_bytes,file_name="ORG_OrderForm_Output_QC_Report.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",width="stretch",key=f"t3_download_{reset_id}")
                    st.markdown("### Final QC Summary")
                    summary_df = pd.DataFrame([
                        ["Overall Result",overall],["Product Type",product_type],["Static Elements",static_count],
                        ["Variable Elements",variable_count],["Issues Found",issue_count],["Manual Review",review_count],
                        ["Unaccounted / Not Found",missing_count],["Output Pages",len(output_pages)]
                    ],columns=["Metric","Value"])
                    st.dataframe(summary_df,width="stretch",hide_index=True)

    # ============================== RIGHT RAIL ===========================
    with rail_col:
        result = st.session_state.get("t3_result")
        if result and st.session_state.get("t3_active_section") != "Settings":
            # Metrics are calculated from the current selected fields when on Home.
            # Before selected fields exist, show the initial structural scan counts.
            selected = st.session_state.get("t3_selected_fields") or result.get("detected_fields", [])
            df = result["df"]
            org_pages = result["org_pages"]
            output_pages = result["output_pages"]
            product_type = result["product_type"]
            mapping = result.get("mapping", {})
            # Repeat the same field validation for truthful rail metrics.
            vr_live = []
            if selected:
                try:
                    fields_for_match = order_fields_for_matching(selected)
                    osz_group_size = sum(1 for f in selected if get_field_type(f) == "OSZ")
                    for pi,out_page in enumerate(output_pages):
                        page_no = int(out_page.get("page",pi+1))
                        row_idx = mapping.get(page_no,min(pi,len(df)-1))
                        if 0 <= row_idx < len(df):
                            row = df.iloc[row_idx]
                            state = _new_output_state(out_page,product_type)
                            temp={}
                            for field in fields_for_match:
                                expected = "" if is_blank_value(row[field]) else str(row[field]).strip()
                                temp[field]=check_field(expected,field,state,osz_group_size=osz_group_size)
                            orgp = org_pages[pi] if pi < len(org_pages) else None
                            for field in selected:
                                expected = "" if is_blank_value(row[field]) else str(row[field]).strip()
                                base=temp.get(field,{"status":"SKIP","pdf":"—","difference":"Blank Order Form value — field ignored."})
                                vr_live.append(_format_field_result(field,expected,base,out_page,orgp,product_type))
                except Exception:
                    vr_live = []
            static_rows = result.get("static_rows", [])
            metrics = _summary_metrics(static_rows,vr_live)
            st.markdown(_render_right_summary(*metrics),unsafe_allow_html=True)
            st.markdown(_render_highlight_legend(),unsafe_allow_html=True)
            with st.container(border=True):
                st.markdown("<div class='t3-card-title'>Zoom & View</div>",unsafe_allow_html=True)
                z = st.columns(3)
                with z[0]:
                    if st.button("−", key=f"t3_rail_minus_{reset_id}", width="stretch"):
                        st.session_state["t3_zoom"] = max(60,st.session_state["t3_zoom"]-10); st.rerun()
                with z[1]: st.markdown(f"<div class='t3-zoom'>{st.session_state['t3_zoom']}%</div>",unsafe_allow_html=True)
                with z[2]:
                    if st.button("+", key=f"t3_rail_plus_{reset_id}", width="stretch"):
                        st.session_state["t3_zoom"] = min(150,st.session_state["t3_zoom"]+10); st.rerun()
                if st.button("Fit to Screen",key=f"t3_rail_fit_{reset_id}",width="stretch"):
                    st.session_state["t3_zoom"]=100; st.rerun()
        else:
            st.markdown('<div class="t3-result-card"><div class="t3-card-title">Result Summary</div><div class="t3-empty" style="margin-top:10px">Run a check to populate this panel.</div></div>',unsafe_allow_html=True)
            st.markdown(_render_highlight_legend(),unsafe_allow_html=True)

