import base64
import html
import io
import math
import re
import unicodedata
from collections import defaultdict
from copy import deepcopy

import fitz
import pandas as pd
import streamlit as st

from PIL import Image, ImageChops, ImageDraw, ImageFont

# ============================================================================
# TOOL 3
# ORG SPEC + ORDER FORM + OUTPUT QC
#
# Architectural rule:
#   - The existing Order Form -> Output engine remains the source of truth for
#     variable-data validation.
#   - Tool 3 adds ORG analysis, artwork registration, static/variable
#     classification, presentation checks, and interactive visualization.
#   - app.py / dashboard routing is not changed by this file.
# ============================================================================

try:
    from .order_form_output_check import (
        _apply_tool_css as _apply_base_css,
        build_page_state,
        check_field,
        extract_output_pages,
        get_available_fields,
        get_field_region,
        get_field_type,
        is_admin_field,
        is_blank_value,
        load_excel,
        normalize_text,
        order_fields_for_matching,
    )
except Exception:
    from order_form_output_check import (
        _apply_tool_css as _apply_base_css,
        build_page_state,
        check_field,
        extract_output_pages,
        get_available_fields,
        get_field_region,
        get_field_type,
        is_admin_field,
        is_blank_value,
        load_excel,
        normalize_text,
        order_fields_for_matching,
    )


TOOL3_VERSION = "2026-09-11-V3-ORG-BASELINE-REGISTERED-VIEWER"

# ----------------------------- visual palette ------------------------------
STATIC_YELLOW = (250, 204, 21, 110)
STATIC_OUTLINE = (250, 204, 21, 245)
PASS_GREEN = (34, 197, 94, 125)
FAIL_RED = (239, 68, 68, 145)
REVIEW_ORANGE = (249, 115, 22, 145)
MISSING_PURPLE = (168, 85, 247, 140)

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


# ============================================================================
# UI CSS
# ============================================================================

def _tool3_css():
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
        .block-container { max-width: 1540px !important; padding-top: 1rem !important; }

        .t3-header { display:flex; align-items:center; gap:14px; padding:2px 0 13px 0; }
        .t3-logo {
            width:42px; height:42px; border-radius:12px;
            background:linear-gradient(135deg,#0ea5e9,#2563eb);
            display:flex; align-items:center; justify-content:center;
            font-size:22px; box-shadow:0 0 18px rgba(37,99,235,.25);
        }
        .t3-title { font-size:27px; font-weight:800; line-height:1.08; color:#f8fafc; }
        .t3-subtitle { font-size:12px; color:#8fa1b8; margin-top:4px; }
        .t3-rule { height:34px; width:1px; background:#334155; }

        .t3-side, .t3-panel, .t3-card, .t3-rail-card {
            background:#0b1627; border:1px solid #213249; border-radius:11px;
        }
        .t3-side { padding:8px; min-height:650px; }
        .t3-side-note { margin:16px 6px 0 6px; font-size:10px; color:#718399; line-height:1.45; }

        .t3-upload-card {
            background:#0d1829; border:1px solid #22334a; border-radius:10px;
            padding:12px; min-height:150px;
        }
        .t3-card-title { font-size:13px; font-weight:760; color:#f5f8fc; }
        .t3-card-sub { font-size:10px; color:#7f92a8; margin-top:2px; }
        .t3-fileline { font-size:10px; color:#c4d0df; margin-top:8px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
        .t3-ok { color:#22c55e; }
        .t3-bad { color:#fb7185; }
        .t3-control-label { font-size:10px; color:#8fa1b8; margin-bottom:3px; }

        [data-testid="stFileUploader"] {
            background:#0a1423 !important; border:1px solid #263a53 !important;
            border-radius:8px !important; padding:3px !important;
        }
        [data-testid="stFileUploaderDropzone"] {
            background:#0b1524 !important; border:1px dashed #3a5270 !important;
            border-radius:7px !important;
        }
        [data-testid="stFileUploaderDropzoneInstructions"] { font-size:10px !important; }
        [data-testid="stFileUploaderDropzoneInstructions"] span { color:#d8e4f1 !important; }

        [data-baseweb="select"] > div {
            background:#0b1524 !important; border:1px solid #2a3d56 !important;
            color:#f8fafc !important; border-radius:8px !important;
        }
        [data-baseweb="select"] span, [data-baseweb="select"] input { color:#f8fafc !important; }

        .t3-toolbar { display:flex; justify-content:space-between; align-items:center; padding:11px 13px; border-bottom:1px solid #1d2b3f; }
        .t3-toolbar-title { font-size:13px; font-weight:760; }
        .t3-toolbar-meta { font-size:10px; color:#7f92a8; }
        .t3-mini-stat { background:#0d192a; border:1px solid #22334a; border-radius:9px; padding:10px 12px; }
        .t3-mini-stat b { display:block; font-size:20px; color:#f7fbff; }
        .t3-mini-stat span { font-size:9px; color:#8598ad; }

        .t3-empty { color:#8396ab; font-size:11px; padding:18px 4px; }
        .t3-badge { display:inline-block; padding:4px 8px; border-radius:999px; font-size:9px; font-weight:800; }
        .t3-badge.pass { background:#14532d; color:#bbf7d0; }
        .t3-badge.fail { background:#7f1d1d; color:#fecaca; }
        .t3-badge.review { background:#7c2d12; color:#fed7aa; }
        .t3-badge.static { background:#665000; color:#fef08a; }
        .t3-badge.variable { background:#123d79; color:#bfdbfe; }
        .t3-badge.missing { background:#581c87; color:#e9d5ff; }
        .t3-badge.info { background:#334155; color:#dbe7f4; }

        .t3-issue { border:1px solid #203149; border-radius:9px; background:#0d1828; margin:8px 0; padding:10px 11px; }
        .t3-issue-head { display:flex; justify-content:space-between; gap:10px; align-items:center; }
        .t3-issue-field { font-size:11px; font-weight:760; color:#f3f6fb; }
        .t3-issue-detail { font-size:10px; color:#93a6bb; line-height:1.5; margin-top:7px; }
        .t3-note { font-size:10px; color:#8395aa; line-height:1.55; }
        .t3-view-note { font-size:10px; color:#7e91a8; margin:6px 0 0 2px; }
        .t3-section-title { font-size:14px; font-weight:760; margin:10px 0 6px 0; }
        .t3-table-note { font-size:9px; color:#74889f; margin-top:4px; }
        .t3-metric-big { font-size:23px; font-weight:830; }
        .t3-metric-small { font-size:9px; color:#8194aa; }
        </style>
        """,
        unsafe_allow_html=True,
    )


# ============================================================================
# NORMALIZATION / PRESENTATION ANALYSIS
# ============================================================================

def _norm_casefold(text):
    value = unicodedata.normalize("NFKC", str(text or "")).strip()
    value = re.sub(r"\s+", " ", value)
    return value.casefold()


def _compact(text):
    return re.sub(r"[^a-z0-9]+", "", _norm_casefold(text))


def _punct_signature(text):
    value = unicodedata.normalize("NFKC", str(text or ""))
    return [ch for ch in value if not ch.isalnum() and not ch.isspace()]


def _word_case_mode(word):
    word = re.sub(r"[^A-Za-z]", "", str(word or ""))
    if not word:
        return "NONE"
    if word.isupper():
        return "UPPER"
    if word.islower():
        return "LOWER"
    if len(word) > 1 and word[0].isupper() and word[1:].islower():
        return "TITLE"
    return "MIXED"


def _case_requirement(reference, actual):
    """Return whether Output follows the ORG's visible case convention.

    The ORG remains the presentation baseline even when the actual data is
    different because it was supplied by the Order Form.
    """
    ref_words = re.findall(r"[A-Za-z]+", str(reference or ""))
    out_words = re.findall(r"[A-Za-z]+", str(actual or ""))
    if not ref_words or not out_words:
        return True, "No meaningful alphabetic case evidence."

    ref_modes = [_word_case_mode(w) for w in ref_words]
    out_modes = [_word_case_mode(w) for w in out_words]

    # A single consistent ORG style should be enforced across the variable
    # replacement words, even when word count changes.
    non_none = [m for m in ref_modes if m != "NONE"]
    if non_none:
        dominant = max(set(non_none), key=non_none.count)
        dominant_ratio = non_none.count(dominant) / len(non_none)
        if dominant_ratio >= 0.75:
            bad = [m for m in out_modes if m != dominant]
            if not bad:
                return True, f"ORG establishes {dominant.lower()} presentation."
            return False, f"ORG establishes {dominant.lower()} presentation; Output contains a different letter-case pattern."

    # Mixed ORG style: compare positions that exist and tolerate variable word
    # count by checking the common prefix and the dominant residual style.
    common = min(len(ref_modes), len(out_modes))
    if common:
        position_matches = sum(ref_modes[i] == out_modes[i] for i in range(common))
        if position_matches / common >= 0.75:
            return True, "Output follows the visible ORG case pattern."
    return False, "Output does not preserve the visible ORG case pattern."


def _presentation_checks(org_text, output_text, product_type):
    checks = []
    case_pass, case_reason = _case_requirement(org_text, output_text)
    if product_type == "PFL":
        checks.append({
            "kind": "CASE",
            "status": "INFO",
            "reason": "PFL presentation is governed by Order Form case; ORG case is retained as reference only.",
        })
    else:
        checks.append({
            "kind": "CASE",
            "status": "PASS" if case_pass else "FAIL",
            "reason": case_reason,
        })

    org_punct = _punct_signature(org_text)
    out_punct = _punct_signature(output_text)
    punct_pass = org_punct == out_punct
    checks.append({
        "kind": "PUNCTUATION",
        "status": "PASS" if punct_pass else "FAIL",
        "reason": (
            "Output punctuation matches the ORG baseline."
            if punct_pass
            else f"ORG punctuation {org_punct or ['(none)']} differs from Output {out_punct or ['(none)']}."
        ),
    })

    org_tokens = _norm_casefold(org_text).split()
    out_tokens = _norm_casefold(output_text).split()
    checks.append({
        "kind": "TEXT STRUCTURE",
        "status": "PASS" if org_tokens and out_tokens else "REVIEW",
        "reason": "Text exists on both ORG and Output for presentation comparison." if org_tokens and out_tokens else "Insufficient text evidence.",
    })
    return checks


# ============================================================================
# OCR BLOCKS / ARTWORK REGISTRATION
# ============================================================================

def _word_bbox(words):
    if not words:
        return None
    left = min(int(w.get("left", 0)) for w in words)
    top = min(int(w.get("top", 0)) for w in words)
    right = max(int(w.get("left", 0)) + int(w.get("width", 0)) for w in words)
    bottom = max(int(w.get("top", 0)) + int(w.get("height", 0)) for w in words)
    return (left, top, right, bottom) if right > left and bottom > top else None


def _make_visual_blocks(page):
    words = [w for w in page.get("ocr_words", []) if str(w.get("text", "")).strip()]
    grouped = defaultdict(list)
    for word in words:
        key = (word.get("block_num", 0), word.get("par_num", 0), word.get("line_num", 0))
        grouped[key].append(word)

    blocks = []
    for key, items in grouped.items():
        items = sorted(items, key=lambda x: (int(x.get("top", 0)), int(x.get("left", 0))))
        bbox = _word_bbox(items)
        if not bbox:
            continue
        text = " ".join(str(x.get("text", "")).strip() for x in items).strip()
        if not text or not re.search(r"[A-Za-z0-9]", text):
            continue
        l, t, r, b = bbox
        blocks.append({
            "key": key,
            "text": text,
            "norm": _norm_casefold(text),
            "compact": _compact(text),
            "bbox": bbox,
            "cx": (l + r) / 2.0,
            "cy": (t + b) / 2.0,
            "width": max(1, r - l),
            "height": max(1, b - t),
            "words": items,
        })
    blocks.sort(key=lambda b: (b["bbox"][1], b["bbox"][0]))
    for idx, block in enumerate(blocks):
        block["index"] = idx
    return blocks


def _content_bbox(image):
    """Find a practical artwork/content boundary against white page margins."""
    if image is None:
        return None
    rgb = image.convert("RGB")
    bg = Image.new("RGB", rgb.size, (255, 255, 255))
    diff = ImageChops.difference(rgb, bg)
    # Increase sensitivity slightly while suppressing near-white compression noise.
    gray = diff.convert("L").point(lambda p: 0 if p >= 10 else 255)
    bbox = gray.getbbox()
    if not bbox:
        return (0, 0, rgb.width, rgb.height)

    l, t, r, b = bbox
    bw = r - l
    bh = b - t
    # Avoid tiny dust/metadata marks becoming the artwork boundary.
    if bw < rgb.width * 0.05 or bh < rgb.height * 0.05:
        return (0, 0, rgb.width, rgb.height)
    return bbox


def _registration(org_page, output_page):
    """Register ORG artwork space into Output coordinates using content bounds.

    Uniform scaling preserves aspect ratio. Translation aligns content-box
    centers. The algorithm intentionally ignores white page margins.
    """
    org_raw = org_page.get("image_bytes") if org_page else None
    out_raw = output_page.get("image_bytes") if output_page else None
    if not org_raw or not out_raw:
        return None

    org_img = Image.open(io.BytesIO(org_raw)).convert("RGB")
    out_img = Image.open(io.BytesIO(out_raw)).convert("RGB")
    ob = _content_bbox(org_img)
    ub = _content_bbox(out_img)

    ow = max(1, ob[2] - ob[0])
    oh = max(1, ob[3] - ob[1])
    uw = max(1, ub[2] - ub[0])
    uh = max(1, ub[3] - ub[1])

    sx = uw / ow
    sy = uh / oh
    scale = (sx + sy) / 2.0
    # Keep registration robust against wildly different scans/canvas sizes.
    scale = min(4.0, max(0.25, scale))

    org_center = ((ob[0] + ob[2]) / 2.0, (ob[1] + ob[3]) / 2.0)
    out_center = ((ub[0] + ub[2]) / 2.0, (ub[1] + ub[3]) / 2.0)
    tx = out_center[0] - org_center[0] * scale
    ty = out_center[1] - org_center[1] * scale

    aspect_delta = abs(sx - sy) / max(1e-6, (sx + sy) / 2.0)
    warning = ""
    if aspect_delta > 0.12:
        warning = "ORG and Output artwork boundaries have noticeably different aspect ratios; registration uses uniform scale and center anchoring."
    elif aspect_delta > 0.06:
        warning = "Minor aspect-ratio difference detected; uniform scale is being used."

    return {
        "org_image": org_img,
        "output_image": out_img,
        "org_bbox": ob,
        "output_bbox": ub,
        "scale": scale,
        "scale_x": sx,
        "scale_y": sy,
        "tx": tx,
        "ty": ty,
        "aspect_delta": aspect_delta,
        "warning": warning,
        "method": "content-boundary / uniform-scale / center-anchor",
    }


def _transform_bbox(bbox, reg):
    l, t, r, b = bbox
    s = reg["scale"]
    tx = reg["tx"]
    ty = reg["ty"]
    return (l * s + tx, t * s + ty, r * s + tx, b * s + ty)


def _transform_block(block, reg):
    new = dict(block)
    new_bbox = _transform_bbox(block["bbox"], reg)
    l, t, r, b = new_bbox
    new["bbox"] = new_bbox
    new["cx"] = (l + r) / 2.0
    new["cy"] = (t + b) / 2.0
    new["width"] = max(1.0, r - l)
    new["height"] = max(1.0, b - t)
    new["registered"] = True
    return new


def _registered_org_blocks(org_page, reg):
    if not org_page or not reg:
        return []
    return [_transform_block(b, reg) for b in _make_visual_blocks(org_page)]


def _bbox_metrics(a, b, page_w, page_h):
    acx, acy = a["cx"], a["cy"]
    bcx, bcy = b["cx"], b["cy"]
    dx = abs(acx - bcx) / max(1.0, page_w)
    dy = abs(acy - bcy) / max(1.0, page_h)

    al, at, ar, ab = a["bbox"]
    bl, bt, br, bb = b["bbox"]
    aw = max(1.0, ar - al)
    ah = max(1.0, ab - at)
    bw = max(1.0, br - bl)
    bh = max(1.0, bb - bt)

    width_ratio = bw / aw
    height_ratio = bh / ah
    scale_score = min(width_ratio, 1.0 / width_ratio) * min(height_ratio, 1.0 / height_ratio)
    left_shift = abs(al - bl) / max(1.0, page_w)
    right_shift = abs(ar - br) / max(1.0, page_w)
    best_anchor = min(left_shift, right_shift, dx)
    diagonal = math.hypot(page_w, page_h)
    center_diag = math.hypot(acx - bcx, acy - bcy) / max(1.0, diagonal)
    return {
        "dx": dx,
        "dy": dy,
        "center_diag": center_diag,
        "scale_score": scale_score,
        "width_ratio": width_ratio,
        "height_ratio": height_ratio,
        "best_anchor": best_anchor,
    }


def _static_match_score(org_block, out_block, page_w, page_h):
    if org_block["norm"] != out_block["norm"] and org_block["compact"] != out_block["compact"]:
        return None

    m = _bbox_metrics(org_block, out_block, page_w, page_h)
    if m["center_diag"] > 0.050:
        return None
    if m["best_anchor"] > 0.040:
        return None
    if m["width_ratio"] < 0.55 or m["width_ratio"] > 1.85:
        return None
    if m["height_ratio"] < 0.50 or m["height_ratio"] > 1.90:
        return None

    score = (
        (1.0 - min(1.0, m["center_diag"] / 0.050)) * 0.55
        + m["scale_score"] * 0.30
        + (1.0 - min(1.0, m["best_anchor"] / 0.040)) * 0.15
    )
    return max(0.0, min(1.0, score))


def _compare_static_page(org_page, output_page, reg):
    if not org_page or not output_page:
        return []
    org_blocks = _registered_org_blocks(org_page, reg)
    out_blocks = _make_visual_blocks(output_page)
    page_w = max(1, int(output_page.get("image_width", 1)))
    page_h = max(1, int(output_page.get("image_height", 1)))

    candidates = []
    for oi, org in enumerate(org_blocks):
        if len(org["compact"]) < 2:
            continue
        for ui, out in enumerate(out_blocks):
            score = _static_match_score(org, out, page_w, page_h)
            if score is not None:
                candidates.append((score, oi, ui))

    candidates.sort(reverse=True)
    used_org, used_out = set(), set()
    matches = []
    for score, oi, ui in candidates:
        if oi in used_org or ui in used_out:
            continue
        used_org.add(oi)
        used_out.add(ui)
        matches.append({
            "org": org_blocks[oi],
            "output": out_blocks[ui],
            "org_index": oi,
            "output_index": ui,
            "score": score,
            "classification": "STATIC",
            "status": "STATIC",
        })
    return matches


def _nearest_registered_org_block(output_block, org_blocks, excluded=None):
    excluded = excluded or set()
    if not output_block:
        return None, None
    best = None
    best_score = 999.0
    for idx, block in enumerate(org_blocks):
        if idx in excluded:
            continue
        dx = output_block["cx"] - block["cx"]
        dy = output_block["cy"] - block["cy"]
        distance = math.hypot(dx, dy)
        size_penalty = abs(math.log(max(.01, output_block["width"]) / max(.01, block["width"])))
        score = distance + size_penalty * 0.18 * max(output_block["width"], output_block["height"], 1)
        if score < best_score:
            best_score = score
            best = (idx, block)
    return best if best else (None, None)


def _find_blocks_for_text(page, target):
    target_norm = _norm_casefold(target)
    target_compact = _compact(target)
    if not target_norm or target_norm in {"not found", "—", "-"}:
        return []
    blocks = _make_visual_blocks(page)
    exact = [b for b in blocks if b["norm"] == target_norm or b["compact"] == target_compact]
    if exact:
        return exact
    return [
        b for b in blocks
        if target_compact and (target_compact in b["compact"] or b["compact"] in target_compact)
    ]


def _field_color(field_name):
    return FIELD_COLORS.get(get_field_type(field_name), FIELD_COLORS["GENERAL"])


def _clean(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def _field_type_label(field_name):
    kind = get_field_type(field_name)
    region = get_field_region(field_name)
    return f"{kind}{' • ' + region if region else ''}"


# ============================================================================
# ORG BASELINE / VARIABLE RESULT ENRICHMENT
# ============================================================================

def _page_baseline(org_page, output_page):
    reg = _registration(org_page, output_page)
    static_matches = _compare_static_page(org_page, output_page, reg) if reg else []
    return {
        "registration": reg,
        "static_matches": static_matches,
        "registered_org_blocks": _registered_org_blocks(org_page, reg),
        "output_blocks": _make_visual_blocks(output_page),
    }


def _presentation_status(org_block, output_block, product_type):
    if not org_block or not output_block:
        return "REVIEW", "No ORG presentation reference could be mapped to this Output field."
    checks = _presentation_checks(org_block["text"], output_block["text"], product_type)
    failures = [c for c in checks if c["status"] == "FAIL"]
    if failures:
        return "FAIL", " ".join(c["reason"] for c in failures)
    return "PASS", "ORG presentation checks pass for the mapped variable region."


def _enrich_variable_result(field, expected, base_result, output_page, baseline, product_type):
    actual = _clean(base_result.get("pdf", ""))
    if actual.casefold() in {"not found", "—", "-"}:
        actual = ""

    output_candidates = _find_blocks_for_text(output_page, actual)
    output_block = output_candidates[0] if output_candidates else None

    static_out_indexes = {m["output_index"] for m in baseline["static_matches"]}
    static_org_indexes = {m["org_index"] for m in baseline["static_matches"]}
    org_idx, org_block = _nearest_registered_org_block(
        output_block,
        baseline["registered_org_blocks"],
        excluded=static_org_indexes,
    )

    base_status = base_result.get("status", "NOT FOUND")
    status = base_status
    difference = base_result.get("difference", "—") or "—"
    presentation_status = "N/A"
    presentation_reason = ""

    if base_status in {"PASS", "FAIL"} and product_type in {"HTL", "Other"}:
        presentation_status, presentation_reason = _presentation_status(org_block, output_block, product_type)
        if base_status == "PASS" and presentation_status == "FAIL":
            status = "FAIL"
            difference = f"{difference if difference != '—' else ''} {presentation_reason}".strip()
        elif base_status == "PASS" and presentation_status == "REVIEW":
            status = "REVIEW"
            difference = presentation_reason

    # For PFL, keep the proven Tool 1 result as the decision authority.
    classification = "VARIABLE" if status in {"PASS", "FAIL", "REVIEW"} else "NOT SHOWN"

    return {
        "field": field,
        "field_type": _field_type_label(field),
        "expected": expected,
        "actual": actual or "Not found",
        "status": status,
        "base_status": base_status,
        "classification": classification,
        "difference": difference,
        "match_type": base_result.get("match_type", ""),
        "output_block": output_block,
        "org_reference_block": org_block,
        "org_reference_index": org_idx,
        "presentation_status": presentation_status,
        "presentation_reason": presentation_reason,
    }


def _run_shared_field_checks(df, output_pages, org_pages, selected_fields, product_type, mapping):
    rows = []
    if not selected_fields:
        return rows

    fields_for_match = order_fields_for_matching(selected_fields)
    osz_group_size = sum(1 for f in selected_fields if get_field_type(f) == "OSZ")

    for page_idx, output_page in enumerate(output_pages):
        page_no = int(output_page.get("page", page_idx + 1))
        row_idx = int(mapping.get(page_no, min(page_idx, len(df) - 1)))
        if row_idx < 0 or row_idx >= len(df):
            continue
        row = df.iloc[row_idx]
        state = build_page_state(output_page, product_type)
        page_results = {}

        for field in fields_for_match:
            expected = "" if is_blank_value(row[field]) else str(row[field]).strip()
            if not expected:
                continue
            page_results[field] = check_field(expected, field, state, osz_group_size=osz_group_size)

        org_page = org_pages[page_idx] if page_idx < len(org_pages) else None
        baseline = _page_baseline(org_page, output_page)
        for field in selected_fields:
            expected = "" if is_blank_value(row[field]) else str(row[field]).strip()
            if not expected:
                # Preserve Tool 1's blank-field behavior: no row and no table clutter.
                continue
            base = page_results.get(field)
            if base is None:
                continue
            enriched = _enrich_variable_result(field, expected, base, output_page, baseline, product_type)
            enriched.update({"PDF PAGE": page_no, "EXCEL ROW": row_idx + 2})
            if enriched["status"] != "NOT FOUND":
                rows.append(enriched)
    return rows


def _dynamic_auto_detect(df, output_pages, product_type, mapping):
    candidates = [f for f in get_available_fields(df) if not is_admin_field(f)]
    detected = []
    for field in candidates:
        found = False
        for page_idx, output_page in enumerate(output_pages):
            page_no = int(output_page.get("page", page_idx + 1))
            row_idx = int(mapping.get(page_no, min(page_idx, len(df) - 1)))
            if row_idx < 0 or row_idx >= len(df):
                continue
            expected = "" if is_blank_value(df.iloc[row_idx][field]) else str(df.iloc[row_idx][field]).strip()
            if not expected:
                continue
            state = build_page_state(output_page, product_type)
            result = check_field(expected, field, state, osz_group_size=1)
            if result.get("status") in {"PASS", "FAIL"}:
                found = True
                break
        if found:
            detected.append(field)
    return detected


def _unaccounted_org_rows(org_pages, output_pages, baselines):
    """Show only genuine ORG elements with no reasonable Output counterpart.

    This is deliberately different from listing every Tool-1 NOT FOUND field.
    """
    rows = []
    for idx, baseline in enumerate(baselines):
        if idx >= len(org_pages) or idx >= len(output_pages):
            continue
        org_blocks = baseline["registered_org_blocks"]
        out_blocks = baseline["output_blocks"]
        static_org = {m["org_index"] for m in baseline["static_matches"]}
        static_out = {m["output_index"] for m in baseline["static_matches"]}
        page_no = int(output_pages[idx].get("page", idx + 1))
        page_w = max(1, int(output_pages[idx].get("image_width", 1)))
        page_h = max(1, int(output_pages[idx].get("image_height", 1)))

        for oi, org in enumerate(org_blocks):
            if oi in static_org:
                continue
            if len(org.get("compact", "")) < 2:
                continue
            best = None
            best_score = 999.0
            for ui, out in enumerate(out_blocks):
                if ui in static_out:
                    continue
                m = _bbox_metrics(org, out, page_w, page_h)
                score = m["center_diag"] + abs(math.log(max(.01, m["width_ratio"]))) * 0.12
                if score < best_score:
                    best_score = score
                    best = out
            # If an Output block sits in essentially the same registered region,
            # it is accounted for even if its wording is different/variable.
            if best is not None and best_score <= 0.065:
                continue

            rows.append({
                "PDF PAGE": page_no,
                "field": "ORG ELEMENT",
                "field_type": "ORG BASELINE",
                "expected": org["text"],
                "actual": "Not found",
                "status": "MISSING / UNACCOUNTED",
                "base_status": "NOT FOUND",
                "classification": "MISSING / UNACCOUNTED",
                "difference": "ORG element has no reasonable Output counterpart at the registered position.",
                "match_type": "",
                "output_block": None,
                "org_reference_block": org,
                "presentation_status": "N/A",
                "presentation_reason": "",
            })
    return rows


# ============================================================================
# IMAGE ANNOTATION / VIEWER DATA
# ============================================================================

def _resize_width(image, max_width=900):
    if image is None:
        return None
    if image.width <= max_width:
        return image.copy()
    scale = max_width / image.width
    return image.resize((max(1, int(image.width * scale)), max(1, int(image.height * scale))), Image.Resampling.LANCZOS)


def _draw_label(draw, xy, label, fill):
    x, y = xy
    try:
        font = ImageFont.load_default()
    except Exception:
        font = None
    pad = 4
    bbox = draw.textbbox((0, 0), label, font=font)
    w = bbox[2] - bbox[0] + pad * 2
    h = bbox[3] - bbox[1] + pad * 2
    y0 = max(0, y - h)
    draw.rounded_rectangle([x, y0, x + w, y], radius=4, fill=fill[:3] + (225,))
    draw.text((x + pad, y0 + pad - 1), label, fill=(255, 255, 255, 255), font=font)


def _paint_box(image, bbox, fill, label=None, outline=None, width=3):
    left, top, right, bottom = [int(round(v)) for v in bbox]
    left = max(0, min(image.width - 1, left))
    right = max(left + 1, min(image.width, right))
    top = max(0, min(image.height - 1, top))
    bottom = max(top + 1, min(image.height, bottom))
    overlay = Image.new("RGBA", image.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    if outline is None:
        outline = fill
    draw.rounded_rectangle([left, top, right, bottom], radius=4, fill=fill, outline=outline, width=width)
    if label:
        _draw_label(draw, (left, max(top, 26)), label, outline)
    return Image.alpha_composite(image.convert("RGBA"), overlay)


def _build_annotated_output(output_page, static_matches, variable_rows):
    raw = output_page.get("image_bytes")
    if not raw:
        return None
    image = Image.open(io.BytesIO(raw)).convert("RGBA")

    for match in static_matches:
        image = _paint_box(
            image,
            match["output"]["bbox"],
            STATIC_YELLOW,
            "STATIC",
            outline=STATIC_OUTLINE,
            width=3,
        )

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
            color = _field_color(row["field"])
            fill, outline, label = color, color, f"{row['field']} • PASS"
        else:
            continue
        image = _paint_box(image, block["bbox"], fill, label, outline=outline, width=3)
    return image.convert("RGB")


def _build_registered_org_image(reg):
    if not reg:
        return None
    out_img = reg["output_image"].convert("RGB")
    org_img = reg["org_image"].convert("RGBA")
    scaled_size = (
        max(1, int(round(org_img.width * reg["scale"]))),
        max(1, int(round(org_img.height * reg["scale"]))),
    )
    scaled = org_img.resize(scaled_size, Image.Resampling.LANCZOS)
    canvas = Image.new("RGBA", out_img.size, (255, 255, 255, 255))
    canvas.alpha_composite(scaled, (int(round(reg["tx"])), int(round(reg["ty"]))))
    return canvas.convert("RGB")


def _image_data_uri(image, max_width=1400, quality=88):
    if image is None:
        return ""
    resized = _resize_width(image, max_width)
    buf = io.BytesIO()
    resized.save(buf, format="JPEG", quality=quality, optimize=True)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode("ascii")


def _viewer_html(raw_org_image, registered_org_image, output_image, height=670):
    raw_uri = _image_data_uri(raw_org_image)
    reg_uri = _image_data_uri(registered_org_image or raw_org_image)
    out_uri = _image_data_uri(output_image)
    template = r'''
    <div id="viewer" style="height:__HEIGHT__px;background:#070f1c;border:1px solid #22334a;border-radius:10px;overflow:hidden;position:relative;font-family:Inter,Arial,sans-serif;color:#e5edf7">
      <div style="height:42px;display:flex;gap:6px;align-items:center;padding:6px 8px;border-bottom:1px solid #22334a;background:#0b1627;position:absolute;top:0;left:0;right:0;z-index:10">
        <button data-mode="side" class="vbtn active">Side by Side</button>
        <button data-mode="overlay" class="vbtn">Overlay</button>
        <button data-mode="blink" class="vbtn">Blink</button>
        <span style="width:8px"></span>
        <button id="minus" class="vbtn">−</button><span id="zoomtxt" style="font-size:10px;min-width:42px;text-align:center">100%</span><button id="plus" class="vbtn">+</button>
        <button id="fit" class="vbtn">FIT</button><button id="reset" class="vbtn">RESET</button>
        <span id="blinkcontrols" style="display:none;align-items:center;gap:4px;margin-left:4px">
          <span style="font-size:10px;color:#92a4b8">Blink</span>
          <button data-ms="1300" class="speed vbtn active">Slow</button>
          <button data-ms="800" class="speed vbtn">Normal</button>
          <button data-ms="450" class="speed vbtn">Fast</button>
        </span>
        <span id="overlaycontrols" style="display:none;align-items:center;gap:5px;margin-left:4px;font-size:10px;color:#92a4b8">ORG <input id="opacity" type="range" min="5" max="95" value="50" style="width:90px"><span id="opacityTxt">50%</span></span>
        <span style="margin-left:auto;font-size:9px;color:#6f8298">Wheel = Zoom • Drag = Pan</span>
      </div>
      <div id="stage" style="position:absolute;inset:42px 0 0 0;overflow:hidden;cursor:grab;background:#050c16">
        <div id="canvas" style="position:absolute;left:50%;top:50%;transform-origin:center center;display:flex;align-items:center;gap:14px;will-change:transform">
          <div id="side-org" style="position:relative;background:#fff"><img id="rawOrg" src="__RAW__" style="display:block;max-width:47vw;max-height:590px"></div>
          <div id="side-out" style="position:relative;background:#fff"><img id="out" src="__OUT__" style="display:block;max-width:47vw;max-height:590px"></div>
          <div id="overlayBox" style="display:none;position:relative;background:#fff"><img id="overlayOut" src="__OUT__" style="display:block;max-width:76vw;max-height:590px"><img id="overlayOrg" src="__REG__" style="display:block;position:absolute;inset:0;width:100%;height:100%;object-fit:fill;opacity:.5"></div>
          <div id="blinkBox" style="display:none;position:relative;background:#fff"><img id="blinkOrg" src="__REG__" style="display:none;max-width:76vw;max-height:590px"><img id="blinkOut" src="__OUT__" style="display:block;max-width:76vw;max-height:590px"></div>
        </div>
      </div>
    </div>
    <style>
      .vbtn{background:#0d1b2e;color:#cdd9e7;border:1px solid #28405b;border-radius:6px;padding:5px 8px;font-size:10px;cursor:pointer}
      .vbtn:hover{background:#123050;border-color:#2d78b5}
      .vbtn.active{background:#0b4274;border-color:#2f91de;color:#fff}
    </style>
    <script>
    const stage=document.getElementById('stage'), canvas=document.getElementById('canvas');
    const sideOrg=document.getElementById('side-org'), sideOut=document.getElementById('side-out');
    const overlayBox=document.getElementById('overlayBox'), blinkBox=document.getElementById('blinkBox');
    const overlayOrg=document.getElementById('overlayOrg'), blinkOrg=document.getElementById('blinkOrg'), blinkOut=document.getElementById('blinkOut');
    const zoomtxt=document.getElementById('zoomtxt');
    const overlaycontrols=document.getElementById('overlaycontrols'), blinkcontrols=document.getElementById('blinkcontrols');
    let mode='side', zoom=1, px=0, py=0, blinkTimer=null, blinkState=false, blinkMs=800;
    let drag=false, sx=0, sy=0, ox=0, oy=0;

    function apply(){
      canvas.style.transform='translate(-50%,-50%) translate('+px+'px,'+py+'px) scale('+zoom+')';
      zoomtxt.textContent=Math.round(zoom*100)+'%';
    }
    function stopBlink(){ if(blinkTimer){clearInterval(blinkTimer);blinkTimer=null;} }
    function hideAll(){
      sideOrg.style.display='none';sideOut.style.display='none';overlayBox.style.display='none';blinkBox.style.display='none';
      overlaycontrols.style.display='none';blinkcontrols.style.display='none';
    }
    function setMode(m){
      stopBlink(); mode=m;
      document.querySelectorAll('.vbtn[data-mode]').forEach(b=>b.classList.toggle('active',b.dataset.mode===m));
      hideAll();
      if(m==='side'){sideOrg.style.display='block';sideOut.style.display='block';}
      if(m==='overlay'){overlayBox.style.display='block';overlaycontrols.style.display='flex';}
      if(m==='blink'){blinkBox.style.display='block';blinkcontrols.style.display='flex';startBlink();}
      apply();
    }
    function startBlink(){
      if(blinkTimer) clearInterval(blinkTimer);
      blinkState=false; blinkOrg.style.display='none'; blinkOut.style.display='block';
      blinkTimer=setInterval(function(){blinkState=!blinkState;blinkOrg.style.display=blinkState?'block':'none';blinkOut.style.display=blinkState?'none':'block';},blinkMs);
    }
    document.querySelectorAll('.vbtn[data-mode]').forEach(b=>b.addEventListener('click',function(){setMode(b.dataset.mode);}));
    document.querySelectorAll('.speed').forEach(b=>b.addEventListener('click',function(){blinkMs=parseInt(b.dataset.ms);document.querySelectorAll('.speed').forEach(x=>x.classList.remove('active'));b.classList.add('active');if(mode==='blink')startBlink();}));
    document.getElementById('plus').onclick=function(){zoom=Math.min(3,zoom+0.1);apply();};
    document.getElementById('minus').onclick=function(){zoom=Math.max(0.3,zoom-0.1);apply();};
    document.getElementById('fit').onclick=function(){zoom=1;px=0;py=0;apply();};
    document.getElementById('reset').onclick=function(){zoom=1;px=0;py=0;apply();};
    document.getElementById('opacity').oninput=function(e){overlayOrg.style.opacity=(parseInt(e.target.value)/100).toFixed(2);document.getElementById('opacityTxt').textContent=e.target.value+'%';};
    stage.addEventListener('wheel',function(e){e.preventDefault();zoom=Math.max(0.3,Math.min(3,zoom+(e.deltaY<0?0.1:-0.1)));apply();},{passive:false});
    stage.addEventListener('mousedown',function(e){drag=true;stage.style.cursor='grabbing';sx=e.clientX;sy=e.clientY;ox=px;oy=py;});
    window.addEventListener('mouseup',function(){drag=false;stage.style.cursor='grab';});
    window.addEventListener('mousemove',function(e){if(!drag)return;px=ox+(e.clientX-sx);py=oy+(e.clientY-sy);apply();});
    setMode('side');
    </script>
    '''
    return template.replace('__HEIGHT__', str(int(height))).replace('__RAW__', raw_uri).replace('__REG__', reg_uri).replace('__OUT__', out_uri)


# ============================================================================
# REPORT / TABLES / UI HELPERS
# ============================================================================

def _status_badge(text):
    text = str(text or "")
    cls = {
        "PASS": "pass",
        "FAIL": "fail",
        "REVIEW": "review",
        "STATIC": "static",
        "VARIABLE": "variable",
        "MISSING / UNACCOUNTED": "missing",
    }.get(text, "info")
    return f'<span class="t3-badge {cls}">{html.escape(text)}</span>'


def _build_static_rows(baselines, output_pages):
    rows = []
    for idx, baseline in enumerate(baselines):
        page_no = int(output_pages[idx].get("page", idx + 1))
        for m in baseline["static_matches"]:
            metrics = _bbox_metrics(
                m["org"], m["output"],
                max(1, int(output_pages[idx].get("image_width", 1))),
                max(1, int(output_pages[idx].get("image_height", 1))),
            )
            rows.append({
                "PDF PAGE": page_no,
                "Element / Field": "Static Element",
                "Type": "STATIC",
                "ORG Spec": m["org"]["text"],
                "Output": m["output"]["text"],
                "Order Form": "—",
                "Status": "STATIC",
                "Notes": f"Position Δ X {metrics['dx']:.3f}, Y {metrics['dy']:.3f}; scale {metrics['scale_score']:.2f}; confidence {m['score']:.2f}.",
                "org_block": m["org"],
                "output_block": m["output"],
            })
    return rows


def _build_comparison_df(static_rows, variable_rows, unaccounted_rows):
    rows = []
    for r in static_rows:
        rows.append({
            "PDF PAGE": r["PDF PAGE"], "Element / Field": r["Element / Field"], "Type": "STATIC",
            "ORG Spec": r["ORG Spec"], "Output": r["Output"], "Order Form": "—",
            "Status": "STATIC", "Notes": r["Notes"],
        })
    for r in variable_rows:
        org_text = r["org_reference_block"]["text"] if r.get("org_reference_block") else "Not mapped"
        rows.append({
            "PDF PAGE": r["PDF PAGE"], "Element / Field": r["field"], "Type": "VARIABLE",
            "ORG Spec": org_text, "Output": r["actual"], "Order Form": r["expected"],
            "Status": r["status"], "Notes": r["difference"],
        })
    for r in unaccounted_rows:
        rows.append({
            "PDF PAGE": r["PDF PAGE"], "Element / Field": "ORG Element", "Type": "MISSING / UNACCOUNTED",
            "ORG Spec": r["expected"], "Output": r["actual"], "Order Form": "—",
            "Status": "MISSING / UNACCOUNTED", "Notes": r["difference"],
        })
    return pd.DataFrame(rows)


def _metrics(static_rows, variable_rows, unaccounted_rows):
    static_count = len(static_rows)
    variable_count = sum(1 for r in variable_rows if r.get("status") in {"PASS", "FAIL", "REVIEW"})
    fail_count = sum(1 for r in variable_rows if r.get("status") == "FAIL")
    review_count = sum(1 for r in variable_rows if r.get("status") == "REVIEW")
    unaccounted = len(unaccounted_rows)
    overall = "FAIL" if fail_count or unaccounted else ("REVIEW" if review_count else "PASS")
    return overall, static_count, variable_count, fail_count, review_count, unaccounted


def _render_issue_list(rows, title):
    st.markdown(f"<div class='t3-section-title'>{html.escape(title)}</div>", unsafe_allow_html=True)
    if not rows:
        st.markdown("<div class='t3-empty'>No items in this section.</div>", unsafe_allow_html=True)
        return
    for row in rows:
        field = row.get("field") or row.get("Element / Field") or "Item"
        status = row.get("status") or row.get("Status") or "INFO"
        page = row.get("PDF PAGE", "—")
        expected = row.get("expected", row.get("ORG Spec", "—"))
        actual = row.get("actual", row.get("Output", "—"))
        reason = row.get("difference", row.get("Notes", "—"))
        st.markdown(
            f"<div class='t3-issue'><div class='t3-issue-head'><div class='t3-issue-field'>{html.escape(str(field))} • Page {page}</div>{_status_badge(status)}</div>"
            f"<div class='t3-issue-detail'><b>Reference / ORG:</b> {html.escape(str(expected))}<br><b>Output:</b> {html.escape(str(actual))}<br><b>Finding:</b> {html.escape(str(reason))}</div></div>",
            unsafe_allow_html=True,
        )


def _build_excel_report(summary, comparison_df, annotated_images, registration_rows, product_type):
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    dark = PatternFill("solid", fgColor="12233A")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="263A53")

    ws["A1"] = "ORG SPEC + ORDER FORM + OUTPUT QC"
    ws["A1"].font = Font(size=18, bold=True, color="FFFFFF")
    ws["A2"] = "Product Type"; ws["B2"] = product_type
    labels = ["Overall Result", "Static Elements", "Variable Elements", "Issues", "Manual Review", "Unaccounted ORG Elements"]
    for i, value in enumerate(summary, start=3):
        ws.cell(i, 1, labels[i - 3]); ws.cell(i, 2, value)

    for cell in ws[1]:
        cell.fill = dark
        cell.font = header_font

    detail = wb.create_sheet("Field Comparison")
    if not comparison_df.empty:
        for c, name in enumerate(comparison_df.columns, start=1):
            cell = detail.cell(1, c, name); cell.fill = dark; cell.font = header_font; cell.border = Border(bottom=thin)
        for r, vals in enumerate(comparison_df.itertuples(index=False), start=2):
            for c, value in enumerate(vals, start=1):
                detail.cell(r, c, value)
        status_col = list(comparison_df.columns).index("Status") + 1
        fills = {
            "PASS": PatternFill("solid", fgColor="198754"),
            "FAIL": PatternFill("solid", fgColor="DC3545"),
            "STATIC": PatternFill("solid", fgColor="C79B00"),
            "REVIEW": PatternFill("solid", fgColor="E67E22"),
            "MISSING / UNACCOUNTED": PatternFill("solid", fgColor="7E22CE"),
        }
        for r in range(2, detail.max_row + 1):
            status = str(detail.cell(r, status_col).value or "")
            if status in fills:
                detail.cell(r, status_col).fill = fills[status]

    reg = wb.create_sheet("ORG Registration")
    reg_headers = ["Page", "Method", "Scale", "Scale X", "Scale Y", "Aspect Delta", "ORG Boundary", "Output Boundary", "Warning"]
    for c, name in enumerate(reg_headers, start=1):
        cell = reg.cell(1, c, name); cell.fill = dark; cell.font = header_font
    for r, item in enumerate(registration_rows, start=2):
        for c, value in enumerate(item, start=1):
            reg.cell(r, c, value)

    visual = wb.create_sheet("Artwork Visual Validation")
    visual["A1"] = "Artwork Visual Validation"
    visual["A1"].font = Font(size=16, bold=True)
    visual["A3"] = "Yellow = Static | Field colors = variable fields | Red = Fail | Orange = Review | Purple = Unaccounted"
    row_cursor = 5
    for page_no, image in annotated_images.items():
        visual.cell(row_cursor, 1, f"Page {page_no}"); row_cursor += 1
        if image is not None:
            buf = io.BytesIO(); image.save(buf, format="PNG"); buf.seek(0)
            pic = XLImage(buf)
            pic.width = min(560, image.width)
            pic.height = int(image.height * (pic.width / image.width))
            visual.add_image(pic, f"A{row_cursor}")
            row_cursor += max(30, int(pic.height / 14) + 2)

    for sheet in wb.worksheets:
        for col in sheet.columns:
            letter = col[0].column_letter
            if sheet == visual and letter == "A":
                sheet.column_dimensions[letter].width = 25
            else:
                max_len = max(min(60, len(str(cell.value or ""))) for cell in col)
                sheet.column_dimensions[letter].width = max(12, max_len + 2)
        for row in sheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)

    out = io.BytesIO(); wb.save(out); return out.getvalue()


# ============================================================================
# INTERACTIVE VIEWER / SUMMARY UI
# ============================================================================

def _page_registration_note(reg):
    if not reg:
        return "Registration unavailable."
    msg = f"Registered using artwork boundary • scale {reg['scale']:.3f} • X {reg['tx']:.1f} • Y {reg['ty']:.1f}"
    if reg.get("warning"):
        msg += f" • {reg['warning']}"
    return msg


def _render_right_rail(overall, static_count, variable_count, fail_count, review_count, unaccounted):
    st.markdown("<div class='t3-rail-card' style='padding:12px'>", unsafe_allow_html=True)
    status_html = _status_badge(overall)
    st.markdown(f"<div class='t3-card-title'>Result Summary</div><div style='margin:8px 0 12px'>{status_html}</div>", unsafe_allow_html=True)
    for label, value, dot in [
        ("Static Elements", static_count, "#facc15"),
        ("Variable Elements", variable_count, "#3b82f6"),
        ("Issues Found", fail_count, "#ef4444"),
        ("Manual Review", review_count, "#f97316"),
        ("Unaccounted", unaccounted, "#a855f7"),
    ]:
        st.markdown(
            f"<div style='display:flex;justify-content:space-between;align-items:center;padding:7px 0;border-bottom:1px solid #18263a;font-size:10px;color:#a4b4c8'><span><span style='display:inline-block;width:8px;height:8px;border-radius:50%;background:{dot};margin-right:7px'></span>{label}</span><b style='color:#f5f8fc'>{value}</b></div>",
            unsafe_allow_html=True,
        )
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='t3-rail-card' style='padding:12px;margin-top:9px'><div class='t3-card-title'>Highlight Legend</div>", unsafe_allow_html=True)
    legend = [
        ("#facc15", "STATIC"), ("#3b82f6", "VARIABLE / PASS"),
        ("#ef4444", "FAIL"), ("#f97316", "REVIEW"), ("#a855f7", "UNACCOUNTED"),
    ]
    for color, label in legend:
        st.markdown(f"<div style='font-size:9px;color:#9aacbf;margin-top:7px'><span style='display:inline-block;width:10px;height:10px;background:{color};border-radius:2px;margin-right:6px;vertical-align:-1px'></span>{label}</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


def _page_structural_stats(baseline, output_page, variable_rows):
    reg = baseline.get("registration")
    return {
        "ORG blocks": len(baseline.get("registered_org_blocks", [])),
        "Output blocks": len(baseline.get("output_blocks", [])),
        "Static": len(baseline.get("static_matches", [])),
        "Variable checks": len(variable_rows),
        "Scale": round(reg["scale"], 3) if reg else None,
        "Aspect delta": round(reg["aspect_delta"], 3) if reg else None,
        "Registration": _page_registration_note(reg),
    }


def _render_findings_tab(page_no, static_rows, variable_rows, unaccounted_rows):
    page_static = [r for r in static_rows if int(r["PDF PAGE"]) == page_no]
    page_var = [r for r in variable_rows if int(r["PDF PAGE"]) == page_no]
    page_miss = [r for r in unaccounted_rows if int(r["PDF PAGE"]) == page_no]

    findings = []
    for r in page_static:
        findings.append({"PDF PAGE": page_no, "Element / Field": "Static Element", "Type": "STATIC", "ORG Spec": r["ORG Spec"], "Output": r["Output"], "Order Form": "—", "Status": "STATIC", "Notes": r["Notes"]})
    for r in page_var:
        org = r["org_reference_block"]["text"] if r.get("org_reference_block") else "Not mapped"
        findings.append({"PDF PAGE": page_no, "Element / Field": r["field"], "Type": "VARIABLE", "ORG Spec": org, "Output": r["actual"], "Order Form": r["expected"], "Status": r["status"], "Notes": r["difference"]})
    for r in page_miss:
        findings.append({"PDF PAGE": page_no, "Element / Field": "ORG Element", "Type": "MISSING / UNACCOUNTED", "ORG Spec": r["expected"], "Output": "Not found", "Order Form": "—", "Status": "MISSING / UNACCOUNTED", "Notes": r["difference"]})

    if not findings:
        st.info("No classified findings for this page.")
        return
    st.dataframe(
        pd.DataFrame(findings),
        width="stretch",
        hide_index=True,
        height=min(480, 120 + 36 * len(findings)),
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
        key=f"t3_findings_{page_no}",
    )


def _render_details(page_no, org_pages, output_pages, baselines, variable_rows, unaccounted_rows):
    idx = page_no - 1
    if idx < 0 or idx >= len(output_pages):
        st.info("Page is unavailable.")
        return
    baseline = baselines[idx]
    stats = _page_structural_stats(baseline, output_pages[idx], [r for r in variable_rows if r["PDF PAGE"] == page_no])
    cols = st.columns(6)
    for col, (label, value) in zip(cols, stats.items()):
        with col:
            if label == "Registration":
                st.markdown(f"<div class='t3-mini-stat'><b style='font-size:11px'>Registered</b><span>{html.escape(str(value))}</span></div>", unsafe_allow_html=True)
            else:
                st.markdown(f"<div class='t3-mini-stat'><b>{html.escape(str(value))}</b><span>{html.escape(label)}</span></div>", unsafe_allow_html=True)

    st.markdown("<div class='t3-section-title'>ORG Registration</div>", unsafe_allow_html=True)
    st.info(stats["Registration"])

    # Full ORG baseline remains inspectable. This is intentionally separate
    # from the main comparison table so the user can study what the engine
    # extracted from ORG before looking at the Output decision.
    baseline_rows = []
    static_org_indexes = {m["org_index"] for m in baseline["static_matches"]}
    static_by_org = {m["org_index"]: m for m in baseline["static_matches"]}
    out_blocks = baseline.get("output_blocks", [])
    page_w = max(1, int(output_pages[idx].get("image_width", 1)))
    page_h = max(1, int(output_pages[idx].get("image_height", 1)))
    for oi, ob in enumerate(baseline.get("registered_org_blocks", [])):
        if oi in static_org_indexes:
            m = static_by_org[oi]
            classification = "STATIC"
            mapped = m["output"]["text"]
            note = f"Mapped at registered position; confidence {m['score']:.2f}."
        else:
            mapped_block = None
            mapped_score = 999.0
            for out_block in out_blocks:
                met = _bbox_metrics(ob, out_block, page_w, page_h)
                score = met["center_diag"] + abs(math.log(max(.01, met["width_ratio"]))) * 0.12
                if score < mapped_score:
                    mapped_score = score
                    mapped_block = out_block
            if mapped_block is not None and mapped_score <= 0.065:
                classification = "VARIABLE / CANDIDATE"
                mapped = mapped_block["text"]
                note = "Same registered region has Output content; likely variable or presentation change."
            else:
                classification = "MISSING / UNACCOUNTED"
                mapped = "Not found"
                note = "No reasonable Output counterpart at the registered position."
        baseline_rows.append({"ORG #": oi + 1, "ORG Text": ob["text"], "Classification": classification, "Output Counterpart": mapped, "Notes": note})
    if baseline_rows:
        st.dataframe(pd.DataFrame(baseline_rows), width="stretch", hide_index=True, height=min(480, 140 + 32 * len(baseline_rows)), key=f"t3_org_baseline_{page_no}")

    page_static = [m for m in baseline["static_matches"]]
    if page_static:
        st.markdown("<div class='t3-section-title'>Static Elements</div>", unsafe_allow_html=True)
        for m in page_static:
            met = _bbox_metrics(m["org"], m["output"], max(1, int(output_pages[idx].get("image_width", 1))), max(1, int(output_pages[idx].get("image_height", 1))))
            st.markdown(
                f"<div class='t3-issue'><div class='t3-issue-head'><div class='t3-issue-field'>{html.escape(m['org']['text'])}</div>{_status_badge('STATIC')}</div><div class='t3-issue-detail'>Position Δ X {met['dx']:.3f} • Y {met['dy']:.3f} • Scale {met['scale_score']:.2f} • Confidence {m['score']:.2f}</div></div>",
                unsafe_allow_html=True,
            )

    page_var = [r for r in variable_rows if r["PDF PAGE"] == page_no]
    if page_var:
        st.markdown("<div class='t3-section-title'>Variable Field Detail</div>", unsafe_allow_html=True)
        for r in page_var:
            org_text = r["org_reference_block"]["text"] if r.get("org_reference_block") else "Not mapped"
            st.markdown(
                f"<div class='t3-issue'><div class='t3-issue-head'><div class='t3-issue-field'>{html.escape(str(r['field']))}</div>{_status_badge(r['status'])}</div>"
                f"<div class='t3-issue-detail'><b>Type:</b> {html.escape(str(r['field_type']))}<br><b>ORG:</b> {html.escape(org_text)}<br><b>Order Form:</b> {html.escape(str(r['expected']))}<br><b>Output:</b> {html.escape(str(r['actual']))}<br><b>Presentation:</b> {html.escape(str(r['presentation_status']))}<br><b>Finding:</b> {html.escape(str(r['difference']))}</div></div>",
                unsafe_allow_html=True,
            )

    page_miss = [r for r in unaccounted_rows if r["PDF PAGE"] == page_no]
    _render_issue_list(page_miss, "Unaccounted ORG Elements")


def _render_summary(static_rows, variable_rows, unaccounted_rows, product_type, baselines, output_pages):
    overall, static_count, variable_count, fail_count, review_count, unaccounted_count = _metrics(static_rows, variable_rows, unaccounted_rows)
    st.markdown(f"<div class='t3-card' style='padding:14px'><div class='t3-card-title'>Final QC Result</div><div style='margin-top:7px'>{_status_badge(overall)}</div><div class='t3-note' style='margin-top:9px'>Product: <b>{html.escape(product_type)}</b> • ORG is the presentation/spec baseline • Order Form is the variable-data source.</div></div>", unsafe_allow_html=True)

    cols = st.columns(5)
    for col, (n, label) in zip(cols, [(static_count, "Static"), (variable_count, "Variable"), (fail_count, "Issues"), (review_count, "Review"), (unaccounted_count, "Unaccounted")]):
        with col:
            st.markdown(f"<div class='t3-mini-stat'><b>{n}</b><span>{label}</span></div>", unsafe_allow_html=True)

    st.markdown("<div class='t3-section-title'>QC Interpretation</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='t3-card' style='padding:12px'><div class='t3-note'>"
        "<b>Static</b> = ORG and Output text match after artwork registration and are within the allowed positional tolerance.<br><br>"
        "<b>Variable</b> = Output data is validated through the existing Order Form → Output engine. For HTL / Other, ORG presentation remains the reference for case and punctuation.<br><br>"
        "<b>Review</b> = the system has evidence but cannot safely make a fully deterministic decision.<br><br>"
        "<b>Missing / Unaccounted</b> = a meaningful ORG element has no reasonable Output counterpart. Ordinary Tool-1 NOT FOUND fields are intentionally omitted from the main comparison."
        "</div></div>",
        unsafe_allow_html=True,
    )


# ============================================================================
# MAIN
# ============================================================================

def main():
    _tool3_css()

    # Session state -----------------------------------------------------------
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
    if "t3_view_mode" not in st.session_state:
        st.session_state["t3_view_mode"] = "Side by Side"

    if st.session_state.get("t3_version") != TOOL3_VERSION:
        st.session_state["t3_result"] = None
        st.session_state["t3_selected_fields"] = None
        st.session_state["t3_active_section"] = "Comparison View"
        st.session_state["t3_version"] = TOOL3_VERSION

    reset_id = st.session_state["t3_reset_id"]

    # Header ------------------------------------------------------------------
    st.markdown(
        """
        <div class="t3-header">
          <div class="t3-logo">✓</div>
          <div class="t3-rule"></div>
          <div>
            <div class="t3-title">ORG Spec + Order Form + Output Check</div>
            <div class="t3-subtitle">Verify artwork accuracy, variable data, presentation, positioning and layout structure</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    nav_col, workspace_col, rail_col = st.columns([0.85, 4.55, 1.35], gap="small")

    # Left navigation ---------------------------------------------------------
    with nav_col:
        st.markdown('<div class="t3-side">', unsafe_allow_html=True)
        active = st.session_state["t3_active_section"]
        if st.button("⌂  Comparison", width="stretch", type="primary" if active == "Comparison View" else "secondary", key=f"t3_nav_cmp_{reset_id}"):
            st.session_state["t3_active_section"] = "Comparison View"; st.rerun()
        if st.button("＋  New Check", width="stretch", key=f"t3_nav_new_{reset_id}"):
            st.session_state["t3_reset_id"] += 1
            st.session_state["t3_result"] = None
            st.session_state["t3_selected_fields"] = None
            st.session_state["t3_active_section"] = "Comparison View"
            st.rerun()
        if st.button("↶  History", width="stretch", key=f"t3_nav_hist_{reset_id}"):
            st.session_state["t3_active_section"] = "History"; st.rerun()
        if st.button("⚙  Settings", width="stretch", key=f"t3_nav_set_{reset_id}"):
            st.session_state["t3_active_section"] = "Settings"; st.rerun()
        st.markdown(
            "<div class='t3-side-note'>Tool 3 reads the ORG Spec first. It then uses the Order Form as the variable-data source and Output as the final artwork to validate.</div></div>",
            unsafe_allow_html=True,
        )

    with workspace_col:
        # Upload area ---------------------------------------------------------
        upload_cols = st.columns(4, gap="small")
        with upload_cols[0]:
            st.markdown("<div class='t3-upload-card'><div class='t3-card-title'>1. Order Form</div><div class='t3-card-sub'>Excel / CSV</div>", unsafe_allow_html=True)
            order_file = st.file_uploader("Order Form", type=["xlsx", "xls", "csv"], label_visibility="collapsed", key=f"t3_order_{reset_id}")
            if order_file:
                st.markdown(f"<div class='t3-fileline'>{html.escape(order_file.name)} <span class='t3-ok'>✓</span></div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
        with upload_cols[1]:
            st.markdown("<div class='t3-upload-card'><div class='t3-card-title'>2. ORG Spec</div><div class='t3-card-sub'>PDF / Image • analyzed first</div>", unsafe_allow_html=True)
            org_file = st.file_uploader("ORG Spec", type=["pdf", "jpg", "jpeg", "png"], label_visibility="collapsed", key=f"t3_org_{reset_id}")
            if org_file:
                st.markdown(f"<div class='t3-fileline'>{html.escape(org_file.name)} <span class='t3-ok'>✓</span></div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
        with upload_cols[2]:
            st.markdown("<div class='t3-upload-card'><div class='t3-card-title'>3. Output</div><div class='t3-card-sub'>PDF / Image</div>", unsafe_allow_html=True)
            output_file = st.file_uploader("Output", type=["pdf", "jpg", "jpeg", "png"], label_visibility="collapsed", key=f"t3_output_{reset_id}")
            if output_file:
                st.markdown(f"<div class='t3-fileline'>{html.escape(output_file.name)} <span class='t3-ok'>✓</span></div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
        with upload_cols[3]:
            st.markdown("<div class='t3-upload-card'><div class='t3-card-title'>Product Type</div><div class='t3-card-sub'>Controls case behavior</div>", unsafe_allow_html=True)
            product_type = st.selectbox("Product Type", ["----- SELECT -----", "PFL", "HTL", "Other"], key=f"t3_product_{reset_id}")
            ready = bool(order_file and org_file and output_file and product_type != "----- SELECT -----")
            run_clicked = st.button("▶  Analyze & Run QC", width="stretch", type="primary", disabled=not ready, key=f"t3_run_{reset_id}")
            st.markdown("</div>", unsafe_allow_html=True)

        # New run -------------------------------------------------------------
        if run_clicked:
            try:
                with st.spinner("Analyzing ORG Spec first, then registering Output and applying Order Form validation…"):
                    if str(getattr(order_file, "name", "")).casefold().endswith(".csv"):
                        df = pd.read_csv(order_file)
                    else:
                        df = load_excel(order_file)
                    if df.empty:
                        raise ValueError("The Order Form contains no usable data rows.")

                    org_pages = extract_output_pages(org_file)
                    output_pages = extract_output_pages(output_file)
                    if not org_pages:
                        raise ValueError("The ORG Spec could not be read.")
                    if not output_pages:
                        raise ValueError("The Output artwork could not be read.")

                    fields = [f for f in get_available_fields(df) if not is_admin_field(f)]
                    if not fields:
                        raise ValueError("No usable Order Form fields were found.")

                    mapping = {int(p.get("page", i + 1)): min(i, len(df) - 1) for i, p in enumerate(output_pages)}
                    baselines = []
                    page_limit = max(len(org_pages), len(output_pages))
                    for i in range(page_limit):
                        op = org_pages[i] if i < len(org_pages) else None
                        qp = output_pages[i] if i < len(output_pages) else None
                        baselines.append(_page_baseline(op, qp) if op and qp else {"registration": None, "static_matches": [], "registered_org_blocks": [], "output_blocks": []})

                    detected = _dynamic_auto_detect(df, output_pages, product_type, mapping)
                    variable_results = _run_shared_field_checks(df, output_pages, org_pages, detected, product_type, mapping)
                    unaccounted_rows = _unaccounted_org_rows(org_pages, output_pages, baselines)
                    static_rows = _build_static_rows(baselines, output_pages)

                    # Store a compact but complete source/result snapshot.
                    st.session_state["t3_selected_fields"] = detected
                    st.session_state["t3_result"] = {
                        "df": df,
                        "org_pages": org_pages,
                        "output_pages": output_pages,
                        "mapping": mapping,
                        "all_fields": fields,
                        "detected_fields": detected,
                        "product_type": product_type,
                        "baselines": baselines,
                        "static_rows": static_rows,
                        "variable_results": variable_results,
                        "unaccounted_rows": unaccounted_rows,
                        "files": {
                            "order": getattr(order_file, "name", "Order Form"),
                            "org": getattr(org_file, "name", "ORG Spec"),
                            "output": getattr(output_file, "name", "Output"),
                        },
                    }
                    st.session_state["t3_history"].append({
                        **st.session_state["t3_result"]["files"],
                        "product": product_type,
                    })
                    st.session_state["t3_active_section"] = "Comparison View"
                st.success("ORG baseline analysis and QC comparison completed.")
                st.rerun()
            except Exception as exc:
                st.error(f"QC could not be started: {exc}")

        # History / settings --------------------------------------------------
        active = st.session_state["t3_active_section"]
        if active == "History":
            st.markdown("<div class='t3-card' style='padding:13px'><div class='t3-card-title'>Check History</div><div class='t3-note' style='margin-top:4px'>Current Streamlit session only.</div></div>", unsafe_allow_html=True)
            history = st.session_state.get("t3_history", [])
            if not history:
                st.markdown("<div class='t3-empty'>No checks have been run in this session.</div>", unsafe_allow_html=True)
            else:
                for i, item in enumerate(reversed(history[-15:]), 1):
                    st.markdown(
                        f"<div class='t3-issue'><div class='t3-issue-head'><div class='t3-issue-field'>Check {i} • {html.escape(str(item.get('product')))}</div>{_status_badge('PASS')}</div>"
                        f"<div class='t3-issue-detail'>Order Form: {html.escape(str(item.get('order')))}<br>ORG: {html.escape(str(item.get('org')))}<br>Output: {html.escape(str(item.get('output')))}</div></div>",
                        unsafe_allow_html=True,
                    )
        elif active == "Settings":
            st.markdown("<div class='t3-card' style='padding:13px'><div class='t3-card-title'>Viewer / QC Settings</div><div class='t3-note' style='margin-top:4px'>Presentation controls only. Core Order Form validation is unchanged.</div></div>", unsafe_allow_html=True)
            st.checkbox("Show registration diagnostics in Details", value=True, key=f"t3_setting_reg_{reset_id}")
            st.checkbox("Show ORG reference text on variable findings", value=True, key=f"t3_setting_org_{reset_id}")
            st.info("Registration uses proportional scaling and artwork/content boundary anchoring. Future QC rejection scenarios can be added without replacing the shared Tool 1 comparison engine.")
        else:
            result = st.session_state.get("t3_result")
            if not result:
                st.markdown("<div class='t3-card' style='padding:16px;margin-top:10px'><div class='t3-card-title'>Ready for QC</div><div class='t3-empty'>Upload Order Form, ORG Spec and Output, select Product Type, then click <b>Analyze & Run QC</b>.</div></div>", unsafe_allow_html=True)
            else:
                df = result["df"]
                org_pages = result["org_pages"]
                output_pages = result["output_pages"]
                product_type = result["product_type"]
                mapping = result["mapping"]

                # Page mapping / field selection --------------------------------
                controls = st.columns([1.15, 2.75, 1.05], gap="small")
                with controls[0]:
                    with st.expander("🧭 Page Mapping", expanded=len(output_pages) > 1):
                        st.caption("Default: page 1 → Excel row 2, page 2 → Excel row 3. Override as needed.")
                        opts = list(range(len(df)))
                        labels = [f"Excel Row {i + 2}" for i in opts]
                        for i, page in enumerate(output_pages):
                            page_no = int(page.get("page", i + 1))
                            old = int(mapping.get(page_no, min(i, len(df) - 1)))
                            mapping[page_no] = st.selectbox(
                                f"PDF Page {page_no}", opts,
                                index=max(0, min(old, len(df) - 1)),
                                format_func=lambda x, labels=labels: labels[x],
                                key=f"t3_map_{reset_id}_{page_no}",
                            )
                        result["mapping"] = mapping

                with controls[1]:
                    available = result["all_fields"]
                    defaults = [f for f in (st.session_state.get("t3_selected_fields") or result.get("detected_fields", [])) if f in available]
                    selected_fields = st.multiselect(
                        "Field Comparison",
                        available,
                        default=defaults,
                        key=f"t3_fields_{reset_id}",
                        help="Type inside the field selector to search. Blank Order Form values are automatically excluded.",
                    )
                    st.session_state["t3_selected_fields"] = selected_fields

                with controls[2]:
                    st.markdown("<div class='t3-control-label'>Product</div>", unsafe_allow_html=True)
                    st.markdown(f"<div style='font-size:16px;font-weight:850;color:#f6f9fc'>{html.escape(product_type)}</div>", unsafe_allow_html=True)
                    if st.button("Auto Detect", width="stretch", key=f"t3_auto_{reset_id}"):
                        detected = _dynamic_auto_detect(df, output_pages, product_type, mapping)
                        result["detected_fields"] = detected
                        st.session_state["t3_selected_fields"] = detected
                        st.rerun()

                # Recalculate once and reuse across every tab -----------------
                baselines = []
                for i in range(max(len(org_pages), len(output_pages))):
                    op = org_pages[i] if i < len(org_pages) else None
                    qp = output_pages[i] if i < len(output_pages) else None
                    baselines.append(_page_baseline(op, qp) if op and qp else {"registration": None, "static_matches": [], "registered_org_blocks": [], "output_blocks": []})
                variable_results = _run_shared_field_checks(df, output_pages, org_pages, selected_fields, product_type, mapping)
                unaccounted_rows = _unaccounted_org_rows(org_pages, output_pages, baselines)
                static_rows = _build_static_rows(baselines, output_pages)
                result["baselines"] = baselines
                result["variable_results"] = variable_results
                result["unaccounted_rows"] = unaccounted_rows
                result["static_rows"] = static_rows

                overall, static_count, variable_count, fail_count, review_count, unaccounted_count = _metrics(static_rows, variable_results, unaccounted_rows)

                # Top-level application tabs -----------------------------------
                tabs = st.tabs(["👁  Comparison View", "☷  Details", "▥  Summary"])

                # Comparison ---------------------------------------------------
                with tabs[0]:
                    c1, c2 = st.columns([1.0, 2.0])
                    with c1:
                        page_options = [int(p.get("page", i + 1)) for i, p in enumerate(output_pages)]
                        selected_page = st.selectbox("Output Page", page_options, key=f"t3_view_page_{reset_id}")
                    with c2:
                        st.markdown("<div class='t3-view-note'>Viewer modes are now fully interactive. Overlay uses the registered ORG coordinate system; Blink alternates the same registered artwork space.</div>", unsafe_allow_html=True)

                    page_idx = [int(p.get("page", i + 1)) for i, p in enumerate(output_pages)].index(selected_page)
                    out_page = output_pages[page_idx]
                    org_page = org_pages[page_idx] if page_idx < len(org_pages) else None
                    baseline = baselines[page_idx]
                    page_static_rows = [r for r in static_rows if int(r["PDF PAGE"]) == selected_page]
                    page_var_rows = [r for r in variable_results if int(r["PDF PAGE"]) == selected_page]
                    page_miss_rows = [r for r in unaccounted_rows if int(r["PDF PAGE"]) == selected_page]

                    annotated = _build_annotated_output(out_page, baseline["static_matches"], page_var_rows)
                    registered_org = _build_registered_org_image(baseline["registration"]) if baseline.get("registration") else None
                    raw_org = Image.open(io.BytesIO(org_page["image_bytes"])).convert("RGB") if org_page and org_page.get("image_bytes") else None

                    st.components.v1.html(
                        _viewer_html(raw_org, registered_org, annotated),
                        height=700,
                        scrolling=False,
                    )

                    if baseline.get("registration"):
                        st.caption(_page_registration_note(baseline["registration"]))
                    if len(org_pages) != len(output_pages):
                        st.warning(f"ORG pages: {len(org_pages)} • Output pages: {len(output_pages)}. Registration is applied where a same-index page exists.")

                    lower = st.tabs(["Findings", "Field Comparison", "Issues"])
                    with lower[0]:
                        _render_findings_tab(selected_page, static_rows, variable_results, unaccounted_rows)
                    with lower[1]:
                        comp_df = _build_comparison_df(static_rows, variable_results, unaccounted_rows)
                        if comp_df.empty:
                            st.info("No comparison entries were generated.")
                        else:
                            f1, f2, f3 = st.columns(3)
                            pages = ["All"] + sorted({str(x) for x in comp_df["PDF PAGE"].unique()})
                            statuses = ["All"] + sorted({str(x) for x in comp_df["Status"].unique()})
                            types = ["All"] + sorted({str(x) for x in comp_df["Type"].unique()})
                            with f1: pf = st.selectbox("Page", pages, key=f"t3_comp_page_{reset_id}")
                            with f2: sf = st.selectbox("Status", statuses, key=f"t3_comp_status_{reset_id}")
                            with f3: tf = st.selectbox("Type", types, key=f"t3_comp_type_{reset_id}")
                            filtered = comp_df.copy()
                            if pf != "All": filtered = filtered[filtered["PDF PAGE"].astype(str) == pf]
                            if sf != "All": filtered = filtered[filtered["Status"].astype(str) == sf]
                            if tf != "All": filtered = filtered[filtered["Type"].astype(str) == tf]
                            st.dataframe(filtered, width="stretch", hide_index=True, height=min(520, 145 + 36 * len(filtered)), key=f"t3_comp_df_{reset_id}")
                            st.markdown("<div class='t3-table-note'>Ordinary Tool-1 NOT FOUND entries are intentionally omitted. Only classified static, variable, review, fail or genuine ORG-unaccounted findings are shown.</div>", unsafe_allow_html=True)
                    with lower[2]:
                        _render_issue_list([r for r in variable_results if r.get("status") in {"FAIL", "REVIEW"}], "Variable Issues / Reviews")
                        _render_issue_list(unaccounted_rows, "Unaccounted ORG Elements")

                # Details ------------------------------------------------------
                with tabs[1]:
                    page_options = [int(p.get("page", i + 1)) for i, p in enumerate(output_pages)]
                    detail_page = st.selectbox("Page to inspect", page_options, key=f"t3_detail_page_{reset_id}")
                    _render_details(detail_page, org_pages, output_pages, baselines, variable_results, unaccounted_rows)

                # Summary ------------------------------------------------------
                with tabs[2]:
                    _render_summary(static_rows, variable_results, unaccounted_rows, product_type, baselines, output_pages)
                    st.markdown("<div class='t3-section-title'>Downloadable QC Report</div>", unsafe_allow_html=True)
                    annotated_images = {}
                    registration_rows = []
                    for i, out_page in enumerate(output_pages):
                        page_no = int(out_page.get("page", i + 1))
                        baseline = baselines[i] if i < len(baselines) else {"registration": None, "static_matches": []}
                        page_vars = [r for r in variable_results if int(r["PDF PAGE"]) == page_no]
                        annotated_images[page_no] = _build_annotated_output(out_page, baseline["static_matches"], page_vars)
                        reg = baseline.get("registration")
                        if reg:
                            registration_rows.append([
                                page_no, reg["method"], round(reg["scale"], 4), round(reg["scale_x"], 4), round(reg["scale_y"], 4), round(reg["aspect_delta"], 4),
                                str(tuple(round(x, 1) for x in reg["org_bbox"])), str(tuple(round(x, 1) for x in reg["output_bbox"])), reg["warning"],
                            ])
                        else:
                            registration_rows.append([page_no, "Unavailable", "", "", "", "", "", "", ""])
                    comp_df = _build_comparison_df(static_rows, variable_results, unaccounted_rows)
                    report_bytes = _build_excel_report(
                        (overall, static_count, variable_count, fail_count, review_count, unaccounted_count),
                        comp_df, annotated_images, registration_rows, product_type,
                    )
                    st.download_button(
                        "⬇  Download QC Excel Report",
                        data=report_bytes,
                        file_name="ORG_OrderForm_Output_QC_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        width="stretch",
                        key=f"t3_download_{reset_id}",
                    )

    # Right rail --------------------------------------------------------------
    with rail_col:
        result = st.session_state.get("t3_result")
        if result and active not in {"Settings", "History"}:
            _render_right_rail(*_metrics(result.get("static_rows", []), result.get("variable_results", []), result.get("unaccounted_rows", [])))
            st.markdown("<div class='t3-rail-card' style='padding:12px;margin-top:9px'><div class='t3-card-title'>QC Baseline</div><div class='t3-note' style='margin-top:7px'>ORG is always analyzed first. Variable data is then checked through the shared Order Form → Output engine. HTL / Other also validates ORG presentation case and punctuation.</div></div>", unsafe_allow_html=True)
        else:
            st.markdown("<div class='t3-rail-card' style='padding:12px'><div class='t3-card-title'>Result Summary</div><div class='t3-empty'>Run a check to populate this panel.</div></div>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
