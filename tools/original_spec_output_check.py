import streamlit as st
import fitz
from PIL import Image
import io
import base64
import streamlit.components.v1 as components


# =========================================================
# SESSION STATE
# =========================================================

if "spec_reset_id" not in st.session_state:
    st.session_state["spec_reset_id"] = 0

if "spec_compare_started" not in st.session_state:
    st.session_state["spec_compare_started"] = False


# =========================================================
# PDF HELPERS
# =========================================================

def get_pdf_page_count(uploaded_file):
    pdf_bytes = uploaded_file.getvalue()

    doc = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    count = len(doc)
    doc.close()

    return count


def pdf_page_to_image(uploaded_file, page_number=0, scale=2.0):

    pdf_bytes = uploaded_file.getvalue()

    doc = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    page_number = min(
        page_number,
        len(doc) - 1
    )

    page = doc.load_page(page_number)

    matrix = fitz.Matrix(scale, scale)

    pix = page.get_pixmap(
        matrix=matrix,
        alpha=False
    )

    image = Image.frombytes(
        "RGB",
        [pix.width, pix.height],
        pix.samples
    )

    doc.close()

    return image


# =========================================================
# IMAGE HELPERS
# =========================================================

def create_monochrome(image, color):

    gray = image.convert("L")

    width, height = gray.size

    result = Image.new(
        "RGBA",
        (width, height),
        (0, 0, 0, 0)
    )

    r, g, b = color

    gray_pixels = gray.load()
    result_pixels = result.load()

    for y in range(height):

        for x in range(width):

            value = gray_pixels[x, y]

            # Dark artwork = visible
            # White background = transparent
            alpha = 255 - value

            result_pixels[x, y] = (
                r,
                g,
                b,
                alpha
            )

    return result


def image_to_base64(image):

    buffer = io.BytesIO()

    image.save(
        buffer,
        format="PNG"
    )

    return base64.b64encode(
        buffer.getvalue()
    ).decode("utf-8")


# =========================================================
# COMPARISON VIEWER
# =========================================================

def comparison_viewer(
    original_image,
    output_image,
    mode="overlay",
    blink_speed=0.5
):

    # Dark Red
    original_mono = create_monochrome(
        original_image,
        (150, 40, 40)
    )

    # Dark Green
    output_mono = create_monochrome(
        output_image,
        (35, 110, 70)
    )

    original_b64 = image_to_base64(
        original_mono
    )

    output_b64 = image_to_base64(
        output_mono
    )

    # Use normal string replacement instead of f-string
    # This prevents Python SyntaxErrors from JavaScript braces.

    html = """
<!DOCTYPE html>
<html>

<head>

<style>

html, body {
    margin: 0;
    padding: 0;
    background: transparent;
    overflow: hidden;
}

* {
    box-sizing: border-box;
}

html, body {
    width: 100%;
    height: 100%;
}

#wrapper {
    width: 100%;
    height: 100%;
    display: flex;
    justify-content: center;
    align-items: center;
    padding: 0 8px;
}

#viewer {
    width: min(100%, 980px);
    height: 620px;
    background: #ffffff;
    border: 1px solid #333;
    border-radius: 12px;
    position: relative;
    overflow: hidden;
    cursor: grab;
    box-shadow: 0 10px 28px rgba(0,0,0,0.18);
}

#viewer:active {
    cursor: grabbing;
}

canvas {
    width: 100%;
    height: 100%;
    display: block;
}

#info {
    position: absolute;
    top: 12px;
    left: 12px;
    background: rgba(20,20,20,0.85);
    color: #ddd;
    padding: 7px 12px;
    border-radius: 6px;
    font-family: Arial;
    font-size: 12px;
    pointer-events: none;
}

#controls {
    position: absolute;
    bottom: 16px;
    left: 50%;
    transform: translateX(-50%);
    display: flex;
    gap: 8px;
    background: rgba(20,20,20,0.94);
    padding: 8px;
    border-radius: 10px;
    z-index: 20;
    box-shadow: 0 6px 18px rgba(0,0,0,0.28);
}

#controls button {
    background: #252525;
    color: white;
    border: 1px solid #555;
    min-width: 42px;
    padding: 7px 13px;
    border-radius: 7px;
    cursor: pointer;
    font-weight: 700;
    transition: transform 0.15s ease, background 0.15s ease;
}

#controls button:hover {
    background: #444;
    transform: translateY(-1px);
}

#controls button:active {
    transform: scale(0.96);
}

</style>

</head>

<body>

<div id="wrapper">

    <div id="viewer">

        <canvas id="canvas"></canvas>

        <div id="info">
            Loading...
        </div>

        <div id="controls">

            <button id="zoomOut">−</button>

            <button id="fit">FIT</button>

            <button id="zoomIn">+</button>

            <button id="reset">RESET</button>

        </div>

    </div>

</div>


<script>

// =========================================================
// SETTINGS
// =========================================================

const MODE = "__MODE__";

const BLINK_SPEED = __BLINK_SPEED__;
// 1.0 should feel like the previous 1.5 speed.
const BLINK_INTERVAL = 1000 / (BLINK_SPEED * 1.5);

const ORIGINAL_OPACITY = 0.80;

const OUTPUT_OPACITY = 0.80;


// =========================================================
// CANVAS
// =========================================================

const viewer = document.getElementById("viewer");

const canvas = document.getElementById("canvas");

const ctx = canvas.getContext("2d");

const info = document.getElementById("info");


// =========================================================
// IMAGES
// =========================================================

const original = new Image();

const output = new Image();

original.src =
    "data:image/png;base64,__ORIGINAL_IMAGE__";

output.src =
    "data:image/png;base64,__OUTPUT_IMAGE__";

let loaded = 0;


// =========================================================
// VIEW STATE
// =========================================================

let scale = 1;

let offsetX = 0;

let offsetY = 0;

let blinkShowOriginal = true;


// =========================================================
// RESIZE
// =========================================================

let hasInitialFit = false;
let resizeTimer = null;

function resizeCanvas(refit = false) {

    const rect = viewer.getBoundingClientRect();

    const newWidth = Math.max(1, Math.round(rect.width));
    const newHeight = Math.max(1, Math.round(rect.height));

    const changed =
        canvas.width !== newWidth ||
        canvas.height !== newHeight;

    if (!changed) return;

    canvas.width = newWidth;
    canvas.height = newHeight;

    if (refit && original.width && output.width) {
        fitImage();
    } else {
        draw();
    }
}


// =========================================================
// FIT IMAGE
// =========================================================

function getBaseSize() {

    return {
        width: Math.max(original.width, output.width),
        height: Math.max(original.height, output.height)
    };
}


function fitImage() {

    if (!original.width || !output.width) return;

    const padding = 44;
    const base = getBaseSize();

    const scaleX =
        (canvas.width - padding) / base.width;

    const scaleY =
        (canvas.height - padding) / base.height;

    scale = Math.min(scaleX, scaleY);

    offsetX =
        (canvas.width - base.width * scale) / 2;

    offsetY =
        (canvas.height - base.height * scale) / 2;

    draw();
}


function getDrawPosition(image) {

    const base = getBaseSize();

    return {
        x: offsetX + ((base.width - image.width) * scale) / 2,
        y: offsetY + ((base.height - image.height) * scale) / 2
    };
}


// =========================================================
// DRAW
// =========================================================

function draw() {

    ctx.clearRect(
        0,
        0,
        canvas.width,
        canvas.height
    );

    ctx.fillStyle = "#ffffff";

    ctx.fillRect(
        0,
        0,
        canvas.width,
        canvas.height
    );


    // OVERLAY MODE

    if (MODE === "overlay") {

        const originalPos = getDrawPosition(original);
        const outputPos = getDrawPosition(output);

        ctx.globalAlpha =
            ORIGINAL_OPACITY;

        ctx.drawImage(
            original,
            originalPos.x,
            originalPos.y,
            original.width * scale,
            original.height * scale
        );


        ctx.globalAlpha =
            OUTPUT_OPACITY;

        ctx.drawImage(
            output,
            outputPos.x,
            outputPos.y,
            output.width * scale,
            output.height * scale
        );

        info.innerText =
            "🟥 ORIGINAL + 🟢 OUTPUT | "
            + Math.round(scale * 100)
            + "%";
    }


    // BLINK MODE

    if (MODE === "blink") {

        ctx.globalAlpha = 1;

        const activeImage =
            blinkShowOriginal
                ? original
                : output;

        const activePos = getDrawPosition(activeImage);

        ctx.drawImage(
            activeImage,
            activePos.x,
            activePos.y,
            activeImage.width * scale,
            activeImage.height * scale
        );

        info.innerText =
            blinkShowOriginal
                ? "🟥 ORIGINAL SPEC"
                : "🟢 OUTPUT";

        info.innerText +=
            " | "
            + Math.round(scale * 100)
            + "%";
    }

    ctx.globalAlpha = 1;
}


// =========================================================
// IMAGE LOADING
// =========================================================

function onImageLoaded() {

    loaded++;

    if (loaded === 2) {

        // Give Streamlit's iframe/layout time to reach its final size.
        setTimeout(function() {

            resizeCanvas(false);
            fitImage();
            hasInitialFit = true;

            blinkShowOriginal = true;
            draw();

            if (MODE === "blink") {

                setInterval(function() {

                    blinkShowOriginal =
                        !blinkShowOriginal;

                    draw();

                }, BLINK_INTERVAL);
            }

        }, 350);
    }
}

original.onload = onImageLoaded;

output.onload = onImageLoaded;

// =========================================================
// ZOOM
// =========================================================

viewer.addEventListener(
    "wheel",

    function(event) {

        event.preventDefault();

        const factor =
            event.deltaY < 0
                ? 1.12
                : 0.88;

        const rect =
            canvas.getBoundingClientRect();

        const mouseX =
            event.clientX - rect.left;

        const mouseY =
            event.clientY - rect.top;

        offsetX =
            mouseX -
            (mouseX - offsetX)
            * factor;

        offsetY =
            mouseY -
            (mouseY - offsetY)
            * factor;

        scale *= factor;

        scale =
            Math.max(
                0.05,
                Math.min(scale, 20)
            );

        draw();

    },

    { passive: false }
);


// =========================================================
// PAN
// =========================================================

let dragging = false;

let lastX = 0;

let lastY = 0;


viewer.addEventListener(
    "mousedown",

    function(event) {

        dragging = true;

        lastX = event.clientX;

        lastY = event.clientY;

    }
);


window.addEventListener(
    "mousemove",

    function(event) {

        if (!dragging) return;

        const dx =
            event.clientX - lastX;

        const dy =
            event.clientY - lastY;

        offsetX += dx;

        offsetY += dy;

        lastX = event.clientX;

        lastY = event.clientY;

        draw();

    }
);


window.addEventListener(
    "mouseup",

    function() {

        dragging = false;

    }
);


// =========================================================
// BUTTONS
// =========================================================

document
    .getElementById("zoomIn")
    .addEventListener(
        "click",

        function() {

            scale *= 1.25;

            draw();

        }
    );


document
    .getElementById("zoomOut")
    .addEventListener(
        "click",

        function() {

            scale /= 1.25;

            draw();

        }
    );


document
    .getElementById("fit")
    .addEventListener(
        "click",

        fitImage
    );


document
    .getElementById("reset")
    .addEventListener(
        "click",

        fitImage
    );


// =========================================================
// DOUBLE CLICK
// =========================================================

viewer.addEventListener(
    "dblclick",

    function() {

        fitImage();

    }
);


// =========================================================
// START
// =========================================================

window.addEventListener(
    "resize",
    function() {
        resizeCanvas(hasInitialFit);
    }
);

const resizeObserver = new ResizeObserver(function() {

    clearTimeout(resizeTimer);

    resizeTimer = setTimeout(function() {
        resizeCanvas(hasInitialFit);
    }, 80);

});

resizeObserver.observe(viewer);


</script>

</body>
</html>
"""

    html = html.replace(
        "__MODE__",
        str(mode)
    )

    html = html.replace(
        "__BLINK_SPEED__",
        str(blink_speed)
    )

    html = html.replace(
        "__ORIGINAL_IMAGE__",
        original_b64
    )

    html = html.replace(
        "__OUTPUT_IMAGE__",
        output_b64
    )

    components.html(
        html,
        height=670,
        scrolling=False
    )


# =========================================================
# MAIN FUNCTION
# =========================================================

def main():

    st.markdown(
        """
        <style>
        .stTabs [data-baseweb="tab-list"] {
            gap: 12px;
            background: transparent;
            padding: 6px 0 14px 0;
        }

        .stTabs [data-baseweb="tab"] {
            height: 46px;
            min-width: 150px;
            border-radius: 14px;
            padding: 0 24px;
            border: 1px solid #3a4655;
            background: linear-gradient(180deg, #1b222c, #141a22);
            color: #cfd7e3;
            font-weight: 800;
            letter-spacing: .2px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.14);
            transition:
                transform 0.16s ease,
                background 0.18s ease,
                border-color 0.18s ease,
                box-shadow 0.18s ease,
                color 0.18s ease;
        }

        .stTabs [data-baseweb="tab"]:hover {
            transform: translateY(-2px);
            background: linear-gradient(180deg, #263342, #1b2632);
            border-color: #65809c;
            color: #ffffff;
            box-shadow: 0 7px 18px rgba(0,0,0,0.24);
        }

        .stTabs [data-baseweb="tab"]:active {
            transform: scale(0.97);
            box-shadow: 0 2px 8px rgba(0,0,0,0.24);
        }

        .stTabs [aria-selected="true"] {
            background: linear-gradient(135deg, #315a79, #1f3348) !important;
            border-color: #6fb6e8 !important;
            color: #ffffff !important;
            box-shadow:
                0 7px 22px rgba(42, 126, 185, 0.28),
                inset 0 1px 0 rgba(255,255,255,0.10);
            animation: activeTabPulse 1.8s ease-in-out infinite;
        }

        @keyframes activeTabPulse {
            0%, 100% {
                box-shadow:
                    0 7px 22px rgba(42, 126, 185, 0.24),
                    inset 0 1px 0 rgba(255,255,255,0.10);
            }
            50% {
                box-shadow:
                    0 9px 26px rgba(42, 126, 185, 0.34),
                    inset 0 1px 0 rgba(255,255,255,0.14);
            }
        }

        .stTabs [data-baseweb="tab-highlight"] {
            display: none !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.title("🔍 ORIGINAL SPEC TO OUTPUT CHECK")

    st.caption(
        "Visual comparison between Original Specification and Final Output"
    )

    # =====================================================
    # NEW START
    # =====================================================

    reset_col_left, reset_col_right = st.columns([7, 1])

    with reset_col_right:
        if st.button(
            "↻ NEW START",
            key="spec_new_start",
            width="stretch"
        ):
            st.session_state["spec_reset_id"] = (
                st.session_state.get("spec_reset_id", 0) + 1
            )

            st.session_state["spec_compare_started"] = False

            st.rerun()

    reset_id = st.session_state.get("spec_reset_id", 0)

    original_upload_key = f"original_spec_upload_{reset_id}"
    output_upload_key = f"output_spec_upload_{reset_id}"
    page_selector_key = f"spec_page_selector_{reset_id}"
    compare_key = f"spec_compare_{reset_id}"
    blink_speed_key = f"blink_speed_{reset_id}"

    st.divider()


    # =====================================================
    # UPLOADS
    # =====================================================

    col1, col2 = st.columns(2)

    with col1:

        original_spec = st.file_uploader(
            "📄 Upload Original Spec",
            type=["pdf"],
            key=original_upload_key
        )

    with col2:

        output_file = st.file_uploader(
            "📄 Upload Output",
            type=["pdf"],
            key=output_upload_key
        )


    if not original_spec or not output_file:

        st.info(
            "Upload both PDFs to begin comparison."
        )

        return


    # =====================================================
    # PAGE COUNTS
    # =====================================================

    original_pages = get_pdf_page_count(
        original_spec
    )

    output_pages = get_pdf_page_count(
        output_file
    )

    max_pages = min(
        original_pages,
        output_pages
    )


    # =====================================================
    # PAGE SELECTION
    # =====================================================

    if max_pages > 1:

        page_number = st.selectbox(
            "Select Page",
            options=list(range(max_pages)),
            format_func=lambda x:
                f"Page {x + 1}",
            key=page_selector_key
        )

    else:

        page_number = 0


    # =====================================================
    # COMPARE
    # =====================================================

    if st.button(
        "🔍 COMPARE",
        key=compare_key,
        type="primary",
        use_container_width=True
    ):

        st.session_state[
            "spec_compare_started"
        ] = True


    if not st.session_state.get(
        "spec_compare_started",
        False
    ):

        return


    # =====================================================
    # RENDER
    # =====================================================

    with st.spinner(
        "Preparing comparison..."
    ):

        original_image = pdf_page_to_image(
            original_spec,
            page_number,
            scale=2.0
        )

        output_image = pdf_page_to_image(
            output_file,
            page_number,
            scale=2.0
        )


    # =====================================================
    # MODES
    # =====================================================

    overlay_tab, blink_tab = st.tabs([
        "🟥🟢  OVERLAY",
        "👁  BLINK"
    ])


    # OVERLAY

    with overlay_tab:

        st.caption(
            "Original = Dark Red | "
            "Output = Dark Green | "
            "80% opacity"
        )

        comparison_viewer(
            original_image,
            output_image,
            mode="overlay"
        )


    # BLINK

    with blink_tab:

        blink_speed = st.slider(
            "⚡ Blink Speed",
            min_value=0.25,
            max_value=1.5,
            value=1.0,
            step=0.25,
            key=blink_speed_key
        )

        comparison_viewer(
            original_image,
            output_image,
            mode="blink",
            blink_speed=blink_speed
        )

# =========================================================
# RUN
# =========================================================

if __name__ == "__main__":
    main()
