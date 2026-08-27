import streamlit as st
import fitz
from PIL import Image
import io
import base64
import streamlit.components.v1 as components


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

#wrapper {
    width: 100%;
    display: flex;
    justify-content: center;
}

#viewer {
    width: 1100px;
    height: 750px;
    background: #ffffff;
    border: 1px solid #333;
    border-radius: 12px;
    position: relative;
    overflow: hidden;
    cursor: grab;
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
    bottom: 12px;
    left: 50%;
    transform: translateX(-50%);
    display: flex;
    gap: 8px;
    background: rgba(20,20,20,0.9);
    padding: 8px;
    border-radius: 8px;
}

#controls button {
    background: #252525;
    color: white;
    border: 1px solid #555;
    padding: 6px 12px;
    border-radius: 5px;
    cursor: pointer;
}

#controls button:hover {
    background: #444;
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

const BLINK_SPEED = __BLINK_SPEED__ * 1000;

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

function resizeCanvas() {

    const rect =
        viewer.getBoundingClientRect();

    canvas.width = rect.width;

    canvas.height = rect.height;

    draw();
}


// =========================================================
// FIT IMAGE
// =========================================================

function fitImage() {

    if (!original.width) return;

    const padding = 40;

    const scaleX =
        (canvas.width - padding)
        / original.width;

    const scaleY =
        (canvas.height - padding)
        / original.height;

    scale =
        Math.min(scaleX, scaleY);

    offsetX =
        (canvas.width -
        original.width * scale) / 2;

    offsetY =
        (canvas.height -
        original.height * scale) / 2;

    draw();
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

        ctx.globalAlpha =
            ORIGINAL_OPACITY;

        ctx.drawImage(
            original,
            offsetX,
            offsetY,
            original.width * scale,
            original.height * scale
        );


        ctx.globalAlpha =
            OUTPUT_OPACITY;

        ctx.drawImage(
            output,
            offsetX,
            offsetY,
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

        ctx.drawImage(
            activeImage,
            offsetX,
            offsetY,
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

        // Wait briefly so the Streamlit component
        // receives its final width and height.
        setTimeout(function() {

            resizeCanvas();

            // Start every mode centered and fitted
            fitImage();

            // Make sure Blink always starts
            // from the centered Original Spec.
            blinkShowOriginal = true;

            draw();

            // Start blink only after correct positioning
            if (MODE === "blink") {

                setInterval(function() {

                    blinkShowOriginal =
                        !blinkShowOriginal;

                    draw();

                }, BLINK_SPEED);
            }

        }, 200);
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
    resizeCanvas
);


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
        height=680,
        scrolling=False
    )


# =========================================================
# DATA CHECK
# =========================================================

def data_check_view():

    st.subheader("🔍 Data Check")

    st.info(
        "Automatic text and difference validation "
        "will be added here next."
    )

    st.selectbox(
        "Difference Classification",
        [
            "Select Type",
            "STATIC",
            "VARIABLE",
            "IGNORE"
        ],
        key="spec_data_classification"
    )


# =========================================================
# MAIN FUNCTION
# =========================================================

def main():

    st.title("🔍 ORIGINAL SPEC TO OUTPUT CHECK")

    st.caption(
        "Visual comparison between Original Specification and Final Output"
    )

    st.divider()


    # =====================================================
    # UPLOADS
    # =====================================================

    col1, col2 = st.columns(2)

    with col1:

        original_spec = st.file_uploader(
            "📄 Upload Original Spec",
            type=["pdf"],
            key="original_spec_upload"
        )

    with col2:

        output_file = st.file_uploader(
            "📄 Upload Output",
            type=["pdf"],
            key="output_spec_upload"
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
            key="spec_page_selector"
        )

    else:

        page_number = 0


    # =====================================================
    # COMPARE
    # =====================================================

    if st.button(
        "🔍 COMPARE",
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

    overlay_tab, blink_tab, data_tab = st.tabs([
        "🟥🟢 OVERLAY",
        "👁 BLINK",
        "🔍 DATA CHECK"
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
            "⚡ Blink Speed (seconds)",
            min_value=0.25,
            max_value=2.0,
            value=0.5,
            step=0.25,
            key="blink_speed"
        )

        comparison_viewer(
            original_image,
            output_image,
            mode="blink",
            blink_speed=blink_speed
        )


    # DATA CHECK

    with data_tab:

        data_check_view()


# =========================================================
# RUN
# =========================================================

if __name__ == "__main__":
    main()
