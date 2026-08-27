import streamlit as st
from textwrap import dedent


# ==========================================================
# PAGE CONFIG
# ==========================================================

st.set_page_config(
    page_title="QC Validation System",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="collapsed"
)


# ==========================================================
# CUSTOM DESIGN / CSS
# ==========================================================

st.markdown("""
<style>

/* ==========================================================
   MAIN BACKGROUND
========================================================== */

.stApp {
    background:
        radial-gradient(circle at 15% 20%, rgba(33, 150, 243, 0.18), transparent 25%),
        radial-gradient(circle at 85% 15%, rgba(126, 87, 194, 0.18), transparent 25%),
        radial-gradient(circle at 50% 90%, rgba(0, 188, 212, 0.10), transparent 30%),
        #111827;
    color: white;
}

[data-testid="stHeader"] {
    background: transparent;
}

[data-testid="stToolbar"] {
    right: 1rem;
}


/* ==========================================================
   MAIN CONTENT WIDTH
========================================================== */

.block-container {
    padding-top: 2.5rem;
    padding-bottom: 3rem;
    max-width: 1500px;
}


/* ==========================================================
   HEADER
========================================================== */

.system-badge {
    display: inline-block;
    padding: 7px 16px;
    border-radius: 30px;

    background: rgba(33, 150, 243, 0.12);

    border: 1px solid rgba(96, 165, 250, 0.35);

    color: #93c5fd;

    font-size: 13px;
    font-weight: 700;

    letter-spacing: 1px;
}


.main-title {
    font-size: 46px;

    font-weight: 800;

    margin-top: 20px;

    margin-bottom: 10px;

    background: linear-gradient(
        90deg,
        #ffffff,
        #93c5fd,
        #c4b5fd
    );

    -webkit-background-clip: text;

    -webkit-text-fill-color: transparent;
}


.main-subtitle {
    color: #aeb8c7;

    font-size: 17px;

    margin-bottom: 45px;
}


/* ==========================================================
   TOOL CARDS
========================================================== */

.tool-card {
    position: relative;

    height: 340px;

    padding: 24px;

    border-radius: 24px;

    background:
        linear-gradient(
            145deg,
            rgba(30, 41, 59, 0.92),
            rgba(15, 23, 42, 0.92)
        );

    border:
        1px solid rgba(148, 163, 184, 0.20);

    box-shadow:
        0 20px 45px rgba(0, 0, 0, 0.25);

    overflow: hidden;

    transition:
        transform 0.35s ease,
        border 0.35s ease,
        box-shadow 0.35s ease;
}


.tool-card:hover {
    transform: translateY(-8px);

    border:
        1px solid rgba(96, 165, 250, 0.75);

    box-shadow:
        0 25px 60px rgba(37, 99, 235, 0.18);
}


/* ==========================================================
   BACKGROUND DECORATION
========================================================== */

.vector-circle {
    position: absolute;

    width: 150px;

    height: 150px;

    border-radius: 50%;

    background:
        linear-gradient(
            135deg,
            rgba(59, 130, 246, 0.25),
            rgba(139, 92, 246, 0.05)
        );

    right: -45px;

    top: -45px;

    filter: blur(1px);
}


.vector-grid {
    position: absolute;

    width: 140px;

    height: 140px;

    right: -30px;

    bottom: -40px;

    opacity: 0.12;

    background-image:
        linear-gradient(#60a5fa 1px, transparent 1px),
        linear-gradient(90deg, #60a5fa 1px, transparent 1px);

    background-size: 18px 18px;

    transform: rotate(15deg);
}


/* ==========================================================
   TOOL ICON AREA
========================================================== */

.tool-icon {
    width: 70px;

    height: 70px;

    display: flex;

    align-items: center;

    justify-content: center;

    border-radius: 20px;

    font-size: 34px;

    margin-bottom: 25px;

    background:
        linear-gradient(
            135deg,
            rgba(59, 130, 246, 0.25),
            rgba(139, 92, 246, 0.20)
        );

    border:
        1px solid rgba(147, 197, 253, 0.25);

    position: relative;

    z-index: 2;
}


/* ==========================================================
   TEXT
========================================================== */

.tool-number {
    color: #60a5fa;

    font-size: 12px;

    font-weight: 800;

    letter-spacing: 1.5px;

    margin-bottom: 10px;
}


.tool-title {
    color: white;

    font-size: 21px;

    font-weight: 800;

    line-height: 1.25;

    margin-bottom: 14px;
}


.tool-description {
    color: #aab4c3;

    font-size: 14px;

    line-height: 1.55;
}


/* ==========================================================
   BUTTONS
========================================================== */

div.stButton > button {
    width: 100%;

    min-height: 48px;

    border-radius: 14px;

    border:
        1px solid rgba(147, 197, 253, 0.55) !important;

    background:
        linear-gradient(
            90deg,
            #2563eb,
            #4f46e5
        ) !important;

    color: white !important;

    font-weight: 700;

    font-size: 14px;

    transition: all 0.25s ease;
}


div.stButton > button:hover {
    transform: translateY(-2px);

    border-color: #ffffff !important;

    box-shadow:
        0 10px 30px rgba(59, 130, 246, 0.35);
}


/* ==========================================================
   FOOTER
========================================================== */

.dashboard-footer {
    text-align: center;

    margin-top: 55px;

    color: #64748b;

    font-size: 13px;
}


/* ==========================================================
   BACK BUTTON
========================================================== */

.back-title {
    font-size: 18px;

    font-weight: 700;

    color: #ffffff;

    margin-bottom: 25px;
}

</style>
""", unsafe_allow_html=True)


# ==========================================================
# SESSION STATE
# ==========================================================

if "selected_tool" not in st.session_state:
    st.session_state.selected_tool = None


# ==========================================================
# DASHBOARD
# ==========================================================

def show_dashboard():

    # ======================================================
    # HEADER
    # ======================================================

    st.html(dedent("""
    <div style="text-align:center;">

        <div class="system-badge">
            QUALITY CONTROL PLATFORM
        </div>

        <div class="main-title">
            QC VALIDATION SYSTEM
        </div>

        <div class="main-subtitle">
            Smart validation for Order Forms, Specifications and Output Artwork
        </div>

    </div>
    """))


    # ======================================================
    # FOUR CARDS
    # ======================================================

    col1, col2, col3, col4 = st.columns(4, gap="medium")


    # ======================================================
    # TOOL 1
    # ======================================================

    with col1:

        st.html(dedent("""
        <div class="tool-card">

            <div class="vector-circle"></div>

            <div class="vector-grid"></div>

            <div class="tool-icon">
                📋
            </div>

            <div class="tool-number">
                TOOL 01
            </div>

            <div class="tool-title">
                ORDER-FORM<br>
                TO OUTPUT CHECK
            </div>

            <div class="tool-description">
                Validate selected Order Form data
                against the final PDF output and
                identify spelling mistakes,
                missing data and mismatches.
            </div>

        </div>
        """))

        if st.button(
            "OPEN TOOL  →",
            key="tool_1"
        ):
            st.session_state.selected_tool = "order_form"
            st.rerun()


    # ======================================================
    # TOOL 2
    # ======================================================

    with col2:

        st.html(dedent("""
        <div class="tool-card">

            <div class="vector-circle"></div>

            <div class="vector-grid"></div>

            <div class="tool-icon">
                📑
            </div>

            <div class="tool-number">
                TOOL 02
            </div>

            <div class="tool-title">
                ORIGINAL SPEC<br>
                TO OUTPUT CHECK
            </div>

            <div class="tool-description">
                Compare Original Specification
                requirements against the final
                output artwork and identify
                potential QC issues.
            </div>

        </div>
        """))

        if st.button(
            "OPEN TOOL  →",
            key="tool_2"
        ):
            st.session_state.selected_tool = "original_spec"
            st.rerun()


    # ======================================================
    # TOOL 3
    # ======================================================

    with col3:

        st.html(dedent("""
        <div class="tool-card">

            <div class="vector-circle"></div>

            <div class="vector-grid"></div>

            <div class="tool-icon">
                🔄
            </div>

            <div class="tool-number">
                TOOL 03
            </div>

            <div class="tool-title">
                SPEC + ORDER FORM<br>
                + OUTPUT CHECK
            </div>

            <div class="tool-description">
                Cross-check Original Specification,
                Order Form data and final output
                together for complete validation.
            </div>

        </div>
        """))

        if st.button(
            "OPEN TOOL  →",
            key="tool_3"
        ):
            st.session_state.selected_tool = "spec_order_output"
            st.rerun()


    # ======================================================
    # TOOL 4
    # ======================================================

    with col4:

        st.html(dedent("""
        <div class="tool-card">

            <div class="vector-circle"></div>

            <div class="vector-grid"></div>

            <div class="tool-icon">
                🛠️
            </div>

            <div class="tool-number">
                TOOL 04
            </div>

            <div class="tool-title">
                MORE QC<br>
                TOOLS
            </div>

            <div class="tool-description">
                Access additional quality-control
                utilities, validation tools and
                future QC automation features.
            </div>

        </div>
        """))

        if st.button(
            "EXPLORE TOOLS  →",
            key="tool_4"
        ):
            st.session_state.selected_tool = "more_qc"
            st.rerun()


    # ======================================================
    # FOOTER
    # ======================================================

    st.html(dedent("""
    <div class="dashboard-footer">
        SELECT A TOOL TO START YOUR QUALITY CONTROL PROCESS
    </div>
    """))


# ==========================================================
# TOOL ROUTER
# ==========================================================

def show_selected_tool():

    if st.button("←  BACK TO DASHBOARD"):
        st.session_state.selected_tool = None
        st.rerun()

    st.divider()


    # ======================================================
    # TOOL 1
    # ======================================================

    if st.session_state.selected_tool == "order_form":

        from tools.order_form_output_check import main
        main()


    # ======================================================
    # TOOL 2
    # ======================================================

    elif st.session_state.selected_tool == "original_spec":

        from tools.original_spec_output_check import main
        main()


    # ======================================================
    # TOOL 3
    # ======================================================

    elif st.session_state.selected_tool == "spec_order_output":

        from tools.spec_order_form_output_check import main
        main()


    # ======================================================
    # TOOL 4
    # ======================================================

    elif st.session_state.selected_tool == "more_qc":

        from tools.more_qc_tools import main
        main()


# ==========================================================
# RUN APPLICATION
# ==========================================================

if st.session_state.selected_tool is None:

    show_dashboard()

else:

    show_selected_tool()
