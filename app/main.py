"""
Enterprise AI Presentation Architect — Main Application
Production-grade Streamlit UI for AI-powered presentation generation.
"""

import sys
import os
import io
import time
import json
from typing import List, Dict

# Add project root to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st

# ─── Page Config (must be first Streamlit call) ─────────────────────────────────
st.set_page_config(
    page_title="Enterprise AI Presentation Architect",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        "About": "Enterprise AI Presentation Architect v1.0 — Generate stunning presentations with AI."
    }
)

from utils.helpers import (
    init_session_state, reset_session_state, save_session_to_json,
    load_session_from_json, validate_pptx_file, validate_image_file,
    SlideContent, TemplateProfile, format_file_size, get_timestamp_str, logger
)
from core.content_engine import ContentEngine
from core.template_parser import PptxTemplateParser, ImageTemplateParser
from core.ppt_generator import PptGenerator
from core.preview_engine import PreviewEngine
from copy import deepcopy

# ─── Atomizer Engine ────────────────────────────────────────────────────────────

def atomize_slides(slides_list: List[Dict], enabled: bool = True) -> List[Dict]:
    """
    Decomposes dense slides into 'Atomic' Slides for 0% overlap guarantee.
    Only splits when actual chart/table data exists (not just image_prompt).
    """
    if not enabled:
        return slides_list

    new_slides = []
    for s in slides_list:
        bullets = s.get("bullet_points", []) or []
        has_data = bool(s.get("chart_data") or s.get("table_data"))  # only real data
        
        # Only atomize when there's actual visual data AND enough narrative content
        if has_data and len(bullets) >= 2:
            # 1. Narrative Slide (text only)
            s1 = deepcopy(s)
            s1["chart_data"] = None
            s1["table_data"] = None
            s1["image_prompt"] = None
            new_slides.append(s1)
            
            # 2. Visual Insight Slide (data only, clean title)
            s2 = deepcopy(s)
            s2["title"] = f"{s.get('title', '')} - Visual Insight"
            s2["bullet_points"] = []
            s2["subtitle"] = ""
            s2["image_prompt"] = None  # never show grey placeholder on data slides
            new_slides.append(s2)
        elif len(bullets) > 6:
            # Too many bullets even without data - split to keep layout clean
            s1 = deepcopy(s)
            s1["chart_data"] = None
            s1["table_data"] = None
            s1["image_prompt"] = None
            new_slides.append(s1)
        else:
            # single slide - strip image_prompt to avoid grey placeholders
            s["image_prompt"] = None
            new_slides.append(s)
            
    # Re-index
    for i, s in enumerate(new_slides):
        s["slide_number"] = i + 1
        
    return new_slides

# ─── Initialize ─────────────────────────────────────────────────────────────────
init_session_state()
content_engine = ContentEngine()
preview_engine = PreviewEngine()

# ─── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

:root {
    --bg-primary: #0a0a1a;
    --bg-secondary: #111128;
    --bg-card: #1a1a35;
    --accent: #00d4ff;
    --accent-2: #7c3aed;
    --text-primary: #f0f0f5;
    --text-secondary: #a0a0b8;
    --border: rgba(255,255,255,0.08);
    --success: #10b981;
    --warning: #f59e0b;
    --error: #ef4444;
}

html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
}

.stApp {
    background: linear-gradient(180deg, var(--bg-primary) 0%, #0d0d24 50%, var(--bg-primary) 100%) !important;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0f0f28 0%, #1a1a3e 100%) !important;
    border-right: 1px solid var(--border) !important;
}

section[data-testid="stSidebar"] .stMarkdown h1,
section[data-testid="stSidebar"] .stMarkdown h2,
section[data-testid="stSidebar"] .stMarkdown h3 {
    color: var(--accent) !important;
}

/* Main header */
.main-header {
    text-align: center;
    padding: 20px 0 10px;
    margin-bottom: 20px;
}
.main-header h1 {
    font-size: 2.2rem !important;
    font-weight: 800 !important;
    background: linear-gradient(135deg, #00d4ff, #7c3aed, #f472b6);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    margin-bottom: 4px !important;
}
.main-header p {
    color: var(--text-secondary);
    font-size: 0.95rem;
    font-weight: 300;
}

/* Cards */
.metric-card {
    background: linear-gradient(135deg, var(--bg-card), rgba(124,58,237,0.1));
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 18px 20px;
    text-align: center;
    transition: transform 0.2s, box-shadow 0.2s;
}
.metric-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 30px rgba(0,212,255,0.1);
}
.metric-card .value {
    font-size: 1.8rem;
    font-weight: 700;
    color: var(--accent);
}
.metric-card .label {
    font-size: 0.75rem;
    color: var(--text-secondary);
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-top: 4px;
}

/* Status badge */
.status-badge {
    display: inline-block;
    padding: 4px 14px;
    border-radius: 20px;
    font-size: 0.75rem;
    font-weight: 600;
    letter-spacing: 0.5px;
}
.status-idle { background: rgba(160,160,184,0.15); color: var(--text-secondary); }
.status-generating { background: rgba(0,212,255,0.15); color: var(--accent); }
.status-complete { background: rgba(16,185,129,0.15); color: var(--success); }
.status-error { background: rgba(239,68,68,0.15); color: var(--error); }

/* Buttons */
.stButton > button {
    border-radius: 10px !important;
    font-weight: 600 !important;
    letter-spacing: 0.3px !important;
    transition: all 0.3s !important;
}
div.stButton > button:first-child {
    background: linear-gradient(135deg, #00d4ff, #7c3aed) !important;
    color: white !important;
    border: none !important;
    padding: 10px 24px !important;
}
div.stButton > button:first-child:hover {
    box-shadow: 0 4px 20px rgba(0,212,255,0.4) !important;
    transform: translateY(-1px) !important;
}

/* Section dividers */
.section-divider {
    border: none;
    height: 1px;
    background: linear-gradient(90deg, transparent, var(--border), transparent);
    margin: 20px 0;
}

/* Expander styling */
.streamlit-expanderHeader {
    font-weight: 600 !important;
    color: var(--text-primary) !important;
}

/* Toast/alert */
.custom-alert {
    padding: 12px 18px;
    border-radius: 10px;
    margin: 8px 0;
    font-size: 0.85rem;
}
.alert-info { background: rgba(0,212,255,0.1); border-left: 3px solid var(--accent); color: var(--accent); }
.alert-success { background: rgba(16,185,129,0.1); border-left: 3px solid var(--success); color: var(--success); }
.alert-warn { background: rgba(245,158,11,0.1); border-left: 3px solid var(--warning); color: var(--warning); }
.alert-error { background: rgba(239,68,68,0.1); border-left: 3px solid var(--error); color: var(--error); }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 🏗️ Presentation Architect")
    st.markdown('<hr style="border:none;height:1px;background:rgba(255,255,255,0.1);margin:10px 0;">', unsafe_allow_html=True)

    # ── Template Upload ──────────────────────────────────────────────────
    st.markdown("### 📁 Template Upload")
    template_file = st.file_uploader(
        "Upload template (PPTX or Image)",
        type=["pptx", "png", "jpg", "jpeg"],
        key="template_uploader",
        help="Upload a PPTX file or slide image to clone its styling"
    )

    if template_file:
        file_bytes = template_file.getvalue()
        fname = template_file.name.lower()

        if fname != st.session_state.get("template_filename"):
            st.session_state["template_filename"] = fname
            st.session_state["template_bytes"] = file_bytes

            with st.spinner("🔍 Analyzing template..."):
                try:
                    if fname.endswith(".pptx"):
                        if validate_pptx_file(file_bytes):
                            parser = PptxTemplateParser(file_bytes)
                            profile = parser.parse()
                            st.session_state["template_profile"] = profile
                            st.success(f"✅ Template loaded: {profile.get('total_slides', 0)} slides, {len(profile.get('layouts', []))} layouts")
                        else:
                            st.error("❌ Invalid PPTX file")
                    else:
                        if validate_image_file(file_bytes):
                            parser = ImageTemplateParser(file_bytes, fname)
                            profile = parser.parse()
                            st.session_state["template_profile"] = profile
                            st.success("✅ Image template analyzed")
                        else:
                            st.error("❌ Invalid image file")
                except Exception as e:
                    st.error(f"❌ Template parsing error: {e}")
                    logger.error(f"Template parsing failed: {e}")

        if st.session_state.get("template_profile"):
            prof = st.session_state["template_profile"]
            cols = st.columns(2)
            cols[0].metric("Layouts", len(prof.get("layouts", [])))
            cols[1].metric("Slides", prof.get("total_slides", 0))

    st.markdown('<hr style="border:none;height:1px;background:rgba(255,255,255,0.1);margin:15px 0;">', unsafe_allow_html=True)

    # ── Controls ─────────────────────────────────────────────────────────
    st.markdown("### ⚙️ Controls")

    slide_count = st.slider(
        "Number of Slides",
        min_value=1, max_value=100, value=st.session_state.get("slide_count", 10),
        key="slide_count_slider",
        help="Select 1–100 slides. Large counts are batched automatically."
    )
    st.session_state["slide_count"] = slide_count

    # AI Toggle
    ai_enabled = st.toggle(
        "🤖 AI Generation",
        value=st.session_state.get("ai_enabled", True),
        key="ai_toggle",
        help="Enable/disable AI content generation"
    )
    st.session_state["ai_enabled"] = ai_enabled

    # Atomizer Toggle
    atomizer_enabled = st.toggle(
        "⚛️ Gamma Atomizer",
        value=st.session_state.get("atomizer_enabled", True),
        key="atomizer_toggle",
        help="Split dense slides into Narrative/Visual pairs (0% overlap safety)"
    )
    st.session_state["atomizer_enabled"] = atomizer_enabled

    # Model selector
    if ai_enabled:
        st.markdown("### 🧠 AI Model")
        with st.spinner("Loading models..."):
            try:
                model_names = content_engine.get_model_names()
            except Exception:
                model_names = ["llama3-70b-8192"]

        if model_names:
            current_model = st.session_state.get("selected_model")
            default_idx = 0
            if current_model and current_model in model_names:
                default_idx = model_names.index(current_model)

            selected_model = st.selectbox(
                "Select Model",
                model_names,
                index=default_idx,
                key="model_selector",
                help="Dynamically fetched from Groq API"
            )
            st.session_state["selected_model"] = selected_model

            # Show context window
            ctx = content_engine.get_model_context_window(selected_model)
            st.caption(f"Context window: {ctx:,} tokens")

    # ── Help & Documentation ─────────────────────────────────────────────
    st.markdown('<hr style="border:none;height:1px;background:rgba(255,255,255,0.1);margin:15px 0;">', unsafe_allow_html=True)
    with st.sidebar.expander("ℹ️ How to Use & Help"):
        st.markdown("""
        ### Quick Start Guide
        1. **Upload Template**: Drag & drop a PPTX or Image in the 'Assets' tab.
        2. **Set Prompt**: Enter your topic in the main text area.
        3. **Toggle Engine**: Turn on **Gamma Atomizer** if you want automatic Narrative/Visual slide splits.
        4. **Select Model**: Choose your preferred LLM (Llama-3 recommended).
        5. **Generate**: Click the big generate button & wait for the progress bar.
        
        ### Pro Tips
        - **Safe Top Margin**: The engine automatically enforces a **2.2" safe zone** to protect header branding.
        - **Sterile Rendering**: Visual slides use 'Blank' layouts to ensure zero overlap.
        - **Session Save**: Use the **💾 Save** button below to back up your progress as a JSON file.
        """)

    # ── Session Management ───────────────────────────────────────────────
    st.markdown("### 💾 Session")
    col1, col2 = st.columns(2)

    with col1:
        if st.button("🗑️ Reset", use_container_width=True, help="Clear all data"):
            reset_session_state()
            st.rerun()

    with col2:
        # Direct Download Button
        session_json = save_session_to_json()
        st.download_button(
            "💾 Save",
            data=session_json,
            file_name=f"session_{get_timestamp_str()}.json",
            mime="application/json",
            use_container_width=True,
            help="Download current session as a JSON file"
        )

    # Load session
    load_file = st.file_uploader("Load session", type=["json"], key="session_loader")
    if load_file:
        if load_session_from_json(load_file.getvalue().decode()):
            st.success("✅ Session loaded")
            st.rerun()
        else:
            st.error("❌ Failed to load session")

    # ── About Section ─────────────────────────────────────────────────────
    st.markdown('<hr style="border:none;height:1px;background:rgba(255,255,255,0.1);margin:15px 0;">', unsafe_allow_html=True)
    st.caption("🏗️ **Enterprise AI Presentation Architect v1.0**")
    st.caption("Designed for 100% Layout Fidelity.")
    st.markdown(f"[View on GitHub](https://github.com/VijaiVenkatesan/Enterprise_AI_Presentation_Architect_App)")


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN CONTENT AREA
# ═══════════════════════════════════════════════════════════════════════════════

# Header
st.markdown("""
<div class="main-header">
    <h1>🏗️ Enterprise AI Presentation Architect</h1>
    <p>Generate stunning, template-perfect presentations powered by AI + real-time research</p>
</div>
""", unsafe_allow_html=True)

# Status metrics
status = st.session_state.get("generation_status", "idle")
slides_data = st.session_state.get("slides_content", [])
has_template = st.session_state.get("template_profile") is not None

scol1, scol2, scol3, scol4 = st.columns(4)
with scol1:
    st.markdown(f'''<div class="metric-card">
        <div class="value">{"✅" if has_template else "—"}</div>
        <div class="label">Template</div></div>''', unsafe_allow_html=True)
with scol2:
    st.markdown(f'''<div class="metric-card">
        <div class="value">{st.session_state.get("slide_count", 10)}</div>
        <div class="label">Target Slides</div></div>''', unsafe_allow_html=True)
with scol3:
    st.markdown(f'''<div class="metric-card">
        <div class="value">{len(slides_data)}</div>
        <div class="label">Generated</div></div>''', unsafe_allow_html=True)
with scol4:
    status_class = {"idle":"idle","generating":"generating","complete":"complete","error":"error"}.get(status,"idle")
    status_text = {"idle":"Ready","generating":"Generating...","complete":"Complete","error":"Error"}.get(status,"Ready")
    st.markdown(f'''<div class="metric-card">
        <div class="status-badge status-{status_class}">{status_text}</div>
        <div class="label" style="margin-top:8px;">Status</div></div>''', unsafe_allow_html=True)

st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

# ── Topic Input ──────────────────────────────────────────────────────────────
st.markdown("### 💡 Presentation Topic")
topic = st.text_area(
    "Enter your presentation topic or detailed prompt",
    value=st.session_state.get("topic", ""),
    height=100,
    placeholder="e.g., 'AI Transformation in Healthcare: 2025 Trends, Market Analysis, and Implementation Roadmap'",
    key="topic_input",
    help="Be specific — include industry, focus areas, and desired angle"
)
st.session_state["topic"] = topic

# Additional instructions
with st.expander("📝 Additional Instructions (optional)"):
    custom_instructions = st.text_area(
        "Custom instructions for content generation",
        height=80,
        placeholder="e.g., 'Focus on ROI metrics, include competitor analysis, use consulting tone'",
        key="custom_instructions"
    )

# ── Generate Button ──────────────────────────────────────────────────────────
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

gen_col1, gen_col2, gen_col3 = st.columns([2, 1, 1])

with gen_col1:
    generate_clicked = st.button(
        "🚀 Generate Presentation",
        use_container_width=True,
        disabled=(not topic.strip()),
        key="generate_btn"
    )

with gen_col2:
    if slides_data and st.session_state.get("generated_pptx"):
        st.download_button(
            "📥 Download PPTX",
            data=st.session_state["generated_pptx"],
            file_name=f"presentation_{get_timestamp_str()}.pptx",
            mime="application/octet-stream",
            use_container_width=True,
            key="download_pptx_clean"
        )

with gen_col3:
    if slides_data and st.session_state.get("generated_pptx"):
        # PDF export via conversion
        if st.button("📄 Export PDF", use_container_width=True, key="export_pdf"):
            try:
                import base64
                from fpdf import FPDF
                import textwrap
                
                pdf = FPDF(orientation='L', unit='mm', format='A4')
                pdf.set_auto_page_break(auto=True, margin=15)
                pdf.set_margins(15, 15, 15)
                
                # Helper to strip unsupported unicode chars and break long words
                def safe_str(txt):
                    cleaned = str(txt).encode("latin-1", "replace").decode("latin-1")
                    # Break any single word longer than 35 chars to guarantee it fits on A4 Landscape even at 24pt font
                    words = cleaned.split(" ")
                    broken_words = [textwrap.fill(w, 35) if len(w) > 35 else w for w in words]
                    return " ".join(broken_words)
                
                for i, slide in enumerate(slides_data):
                    pdf.add_page()
                    pdf.set_font('Helvetica', 'B', 24)
                    # Use multi_cell for titles to ensure long titles automatically wrap instead of overflowing margins
                    pdf.multi_cell(0, 12, safe_str(slide.get("title", f"Slide {i+1}")), align='C')
                    pdf.ln(5)
                    
                    pdf.set_font('Helvetica', '', 14)
                    if slide.get("subtitle") and str(slide["subtitle"]).strip():
                        pdf.multi_cell(0, 8, safe_str(slide["subtitle"]), align='C')
                        pdf.ln(5)
                    
                    pdf.ln(5)
                    pdf.set_font('Helvetica', '', 12)
                    for bp in slide.get("bullet_points", []):
                        pdf.multi_cell(0, 8, safe_str(f"  -  {bp}"))
                        pdf.ln(2)
                        
                    if slide.get("notes") and str(slide["notes"]).strip():
                        pdf.ln(5)
                        pdf.set_font('Helvetica', 'I', 10)
                        pdf.multi_cell(0, 7, safe_str(f"Notes: {slide['notes']}"))
                        
                pdf_bytes = pdf.output()
                st.download_button(
                    "📥 Download PDF",
                    data=bytes(pdf_bytes),
                    file_name=f"presentation_{get_timestamp_str()}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    key="download_pdf_clean"
                )
            except ImportError:
                st.warning("PDF export requires fpdf2. Install: `pip install fpdf2`")
            except Exception as e:
                st.error(f"PDF export failed: {e}")


# ── Generation Logic ─────────────────────────────────────────────────────────
if generate_clicked:
    if not topic.strip():
        st.error("Please enter a presentation topic.")
    elif st.session_state.get("ai_enabled") and not content_engine.api_key:
        st.error("🔑 Groq API key not found. Add `GROQ_API_KEY` to your Streamlit secrets.")
    else:
        st.session_state["generation_status"] = "generating"
        st.session_state["error_message"] = None

        progress_bar = st.progress(0, text="🔄 Initializing...")
        status_text = st.empty()

        def update_progress(pct, msg):
            progress_bar.progress(min(pct, 1.0), text=msg)
            status_text.markdown(f'<div class="custom-alert alert-info">{msg}</div>', unsafe_allow_html=True)

        try:
            if st.session_state.get("ai_enabled"):
                model = st.session_state.get("selected_model", "llama3-70b-8192")
                ci = st.session_state.get("custom_instructions", "")

                content, error = content_engine.generate_presentation_content(
                    topic=topic,
                    slide_count=slide_count,
                    model_id=model,
                    template_profile=st.session_state.get("template_profile"),
                    custom_instructions=ci,
                    progress_callback=update_progress,
                )

                if error:
                    st.session_state["generation_status"] = "error"
                    st.session_state["error_message"] = error
                    st.error(f"❌ {error}")
                else:
                    slides = content.get("slides", [])
                    
                    # THE GAMMA ATOMIZER: Decompose dense slides into Narrative + Data
                    update_progress(0.85, "🔬 Atomizing dense slides for 0% overlap...")
                    atomizer_on = st.session_state.get("atomizer_enabled", True)
                    
                    if not atomizer_on:
                        print("DEBUG: Atomizer OFF - Stripping visuals for compact layout")
                        for s in slides:
                            s["chart_data"] = None
                            s["table_data"] = None
                            s["image_prompt"] = None
                            
                    slides = atomize_slides(slides, enabled=atomizer_on)
                    print(f"DEBUG: Slides after processing: {len(slides)}")
                    v_slides = [s for s in slides if s.get("chart_data") or s.get("table_data")]
                    print(f"DEBUG: Slides with visuals: {len(v_slides)}")
                    
                    # Add Mandatory THANK YOU slide (User Demand)
                    slides.append({
                        "title": "Thank You",
                        "subtitle": "Questions & Discussion",
                        "bullet_points": [],
                        "notes": "End of presentation."
                    })
                    
                    st.session_state["slides_content"] = slides

                    # Generate PPTX
                    update_progress(0.92, "📦 Building PowerPoint file...")
                    generator = PptGenerator(
                        template_bytes=st.session_state.get("template_bytes"),
                        template_profile=st.session_state.get("template_profile")
                    )

                    def ppt_progress(pct, msg):
                        update_progress(0.92 + pct * 0.08, msg)

                    pptx_bytes = generator.generate(slides, progress_callback=ppt_progress)
                    st.session_state["generated_pptx"] = pptx_bytes
                    st.session_state["generation_status"] = "complete"
                    update_progress(1.0, "✅ Presentation generated successfully!")
                    time.sleep(1)
                    st.rerun()
            else:
                # Manual mode — create placeholder slides
                update_progress(0.3, "📝 Creating slide placeholders...")
                slides = []
                for i in range(slide_count):
                    slides.append({
                        "slide_number": i + 1,
                        "title": f"Slide {i + 1}" if i > 0 else topic[:60],
                        "subtitle": "" if i > 0 else "Edit content below",
                        "bullet_points": ["Click edit to add content"],
                        "chart_data": None, "table_data": None,
                        "diagram_type": None, "image_prompt": None,
                        "notes": "", "layout_index": 1 if i > 0 else 0
                    })
                st.session_state["slides_content"] = slides

                update_progress(0.7, "📦 Building PowerPoint...")
                generator = PptGenerator(
                    template_bytes=st.session_state.get("template_bytes"),
                    template_profile=st.session_state.get("template_profile")
                )
                pptx_bytes = generator.generate(slides)
                st.session_state["generated_pptx"] = pptx_bytes
                st.session_state["generation_status"] = "complete"
                update_progress(1.0, "✅ Placeholders created — edit below!")
                time.sleep(1)
                st.rerun()

        except Exception as e:
            st.session_state["generation_status"] = "error"
            st.session_state["error_message"] = str(e)
            st.error(f"❌ Generation failed: {e}")
            logger.error(f"Generation error: {e}", exc_info=True)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE PREVIEW + EDITING
# ═══════════════════════════════════════════════════════════════════════════════

if slides_data:
    st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
    st.markdown("### 🖼️ Slide Preview & Editor")

    # Navigation
    preview_tab, editor_tab = st.tabs(["📺 Preview", "✏️ Editor"])

    with preview_tab:
        previews = preview_engine.render_all_previews(slides_data)
        # Display in a 2-column grid
        for i in range(0, len(previews), 2):
            cols = st.columns(2)
            with cols[0]:
                st.markdown(previews[i], unsafe_allow_html=True)
            if i + 1 < len(previews):
                with cols[1]:
                    st.markdown(previews[i + 1], unsafe_allow_html=True)

    with editor_tab:
        st.markdown("Edit individual slide content below. Changes are saved when you click **Apply Changes**.")

        for i, slide in enumerate(slides_data):
            with st.expander(f"📝 Slide {slide.get('slide_number', i+1)}: {slide.get('title', '')[:50]}", expanded=False):
                ecol1, ecol2 = st.columns([3, 1])

                with ecol1:
                    new_title = st.text_input(
                        "Title", value=slide.get("title", ""),
                        key=f"edit_title_{i}"
                    )
                    new_subtitle = st.text_input(
                        "Subtitle", value=slide.get("subtitle", ""),
                        key=f"edit_subtitle_{i}"
                    )
                    bullets_text = "\n".join(slide.get("bullet_points", []))
                    new_bullets = st.text_area(
                        "Bullet Points (one per line)",
                        value=bullets_text, height=120,
                        key=f"edit_bullets_{i}"
                    )
                    new_notes = st.text_area(
                        "Speaker Notes",
                        value=slide.get("notes", ""), height=60,
                        key=f"edit_notes_{i}"
                    )

                with ecol2:
                    st.markdown("**Actions**")

                    # Regenerate single slide
                    if st.session_state.get("ai_enabled"):
                        if st.button("🔄 Regenerate", key=f"regen_{i}", use_container_width=True):
                            with st.spinner("Regenerating..."):
                                context = slides_data[max(0,i-1):min(len(slides_data),i+2)]
                                new_slide, err = content_engine.regenerate_single_slide(
                                    slide_number=slide.get("slide_number", i+1),
                                    topic=topic,
                                    model_id=st.session_state.get("selected_model", "llama3-70b-8192"),
                                    context_slides=context
                                )
                                if new_slide and not err:
                                    st.session_state["slides_content"][i] = new_slide
                                    st.success("✅ Regenerated!")
                                    st.rerun()
                                else:
                                    st.error(f"Failed: {err}")

                    # Move up/down
                    mcol1, mcol2 = st.columns(2)
                    with mcol1:
                        if i > 0 and st.button("⬆️", key=f"up_{i}", use_container_width=True):
                            slides_data[i], slides_data[i-1] = slides_data[i-1], slides_data[i]
                            for j, s in enumerate(slides_data):
                                s["slide_number"] = j + 1
                            st.session_state["slides_content"] = slides_data
                            st.rerun()
                    with mcol2:
                        if i < len(slides_data)-1 and st.button("⬇️", key=f"down_{i}", use_container_width=True):
                            slides_data[i], slides_data[i+1] = slides_data[i+1], slides_data[i]
                            for j, s in enumerate(slides_data):
                                s["slide_number"] = j + 1
                            st.session_state["slides_content"] = slides_data
                            st.rerun()

                    # Delete slide
                    if len(slides_data) > 1:
                        if st.button("🗑️ Delete", key=f"del_{i}", use_container_width=True):
                            slides_data.pop(i)
                            for j, s in enumerate(slides_data):
                                s["slide_number"] = j + 1
                            st.session_state["slides_content"] = slides_data
                            st.rerun()

                # Apply edits
                if st.button("✅ Apply Changes", key=f"apply_{i}", use_container_width=True):
                    st.session_state["slides_content"][i]["title"] = new_title
                    st.session_state["slides_content"][i]["subtitle"] = new_subtitle
                    st.session_state["slides_content"][i]["bullet_points"] = [
                        b.strip() for b in new_bullets.split("\n") if b.strip()
                    ]
                    st.session_state["slides_content"][i]["notes"] = new_notes
                    st.success(f"✅ Slide {i+1} updated!")

        # Rebuild PPTX after edits
        st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
        if st.button("🔄 Rebuild Presentation with Edits", use_container_width=True, key="rebuild_btn"):
            with st.spinner("Rebuilding presentation..."):
                try:
                    generator = PptGenerator(
                        template_bytes=st.session_state.get("template_bytes"),
                        template_profile=st.session_state.get("template_profile")
                    )
                    pptx_bytes = generator.generate(st.session_state["slides_content"])
                    st.session_state["generated_pptx"] = pptx_bytes
                    st.success("✅ Presentation rebuilt with edits!")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Rebuild failed: {e}")


# ═══════════════════════════════════════════════════════════════════════════════
# FOOTER
# ═══════════════════════════════════════════════════════════════════════════════

st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
st.markdown("""
<div style="text-align:center; padding:20px 0; color:rgba(255,255,255,0.3); font-size:0.75rem;">
    Enterprise AI Presentation Architect v1.0 · Powered by Groq AI + DuckDuckGo ·
    Built with Streamlit
</div>
""", unsafe_allow_html=True)
