import streamlit as st
import subprocess
import tempfile
import os
import zipfile
import io
from pathlib import Path
import shutil

from PIL import Image
import fitz  # PyMuPDF
try:
    from pdf2docx import Converter
    PDF2DOCX_AVAILABLE = True
except Exception:
    PDF2DOCX_AVAILABLE = False
import pdfplumber
import pandas as pd

# ----------------------------------------------------------------------------
# Page config
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="محوّل الملفات | File Converter",
    page_icon="🔄",
    layout="wide",
)

# ----------------------------------------------------------------------------
# Styling
# ----------------------------------------------------------------------------
st.markdown(
    """
    <link href="https://fonts.googleapis.com/css2?family=Orbitron:wght@600;800;900&family=Cairo:wght@400;600;800&display=swap" rel="stylesheet">
    <style>
    :root {
        --neon-pink: #ff007f;
        --neon-cyan: #00f0ff;
    }

    html, body, [class*="css"] {
        font-family: 'Cairo', sans-serif;
    }

    .stApp {
        background: radial-gradient(ellipse at top, #1a0b2e 0%, #0d0518 45%, #050208 100%);
        background-attachment: fixed;
    }

    .main .block-container {
        direction: rtl;
        max-width: 1150px;
    }
    h1, h2, h3, p, label, .stMarkdown {
        text-align: right;
        color: #e8e6f0;
    }

    .converter-title {
        text-align: center;
        font-family: 'Orbitron', 'Cairo', sans-serif;
        font-weight: 900;
        font-size: 2.6rem;
        letter-spacing: 1px;
        background: linear-gradient(90deg, var(--neon-pink), var(--neon-cyan));
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-shadow: 0 0 25px rgba(255, 0, 127, 0.35);
        margin-bottom: 0;
        animation: glow-pulse 3s ease-in-out infinite;
    }
    @keyframes glow-pulse {
        0%, 100% { filter: drop-shadow(0 0 6px rgba(0, 240, 255, 0.35)); }
        50% { filter: drop-shadow(0 0 18px rgba(255, 0, 127, 0.55)); }
    }
    .converter-subtitle {
        text-align: center;
        color: #a9a3c2;
        margin-top: 0.2rem;
        margin-bottom: 1.6rem;
        font-size: 1rem;
    }

    .section-title {
        font-family: 'Orbitron', 'Cairo', sans-serif;
        font-weight: 800;
        font-size: 1.25rem;
        color: var(--neon-cyan);
        margin: 1.6rem 0 0.8rem 0;
        text-shadow: 0 0 10px rgba(0, 240, 255, 0.35);
    }

    .tool-card {
        background: rgba(255, 255, 255, 0.035);
        border: 1px solid rgba(0, 240, 255, 0.22);
        border-radius: 16px;
        padding: 1rem;
        text-align: center;
        transition: all 0.2s ease;
        margin-bottom: 0.5rem;
    }
    .tool-card:hover {
        border: 1px solid var(--neon-cyan);
        box-shadow: 0 0 18px rgba(0, 240, 255, 0.35);
        transform: translateY(-3px);
    }
    .tool-icon {
        font-size: 2.3rem;
        margin-bottom: 0.2rem;
        filter: drop-shadow(0 0 8px rgba(0, 240, 255, 0.3));
    }

    .stButton>button, .stDownloadButton>button {
        width: 100%;
        border-radius: 12px;
        font-weight: 900;
        font-size: 1.0rem;
        padding: 0.75rem;
        border: 2px solid var(--neon-cyan);
        background: linear-gradient(90deg, var(--neon-pink), var(--neon-cyan));
        color: #050208;
        box-shadow: 0 0 16px rgba(0, 240, 255, 0.5), 0 0 10px rgba(255, 0, 127, 0.4);
        transition: all 0.2s ease;
        letter-spacing: 0.3px;
    }
    .stButton>button:hover, .stDownloadButton>button:hover {
        border: 2px solid #ffffff;
        box-shadow: 0 0 26px rgba(255, 0, 127, 0.75), 0 0 26px rgba(0, 240, 255, 0.6);
        transform: translateY(-2px) scale(1.01);
    }

    .card-btn button {
        background: rgba(255, 255, 255, 0.03) !important;
        border: 1px solid rgba(0, 240, 255, 0.3) !important;
        color: #e8e6f0 !important;
        box-shadow: none !important;
        font-weight: 700 !important;
        font-size: 0.92rem !important;
        padding: 0.5rem !important;
    }
    .card-btn button:hover {
        border: 1px solid var(--neon-pink) !important;
        color: #ffffff !important;
        box-shadow: 0 0 10px rgba(255, 0, 127, 0.45) !important;
    }

    .back-btn button {
        width: auto !important;
        background: rgba(255, 255, 255, 0.05) !important;
        border: 1px solid rgba(0, 240, 255, 0.35) !important;
        color: #e8e6f0 !important;
        box-shadow: none !important;
        font-weight: 700 !important;
        padding: 0.4rem 1rem !important;
    }

    [data-testid="stFileUploaderDropzone"] {
        background: rgba(255, 255, 255, 0.02);
        border: 1.5px dashed rgba(0, 240, 255, 0.4) !important;
        border-radius: 14px;
    }

    .stTextInput input {
        border-radius: 12px;
        border: 1.5px solid rgba(0, 240, 255, 0.3);
    }

    footer {visibility: hidden;}
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    [data-testid="stToolbar"] {visibility: hidden !important; display: none !important;}
    [data-testid="stDecoration"] {display: none !important;}
    .stDeployButton {display: none !important;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<p class="converter-title">🔄 محوّل الملفات</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="converter-subtitle">أكتر من 30 أداة لتحويل المستندات والصور بين الصيغ المختلفة — اختار الأداة وابدأ</p>',
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------

def check_soffice():
    """Check if LibreOffice soffice is available."""
    return shutil.which("soffice") is not None


def soffice_convert(input_path: str, target_format: str, outdir: str) -> str:
    """Convert a file using headless LibreOffice. Returns output file path."""
    if not check_soffice():
        raise RuntimeError(
            "❌ LibreOffice مش موجود على الجهاز.\n"
            "لو شغال على الكمبيوتر: ثبّت LibreOffice من الموقع الرسمي.\n"
            "لو شغال على Streamlit Cloud: أنشئ ملف packages.txt واكتب فيه 'libreoffice'."
        )
    cmd = [
        "soffice", "--headless", "--norestore",
        "--convert-to", target_format,
        "--outdir", outdir, input_path,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    stem = Path(input_path).stem
    out_path = os.path.join(outdir, f"{stem}.{target_format}")
    if not os.path.exists(out_path):
        raise RuntimeError(
            "فشل التحويل عبر LibreOffice.\n"
            f"stdout: {result.stdout}\nstderr: {result.stderr}"
        )
    return out_path


def make_zip(files: dict) -> bytes:
    """files: {filename: bytes}"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)
    buf.seek(0)
    return buf.read()


IMAGE_FORMATS = ["PNG", "JPEG", "WEBP", "BMP", "GIF", "TIFF", "ICO", "PDF"]
IMAGE_EXTS = ["png", "jpg", "jpeg", "webp", "bmp", "gif", "tiff", "tif", "ico"]
IMAGE_EXT_SET = set(IMAGE_EXTS)

PIL_FORMAT = {
    "png": "PNG", "jpg": "JPEG", "jpeg": "JPEG", "webp": "WEBP",
    "bmp": "BMP", "gif": "GIF", "tiff": "TIFF", "tif": "TIFF", "ico": "ICO",
}

UPLOAD_EXTS = {
    "docx": ["docx"], "doc": ["doc"], "odt": ["odt"], "rtf": ["rtf"],
    "txt": ["txt"], "html": ["html", "htm"], "pdf": ["pdf"],
    "xlsx": ["xlsx"], "xls": ["xls"], "ods": ["ods"], "csv": ["csv"],
    "pptx": ["pptx"], "ppt": ["ppt"], "odp": ["odp"],
    "png": ["png"], "jpg": ["jpg", "jpeg"], "webp": ["webp"], "bmp": ["bmp"],
    "gif": ["gif"], "tiff": ["tiff", "tif"], "ico": ["ico"],
}

MIME = {
    "pdf": "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "doc": "application/msword",
    "odt": "application/vnd.oasis.opendocument.text",
    "rtf": "application/rtf",
    "txt": "text/plain",
    "html": "text/html",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "xls": "application/vnd.ms-excel",
    "ods": "application/vnd.oasis.opendocument.spreadsheet",
    "csv": "text/csv",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "ppt": "application/vnd.ms-powerpoint",
    "odp": "application/vnd.oasis.opendocument.presentation",
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "webp": "image/webp",
    "bmp": "image/bmp",
    "gif": "image/gif",
    "tiff": "image/tiff",
    "ico": "image/x-icon",
    "zip": "application/zip",
}

# ----------------------------------------------------------------------------
# Tool registry
# ----------------------------------------------------------------------------
TOOLS = {
    "docx2pdf":  {"icon": "📝", "label": "Word إلى PDF",              "from": "docx", "to": "pdf", "cat": "to_pdf"},
    "doc2pdf":   {"icon": "📝", "label": "Word 97-2003 إلى PDF",      "from": "doc",  "to": "pdf", "cat": "to_pdf"},
    "odt2pdf":   {"icon": "📝", "label": "ODT إلى PDF",                "from": "odt",  "to": "pdf", "cat": "to_pdf"},
    "rtf2pdf":   {"icon": "📝", "label": "RTF إلى PDF",                "from": "rtf",  "to": "pdf", "cat": "to_pdf"},
    "txt2pdf":   {"icon": "📄", "label": "نص إلى PDF",                 "from": "txt",  "to": "pdf", "cat": "to_pdf"},
    "html2pdf":  {"icon": "🌐", "label": "HTML إلى PDF",               "from": "html", "to": "pdf", "cat": "to_pdf"},
    "xlsx2pdf":  {"icon": "📊", "label": "Excel إلى PDF",              "from": "xlsx", "to": "pdf", "cat": "to_pdf"},
    "xls2pdf":   {"icon": "📊", "label": "Excel 97-2003 إلى PDF",      "from": "xls",  "to": "pdf", "cat": "to_pdf"},
    "ods2pdf":   {"icon": "📊", "label": "ODS إلى PDF",                "from": "ods",  "to": "pdf", "cat": "to_pdf"},
    "csv2pdf":   {"icon": "🧮", "label": "CSV إلى PDF",                "from": "csv",  "to": "pdf", "cat": "to_pdf"},
    "pptx2pdf":  {"icon": "📽️", "label": "PowerPoint إلى PDF",         "from": "pptx", "to": "pdf", "cat": "to_pdf"},
    "ppt2pdf":   {"icon": "📽️", "label": "PowerPoint 97-2003 إلى PDF", "from": "ppt",  "to": "pdf", "cat": "to_pdf"},
    "odp2pdf":   {"icon": "📽️", "label": "ODP إلى PDF",                "from": "odp",  "to": "pdf", "cat": "to_pdf"},
    "images2pdf":{"icon": "🖼️", "label": "الصور إلى PDF",              "from": "images", "to": "pdf", "cat": "to_pdf"},

    "pdf2docx":  {"icon": "📕", "label": "PDF إلى Word",               "from": "pdf", "to": "docx",   "cat": "from_pdf"},
    "pdf2xlsx":  {"icon": "📕", "label": "PDF إلى Excel (استخراج جداول)", "from": "pdf", "to": "xlsx", "cat": "from_pdf"},
    "pdf2images":{"icon": "📕", "label": "PDF إلى صور",                "from": "pdf", "to": "images", "cat": "from_pdf"},
    "pdf2txt":   {"icon": "📕", "label": "PDF إلى نص",                 "from": "pdf", "to": "txt",    "cat": "from_pdf"},
    "pdf2odt":   {"icon": "📕", "label": "PDF إلى ODT",                "from": "pdf", "to": "odt",    "cat": "from_pdf"},
    "pdf2html":  {"icon": "📕", "label": "PDF إلى HTML",               "from": "pdf", "to": "html",   "cat": "from_pdf"},

    "png2jpg":   {"icon": "🖼️", "label": "PNG إلى JPG",  "from": "png",  "to": "jpg", "cat": "images"},
    "jpg2png":   {"icon": "🖼️", "label": "JPG إلى PNG",  "from": "jpg",  "to": "png", "cat": "images"},
    "webp2jpg":  {"icon": "🖼️", "label": "WEBP إلى JPG", "from": "webp", "to": "jpg", "cat": "images"},
    "webp2png":  {"icon": "🖼️", "label": "WEBP إلى PNG", "from": "webp", "to": "png", "cat": "images"},
    "png2webp":  {"icon": "🖼️", "label": "PNG إلى WEBP", "from": "png",  "to": "webp","cat": "images"},
    "bmp2png":   {"icon": "🖼️", "label": "BMP إلى PNG",  "from": "bmp",  "to": "png", "cat": "images"},
    "gif2png":   {"icon": "🖼️", "label": "GIF إلى PNG",  "from": "gif",  "to": "png", "cat": "images"},
    "tiff2png":  {"icon": "🖼️", "label": "TIFF إلى PNG", "from": "tiff", "to": "png", "cat": "images"},
    "image_converter": {"icon": "🎛️", "label": "أداة تحويل الصور الشاملة", "from": None, "to": None, "cat": "images"},
}

SECTIONS = [
    ("to_pdf", "🔄 تحويل إلى PDF"),
    ("from_pdf", "🔄 تحويل من PDF"),
    ("images", "🖼️ تحويل الصور"),
]

POPULAR_TOOLS = ["docx2pdf", "pdf2docx", "images2pdf", "pdf2images", "xlsx2pdf", "pptx2pdf"]

if "active_tool" not in st.session_state:
    st.session_state.active_tool = None

# ============================================================================
# HOME
# ============================================================================
if st.session_state.active_tool is None:
    search = st.text_input("🔍 دوّر على أداة (مثال: Word, PDF, صور...)", value="")
    search_l = search.strip().lower()

    def matches(tid):
        if not search_l:
            return True
        t = TOOLS[tid]
        return search_l in t["label"].lower()

    def render_grid(tool_ids, cols_per_row=4):
        visible = [tid for tid in tool_ids if matches(tid)]
        if not visible:
            return False
        cols = st.columns(cols_per_row)
        for i, tid in enumerate(visible):
            tool = TOOLS[tid]
            with cols[i % cols_per_row]:
                with st.container():
                    st.markdown(f'''<div class="tool-card">
                        <div class="tool-icon">{tool["icon"]}</div>
                        <div class="card-btn">''', unsafe_allow_html=True)
                    if st.button(tool["label"], key=f"card_{tid}", use_container_width=True):
                        st.session_state.active_tool = tid
                        st.rerun()
                    st.markdown('</div></div>', unsafe_allow_html=True)
        return True

    if not search_l:
        st.markdown('<p class="section-title">⭐ الأكثر استخدامًا</p>', unsafe_allow_html=True)
        render_grid(POPULAR_TOOLS)

    any_shown = False
    for cat_key, cat_title in SECTIONS:
        ids = [tid for tid, t in TOOLS.items() if t["cat"] == cat_key]
        st.markdown(f'''<p class="section-title">{cat_title}</p>''', unsafe_allow_html=True)
        shown = render_grid(ids)
        any_shown = any_shown or shown
        if not shown:
            st.caption("لا توجد أدوات مطابقة في هذا القسم.")

# ============================================================================
# TOOL PAGE
# ============================================================================
else:
    tool_id = st.session_state.active_tool
    tool = TOOLS[tool_id]

    st.markdown('<div class="back-btn">', unsafe_allow_html=True)
    if st.button("⬅️ كل الأدوات"):
        st.session_state.active_tool = None
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown(f"## {tool['icon']} {tool['label']}")
    st.markdown("---")

    # ------------------------------------------------------------------
    # أداة تحويل الصور الشاملة
    # ------------------------------------------------------------------
    if tool_id == "image_converter":
        files_up = st.file_uploader(
            "ارفع صورة أو أكتر", type=IMAGE_EXTS, accept_multiple_files=True, key="img_uploader"
        )
        target_fmt = st.selectbox("حوّل إلى صيغة", IMAGE_FORMATS)

        quality = 90
        if target_fmt in ("JPEG", "WEBP"):
            quality = st.slider("الجودة", min_value=10, max_value=100, value=90)

        resize_it = st.checkbox("غيّر حجم الصور (اختياري)")
        new_w = new_h = None
        if resize_it:
            c1, c2 = st.columns(2)
            with c1:
                new_w = st.number_input("العرض (بكسل)", min_value=1, value=800)
            with c2:
                new_h = st.number_input("الطول (بكسل)", min_value=1, value=600)

        if files_up and st.button("🚀 حوّل الصور"):
            try:
                with st.spinner("جاري التحويل..."):
                    out_files = {}
                    ext = target_fmt.lower()
                    if ext == "jpeg":
                        ext = "jpg"
                    for uf in files_up:
                        img = Image.open(uf)
                        if target_fmt in ("JPEG", "BMP") and img.mode in ("RGBA", "P"):
                            img = img.convert("RGB")
                        if resize_it and new_w and new_h:
                            img = img.resize((int(new_w), int(new_h)))
                        buf = io.BytesIO()
                        save_kwargs = {}
                        if target_fmt in ("JPEG", "WEBP"):
                            save_kwargs["quality"] = quality
                        img.save(buf, format=target_fmt, **save_kwargs)
                        out_name = f"{Path(uf.name).stem}.{ext}"
                        out_files[out_name] = buf.getvalue()

                if len(out_files) == 1:
                    (name, data), = out_files.items()
                    st.success("تم التحويل بنجاح ✅")
                    st.download_button("⬇️ تحميل الصورة", data, file_name=name, mime=MIME.get(ext, "application/octet-stream"))
                else:
                    zdata = make_zip(out_files)
                    st.success(f"تم تحويل {len(out_files)} صورة بنجاح ✅")
                    st.download_button(
                        "⬇️ تحميل كل الصور (ZIP)", zdata,
                        file_name="converted_images.zip", mime=MIME["zip"],
                    )

                with st.expander("👁️ معاينة"):
                    cols = st.columns(3)
                    for idx, uf in enumerate(files_up[:9]):
                        uf.seek(0)
                        cols[idx % 3].image(uf, use_container_width=True, caption=uf.name)
            except Exception as e:
                st.error(f"حصل خطأ أثناء التحويل: {e}")

    # ------------------------------------------------------------------
    # كل الأدوات ذات صيغة "من ⟶ إلى" ثابتة
    # ------------------------------------------------------------------
    else:
        from_key, to_key = tool["from"], tool["to"]

        dpi = 150
        img_fmt = "PNG"
        if from_key == "pdf" and to_key == "images":
            dpi = st.slider("جودة الصور (DPI)", min_value=72, max_value=300, value=150, step=6)
            img_fmt = st.selectbox("صيغة الصور", ["PNG", "JPEG"])

        f = None
        files_up = None
        if from_key == "images":
            files_up = st.file_uploader(
                "ارفع صورة أو أكتر (هتترتب زي ما ترفعها)",
                type=IMAGE_EXTS, accept_multiple_files=True, key="doc_img_up",
            )
        else:
            f = st.file_uploader(
                f"ارفع ملف {from_key.upper()}",
                type=UPLOAD_EXTS.get(from_key, [from_key]), key="single_file_up",
            )

        convert_clicked = st.button("🚀 حوّل الملف")

        # ---- Images -> PDF ----
        if from_key == "images" and to_key == "pdf":
            if files_up and convert_clicked:
                try:
                    with st.spinner("جاري الدمج..."):
                        images = [Image.open(uf).convert("RGB") for uf in files_up]
                        buf = io.BytesIO()
                        images[0].save(buf, format="PDF", save_all=True, append_images=images[1:])
                        data = buf.getvalue()
                    st.success("تم الدمج بنجاح ✅")
                    st.download_button(
                        "⬇️ تحميل ملف PDF", data,
                        file_name="merged_images.pdf", mime=MIME["pdf"],
                    )
                except Exception as e:
                    st.error(f"حصل خطأ أثناء الدمج: {e}")
            elif convert_clicked:
                st.warning("ارفع صورة واحدة على الأقل الأول.")

        # ---- PDF -> Word ----
        elif from_key == "pdf" and to_key == "docx":
            if f and convert_clicked:
                if not PDF2DOCX_AVAILABLE:
                    st.error("❌ مكتبة pdf2docx مش متوفرة. نزّلها بالأمر: pip install pdf2docx")
                else:
                    with tempfile.TemporaryDirectory() as td:
                        in_path = os.path.join(td, f.name)
                        with open(in_path, "wb") as out:
                            out.write(f.getbuffer())
                        out_path = os.path.join(td, Path(f.name).stem + ".docx")
                        try:
                            with st.spinner("جاري التحويل... (بياخد وقت أطول شوية)"):
                                cv = Converter(in_path)
                                cv.convert(out_path)
                                cv.close()
                            with open(out_path, "rb") as r:
                                data = r.read()
                            st.success("تم التحويل بنجاح ✅")
                            st.download_button(
                                "⬇️ تحميل ملف Word", data,
                                file_name=Path(f.name).stem + ".docx", mime=MIME["docx"],
                            )
                        except Exception as e:
                            st.error(f"حصل خطأ أثناء التحويل: {e}")
            elif convert_clicked:
                st.warning("ارفع ملف PDF الأول.")

        # ---- PDF -> Excel ----
        elif from_key == "pdf" and to_key == "xlsx":
            if f and convert_clicked:
                with tempfile.TemporaryDirectory() as td:
                    in_path = os.path.join(td, f.name)
                    with open(in_path, "wb") as out:
                        out.write(f.getbuffer())
                    try:
                        with st.spinner("جاري استخراج الجداول..."):
                            out_path = os.path.join(td, Path(f.name).stem + ".xlsx")
                            found_any = False
                            with pdfplumber.open(in_path) as pdf, pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                                for i, page in enumerate(pdf.pages, start=1):
                                    tables = page.extract_tables()
                                    for j, table in enumerate(tables, start=1):
                                        if not table or len(table) < 1:
                                            continue
                                        df = pd.DataFrame(table[1:], columns=table[0])
                                        sheet_name = f"page{i}_t{j}"[:31]
                                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                                        found_any = True
                        if not found_any:
                            st.warning(
                                "معرفش ألاقي جداول واضحة في الملف ده — ممكن يكون الجدول عبارة عن صورة أو خطوط غير منتظمة."
                            )
                        else:
                            with open(out_path, "rb") as r:
                                data = r.read()
                            st.success("تم استخراج الجداول بنجاح ✅")
                            st.download_button(
                                "⬇️ تحميل ملف Excel", data,
                                file_name=Path(f.name).stem + ".xlsx", mime=MIME["xlsx"],
                            )
                    except Exception as e:
                        st.error(f"حصل خطأ أثناء التحويل: {e}")
            elif convert_clicked:
                st.warning("ارفع ملف PDF الأول.")

        # ---- PDF -> Images ----
        elif from_key == "pdf" and to_key == "images":
            if f and convert_clicked:
                with tempfile.TemporaryDirectory() as td:
                    in_path = os.path.join(td, f.name)
                    with open(in_path, "wb") as out:
                        out.write(f.getbuffer())
                    try:
                        with st.spinner("جاري تحويل الصفحات..."):
                            doc = fitz.open(in_path)
                            out_files = {}
                            for i, page in enumerate(doc, start=1):
                                pix = page.get_pixmap(dpi=dpi)
                                ext = "png" if img_fmt == "PNG" else "jpg"
                                fname = f"{Path(f.name).stem}_page{i}.{ext}"
                                out_files[fname] = pix.tobytes(ext if ext != "jpg" else "jpeg")
                            doc.close()
                        if len(out_files) == 1:
                            (name, data), = out_files.items()
                            st.success("تم التحويل بنجاح ✅")
                            st.download_button("⬇️ تحميل الصورة", data, file_name=name, mime=MIME[ext])
                        else:
                            zdata = make_zip(out_files)
                            st.success(f"تم تحويل {len(out_files)} صفحة بنجاح ✅")
                            st.download_button(
                                "⬇️ تحميل كل الصور (ZIP)", zdata,
                                file_name=Path(f.name).stem + "_pages.zip", mime=MIME["zip"],
                            )
                    except Exception as e:
                        st.error(f"حصل خطأ أثناء التحويل: {e}")
            elif convert_clicked:
                st.warning("ارفع ملف PDF الأول.")

        # ---- PDF -> Text ----
        elif from_key == "pdf" and to_key == "txt":
            if f and convert_clicked:
                with tempfile.TemporaryDirectory() as td:
                    in_path = os.path.join(td, f.name)
                    with open(in_path, "wb") as out:
                        out.write(f.getbuffer())
                    try:
                        with st.spinner("جاري استخراج النص..."):
                            doc = fitz.open(in_path)
                            text = "\n\n".join(page.get_text() for page in doc)
                            doc.close()
                        data = text.encode("utf-8")
                        st.success("تم الاستخراج بنجاح ✅")
                        st.download_button(
                            "⬇️ تحميل ملف نصي", data,
                            file_name=Path(f.name).stem + ".txt", mime=MIME["txt"],
                        )
                    except Exception as e:
                        st.error(f"حصل خطأ أثناء الاستخراج: {e}")
            elif convert_clicked:
                st.warning("ارفع ملف PDF الأول.")

        # ---- Image -> Image ----
        elif from_key in IMAGE_EXT_SET and to_key in IMAGE_EXT_SET:
            if f and convert_clicked:
                try:
                    with st.spinner("جاري التحويل..."):
                        img = Image.open(f)
                        target_pil_fmt = PIL_FORMAT[to_key]
                        if target_pil_fmt in ("JPEG", "BMP") and img.mode in ("RGBA", "P"):
                            img = img.convert("RGB")
                        buf = io.BytesIO()
                        save_kwargs = {"quality": 90} if target_pil_fmt in ("JPEG", "WEBP") else {}
                        img.save(buf, format=target_pil_fmt, **save_kwargs)
                        data = buf.getvalue()
                    out_ext = "jpg" if to_key == "jpeg" else to_key
                    st.success("تم التحويل بنجاح ✅")
                    st.download_button(
                        "⬇️ تحميل الصورة", data,
                        file_name=Path(f.name).stem + "." + out_ext,
                        mime=MIME.get(out_ext, "application/octet-stream"),
                    )
                except Exception as e:
                    st.error(f"حصل خطأ أثناء التحويل: {e}")
            elif convert_clicked:
                st.warning("ارفع صورة الأول.")

        # ---- General LibreOffice conversions ----
        else:
            if f and convert_clicked:
                with tempfile.TemporaryDirectory() as td:
                    in_path = os.path.join(td, f.name)
                    with open(in_path, "wb") as out:
                        out.write(f.getbuffer())
                    try:
                        with st.spinner("جاري التحويل..."):
                            out_path = soffice_convert(in_path, to_key, td)
                        with open(out_path, "rb") as r:
                            data = r.read()
                        out_name = Path(f.name).stem + "." + to_key
                        st.success("تم التحويل بنجاح ✅")
                        st.download_button(
                            "⬇️ تحميل الملف", data,
                            file_name=out_name, mime=MIME.get(to_key, "application/octet-stream"),
                        )
                    except Exception as e:
                        st.error(f"حصل خطأ أثناء التحويل: {e}")
            elif convert_clicked:
                st.warning("ارفع الملف الأول.")

st.markdown(
    """
    <hr style="border: none; height: 1px; background: linear-gradient(90deg, transparent, rgba(0,240,255,0.4), rgba(255,0,127,0.4), transparent); margin: 2rem 0 1rem;">
    <p style="text-align:center; color:#8b85a8; font-size:0.85rem;">
        تصميم وتطوير: <span style="color:#00f0ff;">المهندس رفيق ناثان</span>
    </p>
    """,
    unsafe_allow_html=True,
)
