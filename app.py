import streamlit as st
import subprocess
import tempfile
import os
import zipfile
import io
from pathlib import Path

from PIL import Image
import fitz  # PyMuPDF
from pdf2docx import Converter
import pdfplumber
import pandas as pd

# ----------------------------------------------------------------------------
# Page config
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="محوّل الملفات | File Converter",
    page_icon="🔄",
    layout="centered",
)

# ----------------------------------------------------------------------------
# Styling — RamboAITV space theme: dark radial gradient, neon pink/cyan,
# Orbitron for headings + Cairo for Arabic body text
# ----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@600;800;900&family=Cairo:wght@400;600;800&display=swap');

    :root {
        --neon-pink: #ff007f;
        --neon-cyan: #00f0ff;
    }

    html, body, [class*="css"] {
        font-family: 'Cairo', sans-serif;
    }

    /* Deep space background */
    .stApp {
        background: radial-gradient(ellipse at top, #1a0b2e 0%, #0d0518 45%, #050208 100%);
        background-attachment: fixed;
    }

    .main .block-container {
        direction: rtl;
    }
    .stTabs [data-baseweb="tab-list"] {
        direction: rtl;
    }
    h1, h2, h3, p, label, .stMarkdown {
        text-align: right;
        color: #e8e6f0;
    }

    /* Glowing hero title */
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
        margin-bottom: 1.8rem;
        font-size: 1rem;
    }

    /* Tabs styled as neon cards */
    .stTabs [data-baseweb="tab"] {
        background: rgba(255, 255, 255, 0.03);
        border: 1px solid rgba(0, 240, 255, 0.25);
        border-radius: 12px 12px 0 0;
        color: #e8e6f0 !important;
        font-weight: 700;
        padding: 0.5rem 1.2rem;
    }
    .stTabs [aria-selected="true"] {
        border: 1px solid var(--neon-cyan);
        box-shadow: 0 0 14px rgba(0, 240, 255, 0.45);
        color: var(--neon-cyan) !important;
    }

    /* Card-like containers */
    div[data-testid="stVerticalBlockBorderWrapper"],
    .stFileUploader, .stSelectbox, .stSlider {
        border-radius: 14px;
    }

    /* Buttons — neon gradient with glow */
    .stButton>button, .stDownloadButton>button {
        width: 100%;
        border-radius: 12px;
        font-weight: 800;
        padding: 0.65rem;
        border: 1px solid rgba(0, 240, 255, 0.4);
        background: linear-gradient(90deg, rgba(255,0,127,0.15), rgba(0,240,255,0.15));
        color: #f5f4ff;
        transition: all 0.25s ease;
    }
    .stButton>button:hover, .stDownloadButton>button:hover {
        border: 1px solid var(--neon-pink);
        box-shadow: 0 0 18px rgba(255, 0, 127, 0.5), 0 0 18px rgba(0, 240, 255, 0.3);
        color: #ffffff;
        transform: translateY(-1px);
    }

    /* File uploader glow border */
    [data-testid="stFileUploaderDropzone"] {
        background: rgba(255, 255, 255, 0.02);
        border: 1.5px dashed rgba(0, 240, 255, 0.4) !important;
        border-radius: 14px;
    }

    footer {visibility: hidden;}
    #MainMenu {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<p class="converter-title">🔄 محوّل الملفات</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="converter-subtitle">حوّل المستندات والصور بين الصيغ المختلفة بضغطة زرار</p>',
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------

def soffice_convert(input_path: str, target_format: str, outdir: str) -> str:
    """Convert a file using headless LibreOffice. Returns output file path."""
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

MIME = {
    "pdf": "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
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
# Tabs
# ----------------------------------------------------------------------------
tab_docs, tab_images = st.tabs(["📄 تحويل المستندات", "🖼️ تحويل الصور"])

# ============================================================================
# TAB 1 — DOCUMENTS
# ============================================================================
with tab_docs:
    doc_options = {
        "PDF ← Word (.docx)": "word2pdf",
        "Word (.docx) ← PDF": "pdf2word",
        "PDF ← Excel (.xlsx)": "excel2pdf",
        "استخراج جداول من PDF إلى Excel": "pdf2excel",
        "PDF ← PowerPoint (.pptx)": "ppt2pdf",
        "صور ← PDF (كل صفحة صورة)": "pdf2images",
        "PDF ← صور (دمج صور في ملف واحد)": "images2pdf_doc",
    }
    choice_label = st.selectbox("اختار نوع التحويل", list(doc_options.keys()))
    choice = doc_options[choice_label]

    # ---- Word -> PDF ----
    if choice == "word2pdf":
        f = st.file_uploader("ارفع ملف Word (.docx / .doc)", type=["docx", "doc"])
        if f and st.button("حوّل إلى PDF"):
            with tempfile.TemporaryDirectory() as td:
                in_path = os.path.join(td, f.name)
                with open(in_path, "wb") as out:
                    out.write(f.getbuffer())
                try:
                    with st.spinner("جاري التحويل..."):
                        out_path = soffice_convert(in_path, "pdf", td)
                    with open(out_path, "rb") as r:
                        data = r.read()
                    st.success("تم التحويل بنجاح ✅")
                    st.download_button(
                        "⬇️ تحميل ملف PDF", data,
                        file_name=Path(f.name).stem + ".pdf", mime=MIME["pdf"],
                    )
                except Exception as e:
                    st.error(f"حصل خطأ أثناء التحويل: {e}")

    # ---- PDF -> Word ----
    elif choice == "pdf2word":
        f = st.file_uploader("ارفع ملف PDF", type=["pdf"])
        if f and st.button("حوّل إلى Word"):
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

    # ---- Excel -> PDF ----
    elif choice == "excel2pdf":
        f = st.file_uploader("ارفع ملف Excel (.xlsx / .xls)", type=["xlsx", "xls"])
        if f and st.button("حوّل إلى PDF"):
            with tempfile.TemporaryDirectory() as td:
                in_path = os.path.join(td, f.name)
                with open(in_path, "wb") as out:
                    out.write(f.getbuffer())
                try:
                    with st.spinner("جاري التحويل..."):
                        out_path = soffice_convert(in_path, "pdf", td)
                    with open(out_path, "rb") as r:
                        data = r.read()
                    st.success("تم التحويل بنجاح ✅")
                    st.download_button(
                        "⬇️ تحميل ملف PDF", data,
                        file_name=Path(f.name).stem + ".pdf", mime=MIME["pdf"],
                    )
                except Exception as e:
                    st.error(f"حصل خطأ أثناء التحويل: {e}")

    # ---- PDF -> Excel (extract tables) ----
    elif choice == "pdf2excel":
        f = st.file_uploader("ارفع ملف PDF فيه جداول", type=["pdf"])
        if f and st.button("استخرج الجداول إلى Excel"):
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

    # ---- PowerPoint -> PDF ----
    elif choice == "ppt2pdf":
        f = st.file_uploader("ارفع ملف PowerPoint (.pptx / .ppt)", type=["pptx", "ppt"])
        if f and st.button("حوّل إلى PDF"):
            with tempfile.TemporaryDirectory() as td:
                in_path = os.path.join(td, f.name)
                with open(in_path, "wb") as out:
                    out.write(f.getbuffer())
                try:
                    with st.spinner("جاري التحويل..."):
                        out_path = soffice_convert(in_path, "pdf", td)
                    with open(out_path, "rb") as r:
                        data = r.read()
                    st.success("تم التحويل بنجاح ✅")
                    st.download_button(
                        "⬇️ تحميل ملف PDF", data,
                        file_name=Path(f.name).stem + ".pdf", mime=MIME["pdf"],
                    )
                except Exception as e:
                    st.error(f"حصل خطأ أثناء التحويل: {e}")

    # ---- PDF -> Images ----
    elif choice == "pdf2images":
        f = st.file_uploader("ارفع ملف PDF", type=["pdf"])
        dpi = st.slider("جودة الصور (DPI)", min_value=72, max_value=300, value=150, step=6)
        img_fmt = st.selectbox("صيغة الصور", ["PNG", "JPEG"])
        if f and st.button("حوّل الصفحات لصور"):
            with tempfile.TemporaryDirectory() as td:
                in_path = os.path.join(td, f.name)
                with open(in_path, "wb") as out:
                    out.write(f.getbuffer())
                try:
                    with st.spinner("جاري تحويل الصفحات..."):
                        doc = fitz.open(in_path)
                        files = {}
                        for i, page in enumerate(doc, start=1):
                            pix = page.get_pixmap(dpi=dpi)
                            ext = "png" if img_fmt == "PNG" else "jpg"
                            fname = f"{Path(f.name).stem}_page{i}.{ext}"
                            files[fname] = pix.tobytes(ext if ext != "jpg" else "jpeg")
                        doc.close()
                    if len(files) == 1:
                        (name, data), = files.items()
                        st.success("تم التحويل بنجاح ✅")
                        st.download_button("⬇️ تحميل الصورة", data, file_name=name, mime=MIME[ext])
                    else:
                        zdata = make_zip(files)
                        st.success(f"تم تحويل {len(files)} صفحة بنجاح ✅")
                        st.download_button(
                            "⬇️ تحميل كل الصور (ZIP)", zdata,
                            file_name=Path(f.name).stem + "_pages.zip", mime=MIME["zip"],
                        )
                except Exception as e:
                    st.error(f"حصل خطأ أثناء التحويل: {e}")

    # ---- Images -> PDF ----
    elif choice == "images2pdf_doc":
        files_up = st.file_uploader(
            "ارفع صورة أو أكتر (هتترتب زي ما ترفعها)",
            type=IMAGE_EXTS, accept_multiple_files=True,
        )
        if files_up and st.button("ادمج الصور في PDF"):
            try:
                with st.spinner("جاري الدمج..."):
                    images = []
                    for uf in files_up:
                        img = Image.open(uf).convert("RGB")
                        images.append(img)
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

# ============================================================================
# TAB 2 — IMAGES
# ============================================================================
with tab_images:
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

    if files_up and st.button("حوّل الصور"):
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

            # Preview
            with st.expander("👁️ معاينة"):
                cols = st.columns(3)
                for idx, uf in enumerate(files_up[:9]):
                    uf.seek(0)
                    cols[idx % 3].image(uf, use_container_width=True, caption=uf.name)
        except Exception as e:
            st.error(f"حصل خطأ أثناء التحويل: {e}")

st.markdown(
    """
    <hr style="border: none; height: 1px; background: linear-gradient(90deg, transparent, rgba(0,240,255,0.4), rgba(255,0,127,0.4), transparent); margin: 2rem 0 1rem;">
    <p style="text-align:center; color:#8b85a8; font-size:0.85rem;">
        تصميم وتطوير: <span style="color:#00f0ff;">المهندس رفيق ناثان</span>
    </p>
    """,
    unsafe_allow_html=True,
)
