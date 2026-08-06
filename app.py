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
# Styling — RamboAITV space theme (base dark theme comes from .streamlit/config.toml
# so every widget — buttons, dropdowns, uploader — is visible by default; this CSS
# only adds the neon glow / fonts on top)
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

    .stApp {
        background: radial-gradient(ellipse at top, #1a0b2e 0%, #0d0518 45%, #050208 100%);
        background-attachment: fixed;
    }

    .main .block-container {
        direction: rtl;
        max-width: 780px;
    }
    h1, h2, h3, p, label, .stMarkdown {
        text-align: right;
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

    /* Always-visible solid buttons — no reliance on hover/focus to show them */
    .stButton>button, .stDownloadButton>button {
        width: 100%;
        border-radius: 12px;
        font-weight: 800;
        padding: 0.7rem;
        border: none;
        background: linear-gradient(90deg, var(--neon-pink), var(--neon-cyan)) !important;
        color: #0a0a12 !important;
        box-shadow: 0 0 10px rgba(255, 0, 127, 0.35);
        transition: all 0.2s ease;
    }
    .stButton>button:hover, .stDownloadButton>button:hover {
        box-shadow: 0 0 22px rgba(255, 0, 127, 0.6), 0 0 22px rgba(0, 240, 255, 0.4);
        transform: translateY(-1px);
    }
    .stButton>button p, .stDownloadButton>button p { color: #0a0a12 !important; font-weight: 800; }

    /* From/To row */
    .conv-arrow {
        text-align: center;
        font-size: 1.8rem;
        color: var(--neon-cyan);
        text-shadow: 0 0 12px rgba(0, 240, 255, 0.6);
        padding-top: 2.1rem;
    }

    [data-testid="stFileUploaderDropzone"] {
        border: 1.5px dashed rgba(0, 240, 255, 0.45) !important;
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
    '<p class="converter-subtitle">اختار تحوّل من إيه لإيه، وارفع الملف</p>',
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------
# Format registry
# ----------------------------------------------------------------------------
FORMATS = {
    "PDF": "pdf",
    "Word (DOCX)": "docx",
    "Excel (XLSX)": "xlsx",
    "PowerPoint (PPTX)": "pptx",
    "نص (TXT)": "txt",
    "جدول (CSV)": "csv",
    "PNG": "png",
    "JPG / JPEG": "jpg",
    "WEBP": "webp",
    "BMP": "bmp",
    "GIF": "gif",
    "TIFF": "tiff",
    "ICO": "ico",
}
EXT_TO_NAME = {v: k for k, v in FORMATS.items()}
IMAGE_EXT_SET = {"png", "jpg", "webp", "bmp", "gif", "tiff", "ico"}
UPLOAD_TYPE_ALIASES = {"jpg": ["jpg", "jpeg"], "tiff": ["tiff", "tif"]}

TARGETS = {
    "pdf": ["docx", "xlsx", "txt", "png", "jpg"],
    "docx": ["pdf", "txt"],
    "xlsx": ["pdf", "csv"],
    "pptx": ["pdf"],
    "txt": ["pdf", "docx"],
    "csv": ["xlsx", "pdf"],
}
for _ext in IMAGE_EXT_SET:
    TARGETS[_ext] = ["pdf"] + sorted(e for e in IMAGE_EXT_SET if e != _ext)

MIME = {
    "pdf": "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "txt": "text/plain",
    "csv": "text/csv",
    "png": "image/png",
    "jpg": "image/jpeg",
    "webp": "image/webp",
    "bmp": "image/bmp",
    "gif": "image/gif",
    "tiff": "image/tiff",
    "ico": "image/x-icon",
    "zip": "application/zip",
}

# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------

def soffice_convert(input_path: str, target_ext: str, outdir: str) -> str:
    cmd = ["soffice", "--headless", "--norestore", "--convert-to", target_ext, "--outdir", outdir, input_path]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    stem = Path(input_path).stem
    out_path = os.path.join(outdir, f"{stem}.{target_ext}")
    if not os.path.exists(out_path):
        raise RuntimeError(f"فشل التحويل. {result.stderr[:300]}")
    return out_path


def make_zip(files: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)
    buf.seek(0)
    return buf.read()


def read_csv_smart(path: str) -> pd.DataFrame:
    for enc in ("utf-8-sig", "utf-8", "cp1256", "latin1"):
        try:
            return pd.read_csv(path, encoding=enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    return pd.read_csv(path, encoding="latin1", errors="replace")


def do_convert(from_ext: str, to_ext: str, files) -> dict:
    """Returns {filename: bytes}"""
    with tempfile.TemporaryDirectory() as td:
        saved_paths = []
        for uf in files:
            p = os.path.join(td, uf.name)
            with open(p, "wb") as out:
                out.write(uf.getbuffer())
            saved_paths.append(p)

        # ---------------- PDF as source ----------------
        if from_ext == "pdf":
            in_path = saved_paths[0]
            stem = Path(in_path).stem
            if to_ext == "docx":
                out_path = os.path.join(td, stem + ".docx")
                cv = Converter(in_path)
                cv.convert(out_path)
                cv.close()
                return {stem + ".docx": open(out_path, "rb").read()}
            if to_ext == "xlsx":
                out_path = os.path.join(td, stem + ".xlsx")
                found = False
                with pdfplumber.open(in_path) as pdf, pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                    for i, page in enumerate(pdf.pages, start=1):
                        for j, table in enumerate(page.extract_tables(), start=1):
                            if not table:
                                continue
                            df = pd.DataFrame(table[1:], columns=table[0])
                            df.to_excel(writer, sheet_name=f"page{i}_t{j}"[:31], index=False)
                            found = True
                if not found:
                    raise RuntimeError("معرفش ألاقي جداول واضحة في الـ PDF ده.")
                return {stem + ".xlsx": open(out_path, "rb").read()}
            if to_ext == "txt":
                doc = fitz.open(in_path)
                text = "\n\n".join(page.get_text() for page in doc)
                doc.close()
                return {stem + ".txt": text.encode("utf-8")}
            if to_ext in ("png", "jpg"):
                doc = fitz.open(in_path)
                out_files = {}
                for i, page in enumerate(doc, start=1):
                    pix = page.get_pixmap(dpi=180)
                    fmt = "png" if to_ext == "png" else "jpeg"
                    out_files[f"{stem}_page{i}.{to_ext}"] = pix.tobytes(fmt)
                doc.close()
                return out_files

        # ---------------- Images as source ----------------
        if from_ext in IMAGE_EXT_SET:
            if to_ext == "pdf":
                images = [Image.open(p).convert("RGB") for p in saved_paths]
                buf = io.BytesIO()
                images[0].save(buf, format="PDF", save_all=True, append_images=images[1:])
                return {"converted.pdf": buf.getvalue()}
            else:
                out_files = {}
                pil_fmt = "JPEG" if to_ext == "jpg" else to_ext.upper()
                for p in saved_paths:
                    img = Image.open(p)
                    if pil_fmt in ("JPEG", "BMP") and img.mode in ("RGBA", "P"):
                        img = img.convert("RGB")
                    buf = io.BytesIO()
                    img.save(buf, format=pil_fmt)
                    out_files[f"{Path(p).stem}.{to_ext}"] = buf.getvalue()
                return out_files

        # ---------------- CSV <-> XLSX (pandas, Arabic-safe) ----------------
        if from_ext == "csv" and to_ext == "xlsx":
            in_path = saved_paths[0]
            stem = Path(in_path).stem
            df = read_csv_smart(in_path)
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                df.to_excel(writer, index=False)
            return {stem + ".xlsx": buf.getvalue()}

        if from_ext == "xlsx" and to_ext == "csv":
            in_path = saved_paths[0]
            stem = Path(in_path).stem
            df = pd.read_excel(in_path, engine="openpyxl")
            buf = io.BytesIO()
            df.to_csv(buf, index=False, encoding="utf-8-sig")
            return {stem + ".csv": buf.getvalue()}

        # ---------------- Everything else -> LibreOffice generic ----------------
        in_path = saved_paths[0]
        stem = Path(in_path).stem
        out_path = soffice_convert(in_path, to_ext, td)
        return {stem + "." + to_ext: open(out_path, "rb").read()}


# ----------------------------------------------------------------------------
# UI — From / Arrow / To
# ----------------------------------------------------------------------------
col_from, col_arrow, col_to = st.columns([5, 1, 5])

with col_from:
    from_name = st.selectbox("حوّل من:", list(FORMATS.keys()), key="from_fmt")
from_ext = FORMATS[from_name]

valid_to_exts = TARGETS.get(from_ext, [])
valid_to_names = [EXT_TO_NAME[e] for e in valid_to_exts]

with col_arrow:
    st.markdown('<div class="conv-arrow">⬅</div>', unsafe_allow_html=True)

with col_to:
    if valid_to_names:
        to_name = st.selectbox("حوّل إلى:", valid_to_names, key=f"to_fmt_{from_ext}")
        to_ext = FORMATS[to_name]
    else:
        st.selectbox("حوّل إلى:", ["لا يوجد تحويل متاح"], disabled=True)
        to_ext = None

st.markdown("")

allow_multi = from_ext in IMAGE_EXT_SET
upload_types = UPLOAD_TYPE_ALIASES.get(from_ext, [from_ext])

uploaded = st.file_uploader(
    f"ارفع ملف {from_name}" + (" (تقدر ترفع أكتر من ملف)" if allow_multi else ""),
    type=upload_types,
    accept_multiple_files=allow_multi,
    key=f"uploader_{from_ext}_{to_ext}",
)
files = uploaded if isinstance(uploaded, list) else ([uploaded] if uploaded else [])

if to_ext and files and st.button(f"🔄 حوّل إلى {to_name}"):
    try:
        with st.spinner("جاري التحويل..."):
            result_files = do_convert(from_ext, to_ext, files)
        st.success("تم التحويل بنجاح ✅")
        if len(result_files) == 1:
            (name, data), = result_files.items()
            st.download_button("⬇️ تحميل الملف", data, file_name=name, mime=MIME.get(to_ext, "application/octet-stream"))
        else:
            zdata = make_zip(result_files)
            st.download_button(f"⬇️ تحميل كل الملفات ({len(result_files)}) — ZIP", zdata, file_name="converted_files.zip", mime=MIME["zip"])
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
