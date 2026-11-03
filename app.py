
from pathlib import Path
import re
import hashlib
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd
import numpy as np
import plotly.express as px
import streamlit as st
from PIL import Image


st.set_page_config(
    page_title="Desaku Maju | Agricultural Intelligence",
    page_icon="🌾",
    layout="wide",
)


# --------------------------------------------------
# KONFIGURASI
# --------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
LOGO_PATH = BASE_DIR / "logo.png"
BANNER_PATH = BASE_DIR / "banner.png"

# Temukan workbook Excel di folder aplikasi. Hindari file sementara Excel (~$...).
excel_files = sorted(
    [
        p for p in BASE_DIR.iterdir()
        if p.is_file()
        and p.suffix.lower() in {".xlsx", ".xlsm", ".xls"}
        and not p.name.startswith("~$")
    ],
    key=lambda p: p.name.casefold(),
)



# --------------------------------------------------
# KNOWLEDGE CENTER: DOKUMEN SOP / PANDUAN
# --------------------------------------------------
KB_DIR = BASE_DIR / "knowledge_docs"
KB_DIR.mkdir(exist_ok=True)

def extract_document_text(path: Path) -> str:
    """Ekstrak teks dari DOCX, PDF, TXT, atau Markdown secara aman."""
    suffix = path.suffix.lower()
    try:
        if suffix == ".docx":
            # Preferred parser: python-docx. If missing, use Python's standard library.
            try:
                from docx import Document
                doc = Document(str(path))
                parts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
                for table in doc.tables:
                    for row in table.rows:
                        cells = [cell.text.strip() for cell in row.cells]
                        if any(cells):
                            parts.append(" | ".join(cells))
                return "\n".join(parts)
            except ImportError:
                with zipfile.ZipFile(path, "r") as archive:
                    xml_bytes = archive.read("word/document.xml")
                root = ET.fromstring(xml_bytes)
                ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                parts = []
                # Paragraphs inside tables are included in document order.
                for paragraph in root.findall(".//w:p", ns):
                    fragments = []
                    for node in paragraph.iter():
                        if node.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t":
                            fragments.append(node.text or "")
                        elif node.tag.endswith("}tab") or node.tag.endswith("}br"):
                            fragments.append(" ")
                    line = "".join(fragments).strip()
                    if line:
                        parts.append(line)
                return "\n".join(parts)
        if suffix == ".pdf":
            try:
                from pypdf import PdfReader
            except ImportError:
                from PyPDF2 import PdfReader
            reader = PdfReader(str(path))
            return "\n".join((page.extract_text() or "") for page in reader.pages)
        if suffix in {".txt", ".md"}:
            return path.read_text(encoding="utf-8", errors="ignore")
    except Exception as exc:
        return f"__ERROR__:{exc}"
    return ""

def split_into_chunks(text: str, chunk_size: int = 900, overlap: int = 150):
    text = re.sub(r"[ \t]+", " ", text or "")
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    if not text:
        return []
    chunks = []
    start = 0
    while start < len(text):
        end = min(len(text), start + chunk_size)
        if end < len(text):
            # Prefer cutting at a sentence or paragraph boundary.
            cut = max(text.rfind("\n", start, end), text.rfind(". ", start, end))
            if cut > start + int(chunk_size * 0.55):
                end = cut + 1
        chunk = text[start:end].strip()
        if chunk:
            chunks.append(chunk)
        if end >= len(text):
            break
        start = max(start + 1, end - overlap)
    return chunks

def load_knowledge_documents():
    """Baca dokumen dari knowledge_docs dan dokumen SOP yang diletakkan di folder app.py."""
    documents = []
    supported = {".docx", ".pdf", ".txt", ".md"}
    candidates = {}
    # Scan folder Knowledge Center plus folder utama agar SOP yang disalin di samping app.py
    # langsung terdeteksi. File Excel dan tipe lain otomatis diabaikan.
    for folder in (KB_DIR, BASE_DIR):
        try:
            for candidate in folder.iterdir():
                if (
                    candidate.is_file()
                    and candidate.suffix.lower() in supported
                    and not candidate.name.casefold().startswith("readme")
                    and "petunjuk_pemasangan" not in candidate.name.casefold()
                ):
                    candidates[str(candidate.resolve())] = candidate
        except OSError:
            pass
    for doc_path in sorted(candidates.values(), key=lambda p: p.name.casefold()):
        extracted = extract_document_text(doc_path)
        if extracted.startswith("__ERROR__:"):
            documents.append({
                "name": doc_path.name, "text": "", "chunks": [],
                "error": extracted.replace("__ERROR__:", "")
            })
            continue
        chunks = split_into_chunks(extracted)
        documents.append({
            "name": doc_path.name, "text": extracted, "chunks": chunks,
            "error": "" if chunks else "Tidak ada teks yang dapat diekstrak."
        })
    return documents

def retrieve_relevant_chunks(question: str, documents, top_k: int = 1):
    """
    Pencarian demo berbasis kalimat: perluasan istilah sesuai maksud pertanyaan,
    buang kalimat umum/daftar referensi, lalu gabungkan maksimal 3 kalimat relevan
    menjadi satu jawaban ringkas dari satu dokumen.
    """
    stopwords = {
        "yang", "dan", "atau", "untuk", "dengan", "dari", "pada", "dalam",
        "apa", "bagaimana", "saya", "kita", "itu", "ini", "agar", "bisa",
        "perlu", "apakah", "saat", "setelah", "sebelum", "dapat", "oleh",
        "the", "and", "for", "with", "from", "what", "how", "perhatikan",
        "diperhatikan"
    }
    q = question.casefold()
    terms = {
        t for t in re.findall(r"[a-zA-ZÀ-ÿ0-9]+", q)
        if len(t) >= 3 and t not in stopwords
    }

    # Query expansion for common farmer questions. This is still keyword search,
    # but handles wording differences such as "setelah panen" vs "pascapanen".
    expansion = set()
    if any(k in q for k in ["setelah panen", "pascapanen", "pasca panen", "habis panen"]):
        expansion.update({
            "pascapanen", "gabah", "pengeringan", "kadar air", "penyimpanan",
            "lembap", "lembap", "panas", "menumpuk", "kehilangan hasil",
            "berat hasil", "evaluasi hasil", "kerusakan mekanis", "wadah bersih"
        })
    if any(k in q for k in ["pupuk", "pemupukan", "memupuk"]):
        expansion.update({"pupuk", "pemupukan", "dosis", "tanah", "berimbang", "rekomendasi", "label"})
    if any(k in q for k in ["hama", "penyakit", "opt"]):
        expansion.update({"hama", "penyakit", "opt", "pengamatan", "pengendalian", "pestisida", "label"})
    if any(k in q for k in ["perawatan", "merawat", "pemeliharaan"]):
        expansion.update({"perawatan", "pemeliharaan", "pengairan", "gulma", "hama", "pengamatan"})

    all_terms = terms | expansion
    candidates = []
    sentence_splitter = re.compile(r"(?<=[.!?])\s+")
    for doc in documents:
        for idx, chunk in enumerate(doc.get("chunks", [])):
            lower_chunk = chunk.casefold()
            if "sumber referensi" in lower_chunk or "https://" in lower_chunk or "http://" in lower_chunk:
                continue
            # Score individual sentences, not entire large chunks.
            for sentence in sentence_splitter.split(re.sub(r"\s+", " ", chunk).strip()):
                sentence = sentence.strip(" -•\t")
                lower = sentence.casefold()
                if len(sentence) < 35:
                    continue
                if any(skip in lower for skip in [
                    "dokumen ini disusun sebagai bahan dummy",
                    "ruang lingkup sop contoh",
                    "sumber referensi",
                    "belum disahkan oleh"
                ]):
                    continue
                matched_query = sum(1 for term in terms if term in lower)
                matched_expansion = sum(1 for term in expansion if term in lower)
                if matched_query == 0 and matched_expansion == 0:
                    continue
                # User's explicit topic terms are weighted more heavily than expansions.
                score = matched_query * 5 + matched_expansion * 2
                if "padi" in terms and ("padi" in lower or "gabah" in lower):
                    score += 3
                if any(k in q for k in ["setelah panen", "pascapanen", "pasca panen", "habis panen"]):
                    if any(k in lower for k in ["pengeringan", "kadar air", "penyimpanan", "lembap", "menumpuk", "berat hasil", "evaluasi hasil"]):
                        score += 8
                candidates.append({
                    "score": score, "filename": doc["name"], "chunk_index": idx,
                    "sentence": sentence, "lower": lower
                })

    if not candidates:
        return []

    candidates.sort(key=lambda item: item["score"], reverse=True)
    best = candidates[0]
    # Combine up to three highly relevant, non-duplicate sentences from the same document.
    # This is presented as one concise answer rather than multiple separate answers.
    selected = [best]
    for candidate in candidates[1:]:
        if candidate["filename"] != best["filename"]:
            continue
        if candidate["sentence"] in [item["sentence"] for item in selected]:
            continue
        if candidate["score"] < max(3, best["score"] * 0.45):
            continue
        selected.append(candidate)
        if len(selected) >= 3:
            break

    answer = " ".join(item["sentence"] for item in selected)
    return [(best["score"], best["filename"], best["chunk_index"], answer)]

def load_workbook(path_string, modified_time):
    # modified_time menjadi bagian cache agar perubahan file
    # dapat memicu pembacaan ulang.
    return pd.read_excel(path_string, sheet_name=None)


@st.cache_data(show_spinner="Membaca sheet Excel...")
def load_sheet_for_display(path_string, sheet_name, modified_time, detect_farmer_header=False):
    """Baca sheet dan deteksi header tabel petani yang berada di bawah judul."""
    if not detect_farmer_header:
        df = pd.read_excel(path_string, sheet_name=sheet_name)
        df.columns = [str(c).strip() for c in df.columns]
        return df, None

    raw_df = pd.read_excel(path_string, sheet_name=sheet_name, header=None)

    header_row = None
    for row_idx in range(min(30, len(raw_df))):
        values = {
            str(value).strip().casefold()
            for value in raw_df.iloc[row_idx].tolist()
            if pd.notna(value)
        }
        if "no" in values and "nama" in values:
            header_row = row_idx
            break

    if header_row is None:
        return raw_df, None

    headers = []
    seen = {}
    for idx, value in enumerate(raw_df.iloc[header_row].tolist()):
        name = str(value).strip() if pd.notna(value) else ""
        if not name or name.casefold() == "nan":
            name = f"Kolom_{idx + 1}"
        seen[name] = seen.get(name, 0) + 1
        if seen[name] > 1:
            name = f"{name}_{seen[name]}"
        headers.append(name)

    df = raw_df.iloc[header_row + 1:].copy()
    df.columns = headers

    # Hilangkan kolom yang benar-benar kosong pada seluruh data.
    df = df.dropna(axis=1, how="all").copy()

    # Untuk workbook data petani, satu petani biasanya memiliki satu baris
    # utama dengan kolom Nama; baris berikutnya kadang hanya berisi tautan peta.
    # Baris pendukung tersebut tidak ditampilkan sebagai record petani terpisah.
    name_col = next(
        (col for col in df.columns if str(col).strip().casefold() == "nama"),
        None,
    )
    if name_col is not None:
        names = df[name_col].astype("string").str.strip()
        valid_names = (
            df[name_col].notna()
            & names.ne("")
            & names.str.casefold().ne("none")
        )
        df = df.loc[valid_names].copy()
    else:
        df = df.dropna(how="all").copy()

    return df, header_row


st.markdown("""
<style>
/* Hilangkan ruang kosong berlebih di atas konten utama. */
[data-testid="stMainBlockContainer"] {
    padding-top: 0.75rem !important;
}
section.main > div.block-container {
    padding-top: 0.75rem !important;
}
/* Desaku Maju — professional sidebar navigation */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #f4f8f5 0%, #eef4f0 100%);
    border-right: 1px solid #dce7df;
}
section[data-testid="stSidebar"] > div {
    padding-top: 1.25rem;
}
.dm-brand {
    padding: 20px 17px 18px 17px;
    margin: 0 0 20px 0;
    border-radius: 18px;
    background: linear-gradient(135deg, #123d2c 0%, #176b4a 100%);
    color: #ffffff;
    box-shadow: 0 8px 22px rgba(18, 61, 44, 0.18);
}
.dm-brand-mark {
    font-size: 31px;
    line-height: 1.1;
    margin-bottom: 10px;
}
.dm-brand-title {
    font-size: 22px;
    font-weight: 800;
    letter-spacing: -0.45px;
    line-height: 1.15;
    color: #ffffff;
}
.dm-brand-subtitle {
    margin-top: 7px;
    font-size: 11px;
    line-height: 1.45;
    letter-spacing: 1.1px;
    text-transform: uppercase;
    color: #d7eee1;
    font-weight: 650;
}
.dm-nav-section {
    font-size: 11px;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    font-weight: 800;
    color: #658071;
    margin: 5px 0 8px 2px;
}
section[data-testid="stSidebar"] [data-testid="stRadio"] > label {
    font-size: 13px !important;
    font-weight: 750 !important;
    color: #345344 !important;
    margin-bottom: 9px;
}
section[data-testid="stSidebar"] [data-testid="stRadio"] div[role="radiogroup"] {
    gap: 7px;
}
section[data-testid="stSidebar"] [data-testid="stRadio"] div[role="radiogroup"] label {
    padding: 10px 12px !important;
    border: 1px solid #dce8df !important;
    border-radius: 12px !important;
    background: rgba(255,255,255,0.78) !important;
    min-height: 43px;
    align-items: center !important;
    transition: border-color 120ms ease, background 120ms ease, box-shadow 120ms ease;
}
section[data-testid="stSidebar"] [data-testid="stRadio"] div[role="radiogroup"] label:hover {
    border-color: #82b49a !important;
    background: #ffffff !important;
}
section[data-testid="stSidebar"] [data-testid="stRadio"] div[role="radiogroup"] label:has(input:checked) {
    background: #e0f2e7 !important;
    border: 1px solid #2d8558 !important;
    box-shadow: 0 3px 10px rgba(31, 111, 69, 0.08);
}
section[data-testid="stSidebar"] [data-testid="stRadio"] div[role="radiogroup"] label p {
    font-size: 14px !important;
    font-weight: 650 !important;
    color: #254536 !important;
    line-height: 1.35 !important;
}
.dm-sidebar-foot {
    border-top: 1px solid #d6e2da;
    margin-top: 22px;
    padding-top: 14px;
    color: #74877c;
    font-size: 11px;
    line-height: 1.6;
}
.dm-status-dot {
    display: inline-block;
    width: 7px;
    height: 7px;
    border-radius: 50%;
    background: #2d9a60;
    margin-right: 7px;
}
@media (max-width: 700px) {
    .dm-brand-title { font-size: 20px; }
    section[data-testid="stSidebar"] [data-testid="stRadio"] div[role="radiogroup"] label {
        padding: 9px 10px !important;
    }
}
</style>
""", unsafe_allow_html=True)


# Banner utama Desaku Maju — ditampilkan paling atas pada area konten.
# Jika banner.png belum tersedia, gunakan logo resmi sebagai fallback.
if BANNER_PATH.exists():
    # Pangkas pita putih kosong di bagian atas banner (jika ada),
    # tanpa mengubah lebar atau memotong bendera, petani, dan padi.
    with Image.open(BANNER_PATH) as banner_file:
        banner_image = banner_file.convert("RGB")
        banner_pixels = np.asarray(banner_image)
        near_white = np.all(banner_pixels >= 245, axis=2)
        white_ratio_by_row = near_white.mean(axis=1)
        content_rows = np.where(white_ratio_by_row < 0.97)[0]
        if len(content_rows):
            crop_top = max(0, int(content_rows[0]) - 2)
            if crop_top > 0:
                banner_image = banner_image.crop(
                    (0, crop_top, banner_image.width, banner_image.height)
                )
        st.image(
            banner_image,
            width="stretch",
            caption=None,
            output_format="PNG",
            clamp=True,
        )
elif LOGO_PATH.exists():
    st.image(str(LOGO_PATH), width="stretch", caption=None, output_format="PNG", alt="Logo resmi Desaku Maju")
else:
    st.title("Desaku Maju")
st.caption("Agricultural Intelligence Platform | Demo Lampung")
st.markdown(
    "Platform digital untuk membantu petani mengelola lahan, "
    "mengakses layanan pertanian, dan memasarkan hasil bumi."
)

with st.sidebar:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=245, caption=None, output_format="PNG")
    else:
        st.markdown(
            """
            <div class="dm-brand">
                <div class="dm-brand-mark">🌾</div>
                <div class="dm-brand-title">Desaku Maju</div>
                <div class="dm-brand-subtitle">Agricultural Intelligence</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    st.markdown('<div class="dm-nav-section">Workspace</div>', unsafe_allow_html=True)
    page = st.radio(
        "Pilih modul",
        [
            "Ringkasan",
            "Prioritas Dryer",
            "Eksplorasi Database",
            "Profil Petani (data sampel)",
            "Portal Vendor",
            "Marketplace",
            "AI Assistant & Learning Center",
        ],
        format_func=lambda item: {
            "Ringkasan": "🏡  Ringkasan",
            "Prioritas Dryer": "🌬️  Prioritas Dryer",
            "Eksplorasi Database": "📊  Eksplorasi Database",
            "Profil Petani (data sampel)": "👩‍🌾  Profil Petani",
            "Portal Vendor": "🤝  Portal Vendor",
            "Marketplace": "🛒  Marketplace",
            "AI Assistant & Learning Center": "✨  AI Assistant & Learning Center",
        }.get(item, item),
        key="main_navigation",
        label_visibility="collapsed",
    )
    st.markdown(
        """
        <div class="dm-sidebar-foot">
            <span class="dm-status-dot"></span><b>DEMO ENVIRONMENT</b><br>
            Sumber data: workbook Excel lokal<br>
            Platform pertanian terpadu
        </div>
        """,
        unsafe_allow_html=True,
    )

if not excel_files:
    st.error(
        "File Excel belum ditemukan. Simpan workbook .xlsx di folder yang sama dengan app.py."
    )
    st.stop()

workbook_paths = {
    p.name: p for p in excel_files
}

if page == "Eksplorasi Database":
    st.header("Eksplorasi Database")

    selected_file = st.selectbox(
        "Pilih workbook",
        list(workbook_paths.keys()),
    )

    path = workbook_paths[selected_file]

    try:
        sheets = load_workbook(
            str(path),
            path.stat().st_mtime,
        )
    except Exception as exc:
        st.error(f"Gagal membaca workbook: {exc}")
        st.stop()

    sheet_name = st.selectbox(
        "Pilih sheet",
        list(sheets.keys()),
    )

    is_farmer_workbook = "petani" in selected_file.lower()
    try:
        df, detected_header_row = load_sheet_for_display(
            str(path),
            sheet_name,
            path.stat().st_mtime,
            detect_farmer_header=is_farmer_workbook,
        )
    except Exception as exc:
        st.error(f"Gagal membaca sheet: {exc}")
        st.stop()

    st.caption(f"File: {selected_file} | Sheet: {sheet_name}")
    if is_farmer_workbook:
        if detected_header_row is not None:
            st.caption(
                f"Header tabel petani terdeteksi pada baris Excel "
                f"{detected_header_row + 1}."
            )
        else:
            st.warning(
                "Header tabel petani dengan kolom 'No' dan 'Nama' "
                "belum terdeteksi; data ditampilkan untuk pemeriksaan."
            )

    c1, c2, c3 = st.columns(3)
    c1.metric("Jumlah baris", f"{len(df):,}")
    c2.metric("Jumlah kolom", f"{len(df.columns):,}")
    c3.metric(
        "Sel kosong",
        f"{int(df.isna().sum().sum()):,}",
    )

    st.subheader("Pratinjau data")
    st.dataframe(df.head(100), use_container_width=True)

    st.subheader("Ringkasan kolom")
    summary = pd.DataFrame({
        "Nama kolom": df.columns,
        "Tipe data": [str(df[c].dtype) for c in df.columns],
        "Jumlah terisi": [int(df[c].notna().sum()) for c in df.columns],
        "Jumlah kosong": [int(df[c].isna().sum()) for c in df.columns],
    })
    st.dataframe(summary, use_container_width=True, hide_index=True)

    numeric_cols = df.select_dtypes(include="number").columns.tolist()

    if numeric_cols:
        st.subheader("Grafik eksplorasi")
        chart_col = st.selectbox(
            "Pilih kolom numerik",
            numeric_cols,
        )

        chart_df = df[[chart_col]].dropna().copy()

        if not chart_df.empty:
            # Gunakan histogram agar bisa dipakai pada beragam data numerik.
            fig = px.histogram(
                chart_df,
                x=chart_col,
                title=f"Distribusi: {chart_col}",
                nbins=25,
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Kolom ini belum memiliki nilai numerik yang terisi.")
    else:
        st.info("Sheet ini tidak memiliki kolom numerik yang terdeteksi.")

elif page == "Prioritas Dryer":
    st.header("Prioritas Fasilitas Pengering Hasil Panen")
    st.write(
        "Daftar lokasi pada workbook yang memiliki isian kelompok tani/BUMDes "
        "terkait program dryer. Kolom kesiapan yang kosong ditampilkan sebagai "
        "'Belum ada data', bukan dianggap tidak siap."
    )

    main_files = [
        p for p in excel_files
        if "DATABASE DESAKU MAJU" in p.name.upper()
    ]

    if not main_files:
        st.warning("Workbook database utama tidak ditemukan.")
        st.stop()

    main_path = main_files[0]
    try:
        main_sheets = load_workbook(str(main_path), main_path.stat().st_mtime)
        if "GABUNGAN" not in main_sheets:
            st.error("Sheet 'GABUNGAN' tidak ditemukan.")
            st.stop()
        dryer_data = main_sheets["GABUNGAN"].copy()
    except Exception as exc:
        st.error(f"Gagal membaca data prioritas dryer: {exc}")
        st.stop()

    dryer_col = "Hibah Dryer  GAPOKTAN/  BUMDES"
    proposal_col = "KETERANGAN (ADA/ TIDAK ADA) PROPOSAL"
    land_col = "KETERSEDIAAN LAHAN"
    house_col = "KESANGGUPAN MEMBANGUN RUMAH DRYER"
    trade_col = "KESEDIAN KERJASAMA TATA NIAGA"

    if dryer_col not in dryer_data.columns:
        st.error(f"Kolom '{dryer_col}' tidak ditemukan di sheet GABUNGAN.")
        st.stop()

    # Workbook saat ini memiliki 82 baris dengan isian kelompok tani/BUMDes.
    # Ini adalah daftar lokasi pada file, bukan klaim bahwa semua sudah siap.
    dryer_data = dryer_data[
        dryer_data[dryer_col].notna()
        & dryer_data[dryer_col].astype(str).str.strip().ne("")
    ].copy()

    def is_filled(series):
        text_values = series.astype("string").str.strip()
        return (
            series.notna()
            & text_values.ne("")
            & text_values.str.casefold().ne("none")
            & text_values.ne("-")
        )

    proposal_filled = is_filled(dryer_data[proposal_col]) if proposal_col in dryer_data.columns else pd.Series(False, index=dryer_data.index)
    proposal_text = dryer_data[proposal_col].astype("string").str.strip().str.casefold() if proposal_col in dryer_data.columns else pd.Series("", index=dryer_data.index)
    proposal_yes = proposal_text.isin(["proposal ada"])
    proposal_no = proposal_text.isin(["proposal belum ada"])

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Lokasi dalam daftar", f"{len(dryer_data):,}")
    k2.metric("Proposal terisi", f"{int(proposal_filled.sum()):,}")
    k3.metric("Proposal tercatat ada", f"{int(proposal_yes.sum()):,}")
    k4.metric("Data kesiapan lahan", f"{int(is_filled(dryer_data[land_col]).sum()) if land_col in dryer_data.columns else 0:,}")

    st.warning(
        "Peringatan kualitas data: kolom ketersediaan lahan, kesanggupan membangun "
        "rumah dryer, dan kesediaan kerja sama tata niaga belum terisi dalam workbook "
        "yang diperiksa. Status tersebut berarti 'belum ada data', bukan 'tidak siap'."
    )

    # Filter wilayah
    district_values = sorted(dryer_data["KABUPATEN/KOTA"].dropna().astype(str).unique().tolist())
    district_filter = st.selectbox(
        "Filter Kabupaten/Kota",
        ["Semua Kabupaten/Kota"] + district_values,
        key="dryer_district_filter",
    )
    visible = dryer_data.copy()
    if district_filter != "Semua Kabupaten/Kota":
        visible = visible[visible["KABUPATEN/KOTA"].astype(str) == district_filter]

    # Kolom tampilan dibuat hanya dari kolom yang tersedia.
    display_candidates = [
        "KABUPATEN/KOTA",
        "KECAMATAN",
        "DESA",
        dryer_col,
        proposal_col,
        land_col,
        house_col,
        trade_col,
    ]
    display_cols = [col for col in display_candidates if col in visible.columns]
    display = visible[display_cols].copy()
    display = display.rename(columns={
        "KABUPATEN/KOTA": "Kabupaten/Kota",
        "KECAMATAN": "Kecamatan",
        "DESA": "Desa",
        dryer_col: "Kelompok Tani / BUMDes",
        proposal_col: "Status Proposal",
        land_col: "Ketersediaan Lahan",
        house_col: "Kesanggupan Rumah Dryer",
        trade_col: "Kesediaan Kerja Sama Tata Niaga",
    })
    for col in display.columns:
        display[col] = display[col].where(display[col].notna(), "Belum ada data")
        display[col] = display[col].replace("", "Belum ada data")
        display[col] = display[col].replace("-", "Belum ada data")

    st.subheader("Daftar lokasi")
    st.dataframe(display, use_container_width=True, hide_index=True)

    st.caption(
        f"Sumber: {main_path.name}, sheet GABUNGAN. "
        "Daftar ditentukan dari baris yang memiliki isian kelompok tani/BUMDes. "
        "Status kesiapan tidak disimpulkan dari kolom kosong."
    )

elif page == "Profil Petani (data sampel)":
    st.header("Dashboard Petani")
    st.caption("Profil petani dan lahan berdasarkan workbook sampel yang tersedia.")

    farmer_files = [p for p in excel_files if "petani" in p.name.lower()]
    if not farmer_files:
        st.warning("Workbook dengan nama yang mengandung 'petani' belum ditemukan.")
        st.stop()

    path = farmer_files[0]
    try:
        sheets = load_workbook(str(path), path.stat().st_mtime)
    except Exception as exc:
        st.error(f"Gagal membaca data petani: {exc}")
        st.stop()

    sheet_name = st.selectbox("Pilih sheet data petani", list(sheets.keys()))
    try:
        df, detected_header_row = load_sheet_for_display(
            str(path),
            sheet_name,
            path.stat().st_mtime,
            detect_farmer_header=True,
        )
    except Exception as exc:
        st.error(f"Gagal membaca sheet data petani: {exc}")
        st.stop()

    if detected_header_row is None:
        st.warning(
            "Header tabel tidak ditemukan otomatis. "
            "Periksa apakah sheet memiliki kolom 'No' dan 'Nama'."
        )
        st.dataframe(df.head(30), use_container_width=True)
        st.stop()

    name_col = next((c for c in df.columns if str(c).strip().casefold() == "nama"), None)
    area_col = next((c for c in df.columns if "luas lahan" in str(c).strip().casefold()), None)
    ownership_col = next((c for c in df.columns if "kepemilikan lahan" in str(c).strip().casefold()), None)
    crop_col = next((c for c in df.columns if "jenis tanaman" in str(c).strip().casefold()), None)
    coord_col = next((c for c in df.columns if "koordinat lahan" in str(c).strip().casefold()), None)

    if name_col is None:
        st.error("Kolom Nama tidak ditemukan setelah header dibaca.")
        st.dataframe(df.head(30), use_container_width=True)
        st.stop()

    # Dashboard hanya menggunakan baris yang memiliki nama petani.
    names = df[name_col].astype("string").str.strip()
    farmer_rows = df[
        df[name_col].notna()
        & names.ne("")
        & names.str.casefold().ne("none")
    ].copy()

    if farmer_rows.empty:
        st.warning("Belum ada baris petani dengan nama yang dapat ditampilkan.")
        st.stop()

    # Ubah teks seperti '0.5 ha' menjadi angka hektare untuk analitik.
    if area_col is not None:
        area_text = farmer_rows[area_col].astype("string").str.replace(",", ".", regex=False)
        area_numeric = pd.to_numeric(
            area_text.str.extract(r"([-+]?[0-9]*\.?[0-9]+)", expand=False),
            errors="coerce",
        )
        farmer_rows["_Luas_ha"] = area_numeric
    else:
        farmer_rows["_Luas_ha"] = pd.NA

    st.info(
        "Catatan: ini adalah data sampel dari workbook, bukan sensus atau "
        "representasi seluruh petani di Lampung. Angka dihitung hanya dari "
        "baris yang memiliki nama petani dan luas lahan yang dapat dibaca."
    )

    # Filter sederhana untuk presentasi.
    filter_col1, filter_col2 = st.columns(2)
    with filter_col1:
        crop_options = ["Semua tanaman"]
        if crop_col is not None:
            crop_options += sorted(
                farmer_rows[crop_col].dropna().astype(str).str.strip().loc[
                    lambda s: s.ne("")
                ].unique().tolist()
            )
        selected_crop = st.selectbox("Filter jenis tanaman", crop_options)

    with filter_col2:
        ownership_options = ["Semua status"]
        if ownership_col is not None:
            ownership_options += sorted(
                farmer_rows[ownership_col].dropna().astype(str).str.strip().loc[
                    lambda s: s.ne("")
                ].unique().tolist()
            )
        selected_ownership = st.selectbox("Filter kepemilikan lahan", ownership_options)

    filtered_farmers = farmer_rows.copy()
    if crop_col is not None and selected_crop != "Semua tanaman":
        filtered_farmers = filtered_farmers[
            filtered_farmers[crop_col].astype(str).str.strip() == selected_crop
        ]
    if ownership_col is not None and selected_ownership != "Semua status":
        filtered_farmers = filtered_farmers[
            filtered_farmers[ownership_col].astype(str).str.strip() == selected_ownership
        ]

    total_farmers = len(filtered_farmers)
    total_area = filtered_farmers["_Luas_ha"].sum(min_count=1)
    valid_area_count = int(filtered_farmers["_Luas_ha"].notna().sum())
    avg_area = filtered_farmers["_Luas_ha"].mean()
    own_count = 0
    if ownership_col is not None:
        own_count = int(
            filtered_farmers[ownership_col].astype(str).str.casefold()
            .str.contains("milik sendiri", na=False).sum()
        )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Petani dalam sampel", f"{total_farmers:,}")
    m2.metric("Total luas lahan teridentifikasi", f"{total_area:,.2f} ha" if pd.notna(total_area) else "Belum tersedia")
    m3.metric("Rata-rata luas per baris", f"{avg_area:,.2f} ha" if pd.notna(avg_area) else "Belum tersedia")
    m4.metric("Milik sendiri", f"{own_count:,}")

    st.divider()
    chart_left, chart_right = st.columns(2)

    with chart_left:
        st.subheader("Petani menurut jenis tanaman")
        if crop_col is not None:
            crop_summary = (
                filtered_farmers[crop_col]
                .fillna("Belum diisi")
                .astype(str)
                .str.strip()
                .replace("", "Belum diisi")
                .value_counts()
                .rename_axis("Jenis tanaman")
                .reset_index(name="Jumlah baris petani")
            )
            if not crop_summary.empty:
                fig_crop = px.bar(
                    crop_summary,
                    x="Jenis tanaman",
                    y="Jumlah baris petani",
                    title="Distribusi tanaman pada sampel",
                    text="Jumlah baris petani",
                )
                fig_crop.update_traces(marker_color="#238636")
                st.plotly_chart(fig_crop, use_container_width=True)
            else:
                st.caption("Belum ada data jenis tanaman.")
        else:
            st.caption("Kolom jenis tanaman tidak tersedia.")

    with chart_right:
        st.subheader("Status kepemilikan lahan")
        if ownership_col is not None:
            ownership_summary = (
                filtered_farmers[ownership_col]
                .fillna("Belum diisi")
                .astype(str)
                .str.strip()
                .replace("", "Belum diisi")
                .value_counts()
                .rename_axis("Status")
                .reset_index(name="Jumlah baris petani")
            )
            if not ownership_summary.empty:
                fig_ownership = px.pie(
                    ownership_summary,
                    names="Status",
                    values="Jumlah baris petani",
                    title="Kepemilikan lahan pada sampel",
                    hole=0.45,
                )
                st.plotly_chart(fig_ownership, use_container_width=True)
            else:
                st.caption("Belum ada data kepemilikan lahan.")
        else:
            st.caption("Kolom kepemilikan lahan tidak tersedia.")

    # Pilih petani terlebih dahulu agar peta mengikuti pilihan pengguna.
    st.subheader("Detail Profil Petani")
    st.caption(
        "Pilih satu nama untuk melihat lokasi dan ringkasan lahannya. "
        "Informasi ini berasal dari data sampel."
    )
    profile_names = (
        filtered_farmers[name_col]
        .dropna()
        .astype(str)
        .str.strip()
    )
    profile_names = sorted(
        name for name in profile_names.unique()
        if name and name.casefold() != "none"
    )

    selected_rows = pd.DataFrame()
    selected_profile = None
    if profile_names:
        selected_profile = st.selectbox(
            "Pilih petani — peta akan mengikuti pilihan ini",
            profile_names,
            key="farmer_profile_select",
        )
        selected_rows = filtered_farmers[
            filtered_farmers[name_col].astype(str).str.strip() == selected_profile
        ].copy()
    else:
        st.info("Belum ada nama petani yang dapat dipilih.")

    # Peta menampilkan lokasi petani terpilih, bukan semua titik sekaligus.
    st.divider()
    st.subheader("Peta Lokasi Lahan Petani Terpilih")
    if coord_col is not None and not selected_rows.empty:
        coord_text = selected_rows[coord_col].astype("string")
        extracted = coord_text.str.extract(
            r"(-?\d+(?:[.,]\d+)?)\s*[,;]\s*(-?\d+(?:[.,]\d+)?)"
        )
        if not extracted.empty:
            selected_rows["_latitude"] = pd.to_numeric(
                extracted[0].str.replace(",", ".", regex=False),
                errors="coerce",
            )
            selected_rows["_longitude"] = pd.to_numeric(
                extracted[1].str.replace(",", ".", regex=False),
                errors="coerce",
            )
            map_points = selected_rows.dropna(
                subset=["_latitude", "_longitude"]
            ).copy()
            map_points = map_points[
                map_points["_latitude"].between(-90, 90)
                & map_points["_longitude"].between(-180, 180)
            ]

            if not map_points.empty:
                location = map_points.iloc[0]
                st.success(f"Peta menunjukkan lokasi lahan milik/garapan: {selected_profile}")
                st.caption(
                    f"Koordinat: {location['_latitude']:.6f}, "
                    f"{location['_longitude']:.6f}. "
                    "Titik ini berubah saat Anda memilih petani lain."
                )
                st.map(
                    pd.DataFrame(
                        {
                            "latitude": [float(location["_latitude"])],
                            "longitude": [float(location["_longitude"])],
                        }
                    ),
                    latitude="latitude",
                    longitude="longitude",
                    size=180,
                    color="#238636",
                    zoom=13,
                )
            else:
                st.warning(
                    f"Koordinat untuk {selected_profile} belum dapat dipetakan. "
                    "Periksa format koordinat pada workbook."
                )
        else:
            st.warning("Kolom koordinat tidak memiliki data yang dapat dipetakan.")
    elif coord_col is None:
        st.info("Kolom 'Koordinat Lahan' tidak ditemukan pada workbook.")
    else:
        st.info("Pilih petani untuk menampilkan lokasi lahannya.")

    if not selected_rows.empty:
        profile = selected_rows.iloc[0]
        p1, p2, p3 = st.columns(3)
        p1.metric("Nama", selected_profile)
        if area_col is not None:
            p1.caption(
                f"Luas lahan sumber: {profile[area_col]}"
                if pd.notna(profile[area_col])
                else "Luas lahan sumber: Belum tersedia"
            )
        if crop_col is not None:
            p2.metric(
                "Jenis tanaman",
                str(profile[crop_col]) if pd.notna(profile[crop_col]) else "Belum diisi",
            )
        if ownership_col is not None:
            p3.metric(
                "Kepemilikan lahan",
                str(profile[ownership_col]) if pd.notna(profile[ownership_col]) else "Belum diisi",
            )
        if coord_col is not None:
            coord_value = profile[coord_col]
            st.text_input(
                "Koordinat lahan (data sumber)",
                value=str(coord_value) if pd.notna(coord_value) else "Belum tersedia",
                disabled=True,
                key="selected_farmer_coordinates",
            )

    st.subheader("Daftar petani")
    search_name = st.text_input("Cari nama petani", placeholder="Ketik nama untuk mencari...")
    visible = filtered_farmers.copy()
    if search_name.strip():
        visible = visible[
            visible[name_col].astype(str).str.contains(
                search_name.strip(), case=False, na=False, regex=False
            )
        ]

    display_cols = [name_col]
    for col in [area_col, crop_col, ownership_col, coord_col]:
        if col is not None and col not in display_cols:
            display_cols.append(col)

    display = visible[display_cols].copy()
    rename_map = {}
    if area_col is not None:
        rename_map[area_col] = "Luas Lahan (data sumber)"
    if crop_col is not None:
        rename_map[crop_col] = "Jenis Tanaman"
    if ownership_col is not None:
        rename_map[ownership_col] = "Kepemilikan Lahan"
    if coord_col is not None:
        rename_map[coord_col] = "Koordinat Lahan"
    display = display.rename(columns=rename_map)
    st.dataframe(display, use_container_width=True, hide_index=True)

    st.caption(
        f"Sumber: {path.name}, sheet {sheet_name}. "
        f"Luas lahan berhasil dibaca untuk {valid_area_count:,} baris dari "
        f"{total_farmers:,} baris petani yang sedang ditampilkan. "
        "Nilai luas berasal dari teks pada workbook dan perlu diverifikasi sebelum "
        "digunakan untuk keputusan operasional atau pembiayaan."
    )

elif page == "Marketplace":
    st.header("Marketplace Desaku Maju")
    st.write(
        "Tempat mempertemukan petani dan pembeli hasil pertanian. "
        "Pada versi demo, produk, volume, harga, dan penjual di bawah adalah data simulasi."
    )

    st.markdown(
        """
        <div style="padding:18px 20px;border-radius:16px;background:linear-gradient(120deg,#123d2c,#218454);color:white;margin-bottom:18px">
            <div style="font-size:12px;letter-spacing:1.5px;text-transform:uppercase;color:#d9f4e4;font-weight:700">Pasar Hasil Tani</div>
            <div style="font-size:25px;font-weight:800;margin:5px 0">Hasil panen bertemu pembeli</div>
            <div style="font-size:14px;color:#effaf3">Jelajahi komoditas, lihat penawaran, dan buat permintaan kontak demo.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    marketplace_products = pd.DataFrame([
        {"Produk": "Beras Premium (Demo)", "Komoditas": "Beras", "Kategori": "Pangan", "Asal": "Lampung Tengah", "Volume": "5 ton", "Harga": "Simulasi — hubungi penjual", "Penjual": "Kelompok Tani Makmur (Demo)", "Keterangan": "Mutu dan kadar air perlu diverifikasi"},
        {"Produk": "Gabah Kering Panen (Demo)", "Komoditas": "Gabah", "Kategori": "Pangan", "Asal": "Lampung Timur", "Volume": "12 ton", "Harga": "Simulasi — hubungi penjual", "Penjual": "Gapoktan Sumber Rezeki (Demo)", "Keterangan": "Harga aktual mengikuti mutu dan lokasi"},
        {"Produk": "Jagung Pipilan (Demo)", "Komoditas": "Jagung", "Kategori": "Pangan", "Asal": "Lampung Selatan", "Volume": "8 ton", "Harga": "Simulasi — hubungi penjual", "Penjual": "Petani Jagung Sejahtera (Demo)", "Keterangan": "Kadar air dan kualitas perlu diperiksa"},
        {"Produk": "Singkong Segar (Demo)", "Komoditas": "Singkong", "Kategori": "Pangan", "Asal": "Lampung Utara", "Volume": "10 ton", "Harga": "Simulasi — hubungi penjual", "Penjual": "Mitra Tani Utara (Demo)", "Keterangan": "Jadwal panen dan pengiriman perlu dikonfirmasi"},
        {"Produk": "Cabai Merah (Demo)", "Komoditas": "Cabai", "Kategori": "Hortikultura", "Asal": "Pesawaran", "Volume": "500 kg", "Harga": "Simulasi — hubungi penjual", "Penjual": "Kebun Sayur Mandiri (Demo)", "Keterangan": "Kesegaran dan grade perlu diverifikasi"},
        {"Produk": "Kopi Robusta (Demo)", "Komoditas": "Kopi", "Kategori": "Perkebunan", "Asal": "Lampung Barat", "Volume": "1 ton", "Harga": "Simulasi — hubungi penjual", "Penjual": "Koperasi Kopi Rakyat (Demo)", "Keterangan": "Bentuk produk dan kualitas perlu dikonfirmasi"},
    ])

    st.warning(
        "KATALOG DEMO: semua produk dan penjual merupakan contoh antarmuka, bukan penawaran nyata. "
        "Harga, ketersediaan, mutu, dan identitas penjual belum diverifikasi."
    )

    search_col, category_col, origin_col = st.columns([2, 1, 1])
    with search_col:
        marketplace_query = st.text_input(
            "Cari komoditas atau produk",
            placeholder="Contoh: jagung, beras, kopi",
            key="marketplace_search",
        )
    with category_col:
        marketplace_categories = ["Semua kategori"] + sorted(marketplace_products["Kategori"].unique().tolist())
        marketplace_category = st.selectbox("Kategori", marketplace_categories, key="marketplace_category")
    with origin_col:
        marketplace_origins = ["Semua wilayah"] + sorted(marketplace_products["Asal"].unique().tolist())
        marketplace_origin = st.selectbox("Asal produk", marketplace_origins, key="marketplace_origin")

    visible_products = marketplace_products.copy()
    if marketplace_query.strip():
        searchable = visible_products[
            ["Produk", "Komoditas", "Penjual", "Keterangan", "Asal"]
        ].fillna("").astype(str).agg(" ".join, axis=1)
        visible_products = visible_products[
            searchable.str.contains(marketplace_query.strip(), case=False, regex=False)
        ]
    if marketplace_category != "Semua kategori":
        visible_products = visible_products[visible_products["Kategori"] == marketplace_category]
    if marketplace_origin != "Semua wilayah":
        visible_products = visible_products[visible_products["Asal"] == marketplace_origin]

    m1, m2, m3 = st.columns(3)
    m1.metric("Penawaran ditemukan", f"{len(visible_products):,}")
    m2.metric("Komoditas tersedia", f"{visible_products['Komoditas'].nunique():,}" if not visible_products.empty else "0")
    m3.metric("Transaksi aktif", "Demo")

    st.subheader("Jelajahi hasil pertanian")
    if visible_products.empty:
        st.info("Produk tidak ditemukan. Coba kata kunci atau filter lain.")
    else:
        card_cols = st.columns(3)
        for idx, (_, product) in enumerate(visible_products.iterrows()):
            with card_cols[idx % 3]:
                st.markdown(
                    f"""
                    <div style="border:1px solid #dce7df;border-radius:14px;padding:16px 16px 14px 16px;margin-bottom:10px;min-height:210px;background:#ffffff">
                        <div style="font-size:11px;color:#27724b;font-weight:800;letter-spacing:1px;text-transform:uppercase">{product['Kategori']}</div>
                        <div style="font-size:18px;font-weight:800;color:#183d2c;margin:8px 0">{product['Produk']}</div>
                        <div style="font-size:13px;color:#53675c;line-height:1.55">
                            <b>Volume:</b> {product['Volume']}<br>
                            <b>Asal:</b> {product['Asal']}<br>
                            <b>Penjual:</b> {product['Penjual']}
                        </div>
                        <div style="margin-top:12px;padding-top:10px;border-top:1px solid #e7eee9;font-size:12px;color:#64776b">{product['Harga']}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                if st.button("Tanyakan produk", key=f"marketplace_inquire_{idx}", use_container_width=True):
                    st.session_state["marketplace_selected_product"] = product["Produk"]

    st.divider()
    st.subheader("Permintaan informasi produk")
    st.caption("Form ini hanya menyimpan permintaan pada sesi demo; belum mengirim pesan ke penjual.")
    selected_product = st.session_state.get("marketplace_selected_product", "")
    with st.form("marketplace_inquiry_form", clear_on_submit=False):
        inquiry_name = st.text_input("Nama petani / pembeli / kelompok", key="marketplace_inquiry_name")
        inquiry_contact = st.text_input("Kontak (opsional)", key="marketplace_inquiry_contact")
        product_options = visible_products["Produk"].tolist() if not visible_products.empty else marketplace_products["Produk"].tolist()
        if selected_product not in product_options:
            selected_product = product_options[0] if product_options else ""
        inquiry_product = st.selectbox(
            "Produk yang ditanyakan",
            product_options,
            index=product_options.index(selected_product) if selected_product in product_options else 0,
            key="marketplace_inquiry_product",
        )
        inquiry_message = st.text_area(
            "Pesan / kebutuhan",
            placeholder="Contoh: mohon informasi harga aktual, mutu, dan jadwal pengiriman.",
            key="marketplace_inquiry_message",
        )
        inquiry_submit = st.form_submit_button("Kirim permintaan demo", type="primary")

    if inquiry_submit:
        if not inquiry_name.strip() or not inquiry_message.strip():
            st.error("Mohon isi nama dan pesan terlebih dahulu.")
        else:
            if "marketplace_demo_inquiries" not in st.session_state:
                st.session_state["marketplace_demo_inquiries"] = []
            st.session_state["marketplace_demo_inquiries"].append({
                "Nama": inquiry_name.strip(),
                "Kontak": inquiry_contact.strip(),
                "Produk": inquiry_product,
                "Pesan": inquiry_message.strip(),
                "Status": "Demo — belum dikirim ke penjual",
            })
            st.success("Permintaan informasi tersimpan pada sesi demo. Belum dikirim ke penjual.")

    inquiry_history = st.session_state.get("marketplace_demo_inquiries", [])
    if inquiry_history:
        with st.expander("Riwayat permintaan pada sesi ini"):
            st.dataframe(pd.DataFrame(inquiry_history), use_container_width=True, hide_index=True)

elif page == "Portal Vendor":
    st.header("Portal Vendor Pertanian")
    st.write(
        "Katalog demo untuk memperlihatkan bagaimana petani dapat menemukan "
        "pemasok sarana dan layanan pertanian. Semua vendor, harga, dan stok di "
        "halaman ini adalah data simulasi, bukan penawaran komersial nyata."
    )

    vendors = pd.DataFrame([
        {
            "Nama Vendor": "Toko Tani Makmur (Demo)",
            "Kategori": "Benih",
            "Produk / Layanan": "Benih padi unggul",
            "Keterangan": "Kemasan dan varietas perlu dikonfirmasi",
            "Harga Indikatif": "Simulasi — belum ditentukan",
            "Wilayah Layanan": "Lampung Tengah",
        },
        {
            "Nama Vendor": "Sarana Jagung Sejahtera (Demo)",
            "Kategori": "Benih",
            "Produk / Layanan": "Benih jagung",
            "Keterangan": "Spesifikasi dan sertifikasi perlu diverifikasi",
            "Harga Indikatif": "Simulasi — belum ditentukan",
            "Wilayah Layanan": "Lampung Timur",
        },
        {
            "Nama Vendor": "Pupuk Mitra Desa (Demo)",
            "Kategori": "Pupuk",
            "Produk / Layanan": "Pupuk dan nutrisi tanaman",
            "Keterangan": "Jenis produk perlu disesuaikan dengan kebutuhan",
            "Harga Indikatif": "Simulasi — belum ditentukan",
            "Wilayah Layanan": "Lampung Selatan",
        },
        {
            "Nama Vendor": "Mekanisasi Tani (Demo)",
            "Kategori": "Alat pertanian",
            "Produk / Layanan": "Sewa alat dan mesin pertanian",
            "Keterangan": "Jadwal dan kapasitas perlu dikonfirmasi",
            "Harga Indikatif": "Simulasi — belum ditentukan",
            "Wilayah Layanan": "Lampung Tengah",
        },
        {
            "Nama Vendor": "Pengering Hasil Panen (Demo)",
            "Kategori": "Pascapanen",
            "Produk / Layanan": "Layanan pengeringan gabah",
            "Keterangan": "Kapasitas, kadar air, dan jadwal perlu dikonfirmasi",
            "Harga Indikatif": "Simulasi — belum ditentukan",
            "Wilayah Layanan": "Lampung Selatan",
        },
        {
            "Nama Vendor": "Angkut Hasil Tani (Demo)",
            "Kategori": "Logistik",
            "Produk / Layanan": "Transportasi hasil panen",
            "Keterangan": "Rute dan muatan perlu dikonfirmasi",
            "Harga Indikatif": "Simulasi — belum ditentukan",
            "Wilayah Layanan": "Lampung Utara",
        },
    ])

    st.warning(
        "MODE DEMO: nama vendor, katalog, dan wilayah layanan di bawah ini "
        "dibuat sebagai contoh antarmuka. Belum ada vendor nyata yang terdaftar."
    )

    f1, f2 = st.columns(2)
    with f1:
        categories = ["Semua kategori"] + sorted(vendors["Kategori"].unique().tolist())
        selected_category = st.selectbox("Kategori", categories, key="vendor_category")
    with f2:
        regions = ["Semua wilayah"] + sorted(vendors["Wilayah Layanan"].unique().tolist())
        selected_region = st.selectbox("Wilayah layanan", regions, key="vendor_region")

    visible_vendors = vendors.copy()
    if selected_category != "Semua kategori":
        visible_vendors = visible_vendors[visible_vendors["Kategori"] == selected_category]
    if selected_region != "Semua wilayah":
        visible_vendors = visible_vendors[visible_vendors["Wilayah Layanan"] == selected_region]

    v1, v2, v3 = st.columns(3)
    v1.metric("Contoh katalog", f"{len(visible_vendors):,}")
    v2.metric("Kategori tersedia", f"{visible_vendors['Kategori'].nunique():,}")
    v3.metric("Vendor terverifikasi", "0")

    st.subheader("Katalog produk dan layanan")
    st.dataframe(
        visible_vendors,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Nama Vendor": st.column_config.TextColumn("Nama Vendor"),
            "Harga Indikatif": st.column_config.TextColumn("Harga Indikatif"),
        },
    )

    st.divider()
    st.subheader("Simulasi permintaan dari petani")
    st.caption(
        "Form ini hanya mendemonstrasikan alur permintaan. Data tidak dikirim "
        "ke vendor atau sistem eksternal."
    )

    with st.form("vendor_request_form", clear_on_submit=False):
        requester = st.text_input("Nama petani / kelompok tani")
        request_region = st.selectbox(
            "Kabupaten/Kota atau wilayah",
            [
                "Lampung Tengah",
                "Lampung Timur",
                "Lampung Selatan",
                "Lampung Utara",
                "Kabupaten/kota lain",
            ],
        )
        request_category = st.selectbox(
            "Kebutuhan",
            ["Benih", "Pupuk", "Alat pertanian", "Pascapanen", "Logistik"],
        )
        request_details = st.text_area(
            "Detail kebutuhan",
            placeholder="Contoh: kebutuhan benih untuk musim tanam berikutnya",
        )
        submitted = st.form_submit_button("Buat permintaan demo")

    if submitted:
        if not requester.strip() or not request_details.strip():
            st.error("Isi nama pemohon dan detail kebutuhan terlebih dahulu.")
        else:
            if "vendor_demo_requests" not in st.session_state:
                st.session_state["vendor_demo_requests"] = []
            st.session_state["vendor_demo_requests"].append({
                "Pemohon": requester.strip(),
                "Wilayah": request_region,
                "Kategori": request_category,
                "Detail": request_details.strip(),
                "Status": "Demo — belum dikirim",
            })
            st.success(
                "Permintaan demo berhasil dibuat di sesi aplikasi ini. "
                "Belum dikirim kepada vendor mana pun."
            )

    demo_requests = st.session_state.get("vendor_demo_requests", [])
    if demo_requests:
        st.subheader("Permintaan demo pada sesi ini")
        st.dataframe(
            pd.DataFrame(demo_requests),
            use_container_width=True,
            hide_index=True,
        )

elif page == "AI Assistant & Learning Center":
    st.header("AI Assistant & Learning Center")
    st.write(
        "Pusat tanya jawab dan pembelajaran pertanian Desaku Maju. "
        "Versi demo ini dapat membaca dokumen SOP yang tersimpan di Knowledge Center "
        "dan menampilkan satu bagian dokumen yang paling relevan."
    )

    tab_ai, tab_learning, tab_kb = st.tabs(
        ["AI Assistant", "Learning Center", "Knowledge Center"]
    )

    # Load the local knowledge folder each time this page is rendered.
    knowledge_docs = load_knowledge_documents()
    readable_docs = [d for d in knowledge_docs if d.get("chunks")]

    with tab_ai:
        st.subheader("Asisten Pertanian")
        st.caption(
            "Jawaban akan didasarkan pada bagian dokumen yang cocok dengan kata kunci. "
            "Ini pencarian sederhana untuk demo, belum menggunakan model AI atau semantic search."
        )

        suggested_questions = [
            "Apa yang perlu diperhatikan setelah panen padi?",
            "Bagaimana pemberian pupuk pada tanaman padi?",
            "Apa yang perlu diperhatikan saat pemupukan jagung?",
            "Bagaimana perawatan tanaman dan pengendalian hama?",
            "Apa saja langkah sebelum dan saat panen?",
        ]
        selected_question = st.selectbox(
            "Pertanyaan contoh",
            ["Pilih pertanyaan contoh..."] + suggested_questions,
            key="ai_suggested_question",
        )
        user_question = st.text_area(
            "Atau tulis pertanyaan Anda",
            placeholder="Contoh: apa yang perlu diperhatikan saat pemberian pupuk?",
            key="ai_user_question",
        )

        if st.button("Cari jawaban dari dokumen", type="primary", key="ai_demo_answer"):
            question = user_question.strip()
            if not question and selected_question != "Pilih pertanyaan contoh...":
                question = selected_question

            if not question:
                st.warning("Pilih pertanyaan contoh atau tulis pertanyaan terlebih dahulu.")
            else:
                st.session_state["desaku_ai_demo_question"] = question
                matches = retrieve_relevant_chunks(question, readable_docs, top_k=1)
                if matches:
                    score, filename, chunk_index, chunk = matches[0]
                    st.markdown("#### Jawaban berdasarkan SOP")
                    st.write(chunk)
                    st.caption(f"Sumber: {filename}, bagian {chunk_index + 1}.")
                    st.info(
                        "Ditampilkan satu bagian yang paling relevan agar jawaban mudah dipahami. "
                        "Bila jawaban belum tepat, coba ubah pertanyaan atau periksa dokumen sumber."
                    )
                    st.caption(
                        "Ikuti panduan varietas dan arahan penyuluh setempat untuk keputusan teknis, "
                        "terutama dosis pupuk dan penggunaan pestisida."
                    )
                else:
                    st.warning(
                        "Belum ditemukan bagian dokumen yang cocok dengan kata kunci pertanyaan. "
                        "Pastikan dokumen SOP sudah ada di Knowledge Center atau coba gunakan kata kunci "
                        "yang muncul di dokumen, misalnya: pupuk, panen, hama, benih, atau pascapanen."
                    )
                st.info(
                    "Catatan: pencarian ini berbasis kecocokan kata, bukan jawaban yang ditulis model AI. "
                    "Jika dokumen tidak membahas pertanyaan, aplikasi tidak akan mengarang jawaban."
                )

        if readable_docs:
            st.caption("Dokumen aktif: " + ", ".join(d["name"] for d in readable_docs))
        else:
            st.warning(
                "Belum ada dokumen yang berhasil dibaca. Buka tab Knowledge Center untuk mengunggah SOP."
            )

    with tab_learning:
        st.subheader("Learning Center")
        st.write("Katalog awal materi penyuluhan yang dapat dikembangkan menjadi LMS.")
        learning_items = pd.DataFrame([
            {
                "Modul": "Budidaya Padi",
                "Topik": "Persiapan lahan, tanam, pemupukan, perawatan, dan panen",
                "Status": "Kerangka demo",
            },
            {
                "Modul": "Budidaya Jagung",
                "Topik": "Penanaman, pemupukan, pemeliharaan, dan panen",
                "Status": "Kerangka demo",
            },
            {
                "Modul": "Penanganan Pascapanen",
                "Topik": "Pencatatan hasil panen dan kebutuhan pengeringan",
                "Status": "Kerangka demo",
            },
            {
                "Modul": "Literasi Data Petani",
                "Topik": "Kualitas data, privasi, dan pembaruan informasi",
                "Status": "Kerangka demo",
            },
        ])
        st.dataframe(learning_items, use_container_width=True, hide_index=True)
        st.caption("Materi ini baru berupa kerangka topik, bukan kursus lengkap atau sertifikasi.")

    with tab_kb:
        st.subheader("Knowledge Center")
        st.write(
            "Unggah dokumen SOP atau panduan pertanian. File akan disimpan ke folder "
            "`knowledge_docs` di samping `app.py`, sehingga tersedia kembali setelah aplikasi dijalankan ulang."
        )
        uploaded_files = st.file_uploader(
            "Unggah dokumen sumber",
            type=["docx", "pdf", "txt", "md"],
            accept_multiple_files=True,
            key="kb_upload_files",
            help="Format yang didukung: Word DOCX, PDF berbasis teks, TXT, dan Markdown."
        )
        if uploaded_files:
            if st.button("Simpan dokumen ke Knowledge Center", type="primary", key="kb_save_files"):
                saved_count = 0
                for uploaded in uploaded_files:
                    # Use the basename only and restrict extensions to supported types.
                    safe_name = Path(uploaded.name).name
                    if Path(safe_name).suffix.lower() not in {".docx", ".pdf", ".txt", ".md"}:
                        continue
                    destination = KB_DIR / safe_name
                    destination.write_bytes(uploaded.getvalue())
                    saved_count += 1
                if saved_count:
                    st.success(f"{saved_count} dokumen disimpan. Aplikasi akan membacanya pada pemuatan berikutnya.")
                    st.rerun()
                else:
                    st.error("Tidak ada dokumen yang berhasil disimpan.")

        # Show the SOP that has been provided alongside any other local documents.
        if knowledge_docs:
            rows = []
            for doc in knowledge_docs:
                rows.append({
                    "Nama dokumen": doc["name"],
                    "Status": "Siap dicari" if doc.get("chunks") else "Perlu diperiksa",
                    "Jumlah bagian": len(doc.get("chunks", [])),
                    "Catatan": doc.get("error", "") or "Teks berhasil diekstrak",
                })
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
            for doc in knowledge_docs:
                if doc.get("text"):
                    with st.expander(f"Pratinjau: {doc['name']}"):
                        st.write(doc["text"][:5000])
                        if len(doc["text"]) > 5000:
                            st.caption("Pratinjau dipotong sampai 5.000 karakter.")
        else:
            st.info(
                "Belum ada dokumen. Unggah file SOP di atas, atau salin "
                "`SOP_Pertanian_Desaku_Maju_Demo.docx` ke folder `knowledge_docs`."
            )
        st.caption(
            "Penting: dokumen SOP demo merupakan bahan awal. Verifikasi dosis pupuk, pestisida, "
            "dan rekomendasi teknis dengan penyuluh atau sumber resmi setempat."
        )

else:
    st.header("Agricultural Intelligence Dashboard")
    st.write(
        "Ringkasan indikator pertanian berdasarkan workbook database Desaku Maju. "
        "Angka merupakan agregasi dari file yang tersedia dan perlu dikonfirmasi "
        "terhadap definisi, periode, serta sumber data aslinya."
    )

    main_files = [
        p for p in excel_files
        if "DATABASE DESAKU MAJU" in p.name.upper()
    ]

    if not main_files:
        st.warning(
            "Workbook database utama belum ditemukan. Pastikan file "
            "'DATABASE DESAKU MAJU ...xlsx' berada di folder yang sama dengan app.py."
        )
    else:
        main_path = main_files[0]
        try:
            main_sheets = load_workbook(
                str(main_path),
                main_path.stat().st_mtime,
            )
            if "GABUNGAN" in main_sheets:
                data = main_sheets["GABUNGAN"].copy()
            else:
                st.error("Sheet 'GABUNGAN' tidak ditemukan di workbook utama.")
                st.stop()
        except Exception as exc:
            st.error(f"Gagal membaca database utama: {exc}")
            st.stop()

        required_cols = [
            "KABUPATEN/KOTA",
            "KECAMATAN",
            "DESA",
            "Luas Lahan Padi",
            "Total Produksi Panen Padi (TON-GKP)",
            "Luas Lahan Jagung",
            "Total Produksi Panen Jagung (Ton)",
            "Luas Lahan Ubi Kayu",
            "Total Produksi Panen Ubi Kayu (Ton)",
        ]
        missing_cols = [col for col in required_cols if col not in data.columns]

        if missing_cols:
            st.error(
                "Kolom yang dibutuhkan tidak ditemukan: "
                + ", ".join(missing_cols)
            )
            st.stop()

        # Pastikan indikator numerik benar-benar numerik.
        numeric_cols = [
            "Luas Lahan Padi",
            "Total Produksi Panen Padi (TON-GKP)",
            "Luas Lahan Jagung",
            "Total Produksi Panen Jagung (Ton)",
            "Luas Lahan Ubi Kayu",
            "Total Produksi Panen Ubi Kayu (Ton)",
        ]
        for col in numeric_cols:
            data[col] = pd.to_numeric(data[col], errors="coerce")

        district_options = ["Semua Kabupaten/Kota"] + sorted(
            data["KABUPATEN/KOTA"].dropna().astype(str).unique().tolist()
        )
        selected_district = st.selectbox(
            "Filter Kabupaten/Kota",
            district_options,
            key="dashboard_district",
        )

        filtered = data.copy()
        if selected_district != "Semua Kabupaten/Kota":
            filtered = filtered[
                filtered["KABUPATEN/KOTA"].astype(str) == selected_district
            ]

        def total_value(column):
            return float(filtered[column].sum(skipna=True))

        village_count = filtered["DESA"].dropna().astype(str).str.strip()
        village_count = village_count[village_count.ne("")].nunique()

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Desa dengan nama tercatat", f"{village_count:,}")
        k2.metric("Luas padi tercatat", f"{total_value('Luas Lahan Padi'):,.0f} ha")
        k3.metric("Luas jagung tercatat", f"{total_value('Luas Lahan Jagung'):,.0f} ha")
        k4.metric("Luas ubi kayu tercatat", f"{total_value('Luas Lahan Ubi Kayu'):,.0f} ha")

        st.divider()
        st.subheader("Produksi komoditas")
        production = pd.DataFrame({
            "Komoditas": ["Padi (GKP)", "Jagung", "Ubi kayu"],
            "Produksi tercatat (ton)": [
                total_value("Total Produksi Panen Padi (TON-GKP)"),
                total_value("Total Produksi Panen Jagung (Ton)"),
                total_value("Total Produksi Panen Ubi Kayu (Ton)"),
            ],
        })
        fig_prod = px.bar(
            production,
            x="Komoditas",
            y="Produksi tercatat (ton)",
            color="Komoditas",
            title="Total produksi tercatat dalam workbook",
            text_auto=".3s",
        )
        fig_prod.update_layout(showlegend=False)
        fig_prod.update_traces(marker_color="#238636")
        st.plotly_chart(fig_prod, use_container_width=True)

        st.subheader("Perbandingan luas lahan")
        land = pd.DataFrame({
            "Komoditas": ["Padi", "Jagung", "Ubi kayu"],
            "Luas tercatat (ha)": [
                total_value("Luas Lahan Padi"),
                total_value("Luas Lahan Jagung"),
                total_value("Luas Lahan Ubi Kayu"),
            ],
        })
        fig_land = px.bar(
            land,
            x="Komoditas",
            y="Luas tercatat (ha)",
            title="Luas lahan tercatat menurut komoditas",
            text_auto=".3s",
        )
        fig_land.update_traces(marker_color="#2878B5")
        st.plotly_chart(fig_land, use_container_width=True)

        st.subheader("Kabupaten/Kota dalam database")
        district_summary = (
            filtered.groupby("KABUPATEN/KOTA", dropna=False)
            .agg(
                Desa=("DESA", lambda s: s.dropna().astype(str).str.strip().replace("", np.nan).nunique()),
                **{
                    "Luas Padi (ha)": ("Luas Lahan Padi", "sum"),
                    "Produksi Padi (ton)": ("Total Produksi Panen Padi (TON-GKP)", "sum"),
                    "Luas Jagung (ha)": ("Luas Lahan Jagung", "sum"),
                    "Produksi Jagung (ton)": ("Total Produksi Panen Jagung (Ton)", "sum"),
                    "Luas Ubi Kayu (ha)": ("Luas Lahan Ubi Kayu", "sum"),
                    "Produksi Ubi Kayu (ton)": ("Total Produksi Panen Ubi Kayu (Ton)", "sum"),
                },
            )
            .reset_index()
            .sort_values("KABUPATEN/KOTA")
        )
        st.dataframe(district_summary, use_container_width=True, hide_index=True)

        st.caption(
            f"Sumber: {main_path.name}, sheet GABUNGAN. "
            "Nilai dijumlahkan dari kolom yang tersedia; angka nol dan data kosong "
            "dipertahankan sesuai isi sumber. Luas lahan antar-komoditas tidak "
            "dijumlahkan menjadi total luas unik karena dapat tumpang tindih."
        )

    st.divider()
    st.subheader("Workbook tersedia")
    for workbook in excel_files:
        st.write(f"📄 {workbook.name}")

st.divider()
st.caption("Desaku Maju | Prototipe demo — data dan fitur perlu diverifikasi.")
