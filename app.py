"""
app.py — Aplikasi Dokumentasi Survey Lapangan
Semua pemrosesan dilakukan in-memory. Tidak ada foto atau data yang disimpan di server.
"""

import streamlit as st
from datetime import date
from utils.doc_builder import build_docx
from utils.preview import docx_to_html

HARI_ID = {
    "Monday": "Senin", "Tuesday": "Selasa", "Wednesday": "Rabu",
    "Thursday": "Kamis", "Friday": "Jumat", "Saturday": "Sabtu", "Sunday": "Minggu"
}
BULAN_ID = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember"
}

def format_tanggal_id(d) -> str:
    """Mengubah objek date menjadi format '16 Juni 2026'."""
    nama_hari = HARI_ID[d.strftime("%A")]
    return f"{nama_hari}, {d.day} {BULAN_ID[d.month]} {d.year}"

# ── Konfigurasi halaman ────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Dokumentasi Survey CAPEX",
    page_icon="📋",
    layout="wide",
)

# ── CSS kustom ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Warna utama PAM JAYA — biru air */
    :root {
        --primary: #005B99;
        --primary-light: #E8F4FB;
        --accent: #00A3D9;
        --border: #C8DFF0;
        --text-muted: #6B8FA8;
        --danger-bg: #FFF3F3;
        --danger-border: #FFBCBC;
    }

    /* Header aplikasi */
    .app-header {
        background: linear-gradient(135deg, #005B99 0%, #00A3D9 100%);
        color: white;
        padding: 20px 28px;
        border-radius: 10px;
        margin-bottom: 28px;
    }
    .app-header h1 {
        margin: 0;
        font-size: 1.5rem;
        font-weight: 700;
        letter-spacing: 0.3px;
    }
    .app-header p {
        margin: 4px 0 0 0;
        font-size: 0.85rem;
        opacity: 0.85;
    }

    /* Kartu seksi */
    .section-card {
        background: white;
        border: 1.5px solid var(--border);
        border-radius: 8px;
        padding: 20px 24px;
        margin-bottom: 20px;
    }
    .section-title {
        font-size: 0.95rem;
        font-weight: 700;
        color: var(--primary);
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 14px;
        padding-bottom: 8px;
        border-bottom: 2px solid var(--primary-light);
    }

    /* Badge jumlah foto */
    .photo-badge {
        display: inline-block;
        background: var(--primary-light);
        color: var(--primary);
        font-size: 0.78rem;
        font-weight: 600;
        padding: 2px 10px;
        border-radius: 12px;
        margin-left: 8px;
        vertical-align: middle;
    }

    /* Info box */
    .info-box {
        background: var(--primary-light);
        border-left: 3px solid var(--accent);
        border-radius: 0 6px 6px 0;
        padding: 10px 14px;
        font-size: 0.83rem;
        color: #1a3a52;
        margin-bottom: 14px;
    }

    /* Warning box */
    .warn-box {
        background: var(--danger-bg);
        border-left: 3px solid #E05252;
        border-radius: 0 6px 6px 0;
        padding: 10px 14px;
        font-size: 0.83rem;
        color: #7a1c1c;
        margin-bottom: 14px;
    }

    /* Thumbnail grid */
    .thumb-label {
        font-size: 0.75rem;
        color: var(--text-muted);
        text-align: center;
        margin-top: 4px;
    }

    /* Tombol download */
    .stDownloadButton > button {
        background: #005B99 !important;
        color: white !important;
        border: none !important;
        border-radius: 6px !important;
        padding: 10px 28px !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        width: 100%;
    }
    .stDownloadButton > button:hover {
        background: #004477 !important;
    }

    /* Sembunyikan label upload bawaan */
    [data-testid="stFileUploaderDropzoneInstructions"] {
        font-size: 0.82rem;
    }

    /* Streamlit default adjustments */
    .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
    .stTextInput label, .stSelectbox label, .stDateInput label { font-size: 0.85rem; }
</style>
""", unsafe_allow_html=True)


# ── Konstanta validasi ─────────────────────────────────────────────────────────
ALLOWED_TYPES   = ["image/jpeg", "image/jpg", "image/png"]
ALLOWED_EXT     = [".jpg", ".jpeg", ".png"]
MAX_FILE_MB     = 10
MAX_FILE_BYTES  = MAX_FILE_MB * 1024 * 1024


# ── Fungsi validasi foto ───────────────────────────────────────────────────────
def validate_photos(files: list, label: str) -> tuple[list, list]:
    """
    Validasi tipe dan ukuran foto.
    Kembalikan (valid_bytes_list, error_messages).
    File bytes dibaca sekali lalu dibuang — tidak disimpan di session_state.
    """
    valid, errors = [], []
    for f in files:
        ext = "." + f.name.rsplit(".", 1)[-1].lower() if "." in f.name else ""
        if ext not in ALLOWED_EXT:
            errors.append(f"❌ **{f.name}**: format tidak didukung (gunakan JPG/PNG)")
            continue
        raw = f.read()
        if len(raw) > MAX_FILE_BYTES:
            errors.append(f"❌ **{f.name}**: ukuran melebihi {MAX_FILE_MB}MB")
            continue
        valid.append(raw)
    return valid, errors


# ── Header aplikasi ────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
    <h1>📋 Dokumentasi Survey Lapangan</h1>
    <p>CAPEX MANAGEMENT · Semua data diproses lokal, tidak tersimpan di server</p>
</div>
""", unsafe_allow_html=True)

# ── Layout utama: form kiri, preview kanan ─────────────────────────────────────
col_form, col_preview = st.columns([1, 1], gap="large")

# ══════════════════════════════════════════════════════════════════════════════
# KOLOM KIRI — Form input
# ══════════════════════════════════════════════════════════════════════════════
with col_form:

    # ── Seksi 1: Data Kegiatan ─────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown('<div class="section-title">① Data Kegiatan</div>', unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            hari_tanggal = st.date_input("Hari / Tanggal", value=date.today())
        with c2:
            jenis_kontrak = st.text_input("Jenis Kontrak", placeholder="Contoh: Non Frame Kontrak")
    
        nama_kegiatan = st.text_input("Nama Kegiatan", placeholder="Contoh: Rehabilitasi Pipa DN 200 mm")
        lokasi        = st.text_input("Lokasi", placeholder="Contoh: Jl. Sudirman No. 10, Jakarta Pusat")
    
        c3, c4 = st.columns(2)
        with c3:
            divisi    = st.text_input("Divisi",    placeholder="Contoh: Divisi Produksi")
        with c4:
            subdivisi = st.text_input("Subdivisi", placeholder="Contoh: Subdivisi Distribusi")
    
        pjk = st.text_input("PJK (Penanggung Jawab Kegiatan)", placeholder="Nama PJK")
    
        c5, c6 = st.columns(2)
        with c5:
            nilai_terkontrak       = st.text_input("Nilai Terkontrak",                  placeholder="Contoh: 147.990.910")
        with c6:
            nilai_terkontrak_final = st.text_input("Nilai Terkontrak Final (Addendum)", placeholder="Contoh: 147.990.910")
    
        c7, c8 = st.columns(2)
        with c7:
            vendor   = st.text_input("Vendor",        placeholder="Nama perusahaan")
        with c8:
            progress = st.text_input("Progress Lap-%", placeholder="Contoh: 75%")
    
        st.markdown('</div>', unsafe_allow_html=True)
    
    # ── Seksi 2: Foto Biasa ────────────────────────────────────────────────────
    with st.container(border=True):

        include_apd = st.toggle("📎 Sertakan foto Penggunaan APD", value=False)
        max_biasa   = 8 if include_apd else 10
    
        st.markdown(
            f'<div class="section-title">② Foto Survey '
            f'<span class="photo-badge">Maks {max_biasa} foto</span></div>',
            unsafe_allow_html=True
        )
        st.markdown(
            f'<div class="info-box">Upload foto dokumentasi lapangan. '
            f'Format JPG/PNG, maks {MAX_FILE_MB}MB per foto. '
            f'{"Karena foto APD diaktifkan, slot foto biasa maks 8." if include_apd else "Maks 10 foto jika tidak ada foto APD."}'
            f'</div>',
            unsafe_allow_html=True
        )
    
        uploaded_biasa = st.file_uploader(
            "Pilih foto",
            type=["jpg", "jpeg", "png"],
            accept_multiple_files=True,
            key="upload_biasa",
            label_visibility="collapsed",
        )
    
        # Validasi & preview thumbnail foto biasa
        valid_biasa, errors_biasa = [], []
        if uploaded_biasa:
            valid_biasa, errors_biasa = validate_photos(uploaded_biasa, "foto biasa")
    
            if len(valid_biasa) > max_biasa:
                st.markdown(
                    f'<div class="warn-box">⚠️ Hanya {max_biasa} foto pertama yang akan digunakan '
                    f'({len(valid_biasa)} foto dipilih).</div>',
                    unsafe_allow_html=True
                )
                valid_biasa = valid_biasa[:max_biasa]
    
            for err in errors_biasa:
                st.markdown(f'<div class="warn-box">{err}</div>', unsafe_allow_html=True)
    
            if valid_biasa:
                st.caption(f"✅ {len(valid_biasa)} foto siap digunakan")
                cols_thumb = st.columns(min(len(valid_biasa), 5))
                for i, raw in enumerate(valid_biasa):
                    with cols_thumb[i % 5]:
                        st.image(raw, use_container_width=True)
                        st.markdown(f'<div class="thumb-label">Foto {i+1}</div>', unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Seksi 3: Foto APD (kondisional) ───────────────────────────────────────
    valid_apd = []
    if include_apd:
       with st.container(border=True):
        st.markdown(
            '<div class="section-title">③ Foto Penggunaan APD '
            '<span class="photo-badge">Maks 2 foto</span></div>',
            unsafe_allow_html=True
        )
        st.markdown(
            '<div class="info-box">Upload 1–2 foto penggunaan APD. '
            'Akan ditempatkan di bagian bawah dokumen dengan label "Penggunaan APD".</div>',
            unsafe_allow_html=True
        )

        uploaded_apd = st.file_uploader(
            "Pilih foto APD",
            type=["jpg", "jpeg", "png"],
            accept_multiple_files=True,
            key="upload_apd",
            label_visibility="collapsed",
        )

        if uploaded_apd:
            valid_apd, errors_apd = validate_photos(uploaded_apd, "foto APD")

            if len(valid_apd) > 2:
                st.markdown(
                    '<div class="warn-box">⚠️ Hanya 2 foto APD pertama yang akan digunakan.</div>',
                    unsafe_allow_html=True
                )
                valid_apd = valid_apd[:2]

            for err in errors_apd:
                st.markdown(f'<div class="warn-box">{err}</div>', unsafe_allow_html=True)

            if valid_apd:
                st.caption(f"✅ {len(valid_apd)} foto APD siap digunakan")
                cols_apd = st.columns(2)
                for i, raw in enumerate(valid_apd):
                    with cols_apd[i]:
                        st.image(raw, use_container_width=True)
                        st.markdown(f'<div class="thumb-label">APD {i+1}</div>', unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# KOLOM KANAN — Preview & Download
# ══════════════════════════════════════════════════════════════════════════════
with col_preview:
    with st.container(border=True):
        st.markdown('<div class="section-title">④ Preview & Unduh Dokumen</div>', unsafe_allow_html=True)

        # Cek kelengkapan minimal
        has_required = bool(nama_kegiatan and lokasi and valid_biasa)

        if not has_required:
            st.markdown("""
            <div class="info-box">
                Isi minimal <strong>Nama Kegiatan</strong>, <strong>Lokasi</strong>, dan upload minimal
                <strong>1 foto</strong> untuk mengaktifkan preview.
            </div>
            """, unsafe_allow_html=True)
            st.markdown("""
            <div style="text-align:center; padding: 60px 20px; color: #6B8FA8;">
                <div style="font-size:3rem;">📄</div>
                <div style="margin-top:12px; font-size:0.9rem;">Preview dokumen akan muncul di sini</div>
            </div>
            """, unsafe_allow_html=True)
        else:
            # Susun data
            form_data = {
                "hari_tanggal": format_tanggal_id(hari_tanggal),
                "nama_kegiatan":           nama_kegiatan,
                "jenis_kontrak":           jenis_kontrak,
                "lokasi":                  lokasi,
                "divisi":                  divisi,
                "subdivisi":               subdivisi,
                "pjk":                     pjk,
                "nilai_terkontrak":       f"Rp {nilai_terkontrak}" if nilai_terkontrak else "",
                "nilai_terkontrak_final": f"Rp {nilai_terkontrak_final}" if nilai_terkontrak_final else "",
                "vendor":                  vendor,
                "progress":                f"{progress}%" if progress else "",
            }

            # Generate docx in-memory
            with st.spinner("Membuat dokumen..."):
                try:
                    docx_bytes = build_docx(
                        data        = form_data,
                        photos      = valid_biasa,
                        apd_photos  = valid_apd,
                        include_apd = include_apd and len(valid_apd) > 0,
                    )

                    # Preview HTML
                    html_preview = docx_to_html(docx_bytes)

                    st.markdown("**Preview Dokumen:**")
                    st.components.v1.html(html_preview, height=700, scrolling=True)

                    st.markdown("<br>", unsafe_allow_html=True)

                    # Nama file output
                    tgl_file  = format_tanggal_id(hari_tanggal)
                    filename  = f"Dokumentasi {nama_kegiatan}_{tgl_file}.docx"

                    st.download_button(
                        label    = "⬇️  Unduh Dokumen Word (.docx)",
                        data     = docx_bytes,
                        file_name= filename,
                        mime     = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )

                    st.markdown(
                        '<div class="info-box" style="margin-top:10px; font-size:0.78rem;">'
                        '🔒 Dokumen dibuat langsung di perangkat Anda. '
                        'Foto dan data tidak tersimpan di server setelah sesi ini berakhir.'
                        '</div>',
                        unsafe_allow_html=True
                    )

                except Exception as e:
                    st.error(f"Gagal membuat dokumen: {e}")

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="text-align:center; color:#6B8FA8; font-size:0.75rem; margin-top:32px; padding-bottom:16px;">
    Dokumentasi Survey CAPEX Management 2026
</div>
""", unsafe_allow_html=True)
