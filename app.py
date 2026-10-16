"""CDSS Skrining Risiko Anemia (Rule-Based) + database SQLite (CRUD lengkap).

Aturan sama dengan file Excel dan PowerPoint:
  R1: G1 dan G2                -> Risiko Sedang
  R2: G1, G2, dan G3           -> Risiko Tinggi
  R3: G4 dan G5                -> Risiko Sedang
  R4: G1, G2, G3, dan G5       -> Risiko Tinggi
Risiko Tinggi jika R2 atau R4 terpenuhi; Risiko Sedang jika R1 atau R3; selain itu Risiko Rendah.
Contoh simulasi tugas, bukan alat diagnosis.
"""
import sqlite3
from contextlib import closing
from datetime import datetime

import pandas as pd
import streamlit as st

DB_NAME = "skrining_anemia_v2.db"

# ----------------------------------------------------------------------------
# KNOWLEDGE BASE
# ----------------------------------------------------------------------------
INDIKATOR = [
    ("G1", "Mudah lelah"),
    ("G2", "Sering pusing/lemas"),
    ("G3", "Pucat"),
    ("G4", "Pola makan kurang baik"),
    ("G5", "Jarang mengonsumsi sumber zat besi"),
    ("G6", "Riwayat anemia"),
]
# (kode, uraian, indeks indikator yang harus bernilai 1, hasil)
ATURAN = [
    ("R1", "IF G1=1 AND G2=1", (0, 1), "Risiko Sedang"),
    ("R2", "IF G1=1 AND G2=1 AND G3=1", (0, 1, 2), "Risiko Tinggi"),
    ("R3", "IF G4=1 AND G5=1", (3, 4), "Risiko Sedang"),
    ("R4", "IF G1=1 AND G2=1 AND G3=1 AND G5=1", (0, 1, 2, 4), "Risiko Tinggi"),
]


def format_uraian_rule(uraian):
    teks = uraian.replace("IF ", "").replace("=1", "")
    for kode, label in INDIKATOR:
        teks = teks.replace(kode, label)
    teks = teks.replace(" AND ", " dan ")
    return teks
BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]
REKOMENDASI = {
    "Risiko Tinggi": "Sangat disarankan segera berkonsultasi dengan tenaga kesehatan untuk pemeriksaan darah lengkap dan penanganan lebih lanjut.",
    "Risiko Sedang": "Pantau kondisi, tingkatkan konsumsi makanan sumber zat besi, dan pertimbangkan konsultasi ke fasilitas kesehatan.",
    "Risiko Rendah": "Belum ada aturan yang terpenuhi. Pertahankan pola makan seimbang dan istirahat cukup.",
}

CUSTOM_CSS = """
<style>
    :root {
        --medical-bg: #f5f7f2;
        --medical-hero1: #315b4b;
        --medical-hero2: #56806a;
        --medical-accent: #426b56;
        --medical-blue: #6a8f73;
        --medical-navy: #263b32;
        --medical-line: #e1e8df;
    }

    html, body, [class*="css"], [data-testid="stAppViewContainer"] {
        color: var(--medical-navy);
        font-family: "Segoe UI", system-ui, -apple-system, sans-serif;
    }
    [data-testid="stAppViewContainer"] { background: linear-gradient(180deg, #fafbf8, var(--medical-bg)); }
    [data-testid="stHeader"] { background-color: #f5f9f7; }
    .main .block-container { padding: 2rem 1rem 3rem; }
    h1, h2, h3, h4 { color: var(--medical-navy); font-weight: 700; letter-spacing: -0.025em; }
    p, label, [data-testid="stMarkdownContainer"] { line-height: 1.6; }

    .hero-card {
        padding: 1.65rem 1.8rem;
        margin-bottom: 1rem;
        border-radius: 22px;
        background: linear-gradient(125deg, var(--medical-hero1), var(--medical-hero2));
        box-shadow: 0 12px 28px rgba(49, 91, 75, 0.16);
    }
    .hero-card h2 { margin: 0 0 .45rem; color: #fff; font-size: clamp(1.45rem, 3vw, 2rem); }
    .hero-card p { margin: 0; color: rgba(255, 255, 255, .9); }
    .hero-card .hello {
        display: inline-block;
        padding: .25rem .8rem;
        border-radius: 999px;
        background: rgba(255, 255, 255, .17);
        color: #fff;
        font-size: .85rem;
    }
    .hero-card .hello { margin-bottom: .6rem; }

    .panel {
        padding: 1rem 1.1rem;
        margin-bottom: 1rem;
        border: 1px solid var(--medical-line);
        border-left: 5px solid var(--medical-blue);
        border-radius: 16px;
        background: #fff;
        box-shadow: 0 4px 14px rgba(38, 59, 50, .05);
        line-height: 1.55;
    }
    .small-note { color: #65756a; font-size: .9rem; }

    .stSelectbox label, .stSelectbox label p,
    .stTextInput label, .stTextInput label p,
    .stCheckbox label, .stCheckbox label p {
        color: var(--medical-accent) !important;
        font-weight: 600;
    }
    div[data-baseweb="select"] > div, div[data-baseweb="input"],
    div[data-baseweb="base-input"], .stTextInput input {
        border-color: #cfe0d8 !important;
        background: #fff !important;
        color: var(--medical-navy) !important;
    }
    div[data-baseweb="select"] svg { fill: var(--medical-accent) !important; }
    div[data-baseweb="select"] input { color: var(--medical-navy) !important; -webkit-text-fill-color: var(--medical-navy); }
    .stTextInput input:disabled {
        background: #edf3ed !important;
        color: var(--medical-accent) !important;
        -webkit-text-fill-color: var(--medical-accent);
        opacity: 1;
        font-weight: 600;
    }
    .stTextInput input:focus, div[data-baseweb="select"] > div:focus-within {
        border-color: var(--medical-accent) !important;
        box-shadow: 0 0 0 .18rem rgba(66, 107, 86, .15);
    }
    div[data-baseweb="popover"] ul, ul[role="listbox"], div[data-baseweb="menu"],
    li[role="option"] { background: #fff !important; color: var(--medical-navy) !important; }
    li[role="option"]:hover, li[role="option"][aria-selected="true"] {
        background: #e3f1ec !important;
        color: #1f4d3d !important;
    }
    [data-baseweb="checkbox"] > span:first-child {
        border: 2px solid #7fb5a0 !important;
        border-radius: 6px !important;
        background: #fff !important;
    }
    [data-baseweb="checkbox"] input:checked + span,
    [data-baseweb="checkbox"] [aria-checked="true"] > span:first-child {
        border-color: #2f8f78 !important;
        background: #2f8f78 !important;
    }
    [data-baseweb="checkbox"] span { color: var(--medical-navy); }

    [data-testid="stTabs"] button[role="tab"],
    [data-testid="stTabs"] button[role="tab"] p {
        color: #263b32 !important;
        font-weight: 600 !important;
    }
    [data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
        color: var(--medical-accent) !important;
        font-weight: 800 !important;
    }
    [data-testid="stTabs"] [data-baseweb="tab-highlight"] { background: var(--medical-accent) !important; }
    [data-testid="stTabs"] [data-baseweb="tab-border"] { background: #d5e3dc !important; }
    .stButton > button, .stDownloadButton > button {
        border-radius: 11px;
        font-family: inherit;
        font-weight: 700 !important;
        transition: background-color 150ms ease, border-color 150ms ease, transform 150ms ease;
    }
    .stButton > button, .stButton > button p, .stDownloadButton > button,
    .stDownloadButton > button p { color: #263b32 !important; }
    .stButton > button[kind="primary"] {
        border: 1.5px solid #2f8f78 !important;
        background: #d3ece2 !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover { border-color: var(--medical-accent) !important; }
    .stDataFrame, [data-testid="stTable"] {
        overflow: hidden;
        border: 1px solid var(--medical-line);
        border-radius: 12px;
        background: #fff;
    }
    [data-testid="stTable"] table { width: 100%; background: #fff; color: var(--medical-navy); }
    [data-testid="stTable"] th { background: #e3f1ec !important; color: #1f4d3d !important; }
    [data-testid="stTable"] td { border-color: var(--medical-line) !important; color: var(--medical-navy) !important; }
    [data-testid="stAlert"] { border-radius: 12px; }
    /* ---- paksa terang walau browser/Streamlit memakai mode gelap ---- */
    :root, html, body, .stApp { color-scheme: light !important; }
    .stApp, [data-testid="stApp"] { background: #f5f9f7 !important; color: #263b32 !important; }
    [data-testid="stSelectbox"] div[data-baseweb="select"] > div,
    [data-testid="stSelectbox"] div[data-baseweb="select"] > div > div,
    [data-testid="stSelectbox"] [data-baseweb="select"] div {
        background-color: #ffffff !important; color: #263b32 !important;
        -webkit-text-fill-color: #263b32 !important;
    }
    [data-testid="stSelectbox"] [data-baseweb="select"] > div { border: 1.5px solid #9fc7b8 !important; border-radius: 10px !important; }
    [data-testid="stSelectbox"] input::placeholder, [data-baseweb="select"] [aria-live] { color: #6b7f75 !important; }
    [data-testid="stTextInput"] input { background:#ffffff !important; color:#263b32 !important; -webkit-text-fill-color:#263b32 !important; }
    [data-testid="stTextInput"] input:disabled { background:#e8f3ee !important; -webkit-text-fill-color:#1f4d3d !important; }
    [data-testid="stTextInput"] [data-baseweb="input"] { border: 1.5px solid #9fc7b8 !important; border-radius:10px !important; background:#fff !important; }
    [data-testid="stCheckbox"] label span, [data-testid="stCheckbox"] label p { color:#263b32 !important; }
    [data-testid="stCheckbox"] label > span:first-child { background:#ffffff !important; border:2px solid #7fb5a0 !important; }
    [data-testid="stCheckbox"] label:has(input:checked) > span:first-child { background:#2f8f78 !important; border-color:#2f8f78 !important; }
    [data-testid="stTab"], [data-testid="stTab"] *, [data-testid="stTabs"] [role="tab"] *,
    button[role="tab"], button[role="tab"] * { color:#111111 !important; opacity:1 !important; -webkit-text-fill-color:#111111 !important; }
    button[role="tab"][aria-selected="true"], button[role="tab"][aria-selected="true"] * { font-weight:800 !important; }
    [data-testid="stTab"][aria-selected="true"] ~ [data-baseweb="tab-highlight"], [data-baseweb="tab-highlight"] { background-color:#111111 !important; }
    .stButton button, .stButton button *, .stDownloadButton button, .stDownloadButton button * {
        color:#111111 !important; -webkit-text-fill-color:#111111 !important; opacity:1 !important; }
    .stButton button[kind="primary"], [data-testid="stBaseButton-primary"] { background:#d3ece2 !important; border:1.5px solid #2f8f78 !important; }
    [data-testid="stBaseButton-secondary"] { background:#ffffff !important; border:1.5px solid #9fc7b8 !important; }
    [data-testid="stTable"], [data-testid="stTable"] * { color:#263b32 !important; opacity:1 !important; background-color:#ffffff; }
    [data-testid="stTable"] th { background-color:#e3f1ec !important; color:#1f4d3d !important; }
    h1,h2,h3,h4,[data-testid="stHeading"] * { color:#263b32 !important; }
    .hero-card h2, .hero-card p, .hero-card .hello { color:#ffffff !important; }
</style>
"""


def evaluasi(gejala):
    """gejala: list 6 nilai bool/int. Mengembalikan (hasil R1..R4, tingkat risiko, rekomendasi)."""
    g = [1 if x else 0 for x in gejala]
    r = [1 if all(g[i] for i in idx) else 0 for _, _, idx, _ in ATURAN]
    if r[1] or r[3]:
        tingkat = "Risiko Tinggi"
    elif r[0] or r[2]:
        tingkat = "Risiko Sedang"
    else:
        tingkat = "Risiko Rendah"
    return r, tingkat, REKOMENDASI[tingkat]


def hitung_usia(bulan_lahir, tahun_lahir, sekarang=None):
    """Mengembalikan (tahun, bulan) atau None bila tanggal lahir belum terjadi."""
    n = sekarang or datetime.now()
    total = (n.year - int(tahun_lahir)) * 12 + (n.month - int(bulan_lahir))
    if total < 0:
        return None
    return total // 12, total % 12


def teks_usia(bulan_lahir, tahun_lahir):
    u = hitung_usia(bulan_lahir, tahun_lahir)
    return "-" if u is None else f"{u[0]} th {u[1]} bln"


# ----------------------------------------------------------------------------
# DATABASE (SQLite) - CREATE, READ, UPDATE, DELETE
# ----------------------------------------------------------------------------
def koneksi():
    return sqlite3.connect(DB_NAME)


def init_db():
    with closing(koneksi()) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS skrining_anemia (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nama TEXT NOT NULL,
                jenis_kelamin TEXT NOT NULL,
                bulan_lahir INTEGER NOT NULL,
                tahun_lahir INTEGER NOT NULL,
                g1 INTEGER DEFAULT 0,
                g2 INTEGER DEFAULT 0,
                g3 INTEGER DEFAULT 0,
                g4 INTEGER DEFAULT 0,
                g5 INTEGER DEFAULT 0,
                g6 INTEGER DEFAULT 0,
                r1 INTEGER DEFAULT 0,
                r2 INTEGER DEFAULT 0,
                r3 INTEGER DEFAULT 0,
                r4 INTEGER DEFAULT 0,
                tingkat_risiko TEXT NOT NULL,
                rekomendasi TEXT,
                tanggal_dibuat TEXT,
                tanggal_diubah TEXT
            )
            """
        )
        conn.commit()


def _sekarang():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def simpan_baru(nama, jk, bulan, tahun, gejala):
    r, tingkat, rek = evaluasi(gejala)
    g = [1 if x else 0 for x in gejala]
    t = _sekarang()
    with closing(koneksi()) as conn:
        conn.execute(
            """INSERT INTO skrining_anemia
               (nama, jenis_kelamin, bulan_lahir, tahun_lahir, g1,g2,g3,g4,g5,g6, r1,r2,r3,r4,
                tingkat_risiko, rekomendasi, tanggal_dibuat, tanggal_diubah)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (nama, jk, int(bulan), int(tahun), *g, *r, tingkat, rek, t, t),
        )
        conn.commit()


def ubah_data(id_pasien, nama, jk, bulan, tahun, gejala):
    r, tingkat, rek = evaluasi(gejala)
    g = [1 if x else 0 for x in gejala]
    with closing(koneksi()) as conn:
        conn.execute(
            """UPDATE skrining_anemia SET nama=?, jenis_kelamin=?, bulan_lahir=?, tahun_lahir=?,
               g1=?,g2=?,g3=?,g4=?,g5=?,g6=?, r1=?,r2=?,r3=?,r4=?,
               tingkat_risiko=?, rekomendasi=?, tanggal_diubah=? WHERE id=?""",
            (nama, jk, int(bulan), int(tahun), *g, *r, tingkat, rek, _sekarang(), int(id_pasien)),
        )
        conn.commit()


def ambil_semua():
    with closing(koneksi()) as conn:
        return pd.read_sql_query("SELECT * FROM skrining_anemia ORDER BY id DESC", conn)


def hapus_data(id_pasien):
    with closing(koneksi()) as conn:
        conn.execute("DELETE FROM skrining_anemia WHERE id = ?", (int(id_pasien),))
        conn.commit()


CONTOH = {
    "M001": [1, 1, 1, 0, 1, 0], "M002": [0, 0, 0, 0, 0, 0], "M003": [1, 1, 0, 1, 1, 0],
    "M004": [0, 1, 0, 1, 1, 0], "M005": [1, 0, 0, 0, 0, 1],
}


def isi_contoh():
    """Mengisi lima pasien contoh (nama fiktif)."""
    for nama, g in CONTOH.items():
        simpan_baru(nama, "Perempuan", 1, 2000, g)


def kode_gejala(row):
    k = [INDIKATOR[i][1] for i in range(6) if row[f"g{i + 1}"]]
    return ", ".join(k) if k else "-"


# ----------------------------------------------------------------------------
# TAMPILAN
# ----------------------------------------------------------------------------
def tampil_hasil(r, tingkat, rek):
    for (_, uraian, _, hasil), ok in zip(ATURAN, r):
        st.write(("✅ " if ok else "⬜ ") + f"{format_uraian_rule(uraian)} → " +
                 (f"terpenuhi ({hasil})" if ok else "tidak terpenuhi"))
    teks = f"**{tingkat}**\n\n{rek}"
    if tingkat == "Risiko Tinggi":
        st.error(teks)
    elif tingkat == "Risiko Sedang":
        st.warning(teks)
    else:
        st.success(teks)


def pilih_tanggal_lahir(prefix, bulan_awal=None, tahun_awal=None):
    tahun_ini = datetime.now().year
    tahun_list = list(range(tahun_ini, 1919, -1))
    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        bln_nama = st.selectbox(
            "Bulan lahir",
            BULAN,
            index=bulan_awal - 1 if bulan_awal is not None else None,
            placeholder="Pilih bulan",
            key=f"{prefix}_bln",
        )
    with c2:
        thn = st.selectbox(
            "Tahun lahir",
            tahun_list,
            index=tahun_list.index(tahun_awal) if tahun_awal in tahun_list else None,
            placeholder="Pilih tahun",
            key=f"{prefix}_thn",
        )
    bln = BULAN.index(bln_nama) + 1 if bln_nama is not None else None
    u = hitung_usia(bln, thn) if bln is not None and thn is not None else None
    with c3:
        if bln is None or thn is None:
            usia = "Pilih bulan dan tahun lahir"
        elif u is None:
            usia = "Tanggal lahir belum terjadi"
        else:
            usia = f"{u[0]} tahun {u[1]} bulan"
        st.text_input(
            "Usia",
            value=usia,
            disabled=True,
            key=f"{prefix}_usia_{bln}_{thn}",
        )
    return bln, thn, bln is not None and thn is not None and u is not None


def gejala_checkbox(prefix, nilai_awal=None):
    nilai_awal = nilai_awal or [0] * 6
    kiri, kanan = st.columns(2)
    hasil = []
    for i, (_, label) in enumerate(INDIKATOR):
        with (kiri if i < 3 else kanan):
            hasil.append(st.checkbox(label, value=bool(nilai_awal[i]), key=f"{prefix}_g{i}"))
    return hasil


def main():
    st.set_page_config(page_title="CDSS Skrining Anemia", page_icon="🩺", layout="wide")
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    init_db()
    st.markdown(
        """
        <div class="hero-card">
            <span class="hello">👋 Selamat datang</span>
            <h2>🩺 CDSS Skrining Risiko Anemia</h2>
            <p>Sistem pendukung keputusan berbasis aturan untuk membantu mengenali risiko anemia secara dini.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    kolom1, kolom2, kolom3 = st.columns(3)
    with kolom1:
        st.markdown("<div class='panel'><b>Pasien</b><br><span class='small-note'>Data diri dan usia</span></div>", unsafe_allow_html=True)
    with kolom2:
        st.markdown("<div class='panel'><b>Gejala</b><br><span class='small-note'>Input faktor risiko</span></div>", unsafe_allow_html=True)
    with kolom3:
        st.markdown("<div class='panel'><b>Hasil</b><br><span class='small-note'>Klasifikasi risiko</span></div>", unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["📝 Form Skrining", "📁 Riwayat & Kelola Data", "📚 Knowledge Base"])

    # ---- TAB 1: form skrining (CREATE) ----
    with tab1:
        st.subheader("Data pasien")
        a, b = st.columns(2)
        with a:
            nama = st.text_input("Nama pasien", key="baru_nama")
        with b:
            jk = st.selectbox("Jenis kelamin", ["Laki-laki", "Perempuan"], key="baru_jk")
        bln, thn, tanggal_valid = pilih_tanggal_lahir("baru")

        st.subheader("Gejala dan faktor risiko")
        gejala = gejala_checkbox("baru")

        st.subheader("Hasil sementara")
        r, tingkat, rek = evaluasi(gejala)
        tampil_hasil(r, tingkat, rek)

        if st.button("💾 Simpan hasil ke database", type="primary", key="baru_simpan"):
            if not nama.strip():
                st.warning("Isi nama pasien terlebih dahulu.")
            elif bln is None or thn is None:
                st.warning("Pilih bulan dan tahun lahir terlebih dahulu.")
            elif not tanggal_valid:
                st.warning("Tanggal lahir belum terjadi. Periksa kembali.")
            else:
                simpan_baru(nama.strip(), jk, bln, thn, gejala)
                st.success(f"Data {nama.strip()} tersimpan. Lihat di tab Riwayat.")

    # ---- TAB 2: riwayat (READ, UPDATE, DELETE) ----
    with tab2:
        df = ambil_semua()
        if df.empty:
            st.info("Belum ada data di database.")
            if st.button("Isi 5 data contoh (M001 sampai M005)", key="seed"):
                isi_contoh()
                st.rerun()
        else:
            tampil = pd.DataFrame({
                "ID": df["id"],
                "Nama": df["nama"],
                "Jenis kelamin": df["jenis_kelamin"],
                "Usia": [teks_usia(b_, t_) for b_, t_ in zip(df["bulan_lahir"], df["tahun_lahir"])],
                "Gejala": [kode_gejala(row) for _, row in df.iterrows()],
                "Aturan terpenuhi": [", ".join(f"R{i}" for i in range(1, 5) if row[f"r{i}"]) or "-" for _, row in df.iterrows()],
                "Hasil": df["tingkat_risiko"],
                "Dibuat": df["tanggal_dibuat"],
                "Diubah": df["tanggal_diubah"],
            })
            st.dataframe(tampil, use_container_width=True, hide_index=True)
            st.download_button("⬇️ Unduh data (CSV)", tampil.to_csv(index=False).encode("utf-8"),
                               file_name="riwayat_skrining_anemia.csv", mime="text/csv")

            st.divider()
            st.subheader("Ubah atau hapus data")
            id_pilihan = st.selectbox("Pilih ID data", df["id"].tolist(), key="pilih_id")
            row = df[df["id"] == id_pilihan].iloc[0]
            tab_ubah, tab_hapus = st.tabs(["✏️ Ubah", "🗑️ Hapus"])

            with tab_ubah:
                p = f"ubah{id_pilihan}"  # key unik per ID agar nilai awal ikut data yang dipilih
                a, b = st.columns(2)
                with a:
                    e_nama = st.text_input("Nama", value=row["nama"], key=f"{p}_nama")
                with b:
                    e_jk = st.selectbox("Jenis kelamin", ["Laki-laki", "Perempuan"],
                                        index=0 if row["jenis_kelamin"] == "Laki-laki" else 1, key=f"{p}_jk")
                e_bln, e_thn, e_valid = pilih_tanggal_lahir(p, int(row["bulan_lahir"]), int(row["tahun_lahir"]))
                e_gejala = gejala_checkbox(p, [row[f"g{i}"] for i in range(1, 7)])
                if st.button("🔄 Perbarui dan hitung ulang hasil", key=f"{p}_btn"):
                    if not e_nama.strip():
                        st.warning("Nama tidak boleh kosong.")
                    elif e_bln is None or e_thn is None:
                        st.warning("Pilih bulan dan tahun lahir terlebih dahulu.")
                    elif not e_valid:
                        st.warning("Tanggal lahir belum terjadi.")
                    else:
                        ubah_data(id_pilihan, e_nama.strip(), e_jk, e_bln, e_thn, e_gejala)
                        st.success("Data diperbarui.")
                        st.rerun()

            with tab_hapus:
                st.error(f"Hapus data **{row['nama']}** (ID {id_pilihan}) secara permanen?")
                yakin = st.checkbox("Ya, saya yakin", key=f"hapus{id_pilihan}_yakin")
                if st.button("🔥 Hapus data ini", disabled=not yakin, key=f"hapus{id_pilihan}_btn"):
                    hapus_data(id_pilihan)
                    st.success("Data terhapus.")
                    st.rerun()

    # ---- TAB 3: knowledge base ----
    with tab3:
        st.subheader("Indikator")
        st.table(pd.DataFrame([(label, kode) for kode, label in INDIKATOR], columns=["Gejala", "Kode"]).assign(Nilai="1 = Ya, 0 = Tidak"))
        st.subheader("Aturan")
        st.table(pd.DataFrame([(k, format_uraian_rule(u), h) for k, u, _, h in ATURAN], columns=["Kode", "Aturan", "Hasil"]))
        st.write("Risiko Tinggi jika R2 atau R4 terpenuhi. Risiko Sedang jika R1 atau R3 terpenuhi. Selain itu Risiko Rendah.")
        

if __name__ == "__main__":
    main()
