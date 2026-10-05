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

DB_NAME = "skrining_anemia.db"

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
BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
         "Agustus", "September", "Oktober", "November", "Desember"]
REKOMENDASI = {
    "Risiko Tinggi": "Sangat disarankan segera berkonsultasi dengan tenaga kesehatan untuk pemeriksaan darah lengkap dan penanganan lebih lanjut.",
    "Risiko Sedang": "Pantau kondisi, tingkatkan konsumsi makanan sumber zat besi, dan pertimbangkan konsultasi ke fasilitas kesehatan.",
    "Risiko Rendah": "Belum ada aturan yang terpenuhi. Pertahankan pola makan seimbang dan istirahat cukup.",
}


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
                g1 INTEGER DEFAULT 0, g2 INTEGER DEFAULT 0, g3 INTEGER DEFAULT 0,
                g4 INTEGER DEFAULT 0, g5 INTEGER DEFAULT 0, g6 INTEGER DEFAULT 0,
                r1 INTEGER DEFAULT 0, r2 INTEGER DEFAULT 0, r3 INTEGER DEFAULT 0, r4 INTEGER DEFAULT 0,
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
               (nama, jenis_kelamin, bulan_lahir, tahun_lahir, g1,g2,g3,g4,g5,g6,
                r1,r2,r3,r4, tingkat_risiko, rekomendasi, tanggal_dibuat, tanggal_diubah)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (nama, jk, int(bulan), int(tahun), *g, *r, tingkat, rek, t, t),
        )
        conn.commit()
    return r, tingkat, rek


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
    k = [INDIKATOR[i][0] for i in range(6) if row[f"g{i + 1}"]]
    return ", ".join(k) if k else "-"


# ----------------------------------------------------------------------------
# TAMPILAN
# ----------------------------------------------------------------------------
def tampil_hasil(r, tingkat, rek):
    for (kode, uraian, _, hasil), ok in zip(ATURAN, r):
        st.write(("✅ " if ok else "⬜ ") + f"**{kode}** {uraian} → " +
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
        bln_nama = st.selectbox("Bulan lahir", BULAN, index=(bulan_awal or 1) - 1, key=f"{prefix}_bln")
    with c2:
        thn = st.selectbox("Tahun lahir", tahun_list,
                           index=tahun_list.index(tahun_awal) if tahun_awal in tahun_list else 25,
                           key=f"{prefix}_thn")
    bln = BULAN.index(bln_nama) + 1
    u = hitung_usia(bln, thn)
    with c3:
        st.text_input("Usia (otomatis)", value="Tanggal lahir belum terjadi" if u is None else f"{u[0]} tahun {u[1]} bulan",
                      disabled=True, key=f"{prefix}_usia_{bln}_{thn}")
    return bln, thn, u is not None


def gejala_checkbox(prefix, nilai_awal=None):
    nilai_awal = nilai_awal or [0] * 6
    kiri, kanan = st.columns(2)
    hasil = []
    for i, (kode, label) in enumerate(INDIKATOR):
        with (kiri if i < 3 else kanan):
            hasil.append(st.checkbox(f"{kode}  {label}", value=bool(nilai_awal[i]), key=f"{prefix}_g{i}"))
    return hasil


def main():
    st.set_page_config(page_title="CDSS Skrining Anemia", page_icon="🩺", layout="wide")
    init_db()

    st.title("🩺 CDSS Skrining Risiko Anemia")
    st.caption("Sistem pendukung keputusan berbasis aturan (rule-based). Contoh simulasi tugas, bukan alat diagnosis.")

    tab1, tab2, tab3 = st.tabs(["📝 Form Skrining", "📁 Riwayat & Kelola Data", "📚 Knowledge Base"])

    # ---- TAB 1: form skrining (CREATE) ----
    with tab1:
        st.subheader("Data pasien")
        a, b = st.columns(2)
        with a:
            nama = st.text_input("Nama pasien (boleh nama fiktif)", key="baru_nama")
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
            elif not tanggal_valid:
                st.warning("Bulan dan tahun lahir belum terjadi. Periksa kembali.")
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
                    elif not e_valid:
                        st.warning("Bulan dan tahun lahir belum terjadi.")
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
        st.table(pd.DataFrame(INDIKATOR, columns=["Kode", "Keterangan"]).assign(Nilai="1 = Ya, 0 = Tidak"))
        st.subheader("Aturan")
        st.table(pd.DataFrame([(k, u, h) for k, u, _, h in ATURAN], columns=["Kode", "Aturan", "Hasil"]))
        st.write("Risiko Tinggi jika R2 atau R4 terpenuhi. Risiko Sedang jika R1 atau R3 terpenuhi. Selain itu Risiko Rendah.")
        st.caption("Aturan adalah contoh simulasi dan perlu divalidasi dengan literatur sebelum dipakai secara klinis.")


if __name__ == "__main__":
    main()
