# CDSS Skrining Risiko Anemia (Rule-Based)

Program web Streamlit dengan database SQLite. Aturan R1-R4 sama dengan file Excel dan PowerPoint tugas.
Contoh simulasi tugas, bukan alat diagnosis.

## Menjalankan di komputer sendiri (VS Code)
1. Buka folder ini di VS Code, lalu buka Terminal (Ctrl + `).
2. Pasang kebutuhan:  `pip install -r requirements.txt`
3. Jalankan:          `streamlit run app.py`
4. Browser terbuka otomatis di http://localhost:8501

File database `skrining_anemia.db` dibuat otomatis di folder yang sama saat pertama dijalankan.
(Kalau `streamlit` tidak dikenali: `python -m streamlit run app.py`.)

## Fitur
- Form skrining: nama, jenis kelamin, bulan dan tahun lahir (usia otomatis), kotak centang gejala G1-G6, hasil langsung.
- Simpan hasil ke database (Create).
- Riwayat, unduh CSV (Read), ubah dengan hitung ulang hasil (Update), hapus dengan konfirmasi (Delete).
- Tab Knowledge Base berisi indikator dan aturan.
- Tombol "Isi 5 data contoh" muncul saat database masih kosong.

## Membuat link yang bisa dibuka dosen (Streamlit Community Cloud)
1. Buat repositori GitHub (publik) berisi `app.py`, `requirements.txt`, dan folder `.streamlit`.
2. Buka https://share.streamlit.io, masuk dengan GitHub, klik "Create app", pilih repositori dan file `app.py`.
3. Setelah selesai dibuat, salin link aplikasi untuk dikumpulkan.

Catatan: di Streamlit Community Cloud file SQLite tidak permanen. Data bisa hilang saat aplikasi restart atau
di-deploy ulang. Untuk tugas ini biasanya cukup, dan tombol "Isi 5 data contoh" membantu mengisi ulang.
Gunakan nama fiktif atau inisial, jangan data kesehatan orang sungguhan.
