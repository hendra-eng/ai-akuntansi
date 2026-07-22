import io

import pandas as pd
import streamlit as st

import akuntansi_ai as ak

st.set_page_config(page_title="AI Kategorisasi Rekening Koran", layout="wide")

POLA_FILE = "pola_belajar.json"

if "pola" not in st.session_state:
    st.session_state.pola = ak.muat_pola(POLA_FILE)
if "hasil" not in st.session_state:
    st.session_state.hasil = None
if "df_coa" not in st.session_state:
    st.session_state.df_coa = pd.DataFrame()

st.title("🤖 AI Kategorisasi & Jurnal Rekening Koran")
st.write(
    "Upload file Excel rekening koran (boleh multi-sheet per bank, format kolom "
    "bebas berbeda tiap bank). Baris yang jurnalnya **sudah diisi** (NO AKUN / NAMA AKUN "
    "debet & kredit) dipakai untuk **mempelajari pola** secara otomatis. Baris yang "
    "**belum ada jurnalnya** akan diisi otomatis: pola historis → kata kunci COA → AI."
)

if not ak.ambil_api_key():
    st.info(
        "ℹ️ Fallback AI belum aktif karena `ANTHROPIC_API_KEY` belum di-set "
        "(environment variable atau `.streamlit/secrets.toml`). Pola historis & "
        "kata kunci COA tetap berjalan normal."
    )

with st.expander(f"📚 Pola yang sudah dipelajari sejauh ini ({len(st.session_state.pola.aturan)} pola)"):
    if st.session_state.pola.aturan:
        baris_pola = []
        for (sig, arah), v in st.session_state.pola.aturan.items():
            baris_pola.append({
                "Signature": sig, "Arah": arah,
                "Akun Debet": f"{v['no_akun_debet']} - {v['nama_akun_debet']}",
                "Akun Kredit": f"{v['no_akun_kredit']} - {v['nama_akun_kredit']}",
                "Status": "Konsisten" if v.get("konsisten") else "Mayoritas (cek manual)",
                "Jumlah contoh": v.get("jumlah_contoh"),
            })
        st.dataframe(pd.DataFrame(baris_pola), use_container_width=True)
        if st.button("🗑️ Reset semua pola yang sudah dipelajari"):
            st.session_state.pola = ak.Pola()
            ak.simpan_pola(st.session_state.pola, POLA_FILE)
            st.rerun()
    else:
        st.write("Belum ada pola tersimpan. Upload file yang jurnalnya sudah lengkap untuk mulai belajar.")

st.subheader("1. Upload file Excel rekening koran")
file = st.file_uploader("File .xlsx (multi-sheet per bank, sheet 'COA' opsional)", type=["xlsx"])

pakai_ai = st.checkbox("Gunakan AI (Claude) untuk transaksi yang tidak cocok pola/kata kunci", value=True)

if file is not None:
    if st.button("🚀 Proses File", type="primary"):
        with st.spinner("Membaca & mendeteksi format sheet..."):
            try:
                df, df_coa, peringatan = ak.muat_workbook_rekening_koran(file)
            except Exception as e:
                st.error(f"Gagal membaca file: {e}")
                df = pd.DataFrame()
                df_coa = pd.DataFrame()
                peringatan = []

        for p in peringatan:
            st.warning(f"⚠️ {p}")

        if df.empty:
            st.error("Tidak ada sheet rekening koran yang berhasil dibaca dari file ini.")
        else:
            st.success(f"✅ Berhasil membaca {len(df)} baris transaksi dari sheet: {', '.join(df['bank'].unique())}")
            if not df_coa.empty:
                st.success(f"✅ Sheet COA ditemukan — {len(df_coa)} akun perusahaan dimuat.")
                st.session_state.df_coa = df_coa
            else:
                st.info("ℹ️ Tidak ada sheet COA di file ini — kategorisasi kata kunci akan pakai penanda generik saja.")

            with st.spinner("Mempelajari pola dari baris yang jurnalnya sudah ada..."):
                pola_baru = ak.pelajari_pola(df)
                st.session_state.pola = ak.gabung_pola(st.session_state.pola, pola_baru)
                ak.simpan_pola(st.session_state.pola, POLA_FILE)
            st.success(f"✅ Mempelajari/memperbarui {len(pola_baru.aturan)} pola dari file ini.")

            with st.spinner("Menerapkan pola + kata kunci" + (" + AI" if pakai_ai else "") + " ke baris yang belum berjurnal..."):
                hasil = ak.proses_dataframe(
                    df, df_coa if not df_coa.empty else st.session_state.df_coa,
                    st.session_state.pola, pakai_ai=pakai_ai,
                )
            st.session_state.hasil = hasil

if st.session_state.hasil is not None:
    hasil = st.session_state.hasil

    st.subheader("2. Hasil Kategorisasi / Jurnal")

    daftar_bank = ["(Semua)"] + sorted(hasil["bank"].unique().tolist())
    pilihan_bank = st.selectbox("Filter bank:", daftar_bank)
    tampil = hasil if pilihan_bank == "(Semua)" else hasil[hasil["bank"] == pilihan_bank]

    kolom_tampil = [
        "bank", "tanggal", "keterangan", "supplier_cust",
        "mutasi_debet", "mutasi_kredit",
        "no_akun_debet", "nama_akun_debet", "jml_debet",
        "no_akun_kredit", "nama_akun_kredit", "jml_kredit",
        "sumber_kategori",
    ]
    st.dataframe(tampil[kolom_tampil], use_container_width=True, height=420)

    st.subheader("3. Ringkasan Sumber Kategorisasi")
    ringkasan_sumber = tampil["sumber_kategori"].value_counts().reset_index()
    ringkasan_sumber.columns = ["Sumber", "Jumlah Baris"]
    st.dataframe(ringkasan_sumber, use_container_width=True)

    perlu_review = tampil[
        tampil["sumber_kategori"].isin([
            "Belum Terkategori - perlu review manual",
            "Pola historis (mayoritas - perlu cek)",
        ])
    ]
    if len(perlu_review) > 0:
        st.warning(f"⚠️ {len(perlu_review)} baris sebaiknya direview manual sebelum diposting ke jurnal.")
        st.dataframe(perlu_review[kolom_tampil], use_container_width=True)

    st.subheader("4. Ringkasan per Akun")
    ring_debet = tampil.groupby(["no_akun_debet", "nama_akun_debet"])["jml_debet"].sum().reset_index()
    ring_debet.columns = ["No Akun", "Nama Akun", "Total Debet"]
    ring_kredit = tampil.groupby(["no_akun_kredit", "nama_akun_kredit"])["jml_kredit"].sum().reset_index()
    ring_kredit.columns = ["No Akun", "Nama Akun", "Total Kredit"]
    col1, col2 = st.columns(2)
    with col1:
        st.write("**Total Debet per Akun**")
        st.dataframe(ring_debet.sort_values("Total Debet", ascending=False), use_container_width=True)
    with col2:
        st.write("**Total Kredit per Akun**")
        st.dataframe(ring_kredit.sort_values("Total Kredit", ascending=False), use_container_width=True)

    st.subheader("5. Download Hasil")
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        hasil[kolom_tampil].to_excel(writer, sheet_name="Jurnal", index=False)
        ringkasan_sumber.to_excel(writer, sheet_name="Ringkasan Sumber", index=False)
    st.download_button(
        "📥 Download hasil (Excel)",
        data=buffer.getvalue(),
        file_name="hasil_jurnal_rekening_koran.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Silakan upload file Excel rekening koran untuk mulai.")