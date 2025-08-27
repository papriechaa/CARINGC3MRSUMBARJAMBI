import streamlit as st
import pandas as pd
import plotly.express as px
import google.generativeai as genai

gemini_api_key = st.secrets["GEMINI_API_KEY"]  # ambil dari Streamlit Secrets

genai.configure(api_key=gemini_api_key)


st.set_page_config(page_title="Dashboard Caring", layout="wide")

# =========================
# AREA UTAMA - TITLE & UPLOAD
# =========================
st.title("📊 Visualisasi Data Caring C3MR - SUMBAR & JAMBI")
uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx"])

if uploaded_file:
    # Ambil hanya sheet yang diawali 'SUMBAR JAMBI'
    excel_file = pd.ExcelFile(uploaded_file)
    valid_sheets = [s for s in excel_file.sheet_names if s.upper().startswith("SUMBAR JAMBI")]

    if not valid_sheets:
        st.error("Tidak ada sheet yang diawali 'SUMBAR JAMBI'.")
        st.stop()

    # Sidebar pilih sheet
    st.sidebar.header("Pilih Data")
    sheet_name = st.sidebar.selectbox("Pilih Sheet", valid_sheets)

    # Baca sheet terpilih
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    df.columns = df.columns.str.strip().str.upper()

    # Bersihkan kolom DATEL
    if "DATEL" in df.columns:
        df["DATEL"] = df["DATEL"].astype(str).str.strip().str.upper()
    else:
        st.error("Kolom 'DATEL' tidak ditemukan.")
        st.stop()

    # Bersihkan kolom STATUS PAID
    if "STATUS PAID" in df.columns:
        df["STATUS PAID"] = df["STATUS PAID"].astype(str).str.strip().str.upper()

    # Bersihkan kolom STATUS CARING
    for col in ["STATUS CARING 1", "STATUS CARING 2", "STATUS CARING"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.upper()

    # =========================
    # ======= PENAMBAHAN ======
    #  (data cleaning helper, branch filter, dedupe caring options, dll)
    # =========================

    # Fungsi bantu: bersihkan opsi (hilangkan duplikat, strip, upper)
    def bersihkan_opsi(series):
        if series is None:
            return []
        return sorted(set(series.dropna().astype(str).str.strip().str.upper()))

    # =========================
    # FILTER SIDEBAR
    # =========================
    st.sidebar.subheader("Filter Data")

    # Filter DATEL (tetap ada)
    datel_list = sorted(df["DATEL"].dropna().unique())
    selected_datel = st.sidebar.selectbox("Pilih DATEL", ["(Semua)"] + datel_list)
    if selected_datel != "(Semua)":
        df = df[df["DATEL"] == selected_datel]

    # (Opsional) multiselect DATEL tambahan — tidak mengganti selectbox di atas
    if len(datel_list) > 1:
        selected_datel_multi = st.sidebar.multiselect("Pilih beberapa DATEL (opsional)", datel_list)
        if selected_datel_multi:
            df = df[df["DATEL"].isin(selected_datel_multi)]

    # Filter HABIT
    if "HABIT" in df.columns:
        habit_options = ["Semua"] + sorted(df["HABIT"].dropna().unique().tolist())
        selected_habit = st.sidebar.selectbox("Pilih Habit", habit_options)
        if selected_habit != "Semua":
            df = df[df["HABIT"] == selected_habit]

    # Filter Status Paid
    if "STATUS PAID" in df.columns:
        paid_list = sorted(df["STATUS PAID"].dropna().unique())
        selected_paid = st.sidebar.selectbox("Status Paid", ["(Semua)"] + paid_list)
        if selected_paid != "(Semua)":
            df = df[df["STATUS PAID"] == selected_paid]

    # Pilihan hasil caring
    caring_choice = []
    if "STATUS CARING 1" in df.columns and "STATUS CARING 2" in df.columns:
        caring_choice = ["Semua", "Status Caring 1", "Status Caring 2"]
    elif "STATUS CARING" in df.columns:
        caring_choice = ["Status Caring"]
    else:
        st.error("Tidak ada kolom STATUS CARING yang valid.")
        st.stop()

    selected_hasil_caring = st.sidebar.selectbox("Pilih Hasil Caring", caring_choice)

    # Pilihan jenis status caring (dengan pembersihan supaya tidak kedouble)
    if selected_hasil_caring == "Status Caring 1":
        caring_options = bersihkan_opsi(df["STATUS CARING 1"])
    elif selected_hasil_caring == "Status Caring 2":
        caring_options = bersihkan_opsi(df["STATUS CARING 2"])
    elif selected_hasil_caring == "Status Caring":
        caring_options = bersihkan_opsi(df["STATUS CARING"])
    else:  # Semua
        caring_options = bersihkan_opsi(pd.concat([
            df.get("STATUS CARING 1", pd.Series(dtype=object)),
            df.get("STATUS CARING 2", pd.Series(dtype=object))
        ]))

    selected_jenis_caring = st.sidebar.selectbox(
        "Pilih Jenis Status Caring", ["(Semua)"] + caring_options
    )

    # Filter berdasarkan hasil caring & jenis caring
    if selected_jenis_caring != "(Semua)":
        if selected_hasil_caring == "Status Caring 1":
            df = df[df["STATUS CARING 1"] == selected_jenis_caring]
        elif selected_hasil_caring == "Status Caring 2":
            df = df[df["STATUS CARING 2"] == selected_jenis_caring]
        elif selected_hasil_caring == "Status Caring":
            df = df[df["STATUS CARING"] == selected_jenis_caring]
        else:  # Semua
            df = df[
                (df["STATUS CARING 1"] == selected_jenis_caring)
                | (df["STATUS CARING 2"] == selected_jenis_caring)
            ]

    # =========================
    # ======= PENAMBAHAN ======
    #  Statistik ringkas, download, insight mingguan, rekomendasi follow-up
    # =========================

    # Statistik Ringkas (summary cards)
    st.subheader("📌 Statistik Ringkas")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Data", len(df))
    col2.metric("Jumlah DATEL", df["DATEL"].nunique() if "DATEL" in df.columns else 0)

    if "STATUS PAID" in df.columns:
        try:
            paid_rate = (df["STATUS PAID"].str.upper().eq("PAID").mean() * 100)
            col3.metric("Paid Rate", f"{paid_rate:.1f}%")
        except Exception:
            col3.metric("Paid Rate", "-")
    else:
        col3.metric("Paid Rate", "-")

    # Top status caring
    top_caring_label = "-"
    if "STATUS CARING 1" in df.columns and not df["STATUS CARING 1"].dropna().empty:
        try:
            top_caring_label = df["STATUS CARING 1"].mode()[0]
        except Exception:
            top_caring_label = df["STATUS CARING 1"].dropna().iloc[0]
    col4.metric("Status Caring Terbanyak", top_caring_label)

    # =========================
    # CHART DISTRIBUSI
    # =========================
    st.subheader("Distribusi Status Caring")
    chart_type = st.radio(
        "Pilih jenis chart:",
        ["Pie Chart", "Bar Chart"],
        horizontal=True
    )

    # Fungsi untuk membuat chart (diperbarui: pie chart tampilkan persentase)
    def buat_chart(data, kolom, judul):
        count_df = data[kolom].value_counts().reset_index()
        count_df.columns = [kolom, "JUMLAH"]
        if chart_type == "Pie Chart":
            fig = px.pie(count_df, names=kolom, values="JUMLAH", title=judul)
            fig.update_traces(textinfo="percent+label", textfont_size=8)
        else:
            fig = px.bar(count_df, x=kolom, y="JUMLAH", title=judul, text="JUMLAH")
            fig.update_traces(textposition="outside")
        return fig

    # Tampilkan chart sesuai pilihan (tetap mempertahankan alur aslinya)
    if selected_hasil_caring == "Status Caring 1":
        st.plotly_chart(buat_chart(df, "STATUS CARING 1", "Distribusi Status Caring 1"), use_container_width=True)
    elif selected_hasil_caring == "Status Caring 2":
        st.plotly_chart(buat_chart(df, "STATUS CARING 2", "Distribusi Status Caring 2"), use_container_width=True)
    elif selected_hasil_caring == "Status Caring":
        st.plotly_chart(buat_chart(df, "STATUS CARING", "Distribusi Status Caring"), use_container_width=True)
    else:  # Semua
        df1 = df.copy()
        if "STATUS CARING 1" in df1.columns:
            st.plotly_chart(buat_chart(df1, "STATUS CARING 1", "Distribusi Status Caring 1"), use_container_width=True)
        if "STATUS CARING 2" in df1.columns:
            st.plotly_chart(buat_chart(df1, "STATUS CARING 2", "Distribusi Status Caring 2"), use_container_width=True)

    # =========================
    # RINGKASAN STATUS CARING KOSONG - OTOMATIS SESUAI PILIHAN
    # =========================
    st.subheader("📌 Ringkasan Data Kosong pada Status Caring")

    # Pastikan kolom DATEL bersih
    df["DATEL"] = df["DATEL"].fillna("").astype(str).str.strip().str.upper()
    df = df[df["DATEL"] != ""]

    # Fungsi bantu: ambil jumlah kosong per DATEL dari kolom tertentu
    def jumlah_kosong_per_datel(df, kolom, nama_output):
        if kolom not in df.columns:
            return pd.DataFrame(columns=["DATEL", nama_output])
        temp = df[[kolom, "DATEL"]].copy()
        temp[kolom] = temp[kolom].fillna("").astype(str).str.strip().str.upper()
        temp = temp[temp[kolom].isin(["", "NAN", "NONE", "NULL", "-"])]
        return temp.groupby("DATEL").size().reset_index(name=nama_output)

    # Logic sesuai pilihan caring
    if selected_hasil_caring == "Status Caring 1":
        kosong1 = jumlah_kosong_per_datel(df, "STATUS CARING 1", "JUMLAH KOSONG CARING 1")
        
        if kosong1.empty:
            st.success("✅ Tidak ada data kosong pada kolom STATUS CARING 1.")
        else:
            kosong1 = kosong1.sort_values(by="JUMLAH KOSONG CARING 1", ascending=False).reset_index(drop=True)
            st.dataframe(kosong1, use_container_width=True)

            fig = px.bar(kosong1, x="DATEL", y="JUMLAH KOSONG CARING 1",
                        title="Distribusi Data Kosong Status Caring 1 per DATEL",
                        text="JUMLAH KOSONG CARING 1")
            fig.update_traces(textposition="outside")
            st.plotly_chart(fig, use_container_width=True)

    elif selected_hasil_caring == "Status Caring 2":
        kosong2 = jumlah_kosong_per_datel(df, "STATUS CARING 2", "JUMLAH KOSONG CARING 2")
        
        if kosong2.empty:
            st.success("✅ Tidak ada data kosong pada kolom STATUS CARING 2.")
        else:
            kosong2 = kosong2.sort_values(by="JUMLAH KOSONG CARING 2", ascending=False).reset_index(drop=True)
            st.dataframe(kosong2, use_container_width=True)

            fig = px.bar(kosong2, x="DATEL", y="JUMLAH KOSONG CARING 2",
                        title="Distribusi Data Kosong Status Caring 2 per DATEL",
                        text="JUMLAH KOSONG CARING 2")
            fig.update_traces(textposition="outside")
            st.plotly_chart(fig, use_container_width=True)

    else:
        # Gabungan Status Caring 1 dan 2
        kosong1 = jumlah_kosong_per_datel(df, "STATUS CARING 1", "JUMLAH KOSONG CARING 1")
        kosong2 = jumlah_kosong_per_datel(df, "STATUS CARING 2", "JUMLAH KOSONG CARING 2")

        if kosong1.empty and kosong2.empty:
            st.success("✅ Tidak ada data kosong pada kolom STATUS CARING 1 maupun STATUS CARING 2.")
        else:
            df_kosong = pd.merge(kosong1, kosong2, on="DATEL", how="outer").fillna(0)
            df_kosong["JUMLAH KOSONG CARING 1"] = df_kosong["JUMLAH KOSONG CARING 1"].astype(int)
            df_kosong["JUMLAH KOSONG CARING 2"] = df_kosong["JUMLAH KOSONG CARING 2"].astype(int)

            df_kosong = df_kosong.sort_values(
                by=["JUMLAH KOSONG CARING 1", "JUMLAH KOSONG CARING 2"],
                ascending=False
            ).reset_index(drop=True)

            st.dataframe(df_kosong, use_container_width=True)

            fig = px.bar(
                df_kosong,
                x="DATEL",
                y=["JUMLAH KOSONG CARING 1", "JUMLAH KOSONG CARING 2"],
                title="Distribusi Data Kosong Status Caring 1 & 2 per DATEL",
                barmode="stack"
            )
            fig.update_traces(texttemplate='%{y}', textposition='outside')
            st.plotly_chart(fig, use_container_width=True)




    # # Insight Mingguan per Branch (jika ada kolom tanggal)
    # if "TANGGAL" in df.columns:
    #     st.subheader("📈 Insight Mingguan")
    #     try:
    #         df["TANGGAL"] = pd.to_datetime(df["TANGGAL"], errors='coerce')
    #         weekly = df.groupby([pd.Grouper(key="TANGGAL", freq="W"), "BRANCH"]).size().reset_index(name="JUMLAH")
    #         if not weekly.empty:
    #             fig_weekly = px.line(
    #                 weekly, x="TANGGAL", y="JUMLAH", color="BRANCH",
    #                 title="Jumlah Kasus per Minggu per Branch", markers=True
    #             )
    #             st.plotly_chart(fig_weekly, use_container_width=True)
    #         else:
    #             st.info("Tidak ada data tanggal yang valid untuk insight mingguan.")
    #     except Exception as e:
    #         st.warning(f"Tidak bisa membuat insight mingguan: {e}")

    # Rekomendasi tindak lanjut otomatis untuk kasus collection (rule-based)
    # st.subheader("💡 Rekomendasi Tindak Lanjut (Follow-up Collection)")
    # rekomendasi = []
    # if "STATUS CARING 1" in df.columns:
    #     s = df["STATUS CARING 1"].astype(str).str.upper()
    #     if s.str.contains("NO NUMBER", na=False).any() or s.str.contains("NO_PHONE", na=False).any():
    #         rekomendasi.append("🔹 Ada record 'NO NUMBER' — lakukan cleansing data nomor, cek CRM/DB lain, atau minta tim lapangan validasi kontak.")
    #     if s.str.contains("REJECTED", na=False).any():
    #         rekomendasi.append("🔹 Ada 'REJECTED' — verifikasi alasan rejection; pertimbangkan eskalasi ke team retention atau kunjungan lapangan.")
    #     if s.str.contains("INVALID", na=False).any():
    #         rekomendasi.append("🔹 Ada 'INVALID' — tandai untuk sinkronisasi dengan sumber data utama dan eliminasi duplikat.")
    # if "STATUS CARING 2" in df.columns:
    #     s2 = df["STATUS CARING 2"].astype(str).str.upper()
    #     if s2.str.contains("NO NUMBER", na=False).any() or s2.str.contains("NO_PHONE", na=False).any():
    #         rekomendasi.append("🔹 Di STATUS CARING 2 ditemukan 'NO NUMBER' — koordinasikan dengan team entry data.")
    #     if s2.str.contains("REJECTED", na=False).any():
    #         rekomendasi.append("🔹 Di STATUS CARING 2 ada 'REJECTED' — buat workflow verifikasi tambahan.")

    # if not rekomendasi:
    #     st.markdown("✅ Tidak terdeteksi status NO NUMBER / REJECTED / INVALID pada data hasil filter.")
    # else:
    #     for r in rekomendasi:
    #         st.markdown(r)

    
    # =========================
    # 🤖 AI GEMINI – SOLUSI OTOMATIS
    # =========================
    st.subheader("🤖 Solusi Otomatis dari AI (Gemini)")

    if gemini_api_key:
        genai.configure(api_key=gemini_api_key)

        # Tentukan kolom caring aktif untuk ringkasan
        caring_col = None
        if selected_hasil_caring == "Status Caring 1" and "STATUS CARING 1" in df.columns:
            caring_col = "STATUS CARING 1"
        elif selected_hasil_caring == "Status Caring 2" and "STATUS CARING 2" in df.columns:
            caring_col = "STATUS CARING 2"
        elif selected_hasil_caring == "Status Caring" and "STATUS CARING" in df.columns:
            caring_col = "STATUS CARING"

        # Bangun ringkasan aman (tanpa KeyError)
        try:
            paid_dist = df["STATUS PAID"].value_counts().to_dict() if "STATUS PAID" in df.columns else "-"
        except Exception:
            paid_dist = "-"

        try:
            if caring_col:
                caring_dist = df[caring_col].value_counts().to_dict()
            else:
                stacks = []
                for cc in ["STATUS CARING 1", "STATUS CARING 2", "STATUS CARING"]:
                    if cc in df.columns:
                        stacks.append(df[cc])
                caring_dist = pd.concat(stacks, ignore_index=True).value_counts().to_dict() if stacks else "-"
        except Exception:
            caring_dist = "-"

        try:
            if "STATUS PAID" in df.columns and "DATEL" in df.columns:
                unpaid_rank = (
                    df.assign(_PAID=df["STATUS PAID"].astype(str).str.upper())
                      .query('_PAID != "PAID"')
                      .groupby("DATEL").size()
                      .sort_values(ascending=False)
                      .head(5)
                      .to_dict()
                )
            else:
                unpaid_rank = "-"
        except Exception:
            unpaid_rank = "-"

        # hitung jumlah kosong di kolom caring aktif (jika ada)
        try:
            if caring_col and caring_col in df.columns:
                kosong_count = int((df[caring_col].isna() | (df[caring_col].str.strip() == "")).sum())
            else:
                kosong_count = "-"
        except Exception:
            kosong_count = "-"

        summary = f"""
Total baris: {len(df)}
Jumlah DATEL unik: {df['DATEL'].nunique() if 'DATEL' in df.columns else 0}
Distribusi STATUS PAID: {paid_dist}
Distribusi STATUS CARING (aktif/combined): {caring_dist}
Top 5 DATEL unpaid (estimasi): {unpaid_rank}
Jumlah data kosong di kolom caring aktif: {kosong_count}
Status caring terbanyak (mode caring 1): {top_caring_label}
"""

        # Tombol untuk menjalankan solusi otomatis
        if st.button("🔎 Jalankan Analisis & Solusi Otomatis"):
            prompt = f"""
Kamu adalah AI analis Collection. Berdasarkan ringkasan berikut, buat SOLUSI OTOMATIS yang:
- Spesifik dan actionable,
- Memprioritaskan DATEL yang paling berisiko,
- Memberi quick wins 7 hari,
- Metrik yang dipantau,
- Risiko & mitigasi singkat.

Ringkasan data:
{summary}

Format jawaban:
1) Prioritas Tindakan (bullet)
2) Quick Wins 7 Hari (bullet)
3) Metrik yang Dipantau (bullet)
4) Risiko & Mitigasi (bullet)
5) Catatan Tambahan (jika perlu)
"""

            try:
                model = genai.GenerativeModel("gemini-1.5-flash")
                resp = model.generate_content(prompt)
                ai_text = resp.text if hasattr(resp, "text") else str(resp)
                st.markdown("### 💡 Hasil Analisis & Solusi AI")
                st.write(ai_text)
            except Exception as e:
                st.error(f"Gagal memanggil Gemini: {e}")
    else:
        st.info("Masukkan Gemini API Key di sidebar untuk mengaktifkan solusi otomatis AI.")

else:
    st.info("Silakan upload file Excel terlebih dahulu.")