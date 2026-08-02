# -*- coding: utf-8 -*-
"""
Dashboard Interaktif: Pengelolaan Sampah & Limbah B3 di Indonesia
Sumber data: TABEL_2_3_3, TABEL_2_3_4, TABEL_2_3_5, TABEL_2_3_6

Cara menjalankan:
    pip install -r requirements.txt
    streamlit run app.py

Struktur folder yang diharapkan:
    app.py
    requirements.txt
    data/
        TABEL_2_3_3.xlsx
        TABEL_2_3_4.xlsx
        TABEL_2_3_5.xlsx
        TABEL_2_3_6.xlsx

Jika file data tidak ditemukan di folder ./data, pengguna bisa mengunggahnya
langsung lewat sidebar aplikasi.
"""

import io
import os
import base64
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

from embedded_data import EMBEDDED_FILES  # data Excel sudah tertanam otomatis di sini

# --------------------------------------------------------------------------------------
# KONFIGURASI HALAMAN
# --------------------------------------------------------------------------------------
st.set_page_config(
    page_title="Dashboard Sampah & Limbah B3 Indonesia",
    page_icon="🗑️",
    layout="wide",
    initial_sidebar_state="expanded",
)

FILES = {
    "2.3.3": "TABEL_2_3_3.xlsx",
    "2.3.4": "TABEL_2_3_4.xlsx",
    "2.3.5": "TABEL_2_3_5.xlsx",
    "2.3.6": "TABEL_2_3_6.xlsx",
}

PALETTE = px.colors.qualitative.Set2
TEMPLATE = "plotly_white"


# --------------------------------------------------------------------------------------
# CSS SEDIKIT MEMPERCANTIK TAMPILAN
# --------------------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .main > div {padding-top: 1.2rem;}
    div[data-testid="stMetric"] {
        background-color: #f7f9f8;
        border: 1px solid #e3e8e6;
        border-radius: 12px;
        padding: 14px 16px 8px 16px;
    }
    div[data-testid="stMetric"] label,
    div[data-testid="stMetric"] div[data-testid="stMetricValue"],
    div[data-testid="stMetric"] div[data-testid="stMetricLabel"] {
        color: #1f3d33 !important;
    }
    h1, h2, h3 { color: #1f3d33; }
    .stTabs [data-baseweb="tab-list"] { gap: 6px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #f0f3f1;
        border-radius: 8px 8px 0 0;
        padding: 8px 16px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# --------------------------------------------------------------------------------------
# FUNGSI BANTUAN: MEMUAT & MEMBERSIHKAN DATA
# --------------------------------------------------------------------------------------
def _read_excel(source):
    return pd.read_excel(source, header=None)


@st.cache_data(show_spinner=False)
def load_table_233(source):
    """Pengelolaan Limbah B3 per sektor (2020-2024, rata-rata)."""
    df = _read_excel(source)
    df = df.iloc[2:20].reset_index(drop=True)
    df.columns = ["Sektor", "Indikator", "Nilai"]
    df["Sektor"] = df["Sektor"].ffill()
    df["Sektor"] = df["Sektor"].str.split("/").str[0].str.strip()
    sektor_map = {
        "Fasilitas Pelayanan Kesehatan Health Service Facilities": "Fasilitas Pelayanan Kesehatan",
        "Pertambangan, Energi, dan Migas Mining, Energy, Oil, and Gas": "Pertambangan, Energi, dan Migas",
    }
    df["Sektor"] = df["Sektor"].replace(sektor_map)

    indikator_map = {
        "Jumlah Perusahaan (industri) Number of Companies (industry)": "Jumlah Perusahaan",
        "Limbah yang dihasilkan (ton) Produced Waste (tonnes)": "Limbah Dihasilkan (ton)",
        "Limbah yang dikelola (ton) Managed Waste (tonnes)": "Limbah Dikelola (ton)",
    }
    df["Indikator"] = df["Indikator"].map(indikator_map).fillna(df["Indikator"])
    df["Nilai"] = pd.to_numeric(df["Nilai"], errors="coerce")

    wide = df.pivot_table(index="Sektor", columns="Indikator", values="Nilai", aggfunc="first").reset_index()
    cols = ["Sektor", "Jumlah Perusahaan", "Limbah Dihasilkan (ton)", "Limbah Dikelola (ton)"]
    wide = wide[[c for c in cols if c in wide.columns]]
    wide["% Limbah Dikelola"] = (wide["Limbah Dikelola (ton)"] / wide["Limbah Dihasilkan (ton)"] * 100).round(2)
    return wide


@st.cache_data(show_spinner=False)
def load_table_234(source):
    """Timbulan & pengelolaan sampah tahunan per provinsi."""
    df = _read_excel(source)
    df = df.iloc[2:40].reset_index(drop=True)
    df.columns = [
        "Provinsi",
        "Timbulan Sampah Tahunan (ton/tahun)",
        "Pengurangan Sampah (%)",
        "Penanganan Sampah (%)",
        "Sampah Terkelola (%)",
    ]
    for c in df.columns[1:]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


@st.cache_data(show_spinner=False)
def load_table_235(source):
    """Komposisi sampah per provinsi (%)."""
    df = _read_excel(source)
    df = df.iloc[2:40].reset_index(drop=True)
    df.columns = [
        "Provinsi", "Sisa Makanan", "Kayu-Ranting", "Kertas-Karton", "Plastik",
        "Logam", "Kain", "Karet-Kulit", "Kaca", "Lainnya",
    ]
    for c in df.columns[1:]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


@st.cache_data(show_spinner=False)
def load_table_236(source):
    """Sumber timbulan sampah per provinsi."""
    df = _read_excel(source)
    df = df.iloc[2:40].reset_index(drop=True)
    df.columns = [
        "Provinsi", "Rumah Tangga", "Perkantoran", "Pasar", "Perniagaan",
        "Fasilitas Publik", "Kawasan", "Lainnya",
    ]
    for c in df.columns[1:]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def get_source(key):
    """Ambil sumber file: upload pengguna jika ada, kalau tidak pakai data yang sudah
    tertanam otomatis (embedded_data.py) — tidak perlu file eksternal / folder data/."""
    uploaded = st.session_state.get(f"upload_{key}")
    if uploaded is not None:
        return uploaded
    b64 = EMBEDDED_FILES.get(key)
    if b64:
        return io.BytesIO(base64.b64decode(b64))
    return None


def df_download_button(df, label, filename, key):
    csv = df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label=f"⬇️ {label}",
        data=csv,
        file_name=filename,
        mime="text/csv",
        key=key,
        use_container_width=True,
    )


def fig_download_button(fig, filename, key, label="⬇️ Unduh Grafik (PNG)"):
    """Tombol unduh grafik sebagai PNG (butuh paket kaleido)."""
    try:
        img_bytes = fig.to_image(format="png", scale=2)
        st.download_button(
            label=label,
            data=img_bytes,
            file_name=filename,
            mime="image/png",
            key=key,
            use_container_width=True,
        )
    except Exception:
        html_bytes = fig.to_html(include_plotlyjs="cdn").encode("utf-8")
        st.download_button(
            label="⬇️ Unduh Grafik (HTML interaktif)",
            data=html_bytes,
            file_name=filename.replace(".png", ".html"),
            mime="text/html",
            key=key + "_html",
            use_container_width=True,
        )
    st.caption("Tips: gunakan ikon kamera 📷 di pojok kanan atas grafik untuk mengunduh langsung juga.")


def render_metric_card_png(df, specs, dpi=200):
    """Gambar kartu KPI datar (bukan speedometer) untuk beberapa metrik rata-rata nasional
    sebagai satu PNG utuh memakai matplotlib. Setiap kartu berisi: label, nilai besar, dan
    mini progress bar yang menunjukkan posisi rata-rata terhadap rentang min-maks provinsi.
    Tidak bergantung pada kaleido/Chrome sehingga tombol unduh PNG selalu berfungsi."""
    n = len(specs)
    fig, axes = plt.subplots(1, n, figsize=(3.7 * n, 2.1))
    fig.patch.set_facecolor("white")
    if n == 1:
        axes = [axes]

    for ax, spec in zip(axes, specs):
        col_data = df[spec["col"]]
        val = float(col_data.mean())
        vmin, vmax = spec.get("fixed_range", (float(col_data.min()), float(col_data.max())))
        if vmax == vmin:
            vmax = vmin + 1
        frac = max(0.0, min(1.0, (val - vmin) / (vmax - vmin)))

        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        ax.axis("off")

        card = mpatches.FancyBboxPatch(
            (0.025, 0.06), 0.95, 0.88,
            boxstyle="round,pad=0,rounding_size=0.09",
            linewidth=1.3, edgecolor="#e3e8e6", facecolor="#f7f9f8",
            transform=ax.transAxes, zorder=1,
        )
        ax.add_patch(card)

        ax.add_patch(mpatches.Circle((0.10, 0.72), 0.028, facecolor=spec["bar_color"],
                                      edgecolor="none", transform=ax.transAxes, zorder=2))
        ax.text(0.155, 0.72, spec["label"], fontsize=11.5, fontweight="bold",
                color=spec["bar_color"], va="center", ha="left", transform=ax.transAxes, zorder=2)

        ax.text(0.09, 0.46, f"{format(val, spec['valueformat'])}{spec['suffix']}",
                fontsize=25, fontweight="bold", color="#1f3d33",
                va="center", ha="left", transform=ax.transAxes, zorder=2)

        # mini progress bar (bukan speedometer) menunjukkan posisi rata-rata nasional
        bar_x0, bar_w, bar_y, bar_h = 0.09, 0.82, 0.20, 0.05
        ax.add_patch(mpatches.FancyBboxPatch((bar_x0, bar_y), bar_w, bar_h,
                     boxstyle="round,pad=0,rounding_size=0.02", linewidth=0,
                     facecolor="#e3e8e6", transform=ax.transAxes, zorder=2))
        ax.add_patch(mpatches.FancyBboxPatch((bar_x0, bar_y), bar_w * max(frac, 0.02), bar_h,
                     boxstyle="round,pad=0,rounding_size=0.02", linewidth=0,
                     facecolor=spec["bar_color"], transform=ax.transAxes, zorder=3))

        ax.text(bar_x0, 0.11, format(vmin, spec["valueformat"]), fontsize=8.5, color="#8a9a94",
                va="center", ha="left", transform=ax.transAxes, zorder=2)
        ax.text(bar_x0 + bar_w, 0.11, format(vmax, spec["valueformat"]), fontsize=8.5, color="#8a9a94",
                va="center", ha="right", transform=ax.transAxes, zorder=2)

    fig.subplots_adjust(left=0.005, right=0.995, top=0.98, bottom=0.02, wspace=0.06)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


NASIONAL_LABEL = "🇮🇩 Nasional (Rata-rata Semua Provinsi)"


def national_row(df, kategori_cols, label=NASIONAL_LABEL):
    """Hitung rata-rata nasional (rata-rata dari semua provinsi) untuk kolom-kolom kategori
    tertentu, dan kembalikan sebagai satu baris DataFrame agar bisa digabung (concat)
    dengan data provinsi terpilih."""
    means = df[kategori_cols].mean(numeric_only=True)
    row = {"Provinsi": label}
    row.update(means.to_dict())
    return pd.DataFrame([row])


# --------------------------------------------------------------------------------------
# SIDEBAR
# --------------------------------------------------------------------------------------
st.sidebar.title("🗑️ Panel Kontrol")
st.sidebar.caption("Dashboard Pengelolaan Sampah & Limbah B3 Indonesia")

with st.sidebar.expander("📤 Ganti data (opsional)", expanded=False):
    st.caption("Data 4 tabel sudah otomatis termuat dari dalam aplikasi. Unggah di sini hanya jika ingin memakai versi data yang lebih baru.")
    for key, fname in FILES.items():
        up = st.file_uploader(f"{fname}", type=["xlsx"], key=f"uploader_{key}")
        if up is not None:
            st.session_state[f"upload_{key}"] = up

st.sidebar.divider()
page = st.sidebar.radio(
    "Pilih Halaman",
    [
        "🏠 Beranda",
        "☣️ Limbah B3 per Sektor",
        "📦 Timbulan Sampah Provinsi",
        "🧩 Komposisi Sampah",
        "🏷️ Sumber Timbulan Sampah",
        "⚖️ Perbandingan Antar Provinsi",
    ],
)
st.sidebar.divider()
st.sidebar.caption(f"Terakhir dibuka: {datetime.now().strftime('%d %b %Y, %H:%M')}")

# --------------------------------------------------------------------------------------
# MEMUAT SEMUA DATA (dengan penanganan file hilang)
# --------------------------------------------------------------------------------------
sources = {k: get_source(k) for k in FILES}

df_233 = load_table_233(sources["2.3.3"])
df_234 = load_table_234(sources["2.3.4"])
df_235 = load_table_235(sources["2.3.5"])
df_236 = load_table_236(sources["2.3.6"])


# ========================================================================================
# HALAMAN: BERANDA
# ========================================================================================
if page == "🏠 Beranda":
    st.title("🗑️ Dashboard Pengelolaan Sampah & Limbah B3 di Indonesia")
    st.markdown(
        """
        Dashboard ini menyajikan visualisasi interaktif dari empat tabel data pengelolaan
        sampah dan limbah Bahan Berbahaya dan Beracun (B3) di Indonesia. Gunakan menu di
        **sidebar kiri** untuk berpindah antar halaman analisis, dan setiap grafik dapat
        **diunduh langsung** (PNG maupun data mentah dalam CSV).
        """
    )

    c1, c2, c3, c4 = st.columns(4)
    total_timbulan = df_234["Timbulan Sampah Tahunan (ton/tahun)"].sum()
    rata_kelola = df_234["Sampah Terkelola (%)"].mean()
    total_b3_dihasilkan = df_233["Limbah Dihasilkan (ton)"].sum()
    total_b3_dikelola = df_233["Limbah Dikelola (ton)"].sum()

    c1.metric("Total Timbulan Sampah / Tahun", f"{total_timbulan:,.0f} ton")
    c2.metric("Rata-rata Sampah Terkelola", f"{rata_kelola:,.1f} %")
    c3.metric("Total Limbah B3 Dihasilkan", f"{total_b3_dihasilkan:,.0f} ton")
    c4.metric(
        "Rasio Limbah B3 Dikelola",
        f"{(total_b3_dikelola/total_b3_dihasilkan*100):,.1f} %",
    )

    st.divider()
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Provinsi dengan Timbulan Sampah Tertinggi")
        top10 = df_234.nlargest(10, "Timbulan Sampah Tahunan (ton/tahun)")
        fig = px.bar(
            top10.sort_values("Timbulan Sampah Tahunan (ton/tahun)"),
            x="Timbulan Sampah Tahunan (ton/tahun)", y="Provinsi",
            orientation="h", template=TEMPLATE,
            color="Timbulan Sampah Tahunan (ton/tahun)",
            color_continuous_scale="Teal",
        )
        fig.update_layout(coloraxis_showscale=False, height=420, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.subheader("Komposisi Limbah B3 Nasional per Sektor")
        fig2 = px.pie(
            df_233, names="Sektor", values="Limbah Dihasilkan (ton)",
            template=TEMPLATE, color_discrete_sequence=PALETTE, hole=0.45,
        )
        fig2.update_layout(height=420, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig2, use_container_width=True)

    st.info(
        "📌 Gunakan menu sidebar untuk eksplorasi lebih dalam: analisis per sektor limbah B3, "
        "timbulan sampah per provinsi, komposisi jenis sampah, sumber timbulan, hingga "
        "perbandingan langsung antar provinsi."
    )


# ========================================================================================
# HALAMAN: LIMBAH B3 PER SEKTOR
# ========================================================================================
elif page == "☣️ Limbah B3 per Sektor":
    st.title("☣️ Pengelolaan Limbah B3 per Sektor (Rata-rata 2020–2024)")

    sektor_pilihan = st.multiselect(
        "Pilih sektor yang ditampilkan", options=df_233["Sektor"].tolist(),
        default=df_233["Sektor"].tolist(),
    )
    dff = df_233[df_233["Sektor"].isin(sektor_pilihan)]

    if dff.empty:
        st.warning("Pilih minimal satu sektor.")
        st.stop()

    tab1, tab2, tab3 = st.tabs(["📊 Limbah Dihasilkan vs Dikelola", "🏭 Jumlah Perusahaan", "✅ Efisiensi Pengelolaan"])

    with tab1:
        fig = go.Figure()
        fig.add_bar(name="Limbah Dihasilkan (ton)", x=dff["Sektor"], y=dff["Limbah Dihasilkan (ton)"], marker_color=PALETTE[1])
        fig.add_bar(name="Limbah Dikelola (ton)", x=dff["Sektor"], y=dff["Limbah Dikelola (ton)"], marker_color=PALETTE[2])
        fig.update_layout(
            barmode="group", template=TEMPLATE, height=480,
            yaxis_title="Ton", legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig, use_container_width=True)
        colA, colB = st.columns(2)
        with colA:
            fig_download_button(fig, "limbah_b3_dihasilkan_vs_dikelola.png", "dl_233_fig1")
        with colB:
            df_download_button(dff, "Unduh data (CSV)", "limbah_b3_sektor.csv", "dl_233_csv1")

    with tab2:
        fig2 = px.bar(
            dff.sort_values("Jumlah Perusahaan"), x="Jumlah Perusahaan", y="Sektor",
            orientation="h", template=TEMPLATE, color="Sektor", color_discrete_sequence=PALETTE,
            text="Jumlah Perusahaan",
        )
        fig2.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
        fig2.update_layout(showlegend=False, height=440)
        st.plotly_chart(fig2, use_container_width=True)
        fig_download_button(fig2, "jumlah_perusahaan_per_sektor.png", "dl_233_fig2")

    with tab3:
        fig3 = px.bar(
            dff.sort_values("% Limbah Dikelola"), x="% Limbah Dikelola", y="Sektor",
            orientation="h", template=TEMPLATE, color="% Limbah Dikelola",
            color_continuous_scale="RdYlGn", range_color=[0, 100], text="% Limbah Dikelola",
        )
        fig3.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
        fig3.update_layout(height=440, coloraxis_showscale=False)
        st.plotly_chart(fig3, use_container_width=True)
        st.caption("Efisiensi = (Limbah Dikelola / Limbah Dihasilkan) × 100%")
        fig_download_button(fig3, "efisiensi_pengelolaan_b3.png", "dl_233_fig3")

    st.divider()
    st.dataframe(dff, use_container_width=True, hide_index=True)


# ========================================================================================
# HALAMAN: TIMBULAN SAMPAH PROVINSI
# ========================================================================================
elif page == "📦 Timbulan Sampah Provinsi":
    st.title("📦 Timbulan & Pengelolaan Sampah per Provinsi")

    # -------------------- SECTION: ANGKA NASIONAL (KARTU KPI) --------------------
    st.subheader("🇮🇩 Angka Nasional (Rata-rata Semua Provinsi)")
    st.caption(
        "Posisi rata-rata nasional dibandingkan rentang (nilai terendah–tertinggi) seluruh 38 provinsi. "
        "Bar tipis di bawah nilai menunjukkan posisi rata-rata; angka di ujungnya adalah skala minimum dan maksimum provinsi."
    )

    metric_specs = [
        {
            "label": "Timbulan Sampah Tahunan",
            "col": "Timbulan Sampah Tahunan (ton/tahun)",
            "suffix": "", "valueformat": ",.0f",
            "bar_color": "#2E86AB",
        },
        {
            "label": "Pengurangan Sampah",
            "col": "Pengurangan Sampah (%)",
            "suffix": "%", "valueformat": ".1f",
            "bar_color": "#8E44AD",
            "fixed_range": (0, 100),
        },
        {
            "label": "Penanganan Sampah",
            "col": "Penanganan Sampah (%)",
            "suffix": "%", "valueformat": ".1f",
            "bar_color": "#16A085",
            "fixed_range": (0, 100),
        },
        {
            "label": "Sampah Terkelola",
            "col": "Sampah Terkelola (%)",
            "suffix": "%", "valueformat": ".1f",
            "bar_color": "#27AE60",
            "fixed_range": (0, 100),
        },
    ]

    metric_card_png = render_metric_card_png(df_234, metric_specs)
    st.image(metric_card_png, use_container_width=True)
    st.download_button(
        label="⬇️ Unduh Kartu Angka Nasional (PNG)",
        data=metric_card_png,
        file_name="kartu_angka_nasional_234.png",
        mime="image/png",
        key="dl_234_gauge_card",
        use_container_width=True,
    )
    st.caption("Dihitung dari rata-rata seluruh 38 provinsi (bukan dari total nasional).")
    st.divider()

    colf1, colf2 = st.columns([2, 1])
    with colf1:
        provinsi_pilihan = st.multiselect(
            "Filter provinsi (kosongkan = semua)", options=df_234["Provinsi"].tolist(), default=[]
        )
    with colf2:
        n_top = st.slider("Tampilkan Top-N (jika tidak memfilter provinsi)", 5, 38, 15)

    dff = df_234[df_234["Provinsi"].isin(provinsi_pilihan)] if provinsi_pilihan else df_234.copy()

    sort_by = st.selectbox(
        "Urutkan berdasarkan",
        ["Timbulan Sampah Tahunan (ton/tahun)", "Pengurangan Sampah (%)", "Penanganan Sampah (%)", "Sampah Terkelola (%)"],
    )
    dff_sorted = dff.sort_values(sort_by, ascending=False)
    if not provinsi_pilihan:
        dff_sorted = dff_sorted.head(n_top)

    fig = px.bar(
        dff_sorted.sort_values(sort_by), x=sort_by, y="Provinsi", orientation="h",
        template=TEMPLATE, color=sort_by, color_continuous_scale="Teal",
        text=sort_by,
    )
    fig.update_traces(texttemplate="%{text:,.1f}", textposition="outside")
    fig.update_layout(height=max(420, 24 * len(dff_sorted)), coloraxis_showscale=False)
    nasional_mean = df_234[sort_by].mean()
    fig.add_vline(
        x=nasional_mean, line_dash="dash", line_color="#d62728",
        annotation_text=f"Rata-rata Nasional: {nasional_mean:,.1f}", annotation_position="top",
    )
    st.plotly_chart(fig, use_container_width=True)
    colA, colB = st.columns(2)
    with colA:
        fig_download_button(fig, "timbulan_sampah_provinsi.png", "dl_234_fig1")
    with colB:
        df_download_button(dff_sorted, "Unduh data (CSV)", "timbulan_sampah_provinsi.csv", "dl_234_csv1")

    st.divider()
    st.subheader("Hubungan Pengurangan vs Penanganan Sampah")
    fig2 = px.scatter(
        dff, x="Pengurangan Sampah (%)", y="Penanganan Sampah (%)",
        size="Timbulan Sampah Tahunan (ton/tahun)", color="Sampah Terkelola (%)",
        hover_name="Provinsi", template=TEMPLATE, color_continuous_scale="RdYlGn",
        size_max=45,
    )
    fig2.update_layout(height=520)
    st.plotly_chart(fig2, use_container_width=True)
    st.caption("Ukuran gelembung = total timbulan sampah tahunan. Warna = persentase sampah terkelola.")
    fig_download_button(fig2, "scatter_pengurangan_penanganan.png", "dl_234_fig2")

    st.divider()
    st.dataframe(dff.sort_values(sort_by, ascending=False), use_container_width=True, hide_index=True)


# ========================================================================================
# HALAMAN: KOMPOSISI SAMPAH
# ========================================================================================
elif page == "🧩 Komposisi Sampah":
    st.title("🧩 Komposisi Jenis Sampah per Provinsi")
    kategori = [c for c in df_235.columns if c != "Provinsi"]

    st.subheader("🇮🇩 Angka Nasional per Jenis Sampah (Rata-rata Semua Provinsi)")
    nasional_235 = df_235[kategori].mean(numeric_only=True).round(2)
    nasional_235_sorted = nasional_235.sort_values()

    colN1, colN2 = st.columns([1, 1.2])
    with colN1:
        fig_nas_pie235 = px.pie(
            names=nasional_235.index, values=nasional_235.values, hole=0.45,
            template=TEMPLATE, color_discrete_sequence=PALETTE,
            title="Proporsi Rata-rata Nasional",
        )
        fig_nas_pie235.update_traces(textinfo="percent+label", textposition="outside")
        fig_nas_pie235.update_layout(height=420, showlegend=False, margin=dict(t=60, b=10, l=10, r=10))
        st.plotly_chart(fig_nas_pie235, use_container_width=True)
    with colN2:
        fig_nas_bar235 = px.bar(
            x=nasional_235_sorted.values, y=nasional_235_sorted.index, orientation="h",
            template=TEMPLATE, color=nasional_235_sorted.values, color_continuous_scale="Greens",
            text=nasional_235_sorted.values, labels={"x": "Rata-rata Nasional (%)", "y": ""},
            title="Peringkat Rata-rata Nasional",
        )
        fig_nas_bar235.update_traces(texttemplate="%{text:.2f}%", textposition="outside")
        fig_nas_bar235.update_layout(height=420, coloraxis_showscale=False, margin=dict(t=60, b=10, l=10, r=10))
        st.plotly_chart(fig_nas_bar235, use_container_width=True)

    colD1, colD2, colD3 = st.columns(3)
    with colD1:
        fig_download_button(fig_nas_pie235, "nasional_komposisi_pie.png", "dl_235_nas_pie")
    with colD2:
        fig_download_button(fig_nas_bar235, "nasional_komposisi_bar.png", "dl_235_nas_bar")
    with colD3:
        df_download_button(
            nasional_235.rename("Rata-rata Nasional (%)").reset_index().rename(columns={"index": "Jenis Sampah"}),
            "Unduh Angka Nasional (CSV)", "angka_nasional_komposisi_sampah.csv", "dl_235_nasional_csv",
        )
    st.divider()

    tab1, tab2 = st.tabs(["🥧 Satu Provinsi (Detail)", "📊 Perbandingan Antar Provinsi"])

    with tab1:
        opsi_provinsi = [NASIONAL_LABEL] + df_235["Provinsi"].tolist()
        prov = st.selectbox("Pilih provinsi", opsi_provinsi, key="prov_235_single")
        if prov == NASIONAL_LABEL:
            data = nasional_235
        else:
            row = df_235[df_235["Provinsi"] == prov].iloc[0]
            data = row[kategori].astype(float)
        fig = px.pie(
            names=data.index, values=data.values, hole=0.4,
            template=TEMPLATE, color_discrete_sequence=PALETTE,
            title=f"Komposisi Sampah — {prov}",
        )
        fig.update_layout(height=480)
        st.plotly_chart(fig, use_container_width=True)
        fig_download_button(fig, f"komposisi_sampah_{prov}.png", "dl_235_fig1")

    with tab2:
        provinsi_pilihan = st.multiselect(
            "Pilih provinsi untuk dibandingkan", df_235["Provinsi"].tolist(),
            default=df_235["Provinsi"].tolist()[:8],
        )
        kategori_pilihan = st.multiselect("Pilih jenis sampah", kategori, default=kategori)
        sertakan_nasional = st.checkbox("Sertakan Rata-rata Nasional sebagai pembanding", value=True, key="nasional_235_tab2")

        if provinsi_pilihan and kategori_pilihan:
            dff = df_235[df_235["Provinsi"].isin(provinsi_pilihan)]
            if sertakan_nasional:
                dff = pd.concat([dff, national_row(df_235, kategori)], ignore_index=True)
            melted = dff.melt(id_vars="Provinsi", value_vars=kategori_pilihan, var_name="Jenis Sampah", value_name="Persentase")
            fig2 = px.bar(
                melted, x="Persentase", y="Provinsi", color="Jenis Sampah",
                orientation="h", template=TEMPLATE, color_discrete_sequence=PALETTE,
                barmode="stack",
            )
            fig2.update_layout(height=max(420, 30 * len(provinsi_pilihan)), legend=dict(orientation="h", yanchor="bottom", y=1.02))
            st.plotly_chart(fig2, use_container_width=True)
            colA, colB = st.columns(2)
            with colA:
                fig_download_button(fig2, "komposisi_sampah_perbandingan.png", "dl_235_fig2")
            with colB:
                df_download_button(dff, "Unduh data (CSV)", "komposisi_sampah.csv", "dl_235_csv1")
        else:
            st.warning("Pilih minimal satu provinsi dan satu jenis sampah.")

    st.divider()
    st.dataframe(df_235, use_container_width=True, hide_index=True)


# ========================================================================================
# HALAMAN: SUMBER TIMBULAN SAMPAH
# ========================================================================================
elif page == "🏷️ Sumber Timbulan Sampah":
    st.title("🏷️ Sumber Timbulan Sampah per Provinsi")
    kategori = [c for c in df_236.columns if c != "Provinsi"]

    st.subheader("🇮🇩 Angka Nasional per Sumber Sampah (Rata-rata Semua Provinsi)")
    nasional_236 = df_236[kategori].mean(numeric_only=True).round(2)
    nasional_236_sorted = nasional_236.sort_values()

    colN1, colN2 = st.columns([1, 1.2])
    with colN1:
        fig_nas_pie236 = px.pie(
            names=nasional_236.index, values=nasional_236.values, hole=0.45,
            template=TEMPLATE, color_discrete_sequence=PALETTE,
            title="Proporsi Rata-rata Nasional",
        )
        fig_nas_pie236.update_traces(textinfo="percent+label", textposition="outside")
        fig_nas_pie236.update_layout(height=420, showlegend=False, margin=dict(t=60, b=10, l=10, r=10))
        st.plotly_chart(fig_nas_pie236, use_container_width=True)
    with colN2:
        fig_nas_bar236 = px.bar(
            x=nasional_236_sorted.values, y=nasional_236_sorted.index, orientation="h",
            template=TEMPLATE, color=nasional_236_sorted.values, color_continuous_scale="Oranges",
            text=nasional_236_sorted.values, labels={"x": "Rata-rata Nasional", "y": ""},
            title="Peringkat Rata-rata Nasional",
        )
        fig_nas_bar236.update_traces(texttemplate="%{text:.1f}", textposition="outside")
        fig_nas_bar236.update_layout(height=420, coloraxis_showscale=False, margin=dict(t=60, b=10, l=10, r=10))
        st.plotly_chart(fig_nas_bar236, use_container_width=True)
    st.caption(
        "Catatan: 'Rumah Tangga' bernilai jauh lebih besar dari sumber lain karena satuan datanya "
        "mencakup skala timbulan yang lebih luas dibanding sumber lain seperti Perkantoran atau Pasar."
    )

    colD1, colD2, colD3 = st.columns(3)
    with colD1:
        fig_download_button(fig_nas_pie236, "nasional_sumber_pie.png", "dl_236_nas_pie")
    with colD2:
        fig_download_button(fig_nas_bar236, "nasional_sumber_bar.png", "dl_236_nas_bar")
    with colD3:
        df_download_button(
            nasional_236.rename("Rata-rata Nasional").reset_index().rename(columns={"index": "Sumber"}),
            "Unduh Angka Nasional (CSV)", "angka_nasional_sumber_sampah.csv", "dl_236_nasional_csv",
        )
    st.divider()

    tab1, tab2 = st.tabs(["🥧 Satu Provinsi (Detail)", "📊 Perbandingan Antar Provinsi"])

    with tab1:
        opsi_provinsi = [NASIONAL_LABEL] + df_236["Provinsi"].tolist()
        prov = st.selectbox("Pilih provinsi", opsi_provinsi, key="prov_236_single")
        if prov == NASIONAL_LABEL:
            data = nasional_236
        else:
            row = df_236[df_236["Provinsi"] == prov].iloc[0]
            data = row[kategori].astype(float)
        fig = px.bar(
            x=data.values, y=data.index, orientation="h", template=TEMPLATE,
            color=data.index, color_discrete_sequence=PALETTE,
            labels={"x": "Nilai", "y": "Sumber"}, title=f"Sumber Sampah — {prov}",
        )
        fig.update_layout(height=440, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
        fig_download_button(fig, f"sumber_sampah_{prov}.png", "dl_236_fig1")

    with tab2:
        provinsi_pilihan = st.multiselect(
            "Pilih provinsi untuk dibandingkan", df_236["Provinsi"].tolist(),
            default=df_236["Provinsi"].tolist()[:8], key="prov_236_multi",
        )
        normalisasi = st.checkbox("Tampilkan sebagai proporsi 100% (dinormalisasi)", value=True)
        sertakan_nasional = st.checkbox("Sertakan Rata-rata Nasional sebagai pembanding", value=True, key="nasional_236_tab2")

        if provinsi_pilihan:
            dff = df_236[df_236["Provinsi"].isin(provinsi_pilihan)].copy()
            if normalisasi:
                totals = dff[kategori].sum(axis=1)
                for c in kategori:
                    dff[c] = dff[c] / totals * 100
            if sertakan_nasional:
                nas_df = national_row(df_236, kategori)
                if normalisasi:
                    tot_nas = nas_df[kategori].sum(axis=1)
                    for c in kategori:
                        nas_df[c] = nas_df[c] / tot_nas * 100
                dff = pd.concat([dff, nas_df], ignore_index=True)
            melted = dff.melt(id_vars="Provinsi", value_vars=kategori, var_name="Sumber", value_name="Nilai")
            fig2 = px.bar(
                melted, x="Nilai", y="Provinsi", color="Sumber", orientation="h",
                template=TEMPLATE, color_discrete_sequence=PALETTE, barmode="stack",
            )
            fig2.update_layout(
                height=max(420, 30 * len(provinsi_pilihan)),
                xaxis_title="Persentase (%)" if normalisasi else "Nilai",
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
            )
            st.plotly_chart(fig2, use_container_width=True)
            colA, colB = st.columns(2)
            with colA:
                fig_download_button(fig2, "sumber_sampah_perbandingan.png", "dl_236_fig2")
            with colB:
                df_download_button(dff, "Unduh data (CSV)", "sumber_sampah.csv", "dl_236_csv1")
        else:
            st.warning("Pilih minimal satu provinsi.")

    st.divider()
    st.dataframe(df_236, use_container_width=True, hide_index=True)


# ========================================================================================
# HALAMAN: PERBANDINGAN ANTAR PROVINSI
# ========================================================================================
elif page == "⚖️ Perbandingan Antar Provinsi":
    st.title("⚖️ Perbandingan Menyeluruh Antar Provinsi")
    st.caption("Menggabungkan indikator timbulan sampah, komposisi, dan sumber untuk provinsi terpilih.")

    all_provinsi = df_234["Provinsi"].tolist()
    provinsi_pilihan = st.multiselect(
        "Pilih 2–6 provinsi untuk dibandingkan", all_provinsi, default=all_provinsi[:3], max_selections=6,
    )

    if len(provinsi_pilihan) < 2:
        st.warning("Pilih minimal dua provinsi untuk perbandingan.")
        st.stop()

    sertakan_nasional_cmp = st.checkbox("Sertakan Rata-rata Nasional dalam perbandingan", value=True, key="nasional_cmp")

    merged = df_234[df_234["Provinsi"].isin(provinsi_pilihan)].copy()
    if sertakan_nasional_cmp:
        merged = pd.concat(
            [merged, national_row(df_234, [c for c in df_234.columns if c != "Provinsi"])],
            ignore_index=True,
        )
    st.subheader("Indikator Utama")
    fig = go.Figure()
    metrics = ["Pengurangan Sampah (%)", "Penanganan Sampah (%)", "Sampah Terkelola (%)"]
    for _, r in merged.iterrows():
        fig.add_trace(go.Scatterpolar(
            r=[r[m] for m in metrics] + [r[metrics[0]]],
            theta=metrics + [metrics[0]],
            fill="toself", name=r["Provinsi"],
        ))
    fig.update_layout(template=TEMPLATE, height=500, polar=dict(radialaxis=dict(visible=True, range=[0, 100])))
    st.plotly_chart(fig, use_container_width=True)
    fig_download_button(fig, "radar_perbandingan_provinsi.png", "dl_cmp_fig1")

    st.divider()
    st.subheader("Komposisi Sampah — Perbandingan")
    kat5 = [c for c in df_235.columns if c != "Provinsi"]
    dff5 = df_235[df_235["Provinsi"].isin(provinsi_pilihan)]
    if sertakan_nasional_cmp:
        dff5 = pd.concat([dff5, national_row(df_235, kat5)], ignore_index=True)
    melted5 = dff5.melt(id_vars="Provinsi", value_vars=kat5, var_name="Jenis Sampah", value_name="Persentase")
    fig2 = px.bar(
        melted5, x="Provinsi", y="Persentase", color="Jenis Sampah",
        template=TEMPLATE, color_discrete_sequence=PALETTE, barmode="stack",
    )
    fig2.update_layout(height=480, legend=dict(orientation="h", yanchor="bottom", y=1.02))
    st.plotly_chart(fig2, use_container_width=True)
    fig_download_button(fig2, "komposisi_perbandingan.png", "dl_cmp_fig2")

    st.divider()
    st.subheader("Sumber Sampah — Perbandingan (dinormalisasi %)")
    kat6 = [c for c in df_236.columns if c != "Provinsi"]
    dff6 = df_236[df_236["Provinsi"].isin(provinsi_pilihan)].copy()
    totals6 = dff6[kat6].sum(axis=1)
    for c in kat6:
        dff6[c] = dff6[c] / totals6 * 100
    if sertakan_nasional_cmp:
        nas6 = national_row(df_236, kat6)
        tot_nas6 = nas6[kat6].sum(axis=1)
        for c in kat6:
            nas6[c] = nas6[c] / tot_nas6 * 100
        dff6 = pd.concat([dff6, nas6], ignore_index=True)
    melted6 = dff6.melt(id_vars="Provinsi", value_vars=kat6, var_name="Sumber", value_name="Persentase")
    fig3 = px.bar(
        melted6, x="Provinsi", y="Persentase", color="Sumber",
        template=TEMPLATE, color_discrete_sequence=PALETTE, barmode="stack",
    )
    fig3.update_layout(height=480, legend=dict(orientation="h", yanchor="bottom", y=1.02))
    st.plotly_chart(fig3, use_container_width=True)
    fig_download_button(fig3, "sumber_perbandingan.png", "dl_cmp_fig3")

    st.divider()
    st.subheader("Tabel Ringkasan")
    st.dataframe(merged, use_container_width=True, hide_index=True)
    df_download_button(merged, "Unduh Ringkasan (CSV)", "perbandingan_provinsi.csv", "dl_cmp_csv1")


# --------------------------------------------------------------------------------------
# FOOTER
# --------------------------------------------------------------------------------------
st.sidebar.divider()
st.sidebar.caption("Dibuat dengan ❤️ menggunakan Streamlit & Plotly.")
