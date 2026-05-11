import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

st.set_page_config(page_title="Review Akses Dashboard", layout="wide")

st.title("📊 Dashboard Review Akses")

# =========================
# HELPER FUNCTIONS
# =========================

def normalize_columns(df):
    df.columns = [str(col).strip().lower() for col in df.columns]
    return df


def find_column(df, possible_names):
    for col in df.columns:
        for name in possible_names:
            if name.lower() in col.lower():
                return col
    return None


# =========================
# SIDEBAR UPLOAD
# =========================

st.sidebar.header("Dashboard")

menu = st.sidebar.selectbox(
    "Pilih Menu",
    [
        "Portalisasi",
        "Database"
    ]
)

st.sidebar.markdown("---")

if menu == "Portalisasi":
    export_file = st.sidebar.file_uploader(
        "Upload Akses History", type=["xlsx"]
    )
    
elif menu == "Database":
    reqdb_file = st.sidebar.file_uploader(
        "Upload Req DB", type=["xlsx"]
    )


# =========================
# REVIEW AKSES PORTALISASI
# =========================

if menu == "Portalisasi" and export_file:
    st.header("Portalisasi")

    xls_export = pd.ExcelFile(export_file)

    # Semua sheet kecuali Summary
    sheets = [s for s in xls_export.sheet_names if s.lower() != "summary"]

    all_data = []

    for sheet in sheets:
        try:
            df = pd.read_excel(export_file, sheet_name=sheet)
            # Jika sheet kosong tetap dibuat dataframe kosong
            if df.empty:
                df = pd.DataFrame({
                    "Aplikasi": [sheet],
                    "Status Administrasi": ["Belum Teradministrasi"]
                })

            else:
                df = normalize_columns(df)

            # Tambahkan nama aplikasi dari nama sheet
            df["Aplikasi"] = sheet

            all_data.append(df)

        except Exception as e:
            st.warning(f"Gagal membaca sheet {sheet}: {e}")

    if all_data:
        portal_df = pd.concat(all_data, ignore_index=True)
        all_apps = pd.DataFrame({
            "Aplikasi" : sheets
        })

        # =========================
        # KOLOM YANG DIGUNAKAN
        # =========================

        created_col = find_column(
            portal_df,
            ["created"]
        )

        referensi_col = find_column(
            portal_df,
            ["referensi"]
        )

        # =========================
        # VALIDASI
        # =========================

        if not created_col:
            st.error("Kolom CREATED tidak ditemukan")

        elif not referensi_col:
            st.error("Kolom REFERENSI tidak ditemukan")

        else:
            # =========================
            # NORMALISASI STATUS
            # =========================

            def determine_status(row):
                 # CREATED kosong -> abaikan
                created_val = row[created_col]

                if pd.isna(created_val):
                    return None

                if str(created_val).strip() == "":
                    return None

                # REFERENSI
                ref_val = row[referensi_col]

                # referensi kosong = belum
                if pd.isna(ref_val):
                    return "Belum Teradministrasi"

                ref_val = str(ref_val).strip()

                if ref_val == "":
                    return "Belum Teradministrasi"

                # jika ada tiket/reference
                return "Sudah Teradministrasi"

            portal_df["Status Administrasi"] = portal_df.apply(determine_status, axis=1)
            portal_df = portal_df[
                portal_df["Status Administrasi"].notna()
            ]

            # =========================
            # SUMMARY STATUS
            # =========================

            summary_status = (
                portal_df["Status Administrasi"]
                .value_counts()
                .reset_index()
            )

            summary_status.columns = ["Status", "Jumlah"]

            sudah = chart_sudah = (
                portal_df[portal_df["Status Administrasi"] == "Sudah Teradministrasi"]
                .shape[0]
            )

            belum = (
                portal_df[portal_df["Status Administrasi"] == "Belum Teradministrasi"]
                .shape[0]
            )

            total = sudah + belum

            persen = round((sudah / total) * 100, 1) if total > 0 else 0
            
            def get_status_label(persen):
                """Fungsi untuk menentukan label status berdasarkan persentase"""
                if persen <= 25:
                    return "Sangat Rendah"
                elif persen <= 50:
                    return "Cukup Rendah"
                elif persen <= 75:
                    return "Cukup Tinggi"
                else:
                    return "Sangat Tinggi"

            # =========================
            # CHART 1 - BASED ON APLIKASI
            # =========================

            col1, col2 = st.columns(2)
            
            with col1:
                app_chart = (
                    portal_df.groupby(["Aplikasi", "Status Administrasi"])
                    .size()
                    .reset_index(name="Jumlah")
                )
                
                # memastikan sheet kosong tetap muncul
                app_chart = all_apps.merge(
                    app_chart,
                    on="Aplikasi",
                    how="left"
                )

                app_chart["Jumlah"] = app_chart["Jumlah"].fillna(0)
                app_chart["Status Administrasi"] = app_chart["Status Administrasi"].fillna("Belum Teradministrasi")

                fig_app = px.bar(
                    app_chart,
                    x="Aplikasi",
                    y="Jumlah",
                    color="Status Administrasi",
                    barmode="group",
                    title="Based on Aplikasi",
                    text_auto=True,
                )
                
                for trace in fig_app.data:
                    trace.textposition=[
                        "outside" if val < 10 else "inside"
                        for val in trace.y
                    ]
                
                fig_app.update_layout(
                    title_x=0.5,
                    title_xanchor="center",
                    legend=dict(
                        orientation="h",
                        yanchor="top",
                        y=-0.5,
                        xanchor="center",
                        x=0.5,
                    )
                )

                st.plotly_chart(fig_app, use_container_width=True)

            # =========================
            # CHART 2 - BASED ON PIC (CREATED)
            # =========================
            
            with col2:
                pic_chart = (
                    portal_df.groupby([created_col, "Status Administrasi"])
                    .size()
                    .reset_index(name="Jumlah")
                )

                fig_pic = px.bar(
                    pic_chart,
                    x=created_col,
                    y="Jumlah",
                    color="Status Administrasi",
                    barmode="group",
                    title="Based on PIC",
                    text_auto=True,
                )
                
                for trace in fig_pic.data:
                    trace.textposition=[
                        "outside" if val < 10 else "inside"
                        for val in trace.y
                    ]

                fig_pic.update_layout(
                    xaxis_title="",
                    title_x=0.5,
                    title_xanchor="center",
                    legend=dict(
                        orientation="h",
                        yanchor="top",
                        y=-0.5,
                        xanchor="center",
                        x=0.5
                    )
                )

                st.plotly_chart(fig_pic, use_container_width=True)

            # =========================
            # CHART 3 - DONUT KETERTIBAN
            # =========================
            
            col3, col4 = st.columns([1,2])
            
            with col3:
                fig_donut1 = go.Figure(
                    data=[
                        go.Pie(
                            labels=[
                                "Sudah Teradministrasi",
                                "Belum Teradministrasi",
                            ],
                            values=[sudah, belum],
                            hole=0.6,
                            textinfo="percent",
                        )
                    ]
                )

                fig_donut1.update_layout(
                    height=400,
                    legend=dict(
                        orientation="h",
                        yanchor="top",
                        y=1.3,
                        xanchor="center",
                        x=0.5
                    )
                )

                st.plotly_chart(fig_donut1, use_container_width=True)

            # =========================
            # PIVOT TABLE
            # =========================

            with col4:
                status_label = get_status_label(persen)

                st.markdown(
                    f"""
                    ### Ketertiban Administrasi

                    Hasil review aktivitas akses pegawai pada beberapa aplikasi menunjukkan bahwa tingkat keteraturan administrasi akses masih berada pada kategori **{status_label.lower()}**, dengan persentase akses yang telah teradministrasi sebesar **{persen}%** dari keseluruhan aktivitas akses yang berhasil direview.
                    """
                )

                pivot_table = pd.pivot_table(
                    portal_df,
                    index="Status Administrasi",
                    columns="Aplikasi",
                    aggfunc="size",
                    fill_value=0,
                )

                pivot_table = pivot_table.reindex(
                    columns=sheets,
                    fill_value=0
                )

                pivot_table["Total"] = pivot_table.sum(axis=1)
                
                total_per_aplikasi = pivot_table.sum(axis=0)
                
                pivot_table.loc[
                    "Total Per Applikasi"
                ] = total_per_aplikasi
                
                persen_row = []

                for col in pivot_table.columns:

                    # persen tiap aplikasi
                    if col in sheets:

                        sudah_app = pivot_table.loc[
                            "Sudah Teradministrasi",
                            col
                        ]

                        belum_app = pivot_table.loc[
                            "Belum Teradministrasi",
                            col
                        ]

                        total_app = sudah_app + belum_app

                        if total_app > 0:
                            persen_app = round(
                                (sudah_app / total_app) * 100,
                                1
                            )
                        else:
                            persen_app = 0

                        persen_row.append(
                            f"{persen_app}%"
                        )

                    # persen total keseluruhan
                    elif col == "Total":

                        if total > 0:
                            total_percent = round(
                                (sudah / total) * 100,
                                1
                            )
                        else:
                            total_percent = 0

                        persen_row.append(
                            f"{total_percent}%"
                        )

                pivot_table.loc[
                    "Persentase Administrasi"
                ] = persen_row

                # =========================
                # RAPIIHKAN INDEX
                # =========================

                pivot_table.index.name = ""

                st.dataframe(pivot_table, use_container_width=True)

# =========================
# REVIEW AKSES DATABASE
# =========================

if menu == "Database" and reqdb_file:
    st.header("Database")

    try:
        # =========================
        # READ ALL SHEETS
        # =========================

        excel_file = pd.ExcelFile(reqdb_file)

        all_data = []

        for sheet in excel_file.sheet_names:

            try:

                raw_df = pd.read_excel(
                    reqdb_file,
                    sheet_name=sheet,
                    header=None
                )

                header_row = None

                for i in range(len(raw_df)):

                    row_values = (
                        raw_df.iloc[i]
                        .astype(str)
                        .str.lower()
                        .str.strip()
                        .tolist()
                    )

                    if any(
                        "masa berlaku dari" in v
                        for v in row_values
                    ):
                        header_row = i
                        break

                if header_row is None:
                    continue

                temp_df = pd.read_excel(
                    reqdb_file,
                    sheet_name=sheet,
                    header=header_row
                )

                temp_df = normalize_columns(temp_df)

                temp_df = temp_df.dropna(
                    how="all"
                )

                temp_df = temp_df.dropna(
                    axis=1,
                    how="all"
                )

                # =========================
                # SAVE SHEET NAME
                # =========================

                temp_df["Sheet_Bulan"] = sheet

                all_data.append(temp_df)

            except:
                continue

        # =========================
        # COMBINE ALL SHEETS
        # =========================

        if len(all_data) == 0:
            st.error("Tidak ada data valid")
            st.stop()

        db_df = pd.concat(
            all_data,
            ignore_index=True
        )

        masa_dari_col = find_column(db_df, ["masa berlaku dari"])
        masa_sd_col = find_column(db_df, [ "masa berlaku s.d", "masa berlaku sd"])
        keterangan_col = find_column(db_df, ["keterangan"])

        if not masa_dari_col:
            st.error("Kolom MASA BERLAKU DARI tidak ditemukan")

        elif not masa_sd_col:
            st.error("Kolom MASA BERLAKU S.D tidak ditemukan")

        else:

            db_df[masa_dari_col] = pd.to_datetime(
                db_df[masa_dari_col],
                errors="coerce"
            )

            db_df[masa_sd_col] = pd.to_datetime(
                db_df[masa_sd_col],
                errors="coerce"
            )
            
            today = pd.Timestamp.today().normalize()

            # =========================
            # MAPPING STATUS
            # =========================

            def map_db_status(row):

                masa_dari = row[masa_dari_col]
                masa_sd = row[masa_sd_col]

                if keterangan_col:

                    ket = str(
                        row[keterangan_col]
                    ).lower().strip()

                    if "resign" in ket:
                        return "Penutupan Akses"
                    
                if (
                    pd.isna(masa_dari)
                    or pd.isna(masa_sd)
                ):
                    return "Tiket Not Found"
                
                if masa_sd < today:
                    return "Expired"
                
                return "Aktif"

            db_df["Status Review"] = db_df.apply(
                map_db_status,
                axis=1
            )

            # =========================
            # CHART DATA
            # =========================

            bulan_order = [
                "Januari",
                "Februari",
                "Maret",
                "April",
                "Mei",
                "Juni",
                "Juli",
                "Agustus",
                "September",
                "Oktober",
                "November",
                "Desember",
            ]

            def detect_bulan(sheet_name):

                sheet_name = str(sheet_name).lower()

                for bulan in bulan_order:

                    if bulan.lower() in sheet_name:
                        return bulan

                return None

            db_df["Bulan"] = (
                db_df["Sheet_Bulan"]
                .apply(detect_bulan)
            )

            chart_data = (
                db_df.groupby(
                    ["Bulan", "Status Review"]
                )
                .size()
                .reset_index(name="Jumlah")
            )
            
            bulan_tersedia = (
                db_df["Bulan"]
                .dropna()
                .unique()
                .tolist()
            )

            bulan_order = [
                b for b in bulan_order
                if b in bulan_tersedia
            ]

            colors = {
                "Aktif": "#4F7DF3",
                "Expired": "#A483FF",
                "Tiket Not Found": "#F0B35A",
                "Penutupan Akses": "#F04F4F",
            }
            
            # =========================
            # TOTAL DATA
            # =========================

            aktif = (
                db_df["Status Review"] == "Aktif"
            ).sum()

            expired = (
                db_df["Status Review"] == "Expired"
            ).sum()

            tiket_not_found = (
                db_df["Status Review"]
                == "Tiket Not Found"
            ).sum()

            penutupan = (
                db_df["Status Review"]
                == "Penutupan Akses"
            ).sum()
            
            total_host = (
                aktif + expired + tiket_not_found + penutupan
            )

            col1, col2 = st.columns([1.5, 1])

            # =========================
            # BAR CHART
            # =========================

            with col1:

                fig_db = go.Figure()

                for status in [
                    "Aktif",
                    "Expired",
                    "Tiket Not Found",
                    "Penutupan Akses",
                ]:

                    temp_df = chart_data[
                        chart_data["Status Review"] == status
                    ]

                    fig_db.add_trace(
                        go.Bar(
                            x=temp_df["Bulan"],
                            y=temp_df["Jumlah"],
                            name=status,
                            marker_color=colors[status],
                            text=temp_df["Jumlah"],
                            textposition="inside"
                        )
                    )

                fig_db.update_layout(
                    barmode="stack",
                    title="Review Akses Database",
                    title_x=0.5,
                    title_xanchor="center",
                    height=500,
                    xaxis=dict(
                        categoryorder="array",
                        categoryarray=bulan_order
                    ),
                    legend=dict(
                        orientation="h",
                        yanchor="top",
                        y=-0.2,
                        xanchor="center",
                        x=0.5
                    )
                )

                st.plotly_chart(
                    fig_db,
                    use_container_width=True
                )

            # =========================
            # RINGKASAN
            # =========================

            with col2:

                st.subheader("Ringkasan Review")

                bulan_list = [
                    b for b in bulan_order
                    if b in db_df["Bulan"].unique()
                ]

                selected_bulan = st.selectbox(
                    "Pilih Bulan",
                    bulan_list
                )

                bulan_df = db_df[
                    db_df["Bulan"] == selected_bulan
                ]

                aktif_bulan = (
                    bulan_df["Status Review"]
                    == "Aktif"
                ).sum()

                expired_bulan = (
                    bulan_df["Status Review"]
                    == "Expired"
                ).sum()

                tiket_bulan = (
                    bulan_df["Status Review"]
                    == "Tiket Not Found"
                ).sum()

                penutupan_bulan = (
                    bulan_df["Status Review"]
                    == "Penutupan Akses"
                ).sum()

                total_bulan = (
                    aktif_bulan
                    + expired_bulan
                    + tiket_bulan
                    + penutupan_bulan
                )

                st.metric(
                    "Total Host",
                    total_bulan
                )

                c1, c2 = st.columns(2)

                with c1:

                    st.metric(
                        "Aktif",
                        aktif_bulan
                    )

                    st.metric(
                        "Expired",
                        expired_bulan
                    )

                with c2:

                    st.metric(
                        "Tiket Not Found",
                        tiket_bulan
                    )

                    st.metric(
                        "Penutupan Akses",
                        penutupan_bulan
                    )

    except Exception as e:
        st.error(f"Gagal membaca file Req DB: {e}")

# =========================
# FOOTER
# =========================

st.markdown("---")
st.caption("Dashboard Review Akses")