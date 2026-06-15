import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# =========================================================
# CONFIG
# =========================================================

st.set_page_config(page_title="Dashboard Review Akses", layout="wide")
st.title("📊 Dashboard Review Akses")

# =========================================================
# CACHE
# =========================================================

@st.cache_data(show_spinner=False)
def read_excel_cached(file, header=0, sheet_name=0):
    return pd.read_excel(file, header=header, sheet_name=sheet_name, engine="openpyxl")

# =========================================================
# HELPER
# =========================================================

def clean_col(df):
    df.columns = (df.columns.astype(str).str.strip().str.upper())
    return df

def normalize_nipp(series):
    return (series.astype(str).str.strip())

def normalize_text(text):
    if pd.isna(text):
        return ""
    return (str(text).strip().lower().replace(".", "").replace("-", "").replace("_", "").replace(" ", ""))

def clean_nama(nama):
    if pd.isna(nama):
        return ""
    return (str(nama).upper().replace(".", "").replace(",", "").replace("-", " ").replace("_", " ").strip())

def filter_login(df):
    if "AKTIVITAS" not in df.columns:
        return set()
    return set(
        df.loc[df["AKTIVITAS"].astype(str).str.contains("S-AKSES LOGIN DIIJINKAN", case=False, na=False), "NIPP"])

def find_column(df, possible_names):
    for col in df.columns:
        for name in possible_names:
            if name.lower() in col.lower():
                return col
    return None

# =========================================================
# SIDEBAR MENU
# =========================================================

menu = st.sidebar.selectbox("Pilih Menu", ["Portalisasi", "Database", "User Tidak Aktif"])

# =========================================================
# PORTALISASI
# =========================================================
MASTER_APLIKASI = ['VASA', 'GENC', 'SPINER', 'MDM', 'INCO', 'PTOSM']
DISPLAY_APLIKASI = {'VASA': 'VASA', 'GENC': 'GENC', 'SPINER': 'SPINER', 'MDM': 'MDM', 'INCO': 'INCO', 'PTOSM': 'PTOS-M'}

def clean_nama(nama):
    if pd.isna(nama):
        return ""

    nama = str(nama).upper()
    nama = (nama.replace(".", "").replace(",", "").replace("-", " ").replace("_", " ").replace("  ", " ").strip())
    return nama

def normalize_text(text):
    if pd.isna(text):
        return ""
    return (str(text).strip().lower().replace(".", "").replace("-", "").replace("_", "").replace(" ", ""))

if menu == "Portalisasi":
    st.sidebar.header("Upload File")

    file_user = st.sidebar.file_uploader("Upload ExportHistoryAkses", type=["xlsx"])
    file_incident = st.sidebar.file_uploader("Upload Incident", type=["xlsx"])
    file_sc_req = st.sidebar.file_uploader("Upload SC_REQ_ITEM", type=["xlsx"])

    if file_user and file_incident and file_sc_req:
        with st.spinner("Memproses data..."):
            df_user_all = pd.read_excel(file_user, sheet_name=None)
            
            df_incident1_all = pd.read_excel(file_incident, sheet_name=None)
            df_incident2_all = pd.read_excel(file_sc_req, sheet_name=None)
            
            df_incident1 = pd.concat(df_incident1_all.values(), ignore_index=True)
            df_incident2 = pd.concat(df_incident2_all.values(), ignore_index=True)

            for sheet_name, df_sheet in df_user_all.items():
                if df_sheet.empty:
                    continue
                
                if 'CREATED' in df_sheet.columns:
                    df_sheet['CREATED'] = (df_sheet['CREATED'].fillna("Unknown").replace("", "Unknown"))
                df_user_all[sheet_name] = df_sheet

            service_alias = {'VASA': ['VASA', 'VASA.'], 'GENC': ['GENC', 'Gen C'], 'SPINER': ['SPINER'], 'MDM': ['MDM', 'Master Data Management', 'Management Data Master'], 'INCO': ['INCO', 'Inco'], 'PTOSM': ['PTOSM', 'PTOS-M', 'PTOS M']}

            df_incident1_subset = df_incident1.iloc[:, [0, 5, 20]].copy()
            df_incident1_subset.columns = ['No. Tiket', 'Nama', 'Service offering']

            df_incident2_subset = df_incident2.iloc[:, [0, 6, 19]].copy()
            df_incident2_subset.columns = ['No. Tiket', 'Nama', 'Service offering']

            df_incident1_subset['Nama'] = (df_incident1_subset['Nama'].apply(clean_nama))
            df_incident2_subset['Nama'] = (df_incident2_subset['Nama'].apply(clean_nama))

            df_all_incident = pd.concat([df_incident1_subset, df_incident2_subset], ignore_index=True)
            df_all_incident = (df_all_incident.dropna(subset=['No. Tiket', 'Nama']))

            df_all_incident = (df_all_incident.drop_duplicates(subset=['Nama', 'Service offering']))

            for sheet_name, df_sheet in df_user_all.items():
                if (df_sheet.empty or len(df_sheet.columns) < 11):
                    continue

                col_nm = df_sheet.columns[3]
                col_ref = df_sheet.columns[10]

                alias_list = service_alias.get(sheet_name, [sheet_name])

                referensi_list = []

                for idx, row in df_sheet.iterrows():
                    nama = clean_nama(row[col_nm])
                    match = df_all_incident[(df_all_incident['Nama'] == nama) & (df_all_incident['Service offering'].apply(normalize_text).isin([normalize_text(x) for x in alias_list]))]

                    if not match.empty:
                        tiket = (match.iloc[0]['No. Tiket'])
                    else:
                        tiket = ""
                    referensi_list.append(tiket)
                df_sheet[col_ref] = referensi_list
                df_user_all[sheet_name] = df_sheet

            summary_data = []

            for aplikasi in MASTER_APLIKASI:
                if aplikasi in df_user_all:
                    df_sheet = df_user_all[aplikasi]
                    if (not df_sheet.empty and 'REFERENSI' in df_sheet.columns):
                        df_sheet['REFERENSI'] = (df_sheet['REFERENSI'].fillna("").astype(str))

                        sudah = int((df_sheet['REFERENSI'].str.strip()!= "").sum())
                        total = len(df_sheet)
                        belum = total - sudah
                    else:
                        sudah = 0
                        belum = 0
                        total = 0
                else:
                    sudah = 0
                    belum = 0
                    total = 0
                persen_sudah = ((sudah / total) * 100 if total > 0 else 0)
                persen_belum = ((belum / total) * 100 if total > 0 else 0)
                summary_data.append({'Aplikasi': DISPLAY_APLIKASI.get(aplikasi, aplikasi),
                    'Sudah Teradministrasi': int(
                        sudah
                    ),

                    'Belum Teradministrasi': int(
                        belum
                    ),

                    'Total': int(
                        total
                    ),

                    '% Sudah': (
                        f"{persen_sudah:.0f}%"
                    ),

                    '% Belum': (
                        f"{persen_belum:.0f}%"
                    )
                })

            df_summary = pd.DataFrame(
                summary_data
            )

            # =========================================================
            # TOTAL
            # =========================================================
            total_sudah = (
                df_summary[
                    'Sudah Teradministrasi'
                ].sum()
            )

            total_belum = (
                df_summary[
                    'Belum Teradministrasi'
                ].sum()
            )

            total_all = (
                df_summary[
                    'Total'
                ].sum()
            )

            persen_total_sudah = (
                (total_sudah / total_all) * 100
                if total_all > 0 else 0
            )

            persen_total_belum = (
                (total_belum / total_all) * 100
                if total_all > 0 else 0
            )

            total_row = pd.DataFrame([{

                'Aplikasi': 'TOTAL',

                'Sudah Teradministrasi': int(
                    total_sudah
                ),

                'Belum Teradministrasi': int(
                    total_belum
                ),

                'Total': int(
                    total_all
                ),

                '% Sudah': (
                    f"{persen_total_sudah:.0f}%"
                ),

                '% Belum': (
                    f"{persen_total_belum:.0f}%"
                )
            }])

            df_summary = pd.concat(
                [
                    df_summary,
                    total_row
                ],
                ignore_index=True
            )

            # =========================================================
            # SUMMARY PIC
            # =========================================================
            pic_data = []

            for sheet_name, df_sheet in df_user_all.items():

                if (
                    df_sheet.empty
                    or 'CREATED' not in df_sheet.columns
                ):
                    continue

                temp = pd.DataFrame({

                    'PIC': df_sheet['CREATED'],

                    'Sudah Teradministrasi':
                        (
                            df_sheet['REFERENSI']
                            .fillna("")
                            .astype(str)
                            .str.strip()
                            != ""
                        ).astype(int),

                    'Belum Teradministrasi':
                        (
                            df_sheet['REFERENSI']
                            .fillna("")
                            .astype(str)
                            .str.strip()
                            == ""
                        ).astype(int)
                })

                pic_data.append(temp)

            if len(pic_data) > 0:

                df_pic = pd.concat(
                    pic_data,
                    ignore_index=True
                )

                df_summary_pic = (
                    df_pic
                    .groupby('PIC')
                    .sum()
                    .reset_index()
                )

            else:

                df_summary_pic = pd.DataFrame()

            # =========================================================
            # DONUT
            # =========================================================
            col1, col2 = st.columns([1, 2])

            with col1:

                st.subheader(
                    "Ketertiban Administrasi"
                )

                fig_donut = go.Figure(
                    data=[
                        go.Pie(
                            labels=[
                                'Belum Teradministrasi',
                                'Sudah Teradministrasi'
                            ],
                            values=[
                                total_belum,
                                total_sudah
                            ],
                            hole=0.7
                        )
                    ]
                )

                fig_donut.update_layout(
                    height=400
                )

                st.plotly_chart(fig_donut, use_container_width=True)

            with col2:
                st.markdown(f""" ## Ketertiban Administrasi Persentase akses yang telah teradministrasi sebesar **{persen_total_sudah:.0f}%** dari keseluruhan aktivitas akses yang berhasil direview. """)
                st.dataframe(df_summary, use_container_width=True)

            st.divider()
            left_chart, right_chart = st.columns(2)

            with left_chart:
                st.markdown("<h2 style='text-align: center;'>Based on Aplikasi</h2>", unsafe_allow_html=True)

                chart_app = df_summary[df_summary['Aplikasi'] != 'TOTAL'].copy()
                chart_app = chart_app.melt(id_vars='Aplikasi', value_vars=['Sudah Teradministrasi', 'Belum Teradministrasi'], var_name='Status', value_name='Jumlah')
                fig_app = px.bar(chart_app, x='Aplikasi', y='Jumlah', color='Status', barmode='group', text='Jumlah')
                fig_app.update_layout(height=500, xaxis_title="", yaxis_title="Jumlah", legend=dict(title="Status", orientation="h", yanchor="top", y=-0.25, xanchor="center", x=0.5))

                st.plotly_chart(fig_app, use_container_width=True)

            with right_chart:
                st.markdown("<h2 style='text-align: center;'>Based on PIC</h2>", unsafe_allow_html=True)

                if not df_summary_pic.empty:
                    chart_pic = (df_summary_pic.melt(id_vars='PIC', value_vars=['Sudah Teradministrasi', 'Belum Teradministrasi'], var_name='Status', value_name='Jumlah'))
                    fig_pic = px.bar(chart_pic, x='PIC', y='Jumlah', color='Status', barmode='group', text='Jumlah')
                    fig_pic.update_layout(height=500, xaxis_tickangle=-45, xaxis_title="", yaxis_title="Jumlah", legend=dict(title="Status", orientation="h", yanchor="top", y=-0.55, xanchor="center", x=0.5))
                    st.plotly_chart(fig_pic, use_container_width=True)
                else:
                    st.warning("Data PIC tidak ditemukan.")

    else:
        st.info("Silakan upload ketiga file terlebih dahulu.")

# =========================================================
# DATABASE
# =========================================================

if menu == "Database":

    reqdb_file = st.sidebar.file_uploader("Upload Req DB", type=["xlsx"])

    if reqdb_file:
        try:
            with st.spinner("Memproses database..."):
                excel = pd.ExcelFile(reqdb_file)
                all_data = []

                for sheet in excel.sheet_names:
                    raw_df = read_excel_cached(reqdb_file, header=None, sheet_name=sheet)
                    header_row = None

                    for i in range(len(raw_df)):
                        row_values = (raw_df.iloc[i].astype(str).str.lower().tolist())

                        if any("masa berlaku dari" in str(v) for v in row_values):
                            header_row = i
                            break

                    if header_row is None:
                        continue

                    temp_df = read_excel_cached(reqdb_file, header=header_row, sheet_name=sheet)
                    temp_df = clean_col(temp_df)
                    temp_df["BULAN"] = sheet
                    all_data.append(temp_df)

                if len(all_data) == 0:
                    st.error("Tidak ada data valid")
                    st.stop()

                db_df = pd.concat(all_data, ignore_index=True)

                masa_dari = find_column(db_df, ["masa berlaku dari"])
                masa_sd = find_column(db_df, ["masa berlaku s.d", "masa berlaku sd"])
                ket_col = find_column(db_df, ["keterangan"])

                db_df[masa_dari] = pd.to_datetime(db_df[masa_dari], errors="coerce")
                db_df[masa_sd] = pd.to_datetime(db_df[masa_sd], errors="coerce")

                today = pd.Timestamp.today()

                def map_status(row):
                    if ket_col:
                        ket = str(row[ket_col]).lower()

                        if "resign" in str(ket):
                            return "Penutupan Akses"

                    if (pd.isna(row[masa_dari]) or pd.isna(row[masa_sd])):
                        return "Tiket Not Found"

                    if row[masa_sd] < today:
                        return "Expired"
                    return "Aktif"

                db_df["STATUS"] = (db_df.apply(map_status, axis=1))

                summary = (db_df.groupby(["BULAN", "STATUS"]).size().reset_index(name="JUMLAH"))

                fig = px.bar(summary, x="BULAN", y="JUMLAH", color="STATUS", barmode="stack", text="JUMLAH", title="Review Database")
                st.plotly_chart(fig, use_container_width=True)
                st.dataframe(db_df, use_container_width=True, height=600)

        except Exception as e:
            st.error(f"Gagal memproses database: {e}")

# =========================================================
# USER TIDAK AKTIF
# =========================================================

if menu == "User Tidak Aktif":
    st.sidebar.header("Upload File")

    bulan1_file = st.sidebar.file_uploader("Upload Bulan 1", type=["xlsx"], key="b1")
    bulan2_file = st.sidebar.file_uploader("Upload Bulan 2", type=["xlsx"], key="b2")
    bulan3_file = st.sidebar.file_uploader("Upload Bulan 3", type=["xlsx"], key="b3")
    users_file = st.sidebar.file_uploader("Upload Export User", type=["xlsx"], key="users")
    mdm_file = st.sidebar.file_uploader("Upload Export Pegawai MDM", type=["xlsx"], key="mdm")

    st.sidebar.divider()

    st.sidebar.subheader("Review Sebelumnya")

    prev_bulan1_file = st.sidebar.file_uploader("Prev Bulan 1", type=["xlsx"], key="pb1")
    prev_bulan2_file = st.sidebar.file_uploader("Prev Bulan 2", type=["xlsx"], key="pb2")
    prev_bulan3_file = st.sidebar.file_uploader("Prev Bulan 3", type=["xlsx"], key="pb3")
    prev_users_file = st.sidebar.file_uploader("Prev Export User", type=["xlsx"], key="pusers")
    prev_mdm_file = st.sidebar.file_uploader("Prev Export Pegawai MDM", type=["xlsx"], key="pmdm")

    required_files = [bulan1_file, bulan2_file, bulan3_file, users_file, mdm_file]

    if all(required_files):
        try:
            with st.spinner("Memproses data besar..."):
                bulan1 = clean_col(read_excel_cached(bulan1_file))
                bulan2 = clean_col(read_excel_cached(bulan2_file))
                bulan3 = clean_col(read_excel_cached(bulan3_file))
                users = clean_col(read_excel_cached(users_file, header=2))
                mdm = clean_col(read_excel_cached(mdm_file, header=2))

                bulan1 = bulan1[["NIPP", "AKTIVITAS", "NAMA"]]
                bulan2 = bulan2[["NIPP", "AKTIVITAS", "NAMA"]]
                bulan3 = bulan3[["NIPP", "AKTIVITAS", "NAMA"]]
                users = users[["NIPP", "STATUS"]]
                mdm = mdm[["PNALT", "PBTXT"]]

                bulan1["NIPP"] = normalize_nipp(bulan1["NIPP"])
                bulan2["NIPP"] = normalize_nipp(bulan2["NIPP"])
                bulan3["NIPP"] = normalize_nipp(bulan3["NIPP"])
                users["NIPP"] = normalize_nipp(users["NIPP"])
                mdm["PNALT"] = normalize_nipp(mdm["PNALT"])

                login1 = filter_login(bulan1)
                login2 = filter_login(bulan2)
                login3 = filter_login(bulan3)

                users["bulan1_tidak_login"] = (~users["NIPP"].map(login1.__contains__)).astype("int8")
                users["bulan2_tidak_login"] = (~users["NIPP"].map(login2.__contains__)).astype("int8")
                users["bulan3_tidak_login"] = (~users["NIPP"].map(login3.__contains__)).astype("int8")
                users["total_tidak_login"] = (users[["bulan1_tidak_login", "bulan2_tidak_login", "bulan3_tidak_login"]].sum(axis=1))

                tidak_aktif = users[(users["total_tidak_login"] == 3) & (users["STATUS"].astype(str).str.upper() == "A")].copy()

                nama_df = pd.concat([bulan1[["NIPP", "NAMA"]], bulan2[["NIPP", "NAMA"]], bulan3[["NIPP", "NAMA"]]])
                nama_df = (nama_df.drop_duplicates("NIPP"))

                tidak_aktif = tidak_aktif.merge(nama_df, on="NIPP", how="left")

                mdm = (mdm.drop_duplicates("PNALT"))

                tidak_aktif = tidak_aktif.merge(mdm, left_on="NIPP", right_on="PNALT", how="left")
                tidak_aktif["PBTXT"] = (tidak_aktif["PBTXT"].fillna("Tidak Ada Wilayah"))

                final_table = tidak_aktif[["NIPP", "NAMA", "PBTXT", "bulan1_tidak_login", "bulan2_tidak_login", "bulan3_tidak_login", "total_tidak_login"]].rename(columns={"PBTXT": "LOKASI"})

                total_user = (users["NIPP"].nunique())
                total_user_aktif = users[users["STATUS"].astype(str).str.upper() == "A"]["NIPP"].nunique()
                total_tidak_aktif = (final_table["NIPP"].nunique())

                # =====================================================
                # REAKTIVASI
                # =====================================================

                reaktivasi_user = 0
                reaktivasi_df = pd.DataFrame()

                prev_ready = all([prev_bulan1_file, prev_bulan2_file, prev_bulan3_file, prev_users_file, prev_mdm_file])

                if prev_ready:
                    prev_b1 = clean_col(read_excel_cached(prev_bulan1_file))
                    prev_b2 = clean_col(read_excel_cached(prev_bulan2_file))
                    prev_b3 = clean_col(read_excel_cached(prev_bulan3_file))
                    prev_users = clean_col(read_excel_cached(prev_users_file, header=2))
                    prev_mdm = clean_col(read_excel_cached(prev_mdm_file))

                    prev_b1["NIPP"] = normalize_nipp(prev_b1["NIPP"])
                    prev_b2["NIPP"] = normalize_nipp(prev_b2["NIPP"])
                    prev_b3["NIPP"] = normalize_nipp(prev_b3["NIPP"])
                    prev_users["NIPP"] = normalize_nipp(prev_users["NIPP"])
                    prev_mdm["PNALT"] = normalize_nipp(prev_mdm["PNALT"])

                    prev_login = (filter_login(prev_b1) | filter_login(prev_b2) | filter_login(prev_b3))
                    prev_users["not_login"] = (~prev_users["NIPP"].map(prev_login.__contains__)).astype("int8")
                    prev_deaktif = prev_users[(prev_users["not_login"] == 1) & (prev_users["STATUS"].astype(str).str.upper() == "A")].copy()

                    current_login = (login1 | login2 | login3)

                    reaktivasi_df = prev_deaktif[prev_deaktif["NIPP"].isin(current_login)].copy()

                    prev_mdm = (prev_mdm[["PNALT", "PBTXT"]].drop_duplicates("PNALT"))

                    reaktivasi_df = reaktivasi_df.merge(prev_mdm, left_on="NIPP", right_on="PNALT", how="left")
                    reaktivasi_df["PBTXT"] = (reaktivasi_df["PBTXT"].fillna("Tidak Ada Wilayah"))
                    reaktivasi_user = (reaktivasi_df["NIPP"].nunique())

                # =====================================================
                # METRIC
                # =====================================================

                c1, c2, c3, c4 = st.columns(4)

                c1.metric("Total User", f"{total_user:,}")
                c2.metric("User Aktif", f"{total_user_aktif:,}")
                c3.metric("Tidak Aktif", f"{total_tidak_aktif:,}")
                c4.metric("Reaktivasi", f"{reaktivasi_user:,}")

                chart_df = pd.DataFrame({"Status": ["Aktif", "Tidak Aktif", "Reaktivasi"], "Jumlah": [total_user_aktif, total_tidak_aktif, reaktivasi_user]})
                fig = px.bar(chart_df, x="Status", y="Jumlah", color="Status", text="Jumlah")
                st.plotly_chart(fig, use_container_width=True)

                wilayah_df = (final_table.groupby("LOKASI").size().reset_index(name="JUMLAH").sort_values(by="JUMLAH",ascending=False))
                fig2 = px.bar(wilayah_df, x="LOKASI", y="JUMLAH", text="JUMLAH", title="Based on Location")
                fig2.update_traces(textposition="outside")
                fig2.update_layout(height=700, title_x=0.5, xaxis_tickangle=-45)

                st.plotly_chart(fig2, use_container_width=True)

                # =====================================================
                # TREND DEAKTIVASI & REAKTIVASI
                # =====================================================

                bulan_chart = ["Sep-25", "Okt-25", "Nov-25", "Des-25", "Jan-26", "Feb-26", "Mar-26", "Apr-26", "Mei-26"]
                deaktif_chart = [421, 4621, 1444, 4423, 1401, 2636, 1437, 2869, total_tidak_aktif]
                reaktif_chart = [174, 2247, 725, 3447, 897, 1985, 830, reaktivasi_user, 0]

                trend_deaktif = pd.DataFrame({"Bulan": bulan_chart, "Jumlah": deaktif_chart, "Status": "Deaktivasi"})
                trend_reaktif = pd.DataFrame({"Bulan": bulan_chart, "Jumlah": reaktif_chart, "Status": "Reaktivasi"})

                trend_df = pd.concat([trend_deaktif, trend_reaktif], ignore_index=True)
                trend_df = trend_df[trend_df["Jumlah"] != 0]

                fig_trend = px.line(trend_df, x="Bulan", y="Jumlah", color="Status", markers=True, text="Jumlah", title="Trend Deaktivasi dan Reaktivasi User")
                fig_trend.update_traces(textposition="top center", line=dict(width=3), marker=dict(size=8))
                fig_trend.update_layout(height=500, title_x=0.5, legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5))

                st.plotly_chart(fig_trend, use_container_width=True)

                # =====================================================
                # DONUT CHART
                # =====================================================
                
                total_distribusi = (total_user_aktif)

                persen_tidak_aktif = round((total_tidak_aktif / total_distribusi) * 100, 2)
                persen_aktif = round(100 - persen_tidak_aktif, 2)

                fig_donut = go.Figure(data=[go.Pie(labels=[f"User Aktif ({persen_aktif}%)", f"User Tidak Aktif ({persen_tidak_aktif}%)"], values=[total_user_aktif, total_tidak_aktif], hole=0.7, textinfo="label")])
                fig_donut.update_layout(height=500, title={"text": "Persentase User Aktif vs Tidak Aktif", "x": 0.5}, showlegend=True)

                st.plotly_chart(fig_donut, use_container_width=True)

        except Exception as e:
            st.error(f"Gagal memproses data: {e}")

# =========================================================
# FOOTER
# =========================================================

st.divider()
st.caption("Dashboard Review Akses")