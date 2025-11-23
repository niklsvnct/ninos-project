import streamlit as st
import pandas as pd
from datetime import time, datetime
import io
import time as t_sleep
import xlsxwriter
import streamlit.components.v1 as components

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="Nino's Project - Command Center",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 2. LINK DATABASE (PASTIKAN BENAR) ---
SHEET_URL_ABSEN = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ_GhoIb1riX98FsP8W4f2-_dH_PLcLDZskjNOyDcnnvOhBg8FUp3xJ-c_YgV0Pw71k4STy4rR0_MS5/pub?output=csv"
# 👇 GANTI INI DENGAN LINK SHEET 2 YANG BARU 👇
SHEET_URL_STATUS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ2QrBN8uTRiHINCEcZBrbdU-gzJ4pN2UljoqYG6NMoUQIK02yj_D1EdlxdPr82Pbr94v2o6V0Vh3Kt/pub?gid=511860805&single=true&output=csv"
GOOGLE_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSeopdaE-lyOtFd2TUr5C3K2DWE3Syt2PaKoXMp0cmWKIFnijw/viewform?usp=header" 

# --- 3. SETUP VARIABEL ---
COL_NAMA = 'Person Name'
COL_TIMESTAMP = 'Event Time'
LATE_THRESHOLD = time(7, 5, 0) 
RENTANG_WAKTU = {'Pagi': ('03:00:00', '11:00:00'), 'Siang_1': ('11:29:00', '12:30:59'), 'Siang_2': ('12:31:00', '13:30:59'), 'Sore': ('17:00:00', '23:59:59')}

DATA_DIVISI = {
    "LEADERSHIP": ["Patra Anggana", "Su Adam", "Budiman Arifin", "Rifaldy Ilham Bhagaskara", "Marwan S Halid", "Budiono"],
    "TLB": ["M. Ansori", "Bayu Pratama Putra Katuwu", "Yoga Nugraha Putra Pasaribu", "Junaidi Taib", "Muhammad Rizal Amra", "Rusli Dj"],
    "TBL": ["Venesia Aprilia Ineke", "Muhammad Naufal Ramadhan", "Yuzak Gerson Puturuhu", "Muhamad Alief Wildan", "Gafur Hamisi", "Jul Akbar M. Nur", "Sarni Massiri", "Adrianto Laundang", "Wahyudi Ismail"],
    "TRANS APRON": ["Marichi Gita Rusdi", "Ilham Rahim", "Abdul Mu Iz Simal", "Dwiki Agus Saputro", "Moh. Sofyan", "Faisal M. Kadir", "Amirudin Rustam", "Faturrahman Kaunar", "Wawan Hermawan", "Rahmat Joni", "Nur Ichsan"],
    "ATS": ["Nurultanti", "Firlon Paembong", "Irwan Rezky Setiawan", "Yusuf Arviansyah", "Nurdahlia Is. Folaimam", "Ghaly Rabbani Panji Indra", "Ikhsan Wahyu Vebriyan", "Rizki Mahardhika Ardi Tigo", "Nikolaus Vincent Quirino"],
    "ADM COMPLIANCE": ["Yessicha Aprilyona Siregar", "Gabriela Margrith Louisa Klavert", "Aldi Saptono"],
    "TRANSLATOR": ["Wilyam Candra", "Norika Joselyn Modnissa"],
    "AVSEC": ["Andrian Maranatha", "Toni Nugroho Simarmata", "Muhamad Albi Ferano", "Andreas Charol Tandjung", "Sabadia Mahmud", "Rusdin Malagapi", "Muhamad Judhytia Winli", "Wahyu Samsudin", "Fientje Elisabeth Joseph", "Anglie Fitria Desiana Mamengko", "Dwi Purnama Bimasakti", "Windi Angriani Sulaeman", "Megawati A. Rauf"],
    "GROUND HANDLING": ["Yuda Saputra.", "Tesalonika Gratia Putri Toar", "Esi Setia Ningseh", "Ardiyanto Kalatjo", "Febrianti Tikabala"],
    "HELICOPTER": ["Agung Sabar Santoso Taufik", "Recky Irwan R. A Arsyad", "Farok Abdul", "Achmad Rizky Ariz", "Yus Andi", "Muh. Noval Kipudjena"],
    "AMC & TERMINAL": ["Risky Sulung", "Muchamad Nur Syaifulrahman", "Muhammad Tunjung Rohmatullah", "Sunarty Fakir", "Albert Papuling", "Gibhran Fitransyah Yusri", "Muhdi R Tomia", "Riski Rifaldo Theofilus Anu", "Eko"],
    "SAFETY OFFICER": ["Hildan Ahmad Zaelani", "Abdurahim Andar"],
    "PKP-PK": ["Andreas Aritonang", "Achmad Alwan Asyhab", "Doni Eka", "Yogi Prasetya Eka Winandra", "Akhsin Aditya Weza Putra", "Fardhan Ahmad Tajali", "Maikel R", "Saldi Sandra", "Hamzah M Ali Gani", "Marfan Mandar", "Julham Keya", "Aditya Sugiantoro A", "M.Usman", "M.Akbar Patty", "Daniel Freski W", "Fandi M Naser", "Agung F", "Deni Hendri", "M Rifai", "Idrus Arsad"]
}

URUTAN_NAMA_CUSTOM = []
for unit, members in DATA_DIVISI.items():
    URUTAN_NAMA_CUSTOM.extend(members)

# --- 4. CSS STYLE (FIXED) ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;500;800&family=JetBrains+Mono:wght@400;700&display=swap');
    .stApp { background-color: #050505; background-image: radial-gradient(at 50% 0%, hsla(225,39%,20%,1) 0, transparent 50%); color: white; }
    section[data-testid="stSidebar"] { background-color: #0a0a0a; border-right: 1px solid #333; }
    .brand-title { font-family: 'Outfit', sans-serif; font-size: 3rem; font-weight: 800; background: linear-gradient(to right, #00c6ff, #0072ff); -webkit-background-clip: text; -webkit-text-fill-color: transparent; margin-bottom: 0; }
    .brand-subtitle { font-family: 'Outfit', sans-serif; color: #888; font-size: 1rem; letter-spacing: 3px; text-transform: uppercase; margin-bottom: 30px; border-bottom: 1px solid #333; padding-bottom: 20px; }
    div[data-testid="stMetric"] { background: rgba(255, 255, 255, 0.05); border: 1px solid rgba(255, 255, 255, 0.1); border-radius: 12px; padding: 15px; text-align: center; }
    div[data-testid="stMetricLabel"] { font-size: 0.9rem; color: #aaa; }
    div[data-testid="stMetricValue"] { font-size: 2rem; font-weight: 800; color: white; }
    
    /* CARD STYLES */
    .card { border-radius: 16px; padding: 20px; margin-bottom: 20px; border: 1px solid rgba(255,255,255,0.05); background: #1a1a1a; transition: transform 0.3s; }
    .card:hover { transform: translateY(-5px); }
    .card-present { border-top: 3px solid #00f260; } 
    .card-partial { border-top: 3px solid #FFC837; } 
    .card-absent { border-top: 3px solid #FF416C; } 
    .card-permit { border-top: 3px solid #d580ff; }
    
    .card-name { font-family: 'Outfit', sans-serif; font-weight: 700; font-size: 0.95rem; margin: 0; color: white; }
    .detail-row { display: flex; justify-content: space-between; margin-bottom: 5px; font-size: 0.8rem; color: #ccc; }
    .value { font-family: 'JetBrains Mono', monospace; font-weight: 600; }
    .late-indicator { color: #ff4b4b; font-weight: bold; font-size: 0.8rem; margin-left: 5px; }
    
    .stDownloadButton button { background: linear-gradient(90deg, #00c6ff, #0072ff) !important; color: white !important; font-weight: 800 !important; border: none !important; width: 100%; }
    .stTextInput input { background: #18181b !important; border: 1px solid #27272a !important; color: white !important; }
</style>
""", unsafe_allow_html=True)

# --- 5. FUNGSI LOGIKA ---
def get_min_time_in_range(group, s, e):
    st_t = time.fromisoformat(s); end_t = time.fromisoformat(e)
    filtered = group[(group['Waktu'] >= st_t) & (group['Waktu'] <= end_t)]
    return filtered[COL_TIMESTAMP].min().strftime('%H:%M') if not filtered.empty else None

def is_late(time_str):
    if not time_str: return False
    try: return datetime.strptime(time_str, '%H:%M').time() > LATE_THRESHOLD
    except: return False

@st.cache_data(ttl=5) 
def load_absen(url):
    try:
        df = pd.read_csv(url); df.columns = df.columns.str.strip()
        if COL_NAMA not in df.columns: return None
        df[COL_NAMA] = df[COL_NAMA].astype(str).str.strip(); df = df[df[COL_NAMA] != ''].copy()
        df[COL_TIMESTAMP] = pd.to_datetime(df[COL_TIMESTAMP])
        df['Tanggal'] = df[COL_TIMESTAMP].dt.date; df['Waktu'] = df[COL_TIMESTAMP].dt.time
        return df
    except: return None

@st.cache_data(ttl=5)
def load_status(url):
    try:
        df = pd.read_csv(url)
        df = df.rename(columns=lambda x: x.strip())
        df['Nama Karyawan'] = df['Nama Karyawan'].astype(str).str.strip()
        
        # JURUS ANTI-GAGAL TANGGAL:
        # Kita buat kolom Tanggal_Str yang formatnya String YYYY-MM-DD
        # Jadi tidak peduli inputnya miring/lurus, kita samakan format teksnya
        df['Tanggal'] = pd.to_datetime(df['Tanggal'], dayfirst=False, errors='coerce') # Coba M/D/Y dulu
        df['Tanggal_Str'] = df['Tanggal'].dt.strftime('%Y-%m-%d')
        
        df['Keterangan'] = df['Keterangan'].astype(str).str.strip().str.upper()
        return df
    except Exception as e:
        return None

# ==========================================
# --- 6. APLIKASI UTAMA ---
# ==========================================

with st.sidebar:
    st.markdown("## ⚡ WedaBayAirport")
    menu = st.radio("Pilih Menu:", ["📊 Dashboard Monitoring", "📝 Input Laporan (Koordinator)"])
    st.markdown("---")
    st.info("Gunakan menu 'Input Laporan' untuk melaporkan Sakit/Izin/Cuti.")

# --- DASHBOARD ---
if menu == "📊 Dashboard Monitoring":
    st.markdown('<div class="brand-title">NINO\'S PROJECT</div>', unsafe_allow_html=True)
    st.markdown('<div class="brand-subtitle">CENTRALIZED COMMAND CENTER</div>', unsafe_allow_html=True)

    df_absen = load_absen(SHEET_URL_ABSEN)
    df_status = load_status(SHEET_URL_STATUS)

    if df_absen is None:
        st.warning("⚠️ Data Absensi Mesin belum terbaca / kosong.")
        df_absen = pd.DataFrame(columns=[COL_NAMA, COL_TIMESTAMP, 'Tanggal', 'Waktu'])

    c1, c2 = st.columns([1, 3])
    with c1:
        avail_dates = sorted(df_absen['Tanggal'].unique())
        if not avail_dates: sel_date = st.date_input("📅 Pilih Tanggal", value=None)
        else: sel_date = st.date_input("📅 Pilih Tanggal", value=avail_dates[-1])
    with c2:
        search_q = st.text_input("🔍 Search Employee", placeholder="Ketik nama...")

    st.markdown("---")

    if sel_date:
        df_today = df_absen[df_absen['Tanggal'] == sel_date]
        status_today = {}
        
        # LOAD STATUS (PAKAI FORMAT STRING BIAR PASTI COCOK)
        if df_status is not None:
            sel_date_str = sel_date.strftime('%Y-%m-%d')
            df_stat_today = df_status[df_status['Tanggal_Str'] == sel_date_str]
            if not df_stat_today.empty:
                status_today = pd.Series(df_stat_today.Keterangan.values, index=df_stat_today['Nama Karyawan']).to_dict()

        # PROSES DATA MESIN
        if df_today.empty: df_res = pd.DataFrame(columns=[COL_NAMA]) 
        else:
            recap_dict = {}; grouped = df_today.groupby([COL_NAMA, 'Tanggal'])
            for cat, (s, e) in RENTANG_WAKTU.items():
                recap_dict[cat] = grouped.apply(lambda x: get_min_time_in_range(x, s, e))
            df_res = pd.DataFrame(recap_dict).reset_index()
            if not df_res.empty: df_res.rename(columns={COL_NAMA: 'Nama Karyawan'}, inplace=True)
        
        df_final = pd.merge(pd.DataFrame({'Nama Karyawan': URUTAN_NAMA_CUSTOM}), df_res, on='Nama Karyawan', how='left')
        for col in list(RENTANG_WAKTU.keys()):
            if col not in df_final.columns: df_final[col] = ''
        df_final[list(RENTANG_WAKTU.keys())] = df_final[list(RENTANG_WAKTU.keys())].fillna('')

        # METRIK
        total_emp = len(df_final); on_time_count = 0; late_count = 0; izin_count = 0; bolos_count = 0
        list_terlambat = []; list_izin = []; list_bolos = []
        
        for _, row in df_final.iterrows():
            nm = row['Nama Karyawan']; pagi = row['Pagi']; times = [row['Pagi'], row['Siang_1'], row['Siang_2'], row['Sore']]; empty = sum(1 for t in times if t == '')
            manual_stat = status_today.get(nm, "")
            
            if manual_stat:
                izin_count += 1; list_izin.append((nm, manual_stat))
            elif empty == 4:
                bolos_count += 1; list_bolos.append(nm)
            else:
                if pagi and is_late(pagi):
                    late_count += 1; list_terlambat.append((nm, pagi))
                else: on_time_count += 1

        hadir_total = on_time_count + late_count
        present_rate = round((hadir_total / total_emp) * 100, 1) if total_emp > 0 else 0

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("TOTAL KARYAWAN", total_emp, "Orang")
        m2.metric("HADIR HARI INI", hadir_total, f"{present_rate}% Rate")
        m3.metric("IZIN / SAKIT", izin_count, "Orang")
        m4.metric("TANPA KETERANGAN", bolos_count, "Orang", delta_color="inverse")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # DAFTAR LIST
        col_telat, col_izin, col_bolos = st.columns(3)
        with col_telat:
            st.markdown("#### ⏰ TERLAMBAT")
            if list_terlambat:
                with st.expander(f"Lihat {len(list_terlambat)} Orang", expanded=False):
                    for nama, jam in list_terlambat: st.markdown(f"<div class='anomaly-box box-telat'><span class='anomaly-name'>{nama}</span><span class='anomaly-info' style='color:#ff4b4b'>{jam}</span></div>", unsafe_allow_html=True)
            else: st.info("Semua Tepat Waktu!")
        with col_izin:
            st.markdown("#### ℹ️ IZIN / SAKIT")
            if list_izin:
                with st.expander(f"Lihat {len(list_izin)} Orang", expanded=False):
                    for nama, alasan in list_izin: st.markdown(f"<div class='anomaly-box box-izin'><span class='anomaly-name'>{nama}</span><span class='anomaly-info' style='color:#d580ff'>{alasan}</span></div>", unsafe_allow_html=True)
            else: st.info("Nihil")
        with col_bolos:
            st.markdown("#### ❌ ALPHA")
            if list_bolos:
                with st.expander(f"Lihat {len(list_bolos)} Orang", expanded=False):
                    for nama in list_bolos: st.markdown(f"<div class='anomaly-box box-alpha'><span class='anomaly-name'>{nama}</span><span class='anomaly-info' style='color:#FFFF00'>--:--</span></div>", unsafe_allow_html=True)
            else: st.success("Nihil")

        st.markdown("---")
        
        # DOWNLOAD
        out = io.BytesIO()
        wb = xlsxwriter.Workbook(out, {'in_memory': True})
        ws = wb.add_worksheet('Rekap')
        fmt_head = wb.add_format({'bold': True, 'fg_color': '#4caf50', 'font_color': 'white', 'border': 1, 'align': 'center'})
        fmt_norm = wb.add_format({'border': 1, 'align': 'center'})
        fmt_miss = wb.add_format({'bg_color': '#FF0000', 'border': 1}) 
        fmt_full = wb.add_format({'bg_color': '#FFFF00', 'border': 1})
        fmt_late = wb.add_format({'font_color': 'red', 'bold': True, 'border': 1, 'align': 'center'})
        headers = ['Nama Karyawan', 'Pagi', 'Siang_1', 'Siang_2', 'Sore', 'Keterangan']
        ws.write_row(0, 0, headers, fmt_head); ws.set_column(0, 0, 30); ws.set_column(1, 5, 15)
        for idx, row in df_final.iterrows():
            nm = row['Nama Karyawan']; pagi = row['Pagi']; manual_stat = status_today.get(nm, "")
            ws.write(idx+1, 0, nm, fmt_norm); ws.write(idx+1, 5, manual_stat, fmt_norm)
            times = [row['Pagi'], row['Siang_1'], row['Siang_2'], row['Sore']]; empty = sum(1 for t in times if t == '')
            if manual_stat:
                for i in range(4): ws.write(idx+1, i+1, "", fmt_norm)
            elif empty == 4:
                for i in range(4): ws.write(idx+1, i+1, "", fmt_full)
            else:
                if pagi == '': ws.write(idx+1, 1, "", fmt_miss)
                else:
                    if is_late(pagi): ws.write(idx+1, 1, pagi, fmt_late)
                    else: ws.write(idx+1, 1, pagi, fmt_norm)
                rest_times = [row['Siang_1'], row['Siang_2'], row['Sore']]
                for i, t in enumerate(rest_times):
                    col_idx = i + 2
                    if t == '': ws.write(idx+1, col_idx, "", fmt_miss)
                    else: ws.write(idx+1, col_idx, t, fmt_norm)
        wb.close(); out.seek(0)
        st.download_button("📥 DOWNLOAD EXCEL REPORT", out, f"Rekap_Smart_{sel_date}.xlsx", use_container_width=True)
        
        st.markdown("<br>", unsafe_allow_html=True)

        # CARD GRID
        with st.expander("🔽 Tampilan Semua detail absen karyawan", expanded=False):
            st.markdown("### 📁 DETAIL PER UNIT")
            for unit_name, member_list in DATA_DIVISI.items():
                member_count = len(member_list)
                df_unit = df_final[df_final['Nama Karyawan'].isin(member_list)]
                if search_q: df_unit = df_unit[df_unit['Nama Karyawan'].str.contains(search_q, case=False, na=False)]

                with st.expander(f"📂 {unit_name} ({len(df_unit)}/{member_count} Personil)", expanded=False):
                    if df_unit.empty: st.info(f"Tidak ada data karyawan di unit {unit_name}")
                    else:
                        COLS = 4; rows = [df_unit.iloc[i:i+COLS] for i in range(0, len(df_unit), COLS)]
                        for r in rows:
                            cols = st.columns(COLS)
                            for i, (idx, row) in enumerate(r.iterrows()):
                                with cols[i]:
                                    nm = row['Nama Karyawan']; pagi = row['Pagi']; times = [pagi, row['Siang_1'], row['Siang_2'], row['Sore']]; empty = sum(1 for t in times if t == '')
                                    manual_stat = status_today.get(nm, "")
                                    if manual_stat: lbl = f"ℹ️ {manual_stat}"; theme = "card-permit"; clr = "#d580ff"
                                    elif empty == 4: lbl = "⛔ TOTAL ABSENT"; theme = "card-absent"; clr = "#FF416C"
                                    elif empty > 0: lbl = "⚠️ PARTIAL"; theme = "card-partial"; clr = "#FFC837"
                                    else: lbl = "✅ FULL PRESENT"; theme = "card-present"; clr = "#00f260"
                                    late_html = '<span class="late-indicator">⏰ LATE</span>' if (pagi and is_late(pagi)) else ""
                                    avt = f"https://ui-avatars.com/api/?name={nm.replace(' ', '+')}&background=random&color=fff"
                                    st.markdown(f"<div class='card {theme}'><div class='card-header'><img src='{avt}' class='avatar'><div><p class='card-name'>{nm}</p><p class='card-id'>NP-{100+idx}</p></div></div><div class='detail-row'><span class='label'>Datang</span><span><span class='value'>{pagi if pagi else '-'}</span> {late_html}</span></div><div class='detail-row'><span class='label'>Pulang</span><span class='value'>{row['Sore'] if row['Sore'] else '-'}</span></div><div style='text-align:right; font-size:0.7rem; font-weight:bold; color:{clr}; margin-top:10px;'>{lbl}</div></div>", unsafe_allow_html=True)
                                    with st.popover("Detail", use_container_width=True):
                                        if manual_stat: st.info(f"Status: {manual_stat}")
                                        c1, c2 = st.columns(2)
                                        p_val = pagi if pagi else "❌"; 
                                        if is_late(pagi): p_val += " (Telat)"
                                        c1.metric("Pagi", p_val); c1.metric("Siang 1", row['Siang_1'] if row['Siang_1'] else "❌")
                                        c2.metric("Siang 2", row['Siang_2'] if row['Siang_2'] else "❌"); c2.metric("Sore", row['Sore'] if row['Sore'] else "❌")

    else: st.info("Menghubungkan Database...")
    
    # --- AREA DEBUGGING SUPER (PASTI KETAHUAN SALAHNYA DIMANA) ---
    st.divider()
    with st.expander("🛠️ DEBUGGING DATA MENTAH (CEK DISINI)", expanded=True):
        st.write("### Data Mentah dari Sheet 2 (Google Form):")
        if df_status is not None:
            st.dataframe(df_status)
            st.write(f"🔍 Website mencari tanggal: **{sel_date}** (Format YYYY-MM-DD)")
            
            # Cek apakah ada data yang cocok
            cocok = df_status[df_status['Tanggal_Str'] == str(sel_date)]
            if cocok.empty:
                st.error(f"❌ TIDAK ADA data yang cocok dengan tanggal {sel_date}.")
                st.info("Tips: Cek kolom 'Tanggal_Str' di tabel atas. Apakah formatnya sudah benar YYYY-MM-DD?")
            else:
                st.success(f"✅ DITEMUKAN {len(cocok)} data cocok!")
                st.dataframe(cocok)
        else:
            st.error("Gagal membaca Sheet 2. Cek Link!")

elif menu == "📝 Input Laporan (Koordinator)":
    st.markdown('<div class="brand-title">INPUT LAPORAN</div>', unsafe_allow_html=True)
    st.markdown('<div class="brand-subtitle">FORMULIR MANUAL KOORDINATOR</div>', unsafe_allow_html=True)
    if "PASTE_LINK" in GOOGLE_FORM_URL: st.warning("⚠️ Link Google Form belum dimasukkan.")
    else:
        st.info("Silakan isi formulir di bawah ini.")
        components.iframe(GOOGLE_FORM_URL, height=1200, scrolling=True)