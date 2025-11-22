import streamlit as st
import pandas as pd
from datetime import time, datetime
import io
import time as t_sleep
import xlsxwriter

# --- KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="Nino's Project - Intelligent",
    page_icon="🧠",
    layout="wide"
)

# ==========================================
# 👇 SETTING BATAS TERLAMBAT DISINI 👇
# ==========================================
LATE_THRESHOLD = time(7, 5, 0) # Jam 07:05:00 (Lewat ini dianggap telat)

# 1. LINK DATA ABSEN (MESIN FINGER)
SHEET_URL_ABSEN = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ_GhoIb1riX98FsP8W4f2-_dH_PLcLDZskjNOyDcnnvOhBg8FUp3xJ-c_YgV0Pw71k4STy4rR0_MS5/pub?output=csv"

# 2. LINK DATA KETERANGAN (GOOGLE FORM)
SHEET_URL_STATUS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ2QrBN8uTRiHINCEcZBrbdU-gzJ4pN2UljoqYG6NMoUQIK02yj_D1EdlxdPr82Pbr94v2o6V0Vh3Kt/pub?output=csv"

# ==========================================

# --- CSS: TAMPILAN KEREN ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;500;700&family=JetBrains+Mono:wght@400;700&display=swap');

    .stApp {
        background-color: #050505;
        background-image: radial-gradient(at 50% 0%, hsla(225,39%,30%,1) 0, transparent 50%);
        color: white;
    }

    /* CARD STYLES */
    .card {
        border-radius: 16px;
        padding: 20px;
        margin-bottom: 20px;
        border: 1px solid rgba(255,255,255,0.05);
        background: #1a1a1a;
        transition: transform 0.3s;
    }
    .card:hover { transform: translateY(-5px); }

    /* WARNA STATUS */
    .card-present { border-top: 3px solid #00f260; }
    .card-partial { border-top: 3px solid #FFC837; }
    .card-absent { border-top: 3px solid #FF416C; }
    .card-permit { border-top: 3px solid #d580ff; }

    /* TEXT */
    .card-name { font-family: 'Outfit', sans-serif; font-weight: 700; font-size: 0.95rem; margin: 0; }
    .detail-row { display: flex; justify-content: space-between; margin-bottom: 5px; font-size: 0.8rem; color: #ccc; }
    .value { font-family: 'JetBrains Mono', monospace; font-weight: 600; }
    
    /* INDIKATOR TELAT DI WEB */
    .late-indicator { color: #ff4b4b; font-weight: bold; font-size: 0.8rem; margin-left: 5px; }

    /* BUTTONS */
    .stDownloadButton button { background: linear-gradient(90deg, #00c6ff, #0072ff) !important; color: white !important; font-weight: 800 !important; border: none !important; }
</style>
""", unsafe_allow_html=True)

st.markdown('<h1 style="text-align: center; background: linear-gradient(to right, #00dbde, #fc00ff); -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-family: Outfit;">NINO\'S PROJECT</h1>', unsafe_allow_html=True)
st.markdown('<p style="text-align: center; color: #aaa; letter-spacing: 2px;">INTELLIGENT ATTENDANCE SYSTEM</p>', unsafe_allow_html=True)

# --- SETUP DATA ---
COL_NAMA = 'Person Name'
COL_TIMESTAMP = 'Event Time'

RENTANG_WAKTU = {
    'Pagi': ('05:00:00', '09:00:00'),
    'Siang_1': ('11:00:00', '12:30:59'),
    'Siang_2': ('12:31:00', '13:30:59'),
    'Sore': ('16:00:00', '23:59:59')
}

URUTAN_NAMA_CUSTOM = [
    "Patra Anggana", "Su Adam", "Budiman Arifin", "Rifaldy Ilham Bhagaskara", "Marwan S Halid", 
    "Budiono", "M. Ansori", "Bayu Pratama Putra Katuwu", "Yoga Nugraha Putra Pasaribu", 
    "Junaidi Taib", "Muhammad Rizal Amra", "Rusli Dj", "Venesia Aprilia Ineke", 
    "Muhammad Naufal Ramadhan", "Yuzak Gerson Puturuhu", "Muhamad Alief Wildan", 
    "Gafur Hamisi", "Jul Akbar M. Nur", "Sarni Massiri", "Adrianto Laundang", 
    "Wahyudi Ismail", "Marichi Gita Rusdi", "Ilham Rahim", "Abdul Mu Iz Simal", 
    "Dwiki Agus Saputro", "Moh. Sofyan", "Faisal M. Kadir", "Amirudin Rustam", 
    "Faturrahman Kaunar", "Wawan Hermawan", "Rahmat Joni", "Nur Ichsan", 
    "Nurultanti", "Firlon Paembong", "Irwan Rezky Setiawan", "Yusuf Arviansyah", 
    "Nurdahlia Is. Folaimam", "Ghaly Rabbani Panji Indra", "Ikhsan Wahyu Vebriyan", 
    "Rizki Mahardhika Ardi Tigo", "Nikolaus Vincent Quirino", "Yessicha Aprilyona Siregar", 
    "Gabriela Margrith Louisa Klavert", "Aldi Saptono", "Wilyam Candra", 
    "Norika Joselyn Modnissa", "Andrian Maranatha", "Toni Nugroho Simarmata", 
    "Muhamad Albi Ferano", "Andreas Charol Tandjung", "Sabadia Mahmud", "Rusdin Malagapi", 
    "Muhamad Judhytia Winli", "Wahyu Samsudin", "Fientje Elisabeth Joseph", 
    "Anglie Fitria Desiana Mamengko", "Dwi Purnama Bimasakti", "Windi Angriani Sulaeman", 
    "Megawati A. Rauf", "Yuda Saputra", "Tesalonika Gratia Putri Toar", "Esi Setia Ningseh", 
    "Ardiyanto Kalatjo", "Febrianti Tikabala", "Agung Sabar S. Taufik", 
    "Recky Irwan R. A Arsyad", "Farok Abdul", "Achmad Rizky Ariz", "Yus Andi", 
    "Muh. Noval Kipudjena", "Risky Sulung", "Muchamad Nur Syaifulrahman", 
    "Muhammad Tunjung Rohmatullah", "Sunarty Fakir", "Albert Papuling", 
    "Gibhran Fitransyah Yusri", "Muhdi R Tomia", "Riski Rifaldo Theofilus Anu", 
    "Eko", "Hildan Ahmad Zaelani", "Abdurahim Andar", "Andreas Aritonang", 
    "Achmad Alwan Asyhab", "Doni Eka Satria", "Yogi Prasetya Eka Winandra", 
    "Akhsin Aditya Weza Putra", "Fardhan Ahmad Tajali", "Maikel Renato Syafaruddin", 
    "Saldi Sandra", "Hamzah M. Ali Gani", "Marfan Mandar", "Julham Keya", 
    "Aditya Sugiantoro Abbas", "Muhamad Usman", "M Akbar D Patty", "Daniel Freski Wangka", 
    "Fandi M.Naser", "Agung Fadjriansyah Ano", "Deni Hendri Bobode", "Muhammad Rifai", 
    "Idrus Arsad, SH"
]

def get_min_time_in_range(group, start_time_str, end_time_str):
    start_t = time.fromisoformat(start_time_str)
    end_t = time.fromisoformat(end_time_str)
    filtered = group[(group['Waktu'] >= start_t) & (group['Waktu'] <= end_t)]
    if not filtered.empty:
        return filtered[COL_TIMESTAMP].min().strftime('%H:%M')
    return None 

# --- FUNGSI CEK TELAT ---
def is_late(time_str):
    if not time_str or time_str == '': return False
    try:
        t = datetime.strptime(time_str, '%H:%M').time()
        return t > LATE_THRESHOLD
    except: return False

@st.cache_data(ttl=30) 
def load_absen(url):
    try:
        df = pd.read_csv(url)
        df.columns = df.columns.str.strip()
        df[COL_NAMA] = df[COL_NAMA].astype(str).str.strip()
        df = df[df[COL_NAMA] != ''].copy()
        df[COL_TIMESTAMP] = pd.to_datetime(df[COL_TIMESTAMP])
        df['Tanggal'] = df[COL_TIMESTAMP].dt.date
        df['Waktu'] = df[COL_TIMESTAMP].dt.time
        return df
    except: return None

@st.cache_data(ttl=30)
def load_status(url):
    try:
        df = pd.read_csv(url)
        df = df.rename(columns=lambda x: x.strip())
        df['Nama Karyawan'] = df['Nama Karyawan'].astype(str).str.strip()
        df['Tanggal'] = pd.to_datetime(df['Tanggal']).dt.date
        df['Keterangan'] = df['Keterangan'].astype(str).str.strip().str.upper()
        return df
    except: return None

# --- MAIN ---
df_absen = load_absen(SHEET_URL_ABSEN)
df_status = load_status(SHEET_URL_STATUS)

if df_absen is not None:
    c1, c2 = st.columns([1, 3])
    with c1:
        avail_dates = sorted(df_absen['Tanggal'].unique())
        sel_date = st.date_input("📅 Pilih Tanggal", value=avail_dates[-1] if avail_dates else None)
    with c2:
        search_q = st.text_input("🔍 Search Employee", placeholder="Ketik nama...")

    st.markdown("---")

    if sel_date:
        df_today = df_absen[df_absen['Tanggal'] == sel_date]
        
        status_today = {}
        if df_status is not None:
            df_stat_today = df_status[df_status['Tanggal'] == sel_date]
            if not df_stat_today.empty:
                status_today = pd.Series(df_stat_today.Keterangan.values, index=df_stat_today['Nama Karyawan']).to_dict()

        recap_dict = {}
        grouped = df_today.groupby([COL_NAMA, 'Tanggal'])
        for cat, (s, e) in RENTANG_WAKTU.items():
            recap_dict[cat] = grouped.apply(lambda x: get_min_time_in_range(x, s, e))

        df_res = pd.DataFrame(recap_dict).reset_index()
        if not df_res.empty: df_res.rename(columns={COL_NAMA: 'Nama Karyawan'}, inplace=True)
        
        df_final = pd.merge(pd.DataFrame({'Nama Karyawan': URUTAN_NAMA_CUSTOM}), df_res, on='Nama Karyawan', how='left')
        df_final[list(RENTANG_WAKTU.keys())] = df_final[list(RENTANG_WAKTU.keys())].fillna('')

        # --- DOWNLOAD EXCEL (DENGAN LOGIKA TELAT & KETERANGAN) ---
        out = io.BytesIO()
        wb = xlsxwriter.Workbook(out, {'in_memory': True})
        ws = wb.add_worksheet('Rekap')
        
        # Format Excel
        fmt_head = wb.add_format({'bold': True, 'fg_color': '#4caf50', 'font_color': 'white', 'border': 1, 'align': 'center'})
        fmt_norm = wb.add_format({'border': 1, 'align': 'center'})
        fmt_miss = wb.add_format({'bg_color': '#FF0000', 'border': 1}) 
        fmt_full = wb.add_format({'bg_color': '#FFFF00', 'border': 1})
        
        # Format Baru: Teks Merah untuk yang Telat
        fmt_late = wb.add_format({'font_color': 'red', 'bold': True, 'border': 1, 'align': 'center'})

        headers = ['Nama Karyawan', 'Pagi', 'Siang_1', 'Siang_2', 'Sore', 'Keterangan']
        ws.write_row(0, 0, headers, fmt_head)
        ws.set_column(0, 0, 30)
        ws.set_column(1, 5, 15)

        for idx, row in df_final.iterrows():
            nm = row['Nama Karyawan']
            pagi = row['Pagi']
            manual_stat = status_today.get(nm, "")
            
            ws.write(idx+1, 0, nm, fmt_norm)
            
            # Tulis Keterangan di Kolom F
            ws.write(idx+1, 5, manual_stat, fmt_norm)

            # Logika Warna Sel
            times = [row['Pagi'], row['Siang_1'], row['Siang_2'], row['Sore']]
            empty = sum(1 for t in times if t == '')

            if manual_stat:
                # Jika ada izin, kolom jam kosong biasa (putih)
                for i, t in enumerate(times):
                    ws.write(idx+1, i+1, "", fmt_norm)
            elif empty == 4:
                # Bolos Total -> Kuning
                for i in range(4):
                    ws.write(idx+1, i+1, "", fmt_full)
            else:
                # Hadir / Partial
                
                # 1. JAM PAGI (Logika Telat)
                if pagi == '':
                    ws.write(idx+1, 1, "", fmt_miss) # Kosong merah
                else:
                    if is_late(pagi):
                        ws.write(idx+1, 1, pagi, fmt_late) # Merah font (Telat)
                    else:
                        ws.write(idx+1, 1, pagi, fmt_norm) # Normal

                # 2. JAM SISANYA (Siang-Sore)
                rest_times = [row['Siang_1'], row['Siang_2'], row['Sore']]
                for i, t in enumerate(rest_times):
                    col_idx = i + 2 # Mulai dari kolom C
                    if t == '':
                        ws.write(idx+1, col_idx, "", fmt_miss)
                    else:
                        ws.write(idx+1, col_idx, t, fmt_norm)

        wb.close()
        out.seek(0)
        st.download_button("📥 Download Smart Report", out, f"Rekap_Smart_{sel_date}.xlsx", use_container_width=True)
        
        st.markdown("<br>", unsafe_allow_html=True)

        # --- GRID CARD ---
        if search_q: df_final = df_final[df_final['Nama Karyawan'].str.contains(search_q, case=False, na=False)]
        
        COLS = 4
        rows = [df_final.iloc[i:i+COLS] for i in range(0, len(df_final), COLS)]
        
        for r in rows:
            cols = st.columns(COLS)
            for i, (idx, row) in enumerate(r.iterrows()):
                with cols[i]:
                    nm = row['Nama Karyawan']
                    pagi = row['Pagi']
                    times = [pagi, row['Siang_1'], row['Siang_2'], row['Sore']]
                    empty = sum(1 for t in times if t == '')
                    manual_stat = status_today.get(nm, "")

                    # Tentukan Tema Kartu
                    if manual_stat:
                        lbl = f"ℹ️ {manual_stat}"; theme = "card-permit"; clr = "#d580ff"
                    elif empty == 4:
                        lbl = "⛔ TOTAL ABSENT"; theme = "card-absent"; clr = "#FF416C"
                    elif empty > 0:
                        lbl = "⚠️ PARTIAL"; theme = "card-partial"; clr = "#FFC837"
                    else:
                        lbl = "✅ FULL PRESENT"; theme = "card-present"; clr = "#00f260"

                    # Indikator Telat di Web
                    late_html = ""
                    if pagi and is_late(pagi):
                        late_html = '<span class="late-indicator">⏰ LATE</span>'

                    avt = f"https://ui-avatars.com/api/?name={nm.replace(' ', '+')}&background=random&color=fff"
                    
                    st.markdown(f"""
                    <div class="card {theme}">
                        <div class="card-header">
                            <img src="{avt}" class="avatar">
                            <div><p class="card-name">{nm}</p><p class="card-id">NP-{100+idx}</p></div>
                        </div>
                        <div class="detail-row">
                            <span class="label">Datang</span>
                            <span><span class="value">{pagi if pagi else '-'}</span> {late_html}</span>
                        </div>
                        <div class="detail-row"><span class="label">Pulang</span><span class="value">{row['Sore'] if row['Sore'] else '-'}</span></div>
                        <div style="text-align:right; font-size:0.7rem; font-weight:bold; color:{clr}; margin-top:10px;">{lbl}</div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    with st.popover("Detail", use_container_width=True):
                        if manual_stat: st.info(f"Status: {manual_stat}")
                        c1, c2 = st.columns(2)
                        # Pagi (Cek Telat di Popover juga)
                        pagi_val = pagi if pagi else "❌"
                        if is_late(pagi): pagi_val += " (Telat)"
                        c1.metric("Pagi", pagi_val)
                        c1.metric("Siang 1", row['Siang_1'] if row['Siang_1'] else "❌")
                        c2.metric("Siang 2", row['Siang_2'] if row['Siang_2'] else "❌")
                        c2.metric("Sore", row['Sore'] if row['Sore'] else "❌")
else:
    st.info("Menghubungkan Database...")