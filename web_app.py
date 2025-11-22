import streamlit as st
import pandas as pd
from datetime import time
import io
import time as t_sleep
import xlsxwriter

# --- KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="Nino's Project - Hybrid System",
    page_icon="🛡️",
    layout="wide"
)

# ==========================================
# 👇 DATABASE TERINTEGRASI 👇
# ==========================================

# 1. LINK DATA ABSEN (MESIN FINGER - DARI DATA SEBELUMNYA)
SHEET_URL_ABSEN = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ_GhoIb1riX98FsP8W4f2-_dH_PLcLDZskjNOyDcnnvOhBg8FUp3xJ-c_YgV0Pw71k4STy4rR0_MS5/pub?output=csv"

# 2. LINK DATA KETERANGAN (GOOGLE FORM - YANG BARU BAPAK KIRIM)
SHEET_URL_STATUS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ2QrBN8uTRiHINCEcZBrbdU-gzJ4pN2UljoqYG6NMoUQIK02yj_D1EdlxdPr82Pbr94v2o6V0Vh3Kt/pub?output=csv"

# ==========================================

# --- CSS: TAMPILAN PRESTIGE DARK ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;500;700&family=JetBrains+Mono:wght@400;700&display=swap');

    /* BACKGROUND */
    .stApp {
        background-color: #050505;
        background-image: 
            radial-gradient(at 0% 0%, hsla(253,16%,7%,1) 0, transparent 50%), 
            radial-gradient(at 50% 0%, hsla(225,39%,30%,1) 0, transparent 50%), 
            radial-gradient(at 100% 0%, hsla(339,49%,30%,1) 0, transparent 50%);
        color: white;
    }

    /* TYPOGRAPHY */
    .brand-title {
        font-family: 'Outfit', sans-serif;
        font-size: 3.5rem;
        font-weight: 800;
        background: linear-gradient(to right, #00dbde, #fc00ff);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0;
    }
    .brand-subtitle {
        font-family: 'Outfit', sans-serif;
        color: #a0a0a0;
        font-size: 1.1rem;
        letter-spacing: 4px;
        text-transform: uppercase;
        margin-bottom: 40px;
        border-bottom: 1px solid #333;
        padding-bottom: 20px;
    }

    /* CARD STYLES */
    .card {
        border-radius: 16px;
        padding: 20px;
        margin-bottom: 20px;
        position: relative;
        overflow: hidden;
        border: 1px solid rgba(255,255,255,0.05);
        transition: transform 0.3s ease, box-shadow 0.3s ease;
        background: #1a1a1a;
    }
    .card:hover { transform: translateY(-5px); z-index: 10; }

    /* --- TEMA WARNA KARTU --- */
    
    /* HIJAU (HADIR) */
    .card-present { background: linear-gradient(160deg, #051a10 0%, #000000 100%); border-top: 3px solid #00f260; }
    .card-present:hover { box-shadow: 0 10px 30px rgba(0, 242, 96, 0.15); border-color: #00f260; }
    .card-present .status-text { color: #00f260; }

    /* ORANYE (PARTIAL) */
    .card-partial { background: linear-gradient(160deg, #1a1205 0%, #000000 100%); border-top: 3px solid #FFC837; }
    .card-partial:hover { box-shadow: 0 10px 30px rgba(255, 200, 55, 0.15); border-color: #FFC837; }
    .card-partial .status-text { color: #FFC837; }

    /* MERAH (BOLOS) */
    .card-absent { background: linear-gradient(160deg, #1a0505 0%, #000000 100%); border-top: 3px solid #FF416C; }
    .card-absent:hover { box-shadow: 0 10px 30px rgba(255, 65, 108, 0.15); border-color: #FF416C; }
    .card-absent .status-text { color: #FF416C; }

    /* UNGU (IZIN/SAKIT/CR - DARI GOOGLE FORM) */
    .card-permit { 
        background: linear-gradient(160deg, #12051a 0%, #000000 100%); 
        border-top: 3px solid #d580ff; 
    }
    .card-permit:hover { 
        box-shadow: 0 10px 30px rgba(213, 128, 255, 0.15); 
        border-color: #d580ff; 
    }
    .card-permit .status-text { color: #d580ff; }

    /* ELEMENTS */
    .card-header { display: flex; align-items: center; gap: 12px; margin-bottom: 15px; border-bottom: 1px solid rgba(255,255,255,0.1); padding-bottom: 10px; }
    .avatar { width: 45px; height: 45px; border-radius: 12px; border: 2px solid rgba(255,255,255,0.2); }
    .card-name { font-family: 'Outfit', sans-serif; font-weight: 700; font-size: 0.95rem; color: white; margin: 0; }
    .card-id { font-size: 0.7rem; color: #888; font-family: 'JetBrains Mono', monospace; margin: 0; }
    .detail-row { display: flex; justify-content: space-between; margin-bottom: 5px; font-size: 0.8rem; color: #ccc; }
    .status-text { font-size: 0.7rem; text-transform: uppercase; letter-spacing: 1px; font-weight: bold; text-align: right; margin-top: 10px; }

    /* INPUT & BUTTON */
    .stTextInput input { background: #121212 !important; border: 1px solid #333 !important; color: white !important; border-radius: 8px; }
    .stDownloadButton button { background: linear-gradient(90deg, #00c6ff, #0072ff) !important; color: white !important; font-weight: 800 !important; border: none !important; }
    div[data-testid="stPopover"] button { width: 100%; border: 1px solid rgba(255,255,255,0.2); background: transparent; color: #aaa; font-size: 0.7rem; }
</style>
""", unsafe_allow_html=True)

# --- HEADER ---
st.markdown('<div class="brand-title">NINO\'S PROJECT</div>', unsafe_allow_html=True)
st.markdown('<div class="brand-subtitle">HYBRID ATTENDANCE SYSTEM (AUTO + FORM)</div>', unsafe_allow_html=True)

# --- SETUP DATA ---
COL_NAMA = 'Person Name'
COL_TIMESTAMP = 'Event Time'
ABSENCE_MARKER = '' 

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

# --- LOAD DATA FUNCTIONS ---
@st.cache_data(ttl=30) 
def load_absen(url):
    try:
        df = pd.read_csv(url)
        df.columns = df.columns.str.strip() # Bersihkan spasi di nama kolom
        df[COL_NAMA] = df[COL_NAMA].astype(str).str.strip()
        df = df[df[COL_NAMA] != ''].copy()
        df[COL_TIMESTAMP] = pd.to_datetime(df[COL_TIMESTAMP])
        df['Tanggal'] = df[COL_TIMESTAMP].dt.date
        df['Waktu'] = df[COL_TIMESTAMP].dt.time
        return df
    except Exception as e:
        return None

@st.cache_data(ttl=30)
def load_status(url):
    try:
        df = pd.read_csv(url)
        df.columns = df.columns.str.strip() # Bersihkan spasi header
        # Pastikan nama kolom sesuai Google Form (Biasanya: Timestamp, Tanggal, Nama Karyawan, Keterangan)
        # Kita rename biar aman
        df = df.rename(columns=lambda x: x.strip()) # Strip semua header
        
        # Konversi
        df['Nama Karyawan'] = df['Nama Karyawan'].astype(str).str.strip()
        df['Tanggal'] = pd.to_datetime(df['Tanggal']).dt.date
        df['Keterangan'] = df['Keterangan'].astype(str).str.strip().str.upper()
        return df
    except Exception as e:
        return None

# --- MAIN LOGIC ---
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

    # 1. Filter Absen Harian
    if sel_date:
        df_today = df_absen[df_absen['Tanggal'] == sel_date]
    else:
        df_today = pd.DataFrame(columns=[COL_NAMA, 'Tanggal']) # Kosong

    # 2. Filter Status Harian (Izin/Sakit)
    status_today = {}
    if df_status is not None and sel_date:
        # Ambil data status yang tanggalnya sama dengan selected_date
        df_status_today = df_status[df_status['Tanggal'] == sel_date]
        if not df_status_today.empty:
            # Mapping: {'Nama Orang': 'SAKIT', 'Nama Lain': 'IZIN'}
            status_today = pd.Series(df_status_today.Keterangan.values, index=df_status_today['Nama Karyawan']).to_dict()

    # 3. Proses Rekap Absensi
    recap_dict = {}
    grouped = df_today.groupby([COL_NAMA, 'Tanggal'])
    for cat, (s, e) in RENTANG_WAKTU.items():
        recap_dict[cat] = grouped.apply(lambda x: get_min_time_in_range(x, s, e))

    df_res = pd.DataFrame(recap_dict).reset_index()
    if not df_res.empty:
        df_res.rename(columns={COL_NAMA: 'Nama Karyawan'}, inplace=True)
    
    # Gabung dengan Master 101 Nama
    df_final = pd.merge(pd.DataFrame({'Nama Karyawan': URUTAN_NAMA_CUSTOM}), df_res, on='Nama Karyawan', how='left')
    df_final[list(RENTANG_WAKTU.keys())] = df_final[list(RENTANG_WAKTU.keys())].fillna('')

    # --- DOWNLOAD EXCEL (HYBRID) ---
    out = io.BytesIO()
    wb = xlsxwriter.Workbook(out, {'in_memory': True})
    ws = wb.add_worksheet('Rekap')
    
    # Format
    fmt_head = wb.add_format({'bold': True, 'fg_color': '#4caf50', 'font_color': 'white', 'border': 1, 'align': 'center'})
    fmt_norm = wb.add_format({'border': 1, 'align': 'center'})
    
    # Warna Warni Excel
    fmt_miss = wb.add_format({'bg_color': '#FF0000', 'border': 1}) # Merah (Bolong)
    fmt_full_absent = wb.add_format({'bg_color': '#FFFF00', 'border': 1}) # Kuning (Bolos Total)
    fmt_permit = wb.add_format({'bg_color': '#d580ff', 'border': 1, 'align': 'center'}) # Ungu (Izin)

    headers = ['Nama Karyawan', 'Pagi', 'Siang_1', 'Siang_2', 'Sore', 'Keterangan']
    ws.write_row(0, 0, headers, fmt_head)
    ws.set_column(0, 0, 30) # Lebar Nama
    ws.set_column(1, 5, 15) # Lebar Jam & Ket

    for idx, row in df_final.iterrows():
        nm = row['Nama Karyawan']
        times = [row['Pagi'], row['Siang_1'], row['Siang_2'], row['Sore']]
        empty_count = sum(1 for t in times if t == '')
        
        # Cek apakah dia ada di Daftar Izin?
        manual_status = status_today.get(nm, "") # Ambil status (SAKIT/IZIN), kosong jika tidak ada
        
        ws.write(idx+1, 0, nm, fmt_norm)
        
        if manual_status:
            # JIKA ADA IZIN -> TULIS DI KETERANGAN & WARNA UNGU
            # (Opsional: Mau di-merge kolom jamnya atau dibiarkan kosong berwarna ungu)
            # Kita buat kolom jamnya kosong tapi ungu, kolom keterangan diisi
            for i in range(4):
                ws.write(idx+1, i+1, "", fmt_permit)
            ws.write(idx+1, 5, manual_status, fmt_permit)
            
        elif empty_count == 4:
            # BOLOS TOTAL -> KUNING
            for i in range(4):
                ws.write(idx+1, i+1, "", fmt_full_absent)
            ws.write(idx+1, 5, "", fmt_full_absent)
            
        else:
            # HADIR / PARTIAL
            for i, t in enumerate(times):
                if t == '':
                    ws.write(idx+1, i+1, "", fmt_miss) # Merah kalau bolong
                else:
                    ws.write(idx+1, i+1, t, fmt_norm)
            ws.write(idx+1, 5, "", fmt_norm)

    wb.close()
    out.seek(0)
    st.download_button("📥 Download Absence Report Here", out, f"Rekap_{sel_date}.xlsx", use_container_width=True)
    
    st.markdown("<br>", unsafe_allow_html=True)

    # --- GRID CARD DISPLAY ---
    if search_q:
        df_final = df_final[df_final['Nama Karyawan'].str.contains(search_q, case=False, na=False)]

    COLS = 4
    rows = [df_final.iloc[i:i+COLS] for i in range(0, len(df_final), COLS)]
    
    for r in rows:
        cols = st.columns(COLS)
        for i, (idx, row) in enumerate(r.iterrows()):
            with cols[i]:
                nm = row['Nama Karyawan']
                times = [row['Pagi'], row['Siang_1'], row['Siang_2'], row['Sore']]
                empty = sum(1 for t in times if t == '')
                
                # Cek Status Manual dari Google Form
                manual_stat = status_today.get(nm, None)

                if manual_stat:
                    # TAMPILAN KARTU IZIN (UNGU)
                    lbl = f"ℹ️ {manual_stat}"
                    theme = "card-permit"
                    clr = "#d580ff"
                elif empty == 4:
                    lbl = "⛔ TOTAL ABSENT"; theme = "card-absent"; clr = "#FF416C"
                elif empty > 0:
                    lbl = "⚠️ PARTIAL"; theme = "card-partial"; clr = "#FFC837"
                else:
                    lbl = "✅ FULL PRESENT"; theme = "card-present"; clr = "#00f260"

                avt = f"https://ui-avatars.com/api/?name={nm.replace(' ', '+')}&background=random&color=fff"
                
                # RENDER HTML
                st.markdown(f"""
                <div class="card {theme}">
                    <div class="card-header">
                        <img src="{avt}" class="avatar">
                        <div><p class="card-name">{nm}</p><p class="card-id">NP-{100+idx}</p></div>
                    </div>
                    <div class="detail-row"><span class="label">In</span><span class="value">{row['Pagi'] if row['Pagi'] else '-'}</span></div>
                    <div class="detail-row"><span class="label">Out</span><span class="value">{row['Sore'] if row['Sore'] else '-'}</span></div>
                    <div class="status-text" style="color: {clr};">{lbl}</div>
                </div>
                """, unsafe_allow_html=True)
                
                # POPOVER DETAIL
                with st.popover("Show Detail", use_container_width=True):
                    st.caption(f"Record for: {nm}")
                    if manual_stat:
                        st.info(f"Status: {manual_stat}")
                    else:
                        c1, c2 = st.columns(2)
                        c1.metric("Pagi", row['Pagi'] if row['Pagi'] else "-")
                        c1.metric("Siang 1", row['Siang_1'] if row['Siang_1'] else "-")
                        c2.metric("Siang 2", row['Siang_2'] if row['Siang_2'] else "-")
                        c2.metric("Sore", row['Sore'] if row['Sore'] else "-")

else:
    st.info("Connecting to Database...")