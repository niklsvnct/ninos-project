import streamlit as st
import pandas as pd
from datetime import time
import io
import time as t_sleep
import xlsxwriter

# --- KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="Nino's Project - V10",
    page_icon="🧬",
    layout="wide"
)

# --- LINK DATABASE ---
SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ_GhoIb1riX98FsP8W4f2-_dH_PLcLDZskjNOyDcnnvOhBg8FUp3xJ-c_YgV0Pw71k4STy4rR0_MS5/pub?output=csv"

# --- CSS: AMBIENT GRADIENT CARDS ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;500;700&family=JetBrains+Mono:wght@400;700&display=swap');

    /* BACKGROUND UTAMA (TETAP GELAP BERTEKSTUR) */
    .stApp {
        background-color: #0a0a0a;
        background-image: 
            linear-gradient(rgba(255, 255, 255, 0.03) 1px, transparent 1px),
            linear-gradient(90deg, rgba(255, 255, 255, 0.03) 1px, transparent 1px);
        background-size: 30px 30px;
        color: white;
    }

    /* TYPOGRAPHY */
    .brand-title {
        font-family: 'Outfit', sans-serif;
        font-size: 3.5rem;
        font-weight: 800;
        background: linear-gradient(to right, #4facfe 0%, #00f2fe 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0;
    }
    .brand-subtitle {
        font-family: 'Outfit', sans-serif;
        color: #888;
        font-size: 1rem;
        letter-spacing: 4px;
        text-transform: uppercase;
        margin-bottom: 40px;
        border-bottom: 1px solid #333;
        padding-bottom: 20px;
    }

    /* --- NEW CARD DESIGN: AMBIENT GRADIENTS --- */
    
    .card {
        border-radius: 16px;
        padding: 20px;
        margin-bottom: 20px;
        position: relative;
        overflow: hidden;
        border: 1px solid rgba(255,255,255,0.05);
        transition: transform 0.3s ease, box-shadow 0.3s ease;
    }
    
    /* 1. TEMA PRESENT (Hadir Full - Gradasi Hijau Tua Mewah) */
    .card-present {
        background: linear-gradient(160deg, #051a10 0%, #000000 100%);
        border-top: 3px solid #00f260;
    }
    .card-present:hover {
        box-shadow: 0 10px 30px rgba(0, 242, 96, 0.15);
        border-color: #00f260;
    }

    /* 2. TEMA PARTIAL (Bolong - Gradasi Emas Tua) */
    .card-partial {
        background: linear-gradient(160deg, #1a1205 0%, #000000 100%);
        border-top: 3px solid #FFC837;
    }
    .card-partial:hover {
        box-shadow: 0 10px 30px rgba(255, 200, 55, 0.15);
        border-color: #FFC837;
    }

    /* 3. TEMA ABSENT (Bolos - Gradasi Merah Maroon Gelap) */
    .card-absent {
        background: linear-gradient(160deg, #1a0505 0%, #000000 100%);
        border-top: 3px solid #FF416C;
    }
    .card-absent:hover {
        box-shadow: 0 10px 30px rgba(255, 65, 108, 0.15);
        border-color: #FF416C;
    }

    /* ELEMENT DALAM KARTU */
    .card-header {
        display: flex;
        align-items: center;
        gap: 15px;
        margin-bottom: 15px;
        padding-bottom: 15px;
        border-bottom: 1px solid rgba(255,255,255,0.05);
    }
    .avatar {
        width: 50px;
        height: 50px;
        border-radius: 12px; /* Squircle style */
        border: 2px solid rgba(255,255,255,0.1);
    }
    .card-name {
        font-family: 'Outfit', sans-serif;
        font-weight: 700;
        font-size: 1rem;
        color: white;
        margin: 0;
    }
    .card-id {
        font-size: 0.7rem;
        color: #888;
        font-family: monospace;
        margin: 0;
    }
    
    .detail-row {
        display: flex;
        justify-content: space-between;
        margin-bottom: 8px;
        font-size: 0.85rem;
    }
    .label { color: #aaa; }
    .value { color: #fff; font-weight: 600; font-family: 'JetBrains Mono', monospace;}

    /* STATUS BADGE DI DALAM KARTU */
    .status-text {
        font-size: 0.7rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        font-weight: bold;
        text-align: right;
        margin-top: 10px;
        opacity: 0.8;
    }

    /* POPOVER BUTTON */
    div[data-testid="stPopover"] button {
        width: 100%;
        border: 1px solid rgba(255,255,255,0.1);
        background: rgba(255,255,255,0.05);
        color: white;
        font-weight: bold;
        font-size: 0.8rem;
        text-transform: uppercase;
    }
    div[data-testid="stPopover"] button:hover {
        border-color: white;
        background: rgba(255,255,255,0.1);
    }

    /* INPUT & BUTTON LAIN */
    .stTextInput input {
        background: #121212 !important;
        border: 1px solid #333 !important;
        color: white !important;
    }
    .stDownloadButton button {
        background: linear-gradient(90deg, #4facfe 0%, #00f2fe 100%) !important;
        border: none !important;
        color: #000 !important;
        font-weight: 800 !important;
    }
</style>
""", unsafe_allow_html=True)

# --- HEADER ---
st.markdown('<div class="brand-title">NINO\'S PROJECT</div>', unsafe_allow_html=True)
st.markdown('<div class="brand-subtitle">ADVANCED ATTENDANCE ANALYTICS</div>', unsafe_allow_html=True)

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

# --- LOAD DATA (AUTO REFRESH) ---
@st.cache_data(ttl=60) 
def load_data(url):
    try:
        df = pd.read_csv(url)
        df.columns = df.columns.str.strip()
        if COL_NAMA not in df.columns or COL_TIMESTAMP not in df.columns:
            st.error(f"⚠️ ERROR: Kolom '{COL_NAMA}' atau '{COL_TIMESTAMP}' tidak ditemukan di Google Sheet.")
            st.stop()
        df[COL_NAMA] = df[COL_NAMA].astype(str).str.strip()
        df = df[df[COL_NAMA] != ''].copy()
        df[COL_TIMESTAMP] = pd.to_datetime(df[COL_TIMESTAMP])
        df['Tanggal'] = df[COL_TIMESTAMP].dt.date
        df['Waktu'] = df[COL_TIMESTAMP].dt.time
        return df
    except Exception as e:
        st.error(f"Gagal membaca Google Sheet: {e}")
        return None

# --- MAIN APP ---
df_raw = load_data(SHEET_URL)

if df_raw is not None:
    col_date, col_search = st.columns([1, 3])
    with col_date:
        available_dates = sorted(df_raw['Tanggal'].unique())
        default_date = available_dates[-1] if available_dates else None
        selected_date = st.date_input("📅 Pilih Tanggal", value=default_date)
    with col_search:
        search_query = st.text_input("🔍 Search Employee", placeholder="Ketik nama karyawan...")

    st.markdown("---")

    if selected_date not in available_dates:
        st.warning(f"⚠️ Tidak ada data absensi untuk tanggal: {selected_date}")
    else:
        df_today = df_raw[df_raw['Tanggal'] == selected_date]
        recap_dict = {}
        grouped = df_today.groupby([COL_NAMA, 'Tanggal'])
        for category, (start_t, end_t) in RENTANG_WAKTU.items():
            recap_dict[category] = grouped.apply(lambda x: get_min_time_in_range(x, start_t, end_t))
        df_hasil_scan = pd.DataFrame(recap_dict).reset_index()
        df_hasil_scan.rename(columns={COL_NAMA: 'Nama Karyawan'}, inplace=True)
        df_master = pd.DataFrame({'Nama Karyawan': URUTAN_NAMA_CUSTOM})
        df_final = pd.merge(df_master, df_hasil_scan, on='Nama Karyawan', how='left')
        col_fill = list(RENTANG_WAKTU.keys())
        df_final[col_fill] = df_final[col_fill].fillna('')

        # --- DOWNLOAD ---
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Rekap Absensi')
        fmt_header = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#4caf50', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        fmt_normal = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        fmt_red = workbook.add_format({'bg_color': '#FF0000', 'border': 1, 'align': 'center'}) 
        fmt_yellow = workbook.add_format({'bg_color': '#FFFF00', 'border': 1, 'align': 'center'}) 
        headers = ['Nama Karyawan', 'Pagi', 'Siang_1', 'Siang_2', 'Sore']
        for col_num, header in enumerate(headers):
            worksheet.write(0, col_num, header, fmt_header)
            worksheet.set_column(col_num, col_num, 15)
        worksheet.set_column(0, 0, 35)
        for row_num, row_data in df_final.iterrows():
            nama = row_data['Nama Karyawan']
            times = [row_data['Pagi'], row_data['Siang_1'], row_data['Siang_2'], row_data['Sore']]
            worksheet.write(row_num + 1, 0, nama, fmt_normal)
            empty_count = sum(1 for t in times if t == '')
            if empty_count == 4:
                for col_idx in range(4):
                    worksheet.write(row_num + 1, col_idx + 1, "", fmt_yellow)
            else:
                for col_idx, time_val in enumerate(times):
                    if time_val == '':
                        worksheet.write(row_num + 1, col_idx + 1, "", fmt_red)
                    else:
                        worksheet.write(row_num + 1, col_idx + 1, time_val, fmt_normal)
        workbook.close()
        output.seek(0)
        st.download_button(label=f"📥 Download Absence Employe here ({selected_date})", data=output, file_name=f"Rekap_{selected_date}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        st.markdown("<br>", unsafe_allow_html=True)

        # --- GRID CARD (NEW DESIGN) ---
        if search_query:
            df_display = df_final[df_final['Nama Karyawan'].str.contains(search_query, case=False, na=False)]
        else:
            df_display = df_final

        COLS_PER_ROW = 4
        rows = [df_display.iloc[i:i+COLS_PER_ROW] for i in range(0, len(df_display), COLS_PER_ROW)]
        for row_chunk in rows:
            cols = st.columns(COLS_PER_ROW)
            for idx, (index, row) in enumerate(row_chunk.iterrows()):
                with cols[idx]:
                    
                    # LOGIC STATUS & CARD THEME
                    times = [row['Pagi'], row['Siang_1'], row['Siang_2'], row['Sore']]
                    empty_count = sum(1 for t in times if t == '')
                    
                    # Menentukan Kelas CSS Berdasarkan Status
                    if empty_count == 4:
                        status_label = "TOTAL ABSENT"
                        card_theme = "card-absent" # Merah Gelap
                        status_color = "#FF416C"
                    elif empty_count > 0:
                        status_label = "PARTIAL"
                        card_theme = "card-partial" # Emas Gelap
                        status_color = "#FFC837"
                    else:
                        status_label = "FULL PRESENT"
                        card_theme = "card-present" # Hijau Gelap
                        status_color = "#00f260"

                    # TAMPILKAN KARTU
                    with st.container():
                        
                        # HTML untuk Kartu Berwarna
                        avatar_url = f"https://ui-avatars.com/api/?name={row['Nama Karyawan'].replace(' ', '+')}&background=random&color=fff&size=128"
                        
                        st.markdown(f"""
                        <div class="card {card_theme}">
                            <div class="card-header">
                                <img src="{avatar_url}" class="avatar">
                                <div>
                                    <p class="card-name">{row['Nama Karyawan']}</p>
                                    <p class="card-id">ID: NP-{100+index}</p>
                                </div>
                            </div>
                            <div class="detail-row">
                                <span class="label">Datang</span><span class="value">{row['Pagi'] if row['Pagi'] else '-'}</span>
                            </div>
                            <div class="detail-row">
                                <span class="label">Pulang</span><span class="value">{row['Sore'] if row['Sore'] else '-'}</span>
                            </div>
                            <div class="status-text" style="color: {status_color};">{status_label}</div>
                        </div>
                        """, unsafe_allow_html=True)

                        # Popover Detail
                        with st.popover("LIHAT DETAIL JAM", use_container_width=True):
                            st.markdown(f"**Detail: {row['Nama Karyawan']}**")
                            st.caption(f"Status: {status_label}")
                            st.divider()
                            c1, c2 = st.columns(2)
                            c1.metric("Pagi", row['Pagi'] if row['Pagi'] else "❌")
                            c1.metric("Siang 1", row['Siang_1'] if row['Siang_1'] else "❌")
                            c2.metric("Siang 2", row['Siang_2'] if row['Siang_2'] else "❌")
                            c2.metric("Sore", row['Sore'] if row['Sore'] else "❌")
else:
    st.write("Menghubungkan ke Database...")