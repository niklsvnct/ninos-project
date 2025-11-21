import streamlit as st
import pandas as pd
from datetime import time
import io
import time as t_sleep
import xlsxwriter

# --- KONFIGURASI HALAMAN ---
st.set_page_config(
    page_title="Nino's Project - Chroma",
    page_icon="🌈",
    layout="wide"
)

# --- CSS: CHROMA GLASSMORPHISM STYLE ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;500;700&family=JetBrains+Mono:wght@400;700&display=swap');

    /* BACKGROUND UTAMA */
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
        letter-spacing: -1px;
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
        display: inline-block;
    }

    /* CARD BASE STYLE */
    .card {
        border-radius: 16px;
        padding: 2px; /* Untuk border gradient */
        position: relative;
        transition: transform 0.3s ease;
        background: #1a1a1a; /* Fallback */
    }
    .card:hover {
        transform: translateY(-5px) scale(1.02);
        z-index: 10;
    }

    /* INNER CONTENT CARD */
    .card-content {
        background: rgba(20, 20, 20, 0.95);
        border-radius: 14px;
        padding: 15px;
        height: 100%;
    }

    /* --- COLOR THEMES (THE MAGIC HAPPENS HERE) --- */
    
    /* THEME: PRESENT (HIJAU-BIRU) */
    .theme-present {
        background: linear-gradient(135deg, #00F260, #0575E6);
        box-shadow: 0 0 20px rgba(0, 242, 96, 0.1);
    }
    .theme-present .card-status { color: #00F260; }
    
    /* THEME: PARTIAL (ORANYE-EMAS) */
    .theme-partial {
        background: linear-gradient(135deg, #FF8008, #FFC837);
        box-shadow: 0 0 20px rgba(255, 128, 8, 0.1);
    }
    .theme-partial .card-status { color: #FF8008; }

    /* THEME: ABSENT (MERAH-PINK) */
    .theme-absent {
        background: linear-gradient(135deg, #EB3349, #F45C43);
        box-shadow: 0 0 20px rgba(235, 51, 73, 0.1);
    }
    .theme-absent .card-status { color: #EB3349; }


    /* CARD ELEMENTS */
    .card-header {
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 15px;
        border-bottom: 1px solid rgba(255,255,255,0.1);
        padding-bottom: 10px;
    }
    .avatar {
        width: 42px;
        height: 42px;
        border-radius: 10px; /* Kotak rounded biar modern */
        border: 2px solid rgba(255,255,255,0.2);
    }
    .card-name {
        font-family: 'Outfit', sans-serif;
        font-weight: 700;
        font-size: 0.9rem;
        color: white;
        margin: 0;
        line-height: 1.2;
    }
    .card-id {
        font-size: 0.7rem;
        color: #888;
        margin: 0;
        font-family: 'JetBrains Mono', monospace;
    }

    .card-body {
        display: flex;
        justify-content: space-between;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.8rem;
        color: #ccc;
        background: rgba(0,0,0,0.3);
        padding: 8px;
        border-radius: 8px;
        margin-bottom: 5px;
    }
    .card-label { color: #666; font-size: 0.7rem; display: block;}
    .card-time { font-weight: bold; }

    /* BUTTON STYLING (POPOVER) */
    div[data-testid="stPopover"] button {
        width: 100%;
        border-radius: 8px;
        font-weight: 800;
        font-size: 0.75rem;
        border: none;
        margin-top: 10px;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: white;
        background: rgba(255,255,255,0.1);
        transition: 0.3s;
    }
    div[data-testid="stPopover"] button:hover {
        background: rgba(255,255,255,0.2);
        transform: scale(1.02);
    }

    /* INPUT & DOWNLOAD */
    .stTextInput input {
        background: #1a1a1a !important;
        border: 1px solid #333 !important;
        color: white !important;
        border-radius: 10px;
    }
    .stDownloadButton button {
        background: linear-gradient(90deg, #11998e, #38ef7d) !important;
        color: white !important;
        font-weight: 800 !important;
        border: none !important;
        border-radius: 12px !important;
        box-shadow: 0 5px 15px rgba(56, 239, 125, 0.3);
    }
</style>
""", unsafe_allow_html=True)

# --- HEADER ---
st.markdown('<div class="brand-title">NINO\'S PROJECT</div>', unsafe_allow_html=True)
st.markdown('<div class="brand-subtitle">Employee Attendance Automation System</div>', unsafe_allow_html=True)

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

# --- LAYOUT UTAMA ---
col_search, col_upload = st.columns([2, 1])

with col_search:
    search_query = st.text_input("🔍 Search Employee", placeholder="Ketik nama karyawan...")

with col_upload:
    uploaded_file = st.file_uploader("Upload Excel", type=['xlsx'], label_visibility="collapsed")

st.markdown("---")

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        df[COL_NAMA] = df[COL_NAMA].astype(str).str.strip()
        df = df[df[COL_NAMA] != ''].copy()
        df[COL_TIMESTAMP] = pd.to_datetime(df[COL_TIMESTAMP])
        df['Tanggal'] = df[COL_TIMESTAMP].dt.date
        df['Waktu'] = df[COL_TIMESTAMP].dt.time

        recap_dict = {}
        grouped = df.groupby([COL_NAMA, 'Tanggal'])
        
        for category, (start_t, end_t) in RENTANG_WAKTU.items():
            recap_dict[category] = grouped.apply(lambda x: get_min_time_in_range(x, start_t, end_t))

        df_hasil_scan = pd.DataFrame(recap_dict).reset_index()
        df_hasil_scan.rename(columns={COL_NAMA: 'Nama Karyawan'}, inplace=True)

        df_master = pd.DataFrame({'Nama Karyawan': URUTAN_NAMA_CUSTOM})
        df_final = pd.merge(df_master, df_hasil_scan, on='Nama Karyawan', how='left')
        
        col_fill = list(RENTANG_WAKTU.keys())
        df_final[col_fill] = df_final[col_fill].fillna('')

        # --- DOWNLOAD EXCEL ---
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
        
        st.download_button(
            label="📥 Download Absence Employe Here",
            data=output,
            file_name="Ninos_Report_V9.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.markdown("<br>", unsafe_allow_html=True)

        # --- RENDER CARD GRID CHROMA ---
        if search_query:
            df_display = df_final[df_final['Nama Karyawan'].str.contains(search_query, case=False, na=False)]
        else:
            df_display = df_final

        # Grid System
        COLS = 4
        rows = [df_display.iloc[i:i+COLS] for i in range(0, len(df_display), COLS)]

        for row_chunk in rows:
            cols = st.columns(COLS)
            for idx, (index, row) in enumerate(row_chunk.iterrows()):
                
                # Logic Status & Warna
                times = [row['Pagi'], row['Siang_1'], row['Siang_2'], row['Sore']]
                empty_count = sum(1 for t in times if t == '')
                
                if empty_count == 4:
                    theme_class = "theme-absent"
                    status_label = "⛔ ABSENT"
                    emoji = "🔴"
                elif empty_count > 0:
                    theme_class = "theme-partial"
                    status_label = "⚠️ PARTIAL"
                    emoji = "🟠"
                else:
                    theme_class = "theme-present"
                    status_label = "✅ PRESENT"
                    emoji = "🟢"

                with cols[idx]:
                    # CONTAINER UTAMA DENGAN CLASS TEMA
                    container = st.container()
                    
                    # Avatar Random Color (biar rame)
                    avatar_url = f"https://ui-avatars.com/api/?name={row['Nama Karyawan'].replace(' ', '+')}&background=random&color=fff&size=128"

                    # HTML VISUAL (GLASSMORPHISM CARD)
                    html_code = f"""
                    <div class="card {theme_class}">
                        <div class="card-content">
                            <div class="card-header">
                                <img src="{avatar_url}" class="avatar">
                                <div>
                                    <p class="card-name">{row['Nama Karyawan']}</p>
                                    <p class="card-id">ID: {100+index}</p>
                                </div>
                            </div>
                            <div class="card-body">
                                <div>
                                    <span class="card-label">DATANG</span>
                                    <span class="card-time">{row['Pagi'] if row['Pagi'] else '--:--'}</span>
                                </div>
                                <div style="text-align:right;">
                                    <span class="card-label">PULANG</span>
                                    <span class="card-time">{row['Sore'] if row['Sore'] else '--:--'}</span>
                                </div>
                            </div>
                        </div>
                    </div>
                    """
                    st.markdown(html_code, unsafe_allow_html=True)

                    # TOMBOL INTERAKTIF (POPOVER)
                    with st.popover(f"{emoji} {status_label}", use_container_width=True):
                        st.markdown(f"**Time Log: {row['Nama Karyawan']}**")
                        st.divider()
                        c1, c2 = st.columns(2)
                        c1.metric("Pagi", row['Pagi'] if row['Pagi'] else "-")
                        c1.metric("Siang 1", row['Siang_1'] if row['Siang_1'] else "-")
                        c2.metric("Siang 2", row['Siang_2'] if row['Siang_2'] else "-")
                        c2.metric("Sore", row['Sore'] if row['Sore'] else "-")

    except Exception as e:
        st.error(f"System Error: {e}")

else:
    st.info("Waiting for data upload...")