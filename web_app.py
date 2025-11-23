import streamlit as st
import pandas as pd
from datetime import time, datetime, timedelta
import io
import xlsxwriter
import streamlit.components.v1 as components

# ==========================================
# 1. PAGE CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="WedaBay Airport Command Center",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# 2. CONSTANTS & CONFIGURATION
# ==========================================
class Config:
    COL_NAMA = 'Person Name'
    COL_TIMESTAMP = 'Event Time'
    LATE_THRESHOLD = time(7, 5, 0)
    
    RENTANG_WAKTU = {
        'Pagi': ('03:00:00', '11:00:00'),
        'Siang_1': ('11:29:00', '12:30:59'),
        'Siang_2': ('12:31:00', '13:30:59'),
        'Sore': ('17:00:00', '23:59:59')
    }
    
    # Database URLs
    SHEET_URL_ABSEN = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ_GhoIb1riX98FsP8W4f2-_dH_PLcLDZskjNOyDcnnvOhBg8FUp3xJ-c_YgV0Pw71k4STy4rR0_MS5/pub?output=csv"
    SHEET_URL_STATUS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ2QrBN8uTRiHINCEcZBrbdU-gzJ4pN2UljoqYG6NMoUQIK02yj_D1EdlxdPr82Pbr94v2o6V0Vh3Kt/pub?gid=511860805&single=true&output=csv"
    GOOGLE_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSeopdaE-lyOtFd2TUr5C3K2DWE3Syt2PaKoXMp0cmWKIFnijw/viewform?usp=header"
    
    # Division Data
    DATA_DIVISI = {
        "LEADERSHIP": { "color": "#FFD700", "icon": "👨‍✈️", "code": "LDR", "members": ["Patra Anggana", "Su Adam", "Budiman Arifin", "Rifaldy Ilham Bhagaskara", "Marwan S Halid", "Budiono"] },
        "TLB": { "color": "#00A8FF", "icon": "🔧", "code": "TLB", "members": ["M. Ansori", "Bayu Pratama Putra Katuwu", "Yoga Nugraha Putra Pasaribu", "Junaidi Taib", "Muhammad Rizal Amra", "Rusli Dj"] },
        "TBL": { "color": "#0097e6", "icon": "📦", "code": "TBL", "members": ["Venesia Aprilia Ineke", "Muhammad Naufal Ramadhan", "Yuzak Gerson Puturuhu", "Muhamad Alief Wildan", "Gafur Hamisi", "Jul Akbar M. Nur", "Sarni Massiri", "Adrianto Laundang", "Wahyudi Ismail"] },
        "TRANS APRON": { "color": "#e1b12c", "icon": "🚌", "code": "APR", "members": ["Marichi Gita Rusdi", "Ilham Rahim", "Abdul Mu Iz Simal", "Dwiki Agus Saputro", "Moh. Sofyan", "Faisal M. Kadir", "Amirudin Rustam", "Faturrahman Kaunar", "Wawan Hermawan", "Rahmat Joni", "Nur Ichsan"] },
        "ATS": { "color": "#44bd32", "icon": "📡", "code": "ATS", "members": ["Nurul Tanti", "Firlon Paembong", "Irwan Rezky Setiawan", "Yusuf Arviansyah", "Nurdahlia Is. Folaimam", "Ghaly Rabbani Panji Indra", "Ikhsan Wahyu Vebriyan", "Rizki Mahardhika Ardi Tigo", "Nikolaus Vincent Quirino"] },
        "ADM COMPLIANCE": { "color": "#8c7ae6", "icon": "📋", "code": "ADM", "members": ["Yessicha Aprilyona Siregar", "Gabriela Margrith Louisa Klavert", "Aldi Saptono"] },
        "TRANSLATOR": { "color": "#00cec9", "icon": "🎧", "code": "TRN", "members": ["Wilyam Candra", "Norika Joselyn Modnissa"] },
        "AVSEC": { "color": "#c23616", "icon": "🛡️", "code": "SEC", "members": ["Andrian Maranatha", "Toni Nugroho Simarmata", "Muhamad Albi Ferano", "Andreas Charol Tandjung", "Sabadia Mahmud", "Rusdin Malagapi", "Muhamad Judhytia Winli", "Wahyu Samsudin", "Fientje Elisabeth Joseph", "Anglie Fitria Desiana Mamengko", "Dwi Purnama Bimasakti", "Windi Angriani Sulaeman", "Megawati A. Rauf"] },
        "GROUND HANDLING": { "color": "#e17055", "icon": "🚜", "code": "GND", "members": ["Yuda Saputra", "Tesalonika Gratia Putri Toar", "Esi Setia Ningseh", "Ardiyanto Kalatjo", "Febrianti Tikabala"] },
        "HELICOPTER": { "color": "#6c5ce7", "icon": "🚁", "code": "HEL", "members": ["Agung Sabar S. Taufik", "Recky Irwan R. A Arsyad", "Farok Abdul", "Achmad Rizky Ariz", "Yus Andi", "Muh. Noval Kipudjena"] },
        "AMC & TERMINAL": { "color": "#0984e3", "icon": "🏢", "code": "AMC", "members": ["Risky Sulung", "Muchamad Nur Syaifulrahman", "Muhammad Tunjung Rohmatullah", "Sunarty Fakir", "Albert Papuling", "Gibhran Fitransyah Yusri", "Muhdi R Tomia", "Riski Rifaldo Theofilus Anu", "Eko"] },
        "SAFETY OFFICER": { "color": "#fd79a8", "icon": "🦺", "code": "SFT", "members": ["Hildan Ahmad Zaelani", "Abdurahim Andar"] },
        "PKP-PK": { "color": "#fab1a0", "icon": "🚒", "code": "RES", "members": ["Andreas Aritonang", "Achmad Alwan Asyhab", "Doni Eka Satria", "Yogi Prasetya Eka Winandra", "Akhsin Aditya Weza Putra", "Fardhan Ahmad Tajali", "Maikel Renato Syafaruddin", "Saldi Sandra", "Hamzah M. Ali Gani", "Marfan Mandar", "Julham Keya", "Aditya Sugiantoro Abbas", "Muhamad Usman", "M Akbar D Patty", "Daniel Freski Wangka", "Fandi M.Naser", "Agung Fadjriansyah Ano", "Deni Hendri Bobode", "Muhammad Rifai", "Idrus Arsad, SH"] }
    }

# ==========================================
# 3. STYLING ENGINE (WEBSITE MAHAL)
# ==========================================
def apply_custom_css():
    st.markdown("""
    
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Rajdhani:wght@400;600;700&family=Inter:wght@300;400;600&family=JetBrains+Mono:wght@400;700&display=swap');
        
        /* === BASE THEME === */
        :root {
            --bg-color: #0b0e14;
            --panel-color: #151922;
            --accent-blue: #00a8ff;
            --accent-cyan: #00cec9;
            --alert-red: #e84118;
            --warning-amber: #fbc531;
            --success-green: #4cd137;
            --text-primary: #f5f6fa;
            --text-secondary: #7f8fa6;
            --border-color: #2f3640;
        }

        .stApp {
            background-color: var(--bg-color);
            background-image: radial-gradient(circle at 50% 0%, #1a237e 0%, transparent 50%);
            font-family: 'Inter', sans-serif;
            color: var(--text-primary);
        }
        
        /* === SIDEBAR === */
        section[data-testid="stSidebar"] {
            background-color: #0f1218;
            border-right: 1px solid var(--border-color);
        }
        
        /* === TYPOGRAPHY === */
        h1, h2, h3 { font-family: 'Rajdhani', sans-serif; text-transform: uppercase; letter-spacing: 1.5px; }
        
        .brand-title {
            font-family: 'Rajdhani', sans-serif;
            font-size: 3rem;
            font-weight: 700;
            background: linear-gradient(90deg, var(--accent-blue), var(--accent-cyan));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin-bottom: 0;
            text-shadow: 0 0 20px rgba(0, 168, 255, 0.3);
        }
        
        .brand-subtitle {
            font-family: 'JetBrains Mono', monospace;
            color: var(--text-secondary);
            font-size: 0.8rem;
            letter-spacing: 3px;
            text-transform: uppercase;
            margin-bottom: 30px;
            border-bottom: 1px solid var(--border-color);
            padding-bottom: 20px;
        }

        /* === METRICS === */
        div[data-testid="stMetric"] {
            background: rgba(21, 25, 34, 0.6);
            border: 1px solid var(--border-color);
            border-left: 3px solid var(--accent-blue);
            border-radius: 4px;
            padding: 15px;
            transition: all 0.3s ease;
        }
        div[data-testid="stMetric"]:hover {
            border-color: var(--accent-cyan);
            box-shadow: 0 0 15px rgba(0, 206, 201, 0.2);
            transform: translateY(-2px);
        }
        div[data-testid="stMetricLabel"] { font-family: 'Rajdhani', sans-serif; font-weight: 600; color: var(--text-secondary); }
        div[data-testid="stMetricValue"] { font-family: 'JetBrains Mono', monospace; font-size: 2rem; color: var(--text-primary); }

        /* === CARDS === */
        .card {
            background: var(--panel-color);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            padding: 0;
            margin-bottom: 15px;
            transition: all 0.3s;
        }
        .card:hover { border-color: var(--accent-blue); box-shadow: 0 4px 20px rgba(0,0,0,0.5); }
        
        .card-header {
            padding: 15px;
            display: flex;
            align-items: center;
            background: rgba(255,255,255,0.03);
            border-bottom: 1px dashed var(--border-color);
        }
        
        .card-body { padding: 15px; }
        .card-name { font-family: 'Rajdhani', sans-serif; font-weight: 700; font-size: 1.1rem; margin: 0; color: #fff; }
        .card-code { font-family: 'JetBrains Mono', monospace; font-size: 0.7rem; color: var(--text-secondary); letter-spacing: 1px; }
        
        .flight-row { display: flex; justify-content: space-between; margin-bottom: 8px; }
        .flight-label { font-size: 0.7rem; color: var(--text-secondary); text-transform: uppercase; }
        .flight-value { font-family: 'JetBrains Mono', monospace; font-weight: 700; font-size: 0.9rem; }
        
        /* Status Badges */
        .status-badge { font-family: 'Rajdhani', sans-serif; font-weight: 700; font-size: 0.8rem; padding: 2px 8px; border-radius: 2px; }
        .status-present { color: var(--success-green); border: 1px solid var(--success-green); background: rgba(76, 209, 55, 0.1); }
        .status-partial { color: var(--warning-amber); border: 1px solid var(--warning-amber); background: rgba(251, 197, 49, 0.1); }
        .status-absent { color: var(--alert-red); border: 1px solid var(--alert-red); background: rgba(232, 65, 24, 0.1); }
        .status-permit { color: #9c88ff; border: 1px solid #9c88ff; background: rgba(156, 136, 255, 0.1); }

        /* === ALERTS === */
        .anomaly-box { padding: 10px; margin-bottom: 8px; border-left: 4px solid; background: rgba(255,255,255,0.03); font-family: 'JetBrains Mono', monospace; font-size: 0.85rem; display: flex; justify-content: space-between; }
        .box-telat { border-color: var(--alert-red); background: linear-gradient(90deg, rgba(232, 65, 24, 0.1), transparent); }
        .box-izin { border-color: #9c88ff; background: linear-gradient(90deg, rgba(156, 136, 255, 0.1), transparent); }
        .box-alpha { border-color: var(--warning-amber); background: linear-gradient(90deg, rgba(251, 197, 49, 0.1), transparent); }

        /* === COMPONENTS === */
        .stDownloadButton button { background-color: var(--accent-blue) !important; color: #000 !important; font-weight: 800 !important; border: none !important; }
        .stTextInput input, .stDateInput input { background-color: #0f1218 !important; color: white !important; border: 1px solid var(--border-color) !important; }
        
        /* EXPANDER STYLING (Merged from Code 2 logic but adapted) */
        .streamlit-expanderHeader {
            background-color: rgba(255, 255, 255, 0.05);
            color: var(--accent-blue) !important;
            font-family: 'Rajdhani', sans-serif !important;
            font-weight: 600;
            border-radius: 4px;
        }
        div[data-testid="stDataFrame"] {
            background-color: rgba(21, 25, 34, 0.4);
            border: 1px solid var(--border-color);
        }

        /* Radio Button Styling */
        div[role="radiogroup"] { background: rgba(255,255,255,0.05); padding: 10px; border-radius: 8px; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 4. DATA FUNCTIONS
# ==========================================
@st.cache_data(ttl=10)
def load_absen(url):
    try:
        df = pd.read_csv(url)
        df.columns = df.columns.str.strip()
        if Config.COL_NAMA not in df.columns: return None
        df[Config.COL_NAMA] = df[Config.COL_NAMA].astype(str).str.strip()
        df = df[df[Config.COL_NAMA] != ''].copy()
        df[Config.COL_TIMESTAMP] = pd.to_datetime(df[Config.COL_TIMESTAMP])
        df['Tanggal'] = df[Config.COL_TIMESTAMP].dt.date
        df['Waktu'] = df[Config.COL_TIMESTAMP].dt.time
        return df
    except: return None

@st.cache_data(ttl=10)
def load_status(url):
    try:
        df = pd.read_csv(url)
        df = df.rename(columns=lambda x: x.strip())
        df['Nama Karyawan'] = df['Nama Karyawan'].astype(str).str.strip()
        df['Tanggal'] = pd.to_datetime(df['Tanggal'], format='mixed', dayfirst=False, errors='coerce')
        df['Tanggal_Str'] = df['Tanggal'].dt.strftime('%Y-%m-%d')
        df['Keterangan'] = df['Keterangan'].astype(str).str.strip().str.upper()
        return df
    except: return None

def get_min_time_in_range(group, s, e):
    st_t = time.fromisoformat(s)
    end_t = time.fromisoformat(e)
    filtered = group[(group['Waktu'] >= st_t) & (group['Waktu'] <= end_t)]
    return filtered[Config.COL_TIMESTAMP].min().strftime('%H:%M') if not filtered.empty else None

def is_late(time_str):
    if not time_str: return False
    try: return datetime.strptime(time_str, '%H:%M').time() > Config.LATE_THRESHOLD
    except: return False

def get_division_info(employee_name):
    for div_name, data in Config.DATA_DIVISI.items():
        if employee_name in data['members']:
            return div_name, data['code'], data['color']
    return "UNKNOWN", "UNK", "#666"

def get_all_employees():
    emps = []
    for data in Config.DATA_DIVISI.values(): emps.extend(data['members'])
    return emps

def create_excel_report(df_final, status_today, sel_date):
    out = io.BytesIO()
    wb = xlsxwriter.Workbook(out, {'in_memory': True})
    ws = wb.add_worksheet('Log')
    fmt_head = wb.add_format({'bold': True, 'bg_color': '#2f3542', 'font_color': 'white', 'border': 1})
    fmt_norm = wb.add_format({'border': 1})
    fmt_late = wb.add_format({'font_color': 'red', 'bold': True, 'border': 1})
    
    ws.write_row(0, 0, ['No', 'Nama', 'Masuk', 'Istirahat 1', 'Istirahat 2', 'Pulang', 'Ket'], fmt_head)
    
    for i, row in df_final.iterrows():
        nm = row['Nama Karyawan']
        stat = status_today.get(nm, "")
        ws.write(i+1, 0, i+1, fmt_norm)
        ws.write(i+1, 1, nm, fmt_norm)
        ws.write(i+1, 6, stat, fmt_norm)
        
        if stat: 
            for c in range(2, 6): ws.write(i+1, c, "IJIN", fmt_norm)
        else:
            pagi = row['Pagi']
            if is_late(pagi): ws.write(i+1, 2, pagi, fmt_late)
            else: ws.write(i+1, 2, pagi, fmt_norm)
            ws.write(i+1, 3, row['Siang_1'], fmt_norm)
            ws.write(i+1, 4, row['Siang_2'], fmt_norm)
            ws.write(i+1, 5, row['Sore'], fmt_norm)
            
    wb.close()
    out.seek(0)
    return out

# ==========================================
# 5. UI COMPONENTS
# ==========================================
def render_card(row, status_today):
    nm = row['Nama Karyawan']
    pagi = row['Pagi']
    siang1 = row['Siang_1']
    siang2 = row['Siang_2']
    sore = row['Sore']
    
    div_name, div_code, div_color = get_division_info(nm)
    manual_stat = status_today.get(nm, "")
    
    # Logic Status & Warna
    times = [pagi, siang1, siang2, sore]
    empty = sum(1 for t in times if t == '')
    
    if manual_stat:
        cls, txt, border_color = "status-permit", f"PERMIT: {manual_stat}", "#9c88ff"
    elif empty == 4:
        cls, txt, border_color = "status-absent", "ABSENT / NO SHOW", "#e84118"
    elif empty > 0:
        cls, txt, border_color = "status-partial", "PARTIAL DUTY", "#fbc531"
    else:
        cls, txt, border_color = "status-present", "FULL DUTY", "#4cd137"

    # Late Indicator
    is_person_late = pagi and is_late(pagi)
    late_html = "<span style='color:#e84118; font-weight:bold;'>(DELAY)</span>" if is_person_late else ""
    
    # Avatar
    avt = f"https://ui-avatars.com/api/?name={nm.replace(' ','+')}&background=random&color=fff&size=128&bold=true"

    # --- TAMPILAN KARTU UTAMA (BOARDING PASS STYLE) ---
    st.markdown(f"""
    <div class="card" style="border-left: 4px solid {border_color};">
        <div class="card-header">
            <img src="{avt}" style="width:45px;height:45px;border-radius:4px;margin-right:12px;border:2px solid {div_color}">
            <div style="flex:1; min-width:0;">
                <p class="card-name" style="font-size:1rem; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">{nm}</p>
                <span class="card-code" style="color:{div_color}; font-weight:bold;">✈ {div_code}</span>
                <span class="card-code" style="color:#888;"> // {div_name}</span>
            </div>
        </div>
        <div class="card-body" style="padding: 10px 15px;">
            <div class="flight-row">
                <span class="flight-label">ARRIVAL</span>
                <span class="flight-value" style="font-size:1.1rem;">{pagi if pagi else '--:--'} {late_html}</span>
            </div>
            <div class="flight-row">
                <span class="flight-label">DEPARTURE</span>
                <span class="flight-value" style="font-size:1.1rem;">{sore if sore else '--:--'}</span>
            </div>
            <div style="margin-top:8px; text-align:right;">
                <span class="status-badge {cls}">{txt}</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # --- TOMBOL DETAIL LOG (POPOVER) ---
    with st.popover("📋 VIEW LOG", use_container_width=True):
        st.markdown(f"### ✈️ FLIGHT RECORD: {nm}")
        st.markdown(f"**STATUS:** <span style='color:{border_color}; font-weight:bold'>{txt}</span>", unsafe_allow_html=True)
        st.divider()
        
        # Grid Layout untuk Detail Waktu
        c1, c2 = st.columns(2)
        
        with c1:
            st.markdown("#### 🛫 Waktu Sampai")
            val_pagi = pagi if pagi else "❌ --:--"
            if is_person_late: val_pagi += " (LATE)"
            st.info(f"**CHECK-IN:** {val_pagi}")
            
            val_s1 = siang1 if siang1 else "❌ --:--"
            st.write(f"**BREAK OUT:** {val_s1}")

        with c2:
            st.markdown("#### 🛬 Waktu Pulang")
            val_s2 = siang2 if siang2 else "❌ --:--"
            st.write(f"**BREAK IN:** {val_s2}")
            
            val_sore = sore if sore else "❌ --:--"
            st.success(f"**CHECK-OUT:** {val_sore}")

        # Kalkulasi Durasi Kerja (Estimasi Kasar)
        if pagi and sore:
            try:
                t_start = datetime.strptime(pagi, "%H:%M")
                t_end = datetime.strptime(sore, "%H:%M")
                duration = t_end - t_start
                hours = duration.seconds // 3600
                minutes = (duration.seconds % 3600) // 60
                st.divider()
                st.markdown(f"⏱️ **TOTAL FLIGHT TIME:** {hours} Jam {minutes} Menit")
            except:
                pass

# ==========================================
# 6. MAIN APP
# ==========================================
def main():
    apply_custom_css()
    
    # --- SIDEBAR (SIMPLE SESUAI REQUEST) ---
    with st.sidebar:
        st.markdown("### ⚡ WedaBayAirport")
        menu = st.radio("NAVIGASI:", ["📊 Dashboard Monitoring", "📝 Input Laporan"])
        
        st.markdown("---")
        if st.button("🔄 REFRESH DATA"):
            st.cache_data.clear()
            st.rerun()
    
    # --- HALAMAN 1: DASHBOARD (TAMPILAN MAHAL) ---
    if menu == "📊 Dashboard Monitoring":
        st.markdown('<div class="brand-title">WedaBayAirport</div>', unsafe_allow_html=True)
        st.markdown('<div class="brand-subtitle">REAL-TIME PERSONNEL MONITORING SYSTEM</div>', unsafe_allow_html=True)
        
        df_absen = load_absen(Config.SHEET_URL_ABSEN)
        df_status = load_status(Config.SHEET_URL_STATUS)
        
        if df_absen is None:
            st.error("⚠️ SYSTEM OFFLINE. Check Database Connection.")
            st.stop()
            
        # Filters
        c1, c2 = st.columns([1, 3])
        with c1:
            avail_dates = sorted(df_absen['Tanggal'].unique())
            sel_date = st.date_input("📅 OPERATION DATE", value=avail_dates[-1] if avail_dates else datetime.now())
        with c2:
            search_q = st.text_input("🔍 RADAR SEARCH", placeholder="Search crew name...")
            
        st.markdown("---")
        
        # Processing
        if sel_date:
            df_today = df_absen[df_absen['Tanggal'] == sel_date]
            status_today = {}
            
            if df_status is not None:
                d_str = sel_date.strftime('%Y-%m-%d')
                if 'Tanggal_Str' not in df_status.columns: df_status['Tanggal_Str'] = df_status['Tanggal'].astype(str)
                df_s = df_status[df_status['Tanggal_Str'] == d_str]
                if not df_s.empty:
                    status_today = pd.Series(df_s.Keterangan.values, index=df_s['Nama Karyawan']).to_dict()
            
            # Pivot Data
            recap = {}
            if not df_today.empty:
                g = df_today.groupby([Config.COL_NAMA, 'Tanggal'])
                for k, (s, e) in Config.RENTANG_WAKTU.items():
                    recap[k] = g.apply(lambda x: get_min_time_in_range(x, s, e))
            
            df_res = pd.DataFrame(recap).reset_index() if recap else pd.DataFrame(columns=[Config.COL_NAMA])
            if not df_res.empty: df_res.rename(columns={Config.COL_NAMA: 'Nama Karyawan'}, inplace=True)
            
            all_emps = get_all_employees()
            df_final = pd.merge(pd.DataFrame({'Nama Karyawan': all_emps}), df_res, on='Nama Karyawan', how='left')
            df_final.fillna('', inplace=True)
            
            # Metrics
            total = len(df_final)
            hadir = 0; izin = 0; bolos = 0
            l_telat = []; l_izin = []; l_bolos = []
            
            for idx, row in df_final.iterrows():
                nm = row['Nama Karyawan']
                pagi = row['Pagi']
                manual = status_today.get(nm, "")
                empty = sum(1 for t in [pagi, row['Siang_1'], row['Siang_2'], row['Sore']] if t == '')
                
                if manual: 
                    izin += 1; l_izin.append((nm, manual))
                elif empty == 4: 
                    bolos += 1; l_bolos.append(nm)
                else:
                    if pagi and is_late(pagi): l_telat.append((nm, pagi))
                    hadir += 1
            
            # Render Metrics
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("TOTAL CREW", total)
            m2.metric("ON DUTY", hadir, f"{int(hadir/total*100)}%")
            m3.metric("PERMIT", izin)
            m4.metric("ABSENT", bolos, delta_color="inverse")
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            # Anomalies
            c1, c2, c3 = st.columns(3)
            with c1: 
                st.markdown("#### ⚠️ DELAYED")
                if l_telat:
                    with st.container(height=150):
                        for n, t in l_telat: st.markdown(f"<div class='anomaly-box box-telat'><span>{n}</span><span>{t}</span></div>", unsafe_allow_html=True)
                else: st.success("ON TIME")
            with c2:
                st.markdown("#### ℹ️ PERMIT")
                if l_izin:
                    with st.container(height=150):
                        for n, t in l_izin: st.markdown(f"<div class='anomaly-box box-izin'><span>{n}</span><span>{t}</span></div>", unsafe_allow_html=True)
                else: st.info("NONE")
            with c3:
                st.markdown("#### ❌ ALPHA")
                if l_bolos:
                    with st.container(height=150):
                        for n in l_bolos: st.markdown(f"<div class='anomaly-box box-alpha'><span>{n}</span><span>N/A</span></div>", unsafe_allow_html=True)
                else: st.success("FULL")
            
            # Excel
            st.markdown("---")
            excel = create_excel_report(df_final, status_today, sel_date)
            st.download_button("📥 DOWNLOAD EXCEL LOG", excel, f"Log_{sel_date}.xlsx", use_container_width=True)
            
            # --- FITUR BARU (MERGED) ---
            # Dropdown detail absen karyawan (Diminta user)
            with st.expander("👇 Tampilan Semua detail absen karyawan", expanded=False):
                # Styling untuk tabel agar terlihat rapi dan "Mahal"
                def highlight_late(val):
                    color = 'white'
                    if isinstance(val, str) and ':' in val:
                        # Simple check if standard time string is late
                        try: 
                            t = datetime.strptime(val, '%H:%M').time()
                            if t > Config.LATE_THRESHOLD: color = '#e84118' # Red from CSS
                        except: pass
                    return f'color: {color}'
                
                # Siapkan data untuk ditampilkan
                df_display = df_final.copy()
                
                # Rename kolom agar cantik di tabel
                df_display.rename(columns={
                    'Siang_1': 'Istirahat (Out)', 
                    'Siang_2': 'Istirahat (In)',
                    'Sore': 'Pulang',
                    'Pagi': 'Masuk'
                }, inplace=True)
                
                # Tambahkan kolom Keterangan
                df_display['Keterangan'] = df_display['Nama Karyawan'].apply(lambda x: status_today.get(x, ""))
                
                # Tampilkan DataFrame dengan style
                st.dataframe(
                    df_display[['Nama Karyawan', 'Masuk', 'Istirahat (Out)', 'Istirahat (In)', 'Pulang', 'Keterangan']].style.applymap(highlight_late, subset=['Masuk']),
                    use_container_width=True,
                    height=500,
                    hide_index=True
                )
            # --- END FITUR BARU ---

            # Grid Cards
            st.markdown("### 📋 CREW AIRPORT")
            divs = list(Config.DATA_DIVISI.keys())
            tabs = st.tabs(divs)
            
            for tab, d in zip(tabs, divs):
                with tab:
                    mems = Config.DATA_DIVISI[d]['members']
                    df_d = df_final[df_final['Nama Karyawan'].isin(mems)]
                    if search_q: df_d = df_d[df_d['Nama Karyawan'].str.contains(search_q, case=False)]
                    
                    if not df_d.empty:
                        cols = st.columns(4)
                        for i, (ix, r) in enumerate(df_d.iterrows()):
                            with cols[i%4]: render_card(r, status_today)
                    else: st.info("No Personnel Found")

    # --- HALAMAN 2: INPUT (TAMPILAN SIMPLE) ---
    elif menu == "📝 Input Laporan":
        st.markdown('<div class="brand-title">MANUAL REPORT</div>', unsafe_allow_html=True)
        st.markdown('<div class="brand-subtitle">SUBMIT PERMIT / SICK LEAVE</div>', unsafe_allow_html=True)
        
        if "PASTE_LINK" in Config.GOOGLE_FORM_URL:
            st.warning("Google Form URL not set.")
        else:
            components.iframe(Config.GOOGLE_FORM_URL, height=1000, scrolling=True)

if __name__ == "__main__":
    main()