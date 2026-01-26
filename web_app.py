"""
================================================================================
WEDABAY AIRPORT ABSENCE CENTER - ENTERPRISE PERSONNEL MONITORING SYSTEM
================================================================================
Version: 2.0.1 (Image Replication Build - Fixed)
Architecture: Clean Architecture with MVC Pattern
================================================================================
"""

import streamlit as st
import pandas as pd
from datetime import time, datetime, timedelta, date
import io
import xlsxwriter
import streamlit.components.v1 as components
from typing import Dict, List, Tuple, Optional, Any
from dataclasses import dataclass, field
from enum import Enum
from abc import ABC, abstractmethod
import json
import hashlib
from functools import lru_cache
import plotly.express as px
import plotly.graph_objects as go
from PIL import Image
import base64

# ================================================================================
# SECTION 1: CONFIGURATION & CONSTANTS LAYER
# ================================================================================

class AppConstants:
    """
    Global application constants.
    Updated for DUAL SHIFT (Strict 19:00 Return & 08:15 Cutoff).
    """
    APP_TITLE = "WedaBay Airport Absence Center"
    APP_VERSION = "2.3.0 (Dual Shift Final)"
    APP_ICON = "✈️"
    COMPANY_NAME = "WedaBay Aviation Services"
    
    # Column Names
    COL_PERSON_NAME = 'Person Name'
    COL_EVENT_TIME = 'Event Time'
    COL_EMPLOYEE_NAME = 'Nama Karyawan'
    COL_DATE = 'Tanggal'
    COL_STATUS = 'Keterangan'
    
    # --- LOGIC CONSTANTS ---
    
    # PEMISAH SHIFT (Jantung Logika)
    SHIFT_CUTOFF = time(8, 15, 0)  # <= 08:15 Shift 1, > 08:15 Shift 2
    
    # ATURAN SHIFT 1 (07:00 - 17:00)
    S1_LATE_TOLERANCE = time(7, 5, 0)      # Lewat 07:05 = Merah
    S1_BREAK_OUT_START = time(12, 0, 0)
    S1_BREAK_OUT_END   = time(12, 59, 59)
    S1_BREAK_IN_START  = time(13, 0, 0)
    S1_BREAK_IN_END    = time(14, 0, 0)    # Lewat 14:00 = Merah
    S1_HOME_TIME       = time(17, 0, 0)    # Pulang

    # ATURAN SHIFT 2 (09:00 - 19:00)
    S2_LATE_TOLERANCE  = time(9, 5, 0)     # Lewat 09:05 = Merah
    S2_HOME_TIME       = time(19, 0, 0)    # Pulang (Strict)
    
    # Istirahat Shift 2 (Normal: Senin-Kamis, Sabtu-Minggu)
    S2_NORM_BREAK_OUT_START = time(14, 0, 0)
    S2_NORM_BREAK_OUT_END   = time(14, 59, 59)
    S2_NORM_BREAK_IN_START  = time(15, 0, 0)
    S2_NORM_BREAK_IN_END    = time(16, 0, 0) # Lewat 16:00 = Merah

    # Note: Hari Jumat Shift 2 ikut jam istirahat Shift 1
    
    # System Config
    CACHE_TTL_SECONDS = 10
    MAX_CACHE_ENTRIES = 100
    SIDEBAR_STATE = "expanded"
    LAYOUT_MODE = "wide"
    CARDS_PER_ROW = 4
    AUTO_REFRESH_INTERVAL = 30
    EXCEL_ENGINE = 'xlsxwriter'
    DATE_FORMAT = '%Y-%m-%d'
    TIME_FORMAT = '%H:%M'
    DATETIME_FORMAT = '%Y-%m-%d %H:%M:%S'


class TimeRanges(Enum):
    # ... (Pagi tetap sama)
    MORNING = ('Pagi', '03:00:00', '9:00:00')

    # PERBAIKAN: Batasi Siang 1 sampai 12:35:00
    # Ini akan menangkap Yogi (11:34) dan teman yg telat istirahat (misal 12:31)
    BREAK_OUT = ('Siang_1', '11:29:00', '12:29:00')

    # Mulai Siang 2 dari 12:35:01
    # Ini memastikan Yogi (12:38) masuk ke sini (Siang 2)
    BREAK_IN = ('Siang_2', '12:31:00', '15:00:00')
    
    # ... (Sore tetap sama)
    EVENING = ('Sore', '17:00:00', '23:59:59')

    def __init__(self, label: str, start: str, end: str):
        self.label = label
        self.start_time = start
        self.end_time = end

    @property
    def time_window(self) -> Tuple[str, str]:
        """Returns the time window as a tuple."""
        return (self.start_time, self.end_time)
    


class AttendanceStatus(Enum):
    """
    Enumeration for attendance status types with styling metadata.
    Follows Single Responsibility Principle for status management.
    """
    FULL_DUTY = ("FULL DUTY", "#4cd137", "status-present")
    PARTIAL_DUTY = ("PARTIAL DUTY", "#fbc531", "status-partial")
    ABSENT = ("ABSENT / NO SHOW", "#e84118", "status-absent")
    PERMIT = ("PERMIT", "#9c88ff", "status-permit")
    LATE = ("DELAYED", "#e84118", "status-late")
    
    def __init__(self, display_text: str, color: str, css_class: str):
        self.display_text = display_text
        self.color = color
        self.css_class = css_class


@dataclass
class DivisionConfig:
    """
    Data class representing a division's configuration.
    Immutable configuration following Data Transfer Object pattern.
    """
    name: str
    color: str
    icon: str
    code: str
    members: List[str] = field(default_factory=list)
    description: str = ""
    priority: int = 0
    
    def __hash__(self):
        """Make dataclass hashable for caching purposes."""
        return hash((self.name, self.code))


class DataSourceConfig:
    """
    Configuration for external data sources.
    Centralizes all external URLs and connection parameters.
    """
    
    # Google Sheets Data Sources
    ATTENDANCE_SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ_GhoIb1riX98FsP8W4f2-_dH_PLcLDZskjNOyDcnnvOhBg8FUp3xJ-c_YgV0Pw71k4STy4rR0_MS5/pub?output=csv"
    
    STATUS_SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ2QrBN8uTRiHINCEcZBrbdU-gzJ4pN2UljoqYG6NMoUQIK02yj_D1EdlxdPr82Pbr94v2o6V0Vh3Kt/pub?gid=511860805&single=true&output=csv"
    
    # Google Forms
    REPORT_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSeopdaE-lyOtFd2TUr5C3K2DWE3Syt2PaKoXMp0cmWKIFnijw/viewform?usp=header"
    
    # API Endpoints (for future expansion)
    API_BASE_URL = "https://api.wedabay.airport/v1"
    WEBSOCKET_URL = "wss://ws.wedabay.airport/notifications"


class DivisionRegistry:
    """
    Registry pattern for managing division configurations.
    Provides centralized access to all division data with validation.
    """
    
    _divisions: Dict[str, DivisionConfig] = {}
    
    @classmethod
    def register(cls, division: DivisionConfig) -> None:
        """Register a new division configuration."""
        if division.name in cls._divisions:
            raise ValueError(f"Division {division.name} already registered")
        cls._divisions[division.name] = division
    
    @classmethod
    def get(cls, name: str) -> Optional[DivisionConfig]:
        """Retrieve division configuration by name."""
        return cls._divisions.get(name)
    
    @classmethod
    def get_all(cls) -> Dict[str, DivisionConfig]:
        """Get all registered divisions."""
        return cls._divisions.copy()
    
    @classmethod
    def find_by_member(cls, member_name: str) -> Optional[DivisionConfig]:
        """Find division by member name."""
        for division in cls._divisions.values():
            if member_name in division.members:
                return division
        return None
    
    @classmethod
    def get_all_members(cls) -> List[str]:
        """Get all employee names across all divisions."""
        members = []
        
        # 1. Pastikan urutan divisi berdasarkan Priority (1, 2, 3...)
        # Ini supaya SPT & SPV (Patra, Su Adam) selalu paling atas
        sorted_divisions = sorted(cls._divisions.values(), key=lambda x: x.priority)
        
        for division in sorted_divisions:
            members.extend(division.members)
            
        # 2. Hapus nama ganda TAPI JANGAN di-sort A-Z. 
        # Kita pakai trik 'dict.fromkeys' agar urutan Patra, Su Adam, dst tetap terjaga.
        return list(dict.fromkeys(members))


# Initialize Division Registry with actual data
def initialize_divisions():
    """
    Initialize all division configurations with CUSTOM SORT ORDER defined by User.
    """
    divisions_data = [
        DivisionConfig("SPT & SPV", "#FFD700", "👨‍✈️", "SPT & SPV", description="SPT & SPV", priority=1, 
                       members=["Patra Anggana", "Su Adam", "Budiman Arifin", "Rifaldy Ilham Bhagaskara", "Marwan S Halid", "Budiono"]),
        
        DivisionConfig("TLB", "#00A8FF", "🔧", "TLB", description="Teknik Listrik Bandara", priority=2, 
                       members=["M. Ansori", "Bayu Pratama Putra Katuwu", "Yoga Nugraha Putra Pasaribu", "Junaidi Taib", "Muhammad Rizal Amra", "Rusli Dj"]),
        
        DivisionConfig("TBL", "#0097e6", "📦", "TBL", description="Teknik Bangunan Dan Landasan", priority=3, 
                       members=["Venesia Aprilia Ineke", "Muhammad Naufal Ramadhan", "Yuzak Gerson Puturuhu", "Muhamad Alief Wildan", "Gafur Hamisi", "Jul Akbar M. Nur", "Sarni Massiri", "Adrianto Laundang", "Wahyudi Ismail"]),
        
        DivisionConfig("TRANS APRON", "#e1b12c", "🚌", "APR", description="Trans Apron", priority=4, 
                       members=["Marichi Gita Rusdi", "Ilham Rahim", "Abdul Mu Iz Simal", "Dwiki Agus Saputro", "Moh. Sofyan", "Faisal M. Kadir", "Amirudin Rustam", "Faturrahman Kaunar", "Wawan Hermawan", "Rahmat Joni", "Nur Ichsan"]),
        
        DivisionConfig("ATS", "#44bd32", "📡", "ATS", description="Air Traffic Services", priority=5, 
                       members=["Nurul Tanti", "Firlon Paembong", "Irwan Rezky Setiawan", "Yusuf Arviansyah", "Nurdahlia Is. Folaimam", "Ghaly Rabbani Panji Indra", "Ikhsan Wahyu Vebriyan", "Rizki Mahardhika Ardi Tigo", "Nikolaus Vincent Quirino"]),
        
        DivisionConfig("ADM COMPLIANCE", "#8c7ae6", "📋", "ADM", description="Administration & Compliance", priority=6, 
                       members=["Yessicha Aprilyona Siregar", "Gabriela Margrith Louisa Klavert", "Aldi Saptono"]),
        
        DivisionConfig("TRANSLATOR", "#00cec9", "🎧", "TRN", description="Translation Services", priority=7, 
                       members=["Wilyam Candra", "Norika Joselyn Modnissa"]),
        
        DivisionConfig("AVSEC", "#c23616", "🛡️", "SEC", description="Aviation Security", priority=8, 
                       members=["Andrian Maranatha", "Toni Nugroho Simarmata", "Muhamad Albi Ferano", "Andreas Charol Tandjung", "Sabadia Mahmud", "Rusdin Malagapi", "Muhamad Judhytia Winli", "Wahyu Samsudin", "Fientje Elisabeth Joseph", "Anglie Fitria Desiana Mamengko", "Dwi Purnama Bimasakti", "Windi Angriani Sulaeman", "Megawati A. Rauf"]),
        
        DivisionConfig("GROUND HANDLING", "#e17055", "🚜", "GND", description="Ground Handling Operations", priority=9, 
                       members=["Yuda Saputra", "Tesalonika Gratia Putri Toar", "Esi Setia Ningseh", "Ardiyanto Kalatjo", "Febrianti Tikabala"]),
        
        DivisionConfig("HELICOPTER", "#6c5ce7", "🚁", "HEL", description="Helicopter Operations", priority=10, 
                       members=["Agung Sabar S. Taufik", "Recky Irwan R. A Arsyad", "Farok Abdul", "Achmad Rizky Ariz", "Yus Andi", "Muh. Noval Kipudjena"]),
        
        DivisionConfig("AMC & TERMINAL", "#0984e3", "🏢", "AMC", description="Airport Movement Control & Terminal", priority=11, 
                       members=["Risky Sulung", "Muchamad Nur Syaifulrahman", "Muhammad Tunjung Rohmatullah", "Sunarty Fakir", "Albert Papuling", "Gibhran Fitransyah Yusri", "Muhdi R Tomia", "Riski Rifaldo Theofilus Anu", "Eko"]),
        
        DivisionConfig("SAFETY OFFICER", "#fd79a8", "🦺", "SFT", description="Safety Operations", priority=12, 
                       members=["Hildan Ahmad Zaelani", "Abdurahim Andar"]),
        
        DivisionConfig("PKP-PK", "#fab1a0", "🚒", "RES", description="Fire & Rescue Services", priority=13, 
                       members=["Andreas Aritonang", "Achmad Alwan Asyhab", "Doni Eka Satria", "Yogi Prasetya Eka Winandra", "Akhsin Aditya Weza Putra", "Fardhan Ahmad Tajali", "Maikel Renato Syafaruddin", "Saldi Sandra", "Hamzah M. Ali Gani", "Marfan Mandar", "Julham Keya", "Aditya Sugiantoro Abbas", "Muhamad Usman", "M Akbar D Patty", "Daniel Freski Wangka", "Fandi M.Naser", "Agung Fadjriansyah Ano", "Deni Hendri Bobode", "Muhammad Rifai", "Idrus Arsad, SH"])
    ]
    for division in divisions_data:
        DivisionRegistry.register(division)

# ================================================================================
# SECTION 2: DATA ACCESS LAYER (REPOSITORY PATTERN)
# ================================================================================

class DataRepository(ABC):
    """
    Abstract base class for data repositories.
    Follows Repository Pattern for data access abstraction.
    """
    
    @abstractmethod
    def fetch(self) -> Optional[pd.DataFrame]:
        """Fetch data from source."""
        pass
    
    @abstractmethod
    def validate(self, df: pd.DataFrame) -> bool:
        """Validate fetched data."""
        pass
    
    @abstractmethod
    def transform(self, df: pd.DataFrame) -> pd.DataFrame:
        """Transform raw data to application format."""
        pass


class AttendanceRepository(DataRepository):
    """
    Repository for attendance data management.
    Handles data fetching, validation, and transformation.
    """
    
    def __init__(self, url: str):
        self.url = url
        self._cache: Optional[pd.DataFrame] = None
        self._cache_time: Optional[datetime] = None
    
    @st.cache_data(ttl=AppConstants.CACHE_TTL_SECONDS)
    def fetch(_self) -> Optional[pd.DataFrame]:
        """
        Fetch attendance data from Google Sheets.
        Uses Streamlit caching for performance optimization.
        """
        try:
            df = pd.read_csv(_self.url)
            
            # Standardize column names
            df.columns = df.columns.str.strip()
            
            if not _self.validate(df):
                st.error("❌ Attendance data validation failed")
                return None
            
            return _self.transform(df)
            
        except Exception as e:
            st.error(f"❌ Failed to fetch attendance data: {str(e)}")
            return None
    
    def validate(self, df: pd.DataFrame) -> bool:
        """Validate that required columns exist."""
        required_columns = [AppConstants.COL_PERSON_NAME, AppConstants.COL_EVENT_TIME]
        return all(col in df.columns for col in required_columns)
    
    def transform(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Transform raw attendance data.
        Applies cleaning, type conversion, and feature engineering.
        """
        # Clean employee names
        df[AppConstants.COL_PERSON_NAME] = (
            df[AppConstants.COL_PERSON_NAME]
            .astype(str)
            .str.strip()
        )
        
        # Remove empty records
        df = df[df[AppConstants.COL_PERSON_NAME] != ''].copy()
        
        # Parse timestamps
        df[AppConstants.COL_EVENT_TIME] = pd.to_datetime(
            df[AppConstants.COL_EVENT_TIME],
            errors='coerce'
        )
        
        # Extract date and time components
        df['Tanggal'] = df[AppConstants.COL_EVENT_TIME].dt.date
        df['Waktu'] = df[AppConstants.COL_EVENT_TIME].dt.time
        df['Jam'] = df[AppConstants.COL_EVENT_TIME].dt.hour
        df['Menit'] = df[AppConstants.COL_EVENT_TIME].dt.minute
        
        # Add day of week
        df['Hari'] = df[AppConstants.COL_EVENT_TIME].dt.day_name()
        
        return df


class StatusRepository(DataRepository):
    """
    Repository for employee status (permits, leaves) management.
    """
    
    def __init__(self, url: str):
        self.url = url
    
    @st.cache_data(ttl=AppConstants.CACHE_TTL_SECONDS)
    def fetch(_self) -> Optional[pd.DataFrame]:
        """Fetch status data from Google Sheets."""
        try:
            df = pd.read_csv(_self.url)
            df = df.rename(columns=lambda x: x.strip())
            
            if not _self.validate(df):
                st.warning("⚠️ Status data validation failed")
                return None
            
            return _self.transform(df)
            
        except Exception as e:
            st.warning(f"⚠️ Failed to fetch status data: {str(e)}")
            return None
    
    def validate(self, df: pd.DataFrame) -> bool:
        """Validate status data structure."""
        required = [AppConstants.COL_EMPLOYEE_NAME, AppConstants.COL_DATE, AppConstants.COL_STATUS]
        return all(col in df.columns for col in required)
    
    def transform(self, df: pd.DataFrame) -> pd.DataFrame:
        """Transform status data."""
        # Clean names
        df[AppConstants.COL_EMPLOYEE_NAME] = (
            df[AppConstants.COL_EMPLOYEE_NAME]
            .astype(str)
            .str.strip()
        )
        
        # Parse dates
        df[AppConstants.COL_DATE] = pd.to_datetime(
            df[AppConstants.COL_DATE],
            format='mixed',
            dayfirst=False,
            errors='coerce'
        ).dt.date # Ensure date format consistency
        
        df['Tanggal_Str'] = pd.to_datetime(df[AppConstants.COL_DATE]).dt.strftime(AppConstants.DATE_FORMAT)
        
        # Standardize status text
        df[AppConstants.COL_STATUS] = (
            df[AppConstants.COL_STATUS]
            .astype(str)
            .str.strip()
            .str.upper()
        )
        
        return df


# ================================================================================
# SECTION 3: BUSINESS LOGIC LAYER (SERVICE CLASSES)
# ================================================================================

class TimeService:
    """
    Service class for time-related business logic.
    """
    
    @staticmethod
    def is_late(time_str: Optional[str]) -> Tuple[bool, str]:
        """
        Check late arrival based on Dual Shift logic (Cutoff 08:15).
        Returns: (is_late_bool, shift_label)
        """
        if not time_str:
            return False, ""
        
        try:
            check_time = datetime.strptime(time_str, AppConstants.TIME_FORMAT).time()
            
            # LOGIKA UTAMA: Tentukan Shift berdasarkan jam datang
            if check_time <= AppConstants.SHIFT_CUTOFF:
                # SHIFT 1
                is_late_val = check_time > AppConstants.S1_LATE_TOLERANCE # > 07:05
                return is_late_val, "SHIFT 1"
            else:
                # SHIFT 2
                is_late_val = check_time > AppConstants.S2_LATE_TOLERANCE # > 09:05
                return is_late_val, "SHIFT 2"

        except (ValueError, TypeError):
            return False, ""
    
    @staticmethod
    def calculate_duration(start_time: str, end_time: str) -> Optional[timedelta]:
        try:
            start = datetime.strptime(start_time, AppConstants.TIME_FORMAT)
            end = datetime.strptime(end_time, AppConstants.TIME_FORMAT)
            if end < start: end += timedelta(days=1)
            return end - start
        except (ValueError, TypeError):
            return None
    
    @staticmethod
    def format_duration(duration: timedelta) -> str:
        if not duration: return "N/A"
        hours = duration.seconds // 3600
        minutes = (duration.seconds % 3600) // 60
        return f"{hours}h {minutes}m"
    
    @staticmethod
    def get_time_range_label(check_time: time) -> str:
        """Determine which time range a given time falls into."""
        for time_range in TimeRanges:
            start = time.fromisoformat(time_range.start_time)
            end = time.fromisoformat(time_range.end_time)
            
            if start <= check_time <= end:
                return time_range.label
        
        return "UNKNOWN"


class AttendanceService:
    """
    Core business logic service for attendance processing.
    """
    
    def __init__(self, attendance_repo: AttendanceRepository, status_repo: StatusRepository):
        self.attendance_repo = attendance_repo
        self.status_repo = status_repo
        self.time_service = TimeService()
    
    def get_attendance_for_date(self, target_date: datetime.date) -> Optional[pd.DataFrame]:
        df = self.attendance_repo.fetch()
        if df is None: return None
        return df[df['Tanggal'] == target_date].copy()
    
    def get_status_for_date(self, target_date: datetime.date) -> Dict[str, str]:
        df = self.status_repo.fetch()
        if df is None: return {}
        date_str = target_date.strftime(AppConstants.DATE_FORMAT)
        df_filtered = df[df['Tanggal_Str'] == date_str]
        if df_filtered.empty: return {}
        return pd.Series(
            df_filtered[AppConstants.COL_STATUS].values,
            index=df_filtered[AppConstants.COL_EMPLOYEE_NAME]
        ).to_dict()
    
    def extract_time_ranges(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Extracts attendance into Pagi, Siang1, Siang2, Sore columns.
        Supports Dual Shift & Friday Exception.
        """
        if df.empty: return pd.DataFrame()

        df_clean = df.dropna(subset=[AppConstants.COL_PERSON_NAME, 'Tanggal'])
        if df_clean.empty: return pd.DataFrame()

        df_clean['Waktu_Obj'] = pd.to_datetime(df_clean[AppConstants.COL_EVENT_TIME]).dt.time
        grouped = df_clean.groupby([AppConstants.COL_PERSON_NAME, 'Tanggal'])
        
        def process_group(group):
            result = {'Pagi': '', 'Siang_1': '', 'Siang_2': '', 'Sore': ''}
            
            # Sortir waktu
            sorted_group = group.sort_values(AppConstants.COL_EVENT_TIME)
            if sorted_group.empty: return pd.Series(result)

            # Cek Hari Jumat (0=Senin, 4=Jumat)
            is_friday = sorted_group.iloc[0]['Tanggal'].weekday() == 4
            
            # --- DETEKSI SHIFT BERDASARKAN LOG PERTAMA ---
            first_log = sorted_group.iloc[0]['Waktu_Obj']
            
            # Default anggap Shift 1
            is_shift_2 = False
            
            # Jika log pertama > 08:15, maka dia Shift 2
            if first_log > AppConstants.SHIFT_CUTOFF:
                is_shift_2 = True
            
            # --- SETTING RANGE WAKTU BERDASARKAN SHIFT ---
            
            if not is_shift_2:
                # === SHIFT 1 ===
                # Pagi: < 08:15 (Sebenarnya bisa sampai 11:00 buat jaga2, tapi cutoff 08:15)
                limit_pagi_end = time(11, 0, 0) 
                
                # Istirahat (Shift 1 Sama Terus tiap hari)
                limit_siang_out_start = AppConstants.S1_BREAK_OUT_START # 12:00
                limit_siang_out_end   = AppConstants.S1_BREAK_OUT_END   # 12:59
                limit_siang_in_start  = AppConstants.S1_BREAK_IN_START  # 13:00
                limit_siang_in_end    = time(14, 59, 59) # Kita lebarin dikit capturenya
                
                start_sore = AppConstants.S1_HOME_TIME # 17:00
                
            else:
                # === SHIFT 2 ===
                limit_pagi_end = time(12, 0, 0) # Datang
                
                if is_friday:
                    # JUMAT: Ikut jam Shift 1
                    limit_siang_out_start = AppConstants.S1_BREAK_OUT_START # 12:00
                    limit_siang_out_end   = AppConstants.S1_BREAK_OUT_END   # 12:59
                    limit_siang_in_start  = AppConstants.S1_BREAK_IN_START  # 13:00
                    limit_siang_in_end    = time(14, 59, 59)
                else:
                    # NORMAL: Jam 14 - 16
                    limit_siang_out_start = AppConstants.S2_NORM_BREAK_OUT_START # 14:00
                    limit_siang_out_end   = AppConstants.S2_NORM_BREAK_OUT_END   # 14:59
                    limit_siang_in_start  = AppConstants.S2_NORM_BREAK_IN_START  # 15:00
                    limit_siang_in_end    = time(16, 59, 59) # Lebarin dikit capturenya

                start_sore = AppConstants.S2_HOME_TIME # 19:00 (STRICT)

            # --- MAPPING KE KOLOM ---
            for _, row in sorted_group.iterrows():
                t = row['Waktu_Obj']
                val_str = row[AppConstants.COL_EVENT_TIME].strftime(AppConstants.TIME_FORMAT)
                
                # 1. Pagi (Datang)
                if t < limit_pagi_end:
                    if result['Pagi'] == '': result['Pagi'] = val_str
                
                # 2. Siang 1 (Keluar Istirahat)
                elif limit_siang_out_start <= t <= limit_siang_out_end:
                    if result['Siang_1'] == '': result['Siang_1'] = val_str
                
                # 3. Siang 2 (Masuk Istirahat)
                elif limit_siang_in_start <= t <= limit_siang_in_end:
                    if result['Siang_2'] == '': result['Siang_2'] = val_str
                
                # 4. Sore (Pulang)
                elif t >= start_sore:
                    # Ambil yang paling terakhir
                    result['Sore'] = val_str

            return pd.Series(result)

        if grouped.ngroups == 0: return pd.DataFrame()
        result_df = grouped.apply(process_group).reset_index()
        result_df.rename(columns={AppConstants.COL_PERSON_NAME: AppConstants.COL_EMPLOYEE_NAME}, inplace=True)
        return result_df

    def build_complete_report(self, target_date: datetime.date) -> Tuple[pd.DataFrame, Dict[str, str]]:
        """
        Builds the master dataframe merging attendance times with employee list.
        """
        df_attendance = self.get_attendance_for_date(target_date)
        status_dict = self.get_status_for_date(target_date)
        
        if df_attendance is not None and not df_attendance.empty:
            df_times = self.extract_time_ranges(df_attendance)
        else:
            df_times = pd.DataFrame()
        
        all_employees = DivisionRegistry.get_all_members()
        df_all = pd.DataFrame({AppConstants.COL_EMPLOYEE_NAME: all_employees})
        
        if not df_times.empty:
            df_final = pd.merge(df_all, df_times, on=AppConstants.COL_EMPLOYEE_NAME, how='left')
        else:
            df_final = df_all.copy()
            for col in ['Pagi', 'Siang_1', 'Siang_2', 'Sore']:
                df_final[col] = ''

        df_final.fillna('', inplace=True)
        return df_final, status_dict

    def calculate_metrics(self, df: pd.DataFrame, status_dict: Dict[str, str]) -> Dict[str, Any]:
        """
        Calculates daily statistics (Present, Absent, Late, etc.)
        FIXED: Unpack is_late tuple correctly.
        """
        total_employees = len(df)
        present_count = 0; permit_count = 0; absent_count = 0; late_count = 0
        late_list = []; permit_list = []; absent_list = []; partial_list = []
        
        for idx, row in df.iterrows():
            name = row[AppConstants.COL_EMPLOYEE_NAME]
            times = [row.get('Pagi', ''), row.get('Siang_1', ''), 
                     row.get('Siang_2', ''), row.get('Sore', '')]
            empty_count = sum(1 for t in times if t == '')
            manual_status = status_dict.get(name, "")
            
            if manual_status:
                permit_count += 1
                permit_list.append((name, manual_status))
            elif empty_count == 4:
                absent_count += 1
                absent_list.append(name)
            else:
                present_count += 1
                morning_time = row.get('Pagi', '')
                
                # --- PERBAIKAN DISINI ---
                if morning_time:
                    # Bongkar paket tuple (is_late_status, shift_label)
                    is_late_status, _ = self.time_service.is_late(morning_time)
                    
                    if is_late_status: # Cek nilai boolean-nya saja
                        late_count += 1
                        late_list.append((name, morning_time))
                # ------------------------

                if empty_count > 0:
                    partial_list.append((name, empty_count))
        
        return {
            'total': total_employees, 'present': present_count, 'permit': permit_count,
            'absent': absent_count, 'late': late_count, 'late_list': late_list,
            'permit_list': permit_list, 'absent_list': absent_list, 'partial_list': partial_list,
            'attendance_rate': (present_count / total_employees * 100) if total_employees > 0 else 0,
            'punctuality_rate': ((present_count - late_count) / present_count * 100) if present_count > 0 else 0
        }
class AnalyticsService:
    """
    Service for advanced analytics and reporting.
    Provides statistical analysis and trend detection.
    """
    
    def __init__(self, attendance_repo: AttendanceRepository):
        self.attendance_repo = attendance_repo
    
    def get_weekly_trends(self, end_date: datetime.date, weeks: int = 4) -> pd.DataFrame:
        """
        Get attendance trends over multiple weeks.
        """
        df = self.attendance_repo.fetch()
        if df is None:
            return pd.DataFrame()
        
        start_date = end_date - timedelta(weeks=weeks)
        mask = (df['Tanggal'] >= start_date) & (df['Tanggal'] <= end_date)
        df_period = df[mask].copy()
        
        # Group by week
        df_period['Week'] = df_period[AppConstants.COL_EVENT_TIME].dt.isocalendar().week
        df_period['Year'] = df_period[AppConstants.COL_EVENT_TIME].dt.year
        
        weekly_stats = df_period.groupby(['Year', 'Week']).agg({
            AppConstants.COL_PERSON_NAME: 'nunique',
            AppConstants.COL_EVENT_TIME: 'count'
        }).reset_index()
        
        weekly_stats.columns = ['Year', 'Week', 'Unique_Employees', 'Total_Events']
        
        return weekly_stats
    
    def get_division_statistics(self, target_date: datetime.date) -> Dict[str, Dict]:
        """
        Calculate statistics per division.
        """
        df = self.attendance_repo.fetch()
        if df is None:
            return {}
        
        df_day = df[df['Tanggal'] == target_date]
        stats = {}
        
        for division_name, division_config in DivisionRegistry.get_all().items():
            members = division_config.members
            df_div = df_day[df_day[AppConstants.COL_PERSON_NAME].isin(members)]
            
            present = df_div[AppConstants.COL_PERSON_NAME].nunique()
            total = len(members)
            
            stats[division_name] = {
                'total': total,
                'present': present,
                'absent': total - present,
                'rate': (present / total * 100) if total > 0 else 0,
                'color': division_config.color,
                'icon': division_config.icon
            }
        
        return stats
    
    def detect_anomalies(self, df: pd.DataFrame, threshold_hours: int = 12) -> List[Dict]:
        """
        Detect anomalous attendance patterns.
        """
        anomalies = []
        
        for name in df[AppConstants.COL_PERSON_NAME].unique():
            person_data = df[df[AppConstants.COL_PERSON_NAME] == name].sort_values(AppConstants.COL_EVENT_TIME)
            
            if len(person_data) > 1:
                times = person_data[AppConstants.COL_EVENT_TIME].tolist()
                
                for i in range(len(times) - 1):
                    gap = (times[i+1] - times[i]).total_seconds() / 3600
                    
                    if gap > threshold_hours:
                        anomalies.append({
                            'employee': name,
                            'type': 'LARGE_GAP',
                            'gap_hours': gap,
                            'time1': times[i],
                            'time2': times[i+1]
                        })
        
        return anomalies


# ================================================================================
# SECTION 4: EXPORT & REPORTING LAYER (MODIFIED FOR IMAGE REPLICATION)
# ================================================================================

class ExcelExporter:
    def __init__(self):
        self.workbook = None
        self.formats = {}

    def _init_formats(self, workbook):
        """Helper to initialize formats only once per workbook"""
        # ... (FORMATS SAMA SEPERTI KODINGANMU SEBELUMNYA) ...
        self.fmt_head = workbook.add_format({'bold': True, 'fg_color': '#4caf50', 'font_color': 'white', 'border': 1, 'align': 'center'})
        self.fmt_norm = workbook.add_format({'border': 1, 'align': 'center'})
        self.fmt_miss = workbook.add_format({'bg_color': '#FF0000', 'border': 1, 'align': 'center'}) 
        self.fmt_full = workbook.add_format({'bg_color': '#FFFF00', 'border': 1, 'align': 'center'}) 
        self.fmt_late = workbook.add_format({'font_color': 'red', 'bold': True, 'border': 1, 'align': 'center'})

    # UPDATE: Tambahkan parameter 'target_date' disini
    def _write_sheet_content(self, ws, df: pd.DataFrame, status_dict: Dict[str, str], target_date: date):
        # Headers
        headers = ['Nama Karyawan', 'Pagi', 'Siang_1', 'Siang_2', 'Sore', 'Keterangan']
        ws.write_row(0, 0, headers, self.fmt_head)
        ws.set_column(0, 0, 30)
        ws.set_column(1, 5, 15)
        
        # LOGIC JUMAT LEBIH AMAN (Cek langsung dari target_date)
        is_friday_global = (target_date.weekday() == 4)

        for idx, row in df.iterrows():
            row_num = idx + 1
            
            nm = row[AppConstants.COL_EMPLOYEE_NAME]
            pagi_str = row.get('Pagi', '')
            siang1_str = row.get('Siang_1', '') 
            siang2_str = row.get('Siang_2', '') 
            sore_str = row.get('Sore', '')
            manual_stat = status_dict.get(nm, "")
            
            ws.write(row_num, 0, nm, self.fmt_norm)
            ws.write(row_num, 5, manual_stat, self.fmt_norm)

            # LOGIC 1: Izin Manual -> Kuning
            if manual_stat:
                for i in range(1, 5): ws.write(row_num, i, "", self.fmt_full)
                continue 

            # LOGIC 2: Alpha (Kosong Semua) -> Kuning Full
            if not pagi_str and not siang1_str and not siang2_str and not sore_str:
                for i in range(1, 5): ws.write(row_num, i, "", self.fmt_full)
                continue 

            # --- DETEKSI SHIFT UNTUK PEWARNAAN ---
            is_shift_2 = False
            t_pagi = None
            
            if pagi_str:
                try:
                    t_pagi = datetime.strptime(pagi_str, "%H:%M").time()
                    # Cutoff 08:15 menentukan Shift
                    if t_pagi > AppConstants.SHIFT_CUTOFF:
                        is_shift_2 = True
                except ValueError:
                    pass
            
            # --- PENENTUAN BATAS MERAH ---
            if is_shift_2:
                batas_datang = AppConstants.S2_LATE_TOLERANCE # 09:05
                
                if is_friday_global:
                    batas_balik = AppConstants.S1_BREAK_IN_END # 14:00 (Jumat Shift 2 ikut S1)
                else:
                    batas_balik = AppConstants.S2_NORM_BREAK_IN_END # 16:00
            else:
                # Shift 1
                batas_datang = AppConstants.S1_LATE_TOLERANCE # 07:05
                batas_balik   = AppConstants.S1_BREAK_IN_END   # 14:00

            # --- WRITE CELLS ---

            # 1. Pagi (Datang)
            if pagi_str and t_pagi:
                fmt = self.fmt_late if t_pagi > batas_datang else self.fmt_norm
                ws.write(row_num, 1, pagi_str, fmt)
            else:
                ws.write(row_num, 1, "", self.fmt_miss) # Merah Kosong

            # 2. Siang 1 (Keluar) - Standar
            ws.write(row_num, 2, siang1_str, self.fmt_norm if siang1_str else self.fmt_miss) 

            # 3. Siang 2 (Balik) - Cek Telat Balik
            if siang2_str:
                try:
                    t_balik = datetime.strptime(siang2_str, "%H:%M").time()
                    fmt = self.fmt_late if t_balik > batas_balik else self.fmt_norm
                    ws.write(row_num, 3, siang2_str, fmt)
                except ValueError:
                    ws.write(row_num, 3, siang2_str, self.fmt_norm)
            else:
                ws.write(row_num, 3, "", self.fmt_miss) 

            # 4. Sore (Pulang)
            ws.write(row_num, 4, sore_str, self.fmt_norm if sore_str else self.fmt_miss)

    def create_attendance_report(self, df: pd.DataFrame, status_dict: Dict[str, str], date_obj: date, metrics: Any = None):
        """Creates a single sheet report"""
        output = io.BytesIO()
        self.workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        self._init_formats(self.workbook)
        
        sheet_name = date_obj.strftime('%d-%b')
        ws = self.workbook.add_worksheet(sheet_name)
        
        # Pass date_obj explicitly
        self._write_sheet_content(ws, df, status_dict, target_date=date_obj)
        
        self.workbook.close()
        output.seek(0)
        return output

    def create_range_report(self, data_map: Dict[datetime.date, Tuple[pd.DataFrame, Dict]]):
        """Creates a multi-sheet Excel report for a date range."""
        output = io.BytesIO()
        self.workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        self._init_formats(self.workbook)

        sorted_dates = sorted(data_map.keys())

        for date_obj in sorted_dates:
            df, status_dict = data_map[date_obj]
            sheet_name = date_obj.strftime('%d-%b') 
            ws = self.workbook.add_worksheet(sheet_name)
            # Pass date_obj explicitly agar logic Jumat per sheet benar
            self._write_sheet_content(ws, df, status_dict, target_date=date_obj)

        self.workbook.close()
        output.seek(0)
        return output

# ================================================================================
# SECTION 5: UI STYLING LAYER
# ================================================================================

class ThemeManager:
    """
    Centralized theme and styling management.
    Implements Strategy pattern for theme customization.
    """
    
    @staticmethod
    def apply_global_styles():
        """Apply comprehensive CSS styling."""
        st.markdown("""
        <style>
            /* ========== FONT IMPORTS ========== */
            @import url('https://fonts.googleapis.com/css2?family=Rajdhani:wght@400;600;700&family=Inter:wght@300;400;600;800&family=JetBrains+Mono:wght@400;700&display=swap');
            
            /* ========== CSS VARIABLES ========== */
            :root {
                --bg-primary: #0b0e14;
                --bg-secondary: #151922;
                --bg-tertiary: #1e2530;
                --accent-blue: #00a8ff;
                --accent-cyan: #00cec9;
                --accent-purple: #8c7ae6;
                --alert-red: #e84118;
                --alert-amber: #fbc531;
                --success-green: #4cd137;
                --text-primary: #f5f6fa;
                --text-secondary: #7f8fa6;
                --text-muted: #546e7a;
                --border-color: #2f3640;
                --border-glow: rgba(0, 168, 255, 0.3);
                --shadow-sm: 0 2px 4px rgba(0,0,0,0.3);
                --shadow-md: 0 4px 12px rgba(0,0,0,0.4);
                --shadow-lg: 0 8px 24px rgba(0,0,0,0.5);
                --shadow-glow: 0 0 20px rgba(0, 168, 255, 0.2);
                --transition-fast: 0.2s ease;
                --transition-normal: 0.3s ease;
                --transition-slow: 0.5s ease;
            }
            
            /* ========== BASE STYLES ========== */
            .stApp {
                background: var(--bg-primary);
                background-image: 
                    radial-gradient(circle at 20% 10%, rgba(0, 168, 255, 0.05) 0%, transparent 50%),
                    radial-gradient(circle at 80% 90%, rgba(140, 122, 230, 0.05) 0%, transparent 50%);
                font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
                color: var(--text-primary);
            }
            
            /* ========== SCROLLBAR STYLING ========== */
            ::-webkit-scrollbar {
                width: 8px;
                height: 8px;
            }
            
            ::-webkit-scrollbar-track {
                background: var(--bg-secondary);
            }
            
            ::-webkit-scrollbar-thumb {
                background: var(--accent-blue);
                border-radius: 4px;
            }
            
            ::-webkit-scrollbar-thumb:hover {
                background: var(--accent-cyan);
            }
            
            /* ========== SIDEBAR STYLING ========== */
            section[data-testid="stSidebar"] {
                background: linear-gradient(180deg, #0f1218 0%, #1a1f2e 100%);
                border-right: 1px solid var(--border-color);
                box-shadow: 4px 0 12px rgba(0,0,0,0.3);
            }
            
            section[data-testid="stSidebar"] .stMarkdown {
                padding: 0.5rem 0;
            }
            
            /* ========== TYPOGRAPHY ========== */
            h1, h2, h3, h4, h5, h6 {
                font-family: 'Rajdhani', sans-serif;
                text-transform: uppercase;
                letter-spacing: 1.5px;
                font-weight: 700;
            }
            
            .brand-title {
                font-family: 'Rajdhani', sans-serif;
                font-size: 3.5rem;
                font-weight: 700;
                background: linear-gradient(135deg, var(--accent-blue), var(--accent-cyan), var(--accent-purple));
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                background-clip: text;
                margin-bottom: 0;
                text-shadow: 0 0 30px var(--border-glow);
                animation: titlePulse 3s ease-in-out infinite;
            }
            
            @keyframes titlePulse {
                0%, 100% { opacity: 1; }
                50% { opacity: 0.8; }
            }
            
            .brand-subtitle {
                font-family: 'JetBrains Mono', monospace;
                color: var(--text-secondary);
                font-size: 0.85rem;
                letter-spacing: 4px;
                text-transform: uppercase;
                margin-bottom: 30px;
                border-bottom: 2px solid var(--border-color);
                padding-bottom: 20px;
                position: relative;
            }
            
            .brand-subtitle::after {
                content: '';
                position: absolute;
                bottom: -2px;
                left: 0;
                width: 100px;
                height: 2px;
                background: var(--accent-blue);
                animation: lineExpand 2s ease-in-out infinite;
            }
            
            @keyframes lineExpand {
                0%, 100% { width: 100px; }
                50% { width: 200px; }
            }
            
            /* ========== METRICS STYLING ========== */
            div[data-testid="stMetric"] {
                background: linear-gradient(135deg, rgba(21, 25, 34, 0.8), rgba(30, 37, 48, 0.6));
                border: 1px solid var(--border-color);
                border-left: 4px solid var(--accent-blue);
                border-radius: 8px;
                padding: 20px;
                transition: all var(--transition-normal);
                position: relative;
                overflow: hidden;
            }
            
            div[data-testid="stMetric"]::before {
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: linear-gradient(45deg, transparent, rgba(0, 168, 255, 0.05), transparent);
                transform: translateX(-100%);
                transition: transform var(--transition-slow);
            }
            
            div[data-testid="stMetric"]:hover {
                border-color: var(--accent-cyan);
                box-shadow: var(--shadow-glow);
                transform: translateY(-4px);
            }
            
            div[data-testid="stMetric"]:hover::before {
                transform: translateX(100%);
            }
            
            div[data-testid="stMetricLabel"] {
                font-family: 'Rajdhani', sans-serif;
                font-weight: 600;
                color: var(--text-secondary);
                font-size: 0.9rem;
                text-transform: uppercase;
                letter-spacing: 1px;
            }
            
            div[data-testid="stMetricValue"] {
                font-family: 'JetBrains Mono', monospace;
                font-size: 2.5rem;
                color: var(--text-primary);
                font-weight: 700;
                text-shadow: 0 2px 4px rgba(0,0,0,0.3);
            }
            
            div[data-testid="stMetricDelta"] {
                font-family: 'Inter', sans-serif;
                font-size: 0.85rem;
            }
            
            /* ========== CARD COMPONENTS ========== */
            .card {
                background: linear-gradient(135deg, var(--bg-secondary) 0%, var(--bg-tertiary) 100%);
                border: 1px solid var(--border-color);
                border-radius: 10px;
                margin-bottom: 20px;
                transition: all var(--transition-normal);
                position: relative;
                overflow: hidden;
                box-shadow: var(--shadow-md);
            }
            
            .card::before {
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                height: 3px;
                background: linear-gradient(90deg, var(--accent-blue), var(--accent-cyan));
                transform: scaleX(0);
                transform-origin: left;
                transition: transform var(--transition-normal);
            }
            
            .card:hover {
                border-color: var(--accent-blue);
                box-shadow: var(--shadow-lg), var(--shadow-glow);
                transform: translateY(-4px);
            }
            
            .card:hover::before {
                transform: scaleX(1);
            }
            
            .card-header {
                padding: 18px;
                display: flex;
                align-items: center;
                background: rgba(255, 255, 255, 0.03);
                border-bottom: 1px dashed var(--border-color);
                position: relative;
            }
            
            .card-header::after {
                content: '';
                position: absolute;
                bottom: 0;
                left: 0;
                width: 60px;
                height: 1px;
                background: var(--accent-cyan);
            }
            
            .card-body {
                padding: 18px;
            }
            
            .card-name {
                font-family: 'Rajdhani', sans-serif;
                font-weight: 700;
                font-size: 1.15rem;
                margin: 0;
                color: var(--text-primary);
                line-height: 1.3;
            }
            
            .card-code {
                font-family: 'JetBrains Mono', monospace;
                font-size: 0.75rem;
                color: var(--text-secondary);
                letter-spacing: 1.5px;
                font-weight: 600;
            }
            
            /* ========== FLIGHT LOG STYLES ========== */
            .flight-row {
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 10px;
                padding: 8px 0;
                border-bottom: 1px solid rgba(255,255,255,0.05);
            }
            
            .flight-row:last-child {
                border-bottom: none;
            }
            
            .flight-label {
                font-size: 0.7rem;
                color: var(--text-muted);
                text-transform: uppercase;
                letter-spacing: 1px;
                font-weight: 600;
            }
            
            .flight-value {
                font-family: 'JetBrains Mono', monospace;
                font-weight: 700;
                font-size: 1rem;
                color: var(--text-primary);
            }
            
            /* ========== STATUS BADGES ========== */
            .status-badge {
                font-family: 'Rajdhani', sans-serif;
                font-weight: 700;
                font-size: 0.8rem;
                padding: 4px 12px;
                border-radius: 4px;
                display: inline-block;
                text-transform: uppercase;
                letter-spacing: 1px;
            }
            
            .status-present {
                color: var(--success-green);
                border: 1px solid var(--success-green);
                background: rgba(76, 209, 55, 0.1);
                box-shadow: 0 0 10px rgba(76, 209, 55, 0.2);
            }
            
            .status-partial {
                color: var(--alert-amber);
                border: 1px solid var(--alert-amber);
                background: rgba(251, 197, 49, 0.1);
                box-shadow: 0 0 10px rgba(251, 197, 49, 0.2);
            }
            
            .status-absent {
                color: var(--alert-red);
                border: 1px solid var(--alert-red);
                background: rgba(232, 65, 24, 0.1);
                box-shadow: 0 0 10px rgba(232, 65, 24, 0.2);
            }
            
            .status-permit {
                color: #9c88ff;
                border: 1px solid #9c88ff;
                background: rgba(156, 136, 255, 0.1);
                box-shadow: 0 0 10px rgba(156, 136, 255, 0.2);
            }
            
            /* ========== ANOMALY BOXES ========== */
            .anomaly-box {
                padding: 12px 15px;
                margin-bottom: 10px;
                border-left: 4px solid;
                background: rgba(255, 255, 255, 0.03);
                font-family: 'JetBrains Mono', monospace;
                font-size: 0.85rem;
                display: flex;
                justify-content: space-between;
                align-items: center;
                border-radius: 4px;
                transition: all var(--transition-fast);
            }
            
            .anomaly-box:hover {
                background: rgba(255, 255, 255, 0.06);
                transform: translateX(4px);
            }
            
            .box-telat {
                border-color: var(--alert-red);
                background: linear-gradient(90deg, rgba(232, 65, 24, 0.15), transparent);
            }
            
            .box-izin {
                border-color: #9c88ff;
                background: linear-gradient(90deg, rgba(156, 136, 255, 0.15), transparent);
            }
            
            .box-alpha {
                border-color: var(--alert-amber);
                background: linear-gradient(90deg, rgba(251, 197, 49, 0.15), transparent);
            }
            
            /* ========== BUTTONS ========== */
            .stDownloadButton button, .stButton button {
                background: linear-gradient(135deg, var(--accent-blue), var(--accent-cyan)) !important;
                color: #000 !important;
                font-weight: 800 !important;
                border: none !important;
                border-radius: 6px !important;
                padding: 12px 24px !important;
                text-transform: uppercase !important;
                letter-spacing: 1px !important;
                transition: all var(--transition-normal) !important;
                box-shadow: var(--shadow-md) !important;
            }
            
            .stDownloadButton button:hover, .stButton button:hover {
                transform: translateY(-2px) !important;
                box-shadow: var(--shadow-lg), var(--shadow-glow) !important;
            }
            
            /* ========== INPUT FIELDS ========== */
            .stTextInput input, .stDateInput input, .stSelectbox select {
                background-color: var(--bg-secondary) !important;
                color: var(--text-primary) !important;
                border: 1px solid var(--border-color) !important;
                border-radius: 6px !important;
                padding: 10px !important;
                transition: all var(--transition-fast) !important;
            }
            
            .stTextInput input:focus, .stDateInput input:focus {
                border-color: var(--accent-blue) !important;
                box-shadow: 0 0 0 2px rgba(0, 168, 255, 0.2) !important;
            }
            
            /* ========== EXPANDER ========== */
            .streamlit-expanderHeader {
                background: rgba(255, 255, 255, 0.05) !important;
                color: var(--accent-blue) !important;
                font-family: 'Rajdhani', sans-serif !important;
                font-weight: 600 !important;
                border-radius: 6px !important;
                padding: 12px 16px !important;
                transition: all var(--transition-fast) !important;
            }
            
            .streamlit-expanderHeader:hover {
                background: rgba(255, 255, 255, 0.08) !important;
                border-color: var(--accent-cyan) !important;
            }
            
            /* ========== DATAFRAME STYLING ========== */
            div[data-testid="stDataFrame"] {
                background-color: var(--bg-secondary);
                border: 1px solid var(--border-color);
                border-radius: 8px;
                overflow: hidden;
            }
            
            /* ========== TABS ========== */
            .stTabs [data-baseweb="tab-list"] {
                gap: 8px;
                background-color: var(--bg-secondary);
                border-radius: 8px;
                padding: 8px;
            }
            
            .stTabs [data-baseweb="tab"] {
                height: 50px;
                background-color: transparent;
                border-radius: 6px;
                color: var(--text-secondary);
                font-family: 'Rajdhani', sans-serif;
                font-weight: 600;
                font-size: 1rem;
                transition: all var(--transition-fast);
            }
            
            .stTabs [data-baseweb="tab"]:hover {
                background-color: rgba(255, 255, 255, 0.05);
                color: var(--text-primary);
            }
            
            .stTabs [aria-selected="true"] {
                background: linear-gradient(135deg, var(--accent-blue), var(--accent-cyan)) !important;
                color: #000 !important;
            }
            
            /* ========== RADIO BUTTONS ========== */
            div[role="radiogroup"] {
                background: rgba(255, 255, 255, 0.03);
                padding: 15px;
                border-radius: 8px;
                border: 1px solid var(--border-color);
            }
            
            div[role="radiogroup"] label {
                font-family: 'Rajdhani', sans-serif;
                font-weight: 600;
                padding: 8px 12px;
                cursor: pointer;
                transition: all var(--transition-fast);
            }
            
            div[role="radiogroup"] label:hover {
                color: var(--accent-cyan);
            }
            
            /* ========== POPOVER STYLING ========== */
            [data-testid="stPopover"] {
                background: var(--bg-secondary);
                border: 1px solid var(--border-color);
                border-radius: 8px;
                box-shadow: var(--shadow-lg);
            }
            
            /* ========== DIVIDER ========== */
            hr {
                border-color: var(--border-color);
                margin: 30px 0;
            }
            
            /* ========== LOADING ANIMATION ========== */
            .stSpinner > div {
                border-top-color: var(--accent-blue) !important;
            }
            
            /* ========== TOAST NOTIFICATIONS ========== */
            .stAlert {
                border-radius: 8px;
                border-left: 4px solid;
            }
            
            /* ========== CUSTOM ANIMATIONS ========== */
            @keyframes fadeInUp {
                from {
                    opacity: 0;
                    transform: translateY(20px);
                }
                to {
                    opacity: 1;
                    transform: translateY(0);
                }
            }
            
            .animate-fade-in {
                animation: fadeInUp 0.5s ease-out;
            }
            
            /* ========== UTILITY CLASSES ========== */
            .text-gradient {
                background: linear-gradient(135deg, var(--accent-blue), var(--accent-cyan));
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
            }
            
            .glow-effect {
                box-shadow: 0 0 20px rgba(0, 168, 255, 0.3);
            }
            
            /* ========== RESPONSIVE DESIGN ========== */
            @media (max-width: 768px) {
                .brand-title {
                    font-size: 2rem;
                }
                
                .brand-subtitle {
                    font-size: 0.7rem;
                }
                
                div[data-testid="stMetric"] {
                    padding: 15px;
                }
            }
        </style>
        """, unsafe_allow_html=True)
    
    @staticmethod
    def get_avatar_url(name: str) -> str:
        """Generate avatar URL for employee."""
        encoded_name = name.replace(' ', '+')
        return f"https://ui-avatars.com/api/?name={encoded_name}&background=random&color=fff&size=128&bold=true&rounded=true"


# ================================================================================
# SECTION 6: UI COMPONENT LAYER
# ================================================================================

class ComponentRenderer:
    """
    Centralized component rendering with consistent styling.
    Implements Facade pattern for complex UI operations.
    """
    
    def __init__(self):
        self.time_service = TimeService()
    
    def render_employee_card(
        self, 
        employee_data: pd.Series, 
        status_dict: Dict[str, str]
    ) -> None:
        """
        Render individual employee attendance card with boarding pass design.
        """
        name = employee_data[AppConstants.COL_EMPLOYEE_NAME]
        morning = employee_data.get('Pagi', '')
        break_out = employee_data.get('Siang_1', '')
        break_in = employee_data.get('Siang_2', '')
        evening = employee_data.get('Sore', '')
        
        # Get division info
        division = DivisionRegistry.find_by_member(name)
        if division:
            div_name = division.name
            div_code = division.code
            div_color = division.color
            div_icon = division.icon
        else:
            div_name = "UNKNOWN"
            div_code = "UNK"
            div_color = "#666"
            div_icon = "❓"
        
        # Determine status
        manual_status = status_dict.get(name, "")
        times = [morning, break_out, break_in, evening]
        empty_count = sum(1 for t in times if t == '')
        
        if manual_status:
            status = AttendanceStatus.PERMIT
            status_text = f"PERMIT: {manual_status}"
        elif empty_count == 4:
            status = AttendanceStatus.ABSENT
            status_text = status.display_text
        elif empty_count > 0:
            status = AttendanceStatus.PARTIAL_DUTY
            status_text = status.display_text
        else:
            status = AttendanceStatus.FULL_DUTY
            status_text = status.display_text
        
        # Check if late (UPDATED DUAL SHIFT)
        is_late = False
        shift_label = ""
        
        if morning:
            # Panggil fungsi logic baru yang mengembalikan 2 data
            is_late_status, shift_name = self.time_service.is_late(morning)
            is_late = is_late_status
            shift_label = shift_name

        late_indicator = ""
        if is_late:
            late_indicator = f"<span style='color:#e84118; font-weight:bold; margin-left:8px;'>⚠ TELAT ({shift_label})</span>"
        elif morning:
            late_indicator = f"<span style='color:#4cd137; font-weight:bold; font-size:0.8rem; margin-left:5px;'>✓ {shift_label}</span>"
        
        # Get avatar
        avatar_url = ThemeManager.get_avatar_url(name)
        
        # Render card
        st.markdown(f"""
        <div class="card" style="border-left: 4px solid {status.color};">
            <div class="card-header">
                <img src="{avatar_url}" 
                    style="width:50px; height:50px; border-radius:50%; margin-right:15px; 
                             border:3px solid {div_color}; box-shadow: 0 4px 8px rgba(0,0,0,0.3);">
                <div style="flex:1; min-width:0;">
                    <p class="card-name" style="white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                        {div_icon} {name}
                    </p>
                    <span class="card-code" style="color:{div_color}; font-weight:bold;">
                        ✈ {div_code}
                    </span>
                    <span class="card-code" style="color:#888; margin-left:8px;">
                        // {div_name}
                    </span>
                </div>
            </div>
            <div class="card-body">
                <div class="flight-row">
                    <span class="flight-label">🛫 Jam Datang</span>
                    <span class="flight-value" style="font-size:1.2rem;">
                        {morning if morning else '━━:━━'} {late_indicator}
                    </span>
                </div>
                <div class="flight-row">
                    <span class="flight-label">🛬 Jam Pulang</span>
                    <span class="flight-value" style="font-size:1.2rem;">
                        {evening if evening else '━━:━━'}
                    </span>
                </div>
                <div style="margin-top:12px; text-align:right;">
                    <span class="status-badge {status.css_class}">{status_text}</span>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # Detail popover
        with st.popover("📋 DETAILED FLIGHT LOG", use_container_width=True):
            self._render_detail_popover(name, morning, break_out, break_in, evening, 
                                        status_text, status.color, div_name, is_late)
    
    def _render_detail_popover(
        self, 
        name: str, 
        morning: str, 
        break_out: str, 
        break_in: str, 
        evening: str,
        status_text: str,
        status_color: str,
        division: str,
        is_late: bool
    ) -> None:
        """Render detailed attendance information in popover."""
        st.markdown(f"### ✈️ FLIGHT RECORD: {name}")
        st.markdown(f"**DIVISION:** {division}")
        st.markdown(f"**STATUS:** <span style='color:{status_color}; font-weight:bold'>{status_text}</span>", 
                    unsafe_allow_html=True)
        st.divider()
        
        # Time grid
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### 🛫 JAM DATANG")
            
            check_in = morning if morning else "❌ NOT RECORDED"
            if is_late:
                check_in += " ⚠️ **LATE**"
            st.info(f"**Jam Datang:** {check_in}")
            
            break_out_display = break_out if break_out else "❌ NOT RECORDED"
            st.write(f"**Siang 1:** {break_out_display}")
        
        with col2:
            st.markdown("#### 🛬 JAM PULANG")
            
            break_in_display = break_in if break_in else "❌ NOT RECORDED"
            st.write(f"**Siang 2:** {break_in_display}")
            
            check_out = evening if evening else "❌ NOT RECORDED"
            st.success(f"**Jam Pulang:** {check_out}")
        
        # Calculate work duration
        if morning and evening:
            duration = self.time_service.calculate_duration(morning, evening)
            if duration:
                formatted_duration = self.time_service.format_duration(duration)
                st.divider()
                st.markdown(f"### ⏱️ TOTAL DUTY TIME")
                st.metric("Duration", formatted_duration)
                
                # Add overtime indicator
                if duration.seconds / 3600 > 9:
                    st.warning("⚠️ Extended duty hours detected")
    
    def render_metric_cards(self, metrics: Dict[str, Any]) -> None:
        """Render key metrics in card format."""
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                "👥 TOTAL PERSONNEL",
                metrics['total'],
                help="Total registered employees"
            )
        
        with col2:
            attendance_pct = f"{metrics['attendance_rate']:.1f}%"
            st.metric(
                "✅ ON DUTY",
                metrics['present'],
                delta=attendance_pct,
                help="Employees present today"
            )
        
        with col3:
            st.metric(
                "📝 PERMITS",
                metrics['permit'],
                help="Approved leaves and permits"
            )
        
        with col4:
            absent_delta = f"-{metrics['absent']}" if metrics['absent'] > 0 else "0"
            st.metric(
                "❌ ABSENT",
                metrics['absent'],
                delta=absent_delta,
                delta_color="inverse",
                help="Unexplained absences"
            )
    
    def render_anomaly_section(self, metrics: Dict[str, Any]) -> None:
        """Render anomaly detection section."""
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("#### ⚠️ TERLAMBAT")
            late_list = metrics['late_list']
            
            if late_list:
                with st.container(height=180):
                    for name, time_val in late_list:
                        st.markdown(
                            f"<div class='anomaly-box box-telat'>"
                            f"<span>{name}</span>"
                            f"<span style='font-weight:bold'>{time_val}</span>"
                            f"</div>",
                            unsafe_allow_html=True
                        )
            else:
                st.success("✓ ALL ON TIME")
        
        with col2:
            st.markdown("#### ℹ️ SAKIT/IZIN")
            permit_list = metrics['permit_list']
            
            if permit_list:
                with st.container(height=180):
                    for name, reason in permit_list:
                        st.markdown(
                            f"<div class='anomaly-box box-izin'>"
                            f"<span>{name}</span>"
                            f"<span style='font-weight:bold'>{reason}</span>"
                            f"</div>",
                            unsafe_allow_html=True
                        )
            else:
                st.info("○ NO PERMITS")
        
        with col3:
            st.markdown("#### ❌ YANG TIDAK HADIR")
            absent_list = metrics['absent_list']
            
            if absent_list:
                with st.container(height=180):
                    for name in absent_list:
                        st.markdown(
                            f"<div class='anomaly-box box-alpha'>"
                            f"<span>{name}</span>"
                            f"<span style='color:#e84118'>ABSENT</span>"
                            f"</div>",
                            unsafe_allow_html=True
                        )
            else:
                st.success("✓ FULL ATTENDANCE")
    
    def render_division_tabs(
        self, 
        df: pd.DataFrame, 
        status_dict: Dict[str, str],
        search_query: str = ""
    ) -> None:
        """Render division-based tabs with employee cards."""
        divisions = sorted(
            DivisionRegistry.get_all().items(),
            key=lambda x: x[1].priority
        )
        
        tab_names = [f"{div[1].icon} {div[0]}" for div in divisions]
        tabs = st.tabs(tab_names)
        
        for tab, (div_name, div_config) in zip(tabs, divisions):
            with tab:
                # Division stats
                st.markdown(f"**{div_config.description}**")
                
                # Get members strictly from config order (FORCED SORT ORDER)
                ordered_members = div_config.members
                
                # Filter DF for this division
                df_division = df[df[AppConstants.COL_EMPLOYEE_NAME].isin(ordered_members)].copy()
                
                # Calculate stats
                present = len(df_division[df_division['Pagi'] != ''])
                total = len(df_division)
                rate = (present / total * 100) if total > 0 else 0
                
                st.progress(rate / 100, text=f"Attendance: {present}/{total} ({rate:.1f}%)")
                st.markdown("<br>", unsafe_allow_html=True)
                
                # Apply search filter if exists
                if search_query:
                    ordered_members = [
                        m for m in ordered_members 
                        if search_query.lower() in m.lower()
                    ]
                
                # Create grid
                cols = st.columns(AppConstants.CARDS_PER_ROW)
                card_count = 0
                
                # Loop through ORDERED list, not the DataFrame index
                for member_name in ordered_members:
                    # Find the data row for this member
                    member_row = df_division[df_division[AppConstants.COL_EMPLOYEE_NAME] == member_name]
                    
                    if not member_row.empty:
                        row_data = member_row.iloc[0]
                        with cols[card_count % AppConstants.CARDS_PER_ROW]:
                            self.render_employee_card(row_data, status_dict)
                        card_count += 1
                
                if card_count == 0:
                    st.info(f"No personnel found in {div_name}")

# ================================================================================
# SECTION 7: VISUALIZATION LAYER
# ================================================================================

class ChartBuilder:
    """
    Builder class for creating interactive charts and visualizations.
    Uses Plotly for rich, interactive data visualization.
    """
    
    @staticmethod
    def create_attendance_pie_chart(metrics: Dict[str, Any]) -> go.Figure:
        """Create pie chart for attendance distribution."""
        labels = ['Present', 'Permit', 'Absent']
        values = [metrics['present'], metrics['permit'], metrics['absent']]
        colors = ['#4cd137', '#9c88ff', '#e84118']
        
        fig = go.Figure(data=[go.Pie(
            labels=labels,
            values=values,
            hole=0.4,
            marker=dict(colors=colors, line=dict(color='#0b0e14', width=2)),
            textfont=dict(size=14, color='white', family='Rajdhani'),
            hovertemplate='<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>'
        )])
        
        fig.update_layout(
            title=dict(
                text='Attendance Distribution',
                font=dict(size=20, color='white', family='Rajdhani')
            ),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            showlegend=True,
            legend=dict(
                font=dict(color='white', family='Inter'),
                bgcolor='rgba(21, 25, 34, 0.8)'
            ),
            height=350
        )
        
        return fig
    
    @staticmethod
    def create_division_bar_chart(division_stats: Dict[str, Dict]) -> go.Figure:
        """Create bar chart for division-wise attendance."""
        divisions = []
        present = []
        absent = []
        colors = []
        
        for div_name, stats in sorted(division_stats.items(), key=lambda x: x[1]['rate'], reverse=True):
            divisions.append(div_name)
            present.append(stats['present'])
            absent.append(stats['absent'])
            colors.append(stats['color'])
        
        fig = go.Figure()
        
        fig.add_trace(go.Bar(
            name='Present',
            x=divisions,
            y=present,
            marker_color='#4cd137',
            text=present,
            textposition='auto',
        ))
        
        fig.add_trace(go.Bar(
            name='Absent',
            x=divisions,
            y=absent,
            marker_color='#e84118',
            text=absent,
            textposition='auto',
        ))
        
        fig.update_layout(
            title=dict(
                text='Division-wise Attendance',
                font=dict(size=20, color='white', family='Rajdhani')
            ),
            barmode='stack',
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(21, 25, 34, 0.6)',
            font=dict(color='white', family='Inter'),
            xaxis=dict(
                title='Division',
                gridcolor='rgba(255,255,255,0.1)'
            ),
            yaxis=dict(
                title='Count',
                gridcolor='rgba(255,255,255,0.1)'
            ),
            legend=dict(
                bgcolor='rgba(21, 25, 34, 0.8)'
            ),
            height=400
        )
        
        return fig
    
    @staticmethod
    def create_time_distribution_chart(df: pd.DataFrame) -> go.Figure:
        """Create histogram of arrival times."""
        if df.empty or 'Jam' not in df.columns:
            return go.Figure()
        
        fig = go.Figure()
        
        fig.add_trace(go.Histogram(
            x=df['Jam'],
            nbinsx=24,
            marker=dict(
                color='#00a8ff',
                line=dict(color='#0b0e14', width=1)
            ),
            name='Arrivals'
        ))
        
        # Add late threshold line
        fig.add_vline(
            x=7.083,  # 7:05 AM
            line_dash="dash",
            line_color="#e84118",
            annotation_text="Late Threshold",
            annotation_position="top right"
        )
        
        fig.update_layout(
            title=dict(
                text='Arrival Time Distribution',
                font=dict(size=20, color='white', family='Rajdhani')
            ),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(21, 25, 34, 0.6)',
            font=dict(color='white', family='Inter'),
            xaxis=dict(
                title='Hour of Day',
                gridcolor='rgba(255,255,255,0.1)',
                range=[0, 24]
            ),
            yaxis=dict(
                title='Number of Arrivals',
                gridcolor='rgba(255,255,255,0.1)'
            ),
            height=350
        )
        
        return fig


# ================================================================================
# SECTION 8: APPLICATION CONTROLLER LAYER
# ================================================================================
class AttendanceController:
    """
    Main application controller implementing MVC pattern.
    Orchestrates all business logic and UI rendering.
    """
    
    def __init__(self):
        # Initialize repositories
        self.attendance_repo = AttendanceRepository(DataSourceConfig.ATTENDANCE_SHEET_URL)
        self.status_repo = StatusRepository(DataSourceConfig.STATUS_SHEET_URL)
        
        # Initialize services
        self.attendance_service = AttendanceService(self.attendance_repo, self.status_repo)
        self.analytics_service = AnalyticsService(self.attendance_repo)
        
        # Initialize UI components
        self.component_renderer = ComponentRenderer()
        self.chart_builder = ChartBuilder()
        self.excel_exporter = ExcelExporter()

    def run_dashboard(self) -> None:
        """Main dashboard view."""
        # 1. Header
        st.markdown('<div class="brand-title">WedaBayAirport</div>', unsafe_allow_html=True)
        st.markdown('<div class="brand-subtitle">REAL-TIME PERSONNEL MONITORING SYSTEM</div>', 
                    unsafe_allow_html=True)
        
        # 2. Check data availability
        df_attendance = self.attendance_repo.fetch()
        if df_attendance is None:
            st.error("⚠️ SYSTEM OFFLINE - Unable to connect to attendance database")
            st.stop()
        
        # 3. Filters & Date Selection
        col1, col2, col3 = st.columns([2, 3, 2])
        
        with col1:
            # Ambil data unik dulu
            unfiltered_dates = df_attendance['Tanggal'].unique()
            
            # Filter manual: hanya ambil data yang BUKAN NaT
            clean_dates = [d for d in unfiltered_dates if pd.notna(d)]
            
            # Baru di-sort
            available_dates = sorted(clean_dates, reverse=True)

            # PERBAIKAN: Geser 'if' ke KIRI agar lurus dengan 'available_dates' di atasnya
            if not available_dates:
                st.error("No attendance data available")
                st.stop()
            
            # PERBAIKAN: Geser 'selected_date' ke KIRI agar lurus dengan 'available_dates'
            selected_date = st.date_input(
                "📅 OPERATION DATE",
                value=available_dates[0],
                help="Select date to view attendance"
            )
        
        with col2:
            search_query = st.text_input(
                "🔍 PERSONNEL SEARCH",
                placeholder="Search by name...",
                help="Filter employees by name"
            )
        
        with col3:
            view_mode = st.selectbox(
                "👁️ VIEW MODE",
                ["Cards", "Table", "Analytics"],
                help="Choose display mode"
            )
        
        st.markdown("---")
        
        # 4. Build report variables (df_final & metrics)
        with st.spinner("🔄 Loading flight data..."):
            df_final, status_dict = self.attendance_service.build_complete_report(selected_date)
            metrics = self.attendance_service.calculate_metrics(df_final, status_dict)
        
        # 5. Metrics section UI
        self.component_renderer.render_metric_cards(metrics)
        st.markdown("<br>", unsafe_allow_html=True)
        
        # 6. Anomalies section UI
        with st.container():
            self.component_renderer.render_anomaly_section(metrics)
        
        st.markdown("---")
        
        # 7. EXPORT & REPORTS SECTION
        st.markdown("### 📤 EXPORT REPORTS")
        
        export_tab1, export_tab2 = st.tabs(["📄 Daily Report", "📅 Range Report"])

        # --- TAB 1: DOWNLOAD PER HARI ---
        with export_tab1:
            col_ex1, col_ex2 = st.columns([1, 1])
            with col_ex1:
                st.info(f"Download report for selected date: **{selected_date.strftime('%d %B %Y')}**")
                
                excel_file = self.excel_exporter.create_attendance_report(
                    df_final, status_dict, selected_date, metrics
                )
                
                st.download_button(
                    "📥 DOWNLOAD DAILY EXCEL",
                    data=excel_file,
                    file_name=f"Attendance_{selected_date.strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with col_ex2:
                 if st.button("📊 VIEW ANALYTICS", use_container_width=True):
                    st.session_state['show_analytics'] = True

        # --- TAB 2: DOWNLOAD RANGE TANGGAL ---
        with export_tab2:
            st.write("Select a date range to generate a multi-sheet Excel file.")
            
            # Input Range Tanggal
            col_rng1, col_rng2 = st.columns(2)
            with col_rng1:
                start_date_input = st.date_input("Start Date", value=selected_date - timedelta(days=7))
            with col_rng2:
                end_date_input = st.date_input("End Date", value=selected_date)

            # Tombol Generate
            if st.button("📦 GENERATE RANGE REPORT", use_container_width=True):
                if start_date_input > end_date_input:
                    st.error("Error: Start Date must be before End Date")
                else:
                    with st.spinner(f"Generating report from {start_date_input} to {end_date_input}..."):
                        range_data_map = {}
                        current_loop_date = start_date_input
                        
                        # Loop untuk mengambil data setiap hari dalam range
                        while current_loop_date <= end_date_input:
                            try:
                                # Kita panggil service ulang untuk setiap tanggal
                                day_df, day_status = self.attendance_service.build_complete_report(current_loop_date)
                                range_data_map[current_loop_date] = (day_df, day_status)
                            except Exception:
                                pass # Skip error dates
                            current_loop_date += timedelta(days=1)
                        
                        if range_data_map:
                            range_excel = self.excel_exporter.create_range_report(range_data_map)
                            st.success("✅ Report Generated Successfully!")
                            st.download_button(
                                label=f"📥 DOWNLOAD RANGE REPORT ({start_date_input} - {end_date_input})",
                                data=range_excel,
                                file_name=f"Attendance_Range_{start_date_input}_{end_date_input}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        else:
                            st.warning("No data found for the selected range.")

        st.markdown("---")
        
        # 8. View modes Logic
        if view_mode == "Cards":
            st.markdown("### 📋 PERSONNEL ROSTER")
            self.component_renderer.render_division_tabs(df_final, status_dict, search_query)
        
        elif view_mode == "Table":
            self._render_table_view(df_final, status_dict)
        
        elif view_mode == "Analytics":
            self._render_analytics_view(df_final, status_dict, metrics, selected_date)
        
        # Additional analytics modal
        if st.session_state.get('show_analytics', False):
            with st.expander("📈 ADVANCED ANALYTICS", expanded=True):
                self._render_analytics_view(df_final, status_dict, metrics, selected_date)

    def _render_table_view(self, df: pd.DataFrame, status_dict: Dict[str, str]) -> None:
        """Render table view of attendance."""
        st.markdown("### 📊 DETAILED ATTENDANCE TABLE")
        
        # Prepare display dataframe
        df_display = df.copy()
        
        df_display['Division'] = df_display[AppConstants.COL_EMPLOYEE_NAME].apply(
            lambda x: DivisionRegistry.find_by_member(x).code 
            if DivisionRegistry.find_by_member(x) else "N/A"
        )
        
        df_display['Status'] = df_display[AppConstants.COL_EMPLOYEE_NAME].apply(
            lambda x: status_dict.get(x, "")
        )
        
        # Rename columns for display
        df_display = df_display.rename(columns={
            'Pagi': 'Jam Datang',
            'Siang_1': 'Siang 1',
            'Siang_2': 'Siang 2',
            'Sore': 'Jam Pulang'
        })
        
        display_columns = [
            AppConstants.COL_EMPLOYEE_NAME, 
            'Division', 
            'Jam Datang', 
            'Siang 1', 
            'Siang 2', 
            'Jam Pulang', 
            'Status'
        ]
        
        # Filter existing columns
        final_cols = [c for c in display_columns if c in df_display.columns]
        df_display = df_display[final_cols]
        
        # Styling Function (FIXED)
        def highlight_late(val):
            if isinstance(val, str) and ':' in val:
                # Bongkar paket tuple, ambil status (index 0)
                is_late_status, _ = TimeService.is_late(val)
                if is_late_status:
                    return 'color: #e84118; font-weight: bold'
            return ''
        
        try:
            st.dataframe(
                df_display.style.applymap(highlight_late, subset=['Jam Datang']),
                use_container_width=True,
                height=600,
                hide_index=True
            )
        except Exception:
            st.dataframe(df_display, use_container_width=True, height=600)
        
        # Download Button
        csv = df_display.to_csv(index=False)
        st.download_button(
            "💾 DOWNLOAD CSV",
            data=csv,
            file_name=f"attendance_table_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True
        )

    def _render_analytics_view(
        self, 
        df: pd.DataFrame, 
        status_dict: Dict[str, str],
        metrics: Dict[str, Any],
        selected_date: datetime.date
    ) -> None:
        """Render advanced analytics view."""
        st.markdown("### 📈 ANALYTICS DASHBOARD")
        
        # Charts row 1
        chart_col1, chart_col2 = st.columns(2)
        
        with chart_col1:
            pie_chart = self.chart_builder.create_attendance_pie_chart(metrics)
            st.plotly_chart(pie_chart, use_container_width=True)
        
        with chart_col2:
            division_stats = self.analytics_service.get_division_statistics(selected_date)
            bar_chart = self.chart_builder.create_division_bar_chart(division_stats)
            st.plotly_chart(bar_chart, use_container_width=True)
        
        # Charts row 2
        df_attendance_day = self.attendance_service.get_attendance_for_date(selected_date)
        if df_attendance_day is not None and not df_attendance_day.empty:
            time_dist_chart = self.chart_builder.create_time_distribution_chart(df_attendance_day)
            st.plotly_chart(time_dist_chart, use_container_width=True)
        
        # Key insights
        st.markdown("### 💡 KEY INSIGHTS")
        
        insight_col1, insight_col2, insight_col3 = st.columns(3)
        
        with insight_col1:
            st.metric(
                "Punctuality Rate",
                f"{metrics['punctuality_rate']:.1f}%",
                help="Percentage of on-time arrivals"
            )
        
        with insight_col2:
            best_division = max(
                division_stats.items(),
                key=lambda x: x[1]['rate']
            )
            st.metric(
                "Top Division",
                best_division[0],
                f"{best_division[1]['rate']:.1f}%"
            )
        
        with insight_col3:
            avg_late_time = "N/A"
            if metrics['late_list']:
                late_times = [
                    datetime.strptime(t, '%H:%M').time() 
                    for _, t in metrics['late_list']
                ]
                avg_minutes = sum(
                    t.hour * 60 + t.minute for t in late_times
                ) / len(late_times)
                avg_late_time = f"{int(avg_minutes // 60):02d}:{int(avg_minutes % 60):02d}"
            
            st.metric(
                "Avg Late Arrival",
                avg_late_time,
                help="Average time of late arrivals"
            )

    def run_report_form(self) -> None:
        """Report submission view."""
        st.markdown('<div class="brand-title">MANUAL REPORTING</div>', unsafe_allow_html=True)
        st.markdown('<div class="brand-subtitle">SUBMIT PERMITS & LEAVE REQUESTS</div>', 
                    unsafe_allow_html=True)
        
        st.info("📝 Use the form below to submit permit requests, sick leaves, or other attendance modifications.")
        
        if "PASTE_LINK" in DataSourceConfig.REPORT_FORM_URL:
            st.warning("⚠️ Google Form URL not configured. Please contact system administrator.")
        else:
            components.iframe(
                DataSourceConfig.REPORT_FORM_URL,
                height=1200,
                scrolling=True
            )

# ================================================================================
# SECTION 9: ADDITIONAL FEATURES & UTILITIES
# ================================================================================

class NotificationSystem:
    """
    Notification system for alerts and updates.
    Can be extended to send email/SMS notifications.
    """
    
    @staticmethod
    def show_late_arrivals_alert(late_count: int) -> None:
        """Display alert for late arrivals."""
        if late_count > 0:
            st.warning(f"⚠️ {late_count} personnel arrived late today")
    
    @staticmethod
    def show_absent_alert(absent_count: int, threshold: int = 5) -> None:
        """Display alert for high absence rate."""
        if absent_count >= threshold:
            st.error(f"🚨 ALERT: {absent_count} unexplained absences detected!")
    
    @staticmethod
    def show_success_message(message: str) -> None:
        """Display success notification."""
        st.success(f"✅ {message}")
    
    @staticmethod
    def show_info_message(message: str) -> None:
        """Display info notification."""
        st.info(f"ℹ️ {message}")


class DataValidator:
    """
    Comprehensive data validation utilities.
    Ensures data integrity across the application.
    """
    
    @staticmethod
    def validate_employee_name(name: str) -> bool:
        """Validate employee name format."""
        if not name or not isinstance(name, str):
            return False
        
        # Check minimum length
        if len(name.strip()) < 3:
            return False
        
        # Check for valid characters
        if not all(c.isalpha() or c.isspace() or c in ".,'-" for c in name):
            return False
        
        return True
    
    @staticmethod
    def validate_time_format(time_str: str) -> bool:
        """Validate time string format."""
        try:
            datetime.strptime(time_str, AppConstants.TIME_FORMAT)
            return True
        except (ValueError, TypeError):
            return False
    
    @staticmethod
    def validate_date_range(start_date: datetime.date, end_date: datetime.date) -> bool:
        """Validate date range."""
        if start_date > end_date:
            return False
        
        # Check if range is too large (more than 1 year)
        if (end_date - start_date).days > 365:
            return False
        
        return True
    
    @staticmethod
    def sanitize_input(input_str: str) -> str:
        """Sanitize user input to prevent injection attacks."""
        if not isinstance(input_str, str):
            return ""
        
        # Remove potentially dangerous characters
        sanitized = input_str.strip()
        dangerous_chars = ['<', '>', '{', '}', '|', '\\', '^', '~', '[', ']', '`']
        
        for char in dangerous_chars:
            sanitized = sanitized.replace(char, '')
        
        return sanitized


class PerformanceMonitor:
    """
    Performance monitoring and optimization utilities.
    Tracks application performance metrics.
    """
    
    def __init__(self):
        self.start_time = None
        self.metrics = {}
    
    def start_timer(self, operation: str) -> None:
        """Start timing an operation."""
        self.start_time = datetime.now()
        self.metrics[operation] = {'start': self.start_time}
    
    def end_timer(self, operation: str) -> float:
        """End timing and return duration in seconds."""
        if operation not in self.metrics:
            return 0.0
        
        end_time = datetime.now()
        duration = (end_time - self.metrics[operation]['start']).total_seconds()
        self.metrics[operation]['duration'] = duration
        
        return duration
    
    def get_summary(self) -> Dict[str, float]:
        """Get performance summary."""
        return {
            op: data.get('duration', 0.0) 
            for op, data in self.metrics.items()
        }
    
    @staticmethod
    def optimize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
        """Optimize DataFrame memory usage."""
        for col in df.columns:
            col_type = df[col].dtype
            
            if col_type == 'object':
                df[col] = df[col].astype('category')
            elif col_type == 'float64':
                df[col] = df[col].astype('float32')
            elif col_type == 'int64':
                df[col] = df[col].astype('int32')
        
        return df


class BackupManager:
    """
    Backup and data recovery management.
    Handles data export and archival.
    """
    
    @staticmethod
    def create_backup(df: pd.DataFrame, backup_name: str) -> io.BytesIO:
        """Create backup of attendance data."""
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Backup', index=False)
            
            # Add metadata sheet
            metadata = pd.DataFrame({
                'Key': ['Backup Date', 'Record Count', 'Version'],
                'Value': [
                    datetime.now().strftime(AppConstants.DATETIME_FORMAT),
                    len(df),
                    AppConstants.APP_VERSION
                ]
            })
            metadata.to_excel(writer, sheet_name='Metadata', index=False)
        
        output.seek(0)
        return output
    
    @staticmethod
    def generate_archive_filename(date: datetime.date) -> str:
        """Generate standardized backup filename."""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        return f"WedaBay_Backup_{date}_{timestamp}.xlsx"


class SearchEngine:
    """
    Advanced search functionality for employee records.
    Implements fuzzy matching and filtering.
    """
    
    @staticmethod
    def search_employees(
        df: pd.DataFrame, 
        query: str, 
        search_columns: List[str] = None
    ) -> pd.DataFrame:
        """
        Search employees with fuzzy matching.
        """
        if not query or df.empty:
            return df
        
        if search_columns is None:
            search_columns = [AppConstants.COL_EMPLOYEE_NAME]
        
        # Sanitize query
        query = DataValidator.sanitize_input(query).lower()
        
        # Create filter mask
        mask = pd.Series([False] * len(df))
        
        for col in search_columns:
            if col in df.columns:
                mask |= df[col].astype(str).str.lower().str.contains(query, na=False)
        
        return df
    
    @staticmethod
    def filter_by_division(df: pd.DataFrame, division_name: str) -> pd.DataFrame:
        """Filter employees by division."""
        division = DivisionRegistry.get(division_name)
        if not division:
            return df
        
        return df[df[AppConstants.COL_EMPLOYEE_NAME].isin(division.members)]
    
    @staticmethod
    def filter_by_status(
        df: pd.DataFrame, 
        status_dict: Dict[str, str],
        status_filter: str
    ) -> pd.DataFrame:
        """Filter employees by attendance status."""
        if status_filter == "ALL":
            return df
        
        filtered_names = []
        
        for _, row in df.iterrows():
            name = row[AppConstants.COL_EMPLOYEE_NAME]
            times = [row.get('Pagi', ''), row.get('Siang_1', ''), 
                     row.get('Siang_2', ''), row.get('Sore', '')]
            empty_count = sum(1 for t in times if t == '')
            manual_status = status_dict.get(name, "")
            
            if status_filter == "PRESENT" and not manual_status and empty_count < 4:
                filtered_names.append(name)
            elif status_filter == "ABSENT" and not manual_status and empty_count == 4:
                filtered_names.append(name)
            elif status_filter == "PERMIT" and manual_status:
                filtered_names.append(name)
            elif status_filter == "LATE":
                morning = row.get('Pagi', '')
                if morning and TimeService.is_late(morning):
                    filtered_names.append(name)
        
        return df[df[AppConstants.COL_EMPLOYEE_NAME].isin(filtered_names)]


class ReportGenerator:
    """
    Advanced report generation with multiple formats.
    Supports PDF, Excel, and JSON exports.
    """
    
    @staticmethod
    def generate_summary_report(
        metrics: Dict[str, Any],
        date: datetime.date
    ) -> str:
        """Generate text summary report."""
        report = f"""
╔══════════════════════════════════════════════════════════════╗
║      WEDABAY AIRPORT ATTENDANCE SUMMARY REPORT               ║
╚══════════════════════════════════════════════════════════════╝

Date: {date.strftime('%B %d, %Y')}
Generated: {datetime.now().strftime(AppConstants.DATETIME_FORMAT)}

──────────────────────────────────────────────────────────────
OVERVIEW STATISTICS
──────────────────────────────────────────────────────────────
Total Personnel:            {metrics['total']}
Present:                    {metrics['present']} ({metrics['attendance_rate']:.1f}%)
On Permit/Leave:            {metrics['permit']}
Absent (No Show):           {metrics['absent']}
Late Arrivals:              {metrics['late']}
Punctuality Rate:           {metrics['punctuality_rate']:.1f}%

──────────────────────────────────────────────────────────────
ANOMALIES DETECTED
──────────────────────────────────────────────────────────────
"""
        
        if metrics['late_list']:
            report += "\nLATE ARRIVALS:\n"
            for name, time_val in metrics['late_list']:
                report += f"  • {name}: {time_val}\n"
        
        if metrics['permit_list']:
            report += "\nPERMITS & LEAVES:\n"
            for name, reason in metrics['permit_list']:
                report += f"  • {name}: {reason}\n"
        
        if metrics['absent_list']:
            report += "\nABSENT (NO SHOW):\n"
            for name in metrics['absent_list']:
                report += f"  • {name}\n"
        
        report += f"""
──────────────────────────────────────────────────────────────
System: {AppConstants.APP_TITLE} v{AppConstants.APP_VERSION}
Operator: WedaBay Aviation Services
──────────────────────────────────────────────────────────────
        """
        
        return report
    
    @staticmethod
    def export_to_json(
        df: pd.DataFrame,
        metrics: Dict[str, Any],
        date: datetime.date
    ) -> str:
        """Export data to JSON format."""
        data = {
            'metadata': {
                'date': date.strftime(AppConstants.DATE_FORMAT),
                'generated_at': datetime.now().strftime(AppConstants.DATETIME_FORMAT),
                'version': AppConstants.APP_VERSION
            },
            'metrics': {
                'total': metrics['total'],
                'present': metrics['present'],
                'permit': metrics['permit'],
                'absent': metrics['absent'],
                'late': metrics['late'],
                'attendance_rate': round(metrics['attendance_rate'], 2),
                'punctuality_rate': round(metrics['punctuality_rate'], 2)
            },
            'attendance_records': df.to_dict(orient='records'),
            'anomalies': {
                'late_arrivals': [
                    {'name': name, 'time': time_val} 
                    for name, time_val in metrics['late_list']
                ],
                'permits': [
                    {'name': name, 'reason': reason}
                    for name, reason in metrics['permit_list']
                ],
                'absent': metrics['absent_list']
            }
        }
        
        return json.dumps(data, indent=2, ensure_ascii=False)


# ================================================================================
# SECTION 10: MAIN APPLICATION ENTRY POINT
# ================================================================================

class ConfigurationManager:
    """
    Application configuration management.
    Handles user preferences and settings.
    """
    
    @staticmethod
    def initialize_session_state() -> None:
        """Initialize Streamlit session state variables."""
        if 'show_analytics' not in st.session_state:
            st.session_state['show_analytics'] = False
        
        if 'selected_date' not in st.session_state:
            st.session_state['selected_date'] = datetime.now().date()
        
        if 'view_mode' not in st.session_state:
            st.session_state['view_mode'] = 'Cards'
        
        if 'theme' not in st.session_state:
            st.session_state['theme'] = 'dark'
        
        if 'notifications_enabled' not in st.session_state:
            st.session_state['notifications_enabled'] = True
    
    @staticmethod
    def get_user_preferences() -> Dict[str, Any]:
        """Get user preferences from session state."""
        return {
            'view_mode': st.session_state.get('view_mode', 'Cards'),
            'theme': st.session_state.get('theme', 'dark'),
            'notifications': st.session_state.get('notifications_enabled', True),
            'cards_per_row': AppConstants.CARDS_PER_ROW
        }
    
    @staticmethod
    def save_user_preference(key: str, value: Any) -> None:
        """Save user preference to session state."""
        st.session_state[key] = value


def configure_page() -> None:
    """Configure Streamlit page settings."""
    st.set_page_config(
        page_title=AppConstants.APP_TITLE,
        page_icon=AppConstants.APP_ICON,
        layout=AppConstants.LAYOUT_MODE,
        initial_sidebar_state=AppConstants.SIDEBAR_STATE,
        menu_items={
            'Get Help': 'https://wedabay.airport/support',
            'Report a bug': 'https://wedabay.airport/bug-report',
            'About': f"""
            # {AppConstants.APP_TITLE}
            Version {AppConstants.APP_VERSION}
            
            Enterprise Personnel Monitoring System
            
            © 2025 {AppConstants.COMPANY_NAME}
            """
        }
    )


def render_sidebar() -> str:
    """Render application sidebar and return selected menu."""
    with st.sidebar:
        # Logo/Branding
        st.markdown(f"""
        <div style="text-align: center; padding: 20px 0;">
            <h1 style="font-family: 'Rajdhani', sans-serif; 
                         color: #00a8ff; font-size: 2rem; margin: 0;">
                {AppConstants.APP_ICON} WedaBay
            </h1>
            <p style="color: #7f8fa6; font-size: 0.7rem; 
                      letter-spacing: 2px; margin-top: 5px;">
                ABSENCE CENTER
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        # Navigation
        st.markdown("### 🧭 NAVIGATION")
        menu = st.radio(
            "Select Module:",
            ["📊 Dashboard", "📝 Submit Report", "⚙️ Settings"],
            label_visibility="collapsed"
        )
        
        st.markdown("---")
        
        # Quick actions
        st.markdown("### ⚡ QUICK ACTIONS")
        
        if st.button("🔄 Refresh Data", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
        
        if st.button("📥 Bulk Export", use_container_width=True):
            st.session_state['show_bulk_export'] = True
        
        st.markdown("---")
        
        # System info
        st.markdown("### 📡 SYSTEM STATUS")
        st.success("🟢 Online")
        st.caption(f"Version: {AppConstants.APP_VERSION}")
        st.caption(f"Last Updated: {datetime.now().strftime('%H:%M:%S')}")
        
        # Footer
        st.markdown("---")
        st.caption(f"© 2025 {AppConstants.COMPANY_NAME}")
        st.caption("All rights reserved")
    
    return menu


def render_settings_page() -> None:
    """Render settings and configuration page."""
    st.markdown('<div class="brand-title">SETTINGS</div>', unsafe_allow_html=True)
    st.markdown('<div class="brand-subtitle">SYSTEM CONFIGURATION</div>', 
                unsafe_allow_html=True)
    
    tab1, tab2, tab3 = st.tabs(["⚙️ General", "👥 Division Management", "📊 Reports"])
    
    with tab1:
        st.markdown("### General Settings")
        
        st.checkbox(
            "Enable Notifications",
            value=st.session_state.get('notifications_enabled', True),
            key='notifications_enabled',
            help="Show alerts for anomalies"
        )
        
        st.selectbox(
            "Default View Mode",
            ["Cards", "Table", "Analytics"],
            key='view_mode',
            help="Set default dashboard view"
        )
        
        st.slider(
            "Cards Per Row",
            min_value=2,
            max_value=6,
            value=4,
            help="Number of employee cards per row"
        )
        
        st.number_input(
            "Late Threshold (minutes after 7:00)",
            min_value=0,
            max_value=60,
            value=5,
            help="Minutes after 7:00 AM considered late"
        )
    
    with tab2:
        st.markdown("### Division Management")
        
        divisions = DivisionRegistry.get_all()
        
        for div_name, div_config in divisions.items():
            with st.expander(f"{div_config.icon} {div_name} ({div_config.code})"):
                st.markdown(f"**Description:** {div_config.description}")
                st.markdown(f"**Color:** {div_config.color}")
                st.markdown(f"**Total Members:** {len(div_config.members)}")
                
                st.markdown("**Members:**")
                for member in div_config.members:
                    st.caption(f"• {member}")
    
    with tab3:
        st.markdown("### Report Configuration")
        
        st.selectbox(
            "Default Export Format",
            ["Excel (.xlsx)", "CSV (.csv)", "JSON (.json)", "PDF (.pdf)"],
            help="Default format for exports"
        )
        
        st.checkbox(
            "Include Charts in Reports",
            value=True,
            help="Include visualization charts in exported reports"
        )
        
        st.checkbox(
            "Auto-archive Old Reports",
            value=False,
            help="Automatically archive reports older than 30 days"
        )


def main() -> None:
    """
    Main application entry point.
    Orchestrates the entire application flow.
    """
    # Configure page
    configure_page()
    
    # Initialize configurations
    ConfigurationManager.initialize_session_state()
    initialize_divisions()
    
    # Apply styling
    ThemeManager.apply_global_styles()
    
    # Initialize performance monitoring
    performance_monitor = PerformanceMonitor()
    performance_monitor.start_timer('app_load')
    
    # Render sidebar and get menu selection
    selected_menu = render_sidebar()
    
    # Initialize controller
    controller = AttendanceController()
    
    # Route to appropriate page
    try:
        if selected_menu == "📊 Dashboard":
            controller.run_dashboard()
        
        elif selected_menu == "📝 Submit Report":
            controller.run_report_form()
        
        elif selected_menu == "⚙️ Settings":
            render_settings_page()
        
        # Show performance metrics in debug mode
        load_time = performance_monitor.end_timer('app_load')
        if load_time > 0:
            st.sidebar.caption(f"⏱️ Load time: {load_time:.2f}s")
    
    except Exception as e:
        st.error(f"⚠️ Application Error: {str(e)}")
        st.exception(e)
        
        if st.button("🔄 Reload Application"):
            st.rerun()


# ================================================================================
# APPLICATION EXECUTION
# ================================================================================

if __name__ == "__main__":

    main()
















