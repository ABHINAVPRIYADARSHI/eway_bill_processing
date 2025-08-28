#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import time
import json
import subprocess
from datetime import date, datetime
import streamlit as st
import calendar
from streamlit_autorefresh import st_autorefresh

CONFIG_PATH = os.path.abspath("./input/config.json")
LOG_PATH = os.path.abspath("./input/logs.txt")
os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)


# Initialize session state for log management
if 'log_session_id' not in st.session_state:
    st.session_state.log_session_id = datetime.now().strftime('%Y%m%d_%H%M%S')
    # Clear log file at the start of a new session
    with open(LOG_PATH, "w", encoding="utf-8") as f:
        f.write(f"[SESSION {st.session_state.log_session_id}] Logging started\n")

today = date.today()

hide_toolbar_and_header = """
<style>
/* Hide the toolbar and decoration */
[data-testid="stToolbar"] {visibility: hidden; height: 0; position: fixed;}
[data-testid="stDecoration"] {visibility: hidden; height: 0; position: fixed;}
header {visibility: hidden; height: 0; position: fixed;}

/* Remove top padding/margin so content moves up */
.block-container {
    padding-top: 0rem;
    margin-top: -2rem; /* adjust this value as needed */
}
</style>
"""
st.markdown(hide_toolbar_and_header, unsafe_allow_html=True)
# ---- Custom CSS for Styling & Animations ----
st.markdown(
    """
    <style>
    /* Page Background and Font */
    .stApp {
        background: linear-gradient(135deg, #d0e8d0 0%, #c5e3fc 50%, #d0e8d0 100%);
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        min-height: 100vh;
    }

    /* Main Container */
    .main {
        background: rgba(255, 255, 255, 0.9);
        border-radius: 12px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.1);
        padding: 2rem;
        margin: 1rem auto;
        max-width: 1200px;
    }

    /* Help Tab Styling */
    [data-testid="stMarkdownContainer"] h2 {
        color: #2c3e50;
        font-family: 'Poppins', sans-serif;
        font-weight: 600;
        margin-top: 1.5em;
        margin-bottom: 0.5em;
    }

    [data-testid="stMarkdownContainer"] h3 {
        color: #3498db;
        font-family: 'Poppins', sans-serif;
        font-weight: 500;
        margin-top: 1.2em;
        margin-bottom: 0.5em;
    }

    [data-testid="stMarkdownContainer"] ul {
        margin-top: 0.5em;
        margin-bottom: 1em;
        padding-left: 1.5em;
    }

    [data-testid="stMarkdownContainer"] li {
        margin-bottom: 0.5em;
        line-height: 1.6;
        color: #4a5568;
    }

    /* Add Poppins font from Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&display=swap');

    /* Title */
    h1 {
        text-align: left;
        color: #2c3e50;
        padding: 1rem 0;
        margin: 0 0 1.5rem 0;
        background: linear-gradient(90deg, #3498db, #2ecc71);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 700;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        margin-bottom: 1.5rem;
    }

    .stTabs [data-baseweb="tab"] {
        background: #f8f9fa;
        border-radius: 8px 8px 0 0;
        padding: 0.75rem 1.5rem;
        font-weight: 600;
        color: #6c757d;
        transition: all 0.3s ease;
    }

    .stTabs [aria-selected="true"] {
        background: #3498db;
        color: white;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #f8f9fa 0%, #e9ecef 100%) !important;
        border-right: 1px solid #dee2e6;
    }

    /* Inputs and buttons */
    .stTextInput>div>div>input, .stTextArea>div>textarea {
        border: 1px solid #ced4da;
        border-radius: 8px;
        padding: 0.5rem 1rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }

    /* GSTIN Text Area */
    .stTextArea>div>div>textarea {
        min-height: 200px;  /* Reduced height */
    }

    /* Buttons */
    div.stButton > button {
        background: linear-gradient(45deg, #3498db, #2ecc71);
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        padding: 0.75rem 2rem;
        margin: 1rem 0;
        transition: all 0.3s ease;
        width: 100%;
    }

    div.stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(52, 152, 219, 0.3);
    }

    /* Logs Box */
    .log-box {
        background: #1e1e1e;
        color: #e0e0e0;
        padding: 1rem;
        height: 400px;
        border-radius: 8px;
        border: 1px solid #444;
        font-family: 'Consolas', 'Monaco', monospace;
        white-space: pre-wrap;
        overflow-y: auto;
        line-height: 1.5;
        font-size: 0.9em;
        scroll-behavior: smooth;
    }

    /* Scrollbar styling */
    .log-box::-webkit-scrollbar {
        width: 8px;
    }

    .log-box::-webkit-scrollbar-track {
        background: #2d2d2d;
        border-radius: 4px;
    }

    .log-box::-webkit-scrollbar-thumb {
        background: #555;
        border-radius: 4px;
    }

    .log-box::-webkit-scrollbar-thumb:hover {
        background: #777;
    }

    /* Status messages */
    .stAlert {
        border-radius: 8px;
    }

    /* Footer */
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background: linear-gradient(90deg, #2c3e50, #3498db);
        color: white;
        text-align: center;
        padding: 10px 0;
        font-size: 0.85rem;
        box-shadow: 0 -2px 10px rgba(0,0,0,0.1);
    }

    .footer-content {
        max-width: 1200px;
        margin: 0 auto;
        padding: 0 1rem;
    }

    /* Responsive adjustments */
    @media (max-width: 768px) {
        .stTabs [data-baseweb="tab"] {
            padding: 0.5rem 1rem;
            font-size: 0.9em;
        }
        
        .log-box {
            height: 300px;
        }
    }

    /* Force st.code() background to black and text to light gray */
    pre, pre code {
        background-color: #2d2d2d !important;
        color: #f5f5f5 !important;
        border-radius: 8px;
        font-family: 'Consolas', 'Monaco', monospace !important;
        font-size: 0.95em !important;
    }
    </style>
    """, unsafe_allow_html=True
)

# ---- Load Config ----
def load_config(path):
    try:
        with open(path, 'r') as f:
            config = json.load(f)
            return {
                "url": "https://gstsso.nic.in/",
                "username": config.get("username", ""),
                "password": config.get("password", ""),
                "gstins": config.get("gstins", []),
                "start_month": config.get("start_month", calendar.month_name[today.month]),
                "end_month": config.get("end_month", calendar.month_name[today.month]),
                "start_year": config.get("start_year", today.year-1),
                "end_year": config.get("end_year", today.year),
                "extract_ewb_data_flag": config.get("extract_ewb_data_flag", True),
                "prepare_stock_statement_flag": config.get("prepare_stock_statement_flag", True),
                "check_toll_data_flag": config.get("check_toll_data_flag", True)
            }
    except (FileNotFoundError, json.JSONDecodeError):
        return {"url": "https://gstsso.nic.in/", "username": "", "password": "", "gstins": [],
                "start_month": calendar.month_name[today.month], "end_month": calendar.month_name[today.month],
                "start_year": today.year-1, "end_year": today.year, 
                "extract_ewb_data_flag": True, "prepare_stock_statement_flag": True, "check_toll_data_flag": True}


def run_worker(config_path, log_path):
    worker_script = get_script_path("scraper_worker.py")
    if getattr(sys, "frozen", False):
        # When bundled with PyInstaller, call the separate worker exe
        worker_exe = os.path.join(os.path.dirname(sys.executable), "scraper_worker.exe")
        subprocess.Popen([worker_exe, config_path, log_path])
    else:
        # When running locally (normal Python), call the script directly
        subprocess.Popen([sys.executable, worker_script, config_path, log_path])


def get_script_path(script_name: str) -> str:
    """Return correct path to the worker script both when running source or as PyInstaller exe."""
    if getattr(sys, 'frozen', False):
        # When bundled with PyInstaller
        base_path = sys._MEIPASS
    else:
        # When running normally from source
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, script_name)

config = load_config(CONFIG_PATH)
st.set_page_config(page_title="E-Way Bill Scraper UI", layout="wide")
st.title("🚚 E-Way Bill Mitra")

# Main content with tabs
tab1, tab2, tab3 = st.tabs(["📋 Configuration", "📜 Live Logs", "❓Help"])

with tab1:
    # Two column layout with login on left, GSTIN on right
    col1, col2 = st.columns([1, 1])
    
    with col1:
        # Login section
        st.subheader("🔑 Login Details")
        url = st.text_input("Login URL", value=config["url"])
        username = st.text_input("Username", value=config["username"])
        password = st.text_input("Password", type="password", value=config["password"])
        
        # Date Range below login
        st.subheader("📅 Date Range")
        years = list(range(2017, today.year + 1))
        months = {m: i for i, m in enumerate(calendar.month_name) if m}
        
        col_y1, col_y2 = st.columns(2)
        with col_y1:
            start_year = st.selectbox("Start Year", years, index=years.index(config["start_year"]))
            start_month_options = list(months.keys()) if start_year < today.year else list(months.keys())[:today.month]
            start_month = st.selectbox("Start Month", start_month_options, index=start_month_options.index(config["start_month"]))
        with col_y2:
            end_year = st.selectbox("End Year", years, index=years.index(config["end_year"]))
            end_month_options = list(months.keys()) if end_year < today.year else list(months.keys())[:today.month]
            end_month = st.selectbox("End Month", end_month_options, index=end_month_options.index(config["end_month"]))
        
        # Date validation
        months_map = {m: i for i, m in enumerate(calendar.month_name) if m}
        start_dt = date(start_year, months_map[start_month], 1)
        end_dt = date(end_year, months_map[end_month], calendar.monthrange(end_year, months_map[end_month])[1])
        
        if end_dt < start_dt:
            st.error("⚠️ End date cannot be earlier than start date.")
        else:
            st.success(f"📅 Selected: {start_dt.strftime('%d %b %Y')} to {end_dt.strftime('%d %b %Y')}")
    
    with col2:
        # GSTIN section
        st.subheader("📌 GSTINs")
        gstins_prefill = "\n".join(config["gstins"])
        gstins_str = st.text_area(
            "Enter GSTINs (comma or newline separated)", 
            value=gstins_prefill, 
            height=200  # Reduced height
        )
        gstins = [g.strip() for g in gstins_str.replace("\n", ",").split(",") if g.strip()]
        
        # Operations section below GSTIN
        st.subheader("⚙️ Operations")
        op_col1, op_col2, op_col3 = st.columns(3)
        with op_col1:
            st.markdown("<div style='background: #f8f9fa; padding: 1rem; border-radius: 8px; border-left: 4px solid #3498db;'>", unsafe_allow_html=True)
            extract_ewb_data_flag = st.checkbox("Extract EWB Data", value=config["extract_ewb_data_flag"])
            st.markdown("</div>", unsafe_allow_html=True)
        
        with op_col2:
            st.markdown("<div style='background: #f8f9fa; padding: 1rem; border-radius: 8px; border-left: 4px solid #2ecc71;'>", unsafe_allow_html=True)
            prepare_stock_statement_flag = st.checkbox("Prepare Stock Statement", value=config["prepare_stock_statement_flag"])
            st.markdown("</div>", unsafe_allow_html=True)
        
        with op_col3:
            st.markdown("<div style='background: #f8f9fa; padding: 1rem; border-radius: 8px; border-left: 4px solid #9b59b6;'>", unsafe_allow_html=True)
            check_toll_data_flag = st.checkbox("Check Toll Data", value=config["check_toll_data_flag"])
            st.markdown("</div>", unsafe_allow_html=True)
    
    # Start button centered below both columns
    st.markdown("<div style='text-align: center; margin: 2rem 0;'>", unsafe_allow_html=True)
    run = st.button("🚀 Start Scraping", use_container_width=True, type="primary")
    st.markdown("</div>", unsafe_allow_html=True)

    if run:
        if not username or not password or not gstins:
            st.error("❌ Please fill Username, Password, and GSTINs.")
            st.stop()
        if end_dt < start_dt:
            st.error("❌ End date must be the same or after start date.")
            st.stop()
        st.success("✅ Scraping started! Switch to the 'Live Logs' tab to monitor progress.")

        config_data = {
            "url": url,
            "username": username,
            "password": password,
            "gstins": gstins,
            "start_month": start_month,
            "end_month": end_month,
            "start_year": start_year,
            "end_year": end_year,
            "start_date": start_dt.isoformat(),
            "end_date": end_dt.isoformat(),
            "extract_ewb_data_flag": extract_ewb_data_flag,
            "prepare_stock_statement_flag": prepare_stock_statement_flag,
            "check_toll_data_flag": check_toll_data_flag
        }
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config_data, f, indent=2)
        # Add a clear separator when starting a new scraping job
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"\n[SCRAPER {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Starting new scraping job\n")
        run_worker(CONFIG_PATH, LOG_PATH)
        # subprocess.Popen([sys.executable, "scraper_worker.py", CONFIG_PATH, LOG_PATH])
        # subprocess.Popen([sys.executable, get_script_path("scraper_worker.py"), CONFIG_PATH, LOG_PATH])


# Footer
st.markdown(
    """
    <div class='footer'>
        <div class='footer-content'>
            <div>&copy; 2025 All rights reserved. Designed & developed by CGST & Central Excise, Bhilai (v 2.0.0)</div>
        </div>
    </div>
    """, 
    unsafe_allow_html=True
)

with tab3:
    st.markdown("""
    ## 📚 User Guide
    
    Welcome to the E-Way Bill Mitra! Follow these steps to use the application:
    
    ### 1. Login & Setup
    - Enter your E-Way Bill portal credentials (Username and Password)
    - Select the date range for which you want to download data
    - Enter GSTIN numbers (one per line or comma separated)
    - Select the operations you want to perform (extract EWB data, prepare stock statement, check toll data etc.)
    
    ### 2. Starting the Scraper
    - Click the "Start Scraping" button
    - A new Chromium browser window will open automatically.
    - The system will automatically navigate to the E-Way Bill portal
    
    ### 3. Manual Steps (Required)
    - When prompted, enter the CAPTCHA manually
    - Enter the OTP received on your registered mobile/email
    - The scraper will handle the rest automatically
    
    ### 4. Monitoring Progress
    - Monitor the progress in the "Logs" tab which gets refreshed every 5 seconds
    - Downloaded files will be available in the `./E-Way Mitra/output` folder
    - Each GSTIN will have its own subfolder with the downloaded files
    
    ### ❗ Important Notes
    - Keep the browser window open until the process completes
    - Do not interact with the automated browser window
    - In case of errors, check the logs for details
    - For best results, ensure a stable internet connection
    
    ### 🛠️ Troubleshooting
    - If the browser doesn't open automatically, try:
      1. Check if a browser window is already open in the background
      2. Manually open a browser and navigate to `http://localhost:8501`
      3. Ensure your browser is up to date
    - If you encounter CAPTCHA issues, try refreshing the page and starting again
    - For persistent issues, check the logs and contact support
    """, unsafe_allow_html=True)


# Logs Tab
with tab2:
    st.subheader("📜 Live Logs")
    st.caption("🔄 Auto-refreshes every 5 seconds")
    # Auto-refresh component - this replaces the problematic time.sleep + st.rerun
    count = st_autorefresh(interval=5000, limit=None, key="log_refresh")

    # Initialize session state for logs if it doesn't exist
    if 'log_history' not in st.session_state:
        st.session_state.log_history = ""
    
    # Create a container for logs
    log_container = st.container()
    
    # Function to read new log entries
    def read_new_logs():
        if not os.path.exists(LOG_PATH):
            return None
            
        try:
            with open(LOG_PATH, "r", encoding="utf-8") as f:
                new_logs = f.read()
                return new_logs if new_logs else None
        except Exception as e:
            st.error(f"Error reading logs: {str(e)}")
            return None
    
    # Display logs with auto-refresh
    with log_container:
        # Read new logs
        current_logs = read_new_logs()
        
        # If we have new logs, update the history
        if current_logs:
            st.session_state.log_history = current_logs
        
        # Always display the full log history
        if st.session_state.log_history:
            st.code(st.session_state.log_history, language="bash")
        else:
            st.info("No logs available yet. Start the scraper to see logs here.")
    
    # # Auto-refresh the tab every 5 seconds
    # time.sleep(5)
    # st.rerun()
