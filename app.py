"""
Birla Opus – Employee Wellness Arena
Pure SQLite | No Excel dependency | Full CRUD Admin | Health & Wellbeing
"""
import streamlit as st
import sqlite3
import pandas as pd
import plotly.graph_objects as go
import hashlib, os, io
from datetime import datetime, date

# ── CONFIG ────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Birla Opus · Wellness Arena",
    page_icon="🏆",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={}
)

DB_PATH    = "wellness.db"
ADMIN_PASS = hashlib.sha256("birlaopus2024".encode()).hexdigest()

# Brand palette
CR  = "#C0392B"   # Birla red
CDR = "#922B21"   # Dark red
CG  = "#D4AC0D"   # Gold
CO  = "#E67E22"   # Orange
CGR = "#27AE60"   # Green (health)
CDG = "#1A5276"   # Dark green/teal
CW  = "#FFFFFF"
CLG = "#F8F4EF"   # Warm light bg
CBS = "#2C3E50"   # Body text
CGY = "#7F8C8D"   # Gray
CBR = "#E8DDD0"   # Border warm

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,700;0,900;1,400&family=DM+Sans:wght@300;400;500;600;700&family=Montserrat:wght@400;600;700;800;900&display=swap');

html,body,[class*="css"]{{font-family:'DM Sans',sans-serif;color:{CBS};}}
.stApp{{background:linear-gradient(150deg,#FFFFFF 0%,#FDF9F5 40%,#F8F2EA 100%);background-attachment:fixed;}}

/* ── Sidebar ── */
[data-testid="stSidebar"] > div:first-child {{
    background:linear-gradient(180deg,{CDR} 0%,{CR} 50%,#A93226 100%) !important;
    padding-top:0 !important;
}}
/* Sidebar text — targeted, not wildcard */
[data-testid="stSidebar"] .stMarkdown p,
[data-testid="stSidebar"] .stMarkdown div,
[data-testid="stSidebar"] .stMarkdown span {{
    color:#fff !important;
}}
/* Radio label text specifically */
[data-testid="stSidebar"] .stRadio label {{
    color:#fff !important;
    -webkit-text-fill-color:#fff !important;
}}
[data-testid="stSidebar"] .stRadio label p {{
    color:#fff !important;
    -webkit-text-fill-color:#fff !important;
    font-family:'Montserrat',sans-serif !important;
    font-size:0.78rem !important;
    font-weight:600 !important;
}}
/* Radio option rows */
[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label {{
    background:rgba(255,255,255,0.1) !important;
    border:1px solid rgba(255,255,255,0.18) !important;
    border-radius:10px !important;
    padding:10px 14px !important;
    margin-bottom:5px !important;
    display:flex !important;
    align-items:center !important;
    cursor:pointer !important;
}}
[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:hover {{
    background:rgba(255,255,255,0.2) !important;
}}
/* Button in sidebar */
[data-testid="stSidebar"] .stButton > button {{
    background:rgba(255,255,255,0.15) !important;
    color:#fff !important;
    border:1px solid rgba(255,255,255,0.3) !important;
}}
[data-testid="stSidebar"] .stButton > button:hover {{
    background:rgba(255,255,255,0.25) !important;
    transform:none !important;
}}

/* ── Page Hero ── */
.hero{{
    background:linear-gradient(135deg,{CDR} 0%,{CR} 55%,#A04030 100%);
    border-radius:20px; padding:30px 36px; margin-bottom:28px;
    position:relative; overflow:hidden;
    box-shadow:0 8px 32px rgba(192,57,43,0.28);
}}
.hero::before{{
    content:''; position:absolute; top:-60px; right:-40px;
    width:280px; height:280px; border-radius:50%;
    background:radial-gradient(circle,rgba(255,255,255,0.07),transparent 65%);
}}
.hero::after{{
    content:''; position:absolute; bottom:-50px; left:10%;
    width:180px; height:180px; border-radius:50%;
    background:radial-gradient(circle,rgba(212,172,13,0.14),transparent 65%);
}}
.hero-sport{{background:linear-gradient(135deg,#1A3A6B 0%,#1F618D 55%,#2980B9 100%);box-shadow:0 8px 32px rgba(31,97,141,0.32);}}
.hero-health{{background:linear-gradient(135deg,#0E5E3A 0%,#1A7A4A 55%,#27AE60 100%);box-shadow:0 8px 32px rgba(39,174,96,0.28);}}
.hero-query{{background:linear-gradient(135deg,#4A235A 0%,#6C3483 55%,#8E44AD 100%);box-shadow:0 8px 32px rgba(142,68,173,0.28);}}
.hero-admin{{background:linear-gradient(135deg,#212529 0%,#343A40 55%,#495057 100%);box-shadow:0 8px 32px rgba(0,0,0,0.25);}}
.hero-title{{font-family:'Playfair Display',serif;font-size:2.4rem;font-weight:900;color:#fff;margin:0;line-height:1.1;position:relative;z-index:1;}}
.hero-sub{{font-family:'Montserrat',sans-serif;font-size:0.7rem;font-weight:700;color:rgba(255,255,255,0.65);letter-spacing:3.5px;text-transform:uppercase;margin-top:7px;position:relative;z-index:1;}}

/* ── Cards ── */
.card{{background:{CW};border:1px solid {CBR};border-radius:16px;padding:22px 24px;margin-bottom:14px;
       box-shadow:0 2px 12px rgba(0,0,0,0.055);transition:all 0.25s;position:relative;overflow:hidden;}}
.card::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px;
               background:linear-gradient(90deg,{CR},{CG},{CO});}}
.card:hover{{box-shadow:0 8px 28px rgba(0,0,0,0.1);transform:translateY(-2px);}}
.card-flat{{background:{CW};border:1px solid {CBR};border-radius:12px;padding:16px 18px;margin-bottom:10px;box-shadow:0 1px 5px rgba(0,0,0,0.045);}}
.card-green::before{{background:linear-gradient(90deg,{CGR},#2ECC71,#A9DFBF);}}
.card-blue::before{{background:linear-gradient(90deg,#1F618D,#2980B9,#85C1E9);}}

/* ── KPI ── */
.kpi{{background:{CW};border:1px solid {CBR};border-radius:14px;padding:20px 16px;
      text-align:center;box-shadow:0 2px 10px rgba(0,0,0,0.055);position:relative;overflow:hidden;}}
.kpi::after{{content:'';position:absolute;bottom:0;left:0;right:0;height:3px;
             background:linear-gradient(90deg,{CR},{CG});}}
.kpi-v{{font-family:'Playfair Display',serif;font-size:2.2rem;font-weight:700;color:{CDR};line-height:1;margin-bottom:5px;}}
.kpi-l{{font-family:'Montserrat',sans-serif;font-size:0.62rem;font-weight:800;color:{CGY};text-transform:uppercase;letter-spacing:2px;}}
.kpi-s{{font-size:0.74rem;color:{CO};font-weight:600;margin-top:3px;}}

/* ── Section header ── */
.sh{{font-family:'Montserrat',sans-serif;font-size:0.68rem;font-weight:800;color:{CR};
     text-transform:uppercase;letter-spacing:3px;margin-bottom:13px;padding-bottom:8px;
     border-bottom:2px solid {CBR};}}
.sh-green{{color:{CGR};border-bottom-color:#A9DFBF;}}
.sh-blue{{color:#1F618D;border-bottom-color:#AED6F1;}}

/* ── Podium ── */
.podium{{display:flex;align-items:flex-end;justify-content:center;gap:14px;padding:8px 0 0;}}
.pod{{display:flex;flex-direction:column;align-items:center;}}
.pod-av{{width:68px;height:68px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:1.9rem;margin-bottom:9px;}}
.pod1 .pod-av{{background:linear-gradient(135deg,#F9E94E,{CG});border:3px solid {CG};box-shadow:0 4px 16px rgba(212,172,13,0.4);}}
.pod2 .pod-av{{background:linear-gradient(135deg,#E8E8E8,#A0A0A0);border:3px solid #A0A0A0;box-shadow:0 4px 14px rgba(0,0,0,0.15);}}
.pod3 .pod-av{{background:linear-gradient(135deg,#EDBB99,#A04000);border:3px solid #A04000;box-shadow:0 4px 14px rgba(160,64,0,0.3);}}
.pod-nm{{font-family:'Montserrat',sans-serif;font-weight:700;font-size:0.82rem;color:{CBS};text-align:center;max-width:100px;line-height:1.3;}}
.pod-dp{{font-size:0.64rem;color:{CGY};text-align:center;margin-top:2px;}}
.pod-sc{{font-family:'Playfair Display',serif;font-size:1.05rem;font-weight:700;margin:5px 0 5px;}}
.pod1 .pod-sc{{color:#B7950B;}} .pod2 .pod-sc{{color:#666;}} .pod3 .pod-sc{{color:#7E5109;}}
.pod-st{{width:96px;border-radius:10px 10px 0 0;display:flex;align-items:center;justify-content:center;font-family:'Playfair Display',serif;font-size:1.5rem;font-weight:900;color:#fff;}}
.pod1 .pod-st{{height:96px;background:linear-gradient(180deg,{CG},#9A7D0A);}}
.pod2 .pod-st{{height:70px;background:linear-gradient(180deg,#A0A0A0,#606060);}}
.pod3 .pod-st{{height:52px;background:linear-gradient(180deg,#CD7F32,#7E5109);}}

/* ── LB row ── */
.lb{{display:flex;align-items:center;padding:10px 14px;border-radius:10px;margin-bottom:5px;
     background:{CLG};border:1px solid {CBR};transition:all 0.2s;}}
.lb:hover{{background:#FEF0EC;border-color:#E8A090;transform:translateX(3px);}}

/* ── Schedule pills ── */
.sc-done{{background:#F0FFF4;border:1px solid #A9DFBF;border-radius:10px;padding:12px 16px;margin-bottom:8px;}}
.sc-live{{background:#FFF0EE;border:1.5px solid {CR};border-radius:10px;padding:12px 16px;margin-bottom:8px;animation:pulse-border 2s infinite;}}
.sc-soon{{background:#FFFBEC;border:1px solid #F9E94E;border-radius:10px;padding:12px 16px;margin-bottom:8px;}}
@keyframes pulse-border{{0%,100%{{border-color:{CR}88}}50%{{border-color:{CR}}}}}

/* ── BMI bars ── */
.bmi-normal{{background:#E9F7EF;border-left:4px solid {CGR};border-radius:8px;padding:10px 14px;margin-bottom:6px;}}
.bmi-over{{background:#FEF9E7;border-left:4px solid #F39C12;border-radius:8px;padding:10px 14px;margin-bottom:6px;}}
.bmi-obese{{background:#FDEDEC;border-left:4px solid {CR};border-radius:8px;padding:10px 14px;margin-bottom:6px;}}
.bmi-under{{background:#EAF2FF;border-left:4px solid #2980B9;border-radius:8px;padding:10px 14px;margin-bottom:6px;}}

/* ── Quick Stats Cards ── */
@keyframes qs-shimmer {{
    0% {{ background-position: -200% center; }}
    100% {{ background-position: 200% center; }}
}}
@keyframes qs-pop {{
    0% {{ transform: scale(0.92); opacity:0; }}
    100% {{ transform: scale(1); opacity:1; }}
}}
.qs-card {{
    border-radius: 10px;
    padding: 10px 13px;
    margin-bottom: 7px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    animation: qs-pop 0.4s ease forwards;
    background-size: 200% auto;
    transition: background-position 0.5s ease;
}}
.qs-card:hover {{
    background-position: right center;
}}
.qs-c1 {{ background: linear-gradient(135deg, rgba(212,172,13,0.35), rgba(192,57,43,0.25)); border-left: 3px solid {CG}; }}
.qs-c2 {{ background: linear-gradient(135deg, rgba(39,174,96,0.3), rgba(26,82,118,0.2)); border-left: 3px solid {CGR}; animation-delay:0.05s; }}
.qs-c3 {{ background: linear-gradient(135deg, rgba(41,128,185,0.3), rgba(39,174,96,0.2)); border-left: 3px solid #2980B9; animation-delay:0.1s; }}
.qs-c4 {{ background: linear-gradient(135deg, rgba(142,68,173,0.3), rgba(192,57,43,0.2)); border-left: 3px solid #8E44AD; animation-delay:0.15s; }}
.qs-c5 {{ background: linear-gradient(135deg, rgba(192,57,43,0.35), rgba(212,172,13,0.25)); border-left: 3px solid {CR}; animation-delay:0.2s; }}
.qs-lbl {{ font-family:'Montserrat',sans-serif; font-size:0.62rem; color:rgba(255,255,255,0.7); font-weight:600; letter-spacing:0.5px; }}
.qs-val {{ font-family:'Playfair Display',serif; font-size:1.05rem; font-weight:700; color:#fff; }}

/* ── Health animations ── */
@keyframes heartbeat{{0%,100%{{transform:scale(1)}}14%{{transform:scale(1.15)}}28%{{transform:scale(1)}}42%{{transform:scale(1.1)}}70%{{transform:scale(1)}}}}
.heart-anim{{animation:heartbeat 1.5s ease-in-out infinite;display:inline-block;}}
@keyframes sport-bounce{{0%,100%{{transform:translateY(0)}}50%{{transform:translateY(-8px)}}}}
.sport-anim{{animation:sport-bounce 1.2s ease-in-out infinite;display:inline-block;}}
@keyframes fade-up{{from{{opacity:0;transform:translateY(20px)}}to{{opacity:1;transform:translateY(0)}}}}
.fade-up{{animation:fade-up 0.5s ease forwards;}}
@keyframes shimmer-bar{{0%{{background-position:-200% 0}}100%{{background-position:200% 0}}}}

/* ── Ticker ── */
.ticker{{background:linear-gradient(135deg,{CDR},{CR});border-radius:12px;padding:13px 18px;
         margin-bottom:9px;color:#fff;position:relative;overflow:hidden;
         box-shadow:0 3px 12px rgba(192,57,43,0.22);}}
.ticker::after{{content:'● LIVE';position:absolute;top:10px;right:13px;
                font-size:0.58rem;font-weight:800;letter-spacing:2px;color:rgba(255,255,255,0.75);
                animation:blink 1.4s ease-in-out infinite;}}
@keyframes blink{{0%,100%{{opacity:1}}50%{{opacity:0.25}}}}

/* ── Inputs ── */
.stButton>button{{background:linear-gradient(135deg,{CR},{CDR}) !important;color:#fff !important;
    border:none !important;font-family:'Montserrat',sans-serif !important;font-size:0.7rem !important;
    font-weight:700 !important;letter-spacing:1px !important;text-transform:uppercase !important;
    border-radius:8px !important;padding:10px 20px !important;transition:all 0.25s !important;}}
.stButton>button:hover{{transform:translateY(-2px) !important;box-shadow:0 6px 20px rgba(192,57,43,0.35) !important;}}
[data-testid="stTextInput"] input,[data-testid="stNumberInput"] input{{
    border:1.5px solid {CBR} !important;border-radius:8px !important;
    background:{CW} !important;color:{CBS} !important;font-family:'DM Sans',sans-serif !important;}}
[data-testid="stTextInput"] input:focus,[data-testid="stNumberInput"] input:focus{{
    border-color:{CR} !important;box-shadow:0 0 0 3px rgba(192,57,43,0.1) !important;}}
[data-testid="stTabs"] [role="tab"]{{font-family:'Montserrat',sans-serif !important;font-size:0.66rem !important;
    font-weight:700 !important;letter-spacing:1.5px !important;text-transform:uppercase !important;color:{CGY} !important;}}
[data-testid="stTabs"] [role="tab"][aria-selected="true"]{{color:{CR} !important;border-bottom:2px solid {CR} !important;}}
[data-testid="stDataFrame"]{{border-radius:10px;overflow:hidden;}}
#MainMenu,footer{{visibility:hidden;}}
[data-testid="stDecoration"]{{display:none;}}
/* Hide header bar AND the sidebar collapse button — sidebar is always open */
header{{visibility:hidden;}}
/* Hide collapse button — all states including hover */
[data-testid="collapsedControl"],
[data-testid="stSidebarCollapsedControl"],
button[data-testid="baseButton-header"],
[data-testid="stSidebar"] [data-testid="stBaseButton-header"],
[data-testid="stSidebar"] button[kind="header"],
[data-testid="stSidebar"]:hover [data-testid="stBaseButton-header"],
[data-testid="stSidebar"]:hover button[kind="header"],
[data-testid="stSidebar"] button[aria-label="Close sidebar"],
[data-testid="stSidebar"] button[aria-label="Collapse sidebar"],
section[data-testid="stSidebar"] > div > div > div > button{{
    display:none !important;
    visibility:hidden !important;
    opacity:0 !important;
    pointer-events:none !important;
    width:0 !important;
    height:0 !important;
    overflow:hidden !important;
}}

::-webkit-scrollbar{{width:6px;height:6px;}}
::-webkit-scrollbar-track{{background:{CLG};}}
::-webkit-scrollbar-thumb{{background:#C8A090;border-radius:3px;}}

/* ── Table centre alignment ── */
[data-testid="stDataFrame"] th {{
    text-align: center !important;
    font-family: 'Montserrat', sans-serif !important;
    font-size: 0.68rem !important;
    font-weight: 800 !important;
    color: {CBS} !important;
    letter-spacing: 1px !important;
    text-transform: uppercase !important;
    background: {CLG} !important;
}}
[data-testid="stDataFrame"] td {{
    text-align: center !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.82rem !important;
}}
[data-testid="stDataFrame"] {{
    border-radius: 12px !important;
    overflow: hidden !important;
    border: 1px solid {CBR} !important;
}}

/* ── Insight ── */
.ins{{background:linear-gradient(135deg,#FFF9F5,{CW});border:1px solid {CBR};
      border-left:4px solid {CR};border-radius:10px;padding:13px 16px;margin-bottom:8px;font-size:0.85rem;line-height:1.55;}}
.ins-green{{border-left-color:{CGR};}}

/* ── Status badge ── */
.badge-done{{display:inline-block;background:#E9F7EF;color:#1E8449;font-size:0.62rem;font-weight:800;
             letter-spacing:1.5px;text-transform:uppercase;padding:3px 10px;border-radius:20px;}}
.badge-live{{display:inline-block;background:#FDEDEC;color:{CR};font-size:0.62rem;font-weight:800;
             letter-spacing:1.5px;text-transform:uppercase;padding:3px 10px;border-radius:20px;}}
.badge-soon{{display:inline-block;background:#FEF9E7;color:#9A7D0A;font-size:0.62rem;font-weight:800;
             letter-spacing:1.5px;text-transform:uppercase;padding:3px 10px;border-radius:20px;}}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# DATABASE LAYER
# ══════════════════════════════════════════════════════════════════════════════
def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)

def init_db():
    conn = get_conn()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS Employees (
            EmployeeID INTEGER PRIMARY KEY,
            Name TEXT NOT NULL,
            Department TEXT NOT NULL,
            Designation TEXT DEFAULT '',
            JoinDate TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS Events (
            EventID INTEGER PRIMARY KEY AUTOINCREMENT,
            EventName TEXT NOT NULL UNIQUE,
            Type TEXT DEFAULT 'Indoor',
            Difficulty TEXT DEFAULT 'Casual',
            Multiplier REAL DEFAULT 1.0,
            Description TEXT DEFAULT '',
            PrizeWinner TEXT DEFAULT '',
            PrizeRunnerUp TEXT DEFAULT '',
            Prize2ndRunnerUp TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS Participation (
            PID INTEGER PRIMARY KEY AUTOINCREMENT,
            EmployeeID INTEGER, EmployeeName TEXT, EventName TEXT, Date TEXT,
            Position TEXT, Registered TEXT DEFAULT 'Yes', Participated TEXT DEFAULT 'Yes',
            ParticipationPoints REAL DEFAULT 10,
            BasePoints REAL DEFAULT 0,
            Multiplier REAL DEFAULT 1.0,
            GamePoints REAL DEFAULT 0,
            FinalPoints REAL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS Scores (
            ScoreID INTEGER PRIMARY KEY AUTOINCREMENT,
            EmployeeID INTEGER UNIQUE,
            EmployeeName TEXT, Department TEXT,
            Score REAL DEFAULT 0,
            LastUpdated TEXT
        );
        CREATE TABLE IF NOT EXISTS Schedule (
            SID INTEGER PRIMARY KEY AUTOINCREMENT,
            EventName TEXT, StartTime TEXT,
            Status TEXT DEFAULT 'Upcoming',
            Venue TEXT DEFAULT '', Notes TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS BMI (
            BID INTEGER PRIMARY KEY AUTOINCREMENT,
            EmployeeID INTEGER,
            EmployeeName TEXT,
            Department TEXT,
            Month TEXT,
            Year INTEGER,
            Weight_kg REAL,
            Height_cm REAL,
            BMI REAL,
            Category TEXT,
            RecordedOn TEXT
        );
    """)
    conn.commit()
    # Migrate existing DB: add prize columns to Events if not present
    for col in ["PrizeWinner TEXT DEFAULT ''", "PrizeRunnerUp TEXT DEFAULT ''", "Prize2ndRunnerUp TEXT DEFAULT ''"]:
        try: conn.execute(f"ALTER TABLE Events ADD COLUMN {col}")
        except: pass
    conn.commit()
    conn.close()

def seed_initial_data(path="Employee_Wellness_Scoring_System.xlsx"):
    """
    Smart incremental sync from Excel → SQLite.
    Checks every sheet and inserts ONLY new records (never deletes existing data).
    Returns (True, report_dict) or (False, error_string).
    """
    if not os.path.exists(path):
        return False, f"Excel not found at: {os.path.abspath(path)}"
    try:
        xl   = pd.read_excel(path, sheet_name=None)
        conn = get_conn()
        report = {}

        # ── 1. Employee_Master ────────────────────────────────────────────────
        if 'Employee_Master' in xl:
            emp = xl['Employee_Master'][['Employee_ID','Employee_Name','Department']].dropna(subset=['Employee_ID']).copy()
            emp.columns = ['EmployeeID','Name','Department']
            emp['EmployeeID'] = pd.to_numeric(emp['EmployeeID'], errors='coerce').dropna().astype(int)
            emp = emp.dropna(subset=['EmployeeID'])

            # Fetch existing IDs once for fast lookup
            existing_ids = {r[0] for r in conn.execute("SELECT EmployeeID FROM Employees").fetchall()}

            new_emp = 0
            for _, r in emp.iterrows():
                eid = int(r['EmployeeID'])
                if eid not in existing_ids:
                    conn.execute(
                        "INSERT INTO Employees (EmployeeID,Name,Department,Designation,JoinDate) VALUES (?,?,?,?,?)",
                        (eid, r['Name'], r['Department'], '', '')
                    )
                    existing_ids.add(eid)
                    new_emp += 1
            report['Employees'] = {'new': new_emp, 'skipped': len(emp) - new_emp}

        # ── 2. Event_Master ───────────────────────────────────────────────────
        if 'Event_Master' in xl:
            ev = xl['Event_Master'].dropna(subset=['Event_ID']).copy()

            existing_events = {r[0].strip().lower() for r in conn.execute("SELECT EventName FROM Events").fetchall()}

            new_ev = 0
            for _, r in ev.iterrows():
                ev_name = str(r['Event_Name']).strip()
                if ev_name.lower() not in existing_events:
                    conn.execute(
                        "INSERT INTO Events (EventName,Type,Difficulty,Multiplier) VALUES (?,?,?,?)",
                        (ev_name, r['Type'], r['Difficulty_Level (Casual/Medium/High)'], r['Multiplier'])
                    )
                    existing_events.add(ev_name.lower())
                    new_ev += 1
            report['Events'] = {'new': new_ev, 'skipped': len(ev) - new_ev}

        # ── 3. Participation_Entry ────────────────────────────────────────────
        if 'Participation_Entry' in xl:
            part = xl['Participation_Entry'].dropna(subset=['Employee_ID']).copy()
            part.columns = ['Date','EmployeeID','EmployeeName','EventName','Registered','Participated',
                            'Position','ParticipationPoints','BasePoints','Multiplier','GamePoints','FinalPoints']
            part['EmployeeID'] = pd.to_numeric(part['EmployeeID'], errors='coerce').dropna().astype(int)
            part = part.dropna(subset=['EmployeeID'])
            for c in ['ParticipationPoints','BasePoints','Multiplier','GamePoints','FinalPoints']:
                part[c] = pd.to_numeric(part[c], errors='coerce').fillna(0)

            # Build a set of (EmployeeID, EventName, Date) tuples already in DB
            existing_part = {
                (r[0], str(r[1]).strip().lower(), str(r[2]))
                for r in conn.execute("SELECT EmployeeID, EventName, Date FROM Participation").fetchall()
            }

            new_part = skipped_part = 0
            updated_employees = set()
            for _, r in part.iterrows():
                key = (int(r['EmployeeID']), str(r['EventName']).strip().lower(), str(r['Date']))
                if key not in existing_part:
                    conn.execute(
                        """INSERT INTO Participation
                           (EmployeeID,EmployeeName,EventName,Date,Position,Registered,Participated,
                            ParticipationPoints,BasePoints,Multiplier,GamePoints,FinalPoints)
                           VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (int(r['EmployeeID']), r['EmployeeName'], r['EventName'], str(r['Date']),
                         r['Position'], r['Registered'], r['Participated'],
                         r['ParticipationPoints'], r['BasePoints'], r['Multiplier'],
                         r['GamePoints'], r['FinalPoints'])
                    )
                    existing_part.add(key)
                    updated_employees.add(int(r['EmployeeID']))
                    new_part += 1
                else:
                    skipped_part += 1
            report['Participation'] = {'new': new_part, 'skipped': skipped_part}

            # Recalculate scores only for employees who got new participation rows
            for eid in updated_employees:
                recalc_score(eid, conn)

        # ── 4. Leaderboard (Scores) ───────────────────────────────────────────
        if 'Leaderboard' in xl:
            lb = xl['Leaderboard'][['Employee_ID','Employee_Name','Department','Total_Points']].dropna(subset=['Employee_ID']).copy()
            lb['Employee_ID'] = pd.to_numeric(lb['Employee_ID'], errors='coerce').dropna().astype(int)
            lb = lb.dropna(subset=['Employee_ID'])

            # Only update scores for employees NOT already recalculated above
            already_recalced = report.get('Participation', {}).get('new', 0) > 0
            existing_score_ids = {r[0] for r in conn.execute("SELECT EmployeeID FROM Scores").fetchall()}

            new_sc = upd_sc = 0
            for _, r in lb.iterrows():
                eid = int(r['Employee_ID'])
                if eid not in existing_score_ids:
                    # Genuinely new employee score — insert it
                    conn.execute(
                        "INSERT OR REPLACE INTO Scores (EmployeeID,EmployeeName,Department,Score,LastUpdated) VALUES (?,?,?,?,?)",
                        (eid, r['Employee_Name'], r['Department'], r['Total_Points'], date.today().isoformat())
                    )
                    existing_score_ids.add(eid)
                    new_sc += 1
                else:
                    # Employee exists — only update if Excel score is higher (new data added in Excel)
                    current = conn.execute("SELECT Score FROM Scores WHERE EmployeeID=?", (eid,)).fetchone()
                    if current and float(r['Total_Points']) > float(current[0] or 0):
                        conn.execute(
                            "UPDATE Scores SET Score=?,LastUpdated=? WHERE EmployeeID=?",
                            (r['Total_Points'], date.today().isoformat(), eid)
                        )
                        upd_sc += 1
            report['Leaderboard'] = {'new': new_sc, 'updated': upd_sc, 'skipped': len(lb) - new_sc - upd_sc}

        # ── 5. Schedule ───────────────────────────────────────────────────────
        if 'Schedule' in xl:
            sched = xl['Schedule'].dropna(how='all').iloc[:, :2].copy()
            sched.columns = ['EventName','StartTime']
            sched = sched.dropna(subset=['EventName'])

            existing_sched = {r[0].strip().lower() for r in conn.execute("SELECT EventName FROM Schedule").fetchall()}

            new_sch = 0
            for _, r in sched.iterrows():
                ev_name = str(r['EventName']).strip()
                if ev_name.lower() not in existing_sched:
                    conn.execute(
                        "INSERT INTO Schedule (EventName,StartTime,Status) VALUES (?,?,?)",
                        (ev_name, str(r['StartTime']), 'Upcoming')
                    )
                    existing_sched.add(ev_name.lower())
                    new_sch += 1
            report['Schedule'] = {'new': new_sch, 'skipped': len(sched) - new_sch}

        conn.commit()
        conn.close()
    except Exception as e:
        return False, str(e)
    return True, report

def recalc_score(employee_id, conn):
    """Recalculate total score from all participation entries."""
    total = conn.execute(
        "SELECT COALESCE(SUM(FinalPoints),0) FROM Participation WHERE EmployeeID=?", (employee_id,)
    ).fetchone()[0]
    emp = conn.execute("SELECT Name,Department FROM Employees WHERE EmployeeID=?", (employee_id,)).fetchone()
    if emp:
        conn.execute("""INSERT OR REPLACE INTO Scores (EmployeeID,EmployeeName,Department,Score,LastUpdated)
            VALUES (?,?,?,?,?)""", (employee_id, emp[0], emp[1], total, date.today().isoformat()))

def bmi_category(bmi):
    if bmi < 18.5: return "Underweight"
    elif bmi < 25:  return "Normal"
    elif bmi < 30:  return "Overweight"
    else:           return "Obese"

def bmi_color(cat):
    return {"Underweight":"#2980B9","Normal":CGR,"Overweight":"#F39C12","Obese":CR}.get(cat, CGY)

# ── Queries ───────────────────────────────────────────────────────────────────
def q(sql, params=()):
    conn = get_conn()
    df = pd.read_sql(sql, conn, params=params if params else None)
    conn.close(); return df

def get_leaderboard():
    return q("""SELECT s.EmployeeID, s.EmployeeName AS Name, s.Department,
                       COALESCE(s.Score,0) AS TotalPoints
                FROM Scores s ORDER BY TotalPoints DESC""")

def get_dept_stats():
    emp = q("SELECT Department,COUNT(*) as Total FROM Employees GROUP BY Department")
    par = q("""SELECT e.Department, COUNT(DISTINCT p.EmployeeID) as Participated
               FROM Participation p JOIN Employees e ON p.EmployeeID=e.EmployeeID
               WHERE p.Participated='Yes' GROUP BY e.Department""")
    m = emp.merge(par, on='Department', how='left').fillna(0)
    # Force numeric — fillna(0) can sometimes produce object dtype after merge
    m['Total']       = pd.to_numeric(m['Total'],       errors='coerce').fillna(0).astype(int)
    m['Participated'] = pd.to_numeric(m['Participated'], errors='coerce').fillna(0).astype(int)
    m['Pct'] = (m['Participated'] / m['Total'].replace(0, 1) * 100).round(1)
    return m.sort_values('Total', ascending=False)

def get_game_winners():
    return q("""SELECT EventName,Position,EmployeeName,EmployeeID,FinalPoints,Date
                FROM Participation
                WHERE Position IN ('Winner','Runner-up','2nd Runner-up')
                ORDER BY EventName,
                  CASE Position WHEN 'Winner' THEN 1 WHEN 'Runner-up' THEN 2 ELSE 3 END""")

def get_summary():
    conn = get_conn()
    s = {}
    s['emp']    = conn.execute("SELECT COUNT(*) FROM Employees").fetchone()[0]
    s['part']   = conn.execute("SELECT COUNT(DISTINCT EmployeeID) FROM Participation WHERE Participated='Yes'").fetchone()[0]
    s['events'] = conn.execute("SELECT COUNT(*) FROM Events").fetchone()[0]
    s['top']    = conn.execute("SELECT COALESCE(MAX(Score),0) FROM Scores").fetchone()[0]
    conn.close()
    s['pct'] = round(s['part']/s['emp']*100,1) if s['emp'] else 0
    return s

# ══════════════════════════════════════════════════════════════════════════════
# INIT + SESSION
# ══════════════════════════════════════════════════════════════════════════════
init_db()
conn_chk = get_conn()
emp_count = conn_chk.execute("SELECT COUNT(*) FROM Employees").fetchone()[0]
conn_chk.close()
if emp_count == 0:
    # Try common paths where Excel might be placed
    import pathlib
    possible_paths = [
        "Employee_Wellness_Scoring_System.xlsx",
        str(pathlib.Path(__file__).parent / "Employee_Wellness_Scoring_System.xlsx"),
    ]
    for p in possible_paths:
        if os.path.exists(p):
            seed_initial_data(p)
            break

for k,v in [('admin_auth',False),('gemini_key',''),('nlq',''),('last_refresh',datetime.now())]:
    if k not in st.session_state: st.session_state[k] = v

# Smart auto-refresh JS (pauses on Admin page & when user is typing)
st.markdown("""<script>
(function(){
    var T=60000,timer=null;
    function isAdmin(){
        var r=document.querySelectorAll('[data-testid="stSidebar"] [role="radio"]');
        for(var i=0;i<r.length;i++){if(r[i].getAttribute('aria-checked')==='true'&&r[i].innerText.includes('Admin'))return true;}
        return false;
    }
    function busy(){
        var e=document.activeElement;
        return e&&(e.tagName==='INPUT'||e.tagName==='TEXTAREA'||e.tagName==='SELECT'||e.isContentEditable);
    }
    function go(){clearTimeout(timer);timer=setTimeout(function(){
        if(isAdmin()||busy()){go();return;}
        window.location.reload();
    },T);}
    ['mousemove','mousedown','keydown','touchstart','click'].forEach(function(e){
        document.addEventListener(e,go,{passive:true});
    });
    go();
})();
</script>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
    <div style="text-align:center;padding:20px 0 14px;">
        <div style="width:68px;height:68px;background:#fff;border-radius:14px;
                    margin:0 auto 11px;display:flex;align-items:center;justify-content:center;
                    box-shadow:0 4px 16px rgba(0,0,0,0.22);">
            <b style="font-family:'Playfair Display',serif;font-size:2rem;color:#C0392B;
                      -webkit-text-fill-color:#C0392B;">B</b>
        </div>
        <div style="font-family:'Playfair Display',serif;font-size:1.35rem;font-weight:900;
                    color:#fff;letter-spacing:1px;-webkit-text-fill-color:#fff;">Birla Opus</div>
        <div style="font-family:'Montserrat',sans-serif;font-size:0.56rem;letter-spacing:3px;
                    text-transform:uppercase;margin-top:3px;color:rgba(255,255,255,0.65);
                    -webkit-text-fill-color:rgba(255,255,255,0.65);">Wellness Arena</div>
    </div>
    <hr style="border-color:rgba(255,255,255,0.18);margin:0 0 14px;">
    """, unsafe_allow_html=True)

    st.markdown('<div style="font-family:Montserrat,sans-serif;font-size:0.6rem;font-weight:800;color:rgba(255,255,255,0.55);letter-spacing:2px;text-transform:uppercase;margin-bottom:8px;">NAVIGATION</div>', unsafe_allow_html=True)
    page = st.radio(
        "Navigation",
        options=[
            "🏆  Dashboard",
            "📅  Schedule",
            "🏥  Health & Wellbeing",
            "💬  Smart Query",
            "🔒  Admin Portal",
        ],
        index=0,
        key="main_nav",
        label_visibility="collapsed"
    )

    s = get_summary()
    qs_items = [
        ("qs-c1","👥","Employees",    str(s['emp'])),
        ("qs-c2","🏃","Participants",  str(s['part'])),
        ("qs-c3","📈","Part. Rate",    f"{s['pct']}%"),
        ("qs-c4","⚽","Events Held",   str(s['events'])),
        ("qs-c5","🏆","Top Score",     str(int(s['top']))),
    ]
    qs_html = '''<div style="margin-top:14px;">
        <div style="font-family:Montserrat,sans-serif;font-size:0.58rem;font-weight:800;
                    color:rgba(255,255,255,0.45);letter-spacing:2.5px;text-transform:uppercase;
                    margin-bottom:10px;">QUICK STATS</div>'''
    for cls, icon, lbl, val in qs_items:
        qs_html += f'''<div class="qs-card {cls}">
            <div><div class="qs-lbl">{icon} {lbl}</div></div>
            <div class="qs-val">{val}</div>
        </div>'''
    qs_html += "</div>"
    st.markdown(qs_html, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    is_admin = "Admin" in page
    if not is_admin:
        if st.button("🔄  Refresh Data", use_container_width=True):
            st.session_state.last_refresh = datetime.now(); st.rerun()
    else:
        st.markdown('<div style="text-align:center;background:rgba(255,255,255,0.09);border-radius:8px;padding:10px;font-size:0.68rem;color:rgba(255,255,255,0.6);">🔒 Refresh paused on Admin</div>', unsafe_allow_html=True)

    lr = st.session_state.last_refresh.strftime("%d %b, %H:%M:%S")
    st.markdown(f'<div style="text-align:center;font-size:0.6rem;color:rgba(255,255,255,0.45);margin-top:6px;">Last refreshed: {lr}</div>', unsafe_allow_html=True)
    if not is_admin:
        st.markdown('<div style="text-align:center;font-size:0.58rem;color:rgba(255,255,255,0.35);margin-top:2px;">⟳ Auto every 60s</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 1 — DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
if "Dashboard" in page:
    st.markdown("""
    <div class="hero fade-up">
        <div style="font-size:2rem;margin-bottom:6px;position:relative;z-index:1;">
            <span class="sport-anim">🏆</span>
        </div>
        <div class="hero-title">Performance Dashboard</div>
        <div class="hero-sub">Birla Opus · Wellness Arena · Live Standings</div>
    </div>""", unsafe_allow_html=True)

    lb  = get_leaderboard()
    ds  = get_dept_stats()
    gw  = get_game_winners()
    ps  = q("SELECT EventName,COUNT(*) as Entries,AVG(FinalPoints) as AvgPts FROM Participation GROUP BY EventName ORDER BY Entries DESC")

    # KPI row
    k1,k2,k3,k4 = st.columns(4)
    top_dept = ds.sort_values('Pct',ascending=False).iloc[0]['Department'] if len(ds) else "—"
    for col,(val,lbl,sub,c) in zip([k1,k2,k3,k4],[
        (s['emp'],   "Total Employees",   "",              CR),
        (s['part'],  "Participants",       f"{s['pct']}% rate", CG),
        (f"{s['pct']}%","Participation Rate",f"{s['part']}/{s['emp']} staff",CGR),
        (s['events'],"Events Held",       f"{top_dept} leads",  "#1F618D"),
    ]):
        with col:
            st.markdown(f'<div class="kpi"><div class="kpi-v" style="color:{c};">{val}</div>'
                        f'<div class="kpi-l">{lbl}</div>'
                        f'{f"<div class=kpi-s>{sub}</div>" if sub else ""}</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Podium + Top 3 Depts
    cp, cd = st.columns([1.2,1], gap="large")
    with cp:
        st.markdown('<div class="sh">🏅 Champions Podium</div>', unsafe_allow_html=True)
        top3 = lb.head(3).reset_index(drop=True)
        cfg = [(1,"pod2","🥈","2"),(0,"pod1","🥇","1"),(2,"pod3","🥉","3")]
        html = '<div class="card"><div class="podium">'
        for si,cls,em,num in cfg:
            if si < len(top3):
                r = top3.iloc[si]
                nm = "<br>".join(str(r['Name']).split()[:2])
                html += (f'<div class="pod {cls}"><div class="pod-av">{em}</div>'
                         f'<div class="pod-nm">{nm}</div><div class="pod-dp">{r["Department"]}</div>'
                         f'<div class="pod-sc">{int(r["TotalPoints"])} pts</div>'
                         f'<div class="pod-st">{num}</div></div>')
        html += '</div></div>'
        st.markdown(html, unsafe_allow_html=True)

    with cd:
        st.markdown('<div class="sh">🏢 Top 3 Departments · Participation</div>', unsafe_allow_html=True)
        top3d = ds.sort_values('Pct',ascending=False).head(3).reset_index(drop=True)
        rc = [CG,"#A0A0A0","#CD7F32"]; re = ["🥇","🥈","🥉"]
        for i, row in top3d.iterrows():
            c = rc[i] if i<3 else CR; e = re[i] if i<3 else "▸"
            bg = ["#FFFBEC","#F8F8F8","#FFF5EE"][i] if i<3 else "#FFF"
            st.markdown(
                f'<div style="background:{bg};border:1.5px solid {c}44;border-left:4px solid {c};'
                f'border-radius:10px;padding:12px 16px;margin-bottom:8px;">'
                f'<div style="display:flex;justify-content:space-between;align-items:center;">'
                f'<div><div style="font-family:Montserrat,sans-serif;font-size:0.6rem;color:{c};font-weight:800;letter-spacing:1.5px;">{e} RANK {i+1}</div>'
                f'<div style="font-weight:700;font-size:0.9rem;color:{CBS};margin-top:2px;">{row["Department"]}</div>'
                f'<div style="font-size:0.7rem;color:{CGY};margin-top:2px;">{int(row["Participated"])} of {int(row["Total"])} employees</div></div>'
                f'<div style="text-align:right;"><div style="font-family:Playfair Display,serif;font-size:1.5rem;font-weight:700;color:{c};">{row["Pct"]}%</div>'
                f'<div style="width:72px;height:4px;background:#EEE;border-radius:2px;margin-top:4px;overflow:hidden;">'
                f'<div style="width:{min(row["Pct"],100)}%;height:100%;background:{c};border-radius:2px;"></div></div></div>'
                f'</div></div>', unsafe_allow_html=True)

    # Dept bar chart
    st.markdown('<div class="sh" style="margin-top:4px;">📊 Department Participation Overview</div>', unsafe_allow_html=True)
    ds2 = ds.sort_values('Total', ascending=False)
    fig = go.Figure()
    fig.add_trace(go.Bar(name='Total Employees', x=ds2['Department'], y=ds2['Total'],
        marker=dict(color='#E8DDD0',line=dict(width=0)),
        text=ds2['Total'], textposition='outside', textfont=dict(size=10,color=CGY,family='DM Sans')))
    fig.add_trace(go.Bar(name='Participated', x=ds2['Department'], y=ds2['Participated'],
        marker=dict(color=CR,line=dict(width=0)),
        text=ds2['Participated'], textposition='outside', textfont=dict(size=10,color=CR,family='DM Sans')))
    fig.update_layout(barmode='group', paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        xaxis=dict(tickfont=dict(color=CGY,size=9,family='DM Sans'),gridcolor='rgba(0,0,0,0.04)',tickangle=-30),
        yaxis=dict(tickfont=dict(color=CGY,size=9,family='DM Sans'),gridcolor='rgba(0,0,0,0.05)'),
        legend=dict(font=dict(color=CGY,size=10,family='DM Sans'),bgcolor='rgba(0,0,0,0)',orientation='h',x=0,y=1.12),
        margin=dict(l=10,r=10,t=30,b=70),height=340,bargap=0.22,bargroupgap=0.06)
    st.markdown('<div class="card" style="padding:14px;">', unsafe_allow_html=True)
    st.plotly_chart(fig, use_container_width=True, config={'displayModeBar':False})
    st.markdown('</div>', unsafe_allow_html=True)

    # Game Winners
    st.markdown('<div class="sh">🎮 Event Winners — Gold · Silver · Bronze</div>', unsafe_allow_html=True)
    ev_list = sorted(gw['EventName'].unique()) if len(gw) else []
    pos_cfg = {
        'Winner':       ('🥇',CG,'#FFFBEC',f'{CG}55'),
        'Runner-up':    ('🥈','#909090','#F8F8F8','#90909044'),
        '2nd Runner-up':('🥉','#CD7F32','#FFF5EE','#CD7F3244'),
    }
    if ev_list:
        for rs in range(0,len(ev_list),3):
            batch = ev_list[rs:rs+3]
            gcols = st.columns(len(batch))
            for col, evn in zip(gcols, batch):
                with col:
                    ev_df = gw[gw['EventName']==evn].reset_index(drop=True)
                    ch = (f'<div class="card" style="padding:15px 17px;min-height:180px;">'
                          f'<div style="font-family:Montserrat,sans-serif;font-size:0.7rem;font-weight:800;'
                          f'color:{CR};letter-spacing:2px;text-transform:uppercase;margin-bottom:11px;'
                          f'padding-bottom:7px;border-bottom:2px solid {CBR};">⚽ {evn}</div>')
                    for _, w in ev_df.iterrows():
                        em,cl,bg,bd = pos_cfg.get(w['Position'],('▸',CGY,'#F8F8F8','#EEE'))
                        ch += (f'<div style="display:flex;align-items:center;gap:9px;margin-bottom:7px;'
                               f'background:{bg};border:1.5px solid {bd};border-radius:8px;padding:8px 11px;">'
                               f'<span style="font-size:1.2rem;">{em}</span>'
                               f'<div style="flex:1;"><div style="font-size:0.82rem;font-weight:700;color:{CBS};line-height:1.2;">{w["EmployeeName"]}</div>'
                               f'<div style="font-size:0.62rem;color:{CGY};letter-spacing:1px;text-transform:uppercase;">{w["Position"]}</div></div>'
                               f'<div style="font-family:Playfair Display,serif;font-size:0.9rem;font-weight:700;color:{cl};">{int(w["FinalPoints"])}pts</div>'
                               f'</div>')
                    ch += '</div>'
                    st.markdown(ch, unsafe_allow_html=True)

    # Live insights + upcoming
    st.markdown('<div class="sh" style="margin-top:4px;">⚡ Live Highlights</div>', unsafe_allow_html=True)
    ic1, ic2 = st.columns(2)
    insights = []
    if len(lb): insights.append(f"🏆 <b>{lb.iloc[0]['Name']}</b> leads with <b>{int(lb.iloc[0]['TotalPoints'])} points</b> · {lb.iloc[0]['Department']}")
    td = ds.sort_values('Pct',ascending=False)
    if len(td): insights.append(f"🏢 <b>{td.iloc[0]['Department']}</b> tops participation at <b>{td.iloc[0]['Pct']}%</b>")
    if len(ps): insights.append(f"⚽ <b>{ps.iloc[0]['EventName']}</b> most contested — <b>{int(ps.iloc[0]['Entries'])} entries</b>")
    if s['part']: insights.append(f"📈 <b>{s['pct']}%</b> of all employees have joined at least one event")
    with ic1:
        for ins in insights[:3]:
            st.markdown(f'<div class="ticker"><div style="font-size:0.84rem;line-height:1.5;padding-right:55px;">{ins}</div></div>', unsafe_allow_html=True)
    with ic2:
        st.markdown(f'<div style="font-family:Montserrat,sans-serif;font-size:0.68rem;font-weight:800;color:{CR};letter-spacing:2px;text-transform:uppercase;margin-bottom:10px;">📅 Upcoming Matches</div>', unsafe_allow_html=True)
        sched_up = q("SELECT EventName,StartTime,Status FROM Schedule ORDER BY StartTime ASC LIMIT 5")
        if len(sched_up):
            cards = ""
            for _, row in sched_up.iterrows():
                ev = str(row['EventName']); st_t = str(row['StartTime'])[:16]
                stat = str(row.get('Status','Upcoming'))
                is_done = 'Complet' in stat; is_live = 'Live' in stat
                cls = 'sc-done' if is_done else ('sc-live' if is_live else 'sc-soon')
                icon = '✅' if is_done else ('🔴' if is_live else '🗓️')
                badge_cls = "done" if is_done else ("live" if is_live else "soon")
                badge = f'<span class="badge-{badge_cls}">{stat}</span>'
                cards += (f'<div class="{cls}">'
                          f'<div style="display:flex;justify-content:space-between;align-items:center;">'
                          f'<div><span style="font-size:0.9rem;">{icon}</span>'
                          f'<span style="font-weight:700;font-size:0.88rem;color:{CBS};margin-left:8px;">{ev}</span></div>'
                          f'{badge}</div>'
                          f'<div style="font-size:0.7rem;color:{CGY};margin-top:3px;margin-left:26px;">📅 {st_t}</div>'
                          f'</div>')
            st.markdown(cards, unsafe_allow_html=True)
        else:
            st.markdown('<div class="ins">No schedule data yet. Add via Admin Portal.</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 2 — SCHEDULE
# ══════════════════════════════════════════════════════════════════════════════
elif "Schedule" in page:
    st.markdown("""
    <div class="hero hero-sport fade-up">
        <div style="font-size:2rem;margin-bottom:6px;position:relative;z-index:1;">
            <span class="sport-anim">⚽</span>&nbsp;
            <span class="sport-anim" style="animation-delay:0.2s;">🏸</span>&nbsp;
            <span class="sport-anim" style="animation-delay:0.4s;">🏏</span>
        </div>
        <div class="hero-title">Event Schedule</div>
        <div class="hero-sub">Birla Opus · Upcoming & Completed Competitions</div>
    </div>""", unsafe_allow_html=True)

    sched = q("SELECT * FROM Schedule ORDER BY StartTime ASC")
    ev_names = q("SELECT EventName FROM Events")['EventName'].tolist()

    # Stats
    if len(sched):
        total_s = len(sched)
        done = sched['Status'].str.contains('Complet',case=False,na=False).sum()
        upcoming = total_s - done
        for col,(val,lbl,c) in zip(st.columns(3),[
            (total_s,"Total Scheduled",CR),(done,"Completed",CGR),(upcoming,"Upcoming",CG)]):
            with col:
                st.markdown(f'<div class="kpi"><div class="kpi-v" style="color:{c};">{val}</div><div class="kpi-l">{lbl}</div></div>', unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

    cs, ca = st.columns([1.3,1], gap="large")

    with cs:
        st.markdown('<div class="sh sh-blue">📋 Full Schedule</div>', unsafe_allow_html=True)
        if len(sched):
            cards = ""
            for _, row in sched.iterrows():
                ev = str(row.get('EventName','—')); st_t = str(row.get('StartTime','—'))[:16]
                stat = str(row.get('Status','Upcoming')); ven = str(row.get('Venue',''))
                note = str(row.get('Notes',''))
                is_done = 'Complet' in stat; is_live = 'Live' in stat
                cls = 'sc-done' if is_done else ('sc-live' if is_live else 'sc-soon')
                icon = '✅' if is_done else ('🔴 LIVE' if is_live else '🗓️')
                sc = '#27AE60' if is_done else (CR if is_live else '#9A7D0A')
                badge_cls = "done" if is_done else ("live" if is_live else "soon")
                badge = f'<span class="badge-{badge_cls}">{stat}</span>'
                ven_h = f" &nbsp;·&nbsp; 📍 {ven}" if ven and ven != 'nan' else ""
                note_h = f'<div style="font-size:0.7rem;color:{CGY};margin-top:2px;">📝 {note}</div>' if note and note != 'nan' else ""
                cards += (f'<div class="{cls}" style="margin-bottom:9px;">'
                          f'<div style="display:flex;justify-content:space-between;align-items:flex-start;">'
                          f'<div style="flex:1;"><div style="font-weight:700;font-size:0.92rem;color:{CBS};">{icon} &nbsp;{ev}</div>'
                          f'<div style="font-size:0.72rem;color:{CGY};margin-top:3px;">📅 {st_t}{ven_h}</div>'
                          f'{note_h}</div>{badge}</div></div>')
            st.markdown(cards, unsafe_allow_html=True)
        else:
            st.markdown('<div class="ins">No schedule yet. Add matches using the form →</div>', unsafe_allow_html=True)

    with ca:
        st.markdown('<div class="sh sh-blue">📅 Upcoming Event Results</div>', unsafe_allow_html=True)
        gw2 = get_game_winners()
        if len(gw2):
            ev_f = st.selectbox("Filter by event:", ["All Events"]+sorted(gw2['EventName'].unique().tolist()), key="sched_ef")
            filtered = gw2 if ev_f=="All Events" else gw2[gw2['EventName']==ev_f]
            for _, row in filtered.iterrows():
                em = {'Winner':'🥇','Runner-up':'🥈','2nd Runner-up':'🥉'}.get(row['Position'],'▸')
                st.markdown(
                    f'<div class="card-flat" style="padding:10px 14px;margin-bottom:6px;">'
                    f'<div style="display:flex;justify-content:space-between;align-items:center;">'
                    f'<div>{em} <b>{row["EmployeeName"]}</b> <span style="color:{CGY};font-size:0.8rem;">· {row["EventName"]}</span></div>'
                    f'<div style="font-family:Playfair Display,serif;font-weight:700;color:{CR};">{int(row["FinalPoints"])} pts</div>'
                    f'</div></div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="ins">No results recorded yet.</div>', unsafe_allow_html=True)

        # Public can VIEW but only admin can ADD — show info
        st.markdown(f'<div class="ins" style="margin-top:14px;font-size:0.8rem;color:{CGY};">ℹ️ To add or edit schedule entries, please use the <b>Admin Portal</b>.</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 3 — HEALTH & WELLBEING
# ══════════════════════════════════════════════════════════════════════════════
elif "Health" in page:
    st.markdown("""
    <div class="hero hero-health fade-up">
        <div style="font-size:2rem;margin-bottom:6px;position:relative;z-index:1;">
            <span class="heart-anim">❤️</span>&nbsp;
            <span class="heart-anim" style="animation-delay:0.3s;">💪</span>&nbsp;
            <span class="heart-anim" style="animation-delay:0.6s;">🌿</span>
        </div>
        <div class="hero-title">Health & Wellbeing</div>
        <div class="hero-sub">Birla Opus · BMI Tracking · Month-on-Month Wellness</div>
    </div>""", unsafe_allow_html=True)

    bmi_df = q("SELECT * FROM BMI ORDER BY Year, Month")
    emps   = q("SELECT EmployeeID, Name, Department FROM Employees ORDER BY Name")

    tab_overview, tab_bmi = st.tabs(["🏥 Department Overview", "📊 BMI Tracker"])

    with tab_bmi:
        if len(bmi_df) == 0:
            st.markdown(f"""
            <div class="card card-green" style="text-align:center;padding:36px;">
                <div style="font-size:3rem;margin-bottom:12px;"><span class="heart-anim">💚</span></div>
                <div style="font-family:'Playfair Display',serif;font-size:1.3rem;font-weight:700;color:{CGR};margin-bottom:8px;">No BMI Data Yet</div>
                <div style="font-size:0.88rem;color:{CGY};line-height:1.6;">
                    BMI data will appear here once Admin uploads the monthly Excel sheet.<br>
                    Admin can upload via <b>Admin Portal → BMI Upload tab</b>.
                </div>
            </div>""", unsafe_allow_html=True)
        else:
            # Employee selector
            emp_names = sorted(bmi_df['EmployeeName'].unique().tolist())
            col_sel, col_yr = st.columns([2,1])
            with col_sel:
                sel_emp = st.selectbox("🔍 Select Employee to view BMI trend:", emp_names, key="bmi_emp")
            with col_yr:
                yrs = sorted(bmi_df['Year'].unique().tolist(), reverse=True)
                sel_yr = st.selectbox("Year:", yrs, key="bmi_yr")

            emp_bmi = bmi_df[(bmi_df['EmployeeName']==sel_emp) & (bmi_df['Year']==sel_yr)].copy()

            # Month order
            month_order = ["January","February","March","April","May","June",
                           "July","August","September","October","November","December"]
            emp_bmi['MonthNum'] = emp_bmi['Month'].apply(
                lambda m: month_order.index(m)+1 if m in month_order else 0)
            emp_bmi = emp_bmi.sort_values('MonthNum')

            if len(emp_bmi):
                # Latest reading
                latest = emp_bmi.iloc[-1]
                cat = latest['Category']; bmi_val = round(latest['BMI'],1)
                cat_color = bmi_color(cat)

                # Summary tiles
                st1,st2,st3,st4 = st.columns(4)
                for col,(val,lbl,c) in zip([st1,st2,st3,st4],[
                    (f"{bmi_val}", "Current BMI",    cat_color),
                    (cat,          "Category",        cat_color),
                    (f"{latest['Weight_kg']} kg", "Weight (latest)", CGR),
                    (f"{latest['Height_cm']} cm", "Height",          "#1F618D"),
                ]):
                    with col:
                        st.markdown(f'<div class="kpi"><div class="kpi-v" style="color:{c};font-size:1.8rem;">{val}</div><div class="kpi-l">{lbl}</div></div>', unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown(f'<div class="sh sh-green">📈 {sel_emp} — BMI Trend {sel_yr}</div>', unsafe_allow_html=True)

                # Color each bar by category
                bar_colors = [bmi_color(c) for c in emp_bmi['Category']]

                fig_bmi = go.Figure()
                # Reference bands
                fig_bmi.add_hrect(y0=0,    y1=18.5, fillcolor="#EAF2FF", opacity=0.35, line_width=0)
                fig_bmi.add_hrect(y0=18.5, y1=25,   fillcolor="#E9F7EF", opacity=0.35, line_width=0)
                fig_bmi.add_hrect(y0=25,   y1=30,   fillcolor="#FEF9E7", opacity=0.35, line_width=0)
                fig_bmi.add_hrect(y0=30,   y1=45,   fillcolor="#FDEDEC", opacity=0.35, line_width=0)
                # Ref lines
                for y,lbl,c in [(18.5,"Underweight ↔ Normal","#2980B9"),
                                (25,"Normal ↔ Overweight","#F39C12"),
                                (30,"Overweight ↔ Obese",CR)]:
                    fig_bmi.add_hline(y=y, line_dash="dot", line_color=c, line_width=1.5,
                                      annotation_text=lbl, annotation_position="right",
                                      annotation_font=dict(color=c,size=9))
                # Bars
                fig_bmi.add_trace(go.Bar(
                    x=emp_bmi['Month'], y=emp_bmi['BMI'],
                    marker=dict(color=bar_colors, line=dict(width=0)),
                    text=[f"{round(v,1)}" for v in emp_bmi['BMI']],
                    textposition='outside',
                    textfont=dict(size=11,family='DM Sans'),
                    customdata=list(zip(emp_bmi['Weight_kg'],emp_bmi['Height_cm'],emp_bmi['Category'])),
                    hovertemplate=(
                        "<b>%{x}</b><br>"
                        "BMI: <b>%{y:.1f}</b><br>"
                        "Weight: %{customdata[0]} kg<br>"
                        "Height: %{customdata[1]} cm<br>"
                        "Category: <b>%{customdata[2]}</b><extra></extra>"
                    ),
                    name="BMI"
                ))
                # Trend line
                if len(emp_bmi) > 1:
                    fig_bmi.add_trace(go.Scatter(
                        x=emp_bmi['Month'], y=emp_bmi['BMI'],
                        mode='lines+markers',
                        line=dict(color=CBS, width=2, dash='dot'),
                        marker=dict(size=7, color=bar_colors,
                                    line=dict(width=2,color=CBS)),
                        name='Trend', hoverinfo='skip'
                    ))

                fig_bmi.update_layout(
                    paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                    xaxis=dict(title='Month',tickfont=dict(color=CGY,size=10,family='DM Sans'),
                               gridcolor='rgba(0,0,0,0.04)',title_font=dict(color=CGY)),
                    yaxis=dict(title='BMI',tickfont=dict(color=CGY,size=10,family='DM Sans'),
                               gridcolor='rgba(0,0,0,0.05)',title_font=dict(color=CGY),
                               range=[0, max(emp_bmi['BMI'].max()*1.18, 32)]),
                    showlegend=False,
                    margin=dict(l=10,r=80,t=20,b=10), height=380,
                )
                st.plotly_chart(fig_bmi, use_container_width=True, config={'displayModeBar':False})

                # Color legend
                st.markdown(f"""
                <div style="display:flex;gap:16px;flex-wrap:wrap;margin-top:4px;padding:12px 16px;
                             background:{CW};border:1px solid {CBR};border-radius:10px;">
                    {"".join([
                        f'<div style="display:flex;align-items:center;gap:6px;">'
                        f'<div style="width:14px;height:14px;border-radius:3px;background:{c};"></div>'
                        f'<span style="font-size:0.78rem;color:{CBS};">{lbl}</span></div>'
                        for c,lbl in [("#2980B9","🔵 Underweight (<18.5)"),(CGR,"🟢 Normal (18.5–24.9)"),
                                      ("#F39C12","🟡 Overweight (25–29.9)"),(CR,"🔴 Obese (30+)")]
                    ])}
                </div>""", unsafe_allow_html=True)

                # Monthly detail table
                st.markdown(f'<div class="sh sh-green" style="margin-top:18px;">📋 Monthly Records</div>', unsafe_allow_html=True)
                show_df = emp_bmi[['Month','Year','Weight_kg','Height_cm','BMI','Category']].copy()
                show_df['BMI'] = show_df['BMI'].round(2)
                st.dataframe(show_df, use_container_width=True, hide_index=True)
            else:
                st.markdown(f'<div class="ins ins-green">No BMI data for <b>{sel_emp}</b> in <b>{sel_yr}</b>.</div>', unsafe_allow_html=True)

    with tab_overview:
        if len(bmi_df) == 0:
            st.markdown('<div class="ins ins-green">No BMI data uploaded yet.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="sh sh-green">🏥 Department BMI Overview (Latest readings)</div>', unsafe_allow_html=True)

            # Get latest BMI per employee
            latest_bmi = bmi_df.sort_values(['Year','Month']).groupby('EmployeeID').last().reset_index()

            # Category counts
            cat_counts = latest_bmi['Category'].value_counts().reset_index()
            cat_counts.columns = ['Category','Count']
            cat_colors_map = {'Normal':CGR,'Underweight':'#2980B9','Overweight':'#F39C12','Obese':CR}

            ov1, ov2 = st.columns(2)

            # Build charts first, then render in columns
            fig_donut = go.Figure(go.Pie(
                labels=cat_counts['Category'], values=cat_counts['Count'], hole=0.62,
                marker=dict(colors=[cat_colors_map.get(c,'#999') for c in cat_counts['Category']],
                            line=dict(color=CW,width=2)),
                textfont=dict(color=CBS,size=10,family='DM Sans'),
                textinfo='percent+label',
            ))
            fig_donut.add_annotation(text=f"<b>{len(latest_bmi)}</b><br>Employees",
                x=0.5,y=0.5,xref='paper',yref='paper',showarrow=False,
                font=dict(color=CBS,size=12,family='DM Sans'))
            fig_donut.update_layout(paper_bgcolor='rgba(0,0,0,0)',
                legend=dict(font=dict(color=CGY,size=9),bgcolor='rgba(0,0,0,0)'),
                margin=dict(l=10,r=10,t=40,b=10),height=300,
                title=dict(text="BMI Category Distribution",font=dict(color=CBS,size=13,family='DM Sans'),x=0.5))

            dept_bmi = latest_bmi.groupby('Department')['BMI'].mean().round(1).reset_index()
            dept_bmi = dept_bmi.sort_values('BMI',ascending=False)
            dept_colors = [bmi_color(bmi_category(v)) for v in dept_bmi['BMI']]
            fig_dept_bmi = go.Figure(go.Bar(
                x=dept_bmi['Department'], y=dept_bmi['BMI'],
                marker=dict(color=dept_colors,line=dict(width=0)),
                text=dept_bmi['BMI'].astype(str), textposition='outside',
                textfont=dict(size=10,color=CGY,family='DM Sans'),
            ))
            fig_dept_bmi.add_hline(y=25,line_dash="dot",line_color="#F39C12",line_width=1.5)
            fig_dept_bmi.update_layout(
                paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',
                xaxis=dict(tickfont=dict(color=CGY,size=9),tickangle=-30,gridcolor='rgba(0,0,0,0.04)'),
                yaxis=dict(tickfont=dict(color=CGY,size=9),gridcolor='rgba(0,0,0,0.05)',range=[0,35]),
                margin=dict(l=10,r=10,t=40,b=60),height=300,
                title=dict(text="Avg BMI by Department",font=dict(color=CBS,size=13,family='DM Sans'),x=0.5))

            with ov1:
                st.plotly_chart(fig_donut, use_container_width=True, config={'displayModeBar':False})
            with ov2:
                st.plotly_chart(fig_dept_bmi, use_container_width=True, config={'displayModeBar':False})

            # Category breakdown tiles
            st.markdown('<div class="sh sh-green" style="margin-top:8px;">BMI Category Breakdown</div>', unsafe_allow_html=True)
            for cat,c in [("Normal",CGR),("Overweight","#F39C12"),("Obese",CR),("Underweight","#2980B9")]:
                grp = latest_bmi[latest_bmi['Category']==cat]
                if len(grp):
                    bmi_cls = f'bmi-{"normal" if cat=="Normal" else "over" if cat=="Overweight" else "obese" if cat=="Obese" else "under"}'
                    names = ", ".join(grp['EmployeeName'].tolist()[:8])
                    if len(grp) > 8: names += f" +{len(grp)-8} more"
                    st.markdown(
                        f'<div class="{bmi_cls}"><b style="color:{c};">{cat}</b> '
                        f'<span style="font-size:0.8rem;color:{CGY};">({len(grp)} employees)</span><br>'
                        f'<span style="font-size:0.78rem;color:{CBS};">{names}</span></div>',
                        unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 4 — SMART QUERY
# ══════════════════════════════════════════════════════════════════════════════
elif "Query" in page:
    st.markdown("""
    <div class="hero hero-query fade-up">
        <div style="font-size:2rem;margin-bottom:6px;position:relative;z-index:1;">💬</div>
        <div class="hero-title">Smart Query</div>
        <div class="hero-sub">Ask anything · Powered by Gemini Flash (Free AI)</div>
    </div>""", unsafe_allow_html=True)

    SCHEMA = """SQLite DB for Employee Wellness:
- Employees(EmployeeID, Name, Department, Designation, JoinDate)
- Events(EventID, EventName, Type, Difficulty, Multiplier)
- Participation(PID, EmployeeID, EmployeeName, EventName, Date, Position, FinalPoints)
- Scores(ScoreID, EmployeeID, EmployeeName, Department, Score, LastUpdated)
- Schedule(SID, EventName, StartTime, Status, Venue, Notes)
- BMI(BID, EmployeeID, EmployeeName, Department, Month, Year, Weight_kg, Height_cm, BMI, Category)
Return ONLY valid SQLite SELECT. No markdown. Max 50 rows. QUESTION: {q}"""

    def ask_gemini(txt, key):
        import urllib.request, json
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={key}"
        body = json.dumps({"contents":[{"parts":[{"text":SCHEMA.format(q=txt)}]}],
                           "generationConfig":{"temperature":0.1,"maxOutputTokens":512}}).encode()
        req = urllib.request.Request(url,data=body,headers={"Content-Type":"application/json"},method="POST")
        with urllib.request.urlopen(req,timeout=15) as r: data=json.loads(r.read())
        raw = data["candidates"][0]["content"]["parts"][0]["text"].strip().replace("```sql","").replace("```","").strip()
        if not raw.upper().startswith("SELECT"): raise ValueError(raw[:80])
        return raw

    def fallback(txt):
        t = txt.lower()
        # --- person-specific queries: extract name from query ---
        import re
        # Look for "for <name>", "of <name>", "<name>'s", "show <name>", "about <name>"
        name_match = re.search(r"(?:for|of|about|show|details of|data of|info(?:rmation)? (?:of|for|about))\s+([A-Za-z][A-Za-z\s]{2,30}?)(?:\s+(?:in|from|on|at|during|score|bmi|participation|points)|\?|$)", txt, re.IGNORECASE)
        person_name = name_match.group(1).strip() if name_match else None
        # Also try "Rajan Gupta score" pattern — name at start
        if not person_name:
            name_match2 = re.search(r"^([A-Z][a-z]+(?:\s[A-Z][a-z]+)+)", txt.strip())
            if name_match2:
                person_name = name_match2.group(1).strip()

        if person_name:
            pn = person_name.replace("'","''")  # SQL-safe
            if 'bmi' in t:
                return f"SELECT EmployeeName,Month,Year,Weight_kg,Height_cm,ROUND(BMI,2) AS BMI,Category FROM BMI WHERE EmployeeName LIKE '%{pn}%' ORDER BY Year,Month"
            if 'score' in t or 'point' in t or 'leaderboard' in t:
                return f"SELECT EmployeeName,Department,Score FROM Scores WHERE EmployeeName LIKE '%{pn}%'"
            if 'participat' in t or 'event' in t or 'played' in t:
                return f"SELECT EventName,Position,FinalPoints,Date FROM Participation WHERE EmployeeName LIKE '%{pn}%' ORDER BY Date DESC"
            # Generic — all info about the person
            return f"SELECT p.EmployeeName, p.EventName, p.Position, p.FinalPoints, p.Date FROM Participation p WHERE p.EmployeeName LIKE '%{pn}%' ORDER BY p.Date DESC"

        # --- non-person queries ---
        if 'bmi' in t: return "SELECT EmployeeName,Department,Month,Year,ROUND(BMI,2) AS BMI,Category FROM BMI ORDER BY Year DESC,Month DESC LIMIT 30"
        if 'winner' in t or 'champion' in t: return "SELECT EventName,Position,EmployeeName,FinalPoints FROM Participation WHERE Position IN ('Winner','Runner-up','2nd Runner-up') ORDER BY EventName"
        if 'schedule' in t or 'upcoming' in t: return "SELECT EventName,StartTime,Status,Venue FROM Schedule ORDER BY StartTime"
        if 'top' in t or 'leader' in t: return "SELECT EmployeeName,Department,Score FROM Scores ORDER BY Score DESC LIMIT 10"
        if 'dept' in t or 'department' in t: return "SELECT e.Department,COUNT(DISTINCT e.EmployeeID) AS Employees,COALESCE(SUM(s.Score),0) AS TotalScore FROM Employees e LEFT JOIN Scores s ON e.EmployeeID=s.EmployeeID GROUP BY e.Department ORDER BY TotalScore DESC"
        if 'event' in t: return "SELECT EventName,COUNT(*) AS Entries,SUM(FinalPoints) AS TotalPts FROM Participation GROUP BY EventName ORDER BY Entries DESC"
        if 'participat' in t: return "SELECT EmployeeName,EventName,Position,FinalPoints FROM Participation ORDER BY FinalPoints DESC LIMIT 20"
        return "SELECT EmployeeName,Department,Score FROM Scores ORDER BY Score DESC LIMIT 20"

    # Key setup
    st.markdown('<div class="sh">🔑 Gemini API Key</div>', unsafe_allow_html=True)
    with st.expander("Configure Free Gemini Key", expanded=not bool(st.session_state.gemini_key)):
        st.markdown(f'<div class="ins">🆓 <b>Gemini Flash is free.</b> Get key at <b>aistudio.google.com</b> → Sign in → "Get API Key" → starts with <code>AIzaSy...</code></div>', unsafe_allow_html=True)
        kc1,kc2 = st.columns([3,1])
        with kc1: kin = st.text_input("Key",value=st.session_state.gemini_key,type="password",placeholder="AIzaSy...",label_visibility="collapsed")
        with kc2:
            if st.button("SAVE KEY",use_container_width=True): st.session_state.gemini_key=kin.strip(); st.success("Saved!"); st.rerun()

    has_key = bool(st.session_state.gemini_key)
    sc,st_t = (CGR,"GEMINI AI ACTIVE") if has_key else (CO,"KEYWORD FALLBACK")
    st.markdown(f'<div style="display:inline-flex;align-items:center;gap:8px;padding:7px 14px;border-radius:8px;margin-bottom:18px;background:{"#F0FFF4" if has_key else "#FFF8EC"};border:1px solid {sc}44;"><span style="color:{sc};">●</span><span style="font-family:Montserrat,sans-serif;font-size:0.7rem;font-weight:800;color:{sc};">{st_t}</span></div>', unsafe_allow_html=True)

    # Auto insights
    st.markdown('<div class="sh">💡 Auto Insights</div>', unsafe_allow_html=True)
    lb2 = get_leaderboard(); ds3 = get_dept_stats()
    auto = []
    if len(lb2): auto.append(f"🥇 <b>{lb2.iloc[0]['Name']}</b> leads with <b>{int(lb2.iloc[0]['TotalPoints'])} pts</b>")
    td3 = ds3.sort_values('Pct',ascending=False)
    if len(td3): auto.append(f"🏢 <b>{td3.iloc[0]['Department']}</b> tops participation at <b>{td3.iloc[0]['Pct']}%</b>")
    bmi3 = q("SELECT Category,COUNT(*) as n FROM BMI GROUP BY Category ORDER BY n DESC")
    if len(bmi3): auto.append(f"💚 BMI: majority in <b>{bmi3.iloc[0]['Category']}</b> category ({int(bmi3.iloc[0]['n'])} records)")

    aic = st.columns(3)
    for col,ins in zip(aic,auto):
        with col: st.markdown(f'<div class="card" style="padding:15px 18px;min-height:80px;"><div style="font-size:0.84rem;line-height:1.5;">{ins}</div></div>', unsafe_allow_html=True)

    # Examples
    st.markdown('<div class="sh" style="margin-top:8px;">💬 Example Queries</div>', unsafe_allow_html=True)
    exs = [
        "Who won the Cricket event?",
        "Top 10 scorers",
        "Show Rajan Gupta score",
        "Average BMI by department",
        "Show upcoming schedule",
        "BMI data for Sushma Tiwari",
        "Participation details for AASHISH",
        "Which department has most participants?",
    ]
    ec = st.columns(4)
    for i,ex in enumerate(exs):
        with ec[i%4]:
            if st.button(ex,key=f"qex_{i}",use_container_width=True):
                st.session_state.nlq=ex; st.rerun()

    st.markdown('<div class="sh" style="margin-top:8px;">🔍 Ask Your Data</div>', unsafe_allow_html=True)
    qtxt = st.text_area("Query",value=st.session_state.nlq,
                        placeholder="e.g. 'Show me all employees in QA with score above 100'",
                        height=80,label_visibility="collapsed")
    qc1,qc2 = st.columns([1,4])
    with qc1: runq = st.button("⚡ EXECUTE",use_container_width=True)
    with qc2: show_sql = st.checkbox("Show SQL",value=True)

    if runq and qtxt.strip():
        with st.spinner("Thinking..."):
            try:
                if has_key: sql=ask_gemini(qtxt,st.session_state.gemini_key); bdg=f'<span style="background:#E9F7EF;color:{CGR};font-size:0.65rem;font-weight:800;padding:3px 8px;border-radius:4px;">✦ GEMINI</span>'
                else: sql=fallback(qtxt); bdg=f'<span style="background:#FEF9E7;color:#9A7D0A;font-size:0.65rem;font-weight:800;padding:3px 8px;border-radius:4px;">⚙ KEYWORD</span>'
            except Exception as e:
                st.error(f"Gemini error: {e}"); sql=fallback(qtxt); bdg=f'<span style="background:#FEF9E7;color:#9A7D0A;font-size:0.65rem;font-weight:800;padding:3px 8px;border-radius:4px;">⚙ FALLBACK</span>'
        if show_sql: st.markdown(f'{bdg} &nbsp;SQL:',unsafe_allow_html=True); st.code(sql,language='sql')
        try:
            res = q(sql)
            st.markdown(f'<div class="ins">✅ <b>{len(res)}</b> rows &nbsp;{bdg}</div>',unsafe_allow_html=True)
            st.dataframe(res,use_container_width=True,hide_index=True)
            if len(res)>1:
                nc=res.select_dtypes(include='number').columns.tolist()
                sc2=res.select_dtypes(exclude='number').columns.tolist()
                if sc2 and nc:
                    fg=go.Figure(go.Bar(x=res[sc2[0]].astype(str).head(20),y=res[nc[0]].head(20),
                        marker=dict(color=res[nc[0]].head(20),colorscale=[[0,'#FADBD8'],[0.5,'#E07060'],[1,CR]],line=dict(width=0)),
                        text=res[nc[0]].head(20).round(1),textposition='outside',textfont=dict(color=CGY,size=10)))
                    fg.update_layout(paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',
                        xaxis=dict(tickfont=dict(color=CGY,size=9),tickangle=-30,gridcolor='rgba(0,0,0,0.04)'),
                        yaxis=dict(tickfont=dict(color=CGY,size=9),gridcolor='rgba(0,0,0,0.05)'),
                        margin=dict(l=10,r=10,t=20,b=70),height=290)
                    st.plotly_chart(fg,use_container_width=True,config={'displayModeBar':False})
        except Exception as e: st.error(f"SQL error: {e}")

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE 5 — ADMIN PORTAL
# ══════════════════════════════════════════════════════════════════════════════
elif "Admin" in page:
    st.markdown("""
    <div class="hero hero-admin fade-up">
        <div style="font-size:2rem;margin-bottom:6px;position:relative;z-index:1;">🔒</div>
        <div class="hero-title">Admin Portal</div>
        <div class="hero-sub">Birla Opus · Secure Master Data Management</div>
    </div>""", unsafe_allow_html=True)

    # Auth
    if not st.session_state.admin_auth:
        ac1,ac2,ac3 = st.columns([1,1.2,1])
        with ac2:
            st.markdown(f'<div class="card" style="text-align:center;padding:32px 24px;"><div style="font-size:2.5rem;margin-bottom:10px;">🔐</div><div style="font-family:Playfair Display,serif;font-size:1.3rem;font-weight:700;color:{CDR};margin-bottom:6px;">Authentication Required</div><div style="font-size:0.8rem;color:{CGY};">Enter admin password to manage master data</div></div>', unsafe_allow_html=True)
            pwd = st.text_input("Password",type="password",placeholder="Enter admin password",label_visibility="collapsed")
            if st.button("🔓  LOGIN",use_container_width=True):
                if hashlib.sha256(pwd.encode()).hexdigest()==ADMIN_PASS:
                    st.session_state.admin_auth=True; st.rerun()
                else: st.error("⚠️ Incorrect password.")
            st.markdown(f'<div style="text-align:center;color:{CGY};font-size:0.7rem;margin-top:8px;">Default: <code>birlaopus2024</code></div>', unsafe_allow_html=True)
    else:
        ah1,ah2 = st.columns([5,1])
        with ah1: st.success("🔓 Admin session active · All changes save to database immediately")
        with ah2:
            if st.button("🔒 Logout"): st.session_state.admin_auth=False; st.rerun()

        tabs = st.tabs(["👥 Employees","⚽ Events","🎮 Participation","📅 Schedule","📊 Scores","🏥 BMI Upload"])

        # ── TAB 1: EMPLOYEES ──────────────────────────────────────────────────
        with tabs[0]:
            st.markdown(f'<div class="sh" style="margin-top:14px;">👥 Employee Master — Add / Edit / Delete / View</div>', unsafe_allow_html=True)

            emp_df = q("SELECT EmployeeID,Name,Department,Designation,JoinDate FROM Employees ORDER BY Department,Name")

            # Filter + Search
            ef1, ef2 = st.columns([1.5, 2])
            with ef1:
                depts = ["All"] + sorted(emp_df['Department'].unique().tolist()) if len(emp_df) else ["All"]
                dept_f = st.selectbox("Filter by Department:", depts, key="emp_df")
            with ef2:
                emp_search = st.text_input("🔍 Search by Name or ID:", placeholder="Type name or Employee ID...", key="emp_search")
            show_emp = emp_df if dept_f=="All" else emp_df[emp_df['Department']==dept_f]
            if emp_search.strip():
                mask = (show_emp['Name'].str.contains(emp_search.strip(), case=False, na=False) |
                        show_emp['EmployeeID'].astype(str).str.contains(emp_search.strip(), na=False))
                show_emp = show_emp[mask]
            st.markdown(f'<div class="ins" style="font-size:0.78rem;">Showing <b>{len(show_emp)}</b> of <b>{len(emp_df)}</b> employees</div>', unsafe_allow_html=True)
            st.dataframe(show_emp, use_container_width=True, hide_index=True,
                column_config={"EmployeeID": st.column_config.NumberColumn("Employee ID", format="%d")})

            st.markdown("---")
            op = st.radio("Action:", ["➕ Add New","✏️ Edit","🗑️ Delete"], horizontal=True, key="emp_op")

            if op == "➕ Add New":
                c1,c2,c3 = st.columns(3)
                with c1: nid = st.number_input("Employee ID",min_value=100000,max_value=999999,value=500000,key="emp_nid")
                with c2: nname = st.text_input("Full Name",placeholder="e.g. Rahul Sharma",key="emp_nn")
                with c3:
                    existing_depts = sorted(emp_df['Department'].unique().tolist()) if len(emp_df) else []
                    ndept_sel = st.selectbox("Department",existing_depts+["+ New Department"],key="emp_nd")
                ndept = ndept_sel
                if ndept_sel == "+ New Department":
                    ndept = st.text_input("New Department Name",key="emp_ndcustom")
                ndesig = st.text_input("Designation (optional)",placeholder="e.g. Engineer",key="emp_ndesig")
                if st.button("➕ ADD EMPLOYEE",use_container_width=True):
                    if nname and ndept and ndept != "+ New Department":
                        try:
                            conn = get_conn()
                            conn.execute("INSERT OR IGNORE INTO Employees (EmployeeID,Name,Department,Designation,JoinDate) VALUES (?,?,?,?,?)",
                                         (int(nid),nname.strip(),ndept.strip(),ndesig.strip(),''))
                            conn.commit(); conn.close()
                            st.success(f"✅ Added {nname} to {ndept}"); st.rerun()
                        except Exception as e: st.error(str(e))
                    else: st.warning("Fill all required fields.")

            elif op == "✏️ Edit":
                if len(emp_df):
                    opts = {f"{r['Name']} — {r['Department']} ({r['EmployeeID']})": r['EmployeeID'] for _,r in emp_df.iterrows()}
                    sel = st.selectbox("Select Employee to Edit:",list(opts.keys()),key="emp_edit_sel")
                    eid = opts[sel]
                    cur = emp_df[emp_df['EmployeeID']==eid].iloc[0]
                    ec1,ec2,ec3 = st.columns(3)
                    with ec1: uname = st.text_input("Name",value=cur['Name'],key="emp_uname")
                    with ec2:
                        all_depts = sorted(emp_df['Department'].unique().tolist())
                        udept = st.selectbox("Department",all_depts,index=all_depts.index(cur['Department']) if cur['Department'] in all_depts else 0,key="emp_udept")
                    with ec3: udesig = st.text_input("Designation",value=str(cur['Designation']) if str(cur['Designation']) != 'nan' else '',key="emp_udesig")
                    if st.button("💾 SAVE CHANGES",use_container_width=True):
                        conn=get_conn()
                        conn.execute("UPDATE Employees SET Name=?,Department=?,Designation=? WHERE EmployeeID=?",(uname,udept,udesig,eid))
                        conn.commit(); conn.close()
                        st.success(f"✅ Updated {uname}"); st.rerun()

            elif op == "🗑️ Delete":
                if len(emp_df):
                    opts2 = {f"{r['Name']} — {r['Department']} ({r['EmployeeID']})": r['EmployeeID'] for _,r in emp_df.iterrows()}
                    sel2 = st.selectbox("Select Employee to Delete:",list(opts2.keys()),key="emp_del_sel")
                    eid2 = opts2[sel2]
                    st.warning(f"⚠️ This will permanently delete this employee and their participation records.")
                    if st.button("🗑️ CONFIRM DELETE",use_container_width=True):
                        conn=get_conn()
                        conn.execute("DELETE FROM Employees WHERE EmployeeID=?",(eid2,))
                        conn.execute("DELETE FROM Participation WHERE EmployeeID=?",(eid2,))
                        conn.execute("DELETE FROM Scores WHERE EmployeeID=?",(eid2,))
                        conn.execute("DELETE FROM BMI WHERE EmployeeID=?",(eid2,))
                        conn.commit(); conn.close()
                        st.success("✅ Deleted."); st.rerun()

        # ── TAB 2: EVENTS ──────────────────────────────────────────────────────
        with tabs[1]:
            st.markdown(f'<div class="sh" style="margin-top:14px;">⚽ Event Master — Add / Edit / Delete / View</div>', unsafe_allow_html=True)
            ev_df = q("SELECT EventID,EventName,Type,Difficulty,Multiplier,PrizeWinner,PrizeRunnerUp,Prize2ndRunnerUp FROM Events ORDER BY EventName")
            st.dataframe(ev_df, use_container_width=True, hide_index=True,
                column_config={
                    "EventID": st.column_config.NumberColumn("ID",format="%d"),
                    "Multiplier": st.column_config.NumberColumn("Multiplier",format="%.1f"),
                    "PrizeWinner": st.column_config.TextColumn("🥇 Prize (Winner)"),
                    "PrizeRunnerUp": st.column_config.TextColumn("🥈 Prize (Runner-up)"),
                    "Prize2ndRunnerUp": st.column_config.TextColumn("🥉 Prize (2nd R-up)"),
                })
            st.markdown("---")
            ev_op = st.radio("Action:",["➕ Add New","✏️ Edit","🗑️ Delete"],horizontal=True,key="ev_op")

            if ev_op == "➕ Add New":
                ec1,ec2,ec3,ec4 = st.columns(4)
                with ec1: ev_nm = st.text_input("Event Name",placeholder="e.g. Football",key="ev_nm")
                with ec2: ev_tp = st.selectbox("Type",["Indoor","Outdoor"],key="ev_tp")
                with ec3: ev_df2 = st.selectbox("Difficulty",["Casual","Medium","High"],key="ev_df2")
                with ec4: ev_ml = st.number_input("Multiplier",min_value=0.5,max_value=5.0,value=1.0,step=0.1,key="ev_ml")
                ep1,ep2,ep3 = st.columns(3)
                with ep1: ev_pw  = st.text_input("🥇 Prize for Winner",placeholder="e.g. Trophy + ₹5000",key="ev_pw")
                with ep2: ev_pru = st.text_input("🥈 Prize for Runner-up",placeholder="e.g. Medal + ₹2000",key="ev_pru")
                with ep3: ev_p2r = st.text_input("🥉 Prize for 2nd Runner-up",placeholder="e.g. Certificate",key="ev_p2r")
                if st.button("➕ ADD EVENT",use_container_width=True):
                    if ev_nm:
                        try:
                            conn=get_conn()
                            conn.execute("INSERT OR IGNORE INTO Events (EventName,Type,Difficulty,Multiplier,Description,PrizeWinner,PrizeRunnerUp,Prize2ndRunnerUp) VALUES (?,?,?,?,?,?,?,?)",
                                         (ev_nm.strip(),ev_tp,ev_df2,ev_ml,'',ev_pw,ev_pru,ev_p2r))
                            conn.commit(); conn.close()
                            st.success(f"✅ Added event: {ev_nm}"); st.rerun()
                        except Exception as e: st.error(str(e))

            elif ev_op == "✏️ Edit":
                if len(ev_df):
                    ev_opts = {r['EventName']:r['EventID'] for _,r in ev_df.iterrows()}
                    sel_ev = st.selectbox("Select Event:",list(ev_opts.keys()),key="ev_edit_sel")
                    cur_ev = ev_df[ev_df['EventName']==sel_ev].iloc[0]
                    eec1,eec2,eec3 = st.columns(3)
                    with eec1: utp = st.selectbox("Type",["Indoor","Outdoor"],index=["Indoor","Outdoor"].index(cur_ev['Type']) if cur_ev['Type'] in ["Indoor","Outdoor"] else 0,key="ev_utp")
                    with eec2: udf = st.selectbox("Difficulty",["Casual","Medium","High"],index=["Casual","Medium","High"].index(cur_ev['Difficulty']) if cur_ev['Difficulty'] in ["Casual","Medium","High"] else 0,key="ev_udf")
                    with eec3: uml = st.number_input("Multiplier",value=float(cur_ev['Multiplier']),min_value=0.5,max_value=5.0,step=0.1,key="ev_uml")
                    eep1,eep2,eep3 = st.columns(3)
                    with eep1: upw  = st.text_input("🥇 Prize Winner",  value=str(cur_ev.get('PrizeWinner','') or ''),  key="ev_upw")
                    with eep2: upru = st.text_input("🥈 Prize Runner-up",value=str(cur_ev.get('PrizeRunnerUp','') or ''),key="ev_upru")
                    with eep3: up2r = st.text_input("🥉 Prize 2nd R-up", value=str(cur_ev.get('Prize2ndRunnerUp','') or ''),key="ev_up2r")
                    if st.button("💾 SAVE EVENT",use_container_width=True):
                        conn=get_conn()
                        conn.execute("UPDATE Events SET Type=?,Difficulty=?,Multiplier=?,PrizeWinner=?,PrizeRunnerUp=?,Prize2ndRunnerUp=? WHERE EventName=?",
                                     (utp,udf,uml,upw,upru,up2r,sel_ev))
                        conn.commit(); conn.close(); st.success("✅ Updated!"); st.rerun()

            elif ev_op == "🗑️ Delete":
                if len(ev_df):
                    ev_opts2 = {r['EventName']:r['EventID'] for _,r in ev_df.iterrows()}
                    sel_ev2 = st.selectbox("Select Event to Delete:",list(ev_opts2.keys()),key="ev_del_sel")
                    st.warning("⚠️ This will delete the event. Participation records referencing it will remain.")
                    if st.button("🗑️ CONFIRM DELETE EVENT",use_container_width=True):
                        conn=get_conn(); conn.execute("DELETE FROM Events WHERE EventName=?",(sel_ev2,))
                        conn.commit(); conn.close(); st.success("✅ Deleted."); st.rerun()

        # ── TAB 3: PARTICIPATION ───────────────────────────────────────────────
        with tabs[2]:
            st.markdown(f'<div class="sh" style="margin-top:14px;">🎮 Participation Entry — Add / Edit / Delete / View</div>', unsafe_allow_html=True)
            part_df = q("SELECT PID,EmployeeName,EventName,Date,Position,FinalPoints FROM Participation ORDER BY Date DESC")

            # Filter + Search
            pf1, pf2 = st.columns([1.5, 2])
            with pf1:
                ev_names_p = ["All"] + sorted(part_df['EventName'].unique().tolist()) if len(part_df) else ["All"]
                pf = st.selectbox("Filter by Event:",ev_names_p,key="part_ef")
            with pf2:
                part_search = st.text_input("🔍 Search by Name:", placeholder="Type employee name...", key="part_search")
            show_part = part_df if pf=="All" else part_df[part_df['EventName']==pf]
            if part_search.strip():
                show_part = show_part[show_part['EmployeeName'].str.contains(part_search.strip(), case=False, na=False)]
            st.dataframe(show_part, use_container_width=True, hide_index=True,
                column_config={"FinalPoints": st.column_config.NumberColumn("Final Points", format="%d")})
            st.markdown(f'<div class="ins" style="font-size:0.78rem;">{len(show_part)} records shown</div>', unsafe_allow_html=True)

            st.markdown("---")
            pop = st.radio("Action:",["➕ Add Entry","✏️ Edit Entry","🗑️ Delete Entry"],horizontal=True,key="part_op")

            emp_all = q("SELECT EmployeeID,Name,Department FROM Employees ORDER BY Name")
            ev_all  = q("SELECT EventName,Multiplier FROM Events")
            emp_opts = {f"{r['Name']} — {r['Department']}": (r['EmployeeID'],r['Name'],r['Department']) for _,r in emp_all.iterrows()}

            if pop == "➕ Add Entry":
                pc1,pc2,pc3 = st.columns(3)
                with pc1: pe_sel = st.selectbox("Employee",list(emp_opts.keys()),key="part_emp")
                with pc2: pe_ev  = st.selectbox("Event",ev_all['EventName'].tolist(),key="part_ev")
                with pc3: pe_pos = st.selectbox("Position",["Winner","Runner-up","2nd Runner-up","Semi","Quarter","Participant"],key="part_pos")
                pc4,pc5,pc6 = st.columns(3)
                base_map = {"Winner":100,"Runner-up":70,"2nd Runner-up":50,"Semi":30,"Quarter":20,"Participant":0}
                ev_mult_row = ev_all[ev_all['EventName']==pe_ev]
                ev_mult = float(ev_mult_row.iloc[0]['Multiplier']) if len(ev_mult_row) else 1.0
                with pc4: pe_part_pts = st.number_input("Participation Pts",value=10,min_value=0,key="part_pp")
                with pc5: pe_base = st.number_input("Base Points",value=base_map.get(pe_pos,0),min_value=0,key="part_bp")
                with pc6: pe_date = st.date_input("Date",value=date.today(),key="part_dt")
                gpts = pe_base * ev_mult; fpts = gpts + pe_part_pts
                st.markdown(f'<div class="ins">Auto-calculated: Game Pts = {pe_base} × {ev_mult} = <b>{gpts}</b> &nbsp;|&nbsp; Final = <b>{fpts}</b></div>', unsafe_allow_html=True)

                if st.button("✅ ADD PARTICIPATION",use_container_width=True):
                    eid_p, ename_p, edept_p = emp_opts[pe_sel]
                    conn=get_conn()
                    try:
                        conn.execute("""INSERT INTO Participation
                            (EmployeeID,EmployeeName,EventName,Date,Position,Registered,Participated,
                             ParticipationPoints,BasePoints,Multiplier,GamePoints,FinalPoints)
                            VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                            (eid_p,ename_p,pe_ev,str(pe_date),pe_pos,'Yes','Yes',
                             pe_part_pts,pe_base,ev_mult,gpts,fpts))
                        recalc_score(eid_p, conn)
                        conn.commit()
                        st.success(f"✅ Added! Total score for {ename_p} recalculated."); st.rerun()
                    except Exception as e: st.error(str(e))
                    finally: conn.close()

            elif pop == "✏️ Edit Entry":
                if len(part_df):
                    pid_opts = {f"#{r['PID']} — {r['EmployeeName']} · {r['EventName']} · {r['Position']}": r['PID'] for _,r in part_df.iterrows()}
                    sel_pid = st.selectbox("Select Record:",list(pid_opts.keys()),key="part_edit_sel")
                    pid = pid_opts[sel_pid]
                    cur_p = q(f"SELECT * FROM Participation WHERE PID={pid}").iloc[0]
                    epc1,epc2,epc3 = st.columns(3)
                    with epc1: up_pos = st.selectbox("Position",["Winner","Runner-up","2nd Runner-up","Semi","Quarter","Participant"],
                                                      index=["Winner","Runner-up","2nd Runner-up","Semi","Quarter","Participant"].index(cur_p['Position']) if cur_p['Position'] in ["Winner","Runner-up","2nd Runner-up","Semi","Quarter","Participant"] else 5,key="part_upos")
                    with epc2: up_base = st.number_input("Base Points",value=int(cur_p['BasePoints']),min_value=0,key="part_ubp")
                    with epc3: up_ppts = st.number_input("Participation Pts",value=int(cur_p['ParticipationPoints']),min_value=0,key="part_upp")
                    up_mult = float(cur_p['Multiplier'])
                    up_gpts = up_base * up_mult; up_fpts = up_gpts + up_ppts
                    st.markdown(f'<div class="ins">Recalculated Final = <b>{up_fpts} pts</b></div>', unsafe_allow_html=True)
                    if st.button("💾 SAVE ENTRY",use_container_width=True):
                        conn=get_conn()
                        conn.execute("UPDATE Participation SET Position=?,BasePoints=?,ParticipationPoints=?,GamePoints=?,FinalPoints=? WHERE PID=?",
                                     (up_pos,up_base,up_ppts,up_gpts,up_fpts,pid))
                        conn.commit()
                        recalc_score(int(cur_p['EmployeeID']),conn)
                        conn.commit(); conn.close()
                        st.success(f"✅ Entry updated! Score auto-synced to Scores tab."); st.rerun()

            elif pop == "🗑️ Delete Entry":
                if len(part_df):
                    pid_opts2 = {f"#{r['PID']} — {r['EmployeeName']} · {r['EventName']}": r['PID'] for _,r in part_df.iterrows()}
                    sel_pid2 = st.selectbox("Select to Delete:",list(pid_opts2.keys()),key="part_del_sel")
                    pid2 = pid_opts2[sel_pid2]
                    emp_id_del = part_df[part_df['PID']==pid2].iloc[0]
                    if st.button("🗑️ DELETE ENTRY",use_container_width=True):
                        # Get employee ID before deleting
                        emp_id_to_recalc = part_df[part_df['PID']==pid2]['EmployeeID'].values
                        conn=get_conn()
                        conn.execute("DELETE FROM Participation WHERE PID=?",(pid2,))
                        conn.commit()
                        # Recalculate score after deletion
                        if len(emp_id_to_recalc):
                            recalc_score(int(emp_id_to_recalc[0]), conn)
                            conn.commit()
                        conn.close()
                        st.success("✅ Deleted and score updated."); st.rerun()

        # ── TAB 4: SCHEDULE ────────────────────────────────────────────────────
        with tabs[3]:
            st.markdown(f'<div class="sh" style="margin-top:14px;">📅 Schedule — Add / Edit Status / Delete</div>', unsafe_allow_html=True)
            sched_df = q("SELECT * FROM Schedule ORDER BY StartTime ASC")
            if len(sched_df): st.dataframe(sched_df, use_container_width=True, hide_index=True)
            st.markdown("---")
            sop = st.radio("Action:",["➕ Add Match","✏️ Update Status","🗑️ Delete"],horizontal=True,key="sched_op")
            ev_names_s = q("SELECT EventName FROM Events")['EventName'].tolist()

            if sop == "➕ Add Match":
                sa1,sa2 = st.columns(2)
                with sa1:
                    s_ev = st.selectbox("Event",ev_names_s+["Other..."],key="sched_ev")
                    if s_ev=="Other...": s_ev = st.text_input("Custom event name",key="sched_ev_cust")
                    s_stat = st.selectbox("Status",["Upcoming","Live","Completed","Postponed"],key="sched_stat")
                with sa2:
                    sd,st2 = st.columns(2)
                    with sd: s_date = st.date_input("Date",value=date.today(),key="sched_date")
                    with st2: s_time = st.time_input("Time",key="sched_time")
                    s_venue = st.text_input("Venue",placeholder="e.g. Main Hall",key="sched_venue")
                s_notes = st.text_input("Notes (optional)",key="sched_notes")
                if st.button("📅 ADD TO SCHEDULE",use_container_width=True):
                    sdt = datetime.combine(s_date,s_time).isoformat()
                    conn=get_conn()
                    conn.execute("INSERT INTO Schedule (EventName,StartTime,Status,Venue,Notes) VALUES (?,?,?,?,?)",(s_ev,sdt,s_stat,s_venue,s_notes))
                    conn.commit(); conn.close()
                    st.success("✅ Added to schedule!"); st.rerun()

            elif sop == "✏️ Update Status":
                if len(sched_df):
                    s_opts = {f"{r['EventName']} — {str(r['StartTime'])[:16]}": r['SID'] for _,r in sched_df.iterrows()}
                    s_sel = st.selectbox("Select Match:",list(s_opts.keys()),key="sched_edit_sel")
                    s_sid = s_opts[s_sel]
                    cur_stat = sched_df[sched_df['SID']==s_sid].iloc[0]['Status']
                    new_stat = st.selectbox("New Status",["Upcoming","Live","Completed","Postponed"],
                                            index=["Upcoming","Live","Completed","Postponed"].index(str(cur_stat)) if str(cur_stat) in ["Upcoming","Live","Completed","Postponed"] else 0,key="sched_new_stat")
                    if st.button("💾 UPDATE STATUS",use_container_width=True):
                        conn=get_conn(); conn.execute("UPDATE Schedule SET Status=? WHERE SID=?",(new_stat,s_sid))
                        conn.commit(); conn.close(); st.success("✅ Status updated!"); st.rerun()

            elif sop == "🗑️ Delete":
                if len(sched_df):
                    s_opts3 = {f"{r['EventName']} — {str(r['StartTime'])[:16]}": r['SID'] for _,r in sched_df.iterrows()}
                    s_sel3 = st.selectbox("Select to Delete:",list(s_opts3.keys()),key="sched_del_sel")
                    s_sid3 = s_opts3[s_sel3]
                    if st.button("🗑️ DELETE MATCH",use_container_width=True):
                        conn=get_conn(); conn.execute("DELETE FROM Schedule WHERE SID=?",(s_sid3,))
                        conn.commit(); conn.close(); st.success("✅ Deleted."); st.rerun()

        # ── TAB 5: SCORES ──────────────────────────────────────────────────────
        with tabs[4]:
            st.markdown(f'<div class="sh" style="margin-top:14px;">📊 Scores & Leaderboard</div>', unsafe_allow_html=True)
            sc_df = get_leaderboard()
            sc_search = st.text_input("🔍 Search by Name or Department:", placeholder="Type to filter...", key="sc_search")
            show_sc = sc_df
            if sc_search.strip():
                show_sc = sc_df[sc_df['Name'].str.contains(sc_search.strip(), case=False, na=False) |
                                sc_df['Department'].str.contains(sc_search.strip(), case=False, na=False)]
            sc_df_display = show_sc
            st.dataframe(sc_df_display, use_container_width=True, hide_index=True,
                column_config={
                    "TotalPoints": st.column_config.NumberColumn("Points 🏆", format="%d"),
                    "EmployeeID":  st.column_config.NumberColumn("Employee ID", format="%d"),
                })
            st.markdown(f'<div class="ins" style="font-size:0.78rem;">Showing <b>{len(sc_df_display)}</b> of <b>{len(sc_df)}</b> employees</div>', unsafe_allow_html=True)

            st.markdown("---")
            st.markdown('<div class="sh">✏️ Override Score Manually</div>', unsafe_allow_html=True)
            if len(sc_df):
                sc_opts = {f"{r['Name']} — {r['Department']}": r['EmployeeID'] for _,r in sc_df.iterrows()}
                sc_sel = st.selectbox("Employee:",list(sc_opts.keys()),key="sc_sel")
                sc_eid = sc_opts[sc_sel]
                cur_sc = sc_df[sc_df['EmployeeID']==sc_eid].iloc[0]['TotalPoints']
                sc_val = st.number_input("New Score",value=int(cur_sc),min_value=0,max_value=99999,key="sc_val")
                st.markdown(f'<div class="ins" style="font-size:0.8rem;">⚠️ Manual override bypasses participation calculation. Use only for corrections.</div>', unsafe_allow_html=True)
                if st.button("💾 OVERRIDE SCORE",use_container_width=True):
                    conn=get_conn()
                    conn.execute("UPDATE Scores SET Score=?,LastUpdated=? WHERE EmployeeID=?",(sc_val,date.today().isoformat(),sc_eid))
                    conn.commit(); conn.close()
                    st.success(f"✅ Score updated to {sc_val}"); st.rerun()

            st.markdown("---")
            if st.button("📥 EXPORT LEADERBOARD CSV",use_container_width=True):
                st.download_button("⬇️ Download",sc_df.to_csv(index=False),"leaderboard.csv","text/csv",use_container_width=True)

            st.markdown("---")
            st.markdown('<div class="sh">🔄 Master Data Sync from Excel</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="ins" style="font-size:0.8rem;">Sync data from <b>Employee_Wellness_Scoring_System.xlsx</b> into the database.<br>Place the Excel file in the <b>same folder as app.py</b> then click sync.<br>✅ <b>Smart Sync:</b> Only NEW records are inserted — existing data is never overwritten or deleted. Scores are auto-recalculated for employees with new participation entries.</div>', unsafe_allow_html=True)
            sync_path = st.text_input("Excel file path (optional)", placeholder="Leave blank if Excel is in same folder as app.py", key="sync_path_input")
            if st.button("🔄 SYNC FROM EXCEL NOW", use_container_width=True):
                import pathlib
                final_path = sync_path.strip() if sync_path.strip() else "Employee_Wellness_Scoring_System.xlsx"
                # Also try same directory as app.py
                if not os.path.exists(final_path):
                    app_dir = str(pathlib.Path(__file__).parent / "Employee_Wellness_Scoring_System.xlsx")
                    if os.path.exists(app_dir):
                        final_path = app_dir
                if not os.path.exists(final_path):
                    st.error(f"❌ Excel file not found at: {os.path.abspath(final_path)}\nPlease place Employee_Wellness_Scoring_System.xlsx in the same folder as app.py")
                else:
                    result = seed_initial_data(final_path)
                    ok, payload = result if result else (False, "Unknown error")
                    if ok and isinstance(payload, dict):
                        rpt = payload
                        total_new = sum(v.get('new', 0) for v in rpt.values())
                        if total_new > 0:
                            st.success(f"✅ Sync complete! {total_new} new record(s) added across all sheets.")
                        else:
                            st.info("✅ Sync complete — database is already up to date. No new records found.")
                        # Detailed breakdown
                        rows = []
                        sheet_icons = {
                            'Employees':     '👤',
                            'Events':        '🏆',
                            'Participation': '📋',
                            'Leaderboard':   '📊',
                            'Schedule':      '📅',
                        }
                        for sheet, counts in rpt.items():
                            icon = sheet_icons.get(sheet, '📄')
                            new_c  = counts.get('new', 0)
                            upd_c  = counts.get('updated', 0)
                            skip_c = counts.get('skipped', 0)
                            detail = f"🆕 {new_c} new"
                            if upd_c:  detail += f"  |  ✏️ {upd_c} updated"
                            if skip_c: detail += f"  |  ⏭️ {skip_c} already existed"
                            rows.append({"Sheet": f"{icon} {sheet}", "Result": detail})
                        if rows:
                            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
                    elif ok:
                        st.success("✅ Data synced successfully from Excel!")
                    else:
                        st.error(f"❌ Sync failed: {payload}")
                    st.rerun()

        # ── TAB 6: BMI UPLOAD ──────────────────────────────────────────────────
        with tabs[5]:
            st.markdown(f'<div class="sh" style="margin-top:14px;">🏥 BMI Data Upload — Excel Template</div>', unsafe_allow_html=True)

            # Template download
            st.markdown(f'<div class="ins ins-green">📥 <b>Step 1:</b> Download the blank template → fill Weight & Height → upload below.<br>BMI will be auto-calculated. You can upload monthly — duplicates will be updated.</div>', unsafe_allow_html=True)

            emp_list = q("SELECT EmployeeID,Name,Department FROM Employees ORDER BY Name")
            if len(emp_list):
                template_df = pd.DataFrame({
                    'Employee_ID':   emp_list['EmployeeID'].tolist(),
                    'Employee_Name': emp_list['Name'].tolist(),
                    'Department':    emp_list['Department'].tolist(),
                    'Month':         ['March'] * len(emp_list),
                    'Year':          [2026] * len(emp_list),
                    'Weight_kg':     [''] * len(emp_list),
                    'Height_cm':     [''] * len(emp_list),
                })
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                    template_df.to_excel(writer, index=False, sheet_name='BMI_Data')
                st.download_button(
                    "📥 DOWNLOAD BMI TEMPLATE",
                    buf.getvalue(),
                    "BMI_Template.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown(f'<div class="ins ins-green">📤 <b>Step 2:</b> Upload the filled Excel file below.</div>', unsafe_allow_html=True)

            uploaded = st.file_uploader("Upload filled BMI Excel", type=['xlsx','csv'], key="bmi_upload")

            if uploaded:
                try:
                    if uploaded.name.endswith('.csv'):
                        df_up = pd.read_csv(uploaded)
                    else:
                        df_up = pd.read_excel(uploaded)

                    # Validate columns
                    required = {'Employee_ID','Employee_Name','Month','Year','Weight_kg','Height_cm'}
                    missing = required - set(df_up.columns)
                    if missing:
                        st.error(f"❌ Missing columns: {missing}. Please use the template.")
                    else:
                        df_up = df_up.dropna(subset=['Employee_ID','Weight_kg','Height_cm'])
                        df_up['Weight_kg'] = pd.to_numeric(df_up['Weight_kg'], errors='coerce')
                        df_up['Height_cm'] = pd.to_numeric(df_up['Height_cm'], errors='coerce')
                        df_up = df_up.dropna(subset=['Weight_kg','Height_cm'])
                        df_up['BMI_calc'] = (df_up['Weight_kg'] / (df_up['Height_cm']/100)**2).round(2)
                        df_up['Category'] = df_up['BMI_calc'].apply(bmi_category)

                        # Preview
                        st.markdown(f'<div class="sh sh-green">Preview — {len(df_up)} records to upload</div>', unsafe_allow_html=True)
                        preview = df_up[['Employee_ID','Employee_Name','Month','Year','Weight_kg','Height_cm','BMI_calc','Category']].copy()
                        preview.columns = ['ID','Name','Month','Year','Weight(kg)','Height(cm)','BMI','Category']
                        st.dataframe(preview, use_container_width=True, hide_index=True)

                        if st.button("✅ CONFIRM UPLOAD TO DATABASE", use_container_width=True):
                            conn = get_conn()
                            ok = upd = err = 0
                            today = date.today().isoformat()
                            for _, row in df_up.iterrows():
                                try:
                                    eid = int(row['Employee_ID'])
                                    nm  = str(row['Employee_Name'])
                                    mon = str(row['Month'])
                                    yr  = int(row['Year'])
                                    wt  = float(row['Weight_kg'])
                                    ht  = float(row['Height_cm'])
                                    bv  = float(row['BMI_calc'])
                                    cat = str(row['Category'])
                                    # Get department
                                    dep_row = conn.execute("SELECT Department FROM Employees WHERE EmployeeID=?", (eid,)).fetchone()
                                    dep = dep_row[0] if dep_row else "Unknown"
                                    # Check if record exists for same employee+month+year
                                    exists = conn.execute("SELECT BID FROM BMI WHERE EmployeeID=? AND Month=? AND Year=?",
                                                          (eid,mon,yr)).fetchone()
                                    if exists:
                                        conn.execute("UPDATE BMI SET Weight_kg=?,Height_cm=?,BMI=?,Category=?,RecordedOn=? WHERE BID=?",
                                                     (wt,ht,bv,cat,today,exists[0]))
                                        upd += 1
                                    else:
                                        conn.execute("INSERT INTO BMI (EmployeeID,EmployeeName,Department,Month,Year,Weight_kg,Height_cm,BMI,Category,RecordedOn) VALUES (?,?,?,?,?,?,?,?,?,?)",
                                                     (eid,nm,dep,mon,yr,wt,ht,bv,cat,today))
                                        ok += 1
                                except Exception as ex:
                                    err += 1
                            conn.commit(); conn.close()
                            st.success(f"✅ Upload complete! {ok} new records added, {upd} updated, {err} skipped.")
                            if err > 0: st.warning(f"⚠️ {err} rows skipped — check Employee IDs exist in system.")
                            st.rerun()
                except Exception as e:
                    st.error(f"❌ Error reading file: {e}")

            # Current BMI data
            bmi_cur = q("SELECT EmployeeName,Department,Month,Year,Weight_kg,Height_cm,ROUND(BMI,2) AS BMI,Category FROM BMI ORDER BY Year DESC,EmployeeName")
            if len(bmi_cur):
                st.markdown("---")
                st.markdown(f'<div class="sh sh-green">Current BMI Records ({len(bmi_cur)} entries)</div>', unsafe_allow_html=True)
                st.dataframe(bmi_cur, use_container_width=True, hide_index=True)
                # Delete option
                st.markdown("---")
                if st.checkbox("🗑️ Delete a BMI record", key="bmi_del_cb"):
                    bmi_all = q("SELECT BID,EmployeeName,Month,Year FROM BMI ORDER BY EmployeeName,Year,Month")
                    bmi_del_opts = {f"#{r['BID']} — {r['EmployeeName']} · {r['Month']} {int(r['Year'])}": r['BID'] for _,r in bmi_all.iterrows()}
                    bmi_del_sel = st.selectbox("Select record to delete:",list(bmi_del_opts.keys()),key="bmi_del_sel")
                    bmi_del_id  = bmi_del_opts[bmi_del_sel]
                    if st.button("🗑️ DELETE BMI RECORD",use_container_width=True):
                        conn=get_conn(); conn.execute("DELETE FROM BMI WHERE BID=?",(bmi_del_id,))
                        conn.commit(); conn.close(); st.success("✅ Deleted."); st.rerun()
