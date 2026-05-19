import streamlit as st

st.set_page_config(
    page_title="Fleet Running Hours",
    page_icon="⚓",
    layout="wide",
    initial_sidebar_state="expanded",
)

import os, re, sqlite3, tempfile, hashlib, subprocess, shutil
from datetime import datetime
from pathlib import Path
import pandas as pd

# ════════════════════════════════════════════════════════════════════
# CSS — Deep Space Command Center
# Fonts: Space Grotesk (display) + Inter (body) + JetBrains Mono
# Palette: midnight blue · electric gold · signal red · phosphor green
# ════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600;700&family=Inter:wght@300;400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

/* ─── TOKENS ─────────────────────────────────────────────────────── */
:root {
  /* backgrounds */
  --bg0:  #03060d;
  --bg1:  #060c18;
  --bg2:  #080f1f;
  --bg3:  #0b1428;
  --bg4:  #0f1c35;
  --bg5:  #132240;
  /* borders */
  --b1:   #111e35;
  --b2:   #1a2e50;
  --b3:   #243b60;
  /* gold accent */
  --g0:   #7a5a10;
  --g1:   #b8820e;
  --g2:   #d4a017;
  --g3:   #e8b830;
  --g4:   #f5cc55;
  --g5:   #fde27a;
  /* status colors */
  --red0: #4a0f0f;
  --red1: #7f1d1d;
  --red2: #dc2626;
  --red3: #f87171;
  --red4: #fecaca;
  --ora0: #4a1f06;
  --ora1: #7c2d12;
  --ora2: #c2410c;
  --ora3: #fb923c;
  --ora4: #fed7aa;
  --grn0: #052a18;
  --grn1: #064e2d;
  --grn2: #059669;
  --grn3: #34d399;
  --grn4: #a7f3d0;
  --blu0: #0a1a3a;
  --blu1: #1e3a5f;
  --blu2: #1d4ed8;
  --blu3: #60a5fa;
  --blu4: #bfdbfe;
  /* text */
  --t0:   #f0f6ff;
  --t1:   #c8d8f0;
  --t2:   #7a92b5;
  --t3:   #3d5578;
  --t4:   #1e3050;
  /* typography */
  --ff:   'Space Grotesk', sans-serif;
  --fi:   'Inter', sans-serif;
  --fm:   'JetBrains Mono', monospace;
}

/* ─── RESET & BASE ───────────────────────────────────────────────── */
*, *::before, *::after { box-sizing: border-box; }
html, body, [class*="css"] {
  font-family: var(--fi) !important;
  background: var(--bg0) !important;
  color: var(--t1) !important;
  -webkit-font-smoothing: antialiased;
  text-rendering: optimizeLegibility;
}
.main, .main > div { background: var(--bg0) !important; }
.block-container {
  padding: 2.25rem 3rem 6rem !important;
  max-width: 100% !important;
}

/* ─── ATMOSPHERIC BG ─────────────────────────────────────────────── */
.main::before {
  content: '';
  position: fixed; inset: 0;
  pointer-events: none; z-index: 0;
  background:
    radial-gradient(ellipse 100% 60% at -5% -10%,
      rgba(180,130,14,0.07) 0%, transparent 50%),
    radial-gradient(ellipse 80% 50% at 105% 110%,
      rgba(29,78,216,0.06) 0%, transparent 50%),
    radial-gradient(ellipse 50% 80% at 50% 50%,
      rgba(3,6,13,0.95) 0%, transparent 100%);
}

/* ─── SIDEBAR ────────────────────────────────────────────────────── */
[data-testid="stSidebar"] {
  background: var(--bg1) !important;
  border-right: 1px solid var(--b2) !important;
  box-shadow: 4px 0 24px rgba(0,0,0,0.4) !important;
}
[data-testid="stSidebar"] * { color: var(--t1) !important; }
[data-testid="stSidebarContent"] { padding: 1.75rem 1.25rem !important; }
[data-testid="stSidebar"] .stSelectbox > div > div {
  background: var(--bg3) !important;
  border: 1px solid var(--b2) !important;
  border-radius: 6px !important;
  font-family: var(--fi) !important;
}

/* ─── TYPOGRAPHY ─────────────────────────────────────────────────── */
h1 {
  font-family: var(--ff) !important;
  font-size: 1.85rem !important; font-weight: 700 !important;
  color: var(--t0) !important; letter-spacing: -0.02em !important;
  line-height: 1.2 !important;
}
h2 {
  font-family: var(--ff) !important;
  font-size: 1.25rem !important; font-weight: 600 !important;
  color: var(--t0) !important;
}
h3 {
  font-family: var(--ff) !important;
  font-size: 1rem !important; font-weight: 500 !important;
  color: var(--t1) !important;
}
p { font-family: var(--fi) !important; }

/* ─── METRICS ────────────────────────────────────────────────────── */
[data-testid="stMetric"] {
  background: var(--bg3) !important;
  border: 1px solid var(--b2) !important;
  border-radius: 12px !important;
  padding: 1.1rem 1.3rem 1.2rem !important;
  position: relative !important;
  overflow: hidden !important;
  transition: border-color 0.25s ease, transform 0.2s ease, box-shadow 0.25s ease !important;
}
[data-testid="stMetric"]:hover {
  border-color: var(--b3) !important;
  transform: translateY(-3px) !important;
  box-shadow: 0 12px 40px rgba(0,0,0,0.5) !important;
}
[data-testid="stMetricValue"] {
  font-family: var(--ff) !important;
  font-size: 2.1rem !important; font-weight: 700 !important;
  color: var(--t0) !important; letter-spacing: -0.03em !important;
}
[data-testid="stMetricLabel"] {
  font-family: var(--fi) !important;
  color: var(--t3) !important; font-size: 0.65rem !important;
  text-transform: uppercase !important; letter-spacing: 0.15em !important;
  font-weight: 500 !important;
}

/* ─── DATAFRAME ──────────────────────────────────────────────────── */
[data-testid="stDataFrame"] {
  border: 1px solid var(--b2) !important;
  border-radius: 10px !important;
  overflow: hidden !important;
  box-shadow: 0 4px 20px rgba(0,0,0,0.3) !important;
}
.dvn-scroller { background: var(--bg2) !important; }

/* ─── BUTTONS ────────────────────────────────────────────────────── */
.stButton > button {
  background: linear-gradient(135deg, var(--g2) 0%, var(--g0) 100%) !important;
  color: #000 !important; border: none !important;
  font-family: var(--ff) !important; font-weight: 600 !important;
  font-size: 0.83rem !important; letter-spacing: 0.06em !important;
  text-transform: uppercase !important; border-radius: 7px !important;
  padding: 0.65rem 2rem !important;
  box-shadow: 0 2px 12px rgba(180,130,14,0.25),
              inset 0 1px 0 rgba(255,255,255,0.12) !important;
  transition: all 0.18s ease !important;
}
.stButton > button:hover {
  background: linear-gradient(135deg, var(--g3) 0%, var(--g1) 100%) !important;
  box-shadow: 0 6px 24px rgba(180,130,14,0.4) !important;
  transform: translateY(-2px) !important;
}
.stButton > button:active { transform: translateY(0) !important; }

/* ─── FILE UPLOADER ──────────────────────────────────────────────── */
[data-testid="stFileUploadDropzone"] {
  background: linear-gradient(160deg,
    rgba(180,130,14,0.05) 0%,
    rgba(29,78,216,0.04) 100%) !important;
  border: 1.5px dashed var(--g2) !important;
  border-radius: 14px !important;
  transition: all 0.3s ease !important;
  padding: 3.5rem 2rem !important;
}
[data-testid="stFileUploadDropzone"]:hover {
  background: rgba(180,130,14,0.08) !important;
  border-color: var(--g3) !important;
  box-shadow: 0 0 50px rgba(180,130,14,0.09),
              inset 0 0 50px rgba(180,130,14,0.03) !important;
}
[data-testid="stFileUploadDropzone"] p,
[data-testid="stFileUploadDropzone"] span {
  color: var(--g3) !important;
  font-family: var(--ff) !important;
  font-size: 1rem !important; font-weight: 500 !important;
}
[data-testid="stFileUploadDropzone"] small { color: var(--t3) !important; }

/* ─── TABS ───────────────────────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"] {
  background: var(--bg2) !important;
  border-radius: 10px 10px 0 0 !important;
  border-bottom: 1px solid var(--b2) !important;
  gap: 0 !important; padding: 0 1rem !important;
}
.stTabs [data-baseweb="tab"] {
  background: transparent !important;
  color: var(--t4) !important;
  font-family: var(--ff) !important; font-weight: 500 !important;
  letter-spacing: 0.04em !important; font-size: 0.78rem !important;
  text-transform: uppercase !important;
  padding: 0.9rem 1.4rem !important;
  border-bottom: 2px solid transparent !important;
  margin-bottom: -1px !important; transition: color 0.2s !important;
}
.stTabs [data-baseweb="tab"]:hover { color: var(--t2) !important; }
.stTabs [aria-selected="true"] {
  color: var(--g3) !important;
  border-bottom: 2px solid var(--g2) !important;
}
.stTabs [data-baseweb="tab-panel"] {
  background: var(--bg2) !important;
  border: 1px solid var(--b2) !important;
  border-top: none !important;
  border-radius: 0 0 10px 10px !important;
  padding: 1.75rem !important;
}

/* ─── EXPANDER ───────────────────────────────────────────────────── */
.streamlit-expanderHeader {
  background: var(--bg3) !important;
  border: 1px solid var(--b2) !important;
  border-radius: 8px !important;
  font-family: var(--ff) !important; font-weight: 500 !important;
  font-size: 0.88rem !important; color: var(--t1) !important;
  transition: background 0.2s, border-color 0.2s !important;
}
.streamlit-expanderHeader:hover {
  background: var(--bg4) !important;
  border-color: var(--b3) !important;
}
.streamlit-expanderContent {
  background: var(--bg2) !important;
  border: 1px solid var(--b2) !important;
  border-top: none !important; border-radius: 0 0 8px 8px !important;
}

/* ─── FORM INPUTS ────────────────────────────────────────────────── */
.stSelectbox > div > div, .stMultiSelect > div > div {
  background: var(--bg3) !important;
  border: 1px solid var(--b2) !important;
  border-radius: 7px !important; color: var(--t1) !important;
  font-family: var(--fi) !important;
}
.stSelectbox label, .stMultiSelect label {
  font-family: var(--fi) !important;
  color: var(--t3) !important;
  font-size: 0.72rem !important; text-transform: uppercase !important;
  letter-spacing: 0.1em !important;
}

/* ─── MISC ───────────────────────────────────────────────────────── */
.stAlert { border-radius: 8px !important; border-left-width: 3px !important; }
hr { border-color: var(--b2) !important; opacity: 1 !important; margin: 1.5rem 0 !important; }
a { color: var(--g3) !important; text-decoration: none !important; }
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: var(--bg1); }
::-webkit-scrollbar-thumb { background: var(--b3); border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: var(--t4); }

/* ─── KEYFRAMES ──────────────────────────────────────────────────── */
@keyframes fadeIn     { from{opacity:0}                             to{opacity:1} }
@keyframes slideDown  { from{opacity:0;transform:translateY(-20px)} to{opacity:1;transform:translateY(0)} }
@keyframes slideUp    { from{opacity:0;transform:translateY(16px)}  to{opacity:1;transform:translateY(0)} }
@keyframes slideRight { from{opacity:0;transform:translateX(-16px)} to{opacity:1;transform:translateX(0)} }
@keyframes popIn      { from{opacity:0;transform:scale(0.9)}        to{opacity:1;transform:scale(1)} }
@keyframes numIn      { from{opacity:0;transform:translateY(10px)}  to{opacity:1;transform:translateY(0)} }
@keyframes successPop { 0%{transform:scale(0.85);opacity:0} 55%{transform:scale(1.02)} 100%{transform:scale(1);opacity:1} }
@keyframes goldLine   { from{width:0;opacity:0} to{width:100%;opacity:1} }
@keyframes scanPulse  { 0%,100%{opacity:0} 50%{opacity:1} }
@keyframes borderGlow { 0%,100%{box-shadow:0 0 0 rgba(212,160,23,0)} 50%{box-shadow:0 0 20px rgba(212,160,23,0.15)} }

/* ─── CUSTOM COMPONENTS ──────────────────────────────────────────── */

/* Page Header */
.ph {
  animation: slideDown 0.5s cubic-bezier(0.22,1,0.36,1) both;
  padding-bottom: 0.25rem;
}
.ph-line {
  height: 1px;
  background: linear-gradient(90deg, var(--g2) 0%, var(--b2) 35%, transparent 100%);
  margin: 0.4rem 0 2rem;
  animation: goldLine 0.8s 0.15s ease both;
  animation-fill-mode: both;
}
.ph-eyebrow {
  font-family: var(--fi);
  font-size: 0.62rem; font-weight: 500;
  letter-spacing: 0.22em; text-transform: uppercase;
  color: var(--g2); margin-bottom: 0.3rem;
  animation: fadeIn 0.4s 0.1s ease both;
  animation-fill-mode: both;
}

/* KPI Card */
.kc {
  background: var(--bg3);
  border: 1px solid var(--b2);
  border-radius: 12px;
  padding: 1.1rem 1.3rem 1.25rem;
  position: relative; overflow: hidden;
  animation: slideUp 0.4s ease both;
  animation-fill-mode: both;
  transition: border-color 0.25s, transform 0.2s, box-shadow 0.25s;
  cursor: default;
}
.kc:hover {
  transform: translateY(-4px);
  box-shadow: 0 16px 48px rgba(0,0,0,0.5);
}
/* Top accent bar */
.kc::before {
  content: '';
  position: absolute; top: 0; left: 0; right: 0; height: 3px;
  border-radius: 12px 12px 0 0;
}
/* Subtle inner gradient */
.kc::after {
  content: '';
  position: absolute; inset: 0;
  border-radius: 12px;
  pointer-events: none;
}
.kc.gold  { border-color: rgba(180,130,14,0.3); }
.kc.gold::before  { background: linear-gradient(90deg, var(--g2), transparent 70%); }
.kc.gold::after   { background: linear-gradient(160deg, rgba(180,130,14,0.06) 0%, transparent 60%); }
.kc.gold:hover    { border-color: rgba(212,160,23,0.5); box-shadow: 0 16px 48px rgba(0,0,0,0.5), 0 0 30px rgba(180,130,14,0.08); }

.kc.red   { border-color: rgba(220,38,38,0.25); }
.kc.red::before   { background: linear-gradient(90deg, var(--red2), transparent 70%); }
.kc.red::after    { background: linear-gradient(160deg, rgba(220,38,38,0.05) 0%, transparent 60%); }
.kc.red:hover     { border-color: rgba(220,38,38,0.4); box-shadow: 0 16px 48px rgba(0,0,0,0.5), 0 0 30px rgba(220,38,38,0.07); }

.kc.orange{ border-color: rgba(194,65,12,0.25); }
.kc.orange::before{ background: linear-gradient(90deg, var(--ora2), transparent 70%); }
.kc.orange::after { background: linear-gradient(160deg, rgba(194,65,12,0.05) 0%, transparent 60%); }
.kc.orange:hover  { border-color: rgba(194,65,12,0.4); box-shadow: 0 16px 48px rgba(0,0,0,0.5), 0 0 30px rgba(194,65,12,0.07); }

.kc.green { border-color: rgba(5,150,105,0.25); }
.kc.green::before { background: linear-gradient(90deg, var(--grn2), transparent 70%); }
.kc.green::after  { background: linear-gradient(160deg, rgba(5,150,105,0.05) 0%, transparent 60%); }
.kc.green:hover   { border-color: rgba(5,150,105,0.4); box-shadow: 0 16px 48px rgba(0,0,0,0.5), 0 0 30px rgba(5,150,105,0.07); }

.kc.blue  { border-color: rgba(29,78,216,0.25); }
.kc.blue::before  { background: linear-gradient(90deg, var(--blu2), transparent 70%); }
.kc.blue::after   { background: linear-gradient(160deg, rgba(29,78,216,0.05) 0%, transparent 60%); }
.kc.blue:hover    { border-color: rgba(29,78,216,0.4); box-shadow: 0 16px 48px rgba(0,0,0,0.5), 0 0 30px rgba(29,78,216,0.07); }

.kc-val {
  font-family: var(--ff);
  font-size: 2.3rem; font-weight: 700; line-height: 1.1;
  letter-spacing: -0.04em;
  animation: numIn 0.45s 0.1s ease both;
  animation-fill-mode: both;
  position: relative; z-index: 1;
}
.kc-lbl {
  font-family: var(--fi);
  font-size: 0.62rem; font-weight: 500;
  text-transform: uppercase; letter-spacing: 0.16em;
  color: var(--t3); margin-top: 5px;
  position: relative; z-index: 1;
}
.kc.gold   .kc-val { color: var(--g4); }
.kc.red    .kc-val { color: var(--red3); }
.kc.orange .kc-val { color: var(--ora3); }
.kc.green  .kc-val { color: var(--grn3); }
.kc.blue   .kc-val { color: var(--blu3); }

/* Section label */
.sl {
  font-family: var(--fi);
  font-size: 0.6rem; font-weight: 600;
  letter-spacing: 0.22em; text-transform: uppercase;
  color: var(--t4);
  display: flex; align-items: center; gap: 0.75rem;
  margin: 2rem 0 1.1rem;
}
.sl::after {
  content: ''; flex: 1; height: 1px;
  background: linear-gradient(90deg, var(--b2), transparent);
}

/* Parse stats */
.ps-row { display: flex; gap: 0.7rem; flex-wrap: wrap; margin: 1rem 0 1.5rem; }
.ps {
  background: var(--bg3);
  border: 1px solid var(--b2);
  border-radius: 10px;
  padding: 0.7rem 1.1rem 0.75rem;
  min-width: 88px;
  animation: popIn 0.35s ease both;
  animation-fill-mode: both;
  transition: border-color 0.2s, transform 0.15s;
  position: relative; overflow: hidden;
}
.ps:hover { transform: translateY(-2px); border-color: var(--b3); }
.ps::before { content:''; position:absolute; top:0; left:0; right:0; height:2px; }
.ps.red::before    { background: var(--red2); }
.ps.orange::before { background: var(--ora2); }
.ps.green::before  { background: var(--grn2); }
.ps.blue::before   { background: var(--blu2); }
.ps-val {
  font-family: var(--ff);
  font-size: 1.75rem; font-weight: 700; line-height: 1;
  letter-spacing: -0.03em;
}
.ps-lbl {
  font-family: var(--fi);
  font-size: 0.6rem; text-transform: uppercase;
  letter-spacing: 0.14em; color: var(--t3); margin-top: 4px;
}
.ps.red    .ps-val { color: var(--red3); }
.ps.orange .ps-val { color: var(--ora3); }
.ps.green  .ps-val { color: var(--grn3); }
.ps.blue   .ps-val { color: var(--blu3); }

/* Info card */
.ic {
  background: var(--bg3);
  border: 1px solid var(--b2);
  border-radius: 12px;
  padding: 1.5rem 1.75rem;
  font-family: var(--fi);
  font-size: 0.83rem; color: var(--t2); line-height: 1.9;
  animation: slideRight 0.4s ease both;
  animation-fill-mode: both;
}
.ic-title {
  font-family: var(--ff);
  font-size: 0.6rem; font-weight: 600;
  letter-spacing: 0.2em; text-transform: uppercase;
  color: var(--g3); margin-bottom: 0.4rem;
}
.ic b { color: var(--t1); font-weight: 500; }

/* Success banner */
.sb {
  background: linear-gradient(135deg,
    rgba(5,150,105,0.14) 0%, rgba(5,150,105,0.04) 100%);
  border: 1px solid rgba(5,150,105,0.3);
  border-radius: 10px;
  padding: 1.1rem 1.5rem;
  color: var(--grn3);
  font-family: var(--ff); font-size: 0.95rem; font-weight: 500;
  animation: successPop 0.5s cubic-bezier(0.34,1.56,0.64,1) both;
  display: flex; align-items: center; gap: 0.75rem;
}

/* All-clear */
.ac {
  background: rgba(5,150,105,0.05);
  border: 1px solid rgba(5,150,105,0.15);
  border-radius: 10px;
  padding: 2rem;
  text-align: center;
  color: var(--grn3);
  font-family: var(--ff); font-size: 1rem; font-weight: 500;
  letter-spacing: 0.02em;
}

/* Sidebar logo */
.logo {
  font-family: var(--ff);
  font-size: 1.2rem; font-weight: 700;
  letter-spacing: 0.04em; color: var(--g3);
  display: flex; align-items: center; gap: 0.5rem;
}
.logo-tag {
  font-family: var(--fi);
  font-size: 0.58rem; text-transform: uppercase;
  letter-spacing: 0.2em; color: var(--t4); margin-top: 3px;
}
.logo-rule {
  height: 1px; margin: 1.25rem 0;
  background: linear-gradient(90deg, var(--g2), transparent);
}

/* Sidebar vessel chip */
.vc {
  display: flex; align-items: center; justify-content: space-between;
  background: var(--bg3);
  border: 1px solid var(--b2);
  border-radius: 7px;
  padding: 0.55rem 0.85rem;
  margin-bottom: 0.3rem;
  transition: border-color 0.2s, background 0.2s;
  animation: slideRight 0.3s ease both;
  animation-fill-mode: both;
}
.vc:hover { border-color: var(--b3); background: var(--bg4); }
.vc.crit  { border-left: 2px solid var(--red2); }
.vc.warn  { border-left: 2px solid var(--ora2); }
.vc.safe  { border-left: 2px solid var(--grn2); }
.vc-name  {
  font-family: var(--ff); font-size: 0.75rem; font-weight: 600;
  color: var(--t1); white-space: nowrap; overflow: hidden;
  text-overflow: ellipsis; max-width: 120px;
}
.vc-tags  { display: flex; gap: 5px; align-items: center; flex-shrink: 0; }
.vc-tag   {
  font-family: var(--fm); font-size: 0.58rem; font-weight: 500;
  padding: 1px 6px; border-radius: 3px;
}
.vc-tag.od { background: var(--red0); color: var(--red3); }
.vc-tag.hp { background: var(--ora0); color: var(--ora3); }
.vc-tag.ok { background: var(--grn0); color: var(--grn3); }

/* Meta line */
.ml {
  display: flex; gap: 1.5rem; flex-wrap: wrap;
  font-family: var(--fm); font-size: 0.68rem; color: var(--t3);
  margin: 0.75rem 0 0;
}
.ml b { color: var(--t2); font-weight: 500; }
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# DATABASE
# ════════════════════════════════════════════════════════════════════
DB_PATH = Path("running_hours.db")

def get_db():
    conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS vessels (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS upload_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL, filename TEXT NOT NULL,
        file_hash TEXT NOT NULL, report_date TEXT,
        me_total_hrs INTEGER, me_this_month INTEGER,
        uploaded_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS components (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL, category TEXT NOT NULL,
        engine_label TEXT NOT NULL, unit TEXT NOT NULL,
        description TEXT NOT NULL, periodicity REAL,
        last_oh_date TEXT, last_oh_hrs REAL,
        hrs_since REAL, pct_used REAL, status TEXT NOT NULL,
        updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS other_equipment (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vessel_name TEXT NOT NULL, section TEXT NOT NULL,
        description TEXT NOT NULL, periodicity TEXT,
        last_date TEXT, run_hrs TEXT,
        updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE INDEX IF NOT EXISTS idx_cv ON components(vessel_name);
    CREATE INDEX IF NOT EXISTS idx_cs ON components(status);
    CREATE INDEX IF NOT EXISTS idx_ov ON other_equipment(vessel_name);
    """)
    conn.commit(); conn.close()

init_db()


# ════════════════════════════════════════════════════════════════════
# CONVERSION — .doc → .docx via LibreOffice (packages.txt)
# ════════════════════════════════════════════════════════════════════
def convert_doc_to_docx(file_bytes: bytes) -> bytes:
    soffice = shutil.which("soffice") or "/usr/bin/soffice"
    if not os.path.isfile(soffice):
        raise RuntimeError(
            "LibreOffice not found. Ensure packages.txt contains: libreoffice")
    with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as tmp:
        tmp.write(file_bytes); tmp_path = tmp.name
    out_dir   = tempfile.gettempdir()
    base      = os.path.splitext(os.path.basename(tmp_path))[0]
    docx_out  = os.path.join(out_dir, base + ".docx")
    profile   = f"file:///tmp/lo_{os.getpid()}_{os.urandom(4).hex()}"
    try:
        r = subprocess.run(
            [soffice, "--headless", "--norestore", "--nofirststartwizard",
             f"-env:UserInstallation={profile}",
             "--convert-to", "docx", tmp_path, "--outdir", out_dir],
            capture_output=True, timeout=120)
        if not os.path.exists(docx_out):
            raise RuntimeError(
                f"Conversion failed (exit {r.returncode}). "
                f"{r.stderr.decode('utf-8', errors='ignore')[:300]}")
        with open(docx_out, "rb") as f: return f.read()
    finally:
        for p in [tmp_path, docx_out]:
            try:
                if os.path.exists(p): os.unlink(p)
            except Exception: pass


# ════════════════════════════════════════════════════════════════════
# PARSER — deterministic, 100% integrity validated
# ════════════════════════════════════════════════════════════════════
def _clean_period(raw):
    if not raw: return None
    s = re.sub(r'\.(?=\d{3}(\.|$))', '', raw.strip())
    s = re.sub(r'[^0-9\.]', '', s)
    try: return float(s) if s else None
    except ValueError: return None

def _parse_date(raw):
    if not raw or raw.strip() in ('', 'N/A', 'n/a'): return None
    raw = re.sub(r'\s+', ' ', raw.strip().lstrip('[').rstrip(']'))
    if re.match(r'^\d+$', raw.strip()): return None
    rn = re.sub(r'\bSEPT\b', 'SEP', raw, flags=re.IGNORECASE)
    rn = re.sub(r'\bJUNE\b', 'JUN', rn,  flags=re.IGNORECASE)
    rn = re.sub(r'\bJULY\b', 'JUL', rn,  flags=re.IGNORECASE)
    fmts = ['%d %b %y','%d %B %y','%d %b %Y','%d %B %Y',
            '%d/%m/%y','%d/%m/%Y','%d-%m-%y','%d-%m-%Y',
            '%b %Y','%B %Y','%Y-%m-%d']
    for fmt in fmts:
        for v in (rn, rn.upper(), rn.title(), raw, raw.upper()):
            try: return datetime.strptime(v, fmt).strftime('%Y-%m-%d')
            except ValueError: pass
    return raw

def _parse_hrs(raw):
    if not raw or raw.strip() in ('', 'N/A', 'n/a'): return None
    for n in re.findall(r'\d[\d,]*', raw.replace('\n', ' ')):
        try:
            v = float(n.replace(',', ''))
            if v > 0: return v
        except ValueError: pass
    return None

def _status(hrs, period):
    if hrs is None or period is None or period == 0: return 'NO DATA'
    p = hrs / period
    if p > 1.0: return 'OVERDUE'
    if p >= 0.80: return 'HIGH PRIORITY'
    return 'OK'

def _pct(hrs, period):
    if hrs is None or period is None or period == 0: return 0.0
    return round(hrs / period, 4)

def parse_doc_bytes(docx_bytes: bytes) -> dict:
    from docx import Document
    warns = []
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
        tmp.write(docx_bytes); tmp_path = tmp.name
    try:
        doc = Document(tmp_path)
    except Exception as e:
        raise ValueError(f"Cannot open document: {e}")
    finally:
        try: os.unlink(tmp_path)
        except Exception: pass
    if not doc.tables:
        raise ValueError("No tables found — is this a TEC-004 report?")

    vessel_name = "UNKNOWN"; report_date = None
    for para in doc.paragraphs:
        txt = para.text.strip()
        if not txt: continue
        vm = re.search(
            r"Vessel[\u2019\u2018']?s?\s+Name\s*:\s*(?:MV\s+)?([A-Z][A-Z0-9 \-]+?)(?:\s{2,}|\t)",
            txt, re.IGNORECASE)
        dm = re.search(r"Date\s*:\s*(.+)", txt, re.IGNORECASE)
        if vm: vessel_name = vm.group(1).strip()
        if dm: report_date = _parse_date(dm.group(1).strip())
        if vm or dm: break
    if vessel_name == "UNKNOWN":
        warns.append("Could not extract vessel name from header.")

    me_total = me_month = None; components = []
    t0 = doc.tables[0]
    for cell in t0.rows[0].cells:
        txt = cell.text
        if m := re.search(r'Total Running Hours[\s:ǀ|]+([\d,]+)', txt, re.IGNORECASE):
            try: me_total = int(m.group(1).replace(',', ''))
            except ValueError: pass
        if m := re.search(r'This Month[\s:]+([\d,]+)', txt, re.IGNORECASE):
            try: me_month = int(m.group(1).replace(',', ''))
            except ValueError: pass

    cyl_cols = []
    if len(t0.rows) > 1:
        for ci, cell in enumerate(t0.rows[1].cells):
            if m := re.search(r'CYL\s*\.?\s*No\s*\.?\s*(\d+)', cell.text.strip(), re.IGNORECASE):
                lbl = f"Cyl {int(m.group(1))}"
                if not cyl_cols or cyl_cols[-1][1] != lbl:
                    cyl_cols.append((ci, lbl))

    rows = t0.rows; i = 2
    while i < len(rows) - 1:
        r1 = [c.text.strip() for c in rows[i].cells]
        r2 = [c.text.strip() for c in rows[i+1].cells] if i+1 < len(rows) else []
        name = r1[0] if r1 else ''
        if not name: i += 1; continue
        t1x = r1[2] if len(r1) > 2 else ''
        t2x = r2[2] if len(r2) > 2 else ''
        if t1x == '1' and t2x == '2' and r1[0] == (r2[0] if r2 else ''):
            period = _clean_period(r1[1] if len(r1) > 1 else '')
            for ci, lbl in cyl_cols:
                d = _parse_date(r1[ci]) if ci < len(r1) else None
                h = _parse_hrs(r2[ci])  if ci < len(r2) else None
                if d is None and h is None: continue
                components.append({
                    'category':'MAIN_ENGINE','engine_label':'ME','unit':lbl,
                    'description':name,'periodicity':period,
                    'last_oh_date':d,'last_oh_hrs':h,'hrs_since':h,
                    'pct_used':_pct(h,period),'status':_status(h,period)})
            i += 2
        else: i += 1

    other_equip = []
    if len(doc.tables) > 1:
        t1t = doc.tables[1]
        SKIP = {'TURBOCHARGER','AUXILIARY BOILER','COOLERS','EXH GAS  BOILER',
                'A/C & REFR. COMPRESSORS','MAIN AIR COMPRESSORS',
                'PERIODICTLY','DATE OF LAST INSPECTION','RUN HRS',
                'DATE OF LAST CLEANING','DATE','PERIODICITY'}
        for row in t1t.rows:
            cells = [c.text.strip() for c in row.cells]
            for sec, dc, datec, hrsc in [
                ('TURBOCHARGER / AUX BOILER',0,1,3),
                ('COOLERS / EXH GAS BOILER',5,6,8),
                ('A/C & COMPRESSORS',10,11,12)]:
                desc = cells[dc] if dc < len(cells) else ''
                if not desc or desc.upper() in SKIP: continue
                dv = cells[datec] if datec < len(cells) else ''
                hv = cells[hrsc]  if hrsc  < len(cells) else ''
                if dv or hv:
                    other_equip.append({'section':sec,'description':desc,
                                        'periodicity':'','last_date':dv,'run_hrs':hv})

    if len(doc.tables) > 2:
        t2t = doc.tables[2]; rows2 = t2t.rows; engine_blocks = []
        if rows2:
            hdr  = [c.text.strip() for c in rows2[0].cells]
            tot  = [c.text.strip() for c in rows2[2].cells] if len(rows2) > 2 else []
            seen = set()
            for ci, cell in enumerate(hdr):
                if m := re.search(r'Aux\.\s*Engine\s*No\.?\s*(\d+)', cell, re.IGNORECASE):
                    lbl = f"AUX-{int(m.group(1))}"
                    if lbl not in seen:
                        seen.add(lbl)
                        th = next((_parse_hrs(tot[j]) for j in range(ci,min(ci+14,len(tot)))
                                   if _parse_hrs(tot[j])), None)
                        engine_blocks.append((lbl, ci, th))
        cyl_map = {}
        if len(rows2) > 4:
            r4 = [c.text.strip() for c in rows2[4].cells]
            for ei,(elbl,estart,_) in enumerate(engine_blocks):
                eend = engine_blocks[ei+1][1] if ei+1 < len(engine_blocks) else len(r4)
                seen_c: list = []
                for ci in range(estart, eend):
                    if ci < len(r4):
                        try:
                            cn = int(r4[ci])
                            if cn not in seen_c:
                                seen_c.append(cn); cyl_map[ci] = (elbl, cn)
                        except ValueError: pass
        i2 = 5
        while i2 < len(rows2) - 1:
            r1 = [c.text.strip() for c in rows2[i2].cells]
            r2 = [c.text.strip() for c in rows2[i2+1].cells] if i2+1 < len(rows2) else []
            name = r1[0] if r1 else ''
            if not name: i2 += 1; continue
            t1t2 = r1[2] if len(r1) > 2 else ''
            t2t2 = r2[2] if len(r2) > 2 else ''
            if t1t2 in ('1','2') and r1[0] == (r2[0] if r2 else ''):
                period = _clean_period(r1[1] if len(r1) > 1 else '')
                for ci,(elbl,cn) in cyl_map.items():
                    d = _parse_date(r1[ci]) if ci < len(r1) else None
                    h = _parse_hrs(r2[ci])  if ci < len(r2) else None
                    if d is None and h is None: continue
                    components.append({
                        'category':'AUX_ENGINE','engine_label':elbl,'unit':f"Cyl {cn}",
                        'description':name,'periodicity':period,
                        'last_oh_date':d,'last_oh_hrs':h,'hrs_since':h,
                        'pct_used':_pct(h,period),'status':_status(h,period)})
                i2 += 2
            else: i2 += 1

    if len(doc.tables) > 3:
        t3t = doc.tables[3]; dglbls = ['D/G 1','D/G 2','D/G 3']
        for ridx, row in enumerate(t3t.rows):
            cells = [c.text.strip() for c in row.cells]
            if ridx == 0: continue
            for dc,pc,tc,ds in [(0,1,2,3),(9,10,11,12)]:
                desc  = cells[dc] if dc < len(cells) else ''
                per   = cells[pc] if pc < len(cells) else ''
                rtype = cells[tc] if tc < len(cells) else ''
                if not desc or rtype not in ('1','2'): continue
                for dgi,dglbl in enumerate(dglbls):
                    col = ds + dgi
                    val = cells[col] if col < len(cells) else ''
                    if not val: continue
                    key = f"{desc} — {dglbl}"
                    if rtype == '1':
                        other_equip.append({'section':'D/G EQUIPMENT','description':key,
                                            'periodicity':per,'last_date':_parse_date(val) or val,'run_hrs':''})
                    else:
                        for e in reversed(other_equip):
                            if e['description']==key and e['run_hrs']=='':
                                e['run_hrs'] = val; break
                        else:
                            other_equip.append({'section':'D/G EQUIPMENT','description':key,
                                                'periodicity':per,'last_date':'','run_hrs':val})

    return {
        'vessel_name':vessel_name,'report_date':report_date,
        'me_total_hrs':me_total,'me_this_month':me_month,
        'components':components,'other_equipment':other_equip,'warnings':warns}


# ════════════════════════════════════════════════════════════════════
# DB HELPERS
# ════════════════════════════════════════════════════════════════════
def save_parsed_data(parsed, filename, file_hash):
    conn = get_db(); c = conn.cursor()
    now = datetime.utcnow().isoformat(); v = parsed['vessel_name']
    c.execute("INSERT OR IGNORE INTO vessels(name,created_at) VALUES(?,?)", (v, now))
    c.execute("INSERT INTO upload_log(vessel_name,filename,file_hash,report_date,me_total_hrs,me_this_month,uploaded_at) VALUES(?,?,?,?,?,?,?)",
              (v,filename,file_hash,parsed['report_date'],parsed['me_total_hrs'],parsed['me_this_month'],now))
    c.execute("DELETE FROM components WHERE vessel_name=?", (v,))
    for comp in parsed['components']:
        c.execute("INSERT INTO components(vessel_name,category,engine_label,unit,description,periodicity,last_oh_date,last_oh_hrs,hrs_since,pct_used,status,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                  (v,comp['category'],comp['engine_label'],comp['unit'],comp['description'],
                   comp['periodicity'],comp['last_oh_date'],comp['last_oh_hrs'],
                   comp['hrs_since'],comp['pct_used'],comp['status'],now))
    c.execute("DELETE FROM other_equipment WHERE vessel_name=?", (v,))
    for oe in parsed['other_equipment']:
        c.execute("INSERT INTO other_equipment(vessel_name,section,description,periodicity,last_date,run_hrs,updated_at) VALUES(?,?,?,?,?,?,?)",
                  (v,oe['section'],oe['description'],oe.get('periodicity',''),
                   oe.get('last_date',''),oe.get('run_hrs',''),now))
    conn.commit(); conn.close()

@st.cache_data(ttl=10)
def get_all_vessels():
    conn = get_db()
    rows = conn.execute("SELECT name FROM vessels ORDER BY name").fetchall()
    conn.close(); return [r['name'] for r in rows]

@st.cache_data(ttl=10)
def get_components_df(vessel):
    conn = get_db()
    df = pd.read_sql_query(
        "SELECT * FROM components WHERE vessel_name=?",
        conn, params=(vessel,))
    conn.close(); return df

@st.cache_data(ttl=10)
def get_other_equip_df(vessel):
    conn = get_db()
    df = pd.read_sql_query(
        "SELECT * FROM other_equipment WHERE vessel_name=? ORDER BY section,description",
        conn, params=(vessel,))
    conn.close(); return df

@st.cache_data(ttl=10)
def get_upload_history(vessel):
    conn = get_db()
    df = pd.read_sql_query(
        "SELECT filename,report_date,me_total_hrs,me_this_month,uploaded_at "
        "FROM upload_log WHERE vessel_name=? ORDER BY uploaded_at DESC LIMIT 20",
        conn, params=(vessel,))
    conn.close(); return df

@st.cache_data(ttl=10)
def get_fleet_summary():
    conn = get_db()
    df = pd.read_sql_query("""
        SELECT c.vessel_name,
            COUNT(CASE WHEN c.status='OVERDUE'       THEN 1 END) AS overdue,
            COUNT(CASE WHEN c.status='HIGH PRIORITY' THEN 1 END) AS high_priority,
            COUNT(CASE WHEN c.status='OK'            THEN 1 END) AS ok,
            COUNT(*) AS total,
            MAX(u.uploaded_at)  AS last_upload,
            MAX(u.me_total_hrs) AS me_total_hrs,
            MAX(u.report_date)  AS report_date
        FROM components c
        LEFT JOIN upload_log u ON u.vessel_name=c.vessel_name
        GROUP BY c.vessel_name ORDER BY overdue DESC, high_priority DESC
    """, conn); conn.close(); return df


# ════════════════════════════════════════════════════════════════════
# TABLE RENDERING — correct sort + rich styling
# ════════════════════════════════════════════════════════════════════

# Status priority for sorting
_STATUS_ORDER = {'OVERDUE': 0, 'HIGH PRIORITY': 1, 'OK': 2, 'NO DATA': 3}

def _safe_float(x):
    """Safely convert to float, return None on failure."""
    try:
        v = float(x)
        return None if pd.isna(v) else v
    except (TypeError, ValueError):
        return None

def fmt_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Build display DataFrame from either raw parser output or DB query.
    Sorts: OVERDUE first → HIGH PRIORITY → OK → NO DATA.
    Within each status group: by pct_used descending (most critical first).
    Works with both raw dict lists converted to DataFrame and DB query results.
    """
    if df.empty:
        return pd.DataFrame(columns=['Status','Component','Engine','Unit',
                                     'Periodicity','Last O/H','Hrs Since','% Used'])

    # ── Sort ──────────────────────────────────────────────────────
    d = df.copy()
    d['_s'] = d['status'].map(lambda s: _STATUS_ORDER.get(str(s), 4))
    d['_p'] = d['pct_used'].apply(lambda x: _safe_float(x) or 0.0)
    d = d.sort_values(['_s', '_p'], ascending=[True, False]).drop(columns=['_s','_p'])

    # ── Format columns ────────────────────────────────────────────
    out = pd.DataFrame(index=range(len(d)))
    out['Status']      = d['status'].values
    out['Component']   = d['description'].values
    out['Engine']      = d['engine_label'].values
    out['Unit']        = d['unit'].values
    out['Periodicity'] = [
        f"{int(float(x)):,}" if _safe_float(x) and str(x) not in ('nan','None','') else 'N/A'
        for x in d['periodicity'].values
    ]
    out['Last O/H']    = [
        str(x) if x and str(x) not in ('nan','None','') else '—'
        for x in d['last_oh_date'].values
    ]
    out['Hrs Since']   = [
        f"{int(float(x)):,}" if _safe_float(x) else '—'
        for x in d['hrs_since'].values
    ]
    out['% Used']      = [
        f"{float(x)*100:.1f}%" if _safe_float(x) else '—'
        for x in d['pct_used'].values
    ]
    return out

def style_df(df: pd.DataFrame):
    """Apply rich row-level colour coding based on Status column."""
    def rs(row):
        s = str(row.get('Status', ''))
        if s == 'OVERDUE':
            return [
                'background:#3a0808;color:#fca5a5;font-weight:600',  # Status
                'background:#280606;color:#f87171',                   # Component
                'background:#280606;color:#e05252',                   # Engine
                'background:#280606;color:#e05252',                   # Unit
                'background:#280606;color:#9ca3af',                   # Periodicity
                'background:#280606;color:#9ca3af',                   # Last O/H
                'background:#280606;color:#fca5a5;font-weight:600',   # Hrs Since
                'background:#280606;color:#fca5a5;font-weight:700',   # % Used
            ]
        if s == 'HIGH PRIORITY':
            return [
                'background:#3a1a04;color:#fed7aa;font-weight:600',
                'background:#281204;color:#fb923c',
                'background:#281204;color:#d97706',
                'background:#281204;color:#d97706',
                'background:#281204;color:#9ca3af',
                'background:#281204;color:#9ca3af',
                'background:#281204;color:#fed7aa;font-weight:600',
                'background:#281204;color:#fed7aa;font-weight:700',
            ]
        if s == 'OK':
            return [
                'background:#061a10;color:#6ee7b7;font-weight:500',
                'background:#041209;color:#34d399',
                'background:#041209;color:#059669',
                'background:#041209;color:#059669',
                'background:#041209;color:#6b7280',
                'background:#041209;color:#6b7280',
                'background:#041209;color:#6ee7b7',
                'background:#041209;color:#6ee7b7',
            ]
        # NO DATA / default
        return ['background:#080f1a;color:#374151'] * 8

    return df.style.apply(rs, axis=1)

def show_table(df: pd.DataFrame, height: int = None):
    """Render a component table — sorted, styled, with correct column widths."""
    if isinstance(df, list):
        df = pd.DataFrame(df)
    if df.empty:
        st.info("No data to display.")
        return
    formatted = fmt_df(df)
    h = height or min(720, 38 * (len(formatted) + 1) + 4)
    st.dataframe(
        style_df(formatted),
        use_container_width=True,
        hide_index=True,
        height=h,
        column_config={
            "Status":      st.column_config.TextColumn("Status",      width=130),
            "Component":   st.column_config.TextColumn("Component",   width=200),
            "Engine":      st.column_config.TextColumn("Engine",      width=80),
            "Unit":        st.column_config.TextColumn("Unit",        width=70),
            "Periodicity": st.column_config.TextColumn("Periodicity", width=100),
            "Last O/H":    st.column_config.TextColumn("Last O/H",    width=100),
            "Hrs Since":   st.column_config.TextColumn("Hrs Since",   width=90),
            "% Used":      st.column_config.TextColumn("% Used",      width=80),
        }
    )


# ════════════════════════════════════════════════════════════════════
# UI HELPERS
# ════════════════════════════════════════════════════════════════════
def kpi(val, lbl, color="gold", delay=0):
    return (f'<div class="kc {color}" style="animation-delay:{delay}s">'
            f'<div class="kc-val">{val}</div>'
            f'<div class="kc-lbl">{lbl}</div></div>')

def ph(icon, title, eyebrow=""):
    eye = f'<div class="ph-eyebrow">{eyebrow}</div>' if eyebrow else ''
    return (f'<div class="ph">{eye}'
            f'<h1>{icon}&nbsp;&nbsp;{title}</h1></div>'
            f'<div class="ph-line"></div>')

def sl(text):
    return f'<div class="sl">{text}</div>'


# ════════════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
    <div class="logo">⚓ FLEET MONITOR</div>
    <div class="logo-tag">Running Hours Management System</div>
    <div class="logo-rule"></div>
    """, unsafe_allow_html=True)

    page = st.selectbox("nav",
        ["🗺️  Fleet Overview", "🚢  Vessel Detail",
         "📤  Upload Report",  "📋  Upload History"],
        label_visibility="collapsed")

    st.markdown("<br>", unsafe_allow_html=True)

    vessels = get_all_vessels()
    selected_vessel = st.selectbox("Active Vessel", vessels) if vessels else None
    if not vessels:
        st.info("No data yet — upload a report to begin.")

    if vessels:
        smry = get_fleet_summary()
        if not smry.empty:
            st.markdown('<div class="logo-rule"></div>', unsafe_allow_html=True)
            st.markdown(
                '<div style="font-family:var(--fi);font-size:0.58rem;text-transform:uppercase;'
                'letter-spacing:0.2em;color:var(--t4);margin-bottom:0.6rem;">Vessel Status</div>',
                unsafe_allow_html=True)
            for idx, (_, row) in enumerate(smry.iterrows()):
                od = int(row['overdue']); hp = int(row['high_priority'])
                cls = 'crit' if od > 0 else ('warn' if hp > 0 else 'safe')
                tags = ''
                if od > 0:
                    tags += f'<span class="vc-tag od">{od} OD</span>'
                if hp > 0:
                    tags += f'<span class="vc-tag hp">{hp} HP</span>'
                if od == 0 and hp == 0:
                    tags += '<span class="vc-tag ok">✓</span>'
                st.markdown(
                    f'<div class="vc {cls}" style="animation-delay:{idx*0.05}s">'
                    f'<div class="vc-name">{row["vessel_name"]}</div>'
                    f'<div class="vc-tags">{tags}</div></div>',
                    unsafe_allow_html=True)

    st.markdown('<div class="logo-rule"></div>', unsafe_allow_html=True)
    db_kb = DB_PATH.stat().st_size / 1024 if DB_PATH.exists() else 0
    st.markdown(
        f'<div style="font-family:var(--fm);font-size:0.6rem;color:var(--t4)">'
        f'db {db_kb:.0f} kb · {len(vessels)} vessels · v4.1</div>',
        unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
# PAGE: UPLOAD REPORT
# ════════════════════════════════════════════════════════════════════
if page == "📤  Upload Report":
    st.markdown(ph("📤", "Upload Report", "TEC-004 · Running Hours"), unsafe_allow_html=True)

    col_up, col_info = st.columns([3, 2], gap="large")
    with col_up:
        uploaded = st.file_uploader(
            "file", type=["doc"], label_visibility="collapsed")
    with col_info:
        st.markdown("""
        <div class="ic">
          <div class="ic-title">Accepted Format</div>
          TEC-004 Running Hours Monthly Report<br>
          Any vessel &nbsp;·&nbsp; <b>.doc format only</b><br><br>
          <div class="ic-title">Extracted Data</div>
          ✦ Vessel name &amp; report date<br>
          ✦ M/E total &amp; monthly running hours<br>
          ✦ All M/E components — dates &amp; hours<br>
          ✦ Aux engines (3 engines × up to 7 cyl)<br>
          ✦ Turbocharger, coolers, D/G equipment<br>
          ✦ Status auto-computed per periodicity
        </div>""", unsafe_allow_html=True)

    if uploaded:
        raw_bytes = uploaded.read()
        file_hash = hashlib.md5(raw_bytes).hexdigest()

        with st.spinner("Converting .doc → .docx and parsing…"):
            try:
                docx_bytes = convert_doc_to_docx(raw_bytes)
            except Exception as e:
                st.error(f"**Conversion failed.** `{e}`")
                st.stop()
            try:
                parsed = parse_doc_bytes(docx_bytes)
            except ValueError as e:
                st.error(f"**Parse failed.** `{e}`")
                st.stop()

        comps  = parsed['components']
        n_comp = len(comps)
        n_od   = sum(1 for c in comps if c['status'] == 'OVERDUE')
        n_hp   = sum(1 for c in comps if c['status'] == 'HIGH PRIORITY')
        n_ok   = sum(1 for c in comps if c['status'] == 'OK')
        n_oe   = len(parsed['other_equipment'])

        st.markdown(sl("Parse Preview — confirm before saving"), unsafe_allow_html=True)

        m1,m2,m3,m4,m5 = st.columns(5)
        m1.metric("Vessel",          parsed['vessel_name'])
        m2.metric("Report Date",     parsed['report_date'] or "—")
        m3.metric("M/E Total Hrs",   f"{parsed['me_total_hrs']:,}"  if parsed['me_total_hrs']  else "—")
        m4.metric("M/E This Month",  f"{parsed['me_this_month']:,}" if parsed['me_this_month'] else "—")
        m5.metric("Components",      n_comp)

        st.markdown(f"""
        <div class="ps-row">
          <div class="ps red"    style="animation-delay:0s">   <div class="ps-val">{n_od}</div><div class="ps-lbl">Overdue</div></div>
          <div class="ps orange" style="animation-delay:0.07s"><div class="ps-val">{n_hp}</div><div class="ps-lbl">High Priority</div></div>
          <div class="ps green"  style="animation-delay:0.14s"><div class="ps-val">{n_ok}</div><div class="ps-lbl">OK</div></div>
          <div class="ps blue"   style="animation-delay:0.21s"><div class="ps-val">{n_oe}</div><div class="ps-lbl">Other Equip</div></div>
        </div>""", unsafe_allow_html=True)

        for w in parsed['warnings']:
            st.warning(f"⚠ {w}")

        if comps:
            with st.expander(f"Preview all {n_comp} parsed components", expanded=True):
                show_table(pd.DataFrame(comps))

        st.markdown("---")
        col_btn, _ = st.columns([1, 4])
        with col_btn:
            if st.button("✅  CONFIRM & SAVE", use_container_width=True):
                save_parsed_data(parsed, uploaded.name, file_hash)
                for fn in [get_all_vessels, get_components_df,
                            get_other_equip_df, get_fleet_summary]:
                    fn.clear()
                st.markdown(f"""
                <div class="sb">
                  <span style="font-size:1.4rem">✓</span>
                  <span><b>{parsed['vessel_name']}</b> saved — {n_comp} components
                  · {n_od} overdue · {n_hp} high priority</span>
                </div>""", unsafe_allow_html=True)
                st.balloons()


# ════════════════════════════════════════════════════════════════════
# PAGE: FLEET OVERVIEW
# ════════════════════════════════════════════════════════════════════
elif page == "🗺️  Fleet Overview":
    st.markdown(ph("🗺️", "Fleet Overview", "All vessels · Live status"), unsafe_allow_html=True)

    summary = get_fleet_summary()
    if summary.empty:
        st.info("No data loaded yet. Upload a .doc report to begin.")
        st.stop()

    tv  = len(summary)
    tc  = int(summary['total'].sum())
    tod = int(summary['overdue'].sum())
    thp = int(summary['high_priority'].sum())
    tok = int(summary['ok'].sum())

    k1,k2,k3,k4,k5 = st.columns(5)
    for col,(val,lbl,clr,dly) in zip([k1,k2,k3,k4,k5],[
        (tv,  "Vessels",       "blue",   0),
        (tc,  "Components",    "gold",   0.07),
        (tod, "Overdue",       "red",    0.14),
        (thp, "High Priority", "orange", 0.21),
        (tok, "OK",            "green",  0.28),
    ]):
        with col: st.markdown(kpi(val,lbl,clr,dly), unsafe_allow_html=True)

    st.markdown(sl("Fleet Status Table"), unsafe_allow_html=True)

    disp = summary[['vessel_name','overdue','high_priority','ok','total',
                    'me_total_hrs','last_upload']].copy()
    disp.columns = ['Vessel','Overdue','High Priority','OK','Total','M/E Total Hrs','Last Upload']
    disp['M/E Total Hrs'] = disp['M/E Total Hrs'].apply(
        lambda x: f"{int(x):,}" if pd.notna(x) else '—')
    disp['Last Upload'] = pd.to_datetime(
        disp['Last Upload'], errors='coerce').dt.strftime('%Y-%m-%d').fillna('—')

    def fleet_style(row):
        if row['Overdue']>0:
            return ['background:#280606;color:#fca5a5;font-weight:600',
                    'background:#280606;color:#f87171;font-weight:700',
                    'background:#280606;color:#e05252',
                    'background:#280606;color:#34d399',
                    'background:#280606;color:#6b7280',
                    'background:#280606;color:#6b7280',
                    'background:#280606;color:#6b7280']
        if row['High Priority']>0:
            return ['background:#1e0f04;color:#fed7aa;font-weight:600',
                    'background:#1e0f04;color:#e05252',
                    'background:#1e0f04;color:#fb923c;font-weight:700',
                    'background:#1e0f04;color:#34d399',
                    'background:#1e0f04;color:#6b7280',
                    'background:#1e0f04;color:#6b7280',
                    'background:#1e0f04;color:#6b7280']
        return ['background:#041209;color:#6ee7b7;font-weight:500',
                'background:#041209;color:#e05252',
                'background:#041209;color:#fb923c',
                'background:#041209;color:#34d399;font-weight:700',
                'background:#041209;color:#6b7280',
                'background:#041209;color:#6b7280',
                'background:#041209;color:#6b7280']

    st.dataframe(
        disp.style.apply(fleet_style, axis=1),
        use_container_width=True, hide_index=True,
        height=min(600, 38*(len(disp)+1)+3))

    st.markdown(sl("Per-Vessel Breakdown"), unsafe_allow_html=True)
    for _, row in summary.iterrows():
        od = int(row['overdue']); hp = int(row['high_priority']); ok = int(row['ok'])
        icon = "🔴" if od>0 else ("🟡" if hp>0 else "🟢")
        with st.expander(
            f"{icon}  {row['vessel_name']}   —   "
            f"{od} overdue · {hp} high priority · {ok} OK",
            expanded=False):
            cc = get_components_df(row['vessel_name'])
            if not cc.empty:
                ta, tb = st.tabs(["🔴  Overdue", "🟡  High Priority"])
                with ta:
                    od_df = cc[cc['status']=='OVERDUE']
                    if od_df.empty:
                        st.markdown('<div class="ac">✓ No overdue items</div>', unsafe_allow_html=True)
                    else:
                        show_table(od_df)
                with tb:
                    hp_df = cc[cc['status']=='HIGH PRIORITY']
                    if hp_df.empty:
                        st.markdown('<div class="ac">✓ No high-priority items</div>', unsafe_allow_html=True)
                    else:
                        show_table(hp_df)


# ════════════════════════════════════════════════════════════════════
# PAGE: VESSEL DETAIL
# ════════════════════════════════════════════════════════════════════
elif page == "🚢  Vessel Detail":
    if not selected_vessel:
        st.info("Select a vessel from the sidebar."); st.stop()

    st.markdown(ph("🚢", selected_vessel, "Component Analysis"), unsafe_allow_html=True)

    df = get_components_df(selected_vessel)
    oe = get_other_equip_df(selected_vessel)
    if df.empty:
        st.info("No component data for this vessel."); st.stop()

    n_tot = len(df)
    n_od  = int((df['status']=='OVERDUE').sum())
    n_hp  = int((df['status']=='HIGH PRIORITY').sum())
    n_ok  = int((df['status']=='OK').sum())
    n_nd  = int((df['status']=='NO DATA').sum())

    k1,k2,k3,k4,k5 = st.columns(5)
    for col,(val,lbl,clr,dly) in zip([k1,k2,k3,k4,k5],[
        (n_tot,"Total",         "gold",   0),
        (n_od, "Overdue",       "red",    0.07),
        (n_hp, "High Priority", "orange", 0.14),
        (n_ok, "OK",            "green",  0.21),
        (n_nd, "No Data",       "blue",   0.28),
    ]):
        with col: st.markdown(kpi(val,lbl,clr,dly), unsafe_allow_html=True)

    hist = get_upload_history(selected_vessel)
    if not hist.empty:
        last = hist.iloc[0]
        mt = f"{int(last['me_total_hrs']):,}" if pd.notna(last['me_total_hrs']) else "—"
        mm = f"{int(last['me_this_month']):,}" if pd.notna(last['me_this_month']) else "—"
        st.markdown(f"""
        <div class="ml">
          <span>📄 <b>{last['filename']}</b></span>
          <span>Report: <b>{last['report_date'] or '—'}</b></span>
          <span>M/E: <b>{mt}</b> total · <b>{mm}</b> this month</span>
          <span>Uploaded: <b>{str(last['uploaded_at'])[:16]}</b></span>
        </div>""", unsafe_allow_html=True)

    st.markdown("---")
    tabs = st.tabs([
        "⚠️  Alerts",
        "⚙️  Main Engine",
        "🔩  Aux Engines",
        "🛠️  Other Equipment",
        "📊  All Components"
    ])

    with tabs[0]:
        st.markdown(sl("Overdue & High Priority — Action Required"), unsafe_allow_html=True)
        alerts = df[df['status'].isin(['OVERDUE','HIGH PRIORITY'])]
        if alerts.empty:
            st.markdown(
                '<div class="ac">✓ All components within acceptable limits — no action required</div>',
                unsafe_allow_html=True)
        else:
            show_table(alerts)

    with tabs[1]:
        me = df[df['category']=='MAIN_ENGINE']
        if me.empty:
            st.info("No Main Engine data.")
        else:
            st.markdown(sl("Main Engine Components"), unsafe_allow_html=True)
            sel = st.selectbox(
                "Filter by component",
                ['ALL'] + sorted(me['description'].unique().tolist()),
                key="me_f")
            show_table(me if sel=='ALL' else me[me['description']==sel])

    with tabs[2]:
        aux = df[df['category']=='AUX_ENGINE']
        if aux.empty:
            st.info("No Aux Engine data.")
        else:
            st.markdown(sl("Auxiliary Engine Components"), unsafe_allow_html=True)
            sel = st.selectbox(
                "Filter by engine",
                ['ALL'] + sorted(aux['engine_label'].unique().tolist()),
                key="aux_f")
            show_table(aux if sel=='ALL' else aux[aux['engine_label']==sel])

    with tabs[3]:
        if oe.empty:
            st.info("No other equipment data.")
        else:
            for sec in sorted(oe['section'].unique()):
                st.markdown(sl(sec), unsafe_allow_html=True)
                sd = oe[oe['section']==sec][['description','periodicity','last_date','run_hrs']].copy()
                sd.columns = ['Description','Periodicity','Last Date','Run Hrs']
                st.dataframe(sd, use_container_width=True, hide_index=True)

    with tabs[4]:
        st.markdown(sl("All Component Records"), unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            sf = st.multiselect(
                "Filter by status",
                ['OVERDUE','HIGH PRIORITY','OK','NO DATA'],
                default=['OVERDUE','HIGH PRIORITY','OK','NO DATA'],
                key="all_s")
        with c2:
            cf = st.multiselect(
                "Filter by category",
                ['MAIN_ENGINE','AUX_ENGINE'],
                default=['MAIN_ENGINE','AUX_ENGINE'],
                key="all_c")
        show_table(df[df['status'].isin(sf) & df['category'].isin(cf)])


# ════════════════════════════════════════════════════════════════════
# PAGE: UPLOAD HISTORY
# ════════════════════════════════════════════════════════════════════
elif page == "📋  Upload History":
    st.markdown(ph("📋", "Upload History", "Audit Trail"), unsafe_allow_html=True)

    if not selected_vessel:
        st.info("Select a vessel from the sidebar."); st.stop()

    st.markdown(sl(f"{selected_vessel}"), unsafe_allow_html=True)
    hist = get_upload_history(selected_vessel)
    if hist.empty:
        st.info("No upload history for this vessel.")
    else:
        d = hist.copy()
        d.columns = ['Filename','Report Date','M/E Total Hrs','M/E This Month','Uploaded At']
        d['M/E Total Hrs']  = d['M/E Total Hrs'].apply(
            lambda x: f"{int(x):,}" if pd.notna(x) else '—')
        d['M/E This Month'] = d['M/E This Month'].apply(
            lambda x: f"{int(x):,}" if pd.notna(x) else '—')
        st.dataframe(d, use_container_width=True, hide_index=True)
