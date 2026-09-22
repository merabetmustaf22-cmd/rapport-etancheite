import streamlit as st
import pandas as pd
from datetime import date, datetime
import os
import json
import io
import zipfile
import re
from PIL import Image, ImageDraw, ImageFont
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Suivi Chantier Étanchéité",
    page_icon="🏗️",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# --- CSS COMPLET SPÉCIAL SMARTPHONE & INVERSION DES CASES ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        -webkit-tap-highlight-color: transparent;
    }

    .stApp {
        background-color: #0B1120 !important;
        overflow-x: hidden !important;
    }
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 2.5rem !important;
        padding-left: 0.8rem !important;
        padding-right: 0.8rem !important;
        max-width: 100% !important;
    }

    /* TITRES ET TEXTES SUR FOND NOIR */
    .stMarkdown, .stMarkdown p, .stCaption, .stCaption p, [data-testid="stMarkdownContainer"] p {
        color: #FFFFFF !important;
    }
    h1, h2, h3, h4, h5, h6 {
        color: #FFFFFF !important;
        font-weight: 800 !important;
        word-break: normal !important;
        hyphens: none !important;
    }
    .stWidgetLabel p, [data-testid="stWidgetLabel"] p, label p {
        color: #FFFFFF !important;
        font-size: 15px !important;
        font-weight: 700 !important;
        margin-bottom: 5px !important;
    }

    /* ========================================================= */
    /* TOUTES LES CASES : FOND BLANC PUR + TEXTE NOIR NET       */
    /* ========================================================= */
    div[data-testid="stDateInput"] div[data-baseweb="input"],
    div[data-testid="stDateInput"] input,
    div[data-testid="stTextInput"] div[data-baseweb="input"],
    div[data-testid="stTextInput"] input,
    div[data-testid="stNumberInput"] div[data-baseweb="input"],
    div[data-testid="stNumberInput"] input,
    div[data-baseweb="base-input"],
    div[data-baseweb="input"] {
        background-color: #FFFFFF !important;
        background: #FFFFFF !important;
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
        border: 2px solid #CBD5E1 !important;
        border-radius: 10px !important;
        font-size: 16px !important;
        font-weight: 800 !important;
        min-height: 48px !important;
    }

    div[data-testid="stDateInput"] input {
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
        font-size: 16px !important;
        font-weight: 800 !important;
        padding-left: 14px !important;
    }

    .stSelectbox div[data-baseweb="select"] {
        background-color: #FFFFFF !important;
        border: 2px solid #CBD5E1 !important;
        border-radius: 10px !important;
        min-height: 48px !important;
    }
    .stSelectbox div[data-baseweb="select"] * {
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
        font-size: 15px !important;
        font-weight: 700 !important;
    }
    .stMultiSelect div[data-baseweb="select"] {
        background-color: #FFFFFF !important;
        border: 2px solid #CBD5E1 !important;
        border-radius: 10px !important;
        min-height: 48px !important;
    }
    .stMultiSelect [data-baseweb="tag"] {
        background-color: #0F766E !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
    }

    /* MENUS DÉROULANTS */
    div[data-baseweb="popover"], ul[role="listbox"], li[role="option"] {
        background-color: #FFFFFF !important;
    }
    div[data-baseweb="popover"] *, ul[role="listbox"] *, li[role="option"] * {
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
        font-weight: 700 !important;
        font-size: 15px !important;
    }

    /* CALENDRIER DE SÉLECTION */
    div[data-baseweb="calendar"], div[data-baseweb="calendar"] * {
        color: #0F172A !important;
        -webkit-text-fill-color: #0F172A !important;
        font-weight: 700 !important;
    }
    div[data-baseweb="calendar"] button[aria-selected="true"] {
        background-color: #0F766E !important;
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
    }

    /* ========================================================= */
    /* LES ONGLETS (TABS) EN HAUT : NETS ET TRÈS VISIBLES        */
    /* ========================================================= */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px !important;
        margin-bottom: 12px !important;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #1E293B !important;
        border-radius: 8px 8px 0px 0px !important;
        border: 1px solid #334155 !important;
        padding: 10px 16px !important;
        height: 48px !important;
    }
    .stTabs [data-baseweb="tab"] p {
        color: #FFFFFF !important;
        font-weight: 800 !important;
        font-size: 14px !important;
    }
    .stTabs [aria-selected="true"] {
        background-color: #0F172A !important;
        border-bottom: 3px solid #38BDF8 !important;
    }
    .stTabs [aria-selected="true"] p {
        color: #38BDF8 !important;
        font-weight: 900 !important;
    }

    /* EN-TÊTE */
    .header-cadre {
        background-color: #0F172A;
        border-radius: 12px;
        padding: 14px 12px;
        margin-bottom: 14px;
        border: 1px solid #334155;
    }
    .header-title {
        font-size: 18px !important;
        line-height: 1.3 !important;
        font-weight: 900 !important;
        margin: 0 !important;
        color: #FFFFFF !important;
    }
    .header-sub {
        font-size: 12px !important;
        color: #38BDF8 !important;
        margin-top: 4px !important;
        font-weight: 700 !important;
    }
    .badge-pro {
        background: rgba(13, 148, 136, 0.3);
        border: 1px solid #14B8A6;
        color: #5EEAD4 !important;
        padding: 4px 8px;
        border-radius: 6px;
        font-size: 11px;
        font-weight: 800;
        white-space: nowrap !important;
    }

    /* UPLOADER PHOTOS */
    [data-testid="stFileUploader"] section {
        background-color: #FFFFFF !important;
        border: 2px dashed #94A3B8 !important;
        border-radius: 12px !important;
        padding: 16px !important;
    }
    [data-testid="stFileUploader"] section * {
        color: #0F172A !important;
        font-weight: 700 !important;
    }
    [data-testid="stFileUploader"] section button {
        background-color: #0284C7 !important;
        border: none !important;
        border-radius: 8px !important;
        min-height: 44px !important;
    }
    [data-testid="stFileUploader"] section button * {
        color: #FFFFFF !important;
        font-weight: 700 !important;
    }

    /* BOUTONS GENERAUX */
    div.stButton > button {
        background-color: #0284C7 !important;
        color: #FFFFFF !important;
        font-size: 15px !important;
        font-weight: 800 !important;
        border-radius: 10px !important;
        border: none !important;
        min-height: 48px !important;
        width: 100% !important;
    }
    div.stButton > button * {
        color: #FFFFFF !important;
        font-weight: 800 !important;
    }

    .btn-supprimer button {
        background-color: #DC2626 !important;
        color: #FFFFFF !important;
        min-height: 40px !important;
        font-size: 13px !important;
        font-weight: 800 !important;
        border-radius: 8px !important;
    }

    .btn-valider button {
        background: linear-gradient(135deg, #059669 0%, #10B981 100%) !important;
        color: #FFFFFF !important;
        font-size: 16px !important;
        font-weight: 900 !important;
        border-radius: 12px !important;
        border: none !important;
        min-height: 54px !important;
        width: 100% !important;
        box-shadow: 0 4px 14px rgba(16, 185, 129, 0.4) !important;
    }

    div.stDownloadButton > button {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        font-size: 14px !important;
        font-weight: 800 !important;
        border-radius: 10px !important;
        border: 2px solid #CBD5E1 !important;
        min-height: 48px !important;
        width: 100% !important;
    }
    div.stDownloadButton > button * {
        color: #0F172A !important;
        font-weight: 800 !important;
    }

    /* CARTES DES INTERVENTIONS */
    .card-intervention {
        background-color: #111827 !important;
        border-radius: 12px !important;
        border: 1px solid #334155 !important;
        padding: 14px !important;
        margin-bottom: 14px !important;
        border-left: 5px solid #10B981 !important;
    }
    .card-text-white {
        color: #FFFFFF !important;
        font-size: 14px !important;
        font-weight: 600 !important;
        margin: 6px 0 !important;
    }
    .card-text-cyan {
        color: #38BDF8 !important;
        font-size: 13px !important;
        font-weight: 700 !important;
    }
    .card-text-note {
        color: #F8FAFC !important;
        font-size: 14px !important;
        font-style: italic !important;
        margin-top: 6px !important;
    }

    .tag-projet {
        font-size: 15px !important;
        font-weight: 900 !important;
        color: #FFFFFF !important;
    }
    .tag-date {
        font-size: 12px !important;
        font-weight: 700 !important;
        color: #FFFFFF !important;
        background: #1E293B !important;
        padding: 4px 8px !important;
        border-radius: 6px !important;
        border: 1px solid #475569 !important;
    }
    .tag-corps {
        display: inline-block !important;
        background-color: #1E293B !important;
        color: #FFFFFF !important;
        border: 1px solid #475569 !important;
        padding: 4px 8px !important;
        border-radius: 6px !important;
        font-size: 12px !important;
        font-weight: 800 !important;
    }
    .tag-rendement {
        display: inline-block !important;
        background-color: #064E3B !important;
        color: #6EE7B7 !important;
        border: 1px solid #059669 !important;
        padding: 4px 8px !important;
        border-radius: 6px !important;
        font-size: 12px !important;
        font-weight: 900 !important;
        margin-left: 4px !important;
    }
    .tag-materiau {
        display: inline-block !important;
        background-color: #082F49 !important;
        color: #7DD3FC !important;
        border: 1px solid #0284C7 !important;
        padding: 4px 8px !important;
        border-radius: 6px !important;
        font-size: 12px !important;
        font-weight: 800 !important;
        margin: 2px 3px 2px 0 !important;
    }
    hr {
        border-color: #334155 !important;
        margin: 14px 0 !important;
    }
    </style>
""", unsafe_allow_html=True)

PHOTOS_BASE_DIR = "photos_chantier"
os.makedirs(PHOTOS_BASE_DIR, exist_ok=True)
CONFIG_FILE = "config_chantier.json"
CSV_FILE = "suivi_journalier_chantiers.csv"
ADMIN_PIN = "2026"

COLONNES_OFFICIELLES = [
    "Date", "Chantier", "Corps_d_etat", "Phase", 
    "Rendement", "Unite", "Consommation",
    "Effectif", "Nb_Ouvriers", "Photos", "Nb_Photos", "Legende"
]

CONFIG_DEFAUT = {
    "chantiers": [
        "CAC-31-24(oran)",
        "HOB-342-08-24(bechar)",
        "IZM-31-24(CES)",
        "LCP-31-24",
        "VILLA Hasnaoui Outhman",
        "ESC-16-24(ALGER)"
    ],
    "macons": [
        "ADDA Abbess", "MEKHACHEF DJAMEL", "MESTEFAOUI AHMED", "ARGOUB HALIM",
        "FEHIM CHIBANI AZZOUZ", "TAIBI REDA", "ABED OMAR", "ZEGHDAN ABDELKADER",
        "BAGHDADI ALI", "BOUKHELIF KAMEL", "MOKHTARI Omar", "GHEZINI Habib",
        "BERACHEMI AHMED", "ARAR AISSA", "MOKHTARI Djelloul", "HAFDI Rachid",
        "BENHAMMADI Mohamed", "MOUISSI WALID", "BENSEMICHA MEROUANE", "TOUATI Zouaoui",
        "GHRIBI MOHAMED", "TAHAR BOUZIAN YOUCEF"
    ],
    "taches": [
        "BACHE A EAU", "BRICOL", "BRICOL ELASTOTEK", "BRICOL SILICONE", "BRICOL SOUS CARRELAGE",
        "coupe-feu", "Couvre-joint", "DALLE Cheminée", "ELASTOTEK", "Forme de pente",
        "GOURGE", "JOINT DE DILATATION", "PARE-VAPEUR", "PAX", "SOKLE PARE-VAPEUR",
        "SOUS CARRELAGE", "BRICOL PARE-VAPEUR", "ELASTOTEK SAUPOUDRAGE", "BRICOL Cheminée"
    ],
    "materiaux": [
        "ELASTOTEK",
        "MORCEM DRY F",
        "MORCEM DRY SF",
        "EPOTEK",
        "FLINKOTE",
        "FILTEK",
        "ARMA TEK PP",
        "TEKWELD SEA",
        "QUARTEK",
        "Primaire d'accrochage",
        "PAX (Rouleau bitume)",
        "Micro-béton B300",
        "Silicone / Mastic joint",
        "Autre"
    ]
}

def clean_folder_name(name):
    clean = name.split('(')[0].strip()
    return re.sub(r'[^a-zA-Z0-9_-]', '_', clean)

def charger_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                for cle in ["chantiers", "macons", "taches", "materiaux"]:
                    if cle not in data:
                        data[cle] = CONFIG_DEFAUT[cle]
                return data
        except Exception:
            return CONFIG_DEFAUT.copy()
    else:
        sauvegarder_config(CONFIG_DEFAUT)
        return CONFIG_DEFAUT.copy()

def sauvegarder_config(nouvelle_config):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(nouvelle_config, f, ensure_ascii=False, indent=4)
        return True
    except Exception:
        return False

def charger_donnees():
    if not os.path.exists(CSV_FILE):
        return None
    try:
        df = pd.read_csv(CSV_FILE, sep=';', encoding='utf-8-sig', on_bad_lines='skip')
        if len(df.columns) < 5:
            df = pd.read_csv(CSV_FILE, sep=',', encoding='utf-8-sig', on_bad_lines='skip')
        for c in COLONNES_OFFICIELLES:
            if c not in df.columns:
                df[c] = ""
        return df
    except Exception:
        return None

def sauvegarder_donnees(df_to_save):
    try:
        df_to_save.to_csv(CSV_FILE, sep=';', index=False, encoding='utf-8-sig')
        return True
    except Exception:
        return False

def appliquer_watermark(image_file, chantier, tache, phase, date_str):
    try:
        img = Image.open(image_file).convert("RGB")
        w, h = img.size

        overlay = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        draw = ImageDraw.Draw(overlay)

        band_h = max(int(h * 0.11), 80)
        draw.rectangle([(0, h - band_h), (w, h)], fill=(15, 23, 42, 210))

        font_size_titre = max(int(band_h * 0.32), 16)
        font_size_sub = max(int(band_h * 0.24), 13)

        try:
            font_t = ImageFont.truetype("arial.ttf", font_size_titre)
            font_s = ImageFont.truetype("arial.ttf", font_size_sub)
        except Exception:
            font_t = ImageFont.load_default()
            font_s = ImageFont.load_default()

        heure_actuelle = datetime.now().strftime("%H:%M")
        ligne_1 = f"PROJET : {chantier}   |   DATE : {date_str} a {heure_actuelle}"
        ligne_2 = f"OUVRAGE : {tache}   |   PHASE : {phase}"

        marge_x = max(int(w * 0.03), 20)
        draw.text((marge_x, h - band_h + int(band_h * 0.18)), ligne_1, fill=(255, 255, 255, 255), font=font_t)
        draw.text((marge_x, h - band_h + int(band_h * 0.55)), ligne_2, fill=(56, 189, 248, 255), font=font_s)

        img_finale = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
        out_bytes = io.BytesIO()
        img_finale.save(out_bytes, format="JPEG", quality=90)
        return out_bytes.getvalue()
    except Exception:
        return image_file.getbuffer()

def generer_rapport_excel(df_source):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    c_navy = "1E3A8A"
    c_header = "0F766E"
    c_row_alt = "F8FAFC"
    c_total = "E2E8F0"
    
    font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=10, color="0F172A")
    font_total = Font(name="Calibri", size=11, bold=True, color="1E3A8A")

    fill_title = PatternFill(start_color=c_navy, end_color=c_navy, fill_type="solid")
    fill_header = PatternFill(start_color=c_header, end_color=c_header, fill_type="solid")
    fill_alt = PatternFill(start_color=c_row_alt, end_color=c_row_alt, fill_type="solid")
    fill_tot = PatternFill(start_color=c_total, end_color=c_total, fill_type="solid")

    border_thin = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    border_double = Border(
        top=Side(style='thin', color='94A3B8'),
        bottom=Side(style='double', color='0F172A')
    )

    # 1. FEUILLE DE SYNTHÈSE GÉNÉRALE
    ws_main = wb.create_sheet(title="Synthèse Générale")
    ws_main.merge_cells("A1:E1")
    ws_main["A1"] = "TABLEAU DE BORD GLOBAL - SUIVI DES CHANTIERS D'ÉTANCHÉITÉ"
    ws_main["A1"].font = font_title
    ws_main["A1"].fill = fill_title
    ws_main["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_main.row_dimensions[1].height = 36

    headers_syn = ["Chantier / Projet", "Surface Totale (m²)", "Linéaire Total (ML)", "Unités (U)", "Interventions"]
    ws_main.append([])
    ws_main.append(headers_syn)
    ws_main.row_dimensions[3].height = 26

    for col_i in range(1, len(headers_syn) + 1):
        cell = ws_main.cell(row=3, column=col_i)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")

    chantiers_uniques = [c for c in df_source["Chantier"].dropna().unique() if not str(c).startswith("2026-") and str(c).strip()]
    
    tot_m2_glob = 0.0
    tot_ml_glob = 0.0
    tot_u_glob = 0.0
    current_r = 4

    for ch in sorted(chantiers_uniques):
        sub = df_source[df_source["Chantier"] == ch].copy()
        sub["Rendement_num"] = pd.to_numeric(sub.get("Rendement", 0), errors='coerce').fillna(0)
        unite_col = sub.get("Unite", pd.Series([""] * len(sub))).astype(str)

        m2 = float(sub[unite_col.str.contains("m²", na=False)]["Rendement_num"].sum())
        ml = float(sub[unite_col.str.contains("ML", na=False)]["Rendement_num"].sum())
        u = float(sub[unite_col.str.contains("U", na=False)]["Rendement_num"].sum())

        tot_m2_glob += m2
        tot_ml_glob += ml
        tot_u_glob += u

        ws_main.append([ch, round(m2, 2), round(ml, 2), round(u, 2), len(sub)])
        
        for c_idx in range(1, 6):
            c_cell = ws_main.cell(row=current_r, column=c_idx)
            c_cell.font = font_data
            c_cell.border = border_thin
            if current_r % 2 == 1:
                c_cell.fill = fill_alt
            if c_idx >= 2:
                c_cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                c_cell.alignment = Alignment(horizontal="left", vertical="center")
        current_r += 1

    ws_main.append(["TOTAL GÉNÉRAL", round(tot_m2_glob, 2), round(tot_ml_glob, 2), round(tot_u_glob, 2), len(df_source)])
    for c_idx in range(1, 6):
        c_cell = ws_main.cell(row=current_r, column=c_idx)
        c_cell.font = font_total
        c_cell.fill = fill_tot
        c_cell.border = border_double
        if c_idx >= 2:
            c_cell.alignment = Alignment(horizontal="right", vertical="center")

    for col in ws_main.columns:
        max_l = max(len(str(c.value or '')) for c in col)
        col_letter = get_column_letter(col[0].column)
        ws_main.column_dimensions[col_letter].width = max(max_l + 6, 16)

    # 2. FEUILLE PAR CHANTIER
    for ch in sorted(chantiers_uniques):
        sub_ch = df_source[df_source["Chantier"] == ch].copy()
        clean_titre = re.sub(r'[^a-zA-Z0-9_-]', '_', ch.split('(')[0].strip())[:28]
        ws_ch = wb.create_sheet(title=clean_titre)

        ws_ch.merge_cells("A1:H1")
        ws_ch["A1"] = f"REGISTRE D'ATTACHEMENT TECHNIQUE — CHANTIER : {ch.upper()}"
        ws_ch["A1"].font = font_title
        ws_ch["A1"].fill = fill_title
        ws_ch["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws_ch.row_dimensions[1].height = 36

        headers_ch = [
            "Date", "Corps d'État / Ouvrage", "Phase Réalisation", 
            "Rendement", "Unité", "Matériaux Consommés", 
            "Équipe Présente", "Observations / Notes"
        ]
        ws_ch.append([])
        ws_ch.append(headers_ch)
        ws_ch.row_dimensions[3].height = 26

        for col_i in range(1, len(headers_ch) + 1):
            cell = ws_ch.cell(row=3, column=col_i)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center")

        row_ch_idx = 4
        m2_chantier = 0.0
        ml_chantier = 0.0

        for _, r in sub_ch.iterrows():
            c_mat = str(r.get("Consommation", "")).strip()
            if not c_mat or c_mat == "nan":
                c_mat = "Aucun"

            rend_val = pd.to_numeric(r.get("Rendement", 0), errors='coerce') or 0.0
            unite_val = str(r.get("Unite", ""))
            if "m²" in unite_val:
                m2_chantier += rend_val
            elif "ML" in unite_val:
                ml_chantier += rend_val

            ws_ch.append([
                r.get("Date", ""),
                r.get("Corps_d_etat", ""),
                r.get("Phase", ""),
                rend_val,
                unite_val,
                c_mat,
                r.get("Effectif", ""),
                r.get("Legende", "")
            ])

            for c_i in range(1, 9):
                cell = ws_ch.cell(row=row_ch_idx, column=c_i)
                cell.font = font_data
                cell.border = border_thin
                if row_ch_idx % 2 == 1:
                    cell.fill = fill_alt
                if c_i == 4:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif c_i in [1, 5]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

            row_ch_idx += 1

        ws_ch.append([
            "TOTAL RÉALISÉ", "", "", 
            f"{round(m2_chantier, 2)} m² | {round(ml_chantier, 2)} ML", 
            "", "", f"{len(sub_ch)} rapports", ""
        ])
        for c_i in range(1, 9):
            cell = ws_ch.cell(row=row_ch_idx, column=c_i)
            cell.font = font_total
            cell.fill = fill_tot
            cell.border = border_double
            if c_i == 4:
                cell.alignment = Alignment(horizontal="right", vertical="center")

        for col in ws_ch.columns:
            max_l = max(len(str(c.value or '')) for c in col)
            col_letter = get_column_letter(col[0].column)
            ws_ch.column_dimensions[col_letter].width = max(max_l + 5, 14)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

config = charger_config()

if "liste_consommations" not in st.session_state:
    st.session_state.liste_consommations = []

# ONGLETS VISIBLES ET BIEN DÉFINIS
tab_saisie, tab_admin = st.tabs(["📲 Saisie Terrain", "📊 Supervision & Rapports"])

# -------------------------------------------------------------
# ONGLET 1 : SAISIE TERRAIN ULTRA ERGONOMIQUE
# -------------------------------------------------------------
with tab_saisie:
    st.markdown("""
        <div class='header-cadre'>
            <div style='display: flex; justify-content: space-between; align-items: center;'>
                <div>
                    <h1 class='header-title'>Rapport Journalier</h1>
                    <div class='header-sub'>Étanchéité technique & Traitement</div>
                </div>
                <div>
                    <span class='badge-pro'>PRO V1.0</span>
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    st.markdown("### 📍 Localisation & Tâche")
    # Date + Champs en disposition verticale propre
    date_jour = st.date_input("Date des travaux", value=date.today())
    chantier_sel = st.selectbox("Projet / Chantier", config["chantiers"])
    tache_sel = st.selectbox("Corps d'état / Ouvrage", config["taches"])
    phase_travaux = st.selectbox("Phase de réalisation", [
        "Pendant exécution / application",
        "Avant travaux (État du support)",
        "Après achèvement (Finition)",
        "Détail technique / Gorge / Relevé",
        "Épreuve d'eau (Test d'étanchéité)",
        "Autre"
    ])

    st.write("---")
    st.markdown("### 👷 Équipe & Metré")
    macons_presents = st.multiselect("Compagnons présents", config["macons"], placeholder="Sélectionnez l'équipe...")

    c_r1, c_r2 = st.columns([2, 1])
    with c_r1:
        rendement = st.number_input("Métré réalisé", min_value=0.0, step=1.0, format="%.2f")
    with c_r2:
        unite = st.selectbox("Unité", ["m²", "ML", "U"])

    st.write("---")
    st.markdown("### 🧪 Matériaux Consommés")

    if st.session_state.liste_consommations:
        for idx_c, item in enumerate(st.session_state.liste_consommations):
            col_txt, col_sup = st.columns([3, 1])
            with col_txt:
                st.markdown(f"<span class='tag-materiau'>📦 {item['produit']} : <b>{item['quantite']} {item['unite']}</b></span>", unsafe_allow_html=True)
            with col_sup:
                st.markdown('<div class="btn-supprimer">', unsafe_allow_html=True)
                if st.button("Suppr.", key=f"del_{idx_c}"):
                    st.session_state.liste_consommations.pop(idx_c)
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

    with st.expander("➕ Enregistrer un matériau appliqué", expanded=True):
        nouveau_mat = st.selectbox("Produit appliqué", config["materiaux"], key="ajout_mat")
        cp2, cp3 = st.columns([1, 1])
        with cp2:
            nouvelle_qte = st.number_input("Quantité", min_value=0.0, step=1.0, format="%.2f", key="ajout_qte")
        with cp3:
            nouvelle_uni = st.selectbox("Unité", ["Seaux/Bidons", "Sacs", "Rouleaux", "Kg", "Litres", "Cartouches", "U"], key="ajout_uni")

        if st.button("➕ Ajouter au registre", use_container_width=True):
            if nouvelle_qte > 0:
                st.session_state.liste_consommations.append({
                    "produit": nouveau_mat,
                    "quantite": nouvelle_qte,
                    "unite": nouvelle_uni
                })
                st.rerun()
            else:
                st.warning("⚠️ Précisez une quantité > 0.")

    st.write("---")
    st.markdown("### 📸 Photos Justificatives")
    photos_galerie = st.file_uploader(
        "Photos de l'ouvrage (Horodatage automatique)",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True
    )
    legende = st.text_input("Observations techniques", placeholder="Ex : Support sec, dépoussiéré avant primaire...")

    st.write("")
    st.markdown('<div class="btn-valider">', unsafe_allow_html=True)
    envoyer_btn = st.button("🚀 TRANSMETTRE LE RAPPORT", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if envoyer_btn:
        if not macons_presents:
            st.error("⚠️ Sélectionnez au moins un ouvrier.")
        elif not photos_galerie:
            st.error("⚠️ Ajoutez au moins une photo justificative.")
        else:
            try:
                dossier_chantier = clean_folder_name(chantier_sel)
                dossier_date = str(date_jour)

                chemin_cible = os.path.join(PHOTOS_BASE_DIR, dossier_chantier, dossier_date)
                os.makedirs(chemin_cible, exist_ok=True)

                saved_files = []
                tache_clean = re.sub(r'[^a-zA-Z0-9_-]', '_', tache_sel)[:12]

                for idx, p in enumerate(photos_galerie):
                    ext = ".jpg"
                    nom_fichier = f"{tache_clean}_{idx+1}{ext}"
                    chemin_disque = os.path.join(chemin_cible, nom_fichier)

                    photo_bytes = appliquer_watermark(p, chantier_sel, tache_sel, phase_travaux, str(date_jour))
                    with open(chemin_disque, "wb") as f_img:
                        f_img.write(photo_bytes)

                    saved_files.append(f"{dossier_chantier}/{dossier_date}/{nom_fichier}")

                if st.session_state.liste_consommations:
                    conso_finale = " | ".join([f"{item['produit']}: {item['quantite']} {item['unite']}" for item in st.session_state.liste_consommations])
                else:
                    conso_finale = "Aucun"

                nouvelle_ligne = {
                    "Date": str(date_jour),
                    "Chantier": str(chantier_sel),
                    "Corps_d_etat": str(tache_sel),
                    "Phase": str(phase_travaux),
                    "Rendement": str(rendement),
                    "Unite": str(unite),
                    "Consommation": conso_finale,
                    "Effectif": ", ".join(macons_presents),
                    "Nb_Ouvriers": len(macons_presents),
                    "Photos": "|".join(saved_files),
                    "Nb_Photos": len(saved_files),
                    "Legende": str(legende).replace(";", " ").replace("|", " ")
                }

                df_entry = pd.DataFrame([nouvelle_ligne])[COLONNES_OFFICIELLES]

                if not os.path.exists(CSV_FILE):
                    df_entry.to_csv(CSV_FILE, sep=';', index=False, encoding='utf-8-sig')
                else:
                    df_entry.to_csv(CSV_FILE, sep=';', mode='a', header=False, index=False, encoding='utf-8-sig')

                st.session_state.liste_consommations = []
                st.balloons()
                st.success(f"✅ Rapport validé avec succès !")
            except Exception as e:
                st.error(f"❌ Erreur : {e}")

# -------------------------------------------------------------
# ONGLET 2 : SUPERVISION & RAPPORTS
# -------------------------------------------------------------
with tab_admin:
    st.markdown("""
        <div class='header-cadre'>
            <div style='display: flex; justify-content: space-between; align-items: center;'>
                <div>
                    <h1 class='header-title'>Tableau de Bord</h1>
                    <div class='header-sub'>Supervision technique & Rapports</div>
                </div>
                <div>
                    <span class='badge-pro'>ADMIN</span>
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    pin = st.text_input("Authentification (Code PIN) :", type="password", placeholder="Code PIN...")

    if pin == ADMIN_PIN:
        with st.expander("⚙️ Gestion des Listes (Ajouter / Supprimer)", expanded=False):
            # CHANTIERS
            st.markdown("<p style='color: #38BDF8 !important; font-weight: 800;'>🏢 Chantiers / Projets</p>", unsafe_allow_html=True)
            nouveau_chantier = st.text_input("Nouveau chantier", placeholder="Ex : PROJET-CENTRE...", key="add_ch")
            if st.button("➕ Ajouter ce chantier", use_container_width=True):
                val = nouveau_chantier.strip()
                if val and val not in config["chantiers"]:
                    config["chantiers"].append(val)
                    sauvegarder_config(config)
                    st.success(f"Chantier « {val} » ajouté !")
                    st.rerun()

            del_ch = st.selectbox("Supprimer un chantier", ["-- Sélectionner --"] + config["chantiers"], key="del_ch_box")
            st.markdown('<div class="btn-supprimer">', unsafe_allow_html=True)
            if st.button("🗑️ Supprimer ce chantier", use_container_width=True):
                if del_ch != "-- Sélectionner --":
                    config["chantiers"].remove(del_ch)
                    sauvegarder_config(config)
                    st.success(f"Chantier supprimé !")
                    st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

            # TÂCHES
            st.write("---")
            st.markdown("<p style='color: #38BDF8 !important; font-weight: 800;'>🛠️ Corps d'État / Tâches</p>", unsafe_allow_html=True)
            nouvelle_tache = st.text_input("Nouvelle tâche", placeholder="Ex : ÉTANCHÉITÉ...", key="add_tache")
            if st.button("➕ Ajouter cette tâche", use_container_width=True):
                val = nouvelle_tache.strip()
                if val and val not in config["taches"]:
                    config["taches"].append(val)
                    sauvegarder_config(config)
                    st.success(f"Tâche ajoutée !")
                    st.rerun()

            del_tache = st.selectbox("Supprimer une tâche", ["-- Sélectionner --"] + config["taches"], key="del_tache_box")
            st.markdown('<div class="btn-supprimer">', unsafe_allow_html=True)
            if st.button("🗑️ Supprimer cette tâche", use_container_width=True):
                if del_tache != "-- Sélectionner --":
                    config["taches"].remove(del_tache)
                    sauvegarder_config(config)
                    st.success(f"Tâche supprimée !")
                    st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

            # PRODUITS
            st.write("---")
            st.markdown("<p style='color: #38BDF8 !important; font-weight: 800;'>🧪 Matériaux</p>", unsafe_allow_html=True)
            nouveau_mat = st.text_input("Nouveau produit", placeholder="Ex : RESINE...", key="add_mat")
            if st.button("➕ Ajouter ce produit", use_container_width=True):
                val = nouveau_mat.strip()
                if val and val not in config["materiaux"]:
                    config["materiaux"].insert(-1, val)
                    sauvegarder_config(config)
                    st.success(f"Produit ajouté !")
                    st.rerun()

            del_mat = st.selectbox("Supprimer un produit", ["-- Sélectionner --"] + config["materiaux"], key="del_mat_box")
            st.markdown('<div class="btn-supprimer">', unsafe_allow_html=True)
            if st.button("🗑️ Supprimer ce produit", use_container_width=True):
                if del_mat != "-- Sélectionner --":
                    config["materiaux"].remove(del_mat)
                    sauvegarder_config(config)
                    st.success(f"Produit supprimé !")
                    st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

            # COMPAGNONS
            st.write("---")
            st.markdown("<p style='color: #38BDF8 !important; font-weight: 800;'>👷 Compagnons / Équipe</p>", unsafe_allow_html=True)
            nouveau_macon = st.text_input("Nouvel ouvrier", placeholder="Ex : Nom & Prénom...", key="add_macon")
            if st.button("➕ Ajouter cet ouvrier", use_container_width=True):
                val = nouveau_macon.strip()
                if val and val not in config["macons"]:
                    config["macons"].append(val)
                    sauvegarder_config(config)
                    st.success(f"Ouvrier ajouté !")
                    st.rerun()

            del_macon = st.selectbox("Supprimer un ouvrier", ["-- Sélectionner --"] + config["macons"], key="del_macon_box")
            st.markdown('<div class="btn-supprimer">', unsafe_allow_html=True)
            if st.button("🗑️ Supprimer cet ouvrier", use_container_width=True):
                if del_macon != "-- Sélectionner --":
                    config["macons"].remove(del_macon)
                    sauvegarder_config(config)
                    st.success(f"Ouvrier supprimé !")
                    st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

        st.write("---")

        df_all = charger_donnees()

        total_photos = 0
        for root, _, files in os.walk(PHOTOS_BASE_DIR):
            total_photos += len([f for f in files if f.lower().endswith(('.jpg', '.jpeg', '.png'))])

        k1, k2, k3 = st.columns(3)
        with k1:
            st.metric("Rapports", f"{len(df_all) if df_all is not None else 0}")
        with k2:
            st.metric("Photos", f"{total_photos}")
        with k3:
            projets_actifs = len(df_all["Chantier"].dropna().unique()) if df_all is not None and not df_all.empty else 0
            st.metric("Projets", f"{projets_actifs}")

        st.write("---")
        st.markdown("### 📑 Téléchargement Rapport Excel")

        if df_all is not None and not df_all.empty:
            excel_bytes = generer_rapport_excel(df_all)
            st.download_button(
                label="📊 TÉLÉCHARGER LE CLASSEUR EXCEL (.XLSX)",
                data=excel_bytes,
                file_name=f"Synthese_Chantiers_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.button("Export Excel (En attente)", disabled=True, use_container_width=True)

        st.write("---")
        dossiers_chantiers = [d for d in os.listdir(PHOTOS_BASE_DIR) if os.path.isdir(os.path.join(PHOTOS_BASE_DIR, d))]
        if dossiers_chantiers:
            st.markdown("### 📁 Téléchargement Photos ZIP")
            for d_ch in sorted(dossiers_chantiers):
                dir_ch = os.path.join(PHOTOS_BASE_DIR, d_ch)
                fichiers_total = []
                for r, _, f_list in os.walk(dir_ch):
                    for f in f_list:
                        if f.lower().endswith(('.jpg', '.jpeg', '.png')):
                            fichiers_total.append(os.path.join(r, f))

                if fichiers_total:
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        for f_abs in fichiers_total:
                            rel_p = os.path.relpath(f_abs, PHOTOS_BASE_DIR)
                            zf.write(f_abs, arcname=rel_p)

                    st.markdown(f"**📁 {d_ch}** — `{len(fichiers_total)} photos`")
                    st.download_button(label=f"Télécharger ZIP ({d_ch})", data=zip_buf.getvalue(), file_name=f"Photos_{d_ch}.zip", mime="application/zip", key=f"z_{d_ch}", use_container_width=True)
                    st.write("")

        st.write("---")
        if df_all is not None and not df_all.empty and "Chantier" in df_all.columns:
            st.markdown("### 🔍 Registre des Interventions")
            chantiers_bruts = [c for c in df_all["Chantier"].dropna().unique() if not str(c).startswith("2026-") and str(c).strip()]
            f_proj = st.selectbox("Filtrer par projet :", ["Tous les projets"] + list(chantiers_bruts))

            df_show = df_all.copy()
            if f_proj != "Tous les projets":
                df_show = df_show[df_show["Chantier"] == f_proj]

            for orig_idx, row in df_show.iloc[::-1].iterrows():
                if str(row.get("Chantier", "")).startswith("2026-") or not str(row.get("Chantier", "")).strip():
                    continue

                conso_val = str(row.get("Consommation", "")).strip()
                if not conso_val or conso_val == "nan" or conso_val == "Aucun":
                    badge_conso_html = "<span style='color: #38BDF8 !important; font-size: 13px; font-weight: 700;'>📦 Aucune consommation déclarée</span>"
                else:
                    items_conso = conso_val.split(" | ")
                    badge_conso_html = "".join([f"<span class='tag-materiau'>🧪 {c}</span>" for c in items_conso])

                st.markdown(f"""
                    <div class='card-intervention'>
                        <div style='display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #334155; padding-bottom: 8px; margin-bottom: 10px;'>
                            <span class='tag-projet'>🏢 {row.get("Chantier", "")}</span>
                            <span class='tag-date'>📅 {row.get("Date", "")}</span>
                        </div>
                        <div style='margin-bottom: 8px;'>
                            <span class='tag-corps'>🛠️ {row.get("Corps_d_etat", "")}</span>
                            <span class='tag-rendement'>📏 {row.get("Rendement", "")} {row.get("Unite", "")}</span>
                            <div class='card-text-cyan' style='margin-top: 4px;'>Phase : {row.get("Phase", "")}</div>
                        </div>
                        <div style='margin: 8px 0;'>
                            {badge_conso_html}
                        </div>
                        <div class='card-text-white'>
                            👷 <span style='color: #38BDF8 !important; font-weight: 800;'>Équipe :</span> {row.get("Effectif", "")}
                        </div>
                        <div class='card-text-note'>
                            💬 <span style='color: #5EEAD4 !important; font-weight: 800;'>Note :</span> {row.get("Legende", "R.A.S")}
                        </div>
                    </div>
                """, unsafe_allow_html=True)

                raw_photos = str(row.get("Photos", "")).replace(";", "|").replace(",", "|").split("|")
                valid_paths = []
                for ph in raw_photos:
                    ph_clean = ph.strip()
                    p1 = os.path.join(PHOTOS_BASE_DIR, ph_clean)
                    if os.path.exists(p1):
                        valid_paths.append(p1)
                    else:
                        for r, _, fl in os.walk(PHOTOS_BASE_DIR):
                            if os.path.basename(ph_clean) in fl:
                                valid_paths.append(os.path.join(r, os.path.basename(ph_clean)))
                                break

                if valid_paths:
                    cols = st.columns(min(len(valid_paths), 2))
                    for i, p_img in enumerate(valid_paths):
                        try:
                            cols[i % 2].image(Image.open(p_img), use_container_width=True)
                        except Exception:
                            pass

                st.markdown('<div class="btn-supprimer" style="margin-top: 6px;">', unsafe_allow_html=True)
                if st.button(f"🗑️ Supprimer cette fiche", key=f"del_row_{orig_idx}"):
                    df_all = df_all.drop(orig_idx).reset_index(drop=True)
                    sauvegarder_donnees(df_all)
                    st.success("Fiche supprimée avec succès !")
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)
                st.write("")

        with st.expander("⚙️ Maintenance du système"):
            if st.button("🗑️ Réinitialiser tout le registre CSV"):
                if os.path.exists(CSV_FILE):
                    os.remove(CSV_FILE)
                    st.rerun()

    elif pin != "":
        st.error("❌ Code d'accès non autorisé.")
