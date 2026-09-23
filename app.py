import streamlit as st
import pandas as pd
from datetime import date, datetime
import os
import json
import io
import zipfile
import re
from PIL import Image
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Système de Suivi Technique Étanchéité",
    page_icon="🏗️",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# --- CSS COMPLET HAUTE VISIBILITÉ & PLEIN ÉCRAN MOBILE ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    /* 1. SUPPRESSION TOTALE DE LA BARRE DU HAUT ET MENUS STREAMLIT */
    #MainMenu {visibility: hidden !important; display: none !important;}
    header, [data-testid="stHeader"] {display: none !important;}
    footer {visibility: hidden !important; display: none !important;}
    .stDeployButton {display: none !important;}

    /* 2. SUPPRESSION DES BADGES ET BOUTONS STREAMLIT EN BAS */
    [data-testid="manage-app-button"],
    [data-testid="stStatusWidget"],
    .viewerBadge_container__1QSob,
    [class*="viewerBadge"],
    [class*="manageApp"],
    [class*="StatusWidget"],
    [class*="ProfileButton"],
    [class*="FloatingActionButton"],
    button[aria-label*="Manage"],
    a[href*="streamlit.io/cloud"],
    a[href*="share.streamlit.io"],
    div:has(> a[href*="streamlit.io"]),
    div:has(> [data-testid="manage-app-button"]) {
        display: none !important;
        visibility: hidden !important;
        opacity: 0 !important;
        pointer-events: none !important;
        width: 0 !important;
        height: 0 !important;
    }

    /* Fond général sombre */
    .stApp {
        background-color: #0B1120 !important;
        margin-top: -40px !important;
        padding-bottom: 30px !important;
    }

    /* TITRES SANS COUPURE */
    h1, h2, h3, h4, h5, h6, .stMarkdown h1, .stMarkdown h2, .stMarkdown h3, .stMarkdown h4, .stMarkdown h5 {
        color: #FFFFFF !important;
        font-weight: 800 !important;
        word-break: normal !important;
        overflow-wrap: break-word !important;
        hyphens: none !important;
    }

    /* EN-TÊTE SMARTPHONE PRO ÉPURÉ */
    .header-cadre {
        background-color: #0F172A;
        border-radius: 14px;
        padding: 16px 14px;
        box-shadow: 0 10px 24px -5px rgba(0, 0, 0, 0.6);
        margin-bottom: 18px;
        border: 1px solid #334155;
    }
    .header-title {
        font-size: 19px !important;
        line-height: 1.3 !important;
        font-weight: 900 !important;
        margin: 0 !important;
        color: #FFFFFF !important;
        word-break: normal !important;
        hyphens: none !important;
    }
    .header-sub {
        font-size: 12px !important;
        color: #38BDF8 !important;
        margin-top: 5px !important;
        font-weight: 600 !important;
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
        display: inline-block;
    }

    /* ÉCRAN DE VERROUILLAGE ADMIN */
    .lock-card {
        background: linear-gradient(145deg, #0F172A 0%, #1E293B 100%);
        border: 1px solid #334155;
        border-radius: 16px;
        padding: 26px 18px;
        text-align: center;
        box-shadow: 0 14px 30px -5px rgba(0, 0, 0, 0.6);
        margin: 15px auto 25px auto;
    }
    .lock-icon-badge {
        font-size: 34px;
        background: rgba(13, 148, 136, 0.2);
        border: 1px solid #0D9488;
        width: 65px;
        height: 65px;
        line-height: 65px;
        border-radius: 50%;
        margin: 0 auto 12px auto;
        display: block;
    }
    .lock-title {
        color: #FFFFFF !important;
        font-size: 19px;
        font-weight: 800;
        margin: 0 0 6px 0;
    }
    .lock-subtitle {
        color: #94A3B8 !important;
        font-size: 13px;
        line-height: 1.4;
        margin: 0 0 16px 0;
    }

    /* LABELS DES CHAMPS */
    .stWidgetLabel p, [data-testid="stWidgetLabel"] p, label p {
        color: #FFFFFF !important;
        font-size: 14px !important;
        font-weight: 700 !important;
    }

    /* CASES DE FORMULAIRE (FOND BLANC + TEXTE NOIR) */
    .stTextInput input, .stDateInput input, .stNumberInput input {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        font-size: 15px !important;
        font-weight: 700 !important;
        border: 2px solid #CBD5E1 !important;
        border-radius: 8px !important;
    }
    .stSelectbox div[data-baseweb="select"] {
        background-color: #FFFFFF !important;
        border: 2px solid #CBD5E1 !important;
        border-radius: 8px !important;
    }
    .stSelectbox div[data-baseweb="select"] * {
        color: #0F172A !important;
        font-weight: 700 !important;
    }
    .stMultiSelect div[data-baseweb="select"] {
        background-color: #FFFFFF !important;
        border: 2px solid #CBD5E1 !important;
        border-radius: 8px !important;
    }
    .stMultiSelect [data-baseweb="tag"] {
        background-color: #0F766E !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
    }

    div[data-baseweb="popover"], ul[role="listbox"], li[role="option"] {
        background-color: #FFFFFF !important;
    }
    div[data-baseweb="popover"] *, ul[role="listbox"] *, li[role="option"] * {
        color: #0F172A !important;
        font-weight: 700 !important;
    }
    li[role="option"]:hover {
        background-color: #E2E8F0 !important;
    }

    div[data-baseweb="calendar"], div[data-baseweb="calendar"] * {
        color: #0F172A !important;
        font-weight: 700 !important;
    }
    div[data-baseweb="calendar"] button[aria-selected="true"] {
        background-color: #0F766E !important;
        color: #FFFFFF !important;
    }

    /* CADRE PHOTO UPLOAD */
    [data-testid="stFileUploader"] {
        position: relative !important;
    }
    [data-testid="stFileUploader"] section {
        position: relative !important;
        background-color: #FFFFFF !important;
        border: 2px dashed #0284C7 !important;
        border-radius: 12px !important;
        padding: 18px !important;
        cursor: pointer !important;
    }
    [data-testid="stFileUploader"] section * {
        color: #0F172A !important;
        font-weight: 800 !important;
    }
    [data-testid="stFileUploader"] section button {
        background-color: #0284C7 !important;
        border: none !important;
        pointer-events: none !important;
    }
    [data-testid="stFileUploader"] section button * {
        color: #FFFFFF !important;
        font-weight: 800 !important;
    }

    [data-testid="stFileUploader"] input[type="file"] {
        display: block !important;
        position: absolute !important;
        top: 0 !important;
        left: 0 !important;
        width: 100% !important;
        height: 100% !important;
        opacity: 0 !important;
        z-index: 9999 !important;
        cursor: pointer !important;
    }

    /* CARTES D'APERÇU PHOTOS D'UPLOAD */
    .photo-preview-card {
        background: #1E293B;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 8px;
        text-align: center;
        margin-bottom: 12px;
    }

    [data-testid="stMetricValue"] {
        color: #FFFFFF !important;
        font-size: 32px !important;
        font-weight: 900 !important;
    }
    [data-testid="stMetricLabel"] p {
        color: #38BDF8 !important;
        font-size: 14px !important;
        font-weight: 800 !important;
    }

    div.stButton > button {
        background-color: #0284C7 !important;
        color: #FFFFFF !important;
        font-size: 15px !important;
        font-weight: 800 !important;
        border-radius: 10px !important;
        border: none !important;
        padding: 12px 16px !important;
        box-shadow: 0 4px 10px rgba(2, 132, 199, 0.3) !important;
    }
    div.stButton > button * {
        color: #FFFFFF !important;
        font-weight: 800 !important;
    }

    .btn-supprimer button {
        background-color: #DC2626 !important;
        color: #FFFFFF !important;
        padding: 8px 12px !important;
        font-size: 13px !important;
        font-weight: 800 !important;
    }

    .btn-enregistrer-modif button {
        background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%) !important;
        color: #FFFFFF !important;
        font-size: 15px !important;
        font-weight: 800 !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 12px !important;
        width: 100% !important;
    }

    .btn-valider button {
        background: linear-gradient(135deg, #059669 0%, #10B981 100%) !important;
        color: #FFFFFF !important;
        font-size: 16px !important;
        font-weight: 800 !important;
        border-radius: 12px !important;
        border: none !important;
        padding: 16px !important;
        width: 100% !important;
        box-shadow: 0 6px 20px rgba(16, 185, 129, 0.4) !important;
    }

    div.stDownloadButton > button {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        font-size: 14px !important;
        font-weight: 800 !important;
        border-radius: 10px !important;
        border: 2px solid #CBD5E1 !important;
        padding: 10px 16px !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.25) !important;
    }
    div.stDownloadButton > button * {
        color: #0F172A !important;
        font-weight: 800 !important;
    }

    .card-intervention {
        background-color: #111827;
        border-radius: 14px;
        border: 1px solid #334155;
        padding: 16px;
        margin-bottom: 16px;
        border-left: 5px solid #10B981;
    }
    .tag-projet {
        font-size: 15px;
        font-weight: 900;
        color: #FFFFFF !important;
    }
    .tag-date {
        font-size: 12px;
        font-weight: 700;
        color: #FFFFFF !important;
        background: #1E293B;
        padding: 3px 8px;
        border-radius: 6px;
        border: 1px solid #475569;
    }
    .tag-corps {
        display: inline-block;
        background-color: #1E293B;
        color: #FFFFFF !important;
        border: 1px solid #475569;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 13px;
        font-weight: 800;
    }
    .tag-rendement {
        display: inline-block;
        background-color: #064E3B;
        color: #6EE7B7 !important;
        border: 1px solid #059669;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 13px;
        font-weight: 900;
        margin-left: 6px;
    }
    .tag-materiau {
        display: inline-block;
        background-color: #082F49;
        color: #7DD3FC !important;
        border: 1px solid #0284C7;
        padding: 4px 8px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: 800;
        margin: 3px 4px 3px 0;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 44px;
        background-color: #1E293B;
        border-radius: 8px 8px 0px 0px;
        color: #FFFFFF !important;
        font-weight: 800;
        font-size: 13px;
        border: 1px solid #334155;
        padding: 6px 12px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #0F172A !important;
        color: #38BDF8 !important;
        border-top: 3px solid #38BDF8 !important;
    }
    hr {
        border-color: #334155 !important;
        margin: 18px 0 !important;
    }
    </style>

    <script>
    function purgeStreamlitBadges() {
        const sel = '[data-testid="manage-app-button"], [data-testid="stStatusWidget"], [class*="viewerBadge"], [class*="manageApp"], a[href*="streamlit.io"]';
        const elements = window.parent.document.querySelectorAll(sel);
        elements.forEach(el => el.remove());
    }
    setInterval(purgeStreamlitBadges, 600);

    function applyGalleryFix() {
        const doc = window.parent.document;
        const inputs = doc.querySelectorAll('input[type="file"]');
        inputs.forEach(inp => {
            inp.setAttribute('accept', 'image/*');
        });
    }
    setInterval(applyGalleryFix, 400);
    </script>
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

def generer_rapport_excel(df_source):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    border_thin = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    def styliser_feuille(ws, titre_entete):
        for col in range(1, len(titre_entete) + 1):
            cell = ws.cell(row=1, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(titre_entete)):
            for cell in row:
                cell.font = data_font
                cell.border = border_thin
                if isinstance(cell.value, (int, float)):
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

        ws.row_dimensions[1].height = 28
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 5, 13)

    # 1. Synthèse
    ws1 = wb.create_sheet(title="Synthèse Chantiers")
    headers1 = ["Chantier", "Surface Réalisée (m²)", "Linéaire Réalisé (ML)", "Unités (U)", "Total Interventions"]
    ws1.append(headers1)
    chantiers_uniques = [c for c in df_source["Chantier"].dropna().unique() if not str(c).startswith("2026-") and str(c).strip()]
    for ch in sorted(chantiers_uniques):
        sub = df_source[df_source["Chantier"] == ch].copy()
        sub["Rendement_num"] = pd.to_numeric(sub.get("Rendement", 0), errors='coerce').fillna(0)
        unite_col = sub.get("Unite", pd.Series([""] * len(sub))).astype(str)
        m2 = sub[unite_col.str.contains("m²", na=False)]["Rendement_num"].sum()
        ml = sub[unite_col.str.contains("ML", na=False)]["Rendement_num"].sum()
        u = sub[unite_col.str.contains("U", na=False)]["Rendement_num"].sum()
        ws1.append([ch, round(m2, 2), round(ml, 2), round(u, 2), len(sub)])
    styliser_feuille(ws1, headers1)

    # 2. Consommation
    ws2 = wb.create_sheet(title="Consommation Matériaux")
    headers2 = ["Date", "Chantier", "Corps d'État", "Matériaux Consommés", "Remarques"]
    ws2.append(headers2)
    for _, r in df_source.iterrows():
        c_mat = str(r.get("Consommation", "")).strip()
        if not c_mat or c_mat == "nan":
            c_mat = "Aucun"
        ws2.append([r.get("Date", ""), r.get("Chantier", ""), r.get("Corps_d_etat", ""), c_mat, r.get("Legende", "")])
    styliser_feuille(ws2, headers2)

    # 3. Effectif
    ws3 = wb.create_sheet(title="Pointage Ouvriers")
    headers3 = ["Date", "Chantier", "Corps d'État", "Effectif Présent", "Nombre d'Ouvriers"]
    ws3.append(headers3)
    for _, r in df_source.iterrows():
        ws3.append([r.get("Date", ""), r.get("Chantier", ""), r.get("Corps_d_etat", ""), r.get("Effectif", ""), r.get("Nb_Ouvriers", "")])
    styliser_feuille(ws3, headers3)

    # 4. Journal Détaillé
    ws4 = wb.create_sheet(title="Journal Détaillé")
    headers4 = ["Date", "Chantier", "Corps d'État", "Phase", "Rendement", "Unité", "Consommation", "Effectif", "Observation"]
    ws4.append(headers4)
    for _, r in df_source.iterrows():
        c_mat = str(r.get("Consommation", "")).strip()
        if not c_mat or c_mat == "nan":
            c_mat = "-"
        ws4.append([
            r.get("Date", ""),
            r.get("Chantier", ""),
            r.get("Corps_d_etat", ""),
            r.get("Phase", ""),
            pd.to_numeric(r.get("Rendement", 0), errors='coerce') or 0,
            r.get("Unite", ""),
            c_mat,
            r.get("Effectif", ""),
            r.get("Legende", "")
        ])
    styliser_feuille(ws4, headers4)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

config = charger_config()

if "liste_consommations" not in st.session_state:
    st.session_state.liste_consommations = []

tab_saisie, tab_admin = st.tabs(["📲 Saisie Terrain", "📊 Espace Encadrement & Rapports"])

# -------------------------------------------------------------
# ONGLET 1 : SAISIE TERRAIN (CALIBRÉE SMARTPHONE & PLEIN ÉCRAN)
# -------------------------------------------------------------
with tab_saisie:
    st.markdown("""
        <div class='header-cadre'>
            <div style='display: flex; justify-content: space-between; align-items: center;'>
                <div>
                    <h1 class='header-title'>Rapport Journalier d'Exécution</h1>
                    <div class='header-sub'>Étanchéité technique & Traitement des supports</div>
                </div>
                <div>
                    <span class='badge-pro'>PRO V1.0</span>
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    st.markdown("### 📍 Localisation & Tâche")
    c_p1, c_p2 = st.columns([1, 1])
    with c_p1:
        date_jour = st.date_input("Date des travaux", value=date.today())
        chantier_sel = st.selectbox("Projet / Chantier", config["chantiers"])
    with c_p2:
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
    st.markdown("### 👷 Équipe mobilisée & Metré")
    macons_presents = st.multiselect("Personnel d'exécution présent", config["macons"], placeholder="Sélectionnez les compagnons...")

    c_r1, c_r2 = st.columns([2, 1])
    with c_r1:
        rendement = st.number_input("Métré / Rendement réalisé", min_value=0.0, step=1.0, format="%.2f")
    with c_r2:
        unite = st.selectbox("Unité de mesure", ["m²", "ML", "U"])

    st.write("---")
    st.markdown("### 🧪 Matériaux & Produits Appliqués")

    if st.session_state.liste_consommations:
        for idx_c, item in enumerate(st.session_state.liste_consommations):
            col_txt, col_sup = st.columns([4, 1])
            with col_txt:
                st.markdown(f"<span class='tag-materiau'>📦 {item['produit']} : <b>{item['quantite']} {item['unite']}</b></span>", unsafe_allow_html=True)
            with col_sup:
                st.markdown('<div class="btn-supprimer">', unsafe_allow_html=True)
                if st.button("Supprimer", key=f"del_{idx_c}"):
                    st.session_state.liste_consommations.pop(idx_c)
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

    with st.expander("➕ Enregistrer un matériau consommé", expanded=True):
        cp1, cp2, cp3 = st.columns([2, 1, 1])
        with cp1:
            nouveau_mat = st.selectbox("Produit appliqué", config["materiaux"], key="ajout_mat")
        with cp2:
            nouvelle_qte = st.number_input("Quantité", min_value=0.0, step=1.0, format="%.2f", key="ajout_qte")
        with cp3:
            nouvelle_uni = st.selectbox("Conditionnement", ["Seaux/Bidons", "Sacs", "Rouleaux", "Kg", "Litres", "Cartouches", "U"], key="ajout_uni")

        if st.button("➕ Ajouter le produit au registre", use_container_width=True):
            if nouvelle_qte > 0:
                st.session_state.liste_consommations.append({
                    "produit": nouveau_mat,
                    "quantite": nouvelle_qte,
                    "unite": nouvelle_uni
                })
                st.rerun()
            else:
                st.warning("⚠️ Précisez une quantité supérieure à 0.")

    st.write("---")
    st.markdown("### 📸 Pièces Jointes & Justificatifs")
    
    # GESTION PROPRE DES PHOTOS TÉLÉVERSÉES DANS SESSION_STATE
    if "uploaded_photos_list" not in st.session_state:
        st.session_state.uploaded_photos_list = []

    st.markdown("<p style='color:#FFFFFF; font-weight:700; margin-bottom:6px;'>Sélectionner les photos de l'ouvrage :</p>", unsafe_allow_html=True)
    nouvelles_photos = st.file_uploader(
        "Toucher ici pour importer vos photos",
        accept_multiple_files=True,
        key="uploader_input"
    )

    # Synchronisation des photos sélectionnées
    if nouvelles_photos:
        for p in nouvelles_photos:
            deja_present = False
            for existante in st.session_state.uploaded_photos_list:
                if existante["name"] == p.name and existante["size"] == p.size:
                    deja_present = True
                    break
            if not deja_present:
                st.session_state.uploaded_photos_list.append({
                    "name": p.name,
                    "size": p.size,
                    "data": p.getbuffer()
                })

    # APERÇU DIRECT DES PHOTOS AVEC OPTION DE SUPPRIMER
    if st.session_state.uploaded_photos_list:
        st.markdown(f"""
            <div style='background: rgba(16, 185, 129, 0.15); border: 1px solid #10B981; border-radius: 8px; padding: 8px 12px; margin: 10px 0;'>
                <span style='color: #6EE7B7; font-weight: 800; font-size: 14px;'>✅ {len(st.session_state.uploaded_photos_list)} photo(s) prête(s) pour l'envoi</span>
            </div>
        """, unsafe_allow_html=True)

        cols_prev = st.columns(min(len(st.session_state.uploaded_photos_list), 3))
        for p_idx, p_obj in enumerate(st.session_state.uploaded_photos_list):
            c_target = cols_prev[p_idx % 3]
            c_target.image(p_obj["data"], use_container_width=True)
            st.markdown('<div class="btn-supprimer">', unsafe_allow_html=True)
            if c_target.button("🗑️ Enlever", key=f"del_up_photo_{p_idx}", use_container_width=True):
                st.session_state.uploaded_photos_list.pop(p_idx)
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

    st.write("---")
    legende = st.text_input("Observations techniques particulières", placeholder="Ex : Support brossé et dépoussiéré avant couche primaire...")

    st.write("")
    st.markdown('<div class="btn-valider">', unsafe_allow_html=True)
    envoyer_btn = st.button("TRANSMETTRE LE RAPPORT JOURNALIER", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if envoyer_btn:
        if not macons_presents:
            st.error("⚠️ Veuillez déclarer au moins un ouvrier pour cet ouvrage.")
        elif not st.session_state.uploaded_photos_list:
            st.error("⚠️ La conformité technique impose l'ajout d'au moins une photo justificative.")
        else:
            try:
                dossier_chantier = clean_folder_name(chantier_sel)
                dossier_date = str(date_jour)

                chemin_cible = os.path.join(PHOTOS_BASE_DIR, dossier_chantier, dossier_date)
                os.makedirs(chemin_cible, exist_ok=True)

                saved_files = []
                tache_clean = re.sub(r'[^a-zA-Z0-9_-]', '_', tache_sel)[:12]

                for idx, p_item in enumerate(st.session_state.uploaded_photos_list):
                    ext = ".jpg"
                    if "." in p_item["name"]:
                        ext = os.path.splitext(p_item["name"])[1].lower()

                    nom_fichier = f"{tache_clean}_{idx+1}{ext}"
                    chemin_disque = os.path.join(chemin_cible, nom_fichier)

                    with open(chemin_disque, "wb") as f_img:
                        f_img.write(p_item["data"])

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
                st.session_state.uploaded_photos_list = []
                st.balloons()
                st.success(f"✅ Rapport validé et intégré au registre officiel : {dossier_chantier} [{dossier_date}]")
            except Exception as e:
                st.error(f"❌ Erreur de transmission : {e}")

# -------------------------------------------------------------
# ONGLET 2 : ESPACE CADRE, DIRECTION & ÉCRAN DE SÉCURITÉ
# -------------------------------------------------------------
with tab_admin:
    if "admin_logged_in" not in st.session_state:
        st.session_state.admin_logged_in = False

    if not st.session_state.admin_logged_in:
        st.markdown("""
            <div class='lock-card'>
                <div class='lock-icon-badge'>🛡️</div>
                <h2 class='lock-title'>Authentification Supervision</h2>
                <p class='lock-subtitle'>Espace réservé à l'encadrement technique et à la direction des travaux.<br>Saisissez votre code PIN pour déverrouiller l'accès.</p>
            </div>
        """, unsafe_allow_html=True)

        c_lock1, c_lock2, c_lock3 = st.columns([1, 2, 1])
        with c_lock2:
            code_saisi = st.text_input("Code d'accès secret :", type="password", placeholder="Entrez le code PIN...", key="input_pin_auth")
            if st.button("🔓 DÉVERROUILLER L'ESPACE SUPERVISION", use_container_width=True):
                if code_saisi == ADMIN_PIN:
                    st.session_state.admin_logged_in = True
                    st.rerun()
                else:
                    st.error("❌ Code PIN incorrect. Accès refusé.")

    else:
        st.markdown("""
            <div class='header-cadre'>
                <div style='display: flex; justify-content: space-between; align-items: center;'>
                    <div>
                        <h1 class='header-title'>Tableau de Bord & Attachements</h1>
                        <div class='header-sub'>Supervision technique, synthèse des consommations et exports</div>
                    </div>
                    <div>
                        <span class='badge-pro'>SUPERVISION ACTIVE</span>
                    </div>
                </div>
            </div>
        """, unsafe_allow_html=True)

        c_dec1, c_dec2 = st.columns([3, 1])
        with c_dec2:
            st.markdown('<div class="btn-supprimer">', unsafe_allow_html=True)
            if st.button("🔒 Verrouiller la session", use_container_width=True):
                st.session_state.admin_logged_in = False
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

        # GESTION DES LISTES (AJOUT / SUPPRESSION)
        with st.expander("⚙️ Configuration des Listes (Ajouter / Supprimer des options)", expanded=False):
            st.markdown("##### 🛠️ Gestion des Chantiers, Tâches, Matériaux et Compagnons")
            
            # 1. Chantiers
            st.write("---")
            st.markdown("**🏢 1. Chantiers / Projets**")
            c_ch1, c_ch2 = st.columns(2)
            with c_ch1:
                nouveau_chantier = st.text_input("Nouveau chantier à ajouter", placeholder="Ex : NOUVEAU-PROJET...", key="add_ch")
                if st.button("➕ Ajouter ce chantier", use_container_width=True):
                    val = nouveau_chantier.strip()
                    if val and val not in config["chantiers"]:
                        config["chantiers"].append(val)
                        sauvegarder_config(config)
                        st.success(f"Chantier « {val} » ajouté !")
                        st.rerun()
            with c_ch2:
                del_ch = st.selectbox("Sélectionner un chantier à supprimer", ["-- Sélectionner --"] + config["chantiers"], key="del_ch_box")
                st.markdown('<div class="btn-supprimer">', unsafe_allow_html=True)
                if st.button("🗑️ Supprimer le chantier sélectionné", use_container_width=True):
                    if del_ch != "-- Sélectionner --":
                        config["chantiers"].remove(del_ch)
                        sauvegarder_config(config)
                        st.success(f"Chantier « {del_ch} » retiré de la liste !")
                        st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

            # 2. Corps d'état
            st.write("---")
            st.markdown("**🛠️ 2. Corps d'État / Tâches**")
            c_t1, c_t2 = st.columns(2)
            with c_t1:
                nouvelle_tache = st.text_input("Nouvelle tâche à ajouter", placeholder="Ex : ÉTANCHÉITÉ CUVELAGE...", key="add_tache")
                if st.button("➕ Ajouter cette tâche", use_container_width=True):
                    val = nouvelle_tache.strip()
                    if val and val not in config["taches"]:
                        config["taches"].append(val)
                        sauvegarder_config(config)
                        st.success(f"Tâche « {val} » ajoutée !")
                        st.rerun()
            with c_t2:
                del_tache = st.selectbox("Sélectionner une tâche à supprimer", ["-- Sélectionner --"] + config["taches"], key="del_tache_box")
                st.markdown('<div class="btn-supprimer">', unsafe_allow_html=True)
                if st.button("🗑️ Supprimer la tâche sélectionnée", use_container_width=True):
                    if del_tache != "-- Sélectionner --":
                        config["taches"].remove(del_tache)
                        sauvegarder_config(config)
                        st.success(f"Tâche « {del_tache} » retirée de la liste !")
                        st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

            # 3. Matériaux
            st.write("---")
            st.markdown("**🧪 3. Matériaux & Produits Consommés**")
            c_m1, c_m2 = st.columns(2)
            with c_m1:
                nouveau_mat = st.text_input("Nouveau matériau à ajouter", placeholder="Ex : RESINE POLYURETHANE...", key="add_mat")
                if st.button("➕ Ajouter ce produit", use_container_width=True):
                    val = nouveau_mat.strip()
                    if val and val not in config["materiaux"]:
                        config["materiaux"].insert(-1, val)
                        sauvegarder_config(config)
                        st.success(f"Produit « {val} » ajouté !")
                        st.rerun()
            with c_m2:
                del_mat = st.selectbox("Sélectionner un produit à supprimer", ["-- Sélectionner --"] + config["materiaux"], key="del_mat_box")
                st.markdown('<div class="btn-supprimer">', unsafe_allow_html=True)
                if st.button("🗑️ Supprimer le produit sélectionné", use_container_width=True):
                    if del_mat != "-- Sélectionner --":
                        config["materiaux"].remove(del_mat)
                        sauvegarder_config(config)
                        st.success(f"Produit « {del_mat} » retiré de la liste !")
                        st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

            # 4. Compagnons
            st.write("---")
            st.markdown("**👷 4. Compagnons / Équipe**")
            c_o1, c_o2 = st.columns(2)
            with c_o1:
                nouveau_macon = st.text_input("Nouvel ouvrier à ajouter", placeholder="Ex : NOM & Prénom...", key="add_macon")
                if st.button("➕ Ajouter cet ouvrier", use_container_width=True):
                    val = nouveau_macon.strip()
                    if val and val not in config["macons"]:
                        config["macons"].append(val)
                        sauvegarder_config(config)
                        st.success(f"Compagnon « {val} » ajouté !")
                        st.rerun()
            with c_o2:
                del_macon = st.selectbox("Sélectionner un ouvrier à supprimer", ["-- Sélectionner --"] + config["macons"], key="del_macon_box")
                st.markdown('<div class="btn-supprimer">', unsafe_allow_html=True)
                if st.button("🗑️ Supprimer l'ouvrier sélectionné", use_container_width=True):
                    if del_macon != "-- Sélectionner --":
                        config["macons"].remove(del_macon)
                        sauvegarder_config(config)
                        st.success(f"Compagnon « {del_macon} » retiré de la liste !")
                        st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

        st.write("---")

        df_all = charger_donnees()

        total_photos = 0
        for root, _, files in os.walk(PHOTOS_BASE_DIR):
            total_photos += len([f for f in files if f.lower().endswith(('.jpg', '.jpeg', '.png'))])

        kpi1, kpi2, kpi3 = st.columns(3)
        with kpi1:
            st.metric("Rapports Validés", f"{len(df_all) if df_all is not None else 0}")
        with kpi2:
            st.metric("Documentation Photo", f"{total_photos} clichés")
        with kpi3:
            projets_actifs = len(df_all["Chantier"].dropna().unique()) if df_all is not None and not df_all.empty else 0
            st.metric("Projets Renseignés", f"{projets_actifs}")

        st.write("---")
        st.markdown("### 📑 Exportations Documentaires Officielles")
        c_exp1, c_exp2 = st.columns([1, 1])

        with c_exp1:
            if df_all is not None and not df_all.empty:
                excel_bytes = generer_rapport_excel(df_all)
                st.download_button(
                    label="📊 EXPORT CLASSEUR EXCEL (.XLSX)",
                    data=excel_bytes,
                    file_name=f"Synthese_Chantiers_{date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.button("Export Excel (En attente de données)", disabled=True, use_container_width=True)

        with c_exp2:
            st.markdown("""
                <div style='color: #FFFFFF !important; font-size: 14px; font-weight: 600; line-height: 1.5; padding-top: 5px;'>
                    Le classeur inclut : Synthèse des surfaces, ratios de consommation, pointage des équipes et journal brut d'exécution.
                </div>
            """, unsafe_allow_html=True)

        st.write("---")
        dossiers_chantiers = [d for d in os.listdir(PHOTOS_BASE_DIR) if os.path.isdir(os.path.join(PHOTOS_BASE_DIR, d))]
        if dossiers_chantiers:
            st.markdown("### 📁 Dossiers Justificatifs Photographiques")
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

                    c_info, c_btn = st.columns([3, 1])
                    with c_info:
                        st.markdown(f"""
                            <div style='margin-top: 8px;'>
                                <span style='color: #FFFFFF !important; font-size: 16px; font-weight: 800;'>📁 {d_ch}</span>
                                <span style='background-color: #1E293B; color: #38BDF8 !important; padding: 4px 10px; border-radius: 6px; font-size: 13px; font-weight: 700; margin-left: 8px; border: 1px solid #334155;'>
                                    {len(fichiers_total)} photos classées
                                </span>
                            </div>
                        """, unsafe_allow_html=True)
                    with c_btn:
                        st.download_button(label=f"Télécharger ZIP", data=zip_buf.getvalue(), file_name=f"Photos_{d_ch}.zip", mime="application/zip", key=f"z_{d_ch}", use_container_width=True)

        st.write("---")
        # REGISTRE D'ATTACHEMENT AVEC SÉLECTION INTUITIVE POUR LES CORRECTIONS
        if df_all is not None and not df_all.empty and "Chantier" in df_all.columns:
            st.markdown("### 🔍 Registre d'Attachement & Suivi des Ouvrages")
            chantiers_bruts = [c for c in df_all["Chantier"].dropna().unique() if not str(c).startswith("2026-") and str(c).strip()]
            f_proj = st.selectbox("Filtrer par projet / chantier :", ["Tous les projets"] + list(chantiers_bruts))

            df_show = df_all.copy()
            if f_proj != "Tous les projets":
                df_show = df_show[df_show["Chantier"] == f_proj]

            for orig_idx, row in df_show.iloc[::-1].iterrows():
                if str(row.get("Chantier", "")).startswith("2026-") or not str(row.get("Chantier", "")).strip():
                    continue

                conso_val = str(row.get("Consommation", "")).strip()
                if not conso_val or conso_val == "nan" or conso_val == "Aucun":
                    badge_conso_html = "<span style='color: #94A3B8 !important; font-size: 13px; font-style: italic;'>Aucune consommation déclarée</span>"
                else:
                    items_conso = conso_val.split(" | ")
                    badge_conso_html = "".join([f"<span class='tag-materiau'>🧪 {c}</span>" for c in items_conso])

                st.markdown(f"""
                    <div class='card-intervention'>
                        <div style='display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #334155; padding-bottom: 10px; margin-bottom: 12px;'>
                            <span class='tag-projet'>🏢 {row.get("Chantier", "")}</span>
                            <span class='tag-date'>📅 {row.get("Date", "")}</span>
                        </div>
                        <div style='margin-bottom: 10px;'>
                            <span class='tag-corps'>🛠️ {row.get("Corps_d_etat", "")}</span>
                            <span class='tag-rendement'>📏 {row.get("Rendement", "")} {row.get("Unite", "")}</span>
                            <span style='margin-left: 10px; color: #38BDF8 !important; font-size: 13px; font-weight: 700;'>Phase : {row.get("Phase", "")}</span>
                        </div>
                        <div style='margin: 10px 0;'>
                            {badge_conso_html}
                        </div>
                        <div style='color: #FFFFFF !important; font-size: 14px; font-weight: 600; margin-top: 8px;'>
                            👷 <b>Effectif présent :</b> {row.get("Effectif", "")}
                        </div>
                        <div style='color: #E2E8F0 !important; font-size: 14px; margin-top: 6px; font-style: italic;'>
                            💬 <b>Note de chantier :</b> {row.get("Legende", "R.A.S")}
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
                    cols = st.columns(min(len(valid_paths), 4))
                    for i, p_img in enumerate(valid_paths):
                        try:
                            cols[i % 4].image(Image.open(p_img), use_container_width=True)
                        except Exception:
                            pass

                # CORRECTION PAR SÉLECTION INTUITIVE (EFFECTIF MULTISELECT + MATÉRIAUX DYNAMIQUES)
                with st.expander(f"✏️ Corriger / Modifier cette fiche ({row.get('Chantier','')} - {row.get('Date','')})"):
                    st.caption("Modifiez directement les erreurs commises par l'équipe sur le terrain :")
                    
                    # 1. Date et Chantier
                    m_c1, m_c2 = st.columns(2)
                    with m_c1:
                        try:
                            d_val = datetime.strptime(str(row.get("Date", "")), "%Y-%m-%d").date()
                        except Exception:
                            d_val = date.today()
                        m_date = st.date_input("Date", value=d_val, key=f"m_date_{orig_idx}")
                    with m_c2:
                        ch_actuel = str(row.get("Chantier", ""))
                        idx_ch = config["chantiers"].index(ch_actuel) if ch_actuel in config["chantiers"] else 0
                        m_chantier = st.selectbox("Chantier", config["chantiers"], index=idx_ch, key=f"m_ch_{orig_idx}")

                    # 2. Corps d'état et Phase
                    m_c3, m_c4 = st.columns(2)
                    with m_c3:
                        tache_actuelle = str(row.get("Corps_d_etat", ""))
                        idx_t = config["taches"].index(tache_actuelle) if tache_actuelle in config["taches"] else 0
                        m_tache = st.selectbox("Corps d'état", config["taches"], index=idx_t, key=f"m_t_{orig_idx}")
                    with m_c4:
                        phases_possibles = [
                            "Pendant exécution / application",
                            "Avant travaux (État du support)",
                            "Après achèvement (Finition)",
                            "Détail technique / Gorge / Relevé",
                            "Épreuve d'eau (Test d'étanchéité)",
                            "Autre"
                        ]
                        ph_actuelle = str(row.get("Phase", ""))
                        idx_ph = phases_possibles.index(ph_actuelle) if ph_actuelle in phases_possibles else 0
                        m_phase = st.selectbox("Phase", phases_possibles, index=idx_ph, key=f"m_ph_{orig_idx}")

                    # 3. Rendement et Unité
                    m_c5, m_c6 = st.columns(2)
                    with m_c5:
                        try:
                            rend_init = float(row.get("Rendement", 0))
                        except Exception:
                            rend_init = 0.0
                        m_rendement = st.number_input("Rendement réalisé", value=rend_init, step=1.0, format="%.2f", key=f"m_rend_{orig_idx}")
                    with m_c6:
                        unites_possibles = ["m²", "ML", "U"]
                        u_actuelle = str(row.get("Unite", "m²"))
                        idx_u = unites_possibles.index(u_actuelle) if u_actuelle in unites_possibles else 0
                        m_unite = st.selectbox("Unité", unites_possibles, index=idx_u, key=f"m_u_{orig_idx}")

                    # 4. EFFECTIF PAR MULTISELECT DIRECT
                    raw_eff = [e.strip() for e in str(row.get("Effectif", "")).split(",") if e.strip()]
                    def_eff = [e for e in raw_eff if e in config["macons"]]
                    m_effectif = st.multiselect("👷 Ouvriers présents", config["macons"], default=def_eff, key=f"m_eff_{orig_idx}")

                    # 5. CONSOMMATION MATÉRIAUX GÉRÉE PAR SÉLECTION
                    st.write("---")
                    st.markdown("##### 🧪 Consommation Matériaux")
                    
                    state_key = f"m_conso_list_{orig_idx}"
                    if state_key not in st.session_state:
                        liste_init = []
                        conso_str = str(row.get("Consommation", "")).strip()
                        if conso_str and conso_str != "nan" and conso_str != "Aucun":
                            for part in conso_str.split(" | "):
                                if ":" in part:
                                    p_nom, p_qte_uni = part.split(":", 1)
                                    p_nom = p_nom.strip()
                                    q_u = p_qte_uni.strip().split(" ", 1)
                                    try:
                                        q_val = float(q_u[0])
                                    except Exception:
                                        q_val = 1.0
                                    u_val = q_u[1] if len(q_u) > 1 else "Seaux/Bidons"
                                    liste_init.append({"produit": p_nom, "quantite": q_val, "unite": u_val})
                        st.session_state[state_key] = liste_init

                    if st.session_state[state_key]:
                        for c_idx, c_item in enumerate(st.session_state[state_key]):
                            c_col_txt, c_col_del = st.columns([4, 1])
                            with c_col_txt:
                                st.markdown(f"<span class='tag-materiau'>📦 {c_item['produit']} : <b>{c_item['quantite']} {item['unite']}</b></span>", unsafe_allow_html=True)
                            with c_col_del:
                                st.markdown('<div class="btn-supprimer">', unsafe_allow_html=True)
                                if st.button("❌", key=f"del_mconso_{orig_idx}_{c_idx}"):
                                    st.session_state[state_key].pop(c_idx)
                                    st.rerun()
                                st.markdown('</div>', unsafe_allow_html=True)

                    c_mat_col1, c_mat_col2, c_mat_col3 = st.columns([2, 1, 1])
                    with c_mat_col1:
                        new_mat_m = st.selectbox("Produit", config["materiaux"], key=f"sel_m_mat_{orig_idx}")
                    with c_mat_col2:
                        new_qte_m = st.number_input("Quantité", min_value=0.0, step=1.0, format="%.2f", key=f"sel_m_qte_{orig_idx}")
                    with c_mat_col3:
                        new_uni_m = st.selectbox("Unité", ["Seaux/Bidons", "Sacs", "Rouleaux", "Kg", "Litres", "Cartouches", "U"], key=f"sel_m_uni_{orig_idx}")

                    if st.button("➕ Ajouter à la consommation", key=f"btn_add_mat_{orig_idx}", use_container_width=True):
                        if new_qte_m > 0:
                            st.session_state[state_key].append({
                                "produit": new_mat_m,
                                "quantite": new_qte_m,
                                "unite": new_uni_m
                            })
                            st.rerun()
                        else:
                            st.warning("⚠️ Indiquez une quantité supérieure à 0.")

                    # 6. Observation
                    st.write("---")
                    m_legende = st.text_input("Note de chantier / Observation", value=str(row.get("Legende", "")), key=f"m_leg_{orig_idx}")

                    st.write("")
                    st.markdown('<div class="btn-enregistrer-modif">', unsafe_allow_html=True)
                    if st.button("💾 Enregistrer les corrections", key=f"btn_save_{orig_idx}", use_container_width=True):
                        if st.session_state[state_key]:
                            conso_finale_mod = " | ".join([f"{it['produit']}: {it['quantite']} {it['unite']}" for it in st.session_state[state_key]])
                        else:
                            conso_finale_mod = "Aucun"

                        df_all.loc[orig_idx, "Date"] = str(m_date)
                        df_all.loc[orig_idx, "Chantier"] = str(m_chantier)
                        df_all.loc[orig_idx, "Corps_d_etat"] = str(m_tache)
                        df_all.loc[orig_idx, "Phase"] = str(m_phase)
                        df_all.loc[orig_idx, "Rendement"] = str(m_rendement)
                        df_all.loc[orig_idx, "Unite"] = str(m_unite)
                        df_all.loc[orig_idx, "Consommation"] = conso_finale_mod
                        df_all.loc[orig_idx, "Effectif"] = ", ".join(m_effectif)
                        df_all.loc[orig_idx, "Nb_Ouvriers"] = len(m_effectif)
                        df_all.loc[orig_idx, "Legende"] = str(m_legende).replace(";", " ").replace("|", " ")

                        sauvegarder_donnees(df_all)
                        st.success("✅ Fiche corrigée et sauvegardée avec succès !")
                        st.rerun()
                    st.markdown('</div>', unsafe_allow_html=True)

                # SUPPRESSION DÉFINITIVE D'UNE FICHE
                st.markdown('<div class="btn-supprimer" style="margin-top: 4px;">', unsafe_allow_html=True)
                if st.button(f"🗑️ Supprimer cette fiche définitivement", key=f"del_row_{orig_idx}"):
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
