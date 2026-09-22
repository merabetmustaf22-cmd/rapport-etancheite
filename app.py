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
    page_title="Système de Suivi Technique Étanchéité",
    page_icon="🏗️",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# --- CSS HAUTE LUMINOSITÉ SANS AUCUN TEXTE GRIS ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    /* Fond général bleu nuit */
    .stApp {
        background-color: #0B1120 !important;
    }

    /* TOUS LES TITRES EN BLANC PUR ÉCLATANT */
    h1, h2, h3, h4, h5, h6, .stMarkdown h1, .stMarkdown h2, .stMarkdown h3, .stMarkdown h4, .stMarkdown h5 {
        color: #FFFFFF !important;
        font-weight: 800 !important;
        word-break: normal !important;
        hyphens: none !important;
    }

    /* TEXTES GÉNÉRAUX EN BLANC PUR (PLUS DE GRIS) */
    .stMarkdown p, .stCaption p, [data-testid="stMarkdownContainer"] p {
        color: #FFFFFF !important;
    }

    /* EN-TÊTE SMARTPHONE SANS MOT COUPÉ */
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
        font-size: 13px !important;
        color: #38BDF8 !important;
        margin-top: 5px !important;
        font-weight: 700 !important;
        line-height: 1.3 !important;
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

    /* LABELS DES FORMULAIRES */
    .stWidgetLabel p, [data-testid="stWidgetLabel"] p, label p {
        color: #FFFFFF !important;
        font-size: 15px !important;
        font-weight: 700 !important;
    }

    /* CASES DE SAISIE EN FOND BLANC ET TEXTE SOMBRE */
    .stTextInput input, .stDateInput input, .stNumberInput input {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        font-size: 15px !important;
        font-weight: 700 !important;
        border: 2px solid #CBD5E1 !important;
        border-radius: 8px !important;
    }
    .stTextInput input::placeholder {
        color: #64748B !important;
        font-weight: 500 !important;
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

    /* MENUS DÉROULANTS */
    div[data-baseweb="popover"], ul[role="listbox"], li[role="option"] {
        background-color: #FFFFFF !important;
    }
    div[data-baseweb="popover"] *, ul[role="listbox"] *, li[role="option"] * {
        color: #0F172A !important;
        font-weight: 700 !important;
    }

    /* CALENDRIER */
    div[data-baseweb="calendar"], div[data-baseweb="calendar"] * {
        color: #0F172A !important;
        font-weight: 700 !important;
    }
    div[data-baseweb="calendar"] button[aria-selected="true"] {
        background-color: #0F766E !important;
        color: #FFFFFF !important;
    }

    /* SÉLECTEUR DE PHOTOS */
    [data-testid="stFileUploader"] section {
        background-color: #FFFFFF !important;
        border: 2px dashed #94A3B8 !important;
        border-radius: 10px !important;
        padding: 14px !important;
    }
    [data-testid="stFileUploader"] section * {
        color: #0F172A !important;
        font-weight: 700 !important;
    }
    [data-testid="stFileUploader"] section button {
        background-color: #0284C7 !important;
        border: none !important;
    }
    [data-testid="stFileUploader"] section button * {
        color: #FFFFFF !important;
        font-weight: 700 !important;
    }

    /* VOLETS DÉPLIANTS (EXPANDERS) EN BLANC BRILLANT */
    [data-testid="stExpander"] summary {
        background-color: #1E293B !important;
        border-radius: 8px !important;
        border: 1px solid #334155 !important;
    }
    [data-testid="stExpander"] summary * {
        color: #FFFFFF !important;
        font-weight: 800 !important;
    }

    /* CHIFFRES STATISTIQUES / KPI */
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

    /* BOUTONS D'ACTION */
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

    /* BOUTONS DE TÉLÉCHARGEMENT */
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

    /* CARTES DES RAPPORTS ET ATTACHEMENTS */
    .card-intervention {
        background-color: #111827;
        border-radius: 14px;
        border: 1px solid #334155;
        padding: 16px;
        margin-bottom: 16px;
        border-left: 5px solid #10B981;
    }
    .tag-projet {
        font-size: 16px;
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

    /* ONGLETS */
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

# --- FONCTION DE WATERMARKING TECHNIQUE OFFICIEL SUR PHOTOS ---
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
        # En cas d'imprévu, sauvegarde sécurisée de la photo originale
        return image_file.getbuffer()

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
# ONGLET 1 : SAISIE TERRAIN HAUTE LISIBILITÉ
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
    st.markdown("### 📸 Pièces Jointes & Filigrane Technique")
    photos_galerie = st.file_uploader(
        "Photos justificatives de l'ouvrage (Horodatage automatique)",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True
    )
    legende = st.text_input("Observations techniques particulières", placeholder="Ex : Support brossé et dépoussiéré avant couche primaire...")

    st.write("")
    st.markdown('<div class="btn-valider">', unsafe_allow_html=True)
    envoyer_btn = st.button("TRANSMETTRE LE RAPPORT JOURNALIER", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if envoyer_btn:
        if not macons_presents:
            st.error("⚠️ Veuillez déclarer au moins un ouvrier pour cet ouvrage.")
        elif not photos_galerie:
            st.error("⚠️ La conformité technique impose l'ajout d'au moins une photo justificative.")
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

                    # APPLICATION AUTOMATIQUE DU WATERMARK TECHNIQUE
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
                st.success(f"✅ Rapport validé et intégré au registre officiel : {dossier_chantier} [{dossier_date}]")
            except Exception as e:
                st.error(f"❌ Erreur de transmission : {e}")

# -------------------------------------------------------------
# ONGLET 2 : ESPACE CADRE, DIRECTION & GESTION
# -------------------------------------------------------------
with tab_admin:
    st.markdown("""
        <div class='header-cadre'>
            <div style='display: flex; justify-content: space-between; align-items: center;'>
                <div>
                    <h1 class='header-title'>Tableau de Bord & Attachements</h1>
                    <div class='header-sub'>Supervision technique, synthèse des consommations et exports</div>
                </div>
                <div>
                    <span class='badge-pro'>SUPERVISION</span>
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)

    pin = st.text_input("Authentification Responsable (Code PIN) :", type="password", placeholder="Saisir le code d'accès...")

    if pin == ADMIN_PIN:
        with st.expander("⚙️ Configuration des Listes (Ajouter / Supprimer des options)", expanded=False):
            st.markdown("<p style='color: #FFFFFF !important; font-weight: 800; font-size: 15px;'>🛠️ Gestion des Chantiers, Tâches, Matériaux et Compagnons</p>", unsafe_allow_html=True)
            
            # 1. CHANTIERS
            st.write("---")
            st.markdown("<p style='color: #38BDF8 !important; font-weight: 800;'>🏢 1. Chantiers / Projets</p>", unsafe_allow_html=True)
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

            # 2. CORPS D'ÉTAT
            st.write("---")
            st.markdown("<p style='color: #38BDF8 !important; font-weight: 800;'>🛠️ 2. Corps d'État / Tâches</p>", unsafe_allow_html=True)
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

            # 3. MATÉRIAUX
            st.write("---")
            st.markdown("<p style='color: #38BDF8 !important; font-weight: 800;'>🧪 3. Matériaux & Produits Consommés</p>", unsafe_allow_html=True)
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

            # 4. COMPAGNONS
            st.write("---")
            st.markdown("<p style='color: #38BDF8 !important; font-weight: 800;'>👷 4. Compagnons / Équipe</p>", unsafe_allow_html=True)
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
                    badge_conso_html = "<span style='color: #38BDF8 !important; font-size: 13px; font-weight: 600;'>📦 Aucune consommation déclarée</span>"
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
                        <div style='color: #FFFFFF !important; font-size: 14px; font-weight: 700; margin-top: 8px;'>
                            👷 <b>Effectif présent :</b> {row.get("Effectif", "")}
                        </div>
                        <div style='color: #FFFFFF !important; font-size: 14px; margin-top: 6px; font-style: italic;'>
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

                st.markdown('<div class="btn-supprimer" style="margin-top: 6px;">', unsafe_allow_html=True)
                if st.button(f"🗑️ Supprimer cette fiche ({row.get('Chantier','')} - {row.get('Date','')})", key=f"del_row_{orig_idx}"):
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
