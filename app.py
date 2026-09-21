import streamlit as st
import pandas as pd
from datetime import date
import os
import json
import io
import zipfile
import re
from PIL import Image
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Chantier & Suivi", page_icon="📱", layout="centered")

# Style du bouton vert d'envoi
st.markdown("""
    <style>
    div.stButton > button {
        border-radius: 10px !important;
        font-weight: 600 !important;
    }
    .btn-valider button {
        background-color: #10B981 !important;
        color: white !important;
        font-size: 18px !important;
        font-weight: 700 !important;
        border-radius: 12px !important;
        border: none !important;
        padding: 14px 20px !important;
        width: 100% !important;
        box-shadow: 0 4px 6px -1px rgba(16, 185, 129, 0.3) !important;
    }
    .btn-valider button:hover {
        background-color: #059669 !important;
        color: white !important;
    }
    div.stDownloadButton > button {
        background-color: #0F766E !important;
        color: white !important;
        font-weight: 600 !important;
        border-radius: 10px !important;
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
                if "materiaux" not in data:
                    data["materiaux"] = CONFIG_DEFAUT["materiaux"]
                return data
        except Exception:
            return CONFIG_DEFAUT
    return CONFIG_DEFAUT

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

def generer_rapport_excel(df_source):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Arial", size=10)
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

        ws.row_dimensions[1].height = 26
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # Feuille 1: Synthèse
    ws1 = wb.create_sheet(title="Synthèse Chantiers")
    headers1 = ["Chantier", "Surface Réalisée (m²)", "Linéaire Réalisé (ML)", "Unités (U)", "Interventions"]
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

    # Feuille 2: Consommation
    ws2 = wb.create_sheet(title="Consommation Matériaux")
    headers2 = ["Date", "Chantier", "Corps d'État", "Matériaux Consommés", "Remarques"]
    ws2.append(headers2)
    for _, r in df_source.iterrows():
        c_mat = str(r.get("Consommation", "")).strip()
        if not c_mat or c_mat == "nan":
            c_mat = "Aucun"
        ws2.append([r.get("Date", ""), r.get("Chantier", ""), r.get("Corps_d_etat", ""), c_mat, r.get("Legende", "")])
    styliser_feuille(ws2, headers2)

    # Feuille 3: Effectif
    ws3 = wb.create_sheet(title="Pointage Ouvriers")
    headers3 = ["Date", "Chantier", "Corps d'État", "Effectif Présent", "Nombre d'Ouvriers"]
    ws3.append(headers3)
    for _, r in df_source.iterrows():
        ws3.append([r.get("Date", ""), r.get("Chantier", ""), r.get("Corps_d_etat", ""), r.get("Effectif", ""), r.get("Nb_Ouvriers", "")])
    styliser_feuille(ws3, headers3)

    # Feuille 4: Journal Détaillé
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

# Initialisation de la mémoire pour les consommations
if "liste_consommations" not in st.session_state:
    st.session_state.liste_consommations = []

tab_saisie, tab_admin = st.tabs(["📲 Saisie Chantier", "📊 Tableau de Bord"])

# -------------------------------------------------------------
# ONGLET 1 : SAISIE AVEC MULTI-PRODUITS ILLIMITÉ
# -------------------------------------------------------------
with tab_saisie:
    st.markdown("""
        <div style='background-color: #0F766E; padding: 14px; border-radius: 12px; text-align: center; margin-bottom: 15px;'>
            <h2 style='color: white; margin: 0; font-size: 20px;'>📱 Saisie du Rapport</h2>
            <p style='color: #CCFBF1; margin: 4px 0 0 0; font-size: 13px;'>Renseignez les travaux, consommations et photos</p>
        </div>
    """, unsafe_allow_html=True)

    date_jour = st.date_input("📅 Date de la journée", value=date.today())
    chantier_sel = st.selectbox("🏢 Chantier", config["chantiers"])
    tache_sel = st.selectbox("🛠️ Corps d'état", config["taches"])
    phase_travaux = st.selectbox("📌 Étape de réalisation", [
        "Pendant exécution / application",
        "Avant travaux (État du support)",
        "Après achèvement (Finition)",
        "Détail technique / Gorge / Relevé",
        "Épreuve d'eau (Test d'étanchéité)",
        "Autre"
    ])

    macons_presents = st.multiselect("👷 Ouvriers présents", config["macons"], placeholder="Sélectionnez les ouvriers...")

    c_r1, c_r2 = st.columns([2, 1])
    with c_r1:
        rendement = st.number_input("📏 Rendement réalisé", min_value=0.0, step=1.0, format="%.2f")
    with c_r2:
        unite = st.selectbox("Unité", ["m²", "ML", "U"])

    # --- SECTION CONSOMMATION MULTI-PRODUITS ---
    st.write("---")
    st.markdown("### 🧪 Matériaux Consommés")
    
    # Affichage des produits déjà enregistrés dans la liste temporaire
    if st.session_state.liste_consommations:
        st.markdown("**Matériaux ajoutés pour ce rapport :**")
        for idx_c, item in enumerate(st.session_state.liste_consommations):
            col_txt, col_sup = st.columns([4, 1])
            with col_txt:
                st.info(f"✔️ {item['produit']} : **{item['quantite']} {item['unite']}**")
            with col_sup:
                if st.button("❌", key=f"del_{idx_c}"):
                    st.session_state.liste_consommations.pop(idx_c)
                    st.rerun()

    # Formulaire d'ajout rapide d'un produit
    with st.expander("➕ Ajouter un matériau consommé", expanded=True):
        cp1, cp2, cp3 = st.columns([2, 1, 1])
        with cp1:
            nouveau_mat = st.selectbox("Produit", config["materiaux"], key="ajout_mat")
        with cp2:
            nouvelle_qte = st.number_input("Quantité", min_value=0.0, step=1.0, format="%.2f", key="ajout_qte")
        with cp3:
            nouvelle_uni = st.selectbox("Unité", ["Seaux/Bidons", "Sacs", "Rouleaux", "Kg", "Litres", "Cartouches", "U"], key="ajout_uni")

        if st.button("➕ Ajouter ce produit à la liste", use_container_width=True):
            if nouvelle_qte > 0:
                st.session_state.liste_consommations.append({
                    "produit": nouveau_mat,
                    "quantite": nouvelle_qte,
                    "unite": nouvelle_uni
                })
                st.rerun()
            else:
                st.warning("⚠️ Indiquez une quantité supérieure à 0.")

    st.write("---")
    st.markdown("### 📸 Photos du chantier")
    photos_galerie = st.file_uploader(
        "Prenez ou sélectionnez vos photos",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True
    )
    legende = st.text_input("💬 Observation", placeholder="Ex : Travail terminé conformément...")

    st.write("")
    
    # GROS BOUTON VERT FINAL
    st.markdown('<div class="btn-valider">', unsafe_allow_html=True)
    envoyer_btn = st.button("✅ ENVOYER LE RAPPORT DU JOUR", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if envoyer_btn:
        if not macons_presents:
            st.error("⚠️ Veuillez sélectionner au moins un ouvrier présent.")
        elif not photos_galerie:
            st.error("⚠️ Veuillez ajouter au moins une photo pour le rapport.")
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
                    if hasattr(p, "name") and os.path.splitext(p.name)[1]:
                        ext = os.path.splitext(p.name)[1].lower()

                    nom_fichier = f"{tache_clean}_{idx+1}{ext}"
                    chemin_disque = os.path.join(chemin_cible, nom_fichier)

                    with open(chemin_disque, "wb") as f_img:
                        f_img.write(p.getbuffer())

                    saved_files.append(f"{dossier_chantier}/{dossier_date}/{nom_fichier}")

                # Formatage de l'ensemble des consommations enregistrées
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

                # Réinitialisation de la liste des produits après envoi
                st.session_state.liste_consommations = []
                st.balloons()
                st.success(f"🎉 Rapport enregistré avec succès ! Données et photos dans : 📁 {dossier_chantier} / 📅 {dossier_date}")
            except Exception as e:
                st.error(f"❌ Erreur lors de l'enregistrement : {e}")

# -------------------------------------------------------------
# ONGLET 2 : TABLEAU DE BORD
# -------------------------------------------------------------
with tab_admin:
    st.markdown("""
        <div style='background-color: #1E293B; padding: 14px; border-radius: 12px; text-align: center; margin-bottom: 15px;'>
            <h2 style='color: white; margin: 0; font-size: 20px;'>📊 Tableau de Bord</h2>
        </div>
    """, unsafe_allow_html=True)

    pin = st.text_input("Code Administrateur :", type="password", placeholder="Code...")

    if pin == ADMIN_PIN:
        df_all = charger_donnees()

        total_photos = 0
        for root, _, files in os.walk(PHOTOS_BASE_DIR):
            total_photos += len([f for f in files if f.lower().endswith(('.jpg', '.jpeg', '.png'))])

        kpi1, kpi2 = st.columns(2)
        with kpi1:
            st.metric("Photos Reçues", f"{total_photos} 📸")
        with kpi2:
            total_rapports = len(df_all) if df_all is not None else 0
            st.metric("Rapports Validés", f"{total_rapports}")

        st.write("---")
        if df_all is not None and not df_all.empty:
            excel_bytes = generer_rapport_excel(df_all)
            st.download_button(
                label="📊 TÉLÉCHARGER LE RAPPORT EXCEL (.xlsx)",
                data=excel_bytes,
                file_name=f"Rapport_Mensuel_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.info("Aucune donnée enregistrée pour le moment.")

        st.write("---")
        dossiers_chantiers = [d for d in os.listdir(PHOTOS_BASE_DIR) if os.path.isdir(os.path.join(PHOTOS_BASE_DIR, d))]

        if dossiers_chantiers:
            st.markdown("### 📦 Télécharger ZIP Photos")
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

                    c_info, c_btn = st.columns([2, 1])
                    with c_info:
                        st.markdown(f"📁 **{d_ch}** ({len(fichiers_total)} photos)")
                    with c_btn:
                        st.download_button(label=f"⬇️ ZIP", data=zip_buf.getvalue(), file_name=f"photos_{d_ch}.zip", mime="application/zip", key=f"z_{d_ch}", use_container_width=True)

        st.write("---")
        if df_all is not None and not df_all.empty and "Chantier" in df_all.columns:
            chantiers_bruts = [c for c in df_all["Chantier"].dropna().unique() if not str(c).startswith("2026-") and str(c).strip()]
            f_proj = st.selectbox("🔍 Filtrer les fiches :", ["Tous les chantiers"] + chantiers_bruts)

            df_show = df_all.copy()
            if f_proj != "Tous les chantiers":
                df_show = df_show[df_show["Chantier"] == f_proj]

            for _, row in df_show.iloc[::-1].iterrows():
                if str(row.get("Chantier", "")).startswith("2026-"):
                    continue

                conso_val = str(row.get("Consommation", "")).strip()
                if not conso_val or conso_val == "nan":
                    conso_val = "Aucun"

                with st.container():
                    st.markdown(f"""
                        <div style='background-color: #F8FAFC; border-left: 5px solid #0F766E; padding: 12px; border-radius: 8px; margin-bottom: 10px; border: 1px solid #E2E8F0;'>
                            <div style='display: flex; justify-content: space-between;'>
                                <b>🏢 {row.get("Chantier", "")}</b>
                                <span style='color: #64748B;'>📅 {row.get("Date", "")}</span>
                            </div>
                            <div style='margin-top: 5px;'>
                                🛠️ {row.get("Corps_d_etat", "")} &nbsp;|&nbsp; 📏 <b>{row.get("Rendement", "")} {row.get("Unite", "")}</b>
                            </div>
                            <div style='margin-top: 4px; font-size: 13px; color: #0369A1;'>
                                🧪 <b>Matériaux :</b> {conso_val}
                            </div>
                            <div style='color: #64748B; font-size: 13px; margin-top: 3px;'>
                                👷 {row.get("Effectif", "")}
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
                        cols = st.columns(len(valid_paths) if len(valid_paths) <= 3 else 3)
                        for i, p_img in enumerate(valid_paths):
                            try:
                                cols[i % 3].image(Image.open(p_img), use_container_width=True)
                            except Exception:
                                pass
                    st.write("")

        with st.expander("⚙️ Paramètres avancés"):
            if st.button("🗑️ Réinitialiser CSV"):
                if os.path.exists(CSV_FILE):
                    os.remove(CSV_FILE)
                    st.rerun()

    elif pin != "":
        st.error("❌ Code incorrect.")
