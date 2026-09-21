import streamlit as st
import pandas as pd
from datetime import date
import os
import json
import io
import zipfile
import re
import shutil
from PIL import Image

st.set_page_config(page_title="Chantier & Suivi", page_icon="📱", layout="centered")

PHOTOS_BASE_DIR = "photos_chantier"
os.makedirs(PHOTOS_BASE_DIR, exist_ok=True)
CONFIG_FILE = "config_chantier.json"
CSV_FILE = "suivi_journalier_chantiers.csv"
ADMIN_PIN = "2026"

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
    ]
}

def clean_folder_name(name):
    clean = name.split('(')[0].strip()
    return re.sub(r'[^a-zA-Z0-9_-]', '_', clean)

# --- AUTO-MIGRATION DES ANCIENNES PHOTOS VERS LEURS SOUS-DOSSIERS ---
def ranger_photos_orphelines(chantiers_connus):
    for item in os.listdir(PHOTOS_BASE_DIR):
        item_path = os.path.join(PHOTOS_BASE_DIR, item)
        if os.path.isfile(item_path) and item.lower().endswith(('.jpg', '.jpeg', '.png')):
            # Trouver à quel chantier appartient cette photo
            chantier_trouve = "Autre"
            for ch in chantiers_connus:
                ch_key = clean_folder_name(ch)
                if ch_key.lower() in item.lower():
                    chantier_trouve = ch_key
                    break
            
            dossier_cible = os.path.join(PHOTOS_BASE_DIR, chantier_trouve)
            os.makedirs(dossier_cible, exist_ok=True)
            try:
                shutil.move(item_path, os.path.join(dossier_cible, item))
            except Exception:
                pass

def charger_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return CONFIG_DEFAUT
    return CONFIG_DEFAUT

def sauver_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

def charger_donnees():
    if not os.path.exists(CSV_FILE):
        return None
    try:
        df = pd.read_csv(CSV_FILE, sep=';', encoding='utf-8-sig', on_bad_lines='skip')
        if "Chantier" not in df.columns or len(df.columns) < 8:
            df = pd.read_csv(CSV_FILE, sep=',', encoding='utf-8-sig', on_bad_lines='skip')
        return df
    except Exception:
        return None

config = charger_config()
ranger_photos_orphelines(config["chantiers"])

tab_saisie, tab_admin = st.tabs(["📲 Saisie Chantier", "📊 Tableau de Bord"])

# -------------------------------------------------------------
# ONGLET 1 : SAISIE CHANTIER
# -------------------------------------------------------------
with tab_saisie:
    st.markdown("""
        <div style='background-color: #0F766E; padding: 14px; border-radius: 12px; text-align: center; margin-bottom: 15px;'>
            <h2 style='color: white; margin: 0; font-size: 20px;'>📱 Pointage & Photos Chantier</h2>
            <p style='color: #CCFBF1; margin: 4px 0 0 0; font-size: 13px;'>Enregistrez l'avancement et les photos du jour</p>
        </div>
    """, unsafe_allow_html=True)

    with st.form("form_mobile_chantier", clear_on_submit=True):
        date_jour = st.date_input("📅 Date", value=date.today())
        chantier_sel = st.selectbox("🏢 Chantier", config["chantiers"])

        tache_sel = st.selectbox("🛠️ Travail / Corps d'état", config["taches"])
        phase_travaux = st.selectbox("📌 Étape de réalisation", [
            "Pendant exécution / application",
            "Avant travaux (État du support)",
            "Après achèvement (Finition)",
            "Détail technique / Gorge / Relevé",
            "Épreuve d'eau (Test d'étanchéité)",
            "Autre"
        ])

        macons_presents = st.multiselect(
            "👷 Ouvriers présents",
            config["macons"],
            placeholder="Sélectionnez les noms..."
        )

        c_r1, c_r2 = st.columns([2, 1])
        with c_r1:
            rendement = st.number_input("📏 Rendement réalisé", min_value=0.0, step=1.0, format="%.2f")
        with c_r2:
            unite = st.selectbox("Unité", ["m²", "ML", "U"])

        st.write("---")
        st.markdown("### 📸 Photos du travail")
        photos_galerie = st.file_uploader(
            "Prendre une photo ou importer depuis la galerie",
            type=["jpg", "jpeg", "png"],
            accept_multiple_files=True
        )

        legende = st.text_input("💬 Remarque / Observation", placeholder="Ex : Zone A finie...")

        st.write("")
        submitted = st.form_submit_button("🚀 ENVOYER LE RAPPORT", use_container_width=True)

        if submitted:
            if not photos_galerie:
                st.error("⚠️ Veuillez ajouter au moins une photo.")
            elif not macons_presents:
                st.error("⚠️ Veuillez sélectionner au moins un ouvrier présent.")
            else:
                dossier_chantier = clean_folder_name(chantier_sel)
                dossier_complet = os.path.join(PHOTOS_BASE_DIR, dossier_chantier)
                os.makedirs(dossier_complet, exist_ok=True)

                saved_files = []
                tache_clean = re.sub(r'[^a-zA-Z0-9_-]', '_', tache_sel)[:10]

                for idx, p in enumerate(photos_galerie):
                    ext = ".jpg"
                    if hasattr(p, "name") and os.path.splitext(p.name)[1]:
                        ext = os.path.splitext(p.name)[1].lower()

                    nom_fichier = f"{date_jour}_{tache_clean}_{idx+1}{ext}"
                    chemin_dest = os.path.join(dossier_complet, nom_fichier)

                    with open(chemin_dest, "wb") as f_img:
                        f_img.write(p.getbuffer())

                    saved_files.append(f"{dossier_chantier}/{nom_fichier}")

                nouvelle_ligne = {
                    "Date": str(date_jour),
                    "Chantier": str(chantier_sel),
                    "Corps_d_etat": str(tache_sel),
                    "Phase": str(phase_travaux),
                    "Rendement": str(rendement),
                    "Unite": str(unite),
                    "Effectif": ", ".join(macons_presents),
                    "Nb_Ouvriers": len(macons_presents),
                    "Photos": "|".join(saved_files),
                    "Nb_Photos": len(saved_files),
                    "Legende": str(legende).replace(";", " ").replace("|", " ")
                }

                df_entry = pd.DataFrame([nouvelle_ligne])
                colonnes_ordre = ["Date", "Chantier", "Corps_d_etat", "Phase", "Rendement", "Unite", "Effectif", "Nb_Ouvriers", "Photos", "Nb_Photos", "Legende"]
                df_entry = df_entry[colonnes_ordre]

                if os.path.exists(CSV_FILE):
                    df_entry.to_csv(CSV_FILE, sep=';', mode='a', header=False, index=False, encoding='utf-8-sig')
                else:
                    df_entry.to_csv(CSV_FILE, sep=';', index=False, encoding='utf-8-sig')

                st.success(f"✅ Photos enregistrées dans le dossier dédié : 📁 {dossier_chantier}")

# -------------------------------------------------------------
# ONGLET 2 : TABLEAU DE BORD (TÉLÉCHARGEMENT STRICTEMENT SÉPARÉ)
# -------------------------------------------------------------
with tab_admin:
    st.markdown("""
        <div style='background-color: #1E293B; padding: 14px; border-radius: 12px; text-align: center; margin-bottom: 15px;'>
            <h2 style='color: white; margin: 0; font-size: 20px;'>📊 Tableau de Bord Chantier</h2>
            <p style='color: #94A3B8; margin: 4px 0 0 0; font-size: 13px;'>Téléchargement séparé par dossier chantier</p>
        </div>
    """, unsafe_allow_html=True)

    pin = st.text_input("Code Administrateur :", type="password", placeholder="Entrez le code...")

    if pin == ADMIN_PIN:
        df_all = charger_donnees()
        
        st.markdown("### 📥 Télécharger les Photos par Chantier")
        st.caption("Chaque bouton télécharge uniquement les photos du chantier sélectionné.")

        # Récupération de tous les dossiers chantiers existants
        dossiers_existants = [d for d in os.listdir(PHOTOS_BASE_DIR) if os.path.isdir(os.path.join(PHOTOS_BASE_DIR, d))]
        
        if not dossiers_existants:
            st.info("Aucun dossier photo disponible pour l'instant.")
        else:
            for d_name in sorted(dossiers_existants):
                target_dir = os.path.join(PHOTOS_BASE_DIR, d_name)
                fichiers_chantier = [f for f in os.listdir(target_dir) if os.path.isfile(os.path.join(target_dir, f))]
                
                if fichiers_chantier:
                    # Préparation de l'archive spécifique à ce chantier seul
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        for f in fichiers_chantier:
                            zf.write(os.path.join(target_dir, f), arcname=f)

                    col_info, col_dl = st.columns([2, 1])
                    with col_info:
                        st.markdown(f"📁 **{d_name}** ({len(fichiers_chantier)} photos)")
                    with col_dl:
                        st.download_button(
                            label=f"⬇️ Télécharger ({d_name})",
                            data=zip_buf.getvalue(),
                            file_name=f"photos_{d_name}_{date.today()}.zip",
                            mime="application/zip",
                            key=f"dl_{d_name}",
                            use_container_width=True
                        )

        st.write("---")
        # Téléchargement du registre Excel
        if df_all is not None and not df_all.empty:
            csv_bytes = df_all.to_csv(sep=';', index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(
                label="📊 Télécharger Tout le Registre (Excel/CSV)",
                data=csv_bytes,
                file_name="registre_chantier.csv",
                mime="text/csv",
                use_container_width=True
            )

        st.write("---")

        # --- AFFICHAGE DU FLUX DE CARTES & PHOTOS ---
        if df_all is not None and not df_all.empty and "Chantier" in df_all.columns:
            chantiers_bruts = [c for c in df_all["Chantier"].dropna().unique() if not str(c).startswith("2026-")]
            liste_projets = ["Tous les chantiers"] + list(chantiers_bruts)
            f_proj = st.selectbox("🔍 Filtrer les fiches par chantier :", liste_projets)

            df_show = df_all.copy()
            if f_proj != "Tous les chantiers":
                df_show = df_show[df_show["Chantier"] == f_proj]

            st.write(f"Affichage de **{len(df_show)}** fiche(s) :")

            for _, row in df_show.iloc[::-1].iterrows():
                if str(row.get("Chantier", "")).startswith("2026-") or "nan" in str(row.get("Chantier", "")):
                    continue

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
                            <div style='color: #64748B; font-size: 13px; margin-top: 3px;'>
                                👷 {row.get("Effectif", "")}
                            </div>
                        </div>
                    """, unsafe_allow_html=True)

                    raw_photos = str(row.get("Photos", "")).replace(";", "|").replace(",", "|").split("|")
                    valid_paths = []
                    for ph in raw_photos:
                        ph_clean = ph.strip()
                        p_sub = os.path.join(PHOTOS_BASE_DIR, ph_clean)
                        p_root = os.path.join(PHOTOS_BASE_DIR, os.path.basename(ph_clean))
                        if os.path.exists(p_sub):
                            valid_paths.append(p_sub)
                        elif os.path.exists(p_root):
                            valid_paths.append(p_root)

                    if valid_paths:
                        cols = st.columns(len(valid_paths) if len(valid_paths) <= 3 else 3)
                        for i, p_img in enumerate(valid_paths):
                            try:
                                cols[i % 3].image(Image.open(p_img), use_container_width=True)
                            except Exception:
                                pass
                    st.write("")

        with st.expander("⚙️ Options avancées"):
            if st.button("🗑️ Réinitialiser le registre CSV"):
                if os.path.exists(CSV_FILE):
                    os.remove(CSV_FILE)
                    st.success("CSV nettoyé.")
                    st.rerun()

    elif pin != "":
        st.error("❌ Code secret incorrect.")
