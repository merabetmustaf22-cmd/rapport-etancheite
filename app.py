import streamlit as st
import pandas as pd
from datetime import date
import os
import json
import io
import zipfile
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

def nettoyer_nom_dossier(nom):
    return nom.split('(')[0].replace(" ", "_").replace("/", "_").strip()

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
                dossier_chantier = nettoyer_nom_dossier(chantier_sel)
                chemin_dossier = os.path.join(PHOTOS_BASE_DIR, dossier_chantier)
                os.makedirs(chemin_dossier, exist_ok=True)

                noms_sauvegardes_relatifs = []
                tache_clean = tache_sel.replace(" ", "_")[:10]

                for idx, p in enumerate(photos_galerie):
                    ext = ".jpg"
                    if hasattr(p, "name") and os.path.splitext(p.name)[1]:
                        ext = os.path.splitext(p.name)[1].lower()

                    nom_fichier = f"{date_jour}_{tache_clean}_{idx+1}{ext}"
                    chemin_fichier_complet = os.path.join(chemin_dossier, nom_fichier)

                    with open(chemin_fichier_complet, "wb") as f_img:
                        f_img.write(p.getbuffer())

                    noms_sauvegardes_relatifs.append(f"{dossier_chantier}/{nom_fichier}")

                nouvelle_ligne = {
                    "Date": str(date_jour),
                    "Chantier": str(chantier_sel),
                    "Corps_d_etat": str(tache_sel),
                    "Phase": str(phase_travaux),
                    "Rendement": str(rendement),
                    "Unite": str(unite),
                    "Effectif": ", ".join(macons_presents),
                    "Nb_Ouvriers": len(macons_presents),
                    "Photos": "|".join(noms_sauvegardes_relatifs),
                    "Nb_Photos": len(noms_sauvegardes_relatifs),
                    "Legende": str(legende).replace(";", " ").replace("|", " ")
                }

                df_entry = pd.DataFrame([nouvelle_ligne])
                colonnes_ordre = ["Date", "Chantier", "Corps_d_etat", "Phase", "Rendement", "Unite", "Effectif", "Nb_Ouvriers", "Photos", "Nb_Photos", "Legende"]
                df_entry = df_entry[colonnes_ordre]

                if os.path.exists(CSV_FILE):
                    df_entry.to_csv(CSV_FILE, sep=';', mode='a', header=False, index=False, encoding='utf-8-sig')
                else:
                    df_entry.to_csv(CSV_FILE, sep=';', index=False, encoding='utf-8-sig')

                st.success(f"✅ {len(noms_sauvegardes_relatifs)} photo(s) classée(s) dans le dossier [{dossier_chantier}] !")

# -------------------------------------------------------------
# ONGLET 2 : TABLEAU DE BORD RESPONSABLE
# -------------------------------------------------------------
with tab_admin:
    st.markdown("""
        <div style='background-color: #1E293B; padding: 14px; border-radius: 12px; text-align: center; margin-bottom: 15px;'>
            <h2 style='color: white; margin: 0; font-size: 20px;'>📊 Tableau de Bord Chantier</h2>
            <p style='color: #94A3B8; margin: 4px 0 0 0; font-size: 13px;'>Suivi visuel et dossiers photos par chantier</p>
        </div>
    """, unsafe_allow_html=True)

    pin = st.text_input("Code Administrateur :", type="password", placeholder="Entrez le code...")

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
            st.metric("Rapports", f"{total_rapports}")

        st.write("---")

        # --- TÉLÉCHARGEMENT ZIP PAR CHANTIER OU GLOBAL ---
        st.markdown("### 📦 Téléchargements Photos & Données")
        dossiers_existants = [d for d in os.listdir(PHOTOS_BASE_DIR) if os.path.isdir(os.path.join(PHOTOS_BASE_DIR, d))]
        
        col_sel_ch, col_zip_ch = st.columns([2, 1])
        with col_sel_ch:
            option_chantier_zip = st.selectbox("Sélectionnez le chantier à télécharger :", ["Tous les chantiers (Dossiers séparés)"] + dossiers_existants)

        with col_zip_ch:
            st.write("")
            st.write("")
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                if option_chantier_zip == "Tous les chantiers (Dossiers séparés)":
                    for root, _, files in os.walk(PHOTOS_BASE_DIR):
                        for f in files:
                            chemin_complet = os.path.join(root, f)
                            arcname = os.path.relpath(chemin_complet, PHOTOS_BASE_DIR)
                            zip_file.write(chemin_complet, arcname=arcname)
                    nom_zip = f"tous_chantiers_photos_{date.today()}.zip"
                else:
                    dossier_cible = os.path.join(PHOTOS_BASE_DIR, option_chantier_zip)
                    for f in os.listdir(dossier_cible):
                        chemin_f = os.path.join(dossier_cible, f)
                        if os.path.isfile(chemin_f):
                            zip_file.write(chemin_f, arcname=f"{option_chantier_zip}/{f}")
                    nom_zip = f"photos_{option_chantier_zip}_{date.today()}.zip"

            st.download_button(
                label="📥 Télécharger ZIP",
                data=zip_buffer.getvalue(),
                file_name=nom_zip,
                mime="application/zip",
                use_container_width=True
            )

        if df_all is not None and not df_all.empty:
            csv_bytes = df_all.to_csv(sep=';', index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(
                label="📊 Télécharger Registre (Excel/CSV)",
                data=csv_bytes,
                file_name="registre_chantier.csv",
                mime="text/csv",
                use_container_width=True
            )

        st.write("---")

        # --- FILTRE & AFFICHAGE DES CARTES ET PHOTOS ---
        if df_all is not None and not df_all.empty and "Chantier" in df_all.columns:
            chantiers_bruts = [c for c in df_all["Chantier"].dropna().unique() if not str(c).startswith("2026-")]
            liste_chantiers = ["Tous les chantiers"] + list(chantiers_bruts)
            filtre_ch = st.selectbox("🔍 Filtrer par projet :", liste_chantiers)

            df_vue = df_all.copy()
            if filtre_ch != "Tous les chantiers":
                df_vue = df_vue[df_vue["Chantier"] == filtre_ch]

            st.write(f"Affichage de **{len(df_vue)}** fiche(s) :")

            for _, row in df_vue.iloc[::-1].iterrows():
                if str(row.get("Chantier", "")).startswith("2026-") or "nan" in str(row.get("Chantier", "")):
                    continue

                with st.container():
                    st.markdown(f"""
                        <div style='background-color: #F8FAFC; border-left: 5px solid #0F766E; padding: 12px; border-radius: 8px; margin-bottom: 10px; border-top: 1px solid #E2E8F0; border-right: 1px solid #E2E8F0; border-bottom: 1px solid #E2E8F0;'>
                            <div style='display: flex; justify-content: space-between; align-items: center;'>
                                <span style='font-weight: 700; color: #0F172A; font-size: 15px;'>🏢 {row.get("Chantier", "")}</span>
                                <span style='background-color: #E2E8F0; padding: 2px 8px; border-radius: 12px; font-size: 12px; color: #475569;'>📅 {row.get("Date", "")}</span>
                            </div>
                            <div style='margin-top: 6px; font-size: 14px; color: #334155;'>
                                <b>Travaux :</b> <span style='color: #0F766E;'>{row.get("Corps_d_etat", "")}</span> &nbsp;|&nbsp; 
                                <b>Rendement :</b> <b>{row.get("Rendement", "")} {row.get("Unite", "")}</b>
                            </div>
                            <div style='margin-top: 4px; font-size: 13px; color: #64748B;'>
                                👷 <i>{row.get("Effectif", "")}</i>
                            </div>
                            {"<div style='margin-top: 4px; font-size: 13px; color: #0284C7;'>💬 " + str(row.get("Legende")) + "</div>" if pd.notna(row.get("Legende")) and str(row.get("Legende")).strip() and str(row.get("Legende")) != "nan" else ""}
                        </div>
                    """, unsafe_allow_html=True)

                    photos_raw = str(row.get("Photos", ""))
                    fichiers = [f.strip() for f in photos_raw.replace(",", "|").replace(";", "|").split("|") if f.strip()]
                    
                    chemins_trouves = []
                    for f in fichiers:
                        p1 = os.path.join(PHOTOS_BASE_DIR, f)
                        p2 = os.path.join(PHOTOS_BASE_DIR, os.path.basename(f))
                        if os.path.exists(p1):
                            chemins_trouves.append(p1)
                        elif os.path.exists(p2):
                            chemins_trouves.append(p2)

                    if chemins_trouves:
                        cols = st.columns(len(chemins_trouves) if len(chemins_trouves) <= 3 else 3)
                        for i, chemin_img in enumerate(chemins_trouves):
                            try:
                                img = Image.open(chemin_img)
                                cols[i % 3].image(img, use_container_width=True)
                            except Exception:
                                pass

                    st.write("")
        else:
            st.info("Aucune intervention enregistrée pour l'instant.")

        # --- PARAMÈTRES EN BAS ---
        with st.expander("⚙️ Paramètres (Modifier les chantiers ou les maçons)"):
            st.markdown("##### 👷 Ouvriers enregistrés")
            st.caption(", ".join(config["macons"]))
            c_m1, c_m2 = st.columns([3, 1])
            with c_m1:
                n_mac = st.text_input("Nouvel ouvrier", key="add_m")
            with c_m2:
                st.write("")
                st.write("")
                if st.button("➕ Ajouter"):
                    if n_mac.strip() and n_mac.strip() not in config["macons"]:
                        config["macons"].append(n_mac.strip())
                        sauver_config(config)
                        st.rerun()

            st.write("---")
            if st.button("🗑️ Réinitialiser le registre CSV si besoin"):
                if os.path.exists(CSV_FILE):
                    os.remove(CSV_FILE)
                    st.success("Fichier CSV nettoyé.")
                    st.rerun()

    elif pin != "":
        st.error("❌ Code secret incorrect.")
