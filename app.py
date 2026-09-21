import streamlit as st
import pandas as pd
from datetime import date
import os
import json
import io
import zipfile
from PIL import Image

st.set_page_config(page_title="Suivi Photos & Rendement Chantier", page_icon="📸", layout="wide")

PHOTOS_DIR = "photos_chantier"
os.makedirs(PHOTOS_DIR, exist_ok=True)
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
        except:
            return CONFIG_DEFAUT
    return CONFIG_DEFAUT

def sauver_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

config = charger_config()

tab_saisie, tab_admin = st.tabs(["📸 Prise & Envoi des Photos", "🔒 Galerie & Archive Admin"])

# -------------------------------------------------------------
# ONGLET 1 : SAISIE TERRAIN ULTRA-CENTRÉE PHOTOS
# -------------------------------------------------------------
with tab_saisie:
    st.markdown("""
        <div style='background-color: #0F766E; padding: 12px; border-radius: 8px; text-align: center;'>
            <h2 style='color: white; margin: 0;'>📸 Journal Photos & Travaux de Chantier</h2>
            <p style='color: #CCFBF1; margin: 0; font-size: 14px;'>Prenez les photos de l'avancement, indiquez le travail et envoyez direct</p>
        </div>
    """, unsafe_allow_html=True)
    st.write("")

    with st.form("form_photos_chantier", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            date_jour = st.date_input("📅 Date de prise de vue", value=date.today())
            chantier_sel = st.selectbox("🏢 Chantier", config["chantiers"])
        with col2:
            tache_sel = st.selectbox("🛠️ Corps d'état concerné", config["taches"])
            phase_travaux = st.selectbox("📌 Étape de la photo", [
                "Pendant l'application / exécution",
                "Avant travaux (État du support)",
                "Après achèvement (Résultat final)",
                "Détail technique / Finition / Gorge",
                "Épreuve d'eau (Test d'étanchéité)",
                "Autre"
            ])

        col_m, col_r, col_u = st.columns([2, 1, 1])
        with col_m:
            macons_presents = st.multiselect("👷 Équipe présente", config["macons"])
        with col_r:
            rendement = st.number_input("📏 Rendement réalisé", min_value=0.0, step=1.0, format="%.2f")
        with col_u:
            unite = st.selectbox("Unité", ["m²", "ML", "U"])

        st.markdown("### 📷 Vos Photos")
        photos_uploaded = st.file_uploader(
            "Prenez les photos avec la caméra ou choisissez depuis la galerie (Plusieurs photos acceptées)",
            type=["jpg", "jpeg", "png"],
            accept_multiple_files=True
        )

        legende = st.text_input("💬 Légende / Observation sur la photo", placeholder="Ex : Première couche Elastotek sur terrasse A...")

        submitted = st.form_submit_button("🚀 Envoyer & Sauvegarder les Photos", use_container_width=True)

        if submitted:
            if not photos_uploaded:
                st.error("⚠️ Veuillez ajouter au moins une photo avant de valider.")
            elif not macons_presents:
                st.error("⚠️ Veuillez sélectionner au moins un ouvrier dans la liste.")
            else:
                noms_sauvegardes = []
                chantier_clean = chantier_sel.split('(')[0].replace(" ", "_")
                tache_clean = tache_sel.replace(" ", "_")[:12]

                for idx, p in enumerate(photos_uploaded):
                    extension = os.path.splitext(p.name)[1].lower()
                    if extension not in [".jpg", ".jpeg", ".png"]:
                        extension = ".jpg"
                    
                    nom_final = f"{date_jour}_{chantier_clean}_{tache_clean}_{idx+1}{extension}"
                    chemin_disque = os.path.join(PHOTOS_DIR, nom_final)

                    with open(chemin_disque, "wb") as f_img:
                        f_img.write(p.getbuffer())

                    noms_sauvegardes.append(nom_final)

                nouvelle_ligne = {
                    "Date": str(date_jour),
                    "Chantier": chantier_sel,
                    "Corps_d_etat": tache_sel,
                    "Phase": phase_travaux,
                    "Rendement": rendement,
                    "Unite": unite,
                    "Effectif": ", ".join(macons_presents),
                    "Nb_Ouvriers": len(macons_presents),
                    "Photos": ";".join(noms_sauvegardes),
                    "Nb_Photos": len(noms_sauvegardes),
                    "Legende": legende
                }

                df_entry = pd.DataFrame([nouvelle_ligne])
                if os.path.exists(CSV_FILE):
                    df_entry.to_csv(CSV_FILE, mode='a', header=False, index=False, encoding='utf-8-sig')
                else:
                    df_entry.to_csv(CSV_FILE, index=False, encoding='utf-8-sig')

                st.success(f"✅ {len(noms_sauvegardes)} Photo(s) envoyée(s) et classée(s) avec succès !")

# -------------------------------------------------------------
# ONGLET 2 : ESPACE ADMIN - GALERIE & TÉLÉCHARGEMENT ZIP
# -------------------------------------------------------------
with tab_admin:
    st.subheader("🔒 Espace Responsable")
    pin = st.text_input("Code Administrateur :", type="password")

    if pin == ADMIN_PIN:
        st.success("🔓 Accès administrateur accordé.")

        sous_onglets = st.tabs(["🖼️ Galerie Photos & Filtres", "📦 Télécharger ZIP Photos", "📋 Registre Chiffré", "⚙️ Configuration"])

        # 1. GALERIE PHOTOS
        with sous_onglets[0]:
            st.markdown("### 🖼️ Galerie Photos de Chantier")
            if os.path.exists(CSV_FILE):
                df_global = pd.read_csv(CSV_FILE, encoding='utf-8-sig')
                
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    filtre_chantier = st.selectbox("Filtrer par Chantier :", ["Tous"] + list(df_global["Chantier"].unique()))
                with col_f2:
                    filtre_tache = st.selectbox("Filtrer par Corps d'état :", ["Toutes"] + list(df_global["Corps_d_etat"].unique()))

                df_filtre = df_global.copy()
                if filtre_chantier != "Tous":
                    df_filtre = df_filtre[df_filtre["Chantier"] == filtre_chantier]
                if filtre_tache != "Toutes":
                    df_filtre = df_filtre[df_filtre["Corps_d_etat"] == filtre_tache]

                st.write(f"Affichage de **{len(df_filtre)}** enregistrements :")

                for _, row in df_filtre.iterrows():
                    with st.expander(f"📍 {row['Chantier']} - {row['Corps_d_etat']} ({row['Date']}) | Rendement: {row['Rendement']} {row['Unite']}", expanded=True):
                        st.caption(f"👷 Équipe: {row['Effectif']} | Étape: {row.get('Phase', '-')} | Observation: {row.get('Legende', '-')}")
                        
                        fichiers = str(row['Photos']).split(";")
                        cols = st.columns(min(len(fichiers), 3))
                        for i, f_name in enumerate(fichiers):
                            chemin_f = os.path.join(PHOTOS_DIR, f_name)
                            if os.path.exists(chemin_f):
                                try:
                                    img = Image.open(chemin_f)
                                    cols[i % 3].image(img, caption=f_name, use_container_width=True)
                                except:
                                    cols[i % 3].write(f"Image introuvable : {f_name}")
            else:
                st.info("Aucune photo enregistrée.")

        # 2. TÉLÉCHARGEMENT ZIP
        with sous_onglets[1]:
            st.markdown("### 📦 Télécharger toutes les photos d'un coup (Archive ZIP)")
            st.write("Récupérez un fichier compressé (.zip) avec toutes les photos classées pour préparer vos rapports.")

            if st.button("🗂️ Générer le fichier ZIP de toutes les photos"):
                fichiers_disponibles = os.listdir(PHOTOS_DIR)
                if fichiers_disponibles:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for photo in fichiers_disponibles:
                            full_path = os.path.join(PHOTOS_DIR, photo)
                            if os.path.isfile(full_path):
                                zip_file.write(full_path, arcname=photo)
                    
                    st.download_button(
                        label="📥 Télécharger le dossier ZIP des photos",
                        data=zip_buffer.getvalue(),
                        file_name=f"photos_chantiers_{date.today()}.zip",
                        mime="application/zip"
                    )
                else:
                    st.warning("Aucun fichier photo dans le dossier.")

        # 3. REGISTRE EXCEL
        with sous_onglets[2]:
            st.markdown("### 📋 Tableau de bord des rendements")
            if os.path.exists(CSV_FILE):
                df_all = pd.read_csv(CSV_FILE, encoding='utf-8-sig')
                st.dataframe(df_all, use_container_width=True)
                csv_bytes = df_all.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                st.download_button("📥 Télécharger le registre (Excel/CSV)", csv_bytes, "registre_chantier.csv", "text/csv")
            else:
                st.info("Aucune donnée.")

        # 4. CONFIGURATION
        with sous_onglets[3]:
            st.markdown("### ⚙️ Gestion des listes")
            st.write("**Chantiers :**", ", ".join(config["chantiers"]))
            n_c = st.text_input("Ajouter Chantier")
            if st.button("➕ Ajouter Chantier"):
                if n_c.strip() and n_c.strip() not in config["chantiers"]:
                    config["chantiers"].append(n_c.strip())
                    sauver_config(config)
                    st.rerun()

            st.write("**Ouvriers :**", ", ".join(config["macons"]))
            n_m = st.text_input("Ajouter Ouvrier")
            if st.button("➕ Ajouter Ouvrier"):
                if n_m.strip() and n_m.strip() not in config["macons"]:
                    config["macons"].append(n_m.strip())
                    sauver_config(config)
                    st.rerun()

    elif pin != "":
        st.error("❌ Code secret incorrect.")
    else:
        st.info("Saisissez le code pour accéder à la galerie et aux téléchargements.")
