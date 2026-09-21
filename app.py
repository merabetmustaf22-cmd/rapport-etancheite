import streamlit as st
import pandas as pd
from datetime import date
import os
import json
import io
import zipfile
from PIL import Image

# Configuration écran vertical smartphone
st.set_page_config(page_title="Chantier Mobile", page_icon="📱", layout="centered")

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

tab_saisie, tab_admin = st.tabs(["📲 Saisie Terrain", "🔒 Espace Admin"])

# -------------------------------------------------------------
# ONGLET 1 : INTERFACE MOBILE ULTRA-SIMPLE POUR LES MAÇONS
# -------------------------------------------------------------
with tab_saisie:
    st.markdown("""
        <div style='background-color: #0F766E; padding: 14px; border-radius: 12px; text-align: center; margin-bottom: 15px;'>
            <h2 style='color: white; margin: 0; font-size: 22px;'>📱 Pointage & Photos Chantier</h2>
            <p style='color: #CCFBF1; margin: 5px 0 0 0; font-size: 13px;'>Remplissez les champs et envoyez vos photos</p>
        </div>
    """, unsafe_allow_html=True)

    with st.form("form_mobile_chantier", clear_on_submit=True):
        # 1. Date & Chantier
        date_jour = st.date_input("📅 Date du jour", value=date.today())
        chantier_sel = st.selectbox("🏢 Chantier", config["chantiers"])

        # 2. Corps d'état
        tache_sel = st.selectbox("🛠️ Travail / Corps d'état", config["taches"])
        phase_travaux = st.selectbox("📌 Étape de réalisation", [
            "Pendant l'application / exécution",
            "Avant travaux (État du support)",
            "Après achèvement (Finition)",
            "Détail technique / Gorge / Relevé",
            "Épreuve d'eau (Test d'étanchéité)",
            "Autre"
        ])

        # 3. Ouvriers présents
        macons_presents = st.multiselect(
            "👷 Ouvriers présents sur ce travail",
            config["macons"],
            placeholder="Touchez pour choisir..."
        )

        # 4. Rendement
        c_r1, c_r2 = st.columns([2, 1])
        with c_r1:
            rendement = st.number_input("📏 Rendement réalisé", min_value=0.0, step=1.0, format="%.2f")
        with c_r2:
            unite = st.selectbox("Unité", ["m²", "ML", "U"])

        st.write("---")
        st.markdown("### 📸 Photos du travail")
        
        # Option 1 : Caméra directe du smartphone
        photo_camera = st.camera_input("📷 Prendre une photo en direct avec la caméra")

        # Option 2 : Galerie pour photos multiples
        photos_galerie = st.file_uploader(
            "📂 Ou importer des photos depuis la galerie",
            type=["jpg", "jpeg", "png"],
            accept_multiple_files=True
        )

        legende = st.text_input("💬 Remarque / Détail sur la photo", placeholder="Ex : Fin de première couche...")

        st.write("")
        submitted = st.form_submit_button("🚀 ENVOYER LE RAPPORT DU JOUR", use_container_width=True)

        if submitted:
            # Regroupement des photos reçues
            toutes_photos = []
            if photo_camera:
                toutes_photos.append(photo_camera)
            if photos_galerie:
                toutes_photos.extend(photos_galerie)

            if not toutes_photos:
                st.error("⚠️ Veuillez ajouter au moins une photo (caméra ou galerie).")
            elif not macons_presents:
                st.error("⚠️ Veuillez sélectionner au moins un ouvrier présent.")
            else:
                noms_sauvegardes = []
                chantier_clean = chantier_sel.split('(')[0].replace(" ", "_")
                tache_clean = tache_sel.replace(" ", "_")[:12]

                for idx, p in enumerate(toutes_photos):
                    extension = ".jpg"
                    if hasattr(p, "name") and os.path.splitext(p.name)[1]:
                        extension = os.path.splitext(p.name)[1].lower()

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

                st.success(f"✅ {len(noms_sauvegardes)} photo(s) et métré enregistrés avec succès !")

# -------------------------------------------------------------
# ONGLET 2 : ESPACE ADMIN (ADAPTÉ SMARTPHONE)
# -------------------------------------------------------------
with tab_admin:
    st.subheader("🔒 Espace Responsable")
    pin = st.text_input("Code Secret :", type="password")

    if pin == ADMIN_PIN:
        st.success("🔓 Accès déverrouillé.")

        st.markdown("### 📦 Téléchargements")
        # 1. Télécharger le ZIP photos
        if st.button("🗂️ Préparer l'archive ZIP des photos", use_container_width=True):
            fichiers_disponibles = os.listdir(PHOTOS_DIR)
            if fichiers_disponibles:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for photo in fichiers_disponibles:
                        full_path = os.path.join(PHOTOS_DIR, photo)
                        if os.path.isfile(full_path):
                            zip_file.write(full_path, arcname=photo)
                
                st.download_button(
                    label="📥 Télécharger le fichier ZIP",
                    data=zip_buffer.getvalue(),
                    file_name=f"photos_chantiers_{date.today()}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
            else:
                st.warning("Aucune photo disponible.")

        # 2. Télécharger le fichier Excel/CSV
        if os.path.exists(CSV_FILE):
            df_all = pd.read_csv(CSV_FILE, encoding='utf-8-sig')
            csv_bytes = df_all.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(
                "📥 Télécharger le registre (Excel/CSV)",
                csv_bytes,
                "registre_chantier.csv",
                "text/csv",
                use_container_width=True
            )

        st.write("---")
        st.markdown("### 🖼️ Dernières photos reçues")
        if os.path.exists(CSV_FILE):
            df_all = pd.read_csv(CSV_FILE, encoding='utf-8-sig')
            for _, row in df_all.tail(5).iloc[::-1].iterrows():
                st.write(f"**📍 {row['Chantier']} - {row['Corps_d_etat']}** ({row['Date']})")
                st.caption(f"Rendement : {row['Rendement']} {row['Unite']} | 👷 {row['Effectif']}")
                fichiers = str(row['Photos']).split(";")
                for f_name in fichiers:
                    chemin_f = os.path.join(PHOTOS_DIR, f_name)
                    if os.path.exists(chemin_f):
                        img = Image.open(chemin_f)
                        st.image(img, use_container_width=True)
                st.write("---")
        else:
            st.info("Aucune saisie pour le moment.")

    elif pin != "":
        st.error("❌ Code incorrect.")
