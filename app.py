import streamlit as st
import pandas as pd
from datetime import date
import os
import json
import io
import zipfile
from PIL import Image

st.set_page_config(page_title="Chantier & Suivi", page_icon="📱", layout="centered")

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
                noms_sauvegardes = []
                chantier_clean = chantier_sel.split('(')[0].replace(" ", "_")
                tache_clean = tache_sel.replace(" ", "_")[:10]

                for idx, p in enumerate(photos_galerie):
                    ext = ".jpg"
                    if hasattr(p, "name") and os.path.splitext(p.name)[1]:
                        ext = os.path.splitext(p.name)[1].lower()

                    nom_final = f"{date_jour}_{chantier_clean}_{tache_clean}_{idx+1}{ext}"
                    chemin_disque = os.path.join(PHOTOS_DIR, nom_final)

                    with open(chemin_disque, "wb") as f_img:
                        f_img.write(p.getbuffer())

                    noms_sauvegardes.append(nom_final)

                nouvelle_ligne = {
                    "Date": str(date_jour),
                    "Chantier": str(chantier_sel),
                    "Corps_d_etat": str(tache_sel),
                    "Phase": str(phase_travaux),
                    "Rendement": str(rendement),
                    "Unite": str(unite),
                    "Effectif": ", ".join(macons_presents),
                    "Nb_Ouvriers": len(macons_presents),
                    "Photos": "|".join(noms_sauvegardes),
                    "Nb_Photos": len(noms_sauvegardes),
                    "Legende": str(legende).replace(";", " ").replace("|", " ")
                }

                df_entry = pd.DataFrame([nouvelle_ligne])
                colonnes_ordre = ["Date", "Chantier", "Corps_d_etat", "Phase", "Rendement", "Unite", "Effectif", "Nb_Ouvriers", "Photos", "Nb_Photos", "Legende"]
                df_entry = df_entry[colonnes_ordre]

                if os.path.exists(CSV_FILE):
                    df_entry.to_csv(CSV_FILE, sep=';', mode='a', header=False, index=False, encoding='utf-8-sig')
                else:
                    df_entry.to_csv(CSV_FILE, sep=';', index=False, encoding='utf-8-sig')

                st.success(f"✅ {len(noms_sauvegardes)} photo(s) et métré enregistrés avec succès !")

# -------------------------------------------------------------
# ONGLET 2 : TABLEAU DE BORD RESPONSABLE
# -------------------------------------------------------------
with tab_admin:
    st.markdown("""
        <div style='background-color: #1E293B; padding: 14px; border-radius: 12px; text-align: center; margin-bottom: 15px;'>
            <h2 style='color: white; margin: 0; font-size: 20px;'>📊 Tableau de Bord Chantier</h2>
            <p style='color: #94A3B8; margin: 4px 0 0 0; font-size: 13px;'>Suivi visuel des rendements et galerie photos</p>
        </div>
    """, unsafe_allow_html=True)

    pin = st.text_input("Code Administrateur :", type="password", placeholder="Entrez le code...")

    if pin == ADMIN_PIN:
        df_all = charger_donnees()
        total_photos = len(os.listdir(PHOTOS_DIR)) if os.path.exists(PHOTOS_DIR) else 0

        # --- CARTES INDICATEURS ---
        kpi1, kpi2, kpi3 = st.columns(3)
        with kpi1:
            total_m2 = 0.0
            if df_all is not None and "Rendement" in df_all.columns and "Unite" in df_all.columns:
                total_m2 = pd.to_numeric(df_all[df_all["Unite"].str.contains("m²", na=False)]["Rendement"], errors='coerce').sum()
            st.metric("Total Réalisé", f"{total_m2:.1f} m²")
        with kpi2:
            st.metric("Photos Reçues", f"{total_photos} 📸")
        with kpi3:
            total_rapports = len(df_all) if df_all is not None else 0
            st.metric("Rapports", f"{total_rapports}")

        st.write("")

        # --- EXPORT ---
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            fichiers_disponibles = os.listdir(PHOTOS_DIR) if os.path.exists(PHOTOS_DIR) else []
            if fichiers_disponibles:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for photo in fichiers_disponibles:
                        full_path = os.path.join(PHOTOS_DIR, photo)
                        if os.path.isfile(full_path):
                            zip_file.write(full_path, arcname=photo)
                
                st.download_button(
                    label="📦 Télécharger Photos (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"photos_chantiers_{date.today()}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
            else:
                st.button("📦 Photos (0)", disabled=True, use_container_width=True)

        with col_btn2:
            if df_all is not None and not df_all.empty:
                csv_bytes = df_all.to_csv(sep=';', index=False, encoding='utf-8-sig').encode('utf-8-sig')
                st.download_button(
                    label="📊 Télécharger Données (Excel)",
                    data=csv_bytes,
                    file_name="registre_chantier.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            else:
                st.button("📊 Excel (Vide)", disabled=True, use_container_width=True)

        st.write("---")

        # --- FILTRE & AFFICHAGE DES CARTES ---
        if df_all is not None and not df_all.empty and "Chantier" in df_all.columns:
            chantiers_bruts = [c for c in df_all["Chantier"].dropna().unique() if not str(c).startswith("2026-")]
            liste_chantiers = ["Tous les chantiers"] + list(chantiers_bruts)
            filtre_ch = st.selectbox("🔍 Filtrer par projet :", liste_chantiers)

            df_vue = df_all.copy()
            if filtre_ch != "Tous les chantiers":
                df_vue = df_vue[df_vue["Chantier"] == filtre_ch]

            st.write(f"Affichage de **{len(df_vue)}** fiche(s) :")

            for _, row in df_vue.iloc[::-1].iterrows():
                # Ignorer les lignes mal formées
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
                    photos_valides = [f for f in fichiers if os.path.exists(os.path.join(PHOTOS_DIR, f))]

                    if photos_valides:
                        cols = st.columns(len(photos_valides) if len(photos_valides) <= 3 else 3)
                        for i, f_name in enumerate(photos_valides):
                            chemin_f = os.path.join(PHOTOS_DIR, f_name)
                            img = Image.open(chemin_f)
                            cols[i % 3].image(img, use_container_width=True)

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
