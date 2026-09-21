import streamlit as st
import pandas as pd
from datetime import date
import os
import json

st.set_page_config(page_title="Suivi Chantier & Rendement", page_icon="🏗️", layout="centered")

os.makedirs("photos_chantier", exist_ok=True)
CONFIG_FILE = "config_chantier.json"
CSV_FILE = "suivi_journalier_chantiers.csv"

# Configuration complète avec vos chantiers, maçons et corps d'état exacts
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
        "ADDA Abbess",
        "MEKHACHEF DJAMEL",
        "MESTEFAOUI AHMED",
        "ARGOUB HALIM",
        "FEHIM CHIBANI AZZOUZ",
        "TAIBI REDA",
        "ABED OMAR",
        "ZEGHDAN ABDELKADER",
        "BAGHDADI ALI",
        "BOUKHELIF KAMEL",
        "MOKHTARI Omar",
        "GHEZINI Habib",
        "BERACHEMI AHMED",
        "ARAR AISSA",
        "MOKHTARI Djelloul",
        "HAFDI Rachid",
        "BENHAMMADI Mohamed",
        "MOUISSI WALID",
        "BENSEMICHA MEROUANE",
        "TOUATI Zouaoui",
        "GHRIBI MOHAMED",
        "TAHAR BOUZIAN YOUCEF"
    ],
    "taches": [
        "BACHE A EAU",
        "BRICOL",
        "BRICOL ELASTOTEK",
        "BRICOL SILICONE",
        "BRICOL SOUS CARRELAGE",
        "coupe-feu",
        "Couvre-joint",
        "DALLE Cheminée",
        "ELASTOTEK",
        "Forme de pente",
        "GOURGE",
        "JOINT DE DILATATION",
        "PARE-VAPEUR",
        "PAX",
        "SOKLE PARE-VAPEUR",
        "SOUS CARRELAGE",
        "BRICOL PARE-VAPEUR",
        "ELASTOTEK SAUPOUDRAGE",
        "BRICOL Cheminée"
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

tab_saisie, tab_admin = st.tabs(["📲 Saisie Journalière (Chantier)", "⚙️ Configuration (Ajouter / Modifier)"])

# -------------------------------------------------------------
# ONGLET 1 : SAISIE TERRAIN (MAÇONS / APPLICATEURS)
# -------------------------------------------------------------
with tab_saisie:
    st.markdown("""
        <div style='background-color: #1E3A8A; padding: 10px; border-radius: 8px; text-align: center;'>
            <h3 style='color: white; margin: 0;'>🏗️ Pointage & Rendement Journalier</h3>
            <p style='color: #E2E8F0; margin: 0; font-size: 13px;'>Sélectionnez les options et enregistrez le travail du jour</p>
        </div>
    """, unsafe_allow_html=True)
    st.write("")

    with st.form("form_saisie_chantier", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            date_jour = st.date_input("📅 Date du jour", value=date.today())
        with col2:
            chantier_sel = st.selectbox("🏢 Chantier", config["chantiers"])

        tache_sel = st.selectbox("🛠️ Corps d'état / Travaux réalisés", config["taches"])

        macons_presents = st.multiselect(
            "👷 Effectif présent (Sélectionnez un ou plusieurs ouvriers)",
            config["macons"],
            help="Sélectionnez les personnes présentes sur ce travail"
        )

        col3, col4 = st.columns(2)
        with col3:
            rendement = st.number_input("📏 Rendement réalisé", min_value=0.0, step=0.5, format="%.2f")
        with col4:
            unite = st.selectbox("Unité de mesure", ["m² (Surface)", "ML (Mètre Linéaire)", "U (Unité / Socle)"])

        st.write("---")
        photos = st.file_uploader(
            "📸 Photos du travail réalisé (Caméra / Galerie)",
            type=["jpg", "jpeg", "png"],
            accept_multiple_files=True
        )

        observation = st.text_input("📝 Observation / Remarque éventuelle", placeholder="Ex : Zone A terminée, test d'étanchéité...")

        submitted = st.form_submit_button("💾 Valider & Enregistrer la Journée", use_container_width=True)

        if submitted:
            if not macons_presents:
                st.error("Veuillez sélectionner au moins un ouvrier dans la liste.")
            elif rendement <= 0:
                st.warning("Veuillez entrer un rendement supérieur à 0.")
            else:
                noms_photos = []
                if photos:
                    for p in photos:
                        nom_pic = f"{date_jour}_{chantier_sel.split('(')[0]}_{p.name}"
                        with open(os.path.join("photos_chantier", nom_pic), "wb") as f_out:
                            f_out.write(p.getbuffer())
                        noms_photos.append(nom_pic)

                ligne = {
                    "Date": str(date_jour),
                    "Chantier": chantier_sel,
                    "Corps_d_etat": tache_sel,
                    "Rendement": rendement,
                    "Unite": unite,
                    "Effectif_Presents": ", ".join(macons_presents),
                    "Nb_Ouvriers": len(macons_presents),
                    "Photos": ";".join(noms_photos),
                    "Observation": observation
                }

                df_new = pd.DataFrame([ligne])
                if os.path.exists(CSV_FILE):
                    df_new.to_csv(CSV_FILE, mode='a', header=False, index=False, encoding='utf-8-sig')
                else:
                    df_new.to_csv(CSV_FILE, index=False, encoding='utf-8-sig')

                st.success("✅ Journée enregistrée avec succès dans le registre central !")

    st.write("---")
    st.subheader("📋 Dernières saisies enregistrées")
    if os.path.exists(CSV_FILE):
        df_hist = pd.read_csv(CSV_FILE, encoding='utf-8-sig')
        st.dataframe(df_hist.tail(10), use_container_width=True)
        
        csv_data = df_hist.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        st.download_button(
            label="📥 Télécharger le registre cumulé (CSV / Excel)",
            data=csv_data,
            file_name="registre_journalier_chantiers.csv",
            mime="text/csv"
        )
    else:
        st.info("Aucune saisie enregistrée pour le moment.")

# -------------------------------------------------------------
# ONGLET 2 : GESTION DYNAMIQUE DES LISTES (AJOUT / SUPPRESSION)
# -------------------------------------------------------------
with tab_admin:
    st.subheader("⚙️ Modifier les listes déroulantes")
    st.caption("Ajoutez ou supprimez des ouvriers, chantiers ou corps d'état selon vos besoins.")

    # 1. Effectif
    st.markdown("#### 👷 Effectif Global")
    c_m1, c_m2 = st.columns([3, 1])
    with c_m1:
        nouveau_macon = st.text_input("Nom du nouvel ouvrier à ajouter", key="new_mac")
    with c_m2:
        st.write("")
        st.write("")
        if st.button("➕ Ajouter"):
            if nouveau_macon.strip() and nouveau_macon.strip() not in config["macons"]:
                config["macons"].append(nouveau_macon.strip())
                sauver_config(config)
                st.rerun()

    macon_a_suppr = st.selectbox("Supprimer un ouvrier :", ["---"] + config["macons"], key="del_mac")
    if st.button("🗑️ Supprimer cet ouvrier"):
        if macon_a_suppr != "---":
            config["macons"].remove(macon_a_suppr)
            sauver_config(config)
            st.rerun()

    st.write("---")

    # 2. Chantiers
    st.markdown("#### 🏢 Chantiers")
    c_c1, c_c2 = st.columns([3, 1])
    with c_c1:
        nouveau_chantier = st.text_input("Nouveau chantier à ajouter", key="new_ch")
    with c_c2:
        st.write("")
        st.write("")
        if st.button("➕ Ajouter Chantier"):
            if nouveau_chantier.strip() and nouveau_chantier.strip() not in config["chantiers"]:
                config["chantiers"].append(nouveau_chantier.strip())
                sauver_config(config)
                st.rerun()

    chantier_a_suppr = st.selectbox("Supprimer un chantier :", ["---"] + config["chantiers"], key="del_ch")
    if st.button("🗑️ Supprimer ce chantier"):
        if chantier_a_suppr != "---":
            config["chantiers"].remove(chantier_a_suppr)
            sauver_config(config)
            st.rerun()

    st.write("---")

    # 3. Corps d'état / Tâches
    st.markdown("#### 🛠️ Corps d'état & Tâches")
    c_t1, c_t2 = st.columns([3, 1])
    with c_t1:
        nouvelle_tache = st.text_input("Nouveau corps d'état à ajouter", key="new_tk")
    with c_t2:
        st.write("")
        st.write("")
        if st.button("➕ Ajouter Corps d'état"):
            if nouvelle_tache.strip() and nouvelle_tache.strip() not in config["taches"]:
                config["taches"].append(nouvelle_tache.strip())
                sauver_config(config)
                st.rerun()

    tache_a_suppr = st.selectbox("Supprimer un corps d'état :", ["---"] + config["taches"], key="del_tk")
    if st.button("🗑️ Supprimer ce corps d'état"):
        if tache_a_suppr != "---":
            config["taches"].remove(tache_a_suppr)
            sauver_config(config)
            st.rerun()
