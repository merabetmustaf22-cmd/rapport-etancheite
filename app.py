import streamlit as st
import pandas as pd
from datetime import date
from PIL import Image
import os

st.set_page_config(page_title="Pointage & Rendement Chantier", page_icon="🏗️", layout="centered")

# Dossier ta3 les photos w fichier CSV
os.makedirs("photos_chantier", exist_ok=True)
csv_file = "suivi_journalier_chantiers.csv"

st.markdown("""
    <div style='background-color: #1E3A8A; padding: 12px; border-radius: 8px; text-align: center;'>
        <h2 style='color: white; margin: 0;'>🏗️ Saisie Journalière Chantier</h2>
        <p style='color: #E2E8F0; margin: 0; font-size: 14px;'>Enregistrement du rendement, effectif et photos</p>
    </div>
""", unsafe_allow_html=True)

st.write("")

with st.form("form_saisie_journaliere", clear_on_submit=True):
    col1, col2 = st.columns(2)
    with col1:
        date_jour = st.date_input("📅 Date d'intervention", value=date.today())
    with col2:
        chantier = st.selectbox(
            "🏢 Chantier / Projet",
            ["IZEM", "CAC-31-24", "HOB-342-08-24", "Autre"]
        )
    
    if chantier == "Autre":
        chantier = st.text_input("Préciser le nom du chantier :")

    tache = st.selectbox(
        "🛠️ Nature des travaux / Tâche",
        [
            "Application Elastotek (Étanchéité liquide)",
            "Pose Pare-vapeur / Écran d'indépendance",
            "Étanchéité sous-carrelage (Salles humides)",
            "Surfaçage / Ponçage / Décapage support",
            "Réalisation gorges d'angle / Relevés d'étanchéité",
            "Calfeutrement joint coupe-feu",
            "Épreuve d'eau (Mise en eau 24h/72h)",
            "Autre intervention"
        ]
    )

    col3, col4 = st.columns(2)
    with col3:
        rendement = st.number_input("📏 Rendement réalisé", min_value=0.0, step=1.0, format="%.2f")
    with col4:
        unite = st.selectbox("Unité", ["m² (Surfacique)", "ML (Mètre Linéaire)", "U (Unité)"])

    col5, col6 = st.columns(2)
    with col5:
        effectif = st.number_input("👷 Nb Ouvriers sur site", min_value=1, max_value=50, value=2, step=1)
    with col6:
        chef_equipe = st.text_input("Nom du maçon / Chef d'équipe")

    st.write("---")
    st.subheader("📸 Photos du travail réalisé")
    photos = st.file_uploader(
        "Prendre une photo avec l'appareil ou choisir dans la galerie",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True
    )

    observation = st.text_area("📝 Remarques / Observations de la journée", placeholder="Ex : Retard dû au vent, zone A terminée, test d'eau sans fuite...")

    submitted = st.form_submit_button("💾 Enregistrer la journée de travail", use_container_width=True)

    if submitted:
        if not chef_equipe:
            st.error("Veuillez renseigner le nom de la personne qui saisit.")
        elif rendement <= 0:
            st.warning("Veuillez saisir un rendement valide supérieur à 0.")
        else:
            noms_photos = []
            if photos:
                for p in photos:
                    nom_fichier = f"{date_jour}_{chantier}_{p.name}"
                    chemin = os.path.join("photos_chantier", nom_fichier)
                    with open(chemin, "wb") as f:
                        f.write(p.getbuffer())
                    noms_photos.append(nom_fichier)

            nouvelle_ligne = {
                "Date": str(date_jour),
                "Chantier": chantier,
                "Tache": tache,
                "Rendement": rendement,
                "Unite": unite,
                "Effectif": effectif,
                "Chef_Equipe": chef_equipe,
                "Photos": ";".join(noms_photos),
                "Observation": observation
            }

            df_new = pd.DataFrame([nouvelle_ligne])
            if os.path.exists(csv_file):
                df_new.to_csv(csv_file, mode='a', header=False, index=False, encoding='utf-8-sig')
            else:
                df_new.to_csv(csv_file, index=False, encoding='utf-8-sig')

            st.success("✅ Données et photos de la journée enregistrées avec succès !")

st.write("---")
st.subheader("📋 Historique des dernières saisies sur chantier")

if os.path.exists(csv_file):
    df_historique = pd.read_csv(csv_file, encoding='utf-8-sig')
    st.dataframe(df_historique.tail(10), use_container_width=True)
    
    csv_data = df_historique.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
    st.download_button(
        label="📥 Télécharger le registre cumulé (Excel/CSV)",
        data=csv_data,
        file_name="registre_journalier_chantiers.csv",
        mime="text/csv"
    )
else:
    st.info("Aucune saisie enregistrée pour le moment. Remplissez le formulaire ci-dessus.")
