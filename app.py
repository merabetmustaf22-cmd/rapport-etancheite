import streamlit as st
import pandas as pd
import openpyxl
from PIL import Image, ImageOps
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import io

st.set_page_config(page_title="Générateur de Rapports BTPH", layout="centered")

st.title("Générateur de Rapport Mensuel d'Étanchéité")
st.write("Application mobile pour générer automatiquement vos rapports Word (.docx)")

# 1. Sélection des fichiers
st.subheader("1. Fichiers Excel")
pointage_file = st.file_uploader("Fichier de Pointage (Excel)", type=["xlsx"])
conso_file = st.file_uploader("Fichier de Consommations / Sorties (Excel)", type=["xlsx"])

st.subheader("2. Photos de Chantier")
photos = st.file_uploader("Sélectionner les photos (depuis la galerie ou l'appareil)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

# 2. Bouton de génération
if st.button("Générer le Rapport Word"):
    if not photos:
        st.warning("Veuillez ajouter au moins une photo pour le rapport.")
    else:
        st.success(f"{len(photos)} photos prêtes. Traitement en cours...")
        
        # Création du document Word en mémoire
        doc = docx.Document()
        
        # En-tête
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("RAPPORT TECHNIQUE MENSUEL D'ACTIVITÉ")
        r.font.size = Pt(20)
        r.font.bold = True
        r.font.color.rgb = RGBColor(30, 58, 138)
        
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_sub = p_sub.add_run("TRAVAUX D'ÉTANCHÉITÉ & PRÉPARATION DES SUPPORTS")
        r_sub.font.size = Pt(11)
        r_sub.font.bold = True
        
        # Ajout des photos formatées
        doc.add_heading("Reportage Photographique de Chantier", level=1)
        
        for idx, photo in enumerate(photos, start=1):
            img = Image.open(photo).convert('RGB')
            img_fit = ImageOps.fit(img, (800, 600), method=Image.Resampling.LANCZOS)
            img_b1 = ImageOps.expand(img_fit, border=6, fill=(241, 245, 249))
            img_final = ImageOps.expand(img_b1, border=2, fill=(71, 85, 105))
            
            img_byte_arr = io.BytesIO()
            img_final.save(img_byte_arr, format='JPEG', quality=95)
            img_byte_arr.seek(0)
            
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.add_run().add_picture(img_byte_arr, width=Inches(4.5))
            
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_cap = p_cap.add_run(f"Photo {idx} : État d'avancement des travaux sur chantier")
            r_cap.font.size = Pt(9)
            r_cap.font.italic = True
        
        # Sauvegarde en mémoire pour téléchargement direct
        docx_io = io.BytesIO()
        doc.save(docx_io)
        docx_io.seek(0)
        
        st.download_button(
            label="Télécharger le Rapport Word (.docx)",
            data=docx_io,
            file_name="Rapport_Mensuel_Genere.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
