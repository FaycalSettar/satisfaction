import streamlit as st
import pandas as pd
from docx import Document
import os
import tempfile
from zipfile import ZipFile
import re  # Pour les expressions régulières

st.set_page_config(page_title="Générateur de Questionnaires", layout="wide")
st.title("Générateur de Questionnaires de Satisfaction à Chaud")

# Configuration des colonnes requises
REQUIRED_COLS = ['nom', 'prénom', 'email', 'session', 'formation']

# Fonction de remplacement des placeholders
def remplacer_placeholders(paragraph, replacements):
    """Remplace les placeholders dans un paragraphe Word"""
    if not paragraph.text:
        return
    
    original_text = paragraph.text
    for key, value in replacements.items():
        if key in original_text:
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, value)

# Fonction de génération d'un questionnaire
def generer_questionnaire(participant, template_path):
    """Génère un questionnaire personnalisé pour un participant"""
    doc = Document(template_path)
    
    # Préparation des remplacements
    replacements = {
        "{{nom}}": str(participant['nom']),
        "{{prenom}}": str(participant['prénom']),
        "{{email}}": str(participant['email']),
        "{{ref_session}}": str(participant['session']),
        "{{formation}}": str(participant['formation']),  # ✅ Ajout de la variable formation
        "{{formateur}}": "Jean Dupont"  # À personnaliser selon vos besoins
    }

    # Remplacement dans les paragraphes
    for para in doc.paragraphs:
        remplacer_placeholders(para, replacements)

    # Remplacement dans les tableaux
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    remplacer_placeholders(para, replacements)

    # Génération du nom de fichier
    safe_prenom = re.sub(r'[^a-zA-Z0-9]', '_', str(participant['prénom']))
    safe_nom = re.sub(r'[^a-zA-Z0-9]', '_', str(participant['nom']))
    filename = f"Questionnaire_{safe_prenom}_{safe_nom}_{participant['session']}.docx"
    
    # Sauvegarde temporaire
    output_path = os.path.join(tempfile.gettempdir(), filename)
    doc.save(output_path)
    
    return output_path

# Interface utilisateur
st.markdown("### Étape 1: Importation des fichiers")
col1, col2 = st.columns(2)

with col1:
    excel_file = st.file_uploader("Fichier Excel des participants", type="xlsx")
    
with col2:
    template_file = st.file_uploader("Modèle Word (Questionnaire de satisfaction)", type="docx")

# Traitement des données
if excel_file and template_file:
    try:
        df = pd.read_excel(excel_file)

        if not all(col in df.columns for col in REQUIRED_COLS):
            st.error("❌ Le fichier Excel doit contenir toutes les colonnes suivantes : " + ", ".join(REQUIRED_COLS))
            st.stop()

        st.info(f"✅ {len(df)} participants trouvés dans le fichier Excel")

        if st.button("Générer les questionnaires", type="primary"):
            with tempfile.TemporaryDirectory() as tmpdir:
                template_path = os.path.join(tmpdir, "template.docx")
                with open(template_path, "wb") as f:
                    f.write(template_file.getbuffer())

                zip_path = os.path.join(tmpdir, "Questionnaires.zip")

                with ZipFile(zip_path, 'w') as zipf:
                    progress_bar = st.progress(0)

                    for idx, row in df.iterrows():
                        try:
                            output_path = generer_questionnaire(row, template_path)
                            zipf.write(output_path, os.path.basename(output_path))
                            progress_bar.progress((idx + 1)/len(df), text=f"Génération en cours : {idx+1}/{len(df)}")
                        except Exception as e:
                            st.warning(f"⚠️ Échec pour {row['prénom']} {row['nom']} : {str(e)}")
                            continue

                with open(zip_path, "rb") as f:
                    st.success("✅ Génération terminée avec succès !")
                    st.download_button(
                        "📥 Télécharger les questionnaires",
                        data=f,
                        file_name="Questionnaires_Satisfaction.zip",
                        mime="application/zip"
                    )

    except Exception as e:
        st.error(f"❌ Erreur lors de la génération : {str(e)}")
