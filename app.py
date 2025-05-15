import streamlit as st
import pandas as pd
from docx import Document
import os
import tempfile
from zipfile import ZipFile
import re
import random

st.set_page_config(page_title="Générateur de Questionnaires", layout="wide")
st.title("Générateur de Questionnaires de Satisfaction à Chaud")

# Colonnes requises dans Excel
REQUIRED_COLS = ['nom', 'prénom', 'email', 'session', 'formation']

# Fonction pour remplacer les placeholders
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

# Fonction spéciale pour traiter les blocs de satisfaction
def traiter_bloc_satisfaction(bloc_paras):
    """Traite un bloc de 5 checkboxs de satisfaction"""
    # Pour la question sur le handicap, toujours cocher "Non concerné"
    if bloc_paras[0].text.strip() == "Si vous êtes une personne en situation de handicap, êtes-vous satisfait de l’accompagnement et de l’adaptation éventuelle de la formation ? ":
        # On coche uniquement "Non concerné"
        bloc_paras[0].text = bloc_paras[0].text.replace("{{checkbox}}", "☑", 1)  # Non concerné
        # Et on laisse vide les autres options
        for i in range(1, len(bloc_paras)):
            bloc_paras[i].text = bloc_paras[i].text.replace("{{checkbox}}", "☐", 1)
        return

    # Pour les autres questions, choix aléatoire entre "Très satisfait" et "Satisfait"
    reponse_choisie = random.choice([0, 1])  # 0=Très satisfait, 1=Satisfait
    
    for i, para in enumerate(bloc_paras):
        if i == reponse_choisie:
            para.text = para.text.replace("{{checkbox}}", "☑", 1)
        else:
            para.text = para.text.replace("{{checkbox}}", "☐", 1)

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
        "{{formation}}": str(participant['formation']),
        "{{formateur}}": "Jean Dupont"  # À personnaliser selon vos besoins
    }

    # Remplacement des placeholders dans les paragraphes
    for para in doc.paragraphs:
        remplacer_placeholders(para, replacements)

    # Remplacement des placeholders dans les tableaux
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    remplacer_placeholders(para, replacements)

    # Traitement des blocs de satisfaction
    bloc_satisfaction = []
    
    for para in doc.paragraphs:
        texte = para.text.strip()
        
        # Détection des blocs de satisfaction
        if any(prefix in texte for prefix in [
            "Merci de nous partager votre évaluation de la formation :",
            "Qualité du Contenu de la Formation :",
            "Pertinence du Contenu par rapport à vos besoins :",
            "Clarté et Organisation du Contenu :",
            "Qualité des Supports de Formation (pdf, diapositives, etc.) :",
            "Utilité des Supports de Formation pour l'apprentissage :",
            "Compétence et Professionnalisme du Formateur :",
            "Clarté des explications du Formateur :",
            "Capacité du Formateur à répondre aux questions :",
            "Interactivité et Dynamisme du Formateur :",
            "Si vous êtes une personne en situation de handicap, êtes-vous satisfait de l’accompagnement et de l’adaptation éventuelle de la formation ? "
        ]):
            # Si on détecte un nouveau bloc, on traite l'ancien
            if bloc_satisfaction:
                traiter_bloc_satisfaction(bloc_satisfaction)
                bloc_satisfaction = []
        
        # Détection des checkboxs de satisfaction
        if "{{checkbox}}" in para.text:
            bloc_satisfaction.append(para)
    
    # Traitement du dernier bloc
    if bloc_satisfaction:
        traiter_bloc_satisfaction(bloc_satisfaction)

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
