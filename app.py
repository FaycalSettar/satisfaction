import streamlit as st
import pandas as pd
from docx import Document
import os
import tempfile
from zipfile import ZipFile
import re
import random
import requests
import shutil


st.set_page_config(page_title="Générateur de Questionnaires", layout="wide")
st.title("Générateur de Questionnaires de Satisfaction à Chaud")


REQUIRED_COLS = ['nom', 'prénom', 'email', 'session', 'formation']

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

def generer_commentaire_ia(openrouter_api_key, formation="la formation"):
 
    url = "https://api.openrouter.ai/api/v1/chat/completions "
    headers = {
        "Authorization": f"Bearer {openrouter_api_key}",
        "HTTP-Referer": "https://formation-entreprise.com ",
        "X-Title": "Générateur Questionnaires",
        "Content-Type": "application/json"
    }
    
    prompt = (
        f"Génère 3 points forts très courts (3-5 mots chacun) pour une formation en {formation}, "
        "sans numérotation, séparés par des tirets. Exemple : "
        "\"explications claires - formation pratique - supports concrets\""
    )
    
    data = {
        "model": "openai/gpt-4",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.5,
        "max_tokens": 50
    }
    
    try:
        response = requests.post(url, headers=headers, json=data, timeout=10)
        response.raise_for_status()
        return response.json()['choices'][0]['message']['content'].strip()
    except Exception as e:
        st.error(f"Erreur API : {str(e)}")
        return ""

def generer_questionnaire(participant, template_path, commentaire_ia=None):
    doc = Document(template_path)

    replacements = {
        "{{nom}}": str(participant['nom']),
        "{{prenom}}": str(participant['prénom']),
        "{{email}}": str(participant['email']),
        "{{ref_session}}": str(participant['session']),
        "{{formation}}": str(participant['formation']),
        "{{formateur}}": "Jean Dupont",
        "{{commentaire_points_forts}}": commentaire_ia or "",
    }

    current_section = None
    formation_choice = str(participant['formation']).strip().lower()
    answer = None

    for para in doc.paragraphs:
        remplacer_placeholders(para, replacements)

        text = para.text.lower()
        
        # Détection des sections
        if 'formation suivie' in text:
            current_section = 'formation'
            continue
        elif any(keyword in text for keyword in [
            'évaluation de la formation', 
            'qualité du contenu',
            'pertinence du contenu',
            'clarté et organisation',
            'qualité des supports',
            'utilité des supports',
            'compétence et professionnalisme',
            'clarté des explications',
            'capacité à répondre',
            'interactivité et dynamisme',
            'globalement'
        ]):
            current_section = 'satisfaction'
            answer = random.choice(['Très satisfait', 'Satisfait'])
            continue
        elif 'handicap' in text:
            current_section = 'handicap'
            answer = 'Non concerné'
            continue

        # Traitement des checkboxes
        if '{{checkbox}}' in para.text:
            option_text = para.text.replace('{{checkbox}}', '').strip()
            clean_option = option_text.split(']')[-1].strip().lower()

            if current_section == 'formation':
                is_selected = formation_choice == clean_option
                symbol = ' ' if is_selected else '☐'
            elif current_section == 'satisfaction':
                is_selected = answer.lower() == clean_option
                symbol = ' ' if is_selected else '☐'
            elif current_section == 'handicap':
                is_selected = 'non concerné' in clean_option
                symbol = ' ' if is_selected else '☐'
            else:
                symbol = '☐'

            original_text = option_text.split('[')[-1].split(']')[0].strip()
            para.text = f'{symbol} {original_text}'

    # Nom du fichier
    safe_prenom = re.sub(r'[^a-zA-Z0-9]', '_', str(participant['prénom']))
    safe_nom = re.sub(r'[^a-zA-Z0-9]', '_', str(participant['nom']))
    filename = f"Questionnaire_{safe_prenom}_{safe_nom}_{participant['session']}.docx"
    
    output_path = os.path.join(tempfile.gettempdir(), filename)
    doc.save(output_path)
    
    return output_path

# Interface utilisateur
st.markdown("### Étape 1: Importation des fichiers")
col1, col2 = st.columns(2)

with col1:
    excel_file = st.file_uploader("Fichier Excel des participants", type="xlsx")
with col2:
    template_file = st.file_uploader("Modèle Word", type="docx")

st.markdown("### Étape 2: Configuration IA")
generer_ia = st.checkbox("Activer la génération de commentaires IA (nécessite clé API)")
openrouter_api_key = ""
if generer_ia:
    openrouter_api_key = st.text_input("Clé API OpenRouter", type="password")
    st.markdown("[Obtenir une clé API](https://openrouter.ai/keys )")

if excel_file and template_file:
    try:
        df = pd.read_excel(excel_file)

        if not all(col in df.columns for col in REQUIRED_COLS):
            st.error(f" Colonnes requises manquantes : {', '.join(REQUIRED_COLS)}")
            st.stop()

        st.success(f" {len(df)} participants détectés")

        if st.button("Générer les questionnaires", type="primary"):
            with tempfile.TemporaryDirectory() as tmpdir:
                template_path = os.path.join(tmpdir, "template.docx")
                with open(template_path, "wb") as f:
                    template_file.seek(0)
                    f.write(template_file.read())

                zip_path = os.path.join(tmpdir, "Questionnaires.zip")

                with ZipFile(zip_path, 'w') as zipf:
                    progress_bar = st.progress(0)
                    
                    for idx, row in df.iterrows():
                        commentaire = None
                        if generer_ia and openrouter_api_key:
                            try:
                                commentaire = generer_commentaire_ia(openrouter_api_key, row['formation'])
                            except Exception as e:
                                st.warning(f" Erreur IA pour {row['prénom']} : {str(e)}")
                        
                        try:
                            output_path = generer_questionnaire(row, template_path, commentaire)
                            zipf.write(output_path, os.path.basename(output_path))
                            progress_bar.progress((idx + 1)/len(df), text=f"Progress: {idx+1}/{len(df)}")
                        except Exception as e:
                            st.error(f" Échec génération {row['prénom']} : {str(e)}")
                            continue

                with open(zip_path, "rb") as f:
                    st.balloons()
                    st.download_button(
                        " Télécharger les questionnaires",
                        data=f,
                        file_name="Questionnaires_Satisfaction.zip",
                        mime="application/zip"
                    )

    except Exception as e:
        st.error(f" Erreur critique : {str(e)}")
