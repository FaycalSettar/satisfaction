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
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Configuration de la page
st.set_page_config(page_title="Générateur de Questionnaires", layout="wide")
st.title("Générateur de Questionnaires de Satisfaction à Chaud")

REQUIRED_COLS = ['nom', 'prénom', 'email', 'session', 'formation', 'formateur']
SECTION_MARKERS = {
    'formation': '%%SECTION_FORMATION%%',
    'satisfaction': '%%SECTION_SATISFACTION%%',
    'handicap': '%%SECTION_HANDICAP%%'
}

def validate_data(df):
    """Valide l'intégrité des données du DataFrame"""
    if df.isna().any().any():
        missing = df.columns[df.isna().any()].tolist()
        st.error(f"❌ Données manquantes dans les colonnes : {', '.join(missing)}")
        return False
    return True

def remplacer_placeholders(doc, replacements):
    """Remplace les placeholders dans tout le document"""
    for p in doc.paragraphs:
        inline = p.runs
        for i in range(len(inline)):
            text = inline[i].text
            for key, value in replacements.items():
                if key in text:
                    text = text.replace(key, value)
                    inline[i].text = text

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                remplacer_placeholders(cell, replacements)

def generer_commentaire_ia(openrouter_api_key, formation):
    """Génère des commentaires IA avec parsing amélioré"""
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {openrouter_api_key}",
        "HTTP-Referer": "https://formation-entreprise.com",
        "X-Title": "Générateur Questionnaires",
        "Content-Type": "application/json"
    }
    
    prompt = """Génère 10 réponses courtes et variées pour des questionnaires de satisfaction.
    Format requis : 
    1. [Réponse 1]
    2. [Réponse 2]
    ...
    10. [Réponse 10]
    """
    
    data = {
        "model": "openai/gpt-4.1",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.7,
        "max_tokens": 300
    }
    
    try:
        response = requests.post(url, headers=headers, json=data, timeout=15)
        response.raise_for_status()
        raw = response.json()['choices'][0]['message']['content']
        
        # Extraction améliorée avec regex
        options = re.findall(r'\d+\.\s*(.+?)(?=\n\d+\.|\Z)', raw, re.DOTALL)
        return random.choice(options).strip() if options else ""
    except Exception as e:
        st.error(f"Erreur API IA : {str(e)[:200]}")
        return ""

def traiter_checkbox(paragraph, formation_choice, current_section):
    """Génère aléatoirement l'état des checkboxes par section"""
    if '{{checkbox}}' not in paragraph.text:
        return

    option_text = paragraph.text.replace('{{checkbox}}', '').strip()
    clean_option = re.sub(r'^\[.*?\]\s*', '', option_text).strip().lower()

    # Détermination dynamique de la réponse
    if current_section == 'formation':
        selected = (formation_choice == clean_option)
    elif current_section == 'satisfaction':
        reponses = ['très satisfait', 'satisfait', 'insatisfait', 'très insatisfait']
        selected = (random.choice(reponses) == clean_option)
    elif current_section == 'handicap':
        selected = ('non concerné' in clean_option)
    else:
        selected = False

    # Création d'un checkbox Word stylé
    checkbox = parse_xml(
        f'<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:sdtPr><w:checkbox>'
        f'<w:checked>{"1" if selected else "0"}</w:checked>' 
        f'</w:checkbox></w:sdtPr>'
        f'<w:sdtContent><w:r><w:t>{"☑" if selected else "☐"}</w:t></w:r></w:sdtContent>'
        f'</w:sdt>'
    )
    paragraph._element.clear_content()
    paragraph._element.append(checkbox)

def generer_questionnaire(participant, template_path, commentaire_ia=None):
    doc = Document(template_path)
    current_section = None
    formation_choice = str(participant['formation']).strip().lower()

    replacements = {
        "{{nom}}": participant['nom'],
        "{{prenom}}": participant['prénom'],
        "{{email}}": participant['email'],
        "{{ref_session}}": participant['session'],
        "{{formation}}": participant['formation'],
        "{{formateur}}": participant['formateur'],
        "{{commentaire_points_forts}}": commentaire_ia or "",  # Modification ici
    }

    remplacer_placeholders(doc, replacements)

    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        
        # Détection des sections via marqueurs
        if SECTION_MARKERS['formation'] in text:
            current_section = 'formation'
            paragraph.text = paragraph.text.replace(SECTION_MARKERS['formation'], '')
        elif SECTION_MARKERS['satisfaction'] in text:
            current_section = 'satisfaction'
            paragraph.text = paragraph.text.replace(SECTION_MARKERS['satisfaction'], '')
        elif SECTION_MARKERS['handicap'] in text:
            current_section = 'handicap'
            paragraph.text = paragraph.text.replace(SECTION_MARKERS['handicap'], '')
        
        # Traitement des checkboxes
        traiter_checkbox(paragraph, formation_choice, current_section)

    # Génération du nom de fichier sécurisé
    safe_name = re.sub(r'[^a-zA-Z0-9]+', '_', f"{participant['prénom']}_{participant['nom']}")
    filename = f"Questionnaire_{safe_name}_{participant['session']}.docx"
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
openrouter_api_key = st.text_input("Clé API OpenRouter", type="password") if generer_ia else ""

if excel_file and template_file:
    try:
        df = pd.read_excel(excel_file).convert_dtypes().dropna()
        
        if not all(col in df.columns for col in REQUIRED_COLS):
            missing = [col for col in REQUIRED_COLS if col not in df.columns]
            st.error(f"Colonnes requises manquantes : {', '.join(missing)}")
            st.stop()

        if not validate_data(df):
            st.stop()

        st.success(f"✅ {len(df)} participants validés")

        if st.button("Générer les questionnaires", type="primary"):
            with tempfile.TemporaryDirectory() as tmpdir:
                # Sauvegarde du template
                template_path = os.path.join(tmpdir, "template.docx")
                with open(template_path, "wb") as f:
                    template_file.seek(0)
                    f.write(template_file.read())

                # Préparation du ZIP
                zip_path = os.path.join(tmpdir, "Questionnaires.zip")
                total = len(df)
                
                with ZipFile(zip_path, 'w') as zipf:
                    progress_bar = st.progress(0)
                    
                    # Sélection aléatoire de 25% des participants
                    selected_indices = random.sample(range(total), k=max(1, total//4))
                    
                    for idx, (_, row) in enumerate(df.iterrows()):
                        try:
                            # Génération IA seulement pour les participants sélectionnés
                            if generer_ia and openrouter_api_key and idx in selected_indices:
                                commentaire = generer_commentaire_ia(openrouter_api_key, row['formation'])
                            else:
                                commentaire = None
                            
                            doc_path = generer_questionnaire(row, template_path, commentaire)
                            zipf.write(doc_path, os.path.basename(doc_path))
                            progress_bar.progress((idx+1)/total, text=f"Génération {idx+1}/{total}")
                        except Exception as e:
                            st.error(f"Erreur ligne {idx+1}: {str(e)[:200]}")
                            continue

                # Téléchargement
                with open(zip_path, "rb") as f:
                    st.balloons()
                    st.download_button(
                        "⬇️ Télécharger les questionnaires",
                        data=f,
                        file_name="Questionnaires.zip",
                        mime="application/zip"
                    )

    except Exception as e:
        st.error(f"Erreur critique : {str(e)[:200]}")
