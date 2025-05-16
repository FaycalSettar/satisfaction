import streamlit as st
import pandas as pd
from docx import Document  # Nécessite l'installation de python-docx
import os
import tempfile
from zipfile import ZipFile
import re
import random
import requests
import shutil

# Configuration de la page
st.set_page_config(page_title="Générateur de Questionnaires", layout="wide")
st.title("Générateur de Questionnaires de Satisfaction à Chaud")

# Colonnes requises dans le fichier Excel
REQUIRED_COLS = ['nom', 'prénom', 'email', 'session', 'formation']

def remplacer_placeholders(paragraph, replacements):
    """Remplace les placeholders dans un paragraphe Word"""
    if not paragraph.text:
        return
    for key, value in replacements.items():
        if key in paragraph.text:
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, value)

def generer_commentaire_ia(openrouter_api_key, formation="la formation"):
    """Génère plusieurs options de commentaires IA via OpenRouter et en renvoie une aléatoire"""
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {openrouter_api_key}",
        "HTTP-Referer": "https://formation-entreprise.com",
        "X-Title": "Générateur Questionnaires",
        "Content-Type": "application/json"
    }
    prompt = (
        f"""
        Génère 10 phrases courtes, variées et aléatoires décrivant les points forts de la formation en {formation},
        chacune sur une ligne, sans numérotation. Sois concis et humain. Réponse en quelques mots.
        """
    )
    data = {
        "model": "openai/gpt-4.1",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.8,
        "max_tokens": 150
    }
    try:
        resp = requests.post(url, headers=headers, json=data, timeout=10)
        resp.raise_for_status()
        raw = resp.json()['choices'][0]['message']['content'].strip()
        options = [l.strip() for l in raw.splitlines() if l.strip()]
        if options:
            return random.choice(options)
        return ""
    except Exception as e:
        st.warning(f"⚠️ Erreur API IA : {e}")
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
        if 'formation suivie' in text:
            current_section = 'formation'
            continue
        elif any(k in text for k in [
            'évaluation de la formation', 'qualité du contenu', 'pertinence du contenu',
            'clarté et organisation', 'qualité des supports', 'utilité des supports',
            'compétence et professionnalisme', 'clarté des explications',
            'capacité à répondre', 'interactivité et dynamisme', 'globalement'
        ]):
            current_section = 'satisfaction'
            answer = random.choice(['Très satisfait', 'Satisfait'])
            continue
        elif 'handicap' in text:
            current_section = 'handicap'
            answer = 'Non concerné'
            continue
        if '{{checkbox}}' in para.text:
            option_text = para.text.replace('{{checkbox}}', '').strip()
            clean_option = option_text.split(']')[-1].strip().lower()
            if current_section == 'formation':
                symbol = '☑' if formation_choice == clean_option else '☐'
            elif current_section == 'satisfaction':
                symbol = '☑' if answer.lower() == clean_option else '☐'
            elif current_section == 'handicap':
                symbol = '☑' if 'non concerné' in clean_option else '☐'
            else:
                symbol = '☐'
            original = option_text.split('[')[-1].split(']')[0].strip()
            para.text = f"{symbol} {original}"

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
    template_file = st.file_uploader("Modèle Word (.docx)", type="docx")

st.markdown("### Étape 2: Configuration IA")
generer_ia = st.checkbox("Activer la génération de commentaires IA (nécessite clé API)")
openrouter_api_key = ""
if generer_ia:
    openrouter_api_key = st.text_input("Clé API OpenRouter", type="password")

if excel_file and template_file:
    try:
        df = pd.read_excel(excel_file)
        if not all(c in df.columns for c in REQUIRED_COLS):
            st.error(f"❌ Colonnes requises manquantes : {', '.join(REQUIRED_COLS)}")
            st.stop()
        st.success(f"✅ {len(df)} participants détectés")
        if st.button("Générer les questionnaires", type="primary"):
            with tempfile.TemporaryDirectory() as tmpdir:
                tpl_path = os.path.join(tmpdir, "template.docx")
                with open(tpl_path, "wb") as f:
                    template_file.seek(0)
                    f.write(template_file.read())

                zip_path = os.path.join(tmpdir, "Questionnaires.zip")
                with ZipFile(zip_path, 'w') as zipf:
                    progress = st.progress(0)
                    total = len(df)
                    for idx, row in df.iterrows():
                        commentaire = None
                        if generer_ia and openrouter_api_key:
                            commentaire = generer_commentaire_ia(openrouter_api_key, row['formation'])
                        try:
                            out = generer_questionnaire(row, tpl_path, commentaire)
                            zipf.write(out, os.path.basename(out))
                            progress.progress((idx + 1) / total, text=f"{idx+1}/{total}")
                        except Exception as e:
                            st.error(f"❌ Échec pour {row['prénom']}: {e}")
                with open(zip_path, "rb") as f:
                    st.balloons()
                    st.download_button(
                        "⬇️ Télécharger les questionnaires",
                        data=f,
                        file_name="Questionnaires_Satisfaction.zip",
                        mime="application/zip"
                    )
    except Exception as e:
        st.error(f"🚨 Erreur critique : {e}")
