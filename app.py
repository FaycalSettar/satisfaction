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
REQUIRED_COLS = ['nom', 'prénom', 'email', 'session', 'formation', 'formateur']

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
    """Génère un commentaire IA via OpenRouter, en choisissant aléatoirement parmi plusieurs options"""
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {openrouter_api_key}",
        "HTTP-Referer": "https://formation-entreprise.com",
        "X-Title": "Générateur Questionnaires",
        "Content-Type": "application/json"
    }
    prompt = (
        f"""
        ne commence pas ta phrase toujours avec la même accroche propose des réponses avec des phrases plus complète et soit le plus humain possible tu es un apprenant qui vient de réaliser une formation en {formation} pour décrire ton ressenti concernant les points forts de cette formation voici quelques exemples inspire toi dessus et change toujours l'accroche et le sens de la première proposition commence ta phrase directement sans chiffre ou caractère et soit le plus aléatoire sur la première proposition
        1-Explications claires et outils
        2-Formation pratico pratique. On en ressort avec un système en place qui fonctionne
        3-Une formation vraiment au top, je suis ressorti avec pleins de tips
        4-Le contenu, les supports
        5-Le formateur est très pédagogue et maîtrise parfaitement le sujet. Le fait d'être en petit comité est très appréciable.
        6-Ouvert à tous et simple d’utilisation. Résultats concrets
        7-La recherche Boléenne
        8-Les cours qui sont sous format numérique et interactif que l'on peut consulter à la demande.
        9-formateur pédagogue prends son temps
        10-gestion de dossier admin tout est ok en plus de la formation
        réponse en quelques mots
        """
    )
    data = {
        "model": "openai/gpt-4.1",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.5,
        "max_tokens": 100
    }
    try:
        response = requests.post(url, headers=headers, json=data, timeout=10)
        response.raise_for_status()
        raw = response.json()['choices'][0]['message']['content'].strip()
        options = [ligne.strip() for ligne in raw.splitlines() if ligne.strip()]
        return random.choice(options) if options else ""
    except Exception as e:
        st.error(f"Erreur API IA : {e}")
        return ""


def generer_questionnaire(participant, template_path, commentaire_ia=None, commentaire_remarques=None):
    doc = Document(template_path)
    replacements = {
        "{{nom}}": str(participant['nom']),
        "{{prenom}}": str(participant['prénom']),
        "{{email}}": str(participant['email']),
        "{{ref_session}}": str(participant['session']),
        "{{formation}}": str(participant['formation']),
        "{{formateur}}": str(participant['formateur']),
        "{{commentaire_points_forts}}": commentaire_ia or "",
        "{{commentaire_remarques}}": commentaire_remarques or "",
    }
    for para in doc.paragraphs:
        remplacer_placeholders(para, replacements)
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
if generer_ia:
    openrouter_api_key = st.text_input("Clé API OpenRouter", type="password")
else:
    openrouter_api_key = ""

if excel_file and template_file:
    try:
        df = pd.read_excel(excel_file)
        if not all(col in df.columns for col in REQUIRED_COLS):
            st.error(f"❌ Colonnes requises manquantes : {', '.join(REQUIRED_COLS)}")
            st.stop()
        st.success(f"✅ {len(df)} participants détectés")

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
                        remarque = None
                        # Générer un commentaire IA et une remarque pour 1 participant sur 4
                        if generer_ia and openrouter_api_key and idx % 4 == 0:
                            commentaire = generer_commentaire_ia(openrouter_api_key, row['formation'])
                            remarque = generer_commentaire_ia(openrouter_api_key, row['formation'])
                        try:
                            output_path = generer_questionnaire(row, template_path, commentaire, remarque)
                            zipf.write(output_path, os.path.basename(output_path))
                            progress_bar.progress((idx + 1) / len(df), text=f"Progress: {idx+1}/{len(df)}")
                        except Exception as e:
                            st.error(f"❌ Échec génération {row['prénom']} : {str(e)}")
                            continue

                with open(zip_path, "rb") as f:
                    st.balloons()
                    st.download_button(
                        "⬇️ Télécharger les questionnaires",
                        data=f,
                        file_name="Questionnaires_Satisfaction.zip",
                        mime="application/zip"
                    )
    except Exception as e:
        st.error(f"🚨 Erreur critique : {str(e)}")
