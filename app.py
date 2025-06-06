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

REQUIRED_COLS = ['nom', 'prénom', 'email', 'session', 'formation', 'formateur']

def remplacer_placeholders(paragraph, replacements):
    """Version corrigée avec gestion des styles de caractères"""
    if not paragraph.text:
        return
    
    # 1. Fusionner tous les runs du paragraphe en conservant les styles
    full_text = ''.join(run.text for run in paragraph.runs)
    
    # 2. Effectuer tous les remplacements sur le texte complet
    for key, value in replacements.items():
        if key in full_text:
            full_text = full_text.replace(key, value)
    
    # 3. Sauvegarder les propriétés de style du premier run
    if paragraph.runs:
        first_run = paragraph.runs[0]
        font = first_run.font
        saved_style = {
            'bold': font.bold,
            'italic': font.italic,
            'underline': font.underline,
            'size': font.size,
            'color': font.color.rgb if font.color else None,
            'name': font.name
        }
    else:
        saved_style = None
    
    # 4. Effacer le paragraphe et recréer le texte
    paragraph.clear()
    new_run = paragraph.add_run(full_text)
    
    # 5. Restaurer les propriétés de style
    if saved_style:
        new_run.font.bold = saved_style['bold']
        new_run.font.italic = saved_style['italic']
        new_run.font.underline = saved_style['underline']
        if saved_style['size']:
            new_run.font.size = saved_style['size']
        if saved_style['color']:
            new_run.font.color.rgb = saved_style['color']
        if saved_style['name']:
            new_run.font.name = saved_style['name']

def generer_commentaire_ia(openrouter_api_key, formation="la formation"):
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {openrouter_api_key}",
        "HTTP-Referer": "https://formation-entreprise.com",
        "X-Title": "Générateur Questionnaires",
        "Content-Type": "application/json"
    }

    prompt_points_forts = (
        f"""Tu viens de suivre une formation en {formation}. Rédige une seule phrase courte et naturelle pour décrire ton ressenti. N'utilise aucune liste, numéro, puce, tiret, ou mise en forme Markdown (comme **gras** ou *italique*). Ne commence pas par \"1.\", \"2.\", \"-\", ou autre symbole. Ne donne qu’une phrase, sans retour à la ligne et évite de commencer ta phrase par Je mais soit le plus aléatoire possible base toi sur les exemples que je t'ai donné ci dessous.

Inspire-toi librement des idées suivantes (mais sans les reprendre telles quelles) :
Explications claires et outils
Formation pratico-pratique
Plein de tips concrets
Le formateur est pédagogue
Cours interactifs accessibles à tout moment
Supports bien structurés
On repart avec un système clé en main
Résultats concrets dès la fin de la session

Ta réponse doit être 100 % humaine, spontanée et fluide. Une seule phrase. Pas de formatage. Pas de liste. Juste un ressenti authentique."""
    )

    prompt_remarques = (
        f"""Tu viens de terminer une formation en {formation}. Rédige une phrase spontanée, comme si tu laissais un commentaire libre à chaud sur ton ressenti général ou un point que tu voudrais partager, évite de commencer ta phrase par "Je" ou "J'ai" mais soit le plus aléatoire possible dans le début de la phrase.

Ta réponse peut contenir une appréciation générale, une suggestion, un ressenti, une émotion ou une remarque libre. Tu peux être personnel, tant que tu restes positif ou constructif.

⚠️ Important :
- Ne commence pas par un numéro, un tiret ou une liste
- Ne fais pas de mise en forme (pas de gras, italique, etc.)
- Ne retourne qu'une seule phrase courte, humaine, naturelle
- Ne commence pas tes phrases par un vrai coup de boost ou un vrai plaisir soit aléatoire dans tes propositions
Ta réponse doit être 100 % humaine, spontanée et fluide. Une seule phrase. Pas de formatage. Pas de liste. Juste un ressenti authentique."""
    )

    def appeler_api(prompt):
        data = {
            "model": "openai/gpt-4.1",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.5,
            "max_tokens": 100
        }
        response = requests.post(url, headers=headers, json=data, timeout=10)
        response.raise_for_status()
        raw = response.json()['choices'][0]['message']['content'].strip()
        options = [ligne.strip() for ligne in raw.splitlines() if ligne.strip()]
        return random.choice(options) if options else ""

    try:
        commentaire_points_forts = appeler_api(prompt_points_forts)
        commentaire_remarques = appeler_api(prompt_remarques)
        return commentaire_points_forts, commentaire_remarques
    except Exception as e:
        st.error(f"Erreur API IA : {e}")
        return "", ""

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

    current_section = None
    formation_choice = str(participant['formation']).strip().lower()
    answer = None

    # Remplacer les placeholders dans tous les paragraphes
    for para in doc.paragraphs:
        remplacer_placeholders(para, replacements)

        text = para.text.lower()
        
        # Détection des sections
        if 'formation suivie' in text:
            current_section = 'formation'
            continue
        elif any(keyword in text for keyword in [
            'évaluation de la formation',
            'qualité du service',
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
            # Sauvegarder le style d'origine
            original_style = para.style
            
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

            original_text = option_text.split('[')[-1].split(']')[0].strip()
            
            # Réinitialiser le paragraphe avec le nouveau contenu
            para.clear()
            para.add_run(f'{symbol} {original_text}')
            para.style = original_style

    safe_prenom = re.sub(r'[^a-zA-Z0-9]', '_', str(participant['prénom']))
    safe_nom = re.sub(r'[^a-zA-Z0-9]', '_', str(participant['nom']))
    filename = f"Questionnaire_satisfaction_{safe_prenom}_{safe_nom}_{participant['session']}.docx"
    
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
frequence_ia = 1

if generer_ia:
    openrouter_api_key = st.text_input("Clé API OpenRouter", type="password")
    frequence_ia = st.slider("Fréquence de génération IA (1 sur x participants)", min_value=1, max_value=10, value=4, step=1)

if excel_file and template_file:
    try:
        df = pd.read_excel(excel_file)
        if not all(col in df.columns for col in REQUIRED_COLS):
            st.error(f"❌ Colonnes requises manquantes : {', '.join(REQUIRED_COLS)}")
            st.stop()

        st.success(f"✅ {len(df)} participants détectés")

        # ➕ Prévisualisation IA aléatoire
        if generer_ia and openrouter_api_key:
            st.markdown("### 🎲 Prévisualiser un commentaire IA aléatoire")

            if st.button("🧠 Générer une prévisualisation pour un participant sélectionné aléatoirement"):
                candidats = df.iloc[::frequence_ia]
                if not candidats.empty:
                    participant_test = candidats.sample(1).iloc[0]
                    cmt_fort, cmt_libre = generer_commentaire_ia(openrouter_api_key, participant_test['formation'])

                    st.markdown(f"**👤 Participant : {participant_test['prénom']} {participant_test['nom']} – Formation : {participant_test['formation']}**")
                    st.markdown("**🟢 Commentaire : Points forts**")
                    st.info(cmt_fort)
                    st.markdown("**💬 Commentaire : Remarque libre**")
                    st.info(cmt_libre)
                else:
                    st.warning("Aucun participant éligible avec la fréquence définie.")

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
                        commentaire_points_forts, commentaire_remarques = "", ""
                        if generer_ia and openrouter_api_key and idx % frequence_ia == 0:
                            commentaire_points_forts, commentaire_remarques = generer_commentaire_ia(openrouter_api_key, row['formation'])
                        try:
                            output_path = generer_questionnaire(row, template_path, commentaire_points_forts, commentaire_remarques)
                            zipf.write(output_path, os.path.basename(output_path))
                            progress_bar.progress((idx + 1)/len(df), text=f"Progress: {idx+1}/{len(df)}")
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
