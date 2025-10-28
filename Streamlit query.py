import streamlit as st
import openpyxl
from datetime import datetime
from io import BytesIO
import json
import os

# Tenter d'importer cohere seulement si la clé existe
COHERE_KEY = None
try:
    COHERE_KEY = st.secrets.get("COHERE_API_KEY")
except Exception:
    COHERE_KEY = os.getenv("COHERE_API_KEY")
co = None
if COHERE_KEY:
    try:
        import cohere
        co = cohere.Client(api_key=COHERE_KEY)
    except Exception:
        co = None

# --- CSS personnalisé ---
st.markdown("""
    <style>
        .stTextArea textarea, .stTextArea, .full-width-reco {width: 100% !important;}
        .stButton > button {
            color: #FFFFFF !important;
            background: #00E47C !important;
            border: none !important;
            border-radius: 8px !important;
        }
        .progress-label { font-weight:600; color:#08312A; margin-bottom:6px; }
        .sidebar-section { margin-bottom:16px; padding:8px; border-radius:6px; background:#f7fff6; }
        .question-box { padding:12px; border-radius:8px; background:#ffffff; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }
    </style>
""", unsafe_allow_html=True)

# Questions (identiques, mêmes emojis)
questions = {
    "name": {"question": "1.1. Quel est votre nom et prénom ? 😊", "options": []},
    "email": {"question": "1.2. Quelle est votre adresse e-mail ? 📧", "options": []},
    "business_unit": {
        "question": "1.3. Quelle est votre Business Unit ? 🏢",
        "options": ["Pet Vet", "Avian", "Ruminant", "Swine", "Equine", "Pet Retail"]
    },
    "supplier_name": {"question": "2.1. Quel est le nom du fournisseur ? 🏭", "options": []},
    "product_code": {
        "question": "2.2. Le produit a-t-il déjà un code existant ? 🔢",
        "options": ["Oui", "Non"]
    },
    "product_code_yes": {"question": "Veuillez indiquer le SKU actuel : 🆔", "options": []},
    "product_code_no": {"question": "Veuillez indiquer le SKU précédent ou similaire : 🆔", "options": []},
    "product_description": {
        "question": "2.3. Fournissez une brève description du produit : 📝 ou description rattachée en automatique?",
        "options": []
    },
    "supplier_conditions": {
        "question": "3.1. Le fournisseur impose-t-il une quantité minimale de commande, ou taille de lot? 📦",
        "options": ["Oui", "Non"]
    },
    "quantity_minimum_yes": {
        "question": "Indiquez la quantité minimale requise : 🔢, ou à négocier? - Y a t-il des paliers de prix avec remise possible?",
        "options": []
    },
    "coverage_duration": {
        "question": "3.2. Avez-vous une idée de la durée de couverture estimée ? ⏳",
        "options": ["Oui", "Non"]
    },
    "coverage_duration_yes": {
        "question": "Indiquez la durée de couverture estimée (en mois) : 📅, selon l'historique des ventes en N-1",
        "options": []
    },
    "supplier_location": {
        "question": "4.1. Où est basé le fournisseur ? 🌍",
        "options": ["En France", "Europe", "Grand export"]
    },
    "availability_delay": {
        "question": "4.2. Quel est le délai estimé pour la mise à disposition du produit ? ⏱️",
        "options": []
    },
    "storage_location": {
        "question": "5.1. le SKU accompagne-t-il des produits finis? 📍",
        "options": ["Oui", "Non"]
    },
    "sku_open": {
        "question": "5.2. le SKU doit-il être ouvert dans Bi connect?",
        "options": ["Oui", "Non"]
    },
    "sku_frequency": {
        "question": "5.3. le SKU est-il ponctuel ou récurrent?",
        "options": []
    },
    "dotation": {
        "question": "6.1. Le produit est-il destiné à une dotation ? 🎁",
        "options": ["Oui", "Non"]
    },
    "dotation_yes": {
        "question": "Veuillez indiquer les délais impératifs de livraison sur le 3PL : 🚚",
        "options": []
    },
    "additional_requirements": {
        "question": "7.1. Y a-t-il des exigences supplémentaires ? ❓",
        "options": []
    }
}
FINAL_KEY = "final"

def get_next_question(answer, previous_question):
    mapping = {
        "name": "email",
        "email": "business_unit",
        "business_unit": "supplier_name",
        "supplier_name": "product_code",
        "product_code": {"Oui": "product_code_yes", "Non": "product_code_no"},
        "product_code_yes": "product_description",
        "product_code_no": "product_description",
        "product_description": "supplier_conditions",
        "supplier_conditions": {"Oui": "quantity_minimum_yes", "Non": "coverage_duration"},
        "quantity_minimum_yes": "coverage_duration",
        "coverage_duration": {"Oui": "coverage_duration_yes", "Non": "supplier_location"},
        "coverage_duration_yes": "supplier_location",
        "supplier_location": "availability_delay",
        "availability_delay": "storage_location",
        "storage_location": "sku_open",
        "sku_open": "sku_frequency",
        "sku_frequency": "dotation",
        "dotation": {"Oui": "dotation_yes", "Non": "additional_requirements"},
        "dotation_yes": "additional_requirements",
        "additional_requirements": FINAL_KEY
    }
    next_question = mapping.get(previous_question)
    if isinstance(next_question, dict):
        return next_question.get(answer)
    return next_question

def build_path(answers):
    """Construit la séquence de questions selon les réponses actuelles."""
    path = []
    q = list(questions.keys())[0]
    while q and q != FINAL_KEY:
        path.append(q)
        a = answers.get(q)
        q = get_next_question(a, q)
    return path

def get_prev_question(current_question, answers):
    path = build_path(answers)
    if current_question not in path:
        return None
    idx = path.index(current_question)
    return path[idx-1] if idx > 0 else None

def save_answers_to_excel_bytes(recommendation, ai_recommendation, answers):
    user_name = answers.get("name") or "utilisateur"
    current_date = datetime.now().strftime("%Y-%m-%d")
    file_name = f"{user_name}_{current_date}.xlsx"
    output = BytesIO()
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Réponses"
    sheet.cell(row=1, column=1, value="Question")
    sheet.cell(row=1, column=2, value="Réponse")
    for idx, (question, answer) in enumerate(answers.items(), start=2):
        sheet.cell(row=idx, column=1, value=questions.get(question, {}).get("question", question))
        sheet.cell(row=idx, column=2, value=answer)
    sheet.cell(row=len(answers) + 2, column=1, value="Recommandations")
    sheet.cell(row=len(answers) + 2, column=2, value=recommendation)
    sheet.cell(row=len(answers) + 3, column=1, value="Recommandations IA")
    sheet.cell(row=len(answers) + 3, column=2, value=ai_recommendation)
    workbook.save(output)
    output.seek(0)
    return output, file_name

def get_ai_recommendation(answers):
    if not co:
        return "Recommandations IA non disponibles (clé COHERE non configurée)."
    try:
        prompt = "Voici les réponses d'un utilisateur à un questionnaire :\n"
        for question, answer in answers.items():
            prompt += f"- {questions.get(question, {}).get('question','')} : {answer}\n"
        prompt += "En vous basant sur ces réponses, fournissez des recommandations supplémentaires pertinentes (max 30 mots, en français) :"
        response = co.generate(prompt=prompt, model="command")
        return response.generations[0].text.strip()
    except Exception as e:
        return f"Erreur IA : {str(e)}"

# Initialiser l'état
if "current_question" not in st.session_state:
    st.session_state.current_question = list(questions.keys())[0]
if "user_answers" not in st.session_state:
    st.session_state.user_answers = {}
if "ai_recommendation" not in st.session_state:
    st.session_state.ai_recommendation = None

# Barre latérale: navigation, sauvegarder/charger le progrès
with st.sidebar:
    st.markdown("<div class='sidebar-section'>", unsafe_allow_html=True)
    st.markdown("### Navigation rapide")
    path = build_path(st.session_state.user_answers)
    for q in path:
        display = questions[q]["question"]
        if st.button(display, key=f"goto_{q}"):
            st.session_state.current_question = q
            st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div class='sidebar-section'>", unsafe_allow_html=True)
    st.markdown("### Progrès et options")
    if st.button("Réinitialiser l'enquête"):
        st.session_state.current_question = list(questions.keys())[0]
        st.session_state.user_answers = {}
        st.session_state.ai_recommendation = None
        st.rerun()
    st.markdown("Sauvegarder / Charger le progrès")
    json_bytes = json.dumps(st.session_state.user_answers, ensure_ascii=False, indent=2).encode("utf-8")
    st.download_button("Sauvegarder progrès (JSON)", data=json_bytes, file_name="progres_enquete.json", mime="application/json")
    uploaded = st.file_uploader("Charger progrès (JSON)", type=["json"])
    if uploaded:
        try:
            loaded = json.load(uploaded)
            if isinstance(loaded, dict):
                st.session_state.user_answers.update(loaded)
                path2 = build_path(st.session_state.user_answers)
                st.session_state.current_question = path2[-1] if path2 else list(questions.keys())[0]
                st.rerun()
            else:
                st.error("JSON invalide.")
        except Exception as e:
            st.error(f"Erreur lors du chargement JSON: {e}")
    st.markdown("</div>", unsafe_allow_html=True)

# En-tête principal
st.title("Outil Marketing Survey")
st.write("Merci de répondre aux questions pour obtenir des recommandations personnalisées.")

# Barre de progrès calculée
full_path = build_path(st.session_state.user_answers)
total = max(1, len(full_path))
if st.session_state.current_question in full_path:
    idx = full_path.index(st.session_state.current_question)
else:
    idx = len(full_path) - 1 if full_path else 0
progress = idx / total
st.markdown(f"<div class='progress-label'>Progrès: {idx}/{len(full_path)}</div>", unsafe_allow_html=True)
st.progress(progress)

# Zone principale du formulaire
current_question_key = st.session_state.current_question
if current_question_key != FINAL_KEY:
    question_data = questions.get(current_question_key)
    st.markdown("<div class='question-box'>", unsafe_allow_html=True)
    st.subheader(question_data["question"])
    prev_value = st.session_state.user_answers.get(current_question_key, "")
    
    # Widget de question sans formulaire
    if question_data["options"]:
        try:
            index = question_data["options"].index(prev_value) if prev_value in question_data["options"] else 0
        except Exception:
            index = 0
        answer = st.radio("", question_data["options"], index=index, key=f"widget_{current_question_key}")
    else:
        answer = st.text_input("", value=prev_value, key=f"widget_{current_question_key}", placeholder="Saisissez votre réponse ici...")
    
    # Boutons de navigation
    col_left, col_center, col_right = st.columns([1, 1, 1])
    
    with col_left:
        prev_q = get_prev_question(current_question_key, st.session_state.user_answers)
        if prev_q and st.button("Précédent ⬅️", key=f"prev_{current_question_key}"):
            st.session_state.user_answers[current_question_key] = answer
            st.session_state.current_question = prev_q
            st.rerun()
    
    with col_center:
        if st.button("Sauvegarder et quitter", key=f"save_{current_question_key}"):
            st.session_state.user_answers[current_question_key] = answer
            st.success("Progrès sauvegardé. Vous pouvez télécharger le JSON dans la barre latérale.")
    
    with col_right:
        if st.button("Suivant ➡️", key=f"next_{current_question_key}"):
            st.session_state.user_answers[current_question_key] = answer
            next_q = get_next_question(answer, current_question_key)
            st.session_state.current_question = next_q if next_q else FINAL_KEY
            st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)
else:
    st.header("Fin du formulaire 🏁")
    st.markdown("Vérifiez vos réponses avant de générer les recommandations.")
    answers = st.session_state.user_answers
    for q in build_path(answers):
        st.markdown(f"**{questions[q]['question']}**")
        st.write(answers.get(q, "_Pas de réponse_"))
        if st.button(f"Modifier", key=f"edit_{q}"):
            st.session_state.current_question = q
            st.rerun()

    # Règles locales de recommandation
    recommendation = []
    if answers.get("product_code") == "Non":
        recommendation.append("- Créer un nouveau code dans le système avant la première commande.")
    if answers.get("supplier_conditions") == "Oui":
        recommendation.append("- Analyser la consommation historique pour ajuster les hypothèses de réapprovisionnement.")
    if answers.get("supplier_location") == "Grand export":
        recommendation.append("- Anticiper les délais logistiques et créer un buffer de sécurité.")
    if answers.get("dotation") == "Oui":
        recommendation.append("- Coordonner avec le 3PL pour respecter les délais impératifs.")

    st.markdown("#### Recommandations automatiques")
    if st.button("Générer recommandations (IA)"):
        st.session_state.ai_recommendation = get_ai_recommendation(answers)

    ai_text = st.session_state.ai_recommendation or "Aucune recommandation IA générée pour l'instant."
    full_reco = ("\n".join(recommendation) if recommendation else "Aucune recommandation automatique locale.") + "\n\nRecommandations IA :\n" + ai_text
    st.text_area("Recommandations", full_reco, height=300)

    # Télécharger Excel et JSON
    if st.button("Télécharger réponses (Excel)"):
        excel_bytes, excel_filename = save_answers_to_excel_bytes("\n".join(recommendation), ai_text, answers)
        if excel_bytes:
            st.download_button("Télécharger Excel", data=excel_bytes, file_name=excel_filename,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.error("Erreur lors de la génération d'Excel.")

    json_bytes = json.dumps(answers, ensure_ascii=False, indent=2).encode("utf-8")
    st.download_button("Télécharger réponses (JSON)", data=json_bytes, file_name="reponses.json", mime="application/json")