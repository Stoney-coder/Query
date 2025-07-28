import streamlit as st
import openpyxl
from datetime import datetime
from io import BytesIO
import cohere

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
    </style>
""", unsafe_allow_html=True)

co = cohere.Client(api_key="VOTRE_COHERE_API_KEY")  # Remplacez par votre clé

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

def get_prev_question(current_question, answers):
    keys = list(questions.keys())
    if current_question == keys[0]:
        return None
    path = [keys[0]]
    for i in range(len(answers)):
        q = path[-1]
        a = answers.get(q)
        n = get_next_question(a, q)
        if n == current_question:
            return q
        if n:
            path.append(n)
    return None

def save_answers_to_excel(recommendation, ai_recommendation):
    user_name = st.session_state.user_answers.get("name")
    if not user_name:
        st.warning("Le nom de l'utilisateur est manquant.")
        return None, None
    current_date = datetime.now().strftime("%d-%m-%Y")
    file_name = f"{user_name}_{current_date}.xlsx"
    output = BytesIO()
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Réponses"
    sheet.cell(row=1, column=1, value="Question")
    sheet.cell(row=1, column=2, value="Réponse")
    for idx, (question, answer) in enumerate(st.session_state.user_answers.items(), start=2):
        sheet.cell(row=idx, column=1, value=question)
        sheet.cell(row=idx, column=2, value=answer)
    sheet.cell(row=len(st.session_state.user_answers) + 2, column=1, value="Recommandations")
    sheet.cell(row=len(st.session_state.user_answers) + 2, column=2, value=recommendation)
    sheet.cell(row=len(st.session_state.user_answers) + 3, column=1, value="Recommandations IA")
    sheet.cell(row=len(st.session_state.user_answers) + 3, column=2, value=ai_recommendation)
    workbook.save(output)
    output.seek(0)
    return output, file_name

def show_recommendation():
    recommendation = "Recommandations :\n"
    product_code = st.session_state.user_answers.get("product_code")
    if product_code == "Non":
        recommendation += "- Veuillez créer un nouveau code dans le système avant de passer commande.\n"
    quantity_minimum = st.session_state.user_answers.get("supplier_conditions")
    if quantity_minimum == "Oui":
        recommendation += "- Il est conseillé d’analyser la consommation historique pour adapter les hypothèses de réapprovisionnement.\n"
    supplier_location = st.session_state.user_answers.get("supplier_location")
    if supplier_location == "Grand export":
        recommendation += "- Prévoyez un délai logistique plus long et anticipez les commandes.\n"
    dotation = st.session_state.user_answers.get("dotation")
    if dotation == "Oui":
        recommendation += "- Priorisez la planification logistique avec le 3PL pour respecter les délais impératifs.\n"
    def get_ai_recommendation(answers):
        try:
            prompt = "Voici les réponses d'un utilisateur à un questionnaire :\n"
            for question, answer in answers.items():
                prompt += f"- {question}: {answer}\n"
            prompt += "En vous basant sur ces réponses, fournissez des recommandations supplémentaires pertinentes (max 30 mots, en français) :"
            response = co.generate(prompt=prompt, model="command")
            return response.generations[0].text.strip()
        except Exception as e:
            return f"Erreur lors de la génération des recommandations IA : {str(e)}"
    ai_recommendation = get_ai_recommendation(st.session_state.user_answers)
    recommendation_full = f"{recommendation}\nRecommandations IA :\n{ai_recommendation}"
    st.markdown('<div class="full-width-reco">', unsafe_allow_html=True)
    st.text_area("Recommandations", recommendation_full, height=300)
    st.markdown('</div>', unsafe_allow_html=True)
    if st.button("Télécharger les réponses"):
        excel_bytes, excel_filename = save_answers_to_excel(recommendation, ai_recommendation)
        if excel_bytes:
            st.download_button(
                label="Télécharger le fichier Excel",
                data=excel_bytes,
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.info(f"Le mot de passe pour ouvrir le fichier est : {excel_filename.replace('.xlsx','')}")
        else:
            st.error("Erreur lors de la création du fichier Excel.")

def main():
    st.title("Outil Marketing Survey")
    st.write("Merci de répondre aux questions pour obtenir des recommandations personnalisées.")
    # Initialisation de l'état
    if "current_question" not in st.session_state:
        st.session_state.current_question = "name"
    if "user_answers" not in st.session_state:
        st.session_state.user_answers = {}
    current_question_key = st.session_state.current_question
    if current_question_key != FINAL_KEY:
        question_data = questions.get(current_question_key)
        widget_key = f"widget_{current_question_key}"
        st.subheader(question_data["question"])
        # Affichage et réponse
        if question_data["options"]:
            answer = st.radio("Choisissez une option :", question_data["options"], key=widget_key)
        else:
            answer = st.text_input("Votre réponse :", key=widget_key)
        col1, col2 = st.columns([1,1])
        # Bouton Précédent
        prev_question = get_prev_question(current_question_key, st.session_state.user_answers)
        with col1:
            if prev_question:
                if st.button("Précédent ⬅️"):
                    st.session_state.current_question = prev_question
        # Bouton Suivant
        with col2:
            if st.button("Suivant"):
                st.session_state.user_answers[current_question_key] = answer
                next_question = get_next_question(answer, current_question_key)
                if next_question:
                    st.session_state.current_question = next_question
                    next_widget_key = f"widget_{next_question}"
                    if next_widget_key in st.session_state:
                        del st.session_state[next_widget_key]
                else:
                    st.session_state.current_question = FINAL_KEY
        # Message d'aide
        st.markdown("<br><div style='color:#08312A;font-weight:bold;'>Remarque : Pour passer à la prochaine question, veuillez répondre puis cliquer deux fois sur 'Suivant'.</div>", unsafe_allow_html=True)
    else:
        st.header("Fin du formulaire 🏁")
        show_recommendation()

if __name__ == "__main__":
    main()
