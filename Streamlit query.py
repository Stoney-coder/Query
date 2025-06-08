import streamlit as st
import openpyxl
import os
from datetime import datetime

# Directory to save the Excel files
excel_directory = os.path.expanduser("~/Desktop/Query_Answers")
os.makedirs(excel_directory, exist_ok=True)

# Dictionary to store user answers
user_answers = {}

# Dictionary of questions and conditional answers
questions = {
    "name": {
        "question": "Quel est votre nom ?",
        "options": []
    },
    "start": {
        "question": "1.1. Objectifs principaux (sélectionner tout ce qui s’applique) :",
        "options": [
            "Réaliser des économies",
            "Réduire les mouvements inutiles",
            "Limiter les destructions à terme",
            "Améliorer la durabilité",
            "Autre"
        ]
    },
    "supplier_location": {
        "question": "2.1. Où est basé le fournisseur ?",
        "options": ["France", "Europe", "Hors Europe (Grand Export)"]
    },
    "sustainability": {
        "question": "2.2. Le fournisseur utilise-t-il des pratiques durables ?",
        "options": ["Oui", "Non", "Inconnu"]
    },
    "lot_size": {
        "question": "2.3. Y a-t-il une taille de lot minimale imposée ?",
        "options": ["Oui", "Non"]
    },
    "lot_size_yes": {
        "question": "Si oui, précisez la taille minimale :",
        "options": []
    },
    "replace_existing": {
        "question": "2.4. Cet article remplace-t-il un article existant ?",
        "options": ["Oui", "Non"]
    },
    "replace_existing_yes": {
        "question": "Si oui, code précédent :",
        "options": []
    },
    "lot_recommendation": {
        "question": "2.5. La taille de lot permet-elle de suivre les recommandations de couverture ?",
        "options": ["Oui", "Non"]
    },
    "price_tiers": {
        "question": "2.6. Existe-t-il des paliers de prix avec remises ?",
        "options": ["Oui", "Non"]
    },
    "price_tiers_yes": {
        "question": "Si oui, détaillez les paliers de prix :",
        "options": []
    },
    "availability": {
        "question": "2.7. Quel est le délai de mise à disposition ?",
        "options": ["Moins d’1 mois", "1 à 3 mois", "Plus de 3 mois"]
    },
    "client_type": {
        "question": "3.1. À quel(s) type(s) de clients le produit est-il destiné ?",
        "options": [
            "Centrales d’achat",
            "Vétérinaires",
            "Pharmacies",
            "Délégués",
            "Mixte",
            "Autre"
        ]
    },
    "dotation": {
        "question": "3.2. Le produit est-il destiné à une dotation ?",
        "options": ["Oui", "Non"]
    },
    "product_details": {
        "question": "4.1. Description du produit :\nDimensions : ___________\nPoids : ___________\nDate de péremption / validité : ___________",
        "options": []
    },
    "packaging": {
        "question": "5.1. Le code peut-il être expédié en fardelage ?",
        "options": ["Oui", "Non"]
    },
    "packaging_yes": {
        "question": "Si oui, quel type de fardelage ?",
        "options": ["Par 5", "Par 10", "Autre"]
    },
    "final": {
        "question": "Merci pour vos réponses. Cliquez sur 'Afficher les recommandations' pour voir les suggestions.",
        "options": []
    }
}

# Function to determine the next question
def get_next_question(answer, previous_question):
    mapping = {
        "Quel est votre nom ?": "start",
        "1.1. Objectifs principaux (sélectionner tout ce qui s’applique) :": "supplier_location",
        "2.1. Où est basé le fournisseur ?": "sustainability",
        "2.2. Le fournisseur utilise-t-il des pratiques durables ?": "lot_size",
        "2.3. Y a-t-il une taille de lot minimale imposée ?": {
            "Oui": "lot_size_yes",
            "Non": "replace_existing"
        },
        "2.4. Cet article remplace-t-il un article existant ?": {
            "Oui": "replace_existing_yes",
            "Non": "lot_recommendation"
        },
        "2.5. La taille de lot permet-elle de suivre les recommandations de couverture ?": "price_tiers",
        "2.6. Existe-t-il des paliers de prix avec remises ?": {
            "Oui": "price_tiers_yes",
            "Non": "availability"
        },
        "2.7. Quel est le délai de mise à disposition ?": "client_type",
        "3.1. À quel(s) type(s) de clients le produit est-il destiné ?": "dotation",
        "3.2. Le produit est-il destiné à une dotation ?": "product_details",
        "4.1. Description du produit :\nDimensions : ___________\nPoids : ___________\nDate de péremption / validité : ___________": "packaging",
        "5.1. Le code peut-il être expédié en fardelage ?": {
            "Oui": "packaging_yes",
            "Non": "final"
        },
        "5.2. Si oui, quel type de fardelage ?": "final"
    }
    next_question = mapping.get(previous_question)
    if isinstance(next_question, dict):
        return next_question.get(answer)
    return next_question

# Function to save answers to a new Excel file
def save_answers_to_excel():
    user_name = user_answers.get("Quel est votre nom ?")
    if not user_name:
        st.warning("Le nom de l'utilisateur est manquant.")
        return None

    # Generate file name with user name and current date (dd-mm-yyyy format)
    current_date = datetime.now().strftime("%d-%m-%Y")
    file_name = f"{user_name}_{current_date}.xlsx"
    file_path = os.path.join(excel_directory, file_name)

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Answers"

    # Write headers
    sheet.cell(row=1, column=1, value="Question")
    sheet.cell(row=1, column=2, value="Answer")

    # Write answers
    for idx, (question, answer) in enumerate(user_answers.items(), start=2):
        sheet.cell(row=idx, column=1, value=question)
        sheet.cell(row=idx, column=2, value=answer)

    workbook.save(file_path)
    return file_path

# Streamlit Application
st.title("Enquête Logistique Interactive")

# Initialize session state
if "current_question" not in st.session_state:
    st.session_state.current_question = "name"
if "user_answers" not in st.session_state:
    st.session_state.user_answers = {}

# Display the current question
current_key = st.session_state.current_question
if current_key in questions:
    q_data = questions[current_key]
    st.subheader(q_data["question"])

    if not q_data["options"]:  # Open-ended question
        answer = st.text_input("Votre réponse :")
        if st.button("Suivant"):
            if answer.strip():
                st.session_state.user_answers[q_data["question"]] = answer
                st.session_state.current_question = get_next_question(answer, q_data["question"])
            else:
                st.warning("Veuillez entrer une réponse.")
    else:
        answer = st.radio("Choisissez une option :", q_data["options"])
        if st.button("Suivant"):
            st.session_state.user_answers[q_data["question"]] = answer
            st.session_state.current_question = get_next_question(answer, q_data["question"])
else:
    st.success("Merci pour vos réponses !")
    recommendation = "Recommandation :\n"
    supplier_location = st.session_state.user_answers.get("2.1. Où est basé le fournisseur ?")
    if supplier_location == "France":
        recommendation += "- Couverture recommandée : 3 mois\n"
    elif supplier_location == "Europe":
        recommendation += "- Couverture recommandée : 6 mois\n"
    elif supplier_location == "Hors Europe (Grand Export)":
        recommendation += "- Couverture recommandée : 1 an\n"

    sustainability = st.session_state.user_answers.get("2.2. Le fournisseur utilise-t-il des pratiques durables ?")
    if sustainability == "Oui":
        recommendation += "- Privilégiez ce fournisseur pour des objectifs de durabilité.\n"
    elif sustainability == "Non":
        recommendation += "- Envisagez de négocier des pratiques plus durables avec ce fournisseur.\n"

    st.write(recommendation)

    file_path = save_answers_to_excel()
    if file_path:
        with open(file_path, "rb") as file:
            st.download_button(
                label="Télécharger les réponses",
                data=file,
                file_name=os.path.basename(file_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )