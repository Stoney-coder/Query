
# Ensure the necessary libraries are installed before running the Streamlit app.
# You can install the required libraries using the following commands:
# pip install streamlit openpyxl cohere

import streamlit as st
import openpyxl
import os
from datetime import datetime
import cohere

# Check if openpyxl is installed
try:
    import openpyxl
except ImportError:
    st.error("The 'openpyxl' library is not installed. Please install it using 'pip install openpyxl'.")

# Initialize Cohere Client
co = cohere.Client(api_key="CADyn7RJ5sXnikvmipLYLSyWhoUvJS56FksKuAEQ")  # Replace with your actual API key

# Directory to save Excel files
excel_directory = os.path.expanduser("~/Desktop/Query_Answers")
os.makedirs(excel_directory, exist_ok=True)

# Dictionary to store user answers
user_answers = {}

# Dictionary of questions and conditional answers
questions = {
    "name": {
        "question": "1.1. Quel est votre nom complet ? 😊",
        "options": []
    },
    "email": {
        "question": "1.2. Quelle est votre adresse e-mail ? 📧",
        "options": []
    },
    "business_unit": {
        "question": "1.3. Quelle est votre Business Unit ? 🏢",
        "options": ["Pet Vet", "Avian", "Ruminant", "Swine", "Equine", "Pet Retail"]
    },
    "supplier_name": {
        "question": "2.1. Quel est le nom du fournisseur ? 🏭",
        "options": []
    },
    "product_code": {
        "question": "2.2. Le produit a-t-il déjà un code existant ? 🔢",
        "options": ["Oui", "Non"]
    },
    "product_code_yes": {
        "question": "Veuillez indiquer le SKU actuel ou précédent : 🆔",
        "options": []
    },
    "product_description": {
        "question": "2.3. Fournissez une brève description du produit : 📝",
        "options": []
    },
    "supplier_conditions": {
        "question": "3.1. Le fournisseur impose-t-il une quantité minimale ? 📦",
        "options": ["Oui", "Non"]
    },
    "quantity_minimum_yes": {
        "question": "Indiquez la quantité minimale requise : 🔢",
        "options": []
    },
    "coverage_duration": {
        "question": "3.2. Connaissez-vous la durée de couverture estimée pour cette quantité ? ⏳",
        "options": ["Oui", "Non"]
    },
    "coverage_duration_yes": {
        "question": "Indiquez la durée de couverture estimée (en mois) : 📅",
        "options": []
    },
    "coverage_assumption": {
        "question": "Sur quelle hypothèse repose ce réapprovisionnement ? 🤔",
        "options": []
    },
    "supplier_location": {
        "question": "4.1. Où est basé le fournisseur ? 🌍",
        "options": ["France", "Europe (hors France)", "Hors Europe (Grand Export)"]
    },
    "availability_delay": {
        "question": "4.2. Quel est le délai estimé pour la mise à disposition en France (en jours) ? ⏱️",
        "options": []
    },
    "storage_location": {
        "question": "5.1. Connaissez-vous la localisation de stockage ? (Sélection multiple possible) 📍",
        "options": [
            "Centrales d'achat → stockage Movianto",
            "Vétos → stockage Movianto",
            "Mixte",
            "Pharmacies",
            "Délégués"
        ]
    },
    "dotation": {
        "question": "6.1. Le produit est-il destiné à une dotation ? 🎁",
        "options": ["Oui", "Non"]
    },
    "dotation_yes": {
        "question": "Veuillez indiquer les délais impératifs de livraison sur le 3PL pour éviter des livraisons multiples : 🚚",
        "options": []
    },
    "additional_requirements": {
        "question": "7.1. Y a-t-il des exigences ou contraintes supplémentaires pour ce produit ? ❓",
        "options": []
    },
    "final": {
        "question": "Fin du formulaire. Recommandations Automatiques Basées sur Vos Réponses 🏁",
        "options": []
    }
}

# Function to determine the next question
def get_next_question(answer, previous_question):
    mapping = {
        "1.1. Quel est votre nom complet ? 😊": "email",
        "1.2. Quelle est votre adresse e-mail ? 📧": "business_unit",
        "1.3. Quelle est votre Business Unit ? 🏢": "supplier_name",
        "2.1. Quel est le nom du fournisseur ? 🏭": "product_code",
        "2.2. Le produit a-t-il déjà un code existant ? 🔢": {
            "Oui": "product_code_yes",
            "Non": "product_description"
        },
        "Veuillez indiquer le SKU actuel ou précédent : 🆔": "product_description",
        "2.3. Fournissez une brève description du produit : 📝": "supplier_conditions",
        "3.1. Le fournisseur impose-t-il une quantité minimale ? 📦": {
            "Oui": "quantity_minimum_yes",
            "Non": "supplier_location"
        },
        "Indiquez la quantité minimale requise : 🔢": "coverage_duration",
        "3.2. Connaissez-vous la durée de couverture estimée pour cette quantité ? ⏳": {
            "Oui": "coverage_duration_yes",
            "Non": "coverage_assumption"
        },
        "Indiquez la durée de couverture estimée (en mois) : 📅": "supplier_location",
        "Sur quelle hypothèse repose ce réapprovisionnement ? 🤔": "supplier_location",
        "4.1. Où est basé le fournisseur ? 🌍": "availability_delay",
        "4.2. Quel est le délai estimé pour la mise à disposition en France (en jours) ? ⏱️": "storage_location",
        "5.1. Connaissez-vous la localisation de stockage ? (Sélection multiple possible) 📍": "dotation",
        "6.1. Le produit est-il destiné à une dotation ? 🎁": {
            "Oui": "dotation_yes",
            "Non": "additional_requirements"
        },
        "Veuillez indiquer les délais impératifs de livraison sur le 3PL pour éviter des livraisons multiples : 🚚": "additional_requirements",
        "7.1. Y a-t-il des exigences ou contraintes supplémentaires pour ce produit ? ❓": "final"
    }
    next_question = mapping.get(previous_question)
    if isinstance(next_question, dict):
        return next_question.get(answer)
    return next_question

# Function to save answers to an Excel file
def save_answers_to_excel(recommendation, ai_recommendation):
    user_name = user_answers.get("1.1. Quel est votre nom complet ? 😊")
    if not user_name:
        st.warning("Le nom de l'utilisateur est manquant.")
        return
    current_date = datetime.now().strftime("%d-%m-%Y")
    file_name = f"{user_name}_{current_date}.xlsx"
    file_path = os.path.join(excel_directory, file_name)
    try:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Réponses"
        sheet.cell(row=1, column=1, value="Question")
        sheet.cell(row=1, column=2, value="Réponse")
        for idx, (question, answer) in enumerate(user_answers.items(), start=2):
            sheet.cell(row=idx, column=1, value=question)
            sheet.cell(row=idx, column=2, value=answer)
        # Add recommendations
        sheet.cell(row=len(user_answers) + 2, column=1, value="Recommandations")
        sheet.cell(row=len(user_answers) + 2, column=2, value=recommendation)
        sheet.cell(row=len(user_answers) + 3, column=1, value="Recommandations IA")
        sheet.cell(row=len(user_answers) + 3, column=2, value=ai_recommendation)
        workbook.save(file_path)
        st.success(f"Les réponses ont été enregistrées dans {file_path}")
    except PermissionError:
        st.error(f"Impossible d'enregistrer le fichier. Vérifiez les permissions pour le chemin : {file_path}")

# Function to display recommendations based on answers
def show_recommendation():
    recommendation = "Recommandations :\n"
    product_code = user_answers.get("2.2. Le produit a-t-il déjà un code existant ? 🔢")
    if product_code == "Non":
        recommendation += "- Assurez-vous de créer un nouveau code dans le système avant de passer commande.\n"
    quantity_minimum = user_answers.get("3.1. Le fournisseur impose-t-il une quantité minimale ? 📦")
    if quantity_minimum == "Oui":
        recommendation += "- Recommandez une analyse de consommation historique pour ajuster les hypothèses de réapprovisionnement.\n"
    supplier_location = user_answers.get("4.1. Où est basé le fournisseur ? 🌍")
    if supplier_location == "Hors Europe (Grand Export)":
        recommendation += "- Prévoir un délai logistique plus long et anticiper les commandes.\n"
    storage_location = user_answers.get("5.1. Connaissez-vous la localisation de stockage ? (Sélection multiple possible) 📍")
    if storage_location and ("Mixte" in storage_location or "Délégués" in storage_location):
        recommendation += "- Vérifiez la coordination entre les différents points de distribution pour éviter les ruptures.\n"
    dotation = user_answers.get("6.1. Le produit est-il destiné à une dotation ? 🎁")
    if dotation == "Oui":
        recommendation += "- Priorisez la planification logistique avec le 3PL pour respecter les délais impératifs.\n"
    additional_requirements = user_answers.get("7.1. Y a-t-il des exigences ou contraintes supplémentaires pour ce produit ? ❓")
    if additional_requirements:
        recommendation += f"- Notes supplémentaires : {additional_requirements}\n"

    def get_ai_recommendation(answers):
        try:
            # Construct the prompt based on user answers
            prompt = "Voici les réponses d'un utilisateur à un questionnaire :\n"
            for question, answer in answers.items():
                prompt += f"- {question}: {answer}\n"
            prompt += "Basé sur ces réponses, fournissez des recommandations supplémentaires pertinentes 30 mots max:"
            # Use Cohere's chat API for generating recommendations
            response = co.chat(
                message=prompt,
                chat_history=[]  # Optionally, provide previous chat history
            )
            # Extract and return the response text
            return response.text.strip()
        except Exception as e:
            return f"Erreur lors de la génération des recommandations IA : {str(e)}"

    ai_recommendation = get_ai_recommendation(user_answers)
    recommendation += f"\nRecommandations IA :\n{ai_recommendation}"
    st.subheader("Recommandations")
    st.text(recommendation)
    if st.button("Enregistrer les réponses"):
        save_answers_to_excel(recommendation, ai_recommendation)

# Streamlit app
st.title("Outil Marketing Survey")
st.sidebar.title("Navigation")
current_question_key = st.sidebar.selectbox("Questions", list(questions.keys()), index=0)

if current_question_key:
    question_data = questions[current_question_key]
    st.subheader(question_data["question"])
    if question_data["options"]:
        selected_option = st.radio("Options", question_data["options"])
        if st.button("Suivant"):
            user_answers[question_data["question"]] = selected_option
            next_question_key = get_next_question(selected_option, question_data["question"])
            st.sidebar.selectbox("Questions", list(questions.keys()), index=list(questions.keys()).index(next_question_key))
    else:
        user_input = st.text_input("Votre réponse")
        if st.button("Suivant"):
            user_answers[question_data["question"]] = user_input
            next_question_key = get_next_question(user_input, question_data["question"])
            st.sidebar.selectbox("Questions", list(questions.keys()), index=list(questions.keys()).index(next_question_key))

if current_question_key == "final":
    show_recommendation()
