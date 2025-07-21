import streamlit as st
import openpyxl
import os
from datetime import datetime
import cohere

# Initialize Cohere Client
co = cohere.Client(api_key="CADyn7RJ5sXnikvmipLYLSyWhoUvJS56FksKuAEQ")

# Directory to save Excel files
excel_directory = os.path.expanduser("~/Desktop/Query_Answers")
os.makedirs(excel_directory, exist_ok=True)

# Dictionary of questions and conditional answers
questions = {
    "name": {
        "question": "1.1. Quel est votre nom et prénom ? 😊",
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
        "question": "Veuillez indiquer le SKU actuel : 🆔",
        "options": []
    },
    "product_code_no": {
        "question": "Veuillez indiquer le SKU précédent ou similaire : 🆔",
        "options": []
    },
    "product_description": {
        "question": "2.3. Fournissez une brève description du produit : 📝 ou description rattachée en automatique?",
        "options": []
    },
    "supplier_conditions": {
        "question": "3.1. Le fournisseur impose-t-il une quantité minimale de commande, ou taille de lot? 📦",
        "options": ["Oui", "Non"]
    },
    "quantity_minimum_yes": {
        "question": "Indiquez la quantité minimale requise : 🔢, ou à négocier? - Y a t il des paliers de prix avec remise possible?",
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
    },
    "final": {
        "question": "Fin du formulaire 🏁",
        "options": []
    }
}

# Function to determine the next question
def get_next_question(answer, previous_question):
    mapping = {
        "name": "email",
        "email": "business_unit",
        "business_unit": "supplier_name",
        "supplier_name": "product_code",
        "product_code": {
            "Oui": "product_code_yes",
            "Non": "product_code_no"
        },
        "product_code_yes": "product_description",
        "product_code_no": "product_description",
        "product_description": "supplier_conditions",
        "supplier_conditions": {
            "Oui": "quantity_minimum_yes",
            "Non": "coverage_duration"
        },
        "quantity_minimum_yes": "coverage_duration",
        "coverage_duration": {
            "Oui": "coverage_duration_yes",
            "Non": "supplier_location"
        },
        "coverage_duration_yes": "supplier_location",
        "supplier_location": "availability_delay",
        "availability_delay": "storage_location",
        "storage_location": "sku_open",
        "sku_open": "sku_frequency",
        "sku_frequency": "dotation",
        "dotation": {
            "Oui": "dotation_yes",
            "Non": "additional_requirements"
        },
        "dotation_yes": "additional_requirements",
        "additional_requirements": "final"
    }
    next_question = mapping.get(previous_question)
    if isinstance(next_question, dict):
        return next_question.get(answer)
    return next_question

# Function to save answers to an Excel file
def save_answers_to_excel(recommendation, ai_recommendation):
    user_name = st.session_state.user_answers.get("name")
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
        for idx, (question, answer) in enumerate(st.session_state.user_answers.items(), start=2):
            sheet.cell(row=idx, column=1, value=question)
            sheet.cell(row=idx, column=2, value=answer)
        # Add recommendations
        sheet.cell(row=len(st.session_state.user_answers) + 2, column=1, value="Recommandations")
        sheet.cell(row=len(st.session_state.user_answers) + 2, column=2, value=recommendation)
        sheet.cell(row=len(st.session_state.user_answers) + 3, column=1, value="Recommandations IA")
        sheet.cell(row=len(st.session_state.user_answers) + 3, column=2, value=ai_recommendation)
        workbook.save(file_path)
        st.success(f"Les réponses ont été enregistrées dans {file_path}")
    except PermissionError:
        st.error(f"Impossible d'enregistrer le fichier. Vérifiez les permissions pour le chemin : {file_path}")

# Function to display recommendations based on answers
def show_recommendation():
    recommendation = "Recommandations :\n"
    product_code = st.session_state.user_answers.get("product_code")
    if product_code == "Non":
        recommendation += "- Assurez-vous de créer un nouveau code dans le système avant de passer commande.\n"
    quantity_minimum = st.session_state.user_answers.get("supplier_conditions")
    if quantity_minimum == "Oui":
        recommendation += "- Recommandez une analyse de consommation historique pour ajuster les hypothèses de réapprovisionnement.\n"
    supplier_location = st.session_state.user_answers.get("supplier_location")
    if supplier_location == "Grand export":
        recommendation += "- Prévoir un délai logistique plus long et anticiper les commandes.\n"
    dotation = st.session_state.user_answers.get("dotation")
    if dotation == "Oui":
        recommendation += "- Priorisez la planification logistique avec le 3PL pour respecter les délais impératifs.\n"

    def get_ai_recommendation(answers):
        try:
            prompt = "Voici les réponses d'un utilisateur à un questionnaire :\n"
            for question, answer in answers.items():
                prompt += f"- {question}: {answer}\n"
            prompt += "Basé sur ces réponses, fournissez des recommandations supplémentaires pertinentes 30 mots max:"
            response = co.generate(prompt=prompt, model="xlarge")
            return response.generations[0].text.strip()
        except Exception as e:
            return f"Erreur lors de la génération des recommandations IA : {str(e)}"

    ai_recommendation = get_ai_recommendation(st.session_state.user_answers)
    recommendation += f"\nRecommandations IA :\n{ai_recommendation}"
    st.text_area("Recommandations", recommendation)
    if st.button("Enregistrer les réponses"):
        save_answers_to_excel(recommendation, ai_recommendation)

# Main Streamlit application
def main():
    st.title("Outil Marketing Survey")
    st.write("Merci de répondre aux questions pour obtenir des recommandations personnalisées.")

    # Initialize session state variables
    if "current_question" not in st.session_state:
        st.session_state.current_question = "name"
    if "user_answers" not in st.session_state:
        st.session_state.user_answers = {}

    current_question_key = st.session_state.current_question
    question_data = questions.get(current_question_key)

    if question_data:
        st.subheader(question_data["question"])
        if question_data["options"]:
            answer = st.radio("Choisissez une option :", question_data["options"], key=current_question_key)
        else:
            answer = st.text_input("Votre réponse :", key=current_question_key)

        if st.button("Suivant"):
            if answer:
                st.session_state.user_answers[current_question_key] = answer
                st.session_state.current_question = get_next_question(answer, current_question_key)
            else:
                st.warning("Veuillez entrer une réponse.")
    else:
        show_recommendation()

if __name__ == "__main__":
    main()
