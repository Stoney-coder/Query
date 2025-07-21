import streamlit as st
import openpyxl
import os
import cohere
from datetime import datetime

# --- CONFIGURATION & CONSTANTS ---

# Use environment variable for API key
COHERE_API_KEY = os.getenv("COHERE_API_KEY")
if not COHERE_API_KEY:
    st.warning("Cohere API key is not set in your environment variables (COHERE_API_KEY). AI recommendations will not work.")

# Initialize Cohere Client if possible
co = cohere.Client(api_key=COHERE_API_KEY) if COHERE_API_KEY else None

# Directory to save Excel files
EXCEL_DIRECTORY = os.path.expanduser("~/Desktop/Query_Answers")
os.makedirs(EXCEL_DIRECTORY, exist_ok=True)

# --- QUESTIONNAIRE DEFINITION ---

questions = {
    # ... (Keep the same dictionary as before)
}

# --- CORE LOGIC FUNCTIONS ---

def get_next_question(answer, previous_question):
    """Determine the next question key based on current answer and previous question."""
    mapping = {
        # ... (Keep the same mapping as before)
    }
    next_question = mapping.get(previous_question)
    if isinstance(next_question, dict):
        return next_question.get(answer)
    return next_question

def save_answers_to_excel(answers, recommendation, ai_recommendation):
    """Save answers and recommendations to an Excel file."""
    user_name = answers.get("name", "utilisateur_inconnu")
    current_date = datetime.now().strftime("%d-%m-%Y")
    file_name = f"{user_name}_{current_date}.xlsx"
    file_path = os.path.join(EXCEL_DIRECTORY, file_name)
    try:
        with openpyxl.Workbook() as workbook:
            sheet = workbook.active
            sheet.title = "Réponses"
            sheet.append(["Clé", "Question", "Réponse"])
            for key, answer in answers.items():
                question_text = questions.get(key, {}).get("question", key)
                sheet.append([key, question_text, answer])
            # Recommendations
            sheet.append(["", "Recommandations", recommendation])
            sheet.append(["", "Recommandations IA", ai_recommendation])
            workbook.save(file_path)
        st.success(f"Réponses enregistrées dans : {file_path}")
    except PermissionError:
        st.error(f"Impossible d'enregistrer le fichier : {file_path}. Vérifiez les permissions.")

def get_ai_recommendation(answers):
    """Generate AI-powered recommendation using Cohere."""
    if not co:
        return "Cohere API key not set. Impossible de générer des recommandations IA."
    try:
        prompt = "Voici les réponses d'un utilisateur à un questionnaire :\n"
        for question, answer in answers.items():
            prompt += f"- {question}: {answer}\n"
        prompt += "Basé sur ces réponses, fournissez des recommandations supplémentaires pertinentes (30 mots max):"
        response = co.generate(prompt=prompt, model="xlarge")
        return response.generations[0].text.strip()
    except Exception as e:
        return f"Erreur lors de la génération des recommandations IA : {str(e)}"

def generate_recommendation(answers):
    """Generate recommendations based on user answers."""
    recommendation = []
    if answers.get("product_code") == "Non":
        recommendation.append("- Créez un nouveau code produit avant la commande.")
    if answers.get("supplier_conditions") == "Oui":
        recommendation.append("- Analysez la consommation historique pour le réapprovisionnement.")
    if answers.get("supplier_location") == "Grand export":
        recommendation.append("- Anticipez un délai logistique plus long.")
    if answers.get("dotation") == "Oui":
        recommendation.append("- Planifiez la logistique avec le 3PL en avance.")
    return "\n".join(recommendation) if recommendation else "Aucune recommandation spécifique."

def show_recommendation():
    """Display recommendations and save feature."""
    answers = st.session_state.user_answers
    recommendation = generate_recommendation(answers)
    ai_recommendation = get_ai_recommendation(answers)
    st.text_area("Recommandations", f"{recommendation}\n\nRecommandations IA :\n{ai_recommendation}", height=200)
    if st.button("Enregistrer les réponses"):
        save_answers_to_excel(answers, recommendation, ai_recommendation)

# --- MAIN APP LOGIC ---

def main():
    st.title("Outil Marketing Survey")
    st.write("Merci de répondre aux questions pour obtenir des recommandations personnalisées.")

    # Session state initialization
    if "current_question" not in st.session_state:
        st.session_state.current_question = "name"
    if "user_answers" not in st.session_state:
        st.session_state.user_answers = {}

    current_key = st.session_state.current_question
    question_data = questions.get(current_key)

    total_questions = len(questions) - 1  # Exclude 'final'
    progress = len(st.session_state.user_answers)

    # Progress indicator
    st.progress(progress / total_questions)

    if question_data and current_key != "final":
        st.subheader(question_data["question"])
        if question_data["options"]:
            answer = st.radio("Choisissez une option :", question_data["options"], key=current_key)
        else:
            answer = st.text_input("Votre réponse :", key=current_key)

        if st.button("Suivant"):
            if answer:
                st.session_state.user_answers[current_key] = answer
                next_q = get_next_question(answer, current_key)
                if next_q:
                    st.session_state.current_question = next_q
                    st.experimental_rerun()
                else:
                    st.session_state.current_question = "final"
                    st.experimental_rerun()
            else:
                st.warning("Veuillez entrer une réponse.")
    else:
        show_recommendation()

if __name__ == "__main__":
    main()
