import streamlit as st
import openai
import pandas as pd
import openpyxl

# Specify the path to your Excel file
#excel_file_path = 'data/planview_data.xlsx'

# Load the workbook
#workbook = openpyxl.load_workbook(excel_file_path)
# Relative path from project root
df = pd.read_excel("data/planview_data.xlsx")

#https://github.com/diwakar998/hbot/blob/cbe096fc062a9cb344c06b861f278966b3a9055e/data/planview_data.xlsx
# Initialize OpenAI with your secret API key
openai.api_key = st.secrets["OPENAI_API_KEY"]

# Set the app title
st.title("🧑‍💼👩‍💼 PMO Reporting & Governance Agent, Your PMO Expert")

# Initialize chat history with a health-focused system prompt
if "messages" not in st.session_state:
    st.session_state.messages = [
        {
            "role": "system",
            "content": (
                "You are a professional PMO AI assistant specialized in Agile, PMI, PMP, Scrum, Spotify, Lean, Kanban, SixSigma and other framework expert"
                "Your role is to help users understand their project status and suggest possible common risk, causes and abbettment plan"
                "If a user asks about anything unrelated to PMO, Project/program, frameworks, Management, risk, resources, reply: "
                "'I'm here to help with PMO related questions and suggestions. "
                "Please ask about project, program, risks, resources or PMO related questions/concerns.' "
            )
        }
    ]

# Display all previous messages
for msg in st.session_state.messages[1:]:  # Skip system prompt in UI
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# User input
user_input = st.chat_input("How can I help you today...")

# Function to get AI response
def get_response(messages):
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=messages
    )
    return response.choices[0].message["content"]

# Process user input
if user_input:
    # Add user's message to history and show
    st.session_state.messages.append({"role": "user", "content": user_input})
    with st.chat_message("user"):
        st.markdown(user_input)

    # Get AI response
    response = get_response(st.session_state.messages)
    
    # Add assistant response to history and show
    st.session_state.messages.append({"role": "assistant", "content": response})
    with st.chat_message("assistant"):
        st.markdown(response)

# Optional: Add footer disclaimer
st.markdown("---")
st.markdown(
   # "🛑 **Disclaimer:** This Agent does not provide medical advice, diagnosis, or treatment. "
   # "Always consult a qualified healthcare provider for medical concerns.",
    unsafe_allow_html=True
)
