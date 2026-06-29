import streamlit as st
import pdfplumber
import docx
from google import genai
from supabase import create_client

# ---------------------------------------------------------------------------
# 1. DIRECT DATABASE CONFIGURATION & INITIALIZATION (STABLE VERSION)
# ---------------------------------------------------------------------------
st.set_page_config(page_title="SaaS Insurance Portal", layout="wide")

# Target credentials explicitly pointing to your unique cloud instance
target_url = "https://kdihfmnspnylylobrrmk.supabase.co"
target_key = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtkaWhmbW5zcG55bHlsb2Jycm1rIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODI3NDExNjksImV4cCI6MjA5ODMxNzE2OX0.kMY9NwQhEI3VAsmLPR3y5XDOZ6FBcsW-0Y-2SOPudjs"

# Initialize client simply to bypass the recent ClientOptions package bug
supabase = create_client(target_url, target_key)

# Keep track of active agent login state
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
if 'user_email' not in st.session_state:
    st.session_state['user_email'] = None

# ---------------------------------------------------------------------------
# 2. SAAS ENCRYPTED SHIELD LAYER (Login / Signup Panel)
# ---------------------------------------------------------------------------
if not st.session_state['logged_in']:
    st.title("🛡️ BrokerFlow AI - Insurance Portal")
    st.subheader("Enterprise Proposal Suite")
    
    auth_mode = st.radio("Choose Action", ["Login to Account", "Create Free Agent Account"])
    
    email = st.text_input("Agency Email Address")
    password = st.text_input("Password", type="password")
    
    if auth_mode == "Create Free Agent Account":
        if st.button("Register as Agent"):
            if len(password) < 6:
                st.error("❌ Password must be at least 6 characters long!")
            else:
                try:
                    res = supabase.auth.sign_up({"email": email, "password": password})
                    st.success("✅ Account created successfully! Please select 'Login to Account' above to log in.")
                except Exception as e:
                    st.error(f"Registration failed: {str(e)}")
                
    elif auth_mode == "Login to Account":
        if st.button("Secure Login"):
            try:
                res = supabase.auth.sign_in_with_password({"email": email, "password": password})
                if res.user:
                    st.session_state['logged_in'] = True
                    st.session_state['user_email'] = res.user.email
                    st.rerun()
            except Exception as e:
                st.error("❌ Invalid agency credentials. Please check your spelling or register.")
    st.stop()

# ---------------------------------------------------------------------------
# 3. CORE PROCESSING SUITE (Unlocks strictly after verification matches)
# ---------------------------------------------------------------------------
st.sidebar.markdown(f"👤 **Agent Workspace:**\n{st.session_state['user_email']}")
if st.sidebar.button("Logout"):
    st.session_state['logged_in'] = False
    st.session_state['user_email'] = None
    st.rerun()

st.title("📋 AI-Powered Insurance Proposal Generator")
st.caption("Commercial Grade Production Suite v1.0")

# --- SECURE BACKEND API KEY HARDCODING ---
# Paste your actual secret Gemini API Key between the quotes below:
api_key = "AIzaSyCugsZzoIvXYVSZxUDtEPLDLeKoFJW_joA"

# Initialize the generative AI processing client engine on the back-end
client = genai.Client(api_key=api_key)

def extract_text_from_pdf(file_file):
    text = ""
    with pdfplumber.open(file_file) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    return text

def extract_text_from_docx(file_file):
    doc = docx.Document(file_file)
    return "\n".join([para.text for para in doc.paragraphs])

def extract_text_from_txt(file_file):
    return file_file.read().decode("utf-8")

st.header("Step 1: Input Raw Data & Documents")
uploaded_files = st.file_uploader(
    "Upload Quote PDFs, Emails (.txt), Word Docs, or Raw Notes", 
    type=["pdf", "docx", "txt"], 
    accept_multiple_files=True
)

additional_notes = st.text_area("Add any additional email text, client requirements, or manual notes here:")
extracted_corpus = ""

if uploaded_files or additional_notes:
    st.info("Documents received! Extracting data fields safely...")
    
    for uploaded_file in uploaded_files:
        if uploaded_file.name.endswith('.pdf'):
            extracted_corpus += f"\n--- Start of File: {uploaded_file.name} ---\n" + extract_text_from_pdf(uploaded_file)
        elif uploaded_file.name.endswith('.docx'):
            extracted_corpus += f"\n--- Start of File: {uploaded_file.name} ---\n" + extract_text_from_docx(uploaded_file)
        elif uploaded_file.name.endswith('.txt'):
            extracted_corpus += f"\n--- Start of File: {uploaded_file.name} ---\n" + extract_text_from_txt(uploaded_file)
            
    if additional_notes:
        extracted_corpus += f"\n--- Start of Additional Notes ---\n{additional_notes}"

    st.header("Step 2: AI Data Compilation & Analysis")
    if st.button("Generate Proposal Analysis Options"):
        with st.spinner("Analyzing coverage lines, structural variances, limits, and premiums..."):
            
            prompt = f"""
            You are an expert commercial insurance broker and underwriter. Analyze the following extracted raw insurance data, quotes, emails, or notes.
            
            Tasks:
            1. Compile all available insurance quote options found in the text.
            2. For EACH suitable option, create a clean Markdown Table mapping: Coverage Type, Limits, Deductibles, and Annual Premium.
            3. Provide an Analysis Summary comparing the options side-by-side.
            4. Explicitly state which option is best suited for the client and why (Cost vs. Risk coverage).
            
            Raw Extracted Data:
            {extracted_corpus}
            
            Format the response cleanly using markdown headers (e.g., # Insurance Proposal, ## Option 1, ## Final Recommendation).
            """
            
            try:
                response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
                st.session_state['analysis_result'] = response.text
                st.markdown(response.text)
            except Exception as e:
                st.error(f"Processing Error: {e}")

    if 'analysis_result' in st.session_state:
        st.header("Step 3: Export Proposal")
        
        def create_word_doc(markdown_text):
            doc = docx.Document()
            doc.add_heading('Insurance Proposal Options & Analysis', 0)
            lines = markdown_text.split('\n')
            for line in lines:
                if line.startswith('# '):
                    doc.add_heading(line.replace('# ', ''), level=1)
                elif line.startswith('## '):
                    doc.add_heading(line.replace('## ', ''), level=2)
                elif line.startswith('|'):
                    doc.add_paragraph(line)
                else:
                    if line.strip():
                        doc.add_paragraph(line)
            output_path = "Generated_Insurance_Proposal.docx"
            doc.save(output_path)
            return output_path

        if st.button("Build Word Document"):
            doc_file_path = create_word_doc(st.session_state['analysis_result'])
            with open(doc_file_path, "rb") as f:
                st.download_button(
                    label="📥 Download Word Proposal",
                    data=f,
                    file_name="Insurance_Proposal_Analysis.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
