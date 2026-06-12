import streamlit as st
import os
import pdfplumber
import docx
import json
import traceback
import re
from dotenv import load_dotenv
from io import BytesIO
import pandas as pd
import base64

# --- CONFIGURATION & API SETUP ---

GROQ_MODEL ="llama-3.3-70b-versatile"              # -----llama-3.1-8b-instant-------
# Load environment variables (e.g., GROQ_API_KEY)
load_dotenv()
GROQ_API_KEY = os.getenv('GROQ_API_KEY')

# --- Default/Mock Data for Filtering ---
DEFAULT_ROLES = ["Data Scientist", "Cloud Engineer", "Software Engineer", "AI/ML Engineer"]
DEFAULT_JOB_TYPES = ["Full-time", "Contract", "Remote"]
STARTER_KEYWORDS = {
    "Python", "MySQL", "GCP", "cloud computing", "ML", 
    "API services", "LLM integration", "JavaScript", "SQL", "AWS", "MLOps", "Data Visualization"
}
# --- Define MockGroqClient globally (Necessary for testing without API Key) ---

class MockGroqClient:
    """Mock client for local testing when Groq is not available or key is missing."""
    def chat(self):
        class Completions:
            def create(self, **kwargs):
                prompt_content = kwargs.get('messages', [{}])[0].get('content', '')
                
                # --- Specific Mock Logic for Interview Prep ---
                if "Generate a list of interview questions" in prompt_content:
                    if "targeting the **JD**" in prompt_content:
                        section = "Cloud Engineer"
                        mock_questions_raw = f"""
                        [Basic/HR-related]
                        Q1: What excites you most about the field of cloud engineering?
                        
                        [Intermediate/Technical]
                        Q2: Explain how you would implement CI/CD for a project involving Docker and Kubernetes.
                        
                        [Advanced/Experience-based]
                        Q3: Describe a time you had to troubleshoot a production issue related to infrastructure automation and the steps you took.
                        
                        [Basic/Situation-based]
                        Q4: How do you handle disagreements with colleagues regarding technical implementation decisions?
                        
                        [Intermediate/Technical]
                        Q5: Explain the core differences between AWS and GCP services related to the JD.
                        """
                    else:
                        # Resume Section Based Mock (targetting skills section)
                        section_match = re.search(r'targeting the \*\*(.+?)\*\* section', prompt_content)
                        section = section_match.group(1).strip() if section_match else "General Skills"
                        mock_questions_raw = f"""
                        [Basic/HR-related]
                        Q1: Why did you choose to specialize in the **{section}** area?
                        
                        [Intermediate/Technical]
                        Q2: Describe a complex technical challenge you overcame in the **{section}** area (e.g., optimizing Python code).
                        
                        [Advanced/Experience-based]
                        Q3: Provide a detailed example of a project where you used your **{section}** skills to achieve a measurable business outcome.
                        
                        [Intermediate/Situation-based]
                        Q4: How would you deal with a tight deadline for a project involving your **{section}** skills?
                        
                        [Advanced/Technical]
                        Q5: How do you keep up to date with the latest trends in **{section}**?
                        """
                    # Return the raw text as expected by the new parser logic
                    return type('MockResponse', (object,), {'choices': [type('Choice', (object,), {'message': type('Message', (object,), {'content': mock_questions_raw})})()]})

                elif "Evaluate the candidate's answers to the following questions" in prompt_content:
                    # Simple mock evaluation logic
                    if "Q2" in prompt_content and "complex technical challenge" in prompt_content:
                        score = 8
                        feedback = "Excellent structure using the STAR method (simulated). You clearly articulated the situation and your actions. **Focus on quantifying the results.**"
                    else:
                        score = 6
                        feedback = "Good technical detail, but the answers were a bit generic (simulated). Try to connect your skills directly to the business impact."

                    mock_evaluation = f"""
                    --- AI Evaluation Report ---
                    
                    **Overall Score:** {score}/10
                    **Summary:** The candidate provided decent technical background but lacked deep, quantifiable examples for most questions. The answer to Q2 was strong. Performance in **HR-related** was good, but **Situation-based** needs improvement.
                    
                    **Q1 (HR-related) Feedback:** {feedback}
                    
                    **Q2 (Technical) Feedback:** Strong response. Excellent use of technical terms and process.
                    
                    **Q3 (Experience-based) Feedback:** Answer was too theoretical. Need a real-world project example.
                    
                    **Q4 (Situation-based) Feedback:** Lacked a clear structured approach to conflict resolution.
                    
                    **Next Steps:** Review the job description and prepare more quantifiable achievements related to this area.
                    """
                    return type('MockResponse', (object,), {'choices': [type('Choice', (object,), {'message': type('Message', (object,), {'content': mock_evaluation})})()]})

                elif "Generate a detailed course plan and suggest relevant certifications" in prompt_content:
                    gap_match = re.search(r'Gaps Identified:\s*(.*)', prompt_content, re.DOTALL)
                    gap_summary = gap_match.group(1).strip() if gap_match else "Missing key skills in Cloud and CI/CD."
                    
                    mock_plan = f"""
                    ## 💡 Detailed Course Plan: Addressing Gaps in Cloud/CI/CD (Simulated)
                    
                    The goal is to cover the identified gaps: **{gap_summary}**.
                    
                    ### Phase 1: Foundational Cloud Skills (4 Weeks)
                    * **Module 1 (AWS/GCP):** Core services (EC2, S3, IAM, VPC). Focus on security best practices.
                    * **Module 2 (IaC):** Introduction to **Terraform** or CloudFormation/Deployment Manager. Hands-on simple infrastructure provisioning.
                    
                    ### Phase 2: Automation & DevOps (6 Weeks)
                    * **Module 3 (CI/CD Principles):** Theory and practice of continuous integration/delivery using **GitLab CI** or Jenkins.
                    * **Module 4 (Containerization):** Advanced Dockerfile creation and multi-container application deployment with Docker Compose.
                    * **Module 5 (Kubernetes Basics):** Deploying and scaling applications using basic K8s objects (Pods, Deployments, Services).
                    
                    ### Phase 3: Project and Certification Prep (4 Weeks)
                    * **Project:** Build a fully automated CI/CD pipeline deploying a microservice to a managed Kubernetes cluster (EKS/GKE).
                    
                    ---
                    
                    ## 🏅 Suggested Certifications
                    
                    * **For AWS Focus:** **AWS Certified Solutions Architect – Associate** (Covers broad cloud knowledge).
                    * **For GCP Focus:** **Google Cloud Professional Cloud Architect** (A high-value certification).
                    * **For DevOps/CI/CD:** **Certified Kubernetes Administrator (CKA)** or **HashiCorp Certified Terraform Associate**.
                    
                    ---
                    **Next Step:** Focus on the **AWS Certified Solutions Architect** path first, as it provides the quickest return on investment for entry to mid-level cloud roles.
                    """
                    return type('MockResponse', (object,), {'choices': [type('Choice', (object,), {'message': type('Message', (object,), {'content': mock_plan})})()]}) 

                elif "Answer the following question about the Job Description concisely and directly." in prompt_content:
                    question_match = re.search(r'Question:\s*(.*)', prompt_content)
                    question = question_match.group(1).strip() if question_match else "a question"
                    
                    if 'role' in question.lower():
                        return type('MockResponse', (object,), {'choices': [type('Message', (object,), {'content': 'The required role in this Job Description is Cloud Engineer.'})()]})
                    elif 'experience' in question.lower():
                        return type('MockResponse', (object,), {'choices': [type('Message', (object,), {'content': 'The job requires 3+ years of experience in AWS/GCP and infrastructure automation.'})()]})
                    else:
                        return type('MockResponse', (object,), {'choices': [type('Message', (object,), {'content': 'Mock answer for JD question: The JD mentions Python and Docker as key skills.'})()]})

                elif "Answer the following question about the resume concisely and directly." in prompt_content:
                    question_match = re.search(r'Question:\s*(.*)', prompt_content)
                    question = question_match.group(1).strip() if question_match else "a question"
                    
                    if 'name' in question.lower():
                        return type('MockResponse', (object,), {'choices': [type('Message', (object,), {'content': 'The candidate\'s name is Vivek Swamy.'})()]})
                    elif 'skills' in question.lower():
                        return type('MockResponse', (object,), {'choices': [type('Message', (object,), {'content': 'Key skills include Python, SQL, AWS, and MLOps.'})()]})
                    else:
                        return type('MockResponse', (object,), {'choices': [type('Message', (object,), {'content': f'Based on the mock resume data, I can provide a simulated answer to your question about {question}.'})()]})

                elif "You are an expert cover letter generator" in prompt_content:
                    role_match = re.search(r'Job Description Role: (.*?)[\.\n]', prompt_content)
                    role = role_match.group(1).strip() if role_match else "Software Engineer"
                    
                    mock_cover_letter = f"""
                    [Date]
                    
                    [Hiring Manager Name/Title, if known]
                    [Company Name]
                    
                    **Subject: Application for {role} Position - Vivek Swamy**
                    
                    Dear Hiring Manager,
                    
                    I am writing to express my enthusiastic interest in the **{role}** position at MockCorp, as detailed in the attached job description. My background, highlighted by strong skills in Python, AWS, and MLOps, aligns perfectly with your requirements for [Key Requirement from JD - e.g., cloud infrastructure management].
                    
                    During my time at Test Corp (simulated experience), I was responsible for [specific achievement related to JD]. My resume further details my proficiency in [Skill 1] and [Skill 2], which I believe would make me an immediate asset to your team.
                    
                    I am confident in my ability to contribute to your company's goals and I look forward to the opportunity to discuss my application further.
                    
                    Sincerely,
                    
                    Vivek Swamy
                    [vivek.swamy@example.com]
                    """
                    return type('MockResponse', (object,), {'choices': [type('Choice', (object,), {'message': type('Message', (object,), {'content': mock_cover_letter})})()]})
                
          # Mock candidate data (Vivek Swamy) for parsing
                mock_llm_json = {
                    "name": "Vivek Swamy", 
                    "email": "vivek.swamy@example.com", 
                    "phone": "555-1234", 
                    "linkedin": "https://linkedin.com/in/vivek-swamy-mock", 
                    "github": "https://github.com/vivek-mock", 
                    "personal_details": "Mock summary generated for: Vivek Swamy.", 
                    "skills": [
                        "Python", "SQL", "AWS", "Streamlit", 
                        "LLM Integration", "MLOps", "Data Visualization", 
                        "Docker", "Kubernetes", "Java", "API Services" 
                    ], 
                    "education": ["B.S. Computer Science, Mock University, 2020"], 
                    "experience": ["Software Intern, Mock Solutions (2024-2025)", "Data Analyst, Test Corp (2022-2024)"], 
                    "certifications": ["Mock Certification in AWS Cloud"], 
                    "projects": ["Mock Project: Built an MLOps pipeline using Docker and Kubernetes."], 
                    "strength": ["Mock Strength"], 
                }                
                
                # Mock response content for GroqClient initialization check (for parsing)
                message_obj = type('Message', (object,), {'content': json.dumps(mock_llm_json)})()
                choice_obj = type('Choice', (object,), {'message': message_obj})()
                response_obj = type('MockResponse', (object,), {'choices': [choice_obj]})()
                return response_obj
        
        # Add a placeholder for the completions object if we need a mock response for fit evaluation
        class FitCompletions(Completions):
            def create(self, **kwargs):
                prompt_content = kwargs.get('messages', [{}])[0].get('content', '')
                
                if "Evaluate how well the following resume content matches the provided job description" in prompt_content:
                    # SIMULATED FIT LOGIC (Fallback for when the LLM-dependent function tries to run without a key)
                    
                    # Simple heuristic mock score based on role title in the prompt
                    jd_role_match = re.search(r'(?:Role|Engineer|Scientist)[:\s]+([\w\s/-]+)', prompt_content)
                    jd_role = jd_role_match.group(1).lower().strip() if jd_role_match else "default"
                    
                    if 'ai/ml' in jd_role or 'mlops' in jd_role:
                        score = 8
                    elif 'data scientist' in jd_role:
                        score = 7
                    elif 'cloud engineer' in jd_role:
                        score = 6
                    else:
                        score = 5
                        
                    # Calculate percentages based on the score to differentiate the rows
                    skills_p = 50 + (score * 5)
                    exp_p = 60 + (score * 3)
                    edu_p = 70 + (score * 1)
                    
                    # NOTE: This mock output uses the strict format expected by the regex parser below.
                    mock_fit_output = f"""
                    Overall Fit Score: {score}/10
                    
                    --- Section Match Analysis ---
                    Skills Match: {skills_p}%
                    Experience Match: {exp_p}%
                    Education Match: {edu_p}%
                    
                    Strengths/Matches:
                    - Mock Match Point 1 (Role: {jd_role})
                    - Mock Match Point 2
                    
                    Gaps/Areas for Improvement:
                    - Missing hands-on experience in **Terraform**.
                    - Lack of project experience deploying applications to **GCP/EKS**.
                    - Weak documentation skills in CI/CD pipeline development.
                    
                    Overall Summary: Mock summary for score {score}.
                    """
                    return type('MockResponse', (object,), {'choices': [type('Choice', (object,), {'message': type('Message', (object,), {'content': mock_fit_output})})()]})
                
                # If it's not a fit evaluation, run standard Completions logic
                return super().create(**kwargs)

        return FitCompletions()
        
# Initialize the Groq client or the Mock client based on the environment variable
try:
    from groq import Groq
    
    if GROQ_API_KEY:
        client = Groq(api_key=GROQ_API_KEY)
        # Custom flag to indicate a successful connection attempt to the real client
        class GroqPlaceholder(Groq): 
             def __init__(self, api_key): 
                 super().__init__(api_key=api_key)
                 self.client_ready = True
        client = GroqPlaceholder(api_key=GROQ_API_KEY)
    else:
        # Fallback if key is missing but Groq is installed
        raise ValueError("GROQ_API_KEY not set. Using Mock Client.")
        
except (ImportError, ValueError, NameError) as e:
    # Fallback to Mock Client if import fails or key is missing
    client = MockGroqClient()
    
# --- END API SETUP ---


# --- Utility Functions ---

def clear_interview_state(mode):
    """Clears all session state variables related to interview preparation for a specific mode."""
    if mode == 'resume':
        if 'iq_output_resume' in st.session_state: del st.session_state['iq_output_resume']
        if 'interview_qa_resume' in st.session_state: del st.session_state['interview_qa_resume']
        if 'evaluation_report_resume' in st.session_state: del st.session_state['evaluation_report_resume']
    elif mode == 'jd':
        if 'iq_output_jd' in st.session_state: del st.session_state['iq_output_jd']
        if 'interview_qa_jd' in st.session_state: del st.session_state['interview_qa_jd']
        if 'evaluation_report_jd' in st.session_state: del st.session_state['evaluation_report_jd']
    
    # Also clear the gap analysis plan when interview state is cleared (as it's derived from the match)
    if 'gap_analysis_plan' in st.session_state: del st.session_state['gap_analysis_plan']


def get_file_type(file_name):
    """Identifies the file type based on its extension, handling common text formats."""
    ext = os.path.splitext(file_name)[1].lower().strip('.')
    if ext == 'pdf': return 'pdf'
    elif ext in ('docx', 'doc'): return 'docx'
    elif ext in ('txt', 'md', 'markdown', 'rtf'): return 'txt' 
    elif ext == 'json': return 'json'
    elif ext in ('xlsx', 'xls', 'csv'): return 'excel' 
    else: return 'unknown' 

def extract_content(file_type, file_content_bytes, file_name):
    """Extracts text content from uploaded file content (bytes)."""
    text = ''
    excel_data = None
    try:
        if file_type == 'pdf':
            with pdfplumber.open(BytesIO(file_content_bytes)) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + '\n'
        
        elif file_type == 'docx':
            doc = docx.Document(BytesIO(file_content_bytes))
            text = '\n'.join([para.text for para in doc.paragraphs])
        
        elif file_type == 'txt':
            try:
                # Try UTF-8 first, fallback to Latin-1
                text = file_content_bytes.decode('utf-8')
            except UnicodeDecodeError:
                 text = file_content_bytes.decode('latin-1')
        
        elif file_type == 'json':
            try:
                text = file_content_bytes.decode('utf-8')
                text = "--- JSON Content Start ---\n" + text + "\n--- JSON Content End ---"
            except UnicodeDecodeError:
                return f"[Error] JSON content extraction failed: Unicode Decode Error.", None
        
        elif file_type == 'excel':
            try:
                if file_name.endswith('.csv'):
                    df = pd.read_csv(BytesIO(file_content_bytes))
                else: 
                    xls = pd.ExcelFile(BytesIO(file_content_bytes))
                    all_sheets_data = {}
                    for sheet_name in xls.sheet_names:
                        df = pd.read_excel(xls, sheet_name=sheet_name)
                        # Store as JSON strings for LLM input
                        all_sheets_data[sheet_name] = df.to_json(orient='records') 
                        
                    excel_data = all_sheets_data 
                    text = json.dumps(all_sheets_data, indent=2)
                    text = f"[EXCEL_CONTENT] The following structured data was extracted:\n{text}"
                    
            except Exception as e:
                return f"[Error] Excel/CSV file parsing failed. Error: {e}", None

        if not text.strip() and file_type not in ('excel', 'json'): 
            return f"[Error] {file_type.upper()} content extraction failed or file is empty.", None
        
        return text, excel_data
    
    except Exception as e:
        return f"[Error] Fatal Extraction Error: Failed to read file content ({file_type}). Error: {e}\n{traceback.format_exc()}", None

@st.cache_data(show_spinner="Analyzing content with Groq LLM...")
def parse_resume_with_llm(text):
    """
    Sends resume text to the LLM for structured information extraction.
    """
    
    def get_fallback_name():
        return "Vivek Swamy" 

    if text.startswith("[Error"):
        return {"name": "Parsing Error", "error": text}

    json_match_external = re.search(r'--- JSON Content Start ---\s*(.*?)\s*--- JSON Content End ---', text, re.DOTALL)
    
    if json_match_external:
        try:
            json_content = json_match_external.group(1).strip()
            parsed_data = json.loads(json_content)
            
            if not parsed_data.get('name'):
                parsed_data['name'] = get_fallback_name()
                 
            parsed_data['error'] = None 
            
            return parsed_data
        
        except json.JSONDecodeError:
            return {"name": get_fallback_name(), "error": f"LLM Input Error: Could not decode uploaded JSON content into a valid structure."}
            
    if isinstance(client, MockGroqClient) or not GROQ_API_KEY:
        try:
            completion = client.chat().create(model=GROQ_MODEL, messages=[{}])
            content = completion.choices[0].message.content.strip()
            parsed_data = json.loads(content)
            
            if not parsed_data.get('name'):
                parsed_data['name'] = get_fallback_name()
            
            parsed_data['error'] = None 
            return parsed_data
            
        except Exception as e:
            return {"name": get_fallback_name(), "error": f"Mock Client Error: {e}"}
    
    prompt = f"""Extract the following information from the resume in structured JSON.
    Ensure all relevant details for each category are captured.
    - Name, - Email, - - Phone, - Skills (list), - Education (list of degrees/institutions/dates), 
    - Experience (list of job roles/companies/dates/responsibilities), - Certifications (list), 
    - Projects (list of project names/descriptions/technologies), - Strength (list of personal strengths/qualities), 
    - Personal Details (e.g., address, date of birth, nationality), - Github (URL), - LinkedIn (URL)
    
    Resume Text:
    {text}
    
    Provide the output strictly as a JSON object.
    """
    content = ""
    parsed = {}
    json_str = ""
    
    try:
        response = client.chat.completions.create( 
            model=GROQ_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2,
            response_format={"type": "json_object"}
        )
        content = response.choices[0].message.content.strip()

        json_match = re.search(r'\{.*\}', content, re.DOTALL)
        
        if json_match:
            json_str = json_match.group(0).strip()
            
            if json_str.startswith('```json'):
                json_str = json_str[len('```json'):]
            if json_str.endswith('```'):
                json_str = json_str[:-len('```')]
            
            json_str = json_str.strip()
            parsed = json.loads(json_str)
        else:
            raise json.JSONDecodeError("Could not isolate a valid JSON structure from LLM response.", content, 0)
        
        if not parsed.get('name'):
            parsed['name'] = get_fallback_name()
            
        parsed['error'] = None 
        return parsed

    except json.JSONDecodeError as e:
        error_msg = f"JSON decoding error from LLM. LLM returned malformed JSON. Error: {e} | Malformed string segment:\n---\n{json_str[:200]}..."
        return {"name": get_fallback_name(), "error": error_msg}
        
    except Exception as e:
        error_msg = f"LLM API interaction error: {e}"
        return {"name": get_fallback_name(), "error": error_msg}
        
# Updated signature to match the request
def parse_and_store_resume(content_source, file_name_key, source_type):
    """Handles extraction, parsing, and storage of CV data from either a file or pasted text."""
    extracted_text = ""
    excel_data = None
    file_name = "Pasted_Resume"

    if source_type == 'file':
        uploaded_file = content_source
        file_name = uploaded_file.name
        file_type = get_file_type(file_name)
        uploaded_file.seek(0) 
        st.session_state.current_parsing_source_name = file_name 
        extracted_text, excel_data = extract_content(file_type, uploaded_file.getvalue(), file_name)
    elif source_type == 'text':
        extracted_text = content_source.strip()
        file_name = "Pasted_Text"
        st.session_state.current_parsing_source_name = file_name 
    elif source_type == 'compiled':
        # Used for CV Management tab, content_source is already the compiled markdown
        extracted_text = content_source.strip()
        file_name = "Form_Compiled_CV"
        st.session_state.current_parsing_source_name = file_name

    if extracted_text.startswith("[Error"):
        return {"error": extracted_text, "full_text": extracted_text, "excel_data": None, "name": file_name}
    
    parsed_data = parse_resume_with_llm(extracted_text)
    
    if parsed_data.get('error') is not None: 
        error_name = parsed_data.get('name', file_name) 
        return {"error": parsed_data['error'], "full_text": extracted_text, "excel_data": excel_data, "name": error_name}

    compiled_text = ""
    for k, v in parsed_data.items():
        if v and k not in ['error']:
            compiled_text += f"## {k.replace('_', ' ').title()}\n\n"
            if isinstance(v, list):
                compiled_text += "\n".join([f"* {str(item)}" for item in v]) + "\n\n"
            else:
                compiled_text += str(v) + "\n\n"

    final_name = parsed_data.get('name', 'Unknown_Candidate').replace(' ', '_') 
    
    return {
        "parsed": parsed_data, 
        "full_text": compiled_text, 
        "excel_data": excel_data, 
        "name": final_name
    }


def get_download_link(data, filename, file_format, title="Parsed Data"):
    """
    Generates a base64 encoded download link for the given data and format.
    """
    mime_type = "application/octet-stream"
    
    if file_format in ('json', 'markdown', 'text'):
        data_bytes = data.encode('utf-8')
        if file_format == 'json':
            mime_type = "application/json"
        elif file_format == 'markdown':
            mime_type = "text/markdown"
        else:
            mime_type = "text/plain"
            
    elif file_format == 'html':
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>{filename}</title>
            <style>
                body {{ font-family: Arial, sans-serif; padding: 40px; line-height: 1.6; max-width: 800px; margin: auto; }}
                h1 {{ color: #1E90FF; border-bottom: 2px solid #ddd; padding-bottom: 10px; }}
                h2 {{ color: #333; margin-top: 20px; }}
                pre, .cover-letter {{ white-space: pre-wrap; word-wrap: break-word; background: #f4f4f4; padding: 15px; border: 1px solid #ddd; border-radius: 5px; }}
                p {{ margin-bottom: 15px; }}
            </style>
        </head>
        <body>
        <h1>{title}: {filename.replace('.html', '')}</h1>
        <hr/>
        <div class="cover-letter">
        {data.replace('\n', '<br>')}
        </div>
        <p style="margin-top: 30px; font-size: 10px; color: grey;">Generated by PragyanAI</p>
        </body>
        </html>
        """
        data_bytes = html_content.encode('utf-8')
        mime_type = "text/html"
    else:
        return "" 

    b64 = base64.b64encode(data_bytes).decode()
    return f"data:{mime_type};base64,{b64}"


def render_download_button(data_uri, filename, label, color):
    """Renders an HTML button that triggers a file download."""
    if color == 'json':
        bg_color = "#4CAF50"
        icon = "💾"
    elif color == 'markdown':
        bg_color = "#008CBA"
        icon = "⬇️"
    elif color == 'html':
        bg_color = "#f44336"
        icon = "📄"
    elif color == 'cover':
        bg_color = "#FFC300"
        icon = "✉️"
    else:
        bg_color = "#555555"
        icon = ""
        
    st.markdown(
        f"""
        <a href="{data_uri}" download="{filename}" style="text-decoration: none;">
            <button style="
                background-color: {bg_color}; 
                color: white; 
                border: none; 
                padding: 10px 10px; 
                text-align: center; 
                text-decoration: none; 
                display: inline-block; 
                font-size: 14px; 
                margin: 4px 0; 
                cursor: pointer; 
                border-radius: 4px;
                width: 100%;">
                {icon} {label}
            </button>
        </a>
        """, 
        unsafe_allow_html=True
    )
# --- END HELPER FUNCTIONS ---


# --- LLM Functions (Used across tabs) ---

@st.cache_data(show_spinner="Analyzing JD with Groq LLM...")
def extract_jd_metadata(jd_text):
    """
    Extracts high-accuracy metadata (Role, Skills, Job Type) from raw Job Description text 
    using Groq's structured JSON output capabilities.
    """
    global client, GROQ_MODEL, GROQ_API_KEY
    
    # 1. Input Validation and Guard Rails
    if isinstance(jd_text, str) and jd_text.startswith("[Error"):
        return {"role": "Extraction Error", "key_skills": ["Error"], "job_type": "Error"}
    
    if not isinstance(jd_text, str):
        jd_text = str(jd_text)
        
    if not jd_text.strip():
        return {"role": "Empty JD", "key_skills": [], "job_type": "N/A"}

    # 2. Structured Prompt Engineering
    prompt = f"""
    You are an expert HR data parsing system. Your task is to analyze the following Job Description text and extract key metadata into a structured JSON format.
    
    Expected JSON Structure:
    {{
        "role": "The explicit title of the position (e.g., 'AI/ML Engineer', 'Full-Stack Developer')",
        "job_type": "The arrangement. Must be exactly one of: 'Full-time', 'Part-time', 'Contract', 'Remote', 'Hybrid'. Default to 'Full-time' if not mentioned.",
        "key_skills": ["List of technical skills, tools, languages, frameworks, or libraries mentioned explicitly. Use standard capitalizations (e.g., 'Python', 'Streamlit', 'MLOps', 'PyTorch')."]
    }}

    Job Description Content:
    {jd_text}

    Provide the output strictly as a valid JSON object matching the schema above. Do not wrap the JSON in markdown code blocks like ```json ... ```.
    """

    # 3. Fallback Heuristic Handling (If running locally via MockGroqClient or API key is missing)
    if isinstance(client, MockGroqClient) or not GROQ_API_KEY:
        jd_lower = jd_text.lower()
        if 'data scientist' in jd_lower or 'machine learning' in jd_lower:
            role = "Data Scientist/ML Engineer"
        elif 'cloud engineer' in jd_lower or 'aws' in jd_lower:
            role = "Cloud Engineer"
        elif 'ai' in jd_lower or 'ml' in jd_lower:
            role = "AI/ML Engineer"
        else:
            role = "Software Engineer"
            
        skills_found = [s for s in ["Python", "SQL", "AWS", "Docker", "Kubernetes", "Streamlit", "MLOps", "PyTorch"] if s.lower() in jd_lower]
        job_type = "Full-time"
        for t in ["Part-time", "Contract", "Remote", "Hybrid"]:
            if t.lower() in jd_lower:
                job_type = t
                break
                
        return {
            "role": role,
            "key_skills": skills_found if skills_found else ["General Software Engineering"],
            "job_type": job_type
        }

    # 4. Live API Execution Block
    try:
        response = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.1,  # Low temperature guarantees deterministic, analytical extraction
            response_format={"type": "json_object"}
        )
        
        content = response.choices[0].message.content.strip()
        parsed_metadata = json.loads(content)
        
        # Post-processing structural validation
        if not parsed_metadata.get('role'):
            parsed_metadata['role'] = "Target Role"
        if not isinstance(parsed_metadata.get('key_skills'), list):
            parsed_metadata['key_skills'] = []
            
        return parsed_metadata

    except json.JSONDecodeError:
        return {
            "role": "Extraction Error", 
            "key_skills": ["Failed to decode structured metadata"], 
            "job_type": "N/A"
        }
    except Exception as e:
        return {
            "role": "API Error", 
            "key_skills": [f"Connection failed: {str(e)}"], 
            "job_type": "N/A"
        }  


# --- Evaluation JD Fit ---
def evaluate_jd_fit(job_description, parsed_json):
    """
    Evaluates how well a resume fits a given job description, 
    including section-wise scores, by calling the Groq LLM API.
    """
    global client, GROQ_MODEL, GROQ_API_KEY
    
    if parsed_json.get('error') is not None: 
        return f"Cannot evaluate due to resume parsing errors: {parsed_json['error']}"

    if isinstance(client, MockGroqClient) or not GROQ_API_KEY:
        # Mock Client is hardcoded to return a structured output including Gaps.
        response = client.chat().create(model=GROQ_MODEL, messages=[{"role": "user", "content": f"Evaluate how well the following resume content matches the provided job description: {job_description}"}])
        return response.choices[0].message.content.strip()

    if not job_description.strip(): 
        return "Please paste a job description."

    relevant_resume_data = {
        'Skills': parsed_json.get('skills', 'Not found or empty'),
        'Experience': parsed_json.get('experience', 'Not found or empty'),
        'Education': parsed_json.get('education', 'Not found or empty'),
    }
    resume_summary = json.dumps(relevant_resume_data, indent=2)

    prompt = f"""Evaluate how well the following resume content matches the provided job description.
    
    Job Description: {job_description}
    
    Resume Sections for Analysis:
    {resume_summary}
    
    Provide a detailed evaluation structured as follows:
    1.  **Overall Fit Score:** A score out of 10.
    2.  **Section Match Percentages:** A percentage score for the match in the key sections (Skills, Experience, Education).
    3.  **Strengths/Matches:** Key points where the resume aligns well with the JD.
    4.  **Gaps/Areas for Improvement:** Key requirements in the JD that are missing or weak in the resume. Focus on specific technical skills or experience areas.
    5.  **Overall Summary:** A concise summary of the fit.
    
    **Format the output strictly as follows, ensuring the scores are easily parsable (use brackets or no brackets around scores, but they must be present):**
    Overall Fit Score: [Score]/10
    
    --- Section Match Analysis ---
    Skills Match: [XX]%
    Experience Match: [YY]%
    Education Match: [ZZ]%
    
    Strengths/Matches:
    - Point 1
    - Point 2
    
    Gaps/Areas for Improvement:
    - Point 1 (Specific Skill/Experience Gap)
    - Point 2 (Specific Skill/Experience Gap)
    
    Overall Summary: [Concise summary]
    """

    try:
        response = client.chat.completions.create(
            model=GROQ_MODEL, 
            messages=[{"role": "user", "content": prompt}], 
            temperature=0.4
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        error_output = f"AI Evaluation Error: Failed to connect or receive response from LLM. Error: {e}\n{traceback.format_exc()}"
        return error_output
        
# ATS  resume Score -------------------
import streamlit as st
import re
import io
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# --- Ensure required state variables exist globally ---
if "ats_score_calculated" not in st.session_state:
    st.session_state.ats_score_calculated = False
if "ats_score_metrics" not in st.session_state:
    st.session_state.ats_score_metrics = {}
if "ats_original_resume_text" not in st.session_state:
    st.session_state.ats_original_resume_text = ""
if "ats_job_description_text" not in st.session_state:
    st.session_state.ats_job_description_text = ""
if "ats_optimized_resume_text" not in st.session_state:
    st.session_state.ats_optimized_resume_text = ""
if "last_uploaded_file_name" not in st.session_state:
    st.session_state.last_uploaded_file_name = None
if "last_uploaded_jd_name" not in st.session_state:
    st.session_state.last_uploaded_jd_name = None


def optimize_resume_for_ats(resume_text, jd_text, report_metrics):
    """Queries the Groq API to convert the raw profile into a tailored, scanner-compliant resume."""
    global client, GROQ_MODEL, GROQ_API_KEY
    
    if isinstance(client, MockGroqClient) or not GROQ_API_KEY:
        return f"# OPTIMIZED ATS RESUME\n\n{resume_text}\n\n*Note: Add explicit tech keywords to finalize structural tuning.*"

    feedback_context = f"""
    - Personal Info Status: {report_metrics.get('personal_info')}
    - Summary Section Status: {report_metrics.get('summary')}
    - Skills Block Status: {report_metrics.get('skills')}
    - Section Headers Compliance: {report_metrics.get('titles')}
    - Work Experience Layout Strategy: {report_metrics.get('exp_structure')}
    - Work Experience Metric Content: {report_metrics.get('exp_content')}
    - Education Timeline Validation: {report_metrics.get('education_grade')}
    - Engineering Projects Validation: {report_metrics.get('projects_grade')}
    """

    prompt = f"""
    You are an elite expert technical recruiter and specialized ATS compliance scanner engineer.
    Your objective is to ingest the candidate's raw profile text and rewrite it completely to hit a 95%+ pass rating on corporate parser scrapers by resolving the explicit issues identified in the audit report.
    
    {"If provided, optimize it specifically to match this Job Description:" + jd_text if jd_text.strip() else ""}
    
    --- Candidate Raw Resume ---
    {resume_text}
    
    --- Hiring Manager Audit Context (Fix Every Section Marked '0%' or 'Deficient') ---
    {feedback_context}
    
    --- ATS Architectural Requirements ---
    1. Structure the layout cleanly using crisp standard Markdown headers (e.g., # Name, ## Professional Summary, ## Core Technical Skills, ## Professional Experience, ## Education, ## Projects).
    2. Convert all vague descriptions or tasks into impact metrics and action-driven bullet paths (use phrases starting with 'Engineered', 'Optimized', 'Architected', 'Spearheaded' and weave in explicit quantified indicators like %, $, or hours saved where applicable).
    3. Remove all non-standard elements like embedded charts, script symbols, layout tables, columns, sidebars, or progress bar gauges. Convert these strictly into clean, linear chronologies.
    4. Inject clear, standardized technical industry standard keyword terminology so machine search queries flag the profile instantly.
    5. Ensure all list elements use a clean plain text bullet character. Do not use '+', '-', or '*' indicators inside the raw final text payload.
    
    Provide ONLY the completely rewritten, structural Markdown text of the optimized resume. Do not include chat introductory prefaces, greeting notes, meta-commentary, or markdown code fences like ```markdown.
    """
    try:
        response = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.4
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"ATS Generation Error: Failed to re-architect profile matrix context details. Detail: {str(e)}"

#------- download ing in pdf---------------
def generate_pdf_bytes(resume_text):
    """Compiles a cleanly styled, single-column, highly machine-scannable true PDF block using pure ReportLab."""
    buffer = io.BytesIO()
    
    # Establish document blueprint margins optimized for typical parser scanners
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=54,  # 0.75 in
        leftMargin=54,
        topMargin=54,
        bottomMargin=54
    )
    
    styles = getSampleStyleSheet()
    
    # Define sharp, linear, high-opacity scannable text constraints
    body_style = ParagraphStyle(
        'ATSPdfBody',
        parent=styles['Normal'],
        fontName='Times-Roman',
        fontSize=10.5,
        leading=15,
        textColor=colors.HexColor('#111111'),
        spaceAfter=6
    )
    
    story = []
    
    # Parse individual line arrays to structural document flows
    lines = resume_text.split('\n')
    for line in lines:
        cleaned_line = line.strip()
        if not cleaned_line:
            story.append(Spacer(1, 8))
            continue
            
        # Format headers natively to maintain readable typography
        if cleaned_line.startswith('##'):
            header_text = cleaned_line.replace('##', '').strip()
            header_style = ParagraphStyle('H2', fontName='Times-Bold', fontSize=13, leading=18, spaceBefore=12, spaceAfter=6, textColor=colors.HexColor('#222222'))
            story.append(Paragraph(f"<b>{header_text}</b>", header_style))
        elif cleaned_line.startswith('#'):
            header_text = cleaned_line.replace('#', '').strip()
            header_style = ParagraphStyle('H1', fontName='Times-Bold', fontSize=18, leading=22, spaceBefore=4, spaceAfter=8, textColor=colors.HexColor('#111111'))
            story.append(Paragraph(f"<b>{header_text}</b>", header_style))
        else:
            story.append(Paragraph(cleaned_line, body_style))
            
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()
    
# --- Cover Letter Generator Helpers ---
# --- Cover Letter Generator Helpers ---
def extract_basic_entities(resume_text, jd_content):
    """Safely extracts candidate names, target roles, and core skill sets from raw inputs."""
    # 1. Candidate Name Extraction Heuristic
    lines = [line.strip() for line in resume_text.split('\n') if line.strip()]
    cand_name = "Candidate Name"
    if lines:
        potential_name = lines[0].strip('#*[] ')
        if len(potential_name) < 40 and not any(kw in potential_name.lower() for kw in ['resume', 'cv', 'experience', 'education']):
            cand_name = potential_name

    # 2. Target Role Title Extraction Heuristic
    role_title = "Technical Specialist"
    role_match = re.search(r'(?:Role|Position|Title|Job Title)[:\s\n]+([\w\s/-]+)', jd_content, re.IGNORECASE)
    if role_match:
        role_title = role_match.group(1).strip()
    elif 'data scientist' in jd_content.lower():
        role_title = "Data Scientist"
    elif 'ai/ml' in jd_content.lower() or 'machine learning' in jd_content.lower():
        role_title = "AI/ML Engineer"
    elif 'cloud engineer' in jd_content.lower():
        role_title = "Cloud Engineer"

    # Clean character overrides out of structural variable names
    role_title = role_title.replace('#', '').replace('*', '').replace('[', '').replace(']', '').strip()
    cand_name = cand_name.replace('#', '').replace('*', '').replace('[', '').replace(']', '').strip()

    # 3. Core Tech Stack Extraction Heuristic
    skills_inventory = ["Python", "Pandas", "NumPy", "SQL", "Streamlit", "Docker", "Kubernetes", "AWS", "GCP", "Scikit-Learn"]
    extracted_skills = [skill for skill in skills_inventory if skill.lower() in resume_text.lower()]
    skills_phrase = ", ".join(extracted_skills[:4]) if extracted_skills else "software engineering principles and modern frameworks"

    return cand_name, role_title, skills_phrase


# ----------- Template Compiler -----------
def compile_static_template(resume_text, jd_content, template_style):
    """Compiles structurally sound cover letter blueprints natively using candidate context details."""
    cand_name, role_title, skills_phrase = extract_basic_entities(resume_text, jd_content)
    
    # 1. Simple Template Option Blueprint
    if template_style == "Simple":
        return f"""[Date]

Hiring Manager
[Company Name]

Subject: Application for {role_title} Position - {cand_name}

Dear Hiring Manager,

Please accept this letter as formal expression of my interest in the {role_title} position currently open at your company. My background includes technical training combined with hands-on software design work utilizing tools like {skills_phrase}.

Through independent project execution, I have built web applications from structural database setups down to final production tracking systems. I specialize in troubleshooting software complexities, writing maintainable logic configurations, and quickly mastering new development environments.

I am eager to apply my skills to your active engineering objectives. Thank you for your review and evaluation of my attached application documentation.

Sincerely,

{cand_name}"""

    # 2. Professional Template Option Blueprint
    elif template_style == "Professional":
        return f"""[Date]

Hiring Manager
[Company Name]
[Company Address]

Subject: Application for {role_title} - {cand_name}

Dear Hiring Manager,

I am writing to express my strong interest in the {role_title} position at your organization. Given the production parameters and technical criteria outlined in your job specification document, I am confident that my technical capabilities match your engineering needs closely.

My practical execution experience is centered around building robust code layers and automating data workflows. I have practical experience implementing, testing, and maintaining software apps using {skills_phrase}. Managing systems across complete development files has trained me to systematically debug performance constraints.

I am eager to discuss how my technical versatility, analytical thinking capabilities, and commitment to delivery can support your performance targets. Thank you for your consideration.

Sincerely,

{cand_name}"""

    # 3. Modern Template Option Blueprint
    elif template_style == "Modern":
        return f"""[Date]

Hiring Team
[Company Name]

Subject: Re: Innovative {role_title} Application - {cand_name}

Dear Hiring Team,

The opportunity to scale platforms as a {role_title} directly matches my passion for engineering efficient tech layers. I excel at converting messy system logic parameters into high-velocity production systems.

My practical profile highlights active experience building and optimizing with frameworks like {skills_phrase}. I approach product challenges by treating infrastructure automation and clean coding logic as foundational requirements, not optional additions. This structured approach cuts down processing bugs and guarantees operational resilience.

I am looking to bring my energy, fast learning agility, and execution focus straight onto your product roadmap deliverables. Let's connect to review my project portfolio indicators in detail.

Best Regards,

{cand_name}"""

    # 4. Creative Template Option Blueprint
    else:
        return f"""[Date]

Hiring Team / Engineering Division
[Company Name]

Subject: Application for {role_title} - {cand_name}

Dear Creative Team,

Every system architecture tells a story—from the efficiency of database calls to the responsiveness of UI components. I am looking to apply my skills to the open {role_title} role to build creative code solutions that directly address your scalability objectives.

My development journey is defined by a deep curiosity for modern computing workflows. Using tools such as {skills_phrase}, I design solutions around the end-user experience, ensuring processing logic is built for both scale and speed. I bring unique perspective, adaptive learning habits, and rigorous testing habits to the engineering room.

I am excited about your company's commitment to building impactful platforms and would love to join forces to execute your upcoming technical releases.

Warm Regards,

{cand_name}"""


# ------------------------------------------
def generate_tailored_cover_letter(resume_text, jd_content, template_style, cache_bust=None):
    """Queries the Groq API for full semantic contextual cover letter tailoring optimized by selected design tone."""
    global client, GROQ_MODEL, GROQ_API_KEY
    
    cand_name, role_title, skills_phrase = extract_basic_entities(resume_text, jd_content)
    
    if isinstance(client, MockGroqClient) or not GROQ_API_KEY:
        return compile_static_template(resume_text, jd_content, template_style)

    # --- ADVANCED ENTITY EXTRACTION FOR PLACEHOLDERS ---
    # Try to extract a company name from the first few lines of the JD text
    company_name = "the company"
    company_match = re.search(r'(?:Company|Employer|Organization)[:\s\n]+([\w\s.\-]+)', jd_content, re.IGNORECASE)
    if company_match:
        company_name = company_match.group(1).strip()
    elif "--- Simulated JD for:" in jd_content:
        # Fallback handle for simulated mock layouts
        co_line = [line for line in jd_content.split('\n') if "Company:" in line]
        if co_line:
            company_name = co_line[0].replace("Company:", "").strip()

    # Core rules for each selected style type fed directly into the model path
    style_guidelines = {
        "Simple": "Direct, clean, minimalistic structure. Focus straightforwardly on basic capabilities, project building, and explicit interest.",
        "Professional": "Formal corporate tone. Emphasize operational alignment, criteria matching matrices, debugging, and robust processing logic constraints.",
        "Modern": "High-velocity, energetic tech-forward tone. Treat testing and automated infrastructure mechanics as a passion, using modern delivery frameworks.",
        "Creative": "Narrative, architectural storytelling tone. Connect computing workflows, unique perspectives, adaptive learning, and scalability to product roadmaps."
    }

    selected_guideline = style_guidelines.get(template_style, style_guidelines["Professional"])

    # Update current time variables dynamically for header insertion
    from datetime import datetime
    current_date = datetime.now().strftime("%B %d, %Y")

    prompt = f"""
    You are an elite career consultant and executive resume writer. Write a tailored, highly specific cover letter for the position of: {role_title} at {company_name}.
    
    CRITICAL STRUCTURE AND TONE RULE:
    You must draft this document specifically following the "{template_style}" tone design format. 
    Style Context: {selected_guideline}

    --- Target Job Description (JD) ---
    {jd_content}

    --- Candidate Resume Context ---
    {resume_text}

    --- EXPLICIT COMPLIANCE FOR ENTITIES ---
    - Use "{current_date}" for the date block at the top left.
    - If the target company name is explicitly identifiable in the text, replace [Company Name] with it. If not found, output "the company" or leave it as "your organization" smoothly. Do not return empty bracket expressions.
    - Do NOT include markdown styling or headers (No '#', '##', or '**').
    - Provide ONLY the raw text message body. Do not write introductory chatter, contextual meta-commentary, or wrap your code output inside a backend block markdown envelope structure like ```markdown.
    """

    try:
        response = client.chat.completions.create(
            model=GROQ_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.65
        )
        raw_output = response.choices[0].message.content.strip()
        
        # Clean standard structural formatting syntax leaks out of active canvas views completely
        cleaned_output = raw_output.replace('#', '').replace('**', '').replace('```markdown', '').replace('```', '')
        return cleaned_output.strip()
    except Exception as e:
        return f"AI Generation Error: Failed to compile cover letter. Detail: {str(e)}"
        
# GAP course plan -------------
def generate_gap_course_plan(gap_analysis_text, jd_role, candidate_skills):
    """Generates a detailed course plan and certification suggestions to fill identified gaps."""
    global client, GROQ_MODEL, GROQ_API_KEY
    
    if not gap_analysis_text.strip() or "No significant gaps" in gap_analysis_text:
        return "No specific gaps were identified in the match analysis. Focus on advanced skills in your core area."
        
    if isinstance(client, MockGroqClient) or not GROQ_API_KEY:
        # Mock client returns a hardcoded, structured plan (see MockGroqClient)
        response = client.chat().create(model=GROQ_MODEL, messages=[{"role": "user", "content": f"Generate a detailed course plan and suggest relevant certifications for Gaps Identified: {gap_analysis_text}"}])
        return response.choices[0].message.content.strip()

    prompt = f"""
    You are an expert career consultant. Based on the candidate's profile and the identified skill gaps for the role of **{jd_role}**, 
    generate a detailed course plan and suggest relevant certifications.
    
    **Context:**
    - Target Role: {jd_role}
    - Candidate's Current Key Skills: {', '.join(candidate_skills)}
    
    **Gaps Identified:**
    {gap_analysis_text}
    
    **Instructions:**
    1.  **Course Plan:** Structure the plan into 2-3 chronological phases (e.g., Foundational, Intermediate, Advanced/Project). Include specific topics (e.g., Python Basics, Docker Networking, Terraform Modules). Suggest a rough time estimate (e.g., weeks) for each phase.
    2.  **Certifications:** Suggest 2-3 industry-recognized certifications that directly address the identified gaps and enhance the resume for the target role.
    3.  **Output Format:** Use Markdown. Use the headings '## Detailed Course Plan' and '## Suggested Certifications'.
    """

    try:
        response = client.chat.completions.create(
            model=GROQ_MODEL, 
            messages=[{"role": "user", "content": prompt}], 
            temperature=0.6 
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        error_output = f"AI Generation Error: Failed to connect or receive response from LLM for course plan. Error: {e}\n{traceback.format_exc()}"
        return error_output


# --- ADAPTED LLM Functions for Interview Preparation (Modified) ---

def generate_interview_questions(parsed_json, section):
    """Generates categorized interview questions using LLM."""
    if not GROQ_API_KEY:
        return "AI Functions Disabled: GROQ_API_KEY not set."
    if "error" in parsed_json: return "Cannot generate questions due to resume parsing errors."
    
    section_title = section.replace("_", " ").title()
    section_content = parsed_json.get(section, "")
    if isinstance(section_content, (list, dict)):
        section_content = json.dumps(section_content, indent=2)
    elif not isinstance(section_content, str):
        section_content = str(section_content)

    if not section_content.strip():
        return f"No significant content found for the '{section_title}' section in the parsed resume. Please select a section with relevant data to generate questions."

    prompt = f"""Based on the following {section_title} section from the resume: {section_content}
Generate 3 interview questions each for these levels: Generic, Basic, Intermediate, Difficult.
**IMPORTANT: Format the output strictly as follows, with level headers and questions starting with 'Qx:':**
[Generic]
Q1: Question text...
Q2: Question text...
Q3: Question text...
[Basic]
Q1: Question text...
...
[Difficult]
Q3: Question text...
    """
    response = client.chat.completions.create(
        model=GROQ_MODEL, 
        messages=[{"role": "user", "content": prompt}], 
        temperature=0.5
    )
    return response.choices[0].message.content.strip()


# --- NEW FUNCTION: JD CHATBOT Q&A ---
def qa_on_jd(question, selected_jd_name):
    """Chatbot for JD (Q&A) using LLM."""
    if not GROQ_API_KEY:
        return "AI Chatbot Disabled: GROQ_API_KEY not set."

    # Find the JD content from the stored list
    jd_item = next((jd for jd in st.session_state.candidate_jd_list if jd['name'] == selected_jd_name), None)

    if not jd_item:
        return "Error: Could not find the selected Job Description in the loaded list."

    jd_text = jd_item['content']
    jd_metadata = {k: v for k, v in jd_item.items() if k not in ['name', 'content']}

    prompt = f"""Given the following Job Description and its extracted metadata:
    
    Job Description Title: {selected_jd_name}
    JD Metadata (JSON): {json.dumps(jd_metadata, indent=2)}
    JD Full Text:
    ---
    {jd_text}
    ---
    
    Answer the following question about the Job Description concisely and directly.
    If the information is not present in the provided text, state that clearly.
    
    Question: {question}
    """
    
    response = client.chat.completions.create(model=GROQ_MODEL, messages=[{"role": "user", "content": prompt}], temperature=0.4)
    return response.choices[0].message.content.strip()

# # interview evaluation--------------
def evaluate_interview_answers(qa_list, resume_context):
    """
    Evaluates a list of candidate's recorded answers based on the questions and resume context.
    The output is a full markdown report.
    """
    global client, GROQ_MODEL
    
    # Format Q&A for LLM
    qa_exchange = "\n\n--- Candidate Answers ---\n\n"
    for i, item in enumerate(qa_list):
        # Ensure question and answer are strings
        # Remove the (Level/Type) part from the question before sending it to the evaluator if necessary
        question = str(item['question'])
        answer = str(item['answer'])
        qa_exchange += f"Q{i+1}: {question}\n"
        qa_exchange += f"Answer {i+1}: {answer}\n"
        qa_exchange += "---"

    prompt = f"""
    You are an expert interviewer evaluating a candidate's recorded answers.
    
    **Evaluation Task:**
    Evaluate the candidate's answers based on the provided questions and their resume/JD context.
    
    **Instructions for Report:**
    1.  Provide an **Overall Score (X/10)** at the beginning of the report.
    2.  Give a **Summary** of the candidate's performance (e.g., strength in technical depth, weakness in behavioral structure). Include feedback on performance across the four types: HR-related, Experience-based, Situation-based, and Technical.
    3.  For **each question** answered, provide specific, actionable, constructive feedback. Use markdown headings (e.g., **Q1 Feedback**).
    4.  Ensure the report is professional and directly addresses consistency with the context.
    
    --- Context Used for Interview ---
    {resume_context}
    
    --- Interview Exchange ---
    {qa_exchange}
    
    ---
    **Output the evaluation report clearly using markdown.**
    """

    try:
        if isinstance(client, MockGroqClient) or not GROQ_API_KEY:
            response = client.chat().create(model=GROQ_MODEL, messages=[{"role": "user", "content": prompt}])
        else:
            response = client.chat.completions.create(
                model=GROQ_MODEL,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.5
            )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Evaluation Error: Failed to connect to LLM for scoring. Error: {e}"

# --- END ADAPTED LLM Functions ---

# --- Tab Content Functions ---
def resume_parsing_tab():
    # --- TAB 1: Resume Parsing ---
    st.header("📄 Resume Upload and Parsing")
    
    input_method = st.radio(
        "Select Input Method",
        ["Upload File", "Paste Text"],
        key="parsing_input_method"
    )
    
    st.markdown("---")

    if input_method == "Upload File":
        st.markdown("### 1. Upload Resume File") 
        
        uploaded_file = st.file_uploader( 
            "Choose PDF, DOCX, TXT, JSON, MD, CSV, XLSX file", 
            type=["pdf", "docx", "txt", "json", "md", "csv", "xlsx", "markdown", "rtf"], 
            accept_multiple_files=False, 
            key='candidate_file_upload_main'
        )
        
        st.markdown(
            """
            <div style='font-size: 10px; color: grey;'>
            Supported File Types: PDF, DOCX, TXT, JSON, MARKDOWN, CSV, XLSX, RTF
            </div>
            """, 
            unsafe_allow_html=True
        )
        st.markdown("---")

        if "candidate_uploaded_resumes" not in st.session_state: st.session_state.candidate_uploaded_resumes = []
        if "pasted_cv_text" not in st.session_state: st.session_state.pasted_cv_text = ""
        
        if uploaded_file is not None:
            if not st.session_state.candidate_uploaded_resumes or st.session_state.candidate_uploaded_resumes[0].name != uploaded_file.name:
                st.session_state.candidate_uploaded_resumes = [uploaded_file] 
                st.session_state.pasted_cv_text = "" 
                st.toast("Resume file uploaded successfully.")
        elif st.session_state.candidate_uploaded_resumes and uploaded_file is None:
            st.session_state.candidate_uploaded_resumes = []
            st.session_state.parsed = {}
            st.session_state.full_text = ""
            st.session_state.excel_data = None
            st.toast("Upload cleared.")
        
        file_to_parse = st.session_state.candidate_uploaded_resumes[0] if st.session_state.candidate_uploaded_resumes else None
        
        st.markdown("### 2. Parse Uploaded File")
        
        if file_to_parse:
            if st.button(f"Parse and Load: **{file_to_parse.name}**", use_container_width=True):
                with st.spinner(f"Parsing {file_to_parse.name}..."):
                    result = parse_and_store_resume(file_to_parse, file_name_key='single_resume_candidate', source_type='file')
                    
                    if result.get('error') is None:
                        st.session_state.parsed = result['parsed']
                        st.session_state.full_text = result['full_text']
                        st.session_state.excel_data = result['excel_data'] 
                        st.session_state.parsed['name'] = result['name'] 
                        clear_interview_state('resume')
                        clear_interview_state('jd')
                        if 'gap_analysis_plan' in st.session_state: del st.session_state['gap_analysis_plan']
                        st.success(f"✅ Successfully loaded and parsed **{result['name']}**.")
                        st.info("The parsed data is ready for matching.")
                        st.rerun() 
                    else:
                        st.error(f"Parsing failed for {file_to_parse.name}: {result['error']}")
                        st.session_state.parsed = {"error": result['error'], "name": result['name']}
                        st.session_state.full_text = result['full_text'] or ""
                        st.session_state.excel_data = result['excel_data'] 
        else:
            st.info("No resume file is currently uploaded. Please upload a file above.")

    else: # input_method == "Paste Text"
        st.markdown("### 1. Paste Your CV Text")
        
        pasted_text = st.text_area(
            "Copy and paste your entire CV or resume text here.",
            value=st.session_state.get('pasted_cv_text', ''),
            height=300,
            key='pasted_cv_text_input'
        )
        st.session_state.pasted_cv_text = pasted_text 
        
        st.markdown("---")
        st.markdown("### 2. Parse Pasted Text")
        
        if pasted_text.strip():
            if st.button("Parse and Load Pasted Text", use_container_width=True):
                with st.spinner("Parsing pasted text..."):
                    st.session_state.candidate_uploaded_resumes = []
                    
                    result = parse_and_store_resume(pasted_text, file_name_key='single_resume_candidate', source_type='text')
                    
                    if result.get('error') is None:
                        st.session_state.parsed = result['parsed']
                        st.session_state.full_text = result['full_text']
                        st.session_state.excel_data = result['excel_data'] 
                        st.session_state.parsed['name'] = result['name'] 
                        clear_interview_state('resume')
                        clear_interview_state('jd')
                        if 'gap_analysis_plan' in st.session_state: del st.session_state['gap_analysis_plan']
                        st.success(f"✅ Successfully loaded and parsed **{result['name']}**.")
                        st.info("The parsed data is ready for matching.") 
                        st.rerun()
                    else:
                        st.error(f"Parsing failed: {result['error']}")
                        st.session_state.parsed = {"error": result['error'], "name": result['name']}
                        st.session_state.full_text = result['full_text'] or ""
                        st.session_state.excel_data = result['excel_data'] 
        else:
            st.info("Please paste your CV text into the box above.")
            
    st.markdown("---")

import io
import json
import streamlit as st
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# ==============================================================================
# 1. GLOBAL HELPER CODES (Keep unindented at the top level of your file)
# ==============================================================================

def convert_to_pdf_bytes(cv_data):
    """
    Generates a professional, print-ready binary PDF from cv_data using ReportLab
    and returns the raw file data as bytes.
    """
    buffer = io.BytesIO()
    
    # Page Setup & Document Geometry
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40
    )
    
    story = []
    base_styles = getSampleStyleSheet()
    
    PRIMARY_COLOR = colors.HexColor("#1E90FF")  # Dodger Blue
    TEXT_COLOR = colors.HexColor("#333333")     # Charcoal
    MUTED_COLOR = colors.HexColor("#666666")    # Slate Grey
    
    name_style = ParagraphStyle(
        'CV_Name',
        parent=base_styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=26,
        leading=30,
        textColor=PRIMARY_COLOR,
        alignment=TA_CENTER,
        spaceAfter=4
    )
    
    contact_style = ParagraphStyle(
        'CV_Contact',
        parent=base_styles['Normal'],
        fontName='Helvetica',
        fontSize=10,
        leading=14,
        textColor=MUTED_COLOR,
        alignment=TA_CENTER,
        spaceAfter=15
    )
    
    h1_style = ParagraphStyle(
        'CV_H1',
        parent=base_styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=14,
        leading=18,
        textColor=PRIMARY_COLOR,
        spaceBefore=14,
        spaceAfter=6,
        keepWithNext=True
    )
    
    body_style = ParagraphStyle(
        'CV_Body',
        parent=base_styles['Normal'],
        fontName='Helvetica',
        fontSize=10.5,
        leading=15,
        textColor=TEXT_COLOR,
        alignment=TA_LEFT
    )
    
    # Compile Header / Branding Band
    personal = cv_data.get('personal_info', {})
    story.append(Paragraph(personal.get('name', 'Candidate Name').upper(), name_style))
    
    contact_bits = [
        f"Email: {personal.get('email', 'N/A')}",
        f"Phone: {personal.get('phone', 'N/A')}"
    ]
    if personal.get('address'):
        contact_bits.append(f"Address: {personal.get('address')}")
        
    contact_string = "  |  ".join(contact_bits)
    story.append(Paragraph(contact_string, contact_style))
    
    def get_divider(color=colors.HexColor("#ddd"), thickness=1, space=10):
        from reportlab.platypus import HRFlowable
        return HRFlowable(width="100%", thickness=thickness, color=color, spaceBefore=space, spaceAfter=space)
    
    story.append(get_divider(PRIMARY_COLOR, thickness=1.5, space=2))
    
    # Modular Section Renderer
    def render_cv_section(title, data_list):
        if not data_list:
            return
        
        section_elements = []
        section_elements.append(Paragraph(title.upper(), h1_style))
        section_elements.append(get_divider(space=2))
        
        for item in data_list:
            clean_item = str(item).strip()
            if " | " in clean_item:
                clean_item = clean_item.replace(" | ", "<br/>&bull; ")
                
            bullet_html = f"&bull; {clean_item}"
            section_elements.append(Paragraph(bullet_html, body_style))
            section_elements.append(Spacer(1, 4))
            
        section_elements.append(Spacer(1, 6))
        story.append(KeepTogether(section_elements))

    # Process Data Elements Sequential Layout Steps
    render_cv_section("Education", cv_data.get('education'))
    render_cv_section("Professional Experience", cv_data.get('experience'))
    render_cv_section("Key Projects", cv_data.get('projects'))
    render_cv_section("Certifications", cv_data.get('certifications'))
    
    # Process Custom Strengths
    strengths_raw = cv_data.get('strengths_raw', '')
    if strengths_raw.strip():
        strength_lines = []
        for line in strengths_raw.split('\n'):
            if line.strip():
                strength_lines.append(line.strip().lstrip('*+- ').strip())
        render_cv_section("Core Competencies & Expertise", strength_lines)
        
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


# ==============================================================================
# 2. MAIN WORKSPACE DASHBOARD INTERFACES
# ==============================================================================

def cv_management_tab():
    """Tab to allow form-based CV data entry and multi-format preview/download."""
    st.header("📝 CV Management & Form Generation")
    st.markdown("Generate a resume text structure by filling out the sections below. This text can then be parsed in the 'Resume Parsing' tab.")
    
    # --- 1. Personal Info ---
    st.subheader("1. Personal Information")
    col_name, col_email, col_phone = st.columns(3)
    
    with col_name:
        st.session_state.cv_data['personal_info']['name'] = st.text_input(
            "Full Name", 
            value=st.session_state.cv_data['personal_info'].get('name', ''), 
            key='cv_name'
        )
    with col_email:
        st.session_state.cv_data['personal_info']['email'] = st.text_input(
            "Email", 
            value=st.session_state.cv_data['personal_info'].get('email', ''), 
            key='cv_email'
        )
    with col_phone:
        st.session_state.cv_data['personal_info']['phone'] = st.text_input(
            "Phone Number", 
            value=st.session_state.cv_data['personal_info'].get('phone', ''), 
            key='cv_phone'
        )
        
    st.session_state.cv_data['personal_info']['address'] = st.text_input(
        "Communication Address (Optional)",
        value=st.session_state.cv_data['personal_info'].get('address', ''),
        key='cv_address'
    )
    
    st.markdown("---")

    # --- 2. Education ---
    st.subheader("2. Education")
    with st.form("education_form", clear_on_submit=True):
        col_deg, col_uni, col_fy, col_ty, col_score = st.columns([2, 2, 1, 1, 1])
        with col_deg: degree = st.text_input("Degree/Qualification", key='edu_degree')
        with col_uni: university = st.text_input("University/Institution", key='edu_uni')
        with col_fy: year_from = st.text_input("From Year", key='edu_fy')
        with col_ty: year_to = st.text_input("To Year", key='edu_ty')
        with col_score: score = st.text_input("Scores (GPA/%)", key='edu_score', help="e.g., 3.8/4.0 or 85%")
        
        if st.form_submit_button("Add Education"):
            if degree and university:
                score_display = f", Score: {score}" if score else ""
                entry = f"Degree: {degree}, Institution: {university} ({year_from}-{year_to}){score_display}"
                st.session_state.cv_data['education'].append(entry)
                st.success(f"Added: {entry}")
            else: st.error("Please enter Degree and University.")
            
    if st.session_state.cv_data['education']:
        st.dataframe(st.session_state.cv_data['education'], use_container_width=True, hide_index=True)
    st.markdown("---")

    # --- 3. Experience ---
    st.subheader("3. Professional Experience")
    with st.form("experience_form", clear_on_submit=True):
        col_comp, col_role, col_ctc = st.columns([2, 2, 1])
        with col_comp: company = st.text_input("Company Name", key='exp_company')
        with col_role: role = st.text_input("Role/Title", key='exp_role')
        with col_ctc: ctc = st.text_input("CTC (Annual)", key='exp_ctc')
        col_fy, col_ty = st.columns(2)
        with col_fy: year_from = st.text_input("From Year/Date", key='exp_fy')
        with col_ty: year_to = st.text_input("To Year/Date (or Present)", key='exp_ty')
        
        responsibilities = st.text_area("Key Responsibilities (Use bullet points)", key='exp_resp', height=100)
        achievements = st.text_area("Key Achievements/Metrics", key='exp_achiev', height=100)
        
        if st.form_submit_button("Add Experience"):
            if company and role:
                resp_formatted = responsibilities.replace('\n', ' | ').strip()
                achiev_formatted = achievements.replace('\n', ' | ').strip()
                
                desc_parts = []
                if resp_formatted:
                    desc_parts.append(f"Responsibilities: {resp_formatted}")
                if achiev_formatted:
                    desc_parts.append(f"Achievements: {achiev_formatted}")
                
                description_text = ". ".join(desc_parts)
                entry = f"Role: {role} at {company} (CTC: {ctc}) ({year_from}-{year_to}). {description_text}"
                st.session_state.cv_data['experience'].append(entry)
                st.success(f"Added: {role} at {company}")
            else: st.error("Please enter Company Name and Role.")
            
    if st.session_state.cv_data['experience']:
        st.dataframe(st.session_state.cv_data['experience'], use_container_width=True, hide_index=True)
    st.markdown("---")

    # --- 4. Projects ---
    st.subheader("4. Projects")
    with st.form("projects_form", clear_on_submit=True):
        col_name, col_link = st.columns(2)
        with col_name: project_name = st.text_input("Project Name", key='proj_name')
        with col_link: app_link = st.text_input("App/Repo Link", key='proj_link')
        tools = st.text_input("Tools Used (Comma Separated)", key='proj_tools')
        description = st.text_area("Description and Accomplishments", key='proj_desc', height=100)
        
        if st.form_submit_button("Add Project"):
            if project_name:
                desc_formatted = description.replace('\n', ' | ').strip()
                entry = f"Project: {project_name}. Tools: {tools}. Link: {app_link}. Description: {desc_formatted}"
                st.session_state.cv_data['projects'].append(entry)
                st.success(f"Added Project: {project_name}")
            else: st.error("Please enter Project Name.")
            
    if st.session_state.cv_data['projects']:
        st.dataframe(st.session_state.cv_data['projects'], use_container_width=True, hide_index=True)
    st.markdown("---")

    # --- 5. Certifications ---
    st.subheader("5. Certifications")
    with st.form("cert_form", clear_on_submit=True):
        col_title, col_by = st.columns(2)
        with col_title: title = st.text_input("Certificate Title", key='cert_title')
        with col_by: given_by = st.text_input("Given By (Issuing Body)", key='cert_given_by')
        col_rec, col_date = st.columns(2)
        with col_rec: received_by = st.text_input("Received By (Your Name)", key='cert_received_by')
        with col_date: date = st.text_input("Date Received (YYYY-MM-DD)", key='cert_date')
        
        if st.form_submit_button("Add Certification"):
            if title:
                entry = f"Certification: {title} from {given_by}. Received by {received_by} on {date}."
                st.session_state.cv_data['certifications'].append(entry)
                st.success(f"Added Certification: {title}")
            else: st.error("Please enter Certificate Title.")
            
    if st.session_state.cv_data['certifications']:
        st.dataframe(st.session_state.cv_data['certifications'], use_container_width=True, hide_index=True)
    st.markdown("---")

    # --- 6. Key Responsibilities, Expertise, or Leadership Skills ---
    st.subheader("6. Key Responsibilities, Expertise, or Leadership Skills")
    st.session_state.cv_data['strengths_raw'] = st.text_area(
        "Enter relevant expertise, core competencies, or soft/leadership skills (one per line)",
        value=st.session_state.cv_data.get('strengths_raw', ''),
        key='cv_strengths_input',
        height=150
    )
    st.markdown("---")

    # --- 7. Generate and Preview Text ---
    if st.button("Generate CV Data for Parsing & Preview", type="primary", use_container_width=True):
        st.session_state.form_cv_text = generate_cv_text()
        st.info("CV Data Generated. Go to **Resume Parsing** tab and select 'Use Form Data'.")
        st.rerun()  
        
    st.markdown("##### Current Generated Data Preview")
    
    # FIXED: Nesting the preview rendering inside the safe function boundaries
    if st.session_state.get('form_cv_text'):
        markdown_text = st.session_state.form_cv_text
        json_data = json.dumps(st.session_state.cv_data, indent=4) 
        html_content = convert_to_html_content(st.session_state.cv_data) if 'convert_to_html_content' in locals() else ""
        
        with st.spinner("Preparing true PDF downloadable channels..."):
            pdf_data_bytes = convert_to_pdf_bytes(st.session_state.cv_data)

        tab_md, tab_json, tab_pdf_download = st.tabs(["Markdown (.md)", "JSON (.json)", "📕 Export PDF Document"])

        with tab_md:
            st.code(markdown_text, language='markdown')
            st.download_button(
                label="⬇️ Download Markdown (.md)",
                data=markdown_text,
                file_name=f"{st.session_state.cv_data['personal_info']['name'].replace(' ', '_')}_cv.md",
                mime="text/markdown",
                use_container_width=True
            )

        with tab_json:
            st.json(json_data)
            st.download_button(
                label="⬇️ Download JSON (.json)",
                data=json_data,
                file_name=f"{st.session_state.cv_data['personal_info']['name'].replace(' ', '_')}_cv.json",
                mime="application/json",
                use_container_width=True
            )

        with tab_pdf_download:
            st.markdown("### Download Document Profile")
            st.info("Your application dataset has been successfully processed into a print-ready PDF layout matching modern ATS requirements.")
            
            st.download_button(
                label="📕 Download Verified PDF Document (.pdf)",
                data=pdf_data_bytes,
                file_name=f"{st.session_state.cv_data['personal_info']['name'].replace(' ', '_')}_Resume.pdf",
                mime="application/pdf",
                use_container_width=True
            )
            
            if html_content:
                st.markdown("---")
                st.markdown("##### Alternative HTML Layout Preview")
                st.components.v1.html(html_content, height=250, scrolling=True)
    else:
        st.info("No CV text generated yet. Fill out the forms and click the generate button.")

    if st.button("🗑️ Clear All Form Data", key="clear_cv_form_data"):
        st.session_state.cv_data = {
            'personal_info': {'name': '', 'email': '', 'phone': '', 'address': ''},
            'education': [],
            'experience': [],
            'projects': [],
            'certifications': [],
            'strengths_raw': '' 
        }
        st.session_state.form_cv_text = ""
        st.rerun()


def generate_cv_text():
    """Generates the text/markdown format from all stored session state data."""
    data = st.session_state.cv_data
    text = f"# Candidate Resume Data\n\n"
    
    text += f"**Name**: {data['personal_info'].get('name', '')}\n"
    text += f"**Email**: {data['personal_info'].get('email', '')}\n"
    text += f"**Phone**: {data['personal_info'].get('phone', '')}\n"
    if data['personal_info'].get('address'):
        text += f"**Address**: {data['personal_info']['address']}\n"
    text += "\n"
    
    text += "## Education\n"
    if data['education']: text += "* " + "\n* ".join(data['education']) + "\n\n"
    
    text += "## Experience\n"
    if data['experience']: text += "* " + "\n* ".join(data['experience']) + "\n\n"
        
    text += "## Projects\n"
    if data['projects']: text += "* " + "\n* ".join(data['projects']) + "\n\n"
        
    text += "## Certifications\n"
    if data['certifications']: text += "* " + "\n* ".join(data['certifications']) + "\n\n"
        
    strengths_raw_data = data.get('strengths_raw', '')
    if strengths_raw_data:
        strengths_list = [s.strip() for s in strengths_raw_data.split('\n') if s.strip()]
        if strengths_list:
            text += "## Key Responsibilities, Expertise, or Leadership Skills\n"
            text += "* " + "\n* ".join(strengths_list) + "\n\n"
        
    return text.strip()
    
# --- Helper URL Extraction Function (Module Level) ---
def extract_jd_from_linkedin_url(url):
    if "linkedin.com/jobs" not in url:
        return f"[Error] Invalid LinkedIn Job URL: {url}"

    url_lower = url.lower()
    if "data-scientist" in url_lower:
        role = "Data Scientist"
        skills = ["Python", "SQL", "ML", "Data Analysis", "Pytorch", "Visualization"]
        focus = "machine learning and statistical modeling"
    elif "cloud-engineer" in url_lower or "aws" in url_lower:
        role = "Cloud Engineer"
        skills = ["AWS", "Docker", "Kubernetes", "Cloud Services", "GCP", "Terraform"]
        focus = "infrastructure as code and cloud deployment"
    elif "ml-engineer" in url_lower or "ai-engineer" in url_lower:
        role = "AI/ML Engineer"
        skills = ["MLOps", "LLM", "Deep Learning", "Python", "TensorFlow", "API Services"]
        focus = "production-level AI/ML model development and deployment"
    else:
        role = "Software Engineer"
        skills = ["Java", "API", "SQL", "React", "JavaScript"]
        focus = "full-stack application development"
        
    skills_str = ", ".join(skills)

    return f"""
    --- Simulated JD for: {role} ---
    
    Company: MockCorp
    Location: Remote
    
    Job Summary:
    We are seeking a highly skilled **{role}** to join our team. The ideal candidate will have expertise in {skills_str}. Must be focused on **{focus}**. This is a Full-time position.
    
    Responsibilities:
    * Develop and maintain systems using **{skills[0]}** and **{skills[1]}** in a collaborative environment.
    * Manage and deploy applications on **{skills[2]}** platforms.
    * Collaborate with cross-functional teams.
    
    Qualifications:
    * 3+ years of experience.
    * Strong proficiency in **{skills[0]}** and analytical tools.
    * Experience with cloud platforms (e.g., AWS).
    ---
    """
# --- JD Management Tab Function ---
def jd_management_tab_candidate():
    """JD Management Tab."""
    st.header("📚 Manage Job Descriptions for Matching")
    st.markdown("Add multiple JDs here to compare your resume against them in the next tabs.")
    
    if "candidate_jd_list" not in st.session_state: 
        st.session_state.candidate_jd_list = []
    st.markdown("---")
    
    jd_type = st.radio("Select JD Type", ["Single JD", "Multiple JD"], key="jd_type_candidate", index=0)
    st.markdown("### Add JD by:")
    method = st.radio("Choose Method", ["Upload File", "Paste Text", "LinkedIn URL"], key="jd_add_method_candidate", index=0) 
    st.markdown("---")

    if method == "LinkedIn URL": 
        with st.form("jd_url_form_candidate", clear_on_submit=True):
            url_list = st.text_area("Enter one or more URLs (comma separated)" if jd_type == "Multiple JD" else "Enter URL", key="url_list_candidate")
            if st.form_submit_button("Add JD(s) from URL", key="add_jd_url_btn_candidate"):
                if url_list:
                    urls = [u.strip() for u in url_list.split(",")] if jd_type == "Multiple JD" else [url_list.strip()]
                    count = 0
                    for url in urls:
                        if not url: 
                            continue
                        with st.spinner(f"Attempting JD extraction and metadata analysis for: {url}"):
                            jd_text = extract_jd_from_linkedin_url(url)
                            metadata = extract_jd_metadata(jd_text)
                        
                        if metadata.get('role') == 'Extraction Error':
                            st.error(f"Failed to process {url}: {jd_text}")
                            continue
                            
                        name = f"JD for {metadata.get('role', 'Unknown Role')}"
                        st.session_state.candidate_jd_list.append({"name": name, "content": jd_text, **metadata})
                        count += 1
                            
                    if count > 0:
                        st.success(f"✅ {count} JD(s) added successfully!")
                        st.rerun() 
                    else:
                        st.error("No JDs were added successfully.")

    elif method == "Paste Text":
        with st.form("jd_paste_form_candidate", clear_on_submit=True):
            text_list = st.text_area("Paste one or more JD texts (separate by '---')" if jd_type == "Multiple JD" else "Paste JD text here", key="text_list_candidate")
            if st.form_submit_button("Add JD(s) from Text", key="add_jd_text_btn_candidate"):
                if text_list:
                    texts = [t.strip() for t in text_list.split("---")] if jd_type == "Multiple JD" else [text_list.strip()]
                    count = 0
                    for i, text in enumerate(texts):
                        if text:
                            metadata = extract_jd_metadata(text)
                            
                            if metadata.get('role') == 'Extraction Error':
                                st.error(f"Failed to extract metadata for pasted text {i+1}.")
                                continue
                                
                            name_base = metadata.get('role', f"Pasted JD {len(st.session_state.candidate_jd_list) + i + 1}")
                            st.session_state.candidate_jd_list.append({"name": name_base, "content": text, **metadata})
                            count += 1
                    
                    if count > 0:
                        st.success(f"✅ {count} JD(s) added successfully!")
                        st.rerun() 

    elif method == "Upload File":
        jd_file_types = ["pdf", "txt", "docx", "md", "json"]
        uploaded_files = st.file_uploader(
            f"Upload JD file(s) ({', '.join(jd_file_types)})",
            type=jd_file_types,
            accept_multiple_files=(jd_type == "Multiple JD"),
            key="jd_file_uploader_candidate"
        )
        files_to_process = uploaded_files if isinstance(uploaded_files, list) else ([uploaded_files] if uploaded_files else [])
        
        with st.form("jd_upload_form_candidate", clear_on_submit=False):
            if files_to_process:
                st.markdown("##### Files Selected:")
                for file in files_to_process:
                    st.markdown(f"&emsp;📄 **{file.name}** {round(file.size / (1024*1024), 2)}MB")
                    
            if st.form_submit_button("Add JD(s) from File", key="add_jd_file_btn_candidate"):
                if not files_to_process:
                    st.warning("Please upload file(s).")
                    
                count = 0
                for file in files_to_process:
                    if file:
                        with st.spinner(f"Extracting content from {file.name}..."):
                            file_type = get_file_type(file.name)
                            file.seek(0)
                            jd_text, _ = extract_content(file_type, file.getvalue(), file.name)
                            
                        if not jd_text.startswith("[Error"):
                            metadata = extract_jd_metadata(jd_text)
                            
                            if metadata.get('role') == 'Extraction Error': 
                                st.error(f"Failed to extract metadata for {file.name}.")
                                continue
                                
                            st.session_state.candidate_jd_list.append({"name": file.name, "content": jd_text, **metadata})
                            count += 1
                        else:
                            st.error(f"Error extracting content from {file.name}: {jd_text}")
                            
                if count > 0:
                    st.success(f"✅ {count} JD(s) added successfully!")
                    st.rerun()
                elif uploaded_files:
                    st.error("No valid JD files were uploaded or content extraction failed.")

    st.markdown("---")
    if st.session_state.candidate_jd_list:
        col_display_header, col_clear_button = st.columns([3, 1])
        with col_display_header: 
            st.markdown("### ✅ Current JDs Added:")
            
        with col_clear_button:
            if st.button("🗑️ Clear All JDs", key="clear_jds_candidate", use_container_width=True, help="Removes all currently loaded JDs."):
                st.session_state.candidate_jd_list = []
                if 'candidate_match_results' in st.session_state: del st.session_state['candidate_match_results']
                if 'jd_chatbot_history' in st.session_state: del st.session_state['jd_chatbot_history']
                if 'gap_analysis_plan' in st.session_state: del st.session_state['gap_analysis_plan']
                clear_interview_state('jd')
                st.success("All JDs and associated data have been cleared.")
                st.rerun() 

        for idx, jd_item in enumerate(st.session_state.candidate_jd_list, 1):
            title = jd_item.get('name', f'JD {idx}')
            role = jd_item.get('role', 'N/A')
            job_type = jd_item.get('job_type', 'N/A')
            key_skills = jd_item.get('key_skills', ['N/A'])
            content = jd_item.get('content', 'No content extracted.')
            
            display_title = title.replace("--- Simulated JD for: ", "")
            with st.expander(f"**JD {idx}:** {display_title} | Role: {role}"):
                st.markdown(f"**Job Type:** {job_type} | **Key Skills:** `{', '.join(key_skills)}`")
                st.markdown("---")
                st.text(content)
    else:
        st.info("No Job Descriptions added yet.")
        
# --- Batch Match Tab Function (UPDATED) ---

import re
import pandas as pd
import streamlit as st

def jd_batch_match_tab():
    """The Batch JD Match tab logic."""
    st.header("🎯 Batch JD Match: Best Matches")
    st.markdown("Compare your current resume against all saved job descriptions.")
    
    # Determine if a resume/CV is ready
    is_resume_parsed = (
        st.session_state.get('parsed') is not None and
        st.session_state.parsed.get('name') is not None and
        st.session_state.parsed.get('error') is None
    )
    
    # Check if we are running in Mock Mode
    is_mock_mode = isinstance(client, MockGroqClient) and not GROQ_API_KEY
    
    if not is_resume_parsed:
        st.warning("⚠️ Please **upload and parse your resume** in the 'Resume Parsing' tab first.")
        if st.session_state.get('parsed', {}).get('error') is not None:
             st.error(f"Resume Parsing Error: {st.session_state.parsed.get('error')}")

    elif not st.session_state.candidate_jd_list:
        st.error("❌ Please **add Job Descriptions** in the 'JD Management' tab before running batch analysis.")
        
    elif not GROQ_API_KEY and not is_mock_mode:
        st.error("Cannot use JD Match: GROQ_API_KEY is not configured.")
        
    else:
        if not is_mock_mode and (not hasattr(client, 'client_ready') or not client.client_ready):
            st.warning("⚠️ LLM client setup failed. Match analysis may not be available.")

    if "candidate_match_results" not in st.session_state:
        st.session_state.candidate_match_results = []

    all_jd_names = [item['name'] for item in st.session_state.candidate_jd_list]
    
    selected_jd_names = st.multiselect(
        "Select Job Descriptions to Match Against",
        options=all_jd_names,
        default=all_jd_names, 
        key='candidate_batch_jd_select'
    )
    
    jds_to_match = [
        jd_item for jd_item in st.session_state.candidate_jd_list 
        if jd_item['name'] in selected_jd_names
    ]
    
    if st.button(f"Run Match Analysis on **{len(jds_to_match)}** Selected JD(s)"):
        st.session_state.candidate_match_results = []
        if 'gap_analysis_plan' in st.session_state: del st.session_state['gap_analysis_plan']
        
        if not jds_to_match:
            st.warning("Please select at least one Job Description.")
        elif not is_resume_parsed:
             st.warning("Please upload and parse your resume successfully first.")
        else:
            resume_name = st.session_state.parsed.get('name', 'Uploaded Resume')
            parsed_json = st.session_state.parsed
            results_with_score = []

            with st.spinner(f"Matching {resume_name}'s resume against {len(jds_to_match)} JDs..."):
                for jd_item in jds_to_match:
                    jd_name = jd_item['name']
                    jd_content = jd_item['content']

                    # Call the LLM evaluation function
                    fit_output = evaluate_jd_fit(jd_content, parsed_json) 
                    
                    # --- FIXED EXTRACTION LOGIC ---
                    
                    # 1. Improved Score Extraction: Removed the 'text*?' bug
                    score_patterns = [
                        r"Overall Fit Score:\s*\*?\[?\s*(\d+)\s*\]?\s*/\s*10",
                        r"Overall\s*Score:\s*\*?\[?\s*(\d+)\s*\]?\s*/\s*10",
                        r"Fit\s*Score:\s*\*?\[?\s*(\d+)\s*\]?\s*/\s*10"
                    ]
                    
                    overall_score = "N/A"
                    for pattern in score_patterns:
                        match = re.search(pattern, fit_output, re.IGNORECASE)
                        if match:
                            overall_score = match.group(1)
                            break
                    
                    # Fallback for any "Number/10" in the text
                    if overall_score == "N/A":
                        fallback = re.search(r"(\d+)\s*/\s*10", fit_output)
                        if fallback:
                            overall_score = fallback.group(1)

                    # 2. Extract Section Match Analysis block
                    section_analysis_match = re.search(
                        r'--- Section Match Analysis ---\s*(.*?)\s*(?:Strengths|Overall Summary|Gaps|$)', 
                        fit_output, re.DOTALL | re.IGNORECASE
                    )
                    
                    skills_percent, exp_percent, edu_percent = '0', '0', '0'
                    if section_analysis_match:
                        section_text = section_analysis_match.group(1)
                        # Look for digits followed by optional % sign
                        s_m = re.search(r'Skills\s*Match:\s*(\d+)', section_text, re.IGNORECASE)
                        x_m = re.search(r'Experience\s*Match:\s*(\d+)', section_text, re.IGNORECASE)
                        e_m = re.search(r'Education\s*Match:\s*(\d+)', section_text, re.IGNORECASE)
                        
                        if s_m: skills_percent = s_m.group(1)
                        if x_m: exp_percent = x_m.group(1)
                        if e_m: edu_percent = e_m.group(1)

                    # 3. Extract Gaps
                    gaps_match = re.search(r'Gaps/Areas for Improvement:\s*(.*?)\s*(?:Overall Summary|---|$)', fit_output, re.DOTALL | re.IGNORECASE)
                    raw_gaps = gaps_match.group(1).strip() if gaps_match else "See detailed analysis below."
                    
                    if "AI Evaluation Error" in fit_output:
                        overall_score = "Error"
                    
                    results_with_score.append({
                        "jd_name": jd_name,
                        "overall_score": overall_score,
                        "numeric_score": int(overall_score) if str(overall_score).isdigit() else -1, 
                        "skills_percent": skills_percent,
                        "experience_percent": exp_percent, 
                        "education_percent": edu_percent, 
                        "full_analysis": fit_output,
                        "gaps": raw_gaps
                    })
                        
                # Sort by score descending
                results_with_score.sort(key=lambda x: x['numeric_score'], reverse=True)
                
                # Assign Ranks
                for i, item in enumerate(results_with_score):
                    item['rank'] = i + 1
                    
                st.session_state.candidate_match_results = results_with_score
                st.success("Batch analysis complete!")
                st.rerun() 

    # --- Display Results ---
    if st.session_state.get('candidate_match_results'):
         st.markdown("---")
         st.subheader("Match Analysis Summary")
         
         summary_df_data = []
         for res in st.session_state.candidate_match_results:
             summary_df_data.append({
                 "Rank": res.get('rank'),
                 "Job Description": res['jd_name'].replace("JD for ", ""),
                 "Overall Score (10)": res['overall_score'],
                 "Experience %": f"{res['experience_percent']}%",
                 "Education %": f"{res['education_percent']}%",
                 "Skills %": f"{res['skills_percent']}%"
             })
             
         summary_df = pd.DataFrame(summary_df_data)
         
         # Color formatting callback logic
         def color_score(val):
             try:
                 num = int(val)
                 if num >= 8: return 'background-color: #d4edda; color: #155724' # Clean Green
                 elif num >= 6: return '' # Default styling logic
                 return 'background-color: #f8d7da; color: #721c24' # Clean Red
             except: return ''
             
         st.dataframe(
             summary_df.style.map(color_score, subset=['Overall Score (10)']), 
             use_container_width=True,
             column_order=["Rank", "Job Description", "Overall Score (10)", "Experience %", "Education %", "Skills %"],
             hide_index=True
         )

         st.markdown("---")
         st.subheader("Detailed Analysis")
         
         for res in st.session_state.candidate_match_results:
             with st.expander(f"**Rank {res.get('rank')}** | {res['jd_name']} | **Score: {res['overall_score']}/10**"):
                 st.write(f"**Section Matches:** Exp: {res['experience_percent']}% | Edu: {res['education_percent']}% | Skills: {res['skills_percent']}%")
                 st.markdown("---")
                 st.markdown(res['full_analysis'])
    else:
         st.info("Run the match analysis above to evaluate your resume against selected Job Descriptions.")
# --- Filter JD Tab Function (unchanged) ---
def filter_jd_tab_content():
    """Filter JD Tab."""
    st.header("🔍 Filter Job Descriptions by Criteria")
    st.markdown("Use the filters below to narrow down your saved Job Descriptions.")

    if not st.session_state.candidate_jd_list:
        st.info("No Job Descriptions are currently loaded. Please add JDs in the 'JD Management' tab.")
        if 'filtered_jds_display' not in st.session_state:
            st.session_state.filtered_jds_display = []
        return
    
    global DEFAULT_ROLES, DEFAULT_JOB_TYPES, STARTER_KEYWORDS
    
    # Safely extract roles, types, and skills from loaded JDs
    unique_roles = sorted(list(set(
        [item.get('role', 'General Analyst') for item in st.session_state.candidate_jd_list] + DEFAULT_ROLES
    )))
    # Note: Using DEFAULT_JOB_TYPES as a base, ensuring all loaded types are included.
    unique_job_types = sorted(list(set(
        [item.get('job_type', 'Full-time') for item in st.session_state.candidate_jd_list] + DEFAULT_JOB_TYPES
    )))
    
    all_unique_skills = set(STARTER_KEYWORDS)
    for jd in st.session_state.candidate_jd_list:
        valid_skills = [
            skill.strip() for skill in jd.get('key_skills', []) 
            if isinstance(skill, str) and skill.strip()
        ]
        all_unique_skills.update(valid_skills)
    
    unique_skills_list = sorted(list(all_unique_skills))
    
    if not unique_skills_list:
        unique_skills_list = ["No skills extracted from current JDs"]

    all_jd_data = st.session_state.candidate_jd_list

    with st.form(key="jd_filter_form"):
        st.markdown("### Select Filters")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            selected_skills = st.multiselect(
                "Skills Keywords (Select multiple)",
                options=unique_skills_list,
                default=st.session_state.get('last_selected_skills', []),
                key="candidate_filter_skills_multiselect", 
                help="Select one or more skills. JDs containing ANY of the selected skills will be shown."
            )
            
        with col2:
            selected_job_type = st.selectbox(
                "Job Type",
                options=["All Job Types"] + unique_job_types,
                index=0, 
                key="filter_job_type_select"
            )
            
        with col3:
            selected_role = st.selectbox(
                "Role Title",
                options=["All Roles"] + unique_roles,
                index=0, 
                key="filter_role_select"
            )

        apply_filters_button = st.form_submit_button("✅ Apply Filters", type="primary", use_container_width=True)

    if apply_filters_button:
        st.session_state.last_selected_skills = selected_skills

        filtered_jds = []
        selected_skills_lower = [k.strip().lower() for k in selected_skills]
        
        for jd in all_jd_data:
            jd_role = jd.get('role', 'General Analyst')
            jd_job_type = jd.get('job_type', 'Full-time')
            jd_key_skills = [
                s.lower() for s in jd.get('key_skills', []) 
                if isinstance(s, str) and s.strip()
            ]
            
            role_match = (selected_role == "All Roles") or (selected_role == jd_role)
            job_type_match = (selected_job_type == "All Job Types") or (selected_job_type == jd_job_type)
            
            skill_match = True
            if selected_skills_lower:
                if not any(k in jd_key_skills for k in selected_skills_lower):
                    skill_match = False
            
            if role_match and job_type_match and skill_match:
                filtered_jds.append(jd)
                
        st.session_state.filtered_jds_display = filtered_jds
        st.success(f"Filter applied! Found {len(filtered_jds)} matching Job Descriptions.")

    st.markdown("---")
    
    if 'filtered_jds_display' not in st.session_state:
        st.session_state.filtered_jds_display = []
        
    filtered_jds = st.session_state.filtered_jds_display
    
    st.subheader(f"Matching Job Descriptions ({len(filtered_jds)} found)")
    
    if filtered_jds:
        display_data = []
        for jd in filtered_jds:
            display_data.append({
                "Job Description Title": jd.get('name', 'N/A').replace("--- Simulated JD for: ", ""),
                "Role": jd.get('role', 'N/A'),
                "Job Type": jd.get('job_type', 'N/A'),
                "Key Skills": ", ".join(jd.get('key_skills', ['N/A'])[:5]) + "...",
            })
            
        st.dataframe(display_data, use_container_width=True)

        st.markdown("##### Detailed View")
        for idx, jd in enumerate(filtered_jds, 1):
            with st.expander(f"JD {idx}: {jd.get('name', 'N/A').replace('--- Simulated JD for: ', '')} - ({jd.get('role', 'N/A')})"):
                st.markdown(f"**Job Type:** {jd.get('job_type', 'N/A')}")
                st.markdown(f"**Extracted Skills:** {', '.join(jd.get('key_skills', ['N/A']))}")
                st.markdown("---")
                st.text(jd.get('content', 'Content not available'))
    elif st.session_state.candidate_jd_list and apply_filters_button:
        st.info("No Job Descriptions match the selected criteria. Try broadening your filter selections.")
    elif st.session_state.candidate_jd_list and not apply_filters_button:
        st.info("Use the filters above and click **'Apply Filters'** to view matching Job Descriptions.")


# --- Parsed Data Tab (unchanged) ---
def parsed_data_tab():
    """Parsed Data View Tab."""
    st.header("✨ Parsed Resume Data View")
    st.markdown("This tab displays the loaded candidate data and provides download options.")
    st.markdown("---")

    is_data_loaded_and_valid = (
        st.session_state.get('parsed', {}).get('name') is not None and 
        st.session_state.get('parsed', {}).get('error') is None
    )

    if is_data_loaded_and_valid:
        candidate_name = st.session_state.parsed['name']
        
        source_key = st.session_state.get('current_parsing_source_name', 'Unknown Source')
        if source_key == "Pasted_Text":
            source_display = "Pasted CV Data"
        elif source_key == "Form_Compiled_CV":
             source_display = "Manually Compiled CV"
        else:
            source_display = source_key.replace('_', ' ').replace('-', ' ') 

        base_filename = f"{candidate_name.replace(' ', '_')}_Parsed_Resume"
        parsed_json_data = json.dumps(st.session_state.parsed, indent=4)
        parsed_markdown_data = st.session_state.full_text
        
        json_filename = f"{base_filename}.json"
        md_filename = f"{base_filename}.md"
        html_filename = f"{base_filename}.html"
        
        json_data_uri = get_download_link(parsed_json_data, json_filename, 'json', title="Parsed Resume Data")
        md_data_uri = get_download_link(parsed_markdown_data, md_filename, 'markdown', title="Parsed Resume Data")
        html_data_uri = get_download_link(parsed_markdown_data.replace('\n', '<br>').replace('##', '<h2>'), html_filename, 'html', title="Parsed Resume Data") 
        
        tab_markdown, tab_json, tab_download = st.tabs([
            "📄 Markdown View", 
            "💾 JSON View", 
            "⬇️ PDF/HTML Download"
        ])

        with tab_markdown:
            st.markdown(f"**Candidate:** **{candidate_name}**")
            st.caption(f"Source: {source_display}")
            st.markdown("---")
            st.markdown("### Resume Content in Markdown Format")
            st.markdown(st.session_state.full_text)
            
            if st.session_state.excel_data:
                 st.markdown("### Extracted Spreadsheet Data (if applicable)")
                 st.json(st.session_state.excel_data)
                 
            st.markdown("---")
            st.markdown("##### Download Markdown Data")
            render_download_button(
                md_data_uri, 
                md_filename, 
                f"⬇️ Download Markdown (.md)", 
                'markdown'
            )

        with tab_json:
            st.markdown(f"**Candidate:** **{candidate_name}**")
            st.caption(f"Source: {source_display}")
            st.markdown("---")
            st.markdown("### Structured Data in JSON Format")
            st.json(st.session_state.parsed)

            st.markdown("---")
            st.markdown("##### Download JSON Data")
            render_download_button(
                json_data_uri, 
                json_filename, 
                f"💾 Download JSON (.json)", 
                'json'
            )

        with tab_download:
            st.markdown("### Download Viewable Document")
            st.info("This download provides the data in an HTML file that can be easily viewed or printed/saved as a PDF.")
            
            col_html = st.columns(1)[0]

            with col_html:
                st.markdown(f"**{html_filename.replace('.html', '.pdf/html')}**", help="Viewable document format.")
                render_download_button(
                    html_data_uri, 
                    html_filename, 
                    f"📄 Download HTML (PDF Sim.)", 
                    'html'
                )
                
            st.markdown("---")
            st.info("For structured data (JSON) or raw text (Markdown), please check their respective viewing tabs.")

    else:
        st.warning(f"**Status:** ❌ **No Valid Resume Data Loaded**")
        if st.session_state.get('parsed', {}).get('error') is not None:
             st.error(f"Last Parsing Error: {st.session_state.parsed['error']}")
        st.info("Please successfully parse a resume in the **Resume Parsing** tab or compile one in **CV Management**.")


# --- Interview Preparation Tab (UPDATED) ---
def parse_questions_from_raw(raw_questions_response):
    """Parses the structured raw LLM output into a list of Q&A dictionaries."""
    q_list = []
    current_level_type = "General"
    
    for line in raw_questions_response.splitlines():
        line = line.strip()
        if line.startswith('[') and line.endswith(']'):
            current_level_type = line.strip('[]')
        elif line.lower().startswith('q') and ':' in line:
            question_text = line[line.find(':') + 1:].strip()
            
            parts = current_level_type.split('/')
            level = parts[0].strip() if parts else "Unknown"
            q_type = parts[1].strip() if len(parts) > 1 else "General"
            
            q_list.append({
                "question": f"({level}/{q_type}) {question_text}",
                "answer": "", 
                "level": level,
                "type": q_type
            })
    return q_list


def display_evaluation_form(mode, qa_data_list, context_for_eval):
    """Handles the display of the Q&A form and evaluation logic for a given mode."""
    current_qa_key = f'interview_qa_{mode}'
    current_report_key = f'evaluation_report_{mode}'
    
    if qa_data_list:
        st.markdown("---")
        st.subheader("2. Practice and Record Answers")
        
        with st.form(f"interview_practice_form_{mode}"):
            current_qa_list = st.session_state[current_qa_key]
            
            for i, qa_item in enumerate(current_qa_list):
                st.markdown(f"**Question {i+1}:** {qa_item['question']}")
                
                answer_key = f'answer_q_{mode}_{i}'
                
                answer = st.text_area(
                    f"Your Answer for Q{i+1}", 
                    value=current_qa_list[i]['answer'], 
                    height=100,
                    key=answer_key,
                    label_visibility='collapsed'
                )
                
                current_qa_list[i]['answer'] = answer 
                st.markdown("---") 
                
            submit_button = st.form_submit_button("Submit & Evaluate Answers", use_container_width=True, type="primary")

            if submit_button:
                if all(item['answer'].strip() for item in current_qa_list):
                    with st.spinner("Sending answers to AI Evaluator..."):
                        try:
                            report = evaluate_interview_answers(
                                current_qa_list,
                                context_for_eval
                            )
                            st.session_state[current_report_key] = report
                            st.success("Evaluation complete! See the report below.")
                        except Exception as e:
                            st.error(f"Evaluation failed: {e}")
                            st.session_state[current_report_key] = f"Evaluation failed: {e}\n{traceback.format_exc()}"
                else:
                    st.error("Please answer all generated questions before submitting.")
        
        if st.session_state.get(current_report_key):
            st.markdown("---")
            st.subheader("3. AI Evaluation Report")
            st.markdown(st.session_state[current_report_key])
            
        st.markdown("---")
        if st.button(f"🗑️ Clear {mode.upper()} Practice Session (Questions and Answers)", key=f"clear_interview_prep_session_{mode}"):
            clear_interview_state(mode)
            st.success("Practice session cleared.")
            st.rerun()
            
# --- Interview Preparation Tab (UPDATED) ---
def interview_preparation_tab():
    """
    Interview Preparation Tab Logic with two sub-tabs: Resume Based and JD Based.
    """
    st.header("🎤 Interview Preparation Tools")
    
    # Determine if a resume/CV is ready
    is_resume_parsed = (
        st.session_state.get('parsed') is not None and
        isinstance(st.session_state.parsed, dict) and
        st.session_state.parsed.get('error') is None and
        st.session_state.parsed.get('name') is not None
    )
    is_jd_loaded = bool(st.session_state.get('candidate_jd_list'))

    # Check if we are running in Mock Mode
    is_mock_mode = isinstance(client, MockGroqClient) and not GROQ_API_KEY
    if not GROQ_API_KEY and not is_mock_mode:
        st.error("Cannot use Interview Prep: GROQ_API_KEY is not configured.")
        return
        
    # Initialize Interview Prep States
    if 'iq_mode' not in st.session_state: st.session_state.iq_mode = 'resume' 
    if 'iq_output_resume' not in st.session_state: st.session_state.iq_output_resume = ""
    if 'interview_qa_resume' not in st.session_state: st.session_state.interview_qa_resume = [] 
    if 'evaluation_report_resume' not in st.session_state: st.session_state.evaluation_report_resume = "" 
    
    if 'iq_output_jd' not in st.session_state: st.session_state.iq_output_jd = ""
    if 'interview_qa_jd' not in st.session_state: st.session_state.interview_qa_jd = [] 
    if 'evaluation_report_jd' not in st.session_state: st.session_state.evaluation_report_jd = "" 
    
    st.markdown("---")
    tab_resume, tab_jd = st.tabs(["👤 Resume Based Q&A", "💼 JD Based Q&A"])
    
    # --- 1. RESUME BASED SUB-TAB ---
    with tab_resume:
        st.session_state.iq_mode = 'resume'
        
        if not is_resume_parsed:
            st.warning("Please upload and successfully parse a resume or compile one in 'CV Management' first.")
            return

        # Grab and structure resume sections dynamically
        raw_keys = st.session_state.parsed.keys()
        excluded_keys = {'name', 'email', 'phone', 'error', 'linkedin', 'github', 'personal_details', 'summary'}
        
        valid_sections = []
        for key in raw_keys:
            if key.lower() not in excluded_keys:
                value = st.session_state.parsed.get(key)
                if value and str(value).strip() and str(value).strip().lower() != 'none':
                    valid_sections.append(key)

        display_map = {k: k.replace('_', ' ').title() for k in valid_sections}
        question_section_options = sorted(list(display_map.values()))

        if not question_section_options:
            st.error("No deep sections (like Experience, Skills, or Projects) found with data in the parsed resume.")
            return
            
        st.subheader("1. Generate Interview Questions (Resume)")
        
        selected_display = st.selectbox(
            "Select Resume Section to Focus On", 
            question_section_options, 
            key='iq_section_resume_c',
            on_change=lambda: clear_interview_state('resume')
        )
        
        chosen_original_key = next((k for k, v in display_map.items() if v == selected_display), selected_display)
        
        if st.button("Generate Resume Questions", key='iq_btn_resume_c', use_container_width=True):
            with st.spinner("Generating questions based on resume section..."):
                try:
                    clear_interview_state('resume')

                    raw_questions_response = generate_interview_questions(
                        st.session_state.parsed, 
                        chosen_original_key
                    )
                    
                    # CATCH-ALL FOR BACKEND ERRORS: 
                    # Checks if response text looks like an error statement instead of parsed questions
                    if any(raw_questions_response.startswith(prefix) for prefix in ["Error:", "Cannot", "Failed"]):
                        st.error(raw_questions_response)
                        st.session_state.iq_output_resume = raw_questions_response
                        return

                    st.session_state.iq_output_resume = raw_questions_response
                    q_list = parse_questions_from_raw(raw_questions_response)
                    st.session_state.interview_qa_resume = q_list
                    
                    if q_list:
                        st.success(f"Generated {len(q_list)} questions based on your **{selected_display}** section.")
                    else:
                        st.warning("Could not parse any questions from the LLM response.")
                        with st.expander("Show Raw Model Response"):
                            st.code(raw_questions_response)
                    
                except Exception as e:
                    st.error(f"Error generating questions: {e}\nTrace: {traceback.format_exc()}")
                    st.session_state.iq_output_resume = "Error generating questions."
                    st.session_state.interview_qa_resume = []
        
        reference_cv_data = st.session_state.get('full_text', str(st.session_state.parsed))
        display_evaluation_form('resume', st.session_state.interview_qa_resume, reference_cv_data)

    # --- 2. JD BASED SUB-TAB ---
    with tab_jd:
        st.session_state.iq_mode = 'jd'

        if not is_jd_loaded:
            st.warning("Please load Job Descriptions in the 'JD Management' tab first.")
            return
            
        st.subheader("1. Generate Interview Questions (JD)")
        
        jd_names = [jd.get('name') for jd in st.session_state.candidate_jd_list if jd.get('name')]
        selected_jd_name = st.selectbox(
            "Select Job Description",
            options=jd_names,
            key='iq_jd_name_c',
            on_change=lambda: clear_interview_state('jd')
        )

        selected_jd = next((jd for jd in st.session_state.candidate_jd_list if jd.get('name') == selected_jd_name), None)
        
        if st.button("Generate JD Questions", key='iq_btn_jd_c', use_container_width=True):
            if not selected_jd:
                st.error("Please select a Job Description.")
                return

            with st.spinner(f"Generating questions based on JD: {selected_jd_name}..."):
                try:
                    clear_interview_state('jd')
                    
                    jd_content = selected_jd.get('content', '')
                    raw_questions_response = generate_interview_questions(
                        {"job_description_name": selected_jd_name, "content": jd_content}, 
                        "job_description"
                    )
                    
                    if any(raw_questions_response.startswith(prefix) for prefix in ["Error:", "Cannot", "Failed"]):
                        st.error(raw_questions_response)
                        st.session_state.iq_output_jd = raw_questions_response
                        return

                    st.session_state.iq_output_jd = raw_questions_response
                    q_list = parse_questions_from_raw(raw_questions_response)
                    st.session_state.interview_qa_jd = q_list
                    
                    if q_list:
                        st.success(f"Generated {len(q_list)} questions based on **{selected_jd_name}**.")
                    else:
                        st.warning("Could not parse any questions from the LLM response.")
                        with st.expander("Show Raw Model Response"):
                            st.code(raw_questions_response)
                    
                except Exception as e:
                    st.error(f"Error generating questions: {e}\nTrace: {traceback.format_exc()}")
                    st.session_state.iq_output_jd = "Error generating questions."
                    st.session_state.interview_qa_jd = []

        jd_content = selected_jd.get('content', '') if selected_jd else ""
        display_evaluation_form('jd', st.session_state.interview_qa_jd, jd_content)
        
# start  ------------------------------------------- -----------------------------------------------------------       
# ATS Scanner Optimization & Compliance Panel tab --------------------
def ats_optimization_tab():
    """Tab to grade parsing scores, deliver strategic optimizations, and display comparison matrices side-by-side."""
    st.header("Resume Score Checker")
    st.markdown("Upload your resume and optional job description to receive an instant corporate screening score.")
    st.markdown("---")

    col_left, col_right = st.columns(2)

    # --- PANEL 1: INPUT MODALITY HUB ---
    with col_left:
        st.subheader("1. Upload Your Resume")
        ats_input_method = st.radio(
            "Select Resume Source",
            ["Upload File Document", "Paste Raw Text Workspace"],
            key="ats_tab_input_modality_toggle"
        )
        
        if ats_input_method == "Upload File Document":
            uploaded_ats_res = st.file_uploader(
                "Upload Resume (PDF, DOCX, TXT)",
                type=["pdf", "docx", "txt"],
                key="ats_tab_file_uploader_widget"
            )
            
            if uploaded_ats_res is not None:
                if st.session_state.last_uploaded_file_name != uploaded_ats_res.name:
                    f_type = get_file_type(uploaded_ats_res.name)
                    uploaded_ats_res.seek(0)
                    txt_out, _ = extract_content(f_type, uploaded_ats_res.getvalue(), uploaded_ats_res.name)
                    
                    if not txt_out.startswith("[Error"):
                        st.session_state.ats_original_resume_text = txt_out
                        st.session_state.last_uploaded_file_name = uploaded_ats_res.name
                        st.session_state.ats_optimized_resume_text = ""
                        st.session_state.ats_score_calculated = False
                    else:
                        st.error(txt_out)
            
            if st.session_state.ats_original_resume_text:
                st.info(f"Active Resume Context: {st.session_state.last_uploaded_file_name or 'Uploaded Document'}")
                
        else:
            st.session_state.last_uploaded_file_name = None
            pasted_text = st.text_area(
                "Paste candidate resume structural workspace values here:",
                value=st.session_state.ats_original_resume_text,
                height=200,
                key="ats_tab_pasted_text_area_widget"
            )
            st.session_state.ats_original_resume_text = pasted_text

        st.markdown("---")
        
        st.subheader("Upload Job Description")
        jd_input_method = st.radio(
            "Select Job Description Source",
            ["Upload JD Document", "Paste Raw JD Text"],
            key="ats_tab_jd_modality_toggle"
        )
        
        if jd_input_method == "Upload JD Document":
            uploaded_jd = st.file_uploader(
                "Upload Job Description (PDF, DOCX, TXT)",
                type=["pdf", "docx", "txt"],
                key="ats_tab_jd_uploader_widget"
            )
            
            if uploaded_jd is not None:
                if st.session_state.last_uploaded_jd_name != uploaded_jd.name:
                    f_type = get_file_type(uploaded_jd.name)
                    uploaded_jd.seek(0)
                    txt_out, _ = extract_content(f_type, uploaded_jd.getvalue(), uploaded_jd.name)
                    
                    if not txt_out.startswith("[Error"):
                        st.session_state.ats_job_description_text = txt_out
                        st.session_state.last_uploaded_jd_name = uploaded_jd.name
                        st.session_state.ats_score_calculated = False
                    else:
                        st.error(txt_out)
            
            if st.session_state.ats_job_description_text:
                st.info(f"Active Job Context: {st.session_state.last_uploaded_jd_name or 'Uploaded JD Document'}")
        else:
            st.session_state.last_uploaded_jd_name = None
            pasted_jd = st.text_area(
                "Paste corporate job description workspace values here:",
                value=st.session_state.ats_job_description_text,
                height=200,
                key="ats_tab_pasted_jd_area_widget"
            )
            st.session_state.ats_job_description_text = pasted_jd

    # --- PANEL 2: COMPLIANCE INSTRUCTIONS & LOGIC TRIGGERS ---
    with col_right:
        st.subheader("2.ATS compatible Report & Resume Analysis ")
        st.markdown(
            "This scanner automatically scales calculations based on available inputs to produce either a structural audit or a targeted role alignment score."
        )
        
        if st.button("🔍 Scan and Score My Resume", type="secondary", use_container_width=True):
            if not st.session_state.ats_original_resume_text.strip():
                st.error("Validation Halt: Please provide a valid resume profile before running scanner audits.")
            else:
                with st.spinner("Running deep evaluation matrices..."):
                    payload = st.session_state.ats_original_resume_text
                    payload_lower = payload.lower()
                    jd_payload = st.session_state.ats_job_description_text.strip()
                    
                    # --- Algorithmic Extraction Vectors ---
                    has_email = "@" in payload
                    has_phone = len(re.findall(r'\b\d{4,}\b', payload_lower)) >= 1
                    has_skills = any(k in payload_lower for k in ["skills", "technical", "competencies", "expertise", "programming"])
                    has_summary = any(k in payload_lower for k in ["summary", "profile", "objective", "about me", "professional summary"])
                    has_experience = any(k in payload_lower for k in ["experience", "employment", "history", "work"])
                    has_education = any(k in payload_lower for k in ["education", "academic", "degree", "grade"])
                    has_projects = any(k in payload_lower for k in ["projects", "personal projects", "key engineering"])
                    has_location = any(k in payload_lower for k in ["india", "usa", "uk", "remote", "bangalore", "wayand", "hyderabad", "mumbai", "delhi"])
                    
                    metrics_count = len(re.findall(r'\b\d+%\b|\b\$\d+|\b\d+\s*(?:points|hours|x|gb|tb|million|k|cgpa)\b', payload_lower))
                    action_verbs = ["engineered", "optimized", "built", "designed", "implemented", "spearheaded", "architected", "developed", "deployed", "automated", "scaled", "led"]
                    verb_matches = sum(1 for verb in action_verbs if verb in payload_lower)
                    
                    is_jd_mode = bool(jd_payload)
                    
                    if is_jd_mode:
                        jd_words = set(re.findall(r'\b[a-z]{3,12}\b', jd_payload.lower()))
                        core_tech_pool = {"python", "sql", "aws", "docker", "kubernetes", "mlops", "pytorch", "tensorflow", "fastapi", "react", "java", "spark", "azure", "ci/cd", "git", "supabase", "streamlit"}
                        target_keywords = jd_words.intersection(core_tech_pool)
                        
                        if target_keywords:
                            matched_keywords = sum(1 for kw in target_keywords if kw in payload_lower)
                            keyword_match_percentage = int((matched_keywords / len(target_keywords)) * 100)
                        else:
                            keyword_match_percentage = 85
                        
                        skills_status = f"{keyword_match_percentage}% — Skill compatibility match compared to target JD requirements"
                        base_score = 30 + (keyword_match_percentage * 0.3) + (20 if metrics_count >= 3 else min(metrics_count * 7, 15)) + (20 if verb_matches >= 4 else min(verb_matches * 5, 15))
                    else:
                        skills_status = "100% (excellent)" if has_skills else "0% — Skills section not cleanly demarcated"
                        base_score = 40 + (10 if has_skills else 0) + (25 if metrics_count >= 3 else min(metrics_count * 8, 15)) + (25 if verb_matches >= 4 else min(verb_matches * 6, 15))
                    
                    personal_info = "100% (excellent)" if (has_email and has_phone) else "0% — Missing critical contact credentials"
                    titles_status = "100% (excellent)" if (has_experience and has_education and has_summary) else "50% — Headers use non-standard naming schemas"
                    location_status = "100% (excellent)" if has_location else "0% — Missing explicit location string info"
                    summary_status = "100% (excellent)" if has_summary else "0% — Missing professional summary hook section"
                    exp_struct = "100% (excellent)" if (has_experience and verb_matches >= 4) else "0% — 1 issue: Missing chronologically linear layout blocks"
                    exp_content = "100% (excellent)" if metrics_count >= 3 else "0% — 1 issue: Descriptive tasks lack hard quantifiable impact metrics"
                    edu_status = "100% (excellent)" if (has_education and ("20" in payload or "19" in payload)) else ("56% — 2 issues: Graduation timeline or calendar date fields missing" if has_education else "0% — Section unreadable")
                    proj_status = "100% (excellent)" if (has_projects and metrics_count >= 4) else ("0% — 9 issues: Core project bullets lack quantified results" if has_projects else "0% — No projects section parsed")

                    if "|" in payload or "\t" in payload:
                        base_score -= 10
                    final_score = min(max(int(base_score), 15), 98)

                    st.session_state.ats_score_metrics = {
                        "overall": final_score,
                        "mode_label": "Targeted JD Role Match Scan" if is_jd_mode else "General Structural Integrity Audit",
                        "personal_info": personal_info,
                        "skills": skills_status,
                        "summary": summary_status,
                        "titles": titles_status,
                        "location": location_status,
                        "exp_structure": exp_struct,
                        "exp_content": exp_content,
                        "education_grade": edu_status,
                        "projects_grade": proj_status
                    }
                    st.session_state.ats_score_calculated = True
                    st.rerun()

        if st.session_state.ats_score_calculated and st.session_state.ats_score_metrics:
            metrics = st.session_state.ats_score_metrics
            score = metrics.get("overall", 0)
            
            st.markdown(f"**Scan Type Engine Running:** `{metrics.get('mode_label')}`")
            if score >= 80:
                st.success(f"Your résumé ATS scored {score}/100 — Strong.")
            elif score >= 65:
                st.warning(f"Your résumé scored {score}/100 — Action Required.")
            else:
                st.error(f"Your résumé scored {score}/100 — High Risk.")
                
            col_well, col_improve = st.columns(2)
            
            with col_well:
                st.markdown("#### ✓ Working well")
                with st.container(border=True):
                    for label, key in [
                        ("Personal Information", "personal_info"),
                        ("Summary", "summary"),
                        ("Skills", "skills"),
                        ("Section Titles", "titles"),
                        ("Location Format", "location"),
                        ("Work Experience — Structure", "exp_structure"),
                        ("Work Experience — Content", "exp_content"),
                        ("Education", "education_grade"),
                        ("Projects", "projects_grade")
                    ]:
                        val = metrics.get(key, "")
                        if "100%" in val or "80%" in val or "90%" in val or "70%" in val:
                            st.markdown(f"**{label}:** {val}")
            
            with col_improve:
                st.markdown("#### → To improve")
                with st.container(border=True):
                    any_improvements = False
                    for label, key in [
                        ("Personal Information", "personal_info"),
                        ("Summary", "summary"),
                        ("Skills", "skills"),
                        ("Section Titles", "titles"),
                        ("Location Format", "location"),
                        ("Work Experience — Structure", "exp_structure"),
                        ("Work Experience — Content", "exp_content"),
                        ("Education", "education_grade"),
                        ("Projects", "projects_grade")
                    ]:
                        val = metrics.get(key, "")
                        if "100%" not in val and "80%" not in val and "90%" not in val and "70%" not in val:
                            st.markdown(f"**{label}:** {val}")
                            any_improvements = True
                    if not any_improvements:
                        st.write("🎉 None! Your structure is immaculate.")

    # --- SECTION 3: AUTOMATED RE-ARCHITECTURE PIPELINE ---
    if st.session_state.ats_original_resume_text.strip():
        st.markdown("---")
        st.subheader("3.Build Your Optimized ATS Resume")
        
        if st.button("Fix Your resume with AI in Minutes", type="primary", use_container_width=True):
            with st.spinner("Injecting core industry keywords, structuring schemas, and re-writing bullet profiles..."):
                optimized_text = optimize_resume_for_ats(
                    st.session_state.ats_original_resume_text, 
                    st.session_state.ats_job_description_text,
                    st.session_state.ats_score_metrics
                )
                
                cleaned_ats_text = optimized_text.replace('#', '').replace('*', '').strip()
                # Clean mathematical bullet points out right at storage generation phase
                cleaned_ats_text = re.sub(r'^\s*[-+*]\s+', '• ', cleaned_ats_text, flags=re.MULTILINE)
                st.session_state.ats_optimized_resume_text = cleaned_ats_text
                st.rerun()

    # --- SECTION 4: EDITABLE SIDE-BY-SIDE VERIFICATION & COMPARISON MATRIX ---
    if st.session_state.ats_optimized_resume_text:
        st.markdown("---")
        st.subheader("Compare Your Resume with Optimized AI Resume and Dowmload")
        st.caption("Review, edit, and fine-tune your compliance copy directly in the workspace prior to downloading your final PDF.")
        
        col_view_left, col_view_right = st.columns(2)
        
        with col_view_left:
            st.markdown("####  Original Upload / Pasted Resume")
            with st.container(border=True):
                st.text(st.session_state.ats_original_resume_text)
            
        with col_view_right:
            st.markdown("####  Optimized ATS Format Resume")
            
            # CRITICAL ENHANCEMENT: Changed from standard text output to an active, interactive text area
            edited_ats_resume = st.text_area(
                "Make additions, polish text strings, or review structure rules here:",
                value=st.session_state.ats_optimized_resume_text,
                height=500,
                key="ats_tab_editable_optimized_output_workspace"
            )
            # Synchronize any interactive adjustments back to session tracking state
            st.session_state.ats_optimized_resume_text = edited_ats_resume
                
            clean_name = "Candidate"
            if "parsed" in st.session_state and isinstance(st.session_state.parsed, dict) and st.session_state.parsed.get("name"):
                clean_name = st.session_state.parsed["name"].replace(" ", "_")
                
            st.markdown("##### Document Export Channels")
            col_dl_md, col_dl_pdf = st.columns(2)
            
            with col_dl_md:
                st.download_button(
                    label="⬇️ Download Optimized Resume (.txt)",
                    data=st.session_state.ats_optimized_resume_text,
                    file_name=f"{clean_name}_ATS_Optimized_Resume.txt",
                    mime="text/plain",
                    use_container_width=True,
                    key="ats_tab_optimized_md_download_widget"
                )
                
            with col_dl_pdf:
                # --- NATIVE IN-MEMORY PDF FILE SYSTEM EXPORT PIPELINE ---
                try:
                    # Ingests user modifications immediately upon file compilation trigger click
                    pdf_data = generate_pdf_bytes(st.session_state.ats_optimized_resume_text)
                    
                    st.download_button(
                        label="📄 Download PDF Profile (.pdf)",
                        data=pdf_data,
                        file_name=f"{clean_name}_ATS_Optimized_Resume.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key="ats_tab_optimized_native_pdf_download_widget"
                    )
                except Exception as e:
                    st.error(f"Native PDF compilation failed. Error context details: {str(e)}")
                    
# --- Cover Letter Generator Tab ---
def cover_letter_tab():
    """ Tab layout managing text document uploads, layout compilation settings, and downloading blocks. """
    
    # 1. Initialize State Flags Natively to maintain memory persistence across page reruns
    if 'cl_v2_cached_output_string' not in st.session_state:
        st.session_state.cl_v2_cached_output_string = ""
    if 'cl_v2_cached_signature_stamp' not in st.session_state:
        st.session_state.cl_v2_cached_signature_stamp = ""
    if "workspace_editor_content" not in st.session_state:
        st.session_state.workspace_editor_content = ""
    if "cl_edit_mode_active" not in st.session_state:
        st.session_state.cl_edit_mode_active = False

    st.header("✉️ Tailored Cover Letter Generator")
    st.markdown("Provide your core text parameters below to instantly draft a clean, high-impact cover letter.")
    st.markdown("---")

    col_left_panel, col_right_panel = st.columns(2)

    # --- SECTION 1: PROFILE/RESUME DATA PANEL ---
    with col_left_panel:
        st.subheader("1. Profile / Resume Input")
        cl_res_method = st.radio(
            "Select Resume Entry Method", 
            ["Upload File Document", "Paste Raw Text Workspace"], 
            key="cl_tab_v2_res_entry_modality_toggle"
        )
        
        resume_payload_text = ""
        if cl_res_method == "Upload File Document":
            uploaded_res = st.file_uploader(
                "Upload Resume (PDF, DOCX, TXT)", 
                type=["pdf", "docx", "txt"], 
                key="cl_tab_v2_raw_file_resume_uploader_widget"
            )
            if uploaded_res:
                f_type = get_file_type(uploaded_res.name)
                uploaded_res.seek(0)
                txt_out, _ = extract_content(f_type, uploaded_res.getvalue(), uploaded_res.name)
                if not txt_out.startswith("[Error"):
                    resume_payload_text = txt_out
                    st.success(f"Loaded Profile: {uploaded_res.name}")
                else:
                    st.error(txt_out)
        else:
            resume_payload_text = st.text_area(
                "Paste candidate resume text contents here:", 
                height=250, 
                key="cl_tab_v2_raw_pasted_text_resume_area_widget"
            )

    # --- SECTION 2: JOB DESCRIPTION POSTING PANEL ---
    with col_right_panel:
        st.subheader("2. Target Job Requirements")
        cl_jd_method = st.radio(
            "Select JD Entry Method", 
            ["Upload File Document", "Paste Raw Text Workspace"], 
            key="cl_tab_v2_jd_entry_modality_toggle"
        )
        
        jd_payload_text = ""
        if cl_jd_method == "Upload File Document":
            uploaded_jd = st.file_uploader(
                "Upload Job Description (PDF, DOCX, TXT)", 
                type=["pdf", "docx", "txt"], 
                key="cl_tab_v2_raw_file_jd_uploader_widget"
            )
            if uploaded_jd:
                f_type = get_file_type(uploaded_jd.name)
                uploaded_jd.seek(0)
                txt_out, _ = extract_content(f_type, uploaded_jd.getvalue(), uploaded_jd.name)
                if not txt_out.startswith("[Error"):
                    jd_payload_text = txt_out
                    st.success(f"Loaded Job Description Details: {uploaded_jd.name}")
                else:
                    st.error(txt_out)
        else:
            jd_payload_text = st.text_area(
                "Paste raw structural job description details here:", 
                height=250, 
                key="cl_tab_v2_raw_pasted_text_jd_area_widget"
            )

    st.markdown("---")

    # --- SECTION 3: DESIGN/TONE SELECTION MATRIX ---
    st.subheader("3. Select Blueprint Tone & Style")
    template_style = st.selectbox(
        "Design Template Tone / Blueprint Style",
        options=["Simple", "Professional", "Modern", "Creative"],
        index=1,
        key="cl_tab_v2_design_style_dropdown_selector"
    )

    current_input_signature = f"res_{hash(resume_payload_text)}_jd_{hash(jd_payload_text)}_style_{template_style}"

    if st.button("🚀 Process & Generate Cover Letter", type="primary", use_container_width=True, key="cl_tab_v2_master_process_trigger_btn"):
        if not resume_payload_text.strip():
            st.error("Validation Halt: Please provide a valid resume profile before running compilation.")
        elif not jd_payload_text.strip():
            st.error("Validation Halt: Please provide target job description requirements text.")
        else:
            with st.spinner("Processing documents content parameters and engineering structural matching layout layouts..."):
                compiled_result = generate_tailored_cover_letter(
                    resume_text=resume_payload_text,
                    jd_content=jd_payload_text,
                    template_style=template_style,
                    cache_bust=current_input_signature
                )
                
                # Overwrite master caches cleanly on initial generation trigger
                st.session_state.cl_v2_cached_output_string = compiled_result
                st.session_state.workspace_editor_content = compiled_result  
                st.session_state.cl_v2_cached_signature_stamp = current_input_signature
                st.session_state.cl_edit_mode_active = False # Reset edit layout state flags
                st.rerun()

    # --- SECTION 4: INTERACTIVE CANVAS DISPLAY WORKSPACE ---
    if st.session_state.cl_v2_cached_output_string:
        st.markdown("---")
        st.subheader("📝 Live Cover Letter Workspace Canvas")
        
        # Guard rails for processing signature configuration tracking changes
        if current_input_signature != st.session_state.cl_v2_cached_signature_stamp:
            st.caption("⚠️ *Data drift notice: Inputs have changed since this layout was drafted. Click generate to rebuild.*")

        # Action Buttons to toggle between Viewing Mode and Editing Mode
        col_edit, col_submit = st.columns(2)
        
        if not st.session_state.cl_edit_mode_active:
            if col_edit.button("✏️ Edit Cover Letter", use_container_width=True):
                st.session_state.cl_edit_mode_active = True
                st.rerun()
        else:
            if col_submit.button("💾 Confirm & Save Changes", type="primary", use_container_width=True):
                st.session_state.cl_edit_mode_active = False
                # Locking inside current text matrix memory state
                st.session_state.cl_v2_cached_output_string = st.session_state.workspace_editor_content
                st.success("✅ Changes saved successfully! Your custom updates are now locked into the download payloads.")
                st.rerun()

        # Render layout conditionally based on Active state operations flags
        if st.session_state.cl_edit_mode_active:
            st.caption("✏️ *Editing Mode Active: Type below to append or remove elements safely.*")
            # Editable State: Renders a real text area look looking at live keystrokes
            st.text_area(
                "Modify text elements or overwrite placeholder values freely inside the editor canvas below:",
                value=st.session_state.workspace_editor_content,
                height=500,
                key="workspace_editor_content",
                label_visibility="collapsed"
            )
        else:
            st.caption("🔒 *Viewing Mode (Changes Locked): Click 'Edit Cover Letter' above to change text contents.*")
            # Static Safe State: Displays text cleanly so it stays frozen across general download loops
            st.info(st.session_state.cl_v2_cached_output_string)

        # File export naming configurations definitions
        cand_name, role_title, _ = extract_basic_entities(resume_payload_text, jd_payload_text)
        clean_name = cand_name.replace(' ', '_') if isinstance(cand_name, str) else "Candidate"
        clean_role = role_title.replace(' ', '_').replace('/', '_')
        base_export_filename = f"{clean_name}_CoverLetter_{clean_role}"

        st.markdown("##### Document Export Channels")
        col_dl_md, col_dl_html = st.columns(2)
        
        with col_dl_md:
            st.download_button(
                label="⬇️ Download Markdown Document (.md)",
                data=st.session_state.cl_v2_cached_output_string, # Strictly matches frozen or verified text configurations
                file_name=f"{base_export_filename}.md",
                mime="text/markdown",
                key="cl_tab_v2_md_download_action_button_widget",
                use_container_width=True
            )
            
        with col_dl_html:
            html_uri_link = get_download_link(
                data=st.session_state.cl_v2_cached_output_string, # Strictly matches frozen or verified text configurations
                filename=f"{base_export_filename}.html",
                file_format='html',
                title="Tailored Resume Cover Letter Documentation"
            )
            render_download_button(
                data_uri=html_uri_link,
                filename=f"{base_export_filename}.html",
                label="📄 Download HTML Profile (Print to PDF)",
                color='html'
            )
# --------------------------------------------------------------------------------------
# NEW TAB: GAP ANALYSIS & COURSE PLAN
# --------------------------------------------------------------------------------------
def gap_analysis_tab():
    """
    Tab to analyze gaps from the top matched JD and generate a course plan.
    """
    st.header("💡 Gap Analysis & Course Plan")
    st.markdown("This tool analyzes your biggest skill gaps from your best-matched Job Description and suggests a course plan and certifications to close the gap.")
    st.markdown("---")

    is_resume_parsed = (
        st.session_state.get('parsed', {}).get('name') is not None and 
        st.session_state.parsed.get('error') is None
    )

    if not is_resume_parsed:
        st.warning("⚠️ **Course Plan Disabled:** Please upload and successfully parse a resume or compile one in 'CV Management' first.")
        return
        
    if not st.session_state.get('candidate_match_results'):
        st.error("❌ **Course Plan Disabled:** Please run the **Batch JD Match** analysis first to identify your best fit JD.")
        return

    # 1. Identify the Top Matched JD
    top_match = st.session_state.candidate_match_results[0]
    top_jd_name = top_match['jd_name']
    
    # Extract the full JD content for context
    top_jd_item = next((jd for jd in st.session_state.candidate_jd_list if jd.get('name') == top_jd_name), None)
    
    if not top_jd_item:
        st.error("Could not find the full JD content for the top match. Please re-run the Batch Match.")
        return

    # Extract the Gaps/Areas for Improvement section from the full analysis output
    gaps_content = top_match.get('gaps', 'Error: Gaps analysis not found.')
    
    if 'gap_analysis_plan' not in st.session_state:
        st.session_state.gap_analysis_plan = ""

    st.subheader("1. Top Match Analysis")
    st.info(f"The analysis focuses on your best-matching JD: **{top_jd_name}** (Score: **{top_match['overall_score']}/10**)")
    
    st.markdown("##### Identified Skill Gaps from AI Match Report:")
    if "No significant gaps identified" in gaps_content or gaps_content.startswith("Error"):
        st.warning(gaps_content)
        gap_summary = "No immediate, specific technical gaps found. Focus on general upskilling for the target role."
    else:
        st.markdown(gaps_content)
        gap_summary = gaps_content.replace('\n', ' ').strip()
        
    st.markdown("---")

    st.subheader("2. Generate Detailed Course Plan")
    
    if st.button("🚀 Generate Course Plan & Certifications", use_container_width=True, type="primary"):
        with st.spinner(f"Generating comprehensive course plan for **{top_jd_name}**..."):
            candidate_skills = st.session_state.parsed.get('skills', [])
            
            plan = generate_gap_course_plan(
                gap_analysis_text=gap_summary,
                jd_role=top_jd_item.get('role', 'Target Role'),
                candidate_skills=candidate_skills
            )
            st.session_state.gap_analysis_plan = plan
            st.rerun()

    st.markdown("---")
    
    if st.session_state.gap_analysis_plan:
        st.subheader("3. AI-Generated 'How to Fill the Gap' Plan")
        
        if st.session_state.gap_analysis_plan.startswith("AI Generation Error") or st.session_state.gap_analysis_plan.startswith("No specific gaps"):
            st.error(st.session_state.gap_analysis_plan)
        else:
            st.markdown(st.session_state.gap_analysis_plan)
            
        st.markdown("---")
        
        # Download button for the plan
        plan_filename = f"{st.session_state.parsed['name'].replace(' ', '_')}_GapPlan_{top_jd_item.get('role', 'Job').replace('/', '_').replace(' ', '_')}.md"
        plan_data_uri = get_download_link(st.session_state.gap_analysis_plan, plan_filename, 'markdown', title="Gap Analysis Course Plan")

        col_dl, _ = st.columns([1, 3])
        with col_dl:
            render_download_button(
                plan_data_uri, 
                plan_filename, 
                "⬇ Third-party course plan roadmap verification document (.md)", 
                'markdown'
            )
    else:
        st.info("Click the 'Generate Course Plan & Certifications' button above to get your personalized study roadmap.")


# --- Chatbot logic operators ---
def qa_on_resume(question):
    """Chatbot for Resume (Q&A) using LLM."""
    global client, GROQ_MODEL, GROQ_API_KEY
    
    if not GROQ_API_KEY and not isinstance(client, MockGroqClient):
        return "AI Chatbot Disabled: GROQ_API_KEY not set."
        
    parsed_json = st.session_state.parsed
    full_text = st.session_state.full_text
    
    if not parsed_json or parsed_json.get('error') is not None:
         return "Please parse a valid resume first to enable the Q&A feature."

    prompt = f"""Given the following resume information:
    Resume Text: {full_text}
    Parsed Resume Data (JSON): {json.dumps(parsed_json, indent=2)}
    Answer the following question about the resume concisely and directly.
    If the information is not present, state that clearly and briefly (e.g., 'Information not found on the resume.').
    Question: {question}
    """
    
    try:
        response = client.chat.completions.create(
            model=GROQ_MODEL, 
            messages=[{"role": "user", "content": prompt}], 
            temperature=0.4
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"AI Chatbot Error: Failed to get response from LLM. Error: {e}"


def qa_on_jd(question, jd_content):
    """Chatbot for Job Description (Q&A) using LLM."""
    global client, GROQ_MODEL, GROQ_API_KEY
    
    if not GROQ_API_KEY and not isinstance(client, MockGroqClient):
        return "AI Chatbot Disabled: GROQ_API_KEY not set."

    if not jd_content or not jd_content.strip():
        return "Please select a valid Job Description to chat about."

    prompt = f"""Given the following Job Description (JD) text:
    Job Description Text: {jd_content}
    Answer the following question about the Job Description concisely and directly.
    If the information is not present, state that clearly and briefly (e.g., 'The JD does not specify that information.').
    Question: {question}
    """
    
    try:
        response = client.chat.completions.create(
            model=GROQ_MODEL, 
            messages=[{"role": "user", "content": prompt}], 
            temperature=0.4
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"AI Chatbot Error: Failed to get response from LLM. Error: {e}"


def resume_qa_content():
    """Content for the Resume Q&A sub-tab."""
    st.subheader("👤 Resume Q&A Chatbot")
    st.markdown("Ask specific questions about the currently loaded resume.")

    is_data_loaded_and_valid = (
        st.session_state.get('parsed', {}).get('name') is not None and 
        st.session_state.get('parsed', {}).get('error') is None
    )
    
    if not is_data_loaded_and_valid:
        st.warning("⚠️ **Q&A Disabled:** Please parse a valid resume in the 'Resume Parsing' or 'CV Management' tab first.")
        return
    
    if "resume_chatbot_history" not in st.session_state:
        st.session_state.resume_chatbot_history = []

    st.info(f"Chatting about: **{st.session_state.parsed['name']}**")
    st.markdown("---")
    
    for message in st.session_state.resume_chatbot_history:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    if prompt := st.chat_input("Ask a question about the resume...", key="resume_qa_input"):
        st.session_state.resume_chatbot_history.append({"role": "user", "content": prompt})
        
        with st.chat_message("user"):
            st.markdown(prompt)
        
        with st.spinner("Thinking..."):
            ai_response = qa_on_resume(prompt)

        with st.chat_message("assistant"):
            st.markdown(ai_response)
            
        st.session_state.resume_chatbot_history.append({"role": "assistant", "content": ai_response})
        st.rerun()

    if st.session_state.resume_chatbot_history:
        st.markdown("---")
        if st.button("🗑️ Clear Resume Chat History", key="clear_resume_chatbot_history"):
            st.session_state.resume_chatbot_history = []
            st.rerun()


def jd_qa_content():
    """Content for the JD Q&A sub-tab."""
    st.subheader("💼 JD Q&A Chatbot")
    st.markdown("Select a Job Description and ask questions about its requirements.")

    if not st.session_state.get('candidate_jd_list'):
        st.warning("⚠️ **Q&A Disabled:** Please load Job Descriptions in the 'JD Management' tab first.")
        return

    jd_names = [jd.get('name') for jd in st.session_state.candidate_jd_list if jd.get('name')]
    selected_jd_name = st.selectbox(
        "Select Job Description",
        options=jd_names,
        key="selected_jd_for_qa"
    )

    if "jd_chatbot_history" not in st.session_state:
        st.session_state.jd_chatbot_history = {} 

    selected_jd = next((jd for jd in st.session_state.candidate_jd_list if jd.get('name') == selected_jd_name), None)
    jd_content = selected_jd.get('content', '') if selected_jd else ""

    current_jd_history = st.session_state.jd_chatbot_history.setdefault(selected_jd_name, [])

    st.info(f"Chatting about: **{selected_jd_name}** (Role: {selected_jd.get('role', 'N/A')})")
    st.markdown("---")

    for message in current_jd_history:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    if prompt := st.chat_input(f"Ask about the requirements of: {selected_jd_name}...", key="jd_qa_input"):
        current_jd_history.append({"role": "user", "content": prompt})
        
        with st.chat_message("user"):
            st.markdown(prompt)
        
        with st.spinner("Thinking..."):
            ai_response = qa_on_jd(prompt, jd_content)

        with st.chat_message("assistant"):
            st.markdown(ai_response)
            
        current_jd_history.append({"role": "assistant", "content": ai_response})
        st.rerun()

    if current_jd_history:
        st.markdown("---")
        if st.button(f"🗑️ Clear Chat History for {selected_jd_name}", key="clear_jd_chatbot_history"):
            st.session_state.jd_chatbot_history[selected_jd_name] = []
            st.rerun()


def chatbot_tab_content():
    """Main Content for the Chatbot Tab with sub-tabs."""
    st.header("🤖 AI Chatbot Assistant")
    
    tab_resume, tab_jd = st.tabs(["👤 Resume Q&A", "💼 JD Q&A"])
    
    with tab_resume:
        resume_qa_content()
        
    with tab_jd:
        jd_qa_content()


# --- Main Engine Dashboard Form Setup Layout ---
def candidate_dashboard():
    # Set page config once at the start
    st.set_page_config(layout="wide", page_title="PragyanAI Candidate Dashboard")
    
    st.title("🧑‍💻 Candidate Dashboard")
    st.markdown("---")

    # --- Session State Initialization ---
    if "parsed" not in st.session_state: st.session_state.parsed = {} 
    if "full_text" not in st.session_state: st.session_state.full_text = ""
    if "excel_data" not in st.session_state: st.session_state.excel_data = None
    if "candidate_uploaded_resumes" not in st.session_state: st.session_state.candidate_uploaded_resumes = []
    if "pasted_cv_text" not in st.session_state: st.session_state.pasted_cv_text = ""
    if "current_parsing_source_name" not in st.session_state: st.session_state.current_parsing_source_name = None 
    if "form_cv_text" not in st.session_state: st.session_state.form_cv_text = ""
    
    if "cv_data" not in st.session_state:
        st.session_state.cv_data = {
            'personal_info': {'name': '', 'email': '', 'phone': '', 'address': ''},
            'education': [],
            'experience': [],
            'projects': [],
            'certifications': [],
            'strengths_raw': '' 
        }
    
    if "candidate_jd_list" not in st.session_state: st.session_state.candidate_jd_list = []
    if "candidate_match_results" not in st.session_state: st.session_state.candidate_match_results = []
    if 'filtered_jds_display' not in st.session_state: st.session_state.filtered_jds_display = []
    if 'last_selected_skills' not in st.session_state: st.session_state.last_selected_skills = []
    if 'generated_cover_letter' not in st.session_state: st.session_state.generated_cover_letter = "" 
    if 'cl_jd_name' not in st.session_state: st.session_state.cl_jd_name = "" 
    
    # --- INTERVIEW Preparation States ---
    if 'iq_mode' not in st.session_state: st.session_state.iq_mode = 'resume' 
    if 'iq_output_resume' not in st.session_state: st.session_state.iq_output_resume = ""
    if 'interview_qa_resume' not in st.session_state: st.session_state.interview_qa_resume = [] 
    if 'evaluation_report_resume' not in st.session_state: st.session_state.evaluation_report_resume = "" 
    
    if 'iq_output_jd' not in st.session_state: st.session_state.iq_output_jd = ""
    if 'interview_qa_jd' not in st.session_state: st.session_state.interview_qa_jd = [] 
    if 'evaluation_report_jd' not in st.session_state: st.session_state.evaluation_report_jd = "" 
    
    # --- NEW GAP ANALYSIS STATE ---
    if 'gap_analysis_plan' not in st.session_state: st.session_state.gap_analysis_plan = ""
    
    if "resume_chatbot_history" not in st.session_state: st.session_state.resume_chatbot_history = []
    if "jd_chatbot_history" not in st.session_state: st.session_state.jd_chatbot_history = {} 
    
    if 'candidate_job_types' not in st.session_state: 
        st.session_state.candidate_job_types = DEFAULT_JOB_TYPES 

    # --- Main Content Tabs Entry Points Layout Pipeline ---
    tab_parsing, tab_cv_management, tab_data_view, tab_jd, tab_batch_match, tab_filter_jd, tab_ats_optimization, tab_cover_letter, tab_chatbot, tab_interview_prep, tab_gap_analysis = st.tabs(
        [
            "📄 Resume Parsing", 
            "📝 Resume or CV Builder", 
            "✨ Parsed Data View", 
            "📚 JD Management", 
            "🎯 Batch JD Match", 
            "🔍 Filter JD", 
            "🎯 ATS Tool Optimization",
            "✉️ Cover Letters",
            "🤖 Chatbot",  
            "🎤 Interview Preparation",
            "💡 Gap Analysis & Course Plan" 
        ]
    )
    
    with tab_parsing:
        resume_parsing_tab()
        
    with tab_cv_management:
        cv_management_tab()
        
    with tab_data_view:
        parsed_data_tab()
        
    with tab_jd:
        jd_management_tab_candidate()
        
    with tab_batch_match:
        jd_batch_match_tab()
        
    with tab_filter_jd:
        filter_jd_tab_content()
        
    # ... keep your existing tab content routers as they are, and add this at the bottom:
    with tab_ats_optimization:
        ats_optimization_tab()          # ROUTING CONTENT HOOK CALL EXECUTION
        
    with tab_cover_letter:
        cover_letter_tab()
        
    with tab_chatbot:
        chatbot_tab_content()
        
    with tab_interview_prep:
        interview_preparation_tab() 
        
    with tab_gap_analysis:
        gap_analysis_tab()


# --- Execution Hook ---
if __name__ == '__main__':
    candidate_dashboard()
