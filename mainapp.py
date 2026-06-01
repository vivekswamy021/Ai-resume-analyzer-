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

GROQ_MODEL = "llama-3.1-8b-instant"
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

# Dummy helper functions for CV Management Tab preview generation
def convert_to_json(data):
    return json.dumps(data, indent=4)

def convert_to_html_content(data):
    return f"<h3>{data['personal_info'].get('name', 'CV Preview')}</h3><p>Email: {data['personal_info'].get('email', '')}</p>"


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
                    return type('MockResponse', (object,), {'choices': [type('Choice', (object,), {'message': type('Message', (object,), {'content': mock_questions_raw})})()]})

                elif "Evaluate the candidate's answers to the following questions" in prompt_content:
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
                
                message_obj = type('Message', (object,), {'content': json.dumps(mock_llm_json)})()
                choice_obj = type('Choice', (object,), {'message': message_obj})()
                response_obj = type('MockResponse', (object,), {'choices': [choice_obj]})()
                return response_obj
        
        class FitCompletions(Completions):
            def create(self, **kwargs):
                prompt_content = kwargs.get('messages', [{}])[0].get('content', '')
                
                if "Evaluate how well the following resume content matches the provided job description" in prompt_content:
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
                        
                    skills_p = 50 + (score * 5)
                    exp_p = 60 + (score * 3)
                    edu_p = 70 + (score * 1)
                    
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
                
                return super().create(**kwargs)

        return FitCompletions()

# Initialize the Groq client or the Mock client based on the environment variable
try:
    from groq import Groq
    
    if GROQ_API_KEY:
        client = Groq(api_key=GROQ_API_KEY)
        class GroqPlaceholder(Groq): 
             def __init__(self, api_key): 
                 super().__init__(api_key=api_key)
                 self.client_ready = True
        client = GroqPlaceholder(api_key=GROQ_API_KEY)
    else:
        raise ValueError("GROQ_API_KEY not set. Using Mock Client.")
        
except (ImportError, ValueError, NameError) as e:
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
    
    if 'gap_analysis_plan' in st.session_state: del st.session_state['gap_analysis_plan']


def get_file_type(file_name):
    """Identifies the file type based on its extension."""
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
    """Sends resume text to the LLM for structured information extraction."""
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
    - Name, - Email, - Phone, - Skills (list), - Education (list of degrees/institutions/dates), 
    - Experience (list of job roles/companies/dates/responsibilities), - Certifications (list), 
    - Projects (list of project names/descriptions/technologies), - Strength (list of personal strengths/qualities), 
    - Personal Details (e.g., address, date of birth, nationality), - Github (URL), - LinkedIn (URL)
    
    Resume Text:
    {text}
    
    Provide the output strictly as a JSON object.
    """
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
                json_str = json_str[:-len('
```')]
            json_str = json_str.strip()
            parsed = json.loads(json_str)
        else:
            raise json.JSONDecodeError("Could not isolate a valid JSON structure from LLM response.", content, 0)
        
        if not parsed.get('name'):
            parsed['name'] = get_fallback_name()
        parsed['error'] = None 
        return parsed

    except json.JSONDecodeError as e:
        return {"name": get_fallback_name(), "error": f"JSON decoding error from LLM. Segment:\n{content[:200]}"}
    except Exception as e:
        return {"name": get_fallback_name(), "error": f"LLM API interaction error: {e}"}

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
        extracted_text = content_source.strip()
        file_name = "Form_Compiled_CV"
        st.session_state.current_parsing_source_name = file_name

    if extracted_text.startswith("[Error"):
        return {"error": extracted_text, "full_text": extracted_text, "excel_data": None, "name": file_name}
    
    parsed_data = parse_resume_with_llm(extracted_text)
    if parsed_data.get('error') is not None: 
        return {"error": parsed_data['error'], "full_text": extracted_text, "excel_data": excel_data, "name": parsed_data.get('name', file_name)}

    compiled_text = ""
    for k, v in parsed_data.items():
        if v and k not in ['error']:
            compiled_text += f"## {k.replace('_', ' ').title()}\n\n"
            if isinstance(v, list):
                compiled_text += "\n".join([f"* {str(item)}" for item in v]) + "\n\n"
            else:
                compiled_text += str(v) + "\n\n"

    return {
        "parsed": parsed_data, 
        "full_text": compiled_text, 
        "excel_data": excel_data, 
        "name": parsed_data.get('name', 'Unknown_Candidate').replace(' ', '_')
    }

def get_download_link(data, filename, file_format, title="Parsed Data"):
    """Generates a base64 encoded download link for the given data and format."""
    mime_type = "application/octet-stream"
    if file_format in ('json', 'markdown', 'text'):
        data_bytes = data.encode('utf-8')
        if file_format == 'json': mime_type = "application/json"
        elif file_format == 'markdown': mime_type = "text/markdown"
        else: mime_type = "text/plain"
            
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
        <div class="cover-letter">{data.replace('\n', '<br>')}</div>
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
    if color == 'json': bg_color = "#4CAF50"; icon = "💾"
    elif color == 'markdown': bg_color = "#008CBA"; icon = "⬇️"
    elif color == 'html': bg_color = "#f44336"; icon = "📄"
    else: bg_color = "#555555"; icon = ""
        
    st.markdown(
        f"""
        <a href="{data_uri}" download="{filename}" style="text-decoration: none;">
            <button style="background-color: {bg_color}; color: white; border: none; padding: 10px; text-align: center; display: inline-block; font-size: 14px; margin: 4px 0; cursor: pointer; border-radius: 4px; width: 100%;">
                {icon} {label}
            </button>
        </a>
        """, 
        unsafe_allow_html=True
    )
    
# --- END HELPER FUNCTIONS ---


# --- LLM Functions ---

@st.cache_data(show_spinner="Analyzing JD for metadata...")
def extract_jd_metadata(jd_text):
    """Extracts metadata (Role, Skills, Job Type) from JD text."""
    if isinstance(jd_text, str) and jd_text.startswith("[Error"):
        return {"role": "Extraction Error", "key_skills": ["Error"], "job_type": "Error"}
    
    if not isinstance(jd_text, str):
        jd_text = str(jd_text)
    
    role_match = re.search(r'(?:Role|Position|Title|Engineer|Scientist)[:\s\n]+([\w\s/-]+)', jd_text, re.IGNORECASE)
    role = role_match.group(1).strip() if role_match else "Software Engineer (Mock)"
    
    skills_match = re.findall(r'(Python|Java|SQL|AWS|Docker|Kubernetes|React|Streamlit|Cloud|Data|ML|LLM|MLOps|Visualization|Deep Learning|TensorFlow|Pytorch|Terraform|GCP|EKS)', jd_text, re.IGNORECASE)
    if 'data scientist' in jd_text.lower() or 'machine learning' in jd_text.lower():
         role = "Data Scientist/ML Engineer"
    elif 'cloud engineer' in jd_text.lower() or 'aws' in jd_text.lower() or 'gcp' in jd_text.lower():
         role = "Cloud Engineer"
    
    job_type_match = re.search(r'(Full-time|Part-time|Contract|Remote|Hybrid)', jd_text, re.IGNORECASE)
    job_type = job_type_match.group(1) if job_type_match else "Full-time (Mock)"
    
    return {
        "role": role, 
        "key_skills": list(set([s.lower() for s in skills_match])), 
        "job_type": job_type
    }

def evaluate_jd_fit(job_description, parsed_json):
    """Evaluates how well a resume fits a given job description."""
    global client, GROQ_MODEL, GROQ_API_KEY
    if parsed_json.get('error') is not None: 
         return f"Cannot evaluate due to resume parsing errors: {parsed_json['error']}"

    if isinstance(client, MockGroqClient) or not GROQ_API_KEY:
         response = client.chat().create(model=GROQ_MODEL, messages=[{"role": "user", "content": f"Evaluate how well the following resume content matches the provided job description: {job_description}"}])
         return response.choices[0].message.content.strip()

    if not job_description.strip(): return "Please paste a job description."

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
    
    Provide a detailed evaluation structured strictly as follows:
    Overall Fit Score: [Score]/10
    
    --- Section Match Analysis ---
    Skills Match: [XX]%
    Experience Match: [YY]%
    Education Match: [ZZ]%
    
    Strengths/Matches:
    - Point 1
    
    Gaps/Areas for Improvement:
    - Point 1
    
    Overall Summary: [Summary text]
    """
    try:
        response = client.chat.completions.create(
            model=GROQ_MODEL, 
            messages=[{"role": "user", "content": prompt}], 
            temperature=0.3
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"AI Evaluation Error: {e}"

def generate_gap_course_plan(gap_analysis_text, jd_role, candidate_skills):
    """Generates a detailed study roadmap and certification suggestions to close skill gaps."""
    global client, GROQ_MODEL, GROQ_API_KEY
    if not gap_analysis_text.strip() or "No significant gaps" in gap_analysis_text:
        return "No specific gaps were identified. Focus on advanced skills in your core area."
        
    if isinstance(client, MockGroqClient) or not GROQ_API_KEY:
         response = client.chat().create(model=GROQ_MODEL, messages=[{"role": "user", "content": f"Generate a detailed course plan and suggest relevant certifications for Gaps Identified: {gap_analysis_text}"}])
         return response.choices[0].message.content.strip()

    prompt = f"""You are an expert career consultant. Target Role: {jd_role}. Gaps: {gap_analysis_text}. Current Skills: {candidate_skills}.
    Generate a detailed course plan with Chronological Phases (weeks included) and 2-3 Suggested Certifications using Markdown headings '## Detailed Course Plan' and '## Suggested Certifications'.
    """
    try:
        response = client.chat.completions.create(
            model=GROQ_MODEL, 
            messages=[{"role": "user", "content": prompt}], 
            temperature=0.6 
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"AI Generation Error: {e}"


# --- INTERVIEW PREPARATION CONTROLLERS ---

def generate_interview_questions(source_data, source_type, identifier):
    """Generates explicit interview questions based on either resume analytics or a target JD."""
    global client, GROQ_MODEL
    if source_type == 'resume':
        target_section_key = identifier.lower().replace(' ', '_')
        resume_content = source_data.get(target_section_key, "")
        content_str = "\n".join([str(i) for i in resume_content]) if isinstance(resume_content, list) else str(resume_content)
        context_block = f"Resume section '{identifier}':\n{content_str}"
    else:
        context_block = f"Job Description content:\n{identifier}"

    prompt = f"""You are an interviewer. Based on:\n{context_block}\nGenerate 6-8 questions covering HR-related, Experience-based, Situation-based, and Technical types across Basic, Intermediate, and Advanced levels.
    Format output lines exactly like:
    [Level Name/Question Type]
    Q1: Question text...
    """
    try:
        if isinstance(client, MockGroqClient) or not GROQ_API_KEY:
             response = client.chat().create(model=GROQ_MODEL, messages=[{"role": "user", "content": prompt}])
        else:
            response = client.chat.completions.create(model=GROQ_MODEL, messages=[{"role": "user", "content": prompt}], temperature=0.8)
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Error generating questions: {e}"

def evaluate_interview_answers(qa_list, resume_context):
    """Evaluates user input answers to mock interview challenges."""
    global client, GROQ_MODEL
    qa_exchange = ""
    for i, item in enumerate(qa_list):
        qa_exchange += f"Q{i+1}: {item['question']}\nAnswer {i+1}: {item['answer']}\n---\n"

    prompt = f"Evaluate these interview responses given context: {resume_context}\nExchange:\n{qa_exchange}\nProvide an Overall Score (X/10), Performance Summary, and heading critiques per question."
    try:
        if isinstance(client, MockGroqClient) or not GROQ_API_KEY:
             response = client.chat().create(model=GROQ_MODEL, messages=[{"role": "user", "content": prompt}])
        else:
            response = client.chat.completions.create(model=GROQ_MODEL, messages=[{"role": "user", "content": prompt}], temperature=0.5)
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Evaluation Error: {e}"


# --- TAB RENDERING METHODS ---
    
def resume_parsing_tab():
    st.header("📄 Resume Upload and Parsing")
    input_method = st.radio("Select Input Method", ["Upload File", "Paste Text"], key="parsing_input_method")
    st.markdown("---")

    if input_method == "Upload File":
        uploaded_file = st.file_uploader("Choose a document", type=["pdf", "docx", "txt", "json", "md", "csv", "xlsx", "markdown", "rtf"], key='candidate_file_upload_main')
        if uploaded_file is not None:
            if "candidate_uploaded_resumes" not in st.session_state or not st.session_state.candidate_uploaded_resumes or st.session_state.candidate_uploaded_resumes[0].name != uploaded_file.name:
                st.session_state.candidate_uploaded_resumes = [uploaded_file]
                st.toast("Resume file uploaded successfully.")
        
        file_to_parse = st.session_state.candidate_uploaded_resumes[0] if st.session_state.get("candidate_uploaded_resumes") else None
        if file_to_parse:
            if st.button(f"Parse and Load: {file_to_parse.name}", use_container_width=True):
                result = parse_and_store_resume(file_to_parse, 'single_resume_candidate', 'file')
                if result.get('error') is None:
                    st.session_state.parsed = result['parsed']
                    st.session_state.full_text = result['full_text']
                    st.session_state.excel_data = result['excel_data'] 
                    st.session_state.parsed['name'] = result['name'] 
                    clear_interview_state('resume')
                    clear_interview_state('jd')
                    st.success(f"✅ Successfully loaded and parsed {result['name']}.")
                    st.rerun()

    else:
        pasted_text = st.text_area("Copy and paste CV structure here", value=st.session_state.get('pasted_cv_text', ''), height=300)
        st.session_state.pasted_cv_text = pasted_text 
        if st.button("Parse and Load Pasted Text", use_container_width=True):
            if pasted_text.strip():
                result = parse_and_store_resume(pasted_text, 'single_resume_candidate', 'text')
                if result.get('error') is None:
                    st.session_state.parsed = result['parsed']
                    st.session_state.full_text = result['full_text']
                    st.session_state.parsed['name'] = result['name'] 
                    clear_interview_state('resume')
                    clear_interview_state('jd')
                    st.success("✅ Successfully loaded structural text profiles.")
                    st.rerun()

def cv_management_tab():
    st.header("📝 CV Management & Form Generation")
    st.markdown("Generate structured CV properties inside structural storage frameworks natively.")
    
    # Structural Safety Check
    if 'cv_data' not in st.session_state:
        st.session_state.cv_data = {
            'personal_info': {'name': '', 'email': '', 'phone': '', 'address': ''},
            'education': [], 'experience': [], 'projects': [], 'certifications': [], 'strengths_raw': ''
        }
    if 'form_cv_text' not in st.session_state:
        st.session_state.form_cv_text = ""

    st.subheader("1. Personal Information")
    col_name, col_email, col_phone = st.columns(3)
    with col_name: st.session_state.cv_data['personal_info']['name'] = st.text_input("Full Name", value=st.session_state.cv_data['personal_info'].get('name', ''))
    with col_email: st.session_state.cv_data['personal_info']['email'] = st.text_input("Email", value=st.session_state.cv_data['personal_info'].get('email', ''))
    with col_phone: st.session_state.cv_data['personal_info']['phone'] = st.text_input("Phone Number", value=st.session_state.cv_data['personal_info'].get('phone', ''))
    st.session_state.cv_data['personal_info']['address'] = st.text_input("Address Mapping (Optional)", value=st.session_state.cv_data['personal_info'].get('address', ''))

    st.subheader("2. Education Matrix")
    with st.form("education_form", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns([2, 2, 1, 1])
        deg = c1.text_input("Degree")
        uni = c2.text_input("Institution")
        sc = c3.text_input("GPA/%")
        yr = c4.text_input("Year")
        if st.form_submit_button("Add Education"):
            if deg and uni:
                st.session_state.cv_data['education'].append(f"Degree: {deg}, Uni: {uni}, Score: {sc} ({yr})")
    if st.session_state.cv_data['education']:
        st.dataframe(st.session_state.cv_data['education'], use_container_width=True)

    st.subheader("3. Project Records")
    with st.form("projects_form", clear_on_submit=True):
        p_name = st.text_input("Project Name")
        p_desc = st.text_area("Accomplishments / Technology Frameworks")
        if st.form_submit_button("Add Project Entry"):
            if p_name: st.session_state.cv_data['projects'].append(f"Project: {p_name} - {p_desc}")
    if st.session_state.cv_data['projects']:
        st.dataframe(st.session_state.cv_data['projects'], use_container_width=True)

    if st.button("Generate Template Markdown", type="primary", use_container_width=True):
        data = st.session_state.cv_data
        md = f"# {data['personal_info']['name']}\nEmail: {data['personal_info']['email']} | Phone: {data['personal_info']['phone']}\n\n## Education\n"
        md += "\n".join([f"* {i}" for i in data['education']]) + "\n\n## Projects\n" + "\n".join([f"* {i}" for i in data['projects']])
        st.session_state.form_cv_text = md
        st.success("MD Generated! Copy text or download using tabs below.")

    if st.session_state.form_cv_text:
        t_md, t_json = st.tabs(["Markdown", "JSON"])
        t_md.code(st.session_state.form_cv_text, language='markdown')
        t_json.json(st.session_state.cv_data)

def jd_management_tab_candidate():
    st.header("📚 Manage Job Descriptions for Matching")
    if "candidate_jd_list" not in st.session_state: st.session_state.candidate_jd_list = []

    method = st.radio("Choose Method", ["Paste Text", "Upload File"], key="jd_method")
    if method == "Paste Text":
        with st.form("jd_paste_form"):
            txt = st.text_area("Paste job profile requirements here")
            if st.form_submit_button("Save Job Profile"):
                if txt.strip():
                    meta = extract_jd_metadata(txt)
                    st.session_state.candidate_jd_list.append({"name": f"JD - {meta['role']}", "content": txt, **meta})
                    st.success("Job Description cached successfully!")
                    st.rerun()
    else:
        fl = st.file_uploader("Upload job specification text", type=["pdf","txt","docx"])
        if fl and st.button("Extract and Cache File Profile"):
            text, _ = extract_content(get_file_type(fl.name), fl.getvalue(), fl.name)
            meta = extract_jd_metadata(text)
            st.session_state.candidate_jd_list.append({"name": fl.name, "content": text, **meta})
            st.success("File context successfully cached.")
            st.rerun()

    if st.session_state.candidate_jd_list:
        if st.button("Clear Saved Job Profiles"):
            st.session_state.candidate_jd_list = []
            st.rerun()
        for idx, item in enumerate(st.session_state.candidate_jd_list):
            st.markdown(f"**{idx+1}. {item['name']}** ({item['job_type']})")

def jd_batch_match_tab():
    st.header("🎯 Batch JD Match: Best Matches")
    is_resume_parsed = st.session_state.get('parsed') and st.session_state.parsed.get('error') is None

    if not is_resume_parsed:
        st.warning("Please upload and process structural candidate definitions before calculating metrics.")
        return
    if not st.session_state.get('candidate_jd_list'):
        st.error("Add target Job Descriptions before attempting systemic comparison scoring operations.")
        return

    if "candidate_match_results" not in st.session_state: st.session_state.candidate_match_results = []
    all_names = [i['name'] for i in st.session_state.candidate_jd_list]
    selections = st.multiselect("Target Matrices", options=all_names, default=all_names)

    if st.button("Execute Vectorized Semantic Score Maps"):
        st.session_state.candidate_match_results = []
        targets = [i for i in st.session_state.candidate_jd_list if i['name'] in selections]
        
        results_with_score = []
        for target in targets:
            analysis = evaluate_jd_fit(target['content'], st.session_state.parsed)
            score_match = re.search(r"Overall Fit Score:\s*(\d+)", analysis, re.IGNORECASE)
            score = score_match.group(1) if score_match else "6"
            
            s_p, x_p, e_p = "70", "60", "80"
            s_m = re.search(r'Skills Match:\s*(\d+)', analysis); x_m = re.search(r'Experience Match:\s*(\d+)', analysis)
            if s_m: s_p = s_m.group(1)
            if x_m: x_p = x_m.group(1)

            gaps_m = re.search(r'Gaps/Areas for Improvement:\s*(.*?)\s*(?:Overall Summary|$)', analysis, re.DOTALL | re.IGNORECASE)
            gaps_content = gaps_m.group(1).strip() if gaps_m else "Review core technical skills."

            results_with_score.append({
                "jd_name": target['name'], "overall_score": score, "numeric_score": int(score) if score.isdigit() else 6,
                "skills_percent": s_p, "experience_percent": x_p, "education_percent": e_p,
                "full_analysis": analysis, "gaps": gaps_content
            })
            
        results_with_score.sort(key=lambda x: x['numeric_score'], reverse=True)
        for rank, item in enumerate(results_with_score, 1): item['rank'] = rank
        st.session_state.candidate_match_results = results_with_score
        st.rerun()

    if st.session_state.candidate_match_results:
        df_view = []
        for r in st.session_state.candidate_match_results:
            df_view.append({"Rank": r['rank'], "Job Title": r['jd_name'], "Match Score": f"{r['overall_score']}/10", "Skills Match": f"{r['skills_percent']}%"})
        st.dataframe(pd.DataFrame(df_view), use_container_width=True, hide_index=True)

def parsed_data_tab():
    st.header("✨ Parsed Resume Data View")
    if st.session_state.get('parsed') and st.session_state.parsed.get('error') is None:
        st.json(st.session_state.parsed)
    else:
        st.info("No structured profile indices generated inside active working session vectors.")

def chatbot_tab_content():
    st.header("🤖 AI Chatbot Assistant")
    t_res, t_jd = st.tabs(["Resume Q&A Engine", "Job Profile Q&A Engine"])
    
    with t_res:
        if st.session_state.get('parsed'):
            q = st.text_input("Ask about profile qualifications:")
            if q and st.button("Query Profile Context"):
                st.write(qa_on_resume(q))
        else: st.warning("Parse profiles before querying state indexes.")
        
    with t_jd:
        if st.session_state.get('candidate_jd_list'):
            selected = st.selectbox("Job Index Target", options=[i['name'] for i in st.session_state.candidate_jd_list])
            target_content = next(i['content'] for i in st.session_state.candidate_jd_list if i['name'] == selected)
            q = st.text_input("Ask about structural criteria specs:")
            if q and st.button("Query Target Domain Specification"):
                st.write(qa_on_jd(q, target_content))
        else: st.warning("Upload target job specs to load context frames.")

def interview_preparation_tab():
    st.header("🎤 Interview Preparation Tools")
    t_res, t_jd = st.tabs(["Resume-Based Practices", "JD-Requirement Practices"])

    with t_res:
        if not st.session_state.get('parsed'): st.warning("Load profile vectors first."); return
        if st.button("Generate Behavioral & Skill Challenges"):
            resp = generate_interview_questions(st.session_state.parsed, 'resume', 'skills')
            st.session_state.interview_qa_resume = parse_questions_from_raw(resp)
        display_evaluation_form('resume', st.session_state.get('interview_qa_resume', []), st.session_state.full_text)

    with t_jd:
        if not st.session_state.get('candidate_jd_list'): st.warning("Cache requirement frameworks first."); return
        sel = st.selectbox("Practice Target Frame", options=[i['name'] for i in st.session_state.candidate_jd_list], key='interview_jd_sel')
        target_jd = next(i for i in st.session_state.candidate_jd_list if i['name'] == sel)
        if st.button("Generate Context Alignment Challenges"):
            resp = generate_interview_questions(target_jd['name'], 'jd', target_jd['content'])
            st.session_state.interview_qa_jd = parse_questions_from_raw(resp)
        display_evaluation_form('jd', st.session_state.get('interview_qa_jd', []), target_jd['content'])

def gap_analysis_tab():
    st.header("💡 Gap Analysis & Course Plan")
    if not st.session_state.get('candidate_match_results'):
        st.warning("Execute target Batch matching computations before analyzing curriculum deltas.")
        return

    top = st.session_state.candidate_match_results[0]
    st.subheader(f"Analyzing Top Structural Alignment Match: {top['jd_name']}")
    st.markdown(f"**Identified Competitive Skill Gaps:**\n{top['gaps']}")

    if st.button("Synthesize Optimization Study Roadmap"):
        skills = st.session_state.parsed.get('skills', [])
        plan = generate_gap_course_plan(top['gaps'], top['jd_name'], skills)
        st.session_state.gap_analysis_plan = plan
    
    if st.session_state.get('gap_analysis_plan'):
        st.markdown("---")
        st.markdown(st.session_state.gap_analysis_plan)


# --- ADAPTED HELPERS ---

def parse_questions_from_raw(raw_questions_response):
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
            q_list.append({"question": f"({level}/{q_type}) {question_text}", "answer": "", "level": level, "type": q_type})
    return q_list

def display_evaluation_form(mode, qa_data_list, context_for_eval):
    current_qa_key = f'interview_qa_{mode}'
    current_report_key = f'evaluation_report_{mode}'
    
    if qa_data_list:
        with st.form(f"practice_form_{mode}"):
            for i, item in enumerate(st.session_state[current_qa_key]):
                st.markdown(f"**Q{i+1}:** {item['question']}")
                st.session_state[current_qa_key][i]['answer'] = st.text_area(f"Response Matrix Space {i+1}", value=item['answer'], label_visibility='collapsed', key=f'ans_{mode}_{i}')
            if st.form_submit_button("Submit Assessment Frameworks"):
                report = evaluate_interview_answers(st.session_state[current_qa_key], context_for_eval)
                st.session_state[current_report_key] = report
                st.rerun()

        if st.session_state.get(current_report_key):
            st.markdown(st.session_state[current_report_key])

def qa_on_resume(question):
    prompt = f"Context Data: {st.session_state.full_text}\nAnswer question details accurately: {question}"
    try:
        if isinstance(client, MockGroqClient): return "Mock response validation checked."
        r = client.chat.completions.create(model=GROQ_MODEL, messages=[{"role":"user","content":prompt}], temperature=0.4)
        return r.choices[0].message.content.strip()
    except Exception as e: return str(e)

def qa_on_jd(question, jd_content):
    prompt = f"Specification Profiles: {jd_content}\nAnswer question details accurately: {question}"
    try:
        if isinstance(client, MockGroqClient): return "Mock specifications alignment validation passed."
        r = client.chat.completions.create(model=GROQ_MODEL, messages=[{"role":"user","content":prompt}], temperature=0.4)
        return r.choices[0].message.content.strip()
    except Exception as e: return str(e)


# --- APPLICATION DASHBOARD MAIN GATEWAY ENTRY ---

def candidate_dashboard():
    st.set_page_config(layout="wide", page_title="PragyanAI Candidate Dashboard")
    st.title("🧑‍💻 Candidate Dashboard")
    st.markdown("---")

    # App-wide structural pipeline state validations
    if "cv_data" not in st.session_state:
        st.session_state.cv_data = {'personal_info': {}, 'education': [], 'experience': [], 'projects': [], 'certifications': [], 'strengths_raw': ''}
    if "candidate_jd_list" not in st.session_state: st.session_state.candidate_jd_list = []
    if "candidate_match_results" not in st.session_state: st.session_state.candidate_match_results = []

    # Streamlined Operational Tabs (Cover Letter and Filter JD entirely removed)
    tabs = st.tabs([
        "📄 Resume Parsing", "📝 CV Management", "✨ Parsed Data View", 
        "📚 JD Management", "🎯 Batch JD Match", "🤖 AI Chatbot Assistant", 
        "🎤 Interview Preparation", "💡 Gap Analysis & Course Plan"
    ])
    
    with tabs[0]: resume_parsing_tab()
    with tabs[1]: cv_management_tab()
    with tabs[2]: parsed_data_tab()
    with tabs[3]: jd_management_tab_candidate()
    with tabs[4]: jd_batch_match_tab()
    with tabs[5]: chatbot_tab_content()
    with tabs[6]: interview_preparation_tab()
    with tabs[7]: gap_analysis_tab()

if __name__ == '__main__':
    candidate_dashboard()
