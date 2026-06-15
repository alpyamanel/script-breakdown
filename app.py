import os
import json
import io
import time
import pandas as pd
import streamlit as st
import pypdf
import docx
from pydantic import BaseModel, Field
from google import genai
from google.genai import types

# Initialize session state variables to manage workflow
if "step" not in st.session_state:
    st.session_state.step = "input"
if "raw_extraction" not in st.session_state:
    st.session_state.raw_extraction = None
if "confirmed_mappings" not in st.session_state:
    st.session_state.confirmed_mappings = {}
if "final_output" not in st.session_state:
    st.session_state.final_output = None

st.set_page_config(page_title="Secure Script Breakdown", layout="wide")
st.title("🎬 Secure Script Breakdown & Shotlist Generator")
st.write("De-duplicate elements and build a professional production breakdown powered securely by Google Gemini.")

# Set model
model_choice = "gemini-2.5-flash"

# Helper to fetch the secure client using Streamlit Secrets
def get_gemini_client():
    try:
        # Pull key automatically from Streamlit Secrets
        key = os.environ.get("GEMINI_API_KEY") or st.secrets.get("GEMINI_API_KEY")
        if not key:
            st.error("Missing Gemini API Key. Please add it to Streamlit Secrets.")
            return None
        return genai.Client(api_key=key)
    except Exception as e:
        st.error(f"Failed to initialize client: {e}")
        return None

# ----------------------------------------------------
# DOCUMENT PARSING UTILITIES
# ----------------------------------------------------
def extract_text_from_pdf(file_bytes):
    pdf_file = io.BytesIO(file_bytes)
    reader = pypdf.PdfReader(pdf_file)
    text = []
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text.append(page_text)
    return "\n".join(text)

def extract_text_from_docx(file_bytes):
    docx_file = io.BytesIO(file_bytes)
    doc = docx.Document(docx_file)
    text = []
    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            text.append(paragraph.text)
    return "\n".join(text)

# ----------------------------------------------------
# PYDANTIC STRUCTURED OUTPUT SCHEMAS
# ----------------------------------------------------
class DuplicateGroup(BaseModel):
    category: str = Field(description="Must be 'character', 'environment', or 'prop'")
    items: list[str] = Field(description="Items identified as duplicates of each other")
    suggested_canonical_name: str = Field(description="Recommended single name for these items")

class ExtractionResponse(BaseModel):
    all_characters: list[str] = Field(description="All extracted character names")
    all_environments: list[str] = Field(description="All extracted scene environments or locations")
    all_props: list[str] = Field(description="All extracted physical props or key objects")
    potential_duplicates: list[DuplicateGroup] = Field(description="Groupings of identified duplicates across characters, environments, or props")

class Shot(BaseModel):
    scene_num: int = Field(description="Scene index")
    shot_num: int = Field(description="Shot index")
    shot_type: str = Field(description="e.g., ECU, CU, MS, WS, OTS")
    camera_angle: str = Field(description="e.g., Low Angle, Eye Level, High Angle")
    action_description: str = Field(description="Brief description of camera framing and movement")
    elements_involved: list[str] = Field(description="Unified characters, props, or locations visible/acting in the shot")

class SummaryItem(BaseModel):
    name: str = Field(description="Name of the Character, Environment, or Prop")
    summary: str = Field(description="Descriptive context or profile within this scene")

class FinalBreakdownResponse(BaseModel):
    character_summaries: list[SummaryItem] = Field(description="Summary of each unique character")
    environment_summaries: list[SummaryItem] = Field(description="Summary of each unique environment")
    prop_summaries: list[SummaryItem] = Field(description="Summary of each unique prop")
    shotlist: list[Shot] = Field(description="Suggested scene-by-scene camera shotlist")

# ----------------------------------------------------
# EXCEL GENERATION UTILITY
# ----------------------------------------------------
def generate_multi_tab_excel(data):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # TAB 1: SHOT LIST
        shot_list = data.get("shotlist", [])
        df_shots = pd.DataFrame(shot_list)
        if not df_shots.empty:
            df_shots["elements_involved"] = df_shots["elements_involved"].apply(lambda x: ", ".join(x) if isinstance(x, list) else x)
            df_shots.columns = ["Scene #", "Shot #", "Shot Type", "Camera Angle", "Framing & Action Description", "Elements Involved"]
        else:
            df_shots = pd.DataFrame(columns=["Scene #", "Shot #", "Shot Type", "Camera Angle", "Framing & Action Description", "Elements Involved"])
        df_shots.to_excel(writer, sheet_name="Shot List", index=False)
        
        # TAB 2: BREAKDOWN ELEMENTS
        chars = data.get("character_summaries", [])
        envs = data.get("environment_summaries", [])
        props = data.get("prop_summaries", [])
        
        df_chars = pd.DataFrame(chars) if chars else pd.DataFrame(columns=["name", "summary"])
        df_chars.columns = ["Character Name", "Description & Context"]
        df_envs = pd.DataFrame(envs) if envs else pd.DataFrame(columns=["name", "summary"])
        df_envs.columns = ["Environment / Location", "Description"]
        df_props = pd.DataFrame(props) if props else pd.DataFrame(columns=["name", "summary"])
        df_props.columns = ["Prop Name", "Description & Context"]
        
        sheet_name = "Breakdown Elements"
        pd.DataFrame([["👤 CHARACTERS"]]).to_excel(writer, sheet_name=sheet_name, startrow=0, header=False, index=False)
        df_chars.to_excel(writer, sheet_name=sheet_name, startrow=1, index=False)
        
        start_env = len(df_chars) + 4
        pd.DataFrame([["📍 ENVIRONMENTS & LOCATIONS"]]).to_excel(writer, sheet_name=sheet_name, startrow=start_env, header=False, index=False)
        df_envs.to_excel(writer, sheet_name=sheet_name, startrow=start_env + 1, index=False)
        
        start_prop = start_env + len(df_envs) + 4
        pd.DataFrame([["🎒 PROPS & OBJECTS"]]).to_excel(writer, sheet_name=sheet_name, startrow=start_prop, header=False, index=False)
        df_props.to_excel(writer, sheet_name=sheet_name, startrow=start_prop + 1, index=False)

    output.seek(0)
    return output.getvalue()

# ----------------------------------------------------
# STEP 1: SCRIPT INPUT
# ----------------------------------------------------
if st.session_state.step == "input":
    st.header("Step 1: Upload or Paste Your Screenplay Script")
    
    # Toggle between Uploading a Document and Pasting Text
    input_method = st.radio("Choose Input Method:", ["📁 Upload Document (PDF, DOCX, TXT)", "✍️ Paste Script Text"])
    
    script_text = ""
    
    if input_method == "📁 Upload Document (PDF, DOCX, TXT)":
        uploaded_file = st.file_uploader(
            "Drag and drop your script file here", 
            type=["pdf", "docx", "txt"],
            help="Supports standard movie scripts in PDF, Microsoft Word, or plain text format."
        )
        if uploaded_file is not None:
            file_bytes = uploaded_file.read()
            if uploaded_file.name.endswith(".pdf"):
                script_text = extract_text_from_pdf(file_bytes)
            elif uploaded_file.name.endswith(".docx"):
                script_text = extract_text_from_docx(file_bytes)
            else:
                script_text = file_bytes.decode("utf-8")
            
            st.success(f"Successfully loaded '{uploaded_file.name}' ({len(script_text)} characters)")
    else:
        default_script = """SCENE 1 - INT. JOHN'S APARTMENT - DAY
JOHN (30s) paces around the messy kitchen. He grips a silver revolver.
His phone rings. The caller ID displays "OFFICER BOB". He hesitates, then answers.
JOHN
(whispering)
I told you not to call me here, Bob.

INTERCUT WITH:

SCENE 2 - INT. POLICE STATION - DAY
OFFICER BOB (50s) sits at a cluttered desk holding a phone and a mug of coffee.
BOB
You don't have a choice, Johnny. Put the gun down.
John looks at the gun in his hand."""
        script_text = st.text_area("Script Text:", value=default_script, height=350)

    if st.button("Analyze Script & Detect Duplicates", disabled=(not script_text)):
        client = get_gemini_client()
        if client:
            # --- LOADING PROGRESS BAR ENGINE ---
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Step 1: Processing text
            status_text.text("🔄 Step 1 of 4: Parsing text document...")
            progress_bar.progress(15)
            time.sleep(0.5)
            
            # Step 2: Preparing Payload
            status_text.text("⚙️ Step 2 of 4: Structuring extraction schema for Gemini...")
            progress_bar.progress(35)
            time.sleep(0.5)
            
            # Step 3: Running API
            status_text.text("🧠 Step 3 of 4: Running deep analysis with Gemini (extracting characters, locations, & props)...")
            progress_bar.progress(60)
            
            try:
                prompt = f"""
                Analyze the following film script. Extract:
                1. All Characters
                2. All Environments/Locations
                3. All Props
                
                Additionally, flag any potential duplicates/variations (e.g., "John" vs "Johnny", "revolver" vs "gun").
                
                Script:
                {script_text}
                """
                
                response = client.models.generate_content(
                    model=model_choice,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_schema=ExtractionResponse,
                        temperature=0.1
                    )
                )
                
                # Step 4: Complete
                progress_bar.progress(100)
                status_text.text("✅ Analysis complete! Loading duplicate resolver...")
                time.sleep(0.5)
                
                if response.parsed:
                    st.session_state.raw_extraction = response.parsed.model_dump()
                    st.session_state.script_text = script_text
                    st.session_state.step = "duplicate_check"
                    st.rerun()
                else:
                    st.error("No data parsed. Please check model and input.")
            except Exception as e:
                st.error(f"Analysis failed: {e}")

# ----------------------------------------------------
# STEP 2: INTERACTIVE DUPLICATE VERIFICATION
# ----------------------------------------------------
elif st.session_state.step == "duplicate_check":
    st.header("🔍 Step 2: Review and Resolve Duplicates")
    st.write("Select how to merge detected variations. Unselected or customized items will be treated as independent elements.")

    duplicates = st.session_state.raw_extraction.get("potential_duplicates", [])
    user_decisions = {}

    if not duplicates:
        st.success("No duplicates detected! Click proceed to build your breakdown.")
    else:
        for idx, dup in enumerate(duplicates):
            category = dup["category"].upper()
            items_str = ", ".join(f"'{i}'" for i in dup["items"])
            suggested = dup["suggested_canonical_name"]

            st.subheader(f"Group #{idx+1}: {category}")
            st.info(f"AI suggests matching: **{items_str}**")

            choice = st.radio(
                f"Action for Group #{idx+1}:",
                options=[
                    f"Merge all into '{suggested}'",
                    "Keep them separate",
                    "Merge into a custom name..."
                ],
                key=f"choice_{idx}"
            )

            custom_name = ""
            if choice == "Merge into a custom name...":
                custom_name = st.text_input("Custom consolidated name:", value=suggested, key=f"custom_{idx}")

            user_decisions[idx] = {
                "category": dup["category"],
                "items": dup["items"],
                "choice": choice,
                "custom_name": custom_name,
                "suggested": suggested
            }
            st.markdown("---")

    if st.button("Confirm Mappings & Generate Final Breakdown"):
        mappings = {}
        for idx, decision in user_decisions.items():
            if "Merge all into" in decision["choice"]:
                target_name = decision["suggested"]
                for item in decision["items"]:
                    mappings[item] = target_name
            elif decision["choice"] == "Merge into a custom name...":
                target_name = decision["custom_name"]
                for item in decision["items"]:
                    mappings[item] = target_name
            else:
                for item in decision["items"]:
                    mappings[item] = item

        st.session_state.confirmed_mappings = mappings
        
        client = get_gemini_client()
        if client:
            # --- DEPLOY LOADING BAR ---
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            status_text.text("🧱 Step 1 of 3: Mapping resolved duplicates...")
            progress_bar.progress(25)
            time.sleep(0.5)
            
            status_text.text("📝 Step 2 of 3: Generating unified breakdown details and scene-by-scene camera shotlist...")
            progress_bar.progress(60)
            
            try:
                final_prompt = f"""
                Produce a complete production breakdown and shotlist.
                
                CRITICAL REQUIREMENT: Wherever any duplicate refers to these mappings, you MUST map it to the unified name:
                {json.dumps(st.session_state.confirmed_mappings, indent=2)}
                
                Script to analyze:
                {st.session_state.script_text}
                """
                
                response = client.models.generate_content(
                    model=model_choice,
                    contents=final_prompt,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_schema=FinalBreakdownResponse,
                        temperature=0.2
                    )
                )
                
                progress_bar.progress(100)
                status_text.text("🎉 Formatting final production dashboard...")
                time.sleep(0.5)
                
                if response.parsed:
                    st.session_state.final_output = response.parsed.model_dump()
                    st.session_state.step = "final"
                    st.rerun()
                else:
                    st.error("Could not construct breakdown. Please try again.")
            except Exception as e:
                st.error(f"Failed generation step: {e}")

# ----------------------------------------------------
# STEP 3: EXPORT AND PRESENTATION
# ----------------------------------------------------
elif st.session_state.step == "final":
    st.header("🚀 Production Breakdown & Generated Shotlist")
    
    data = st.session_state.final_output
    excel_file = generate_multi_tab_excel(data)
    
    st.download_button(
        label="📥 Download Production Board (Excel .xlsx)",
        data=excel_file,
        file_name="production_breakdown.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch"
    )
    
    st.markdown("---")
    
    tab1, tab2, tab3, tab4 = st.tabs(["🎥 Shotlist", "👥 Characters", "📍 Locations", "🎒 Key Props"])
    
    with tab1:
        st.subheader("Shotlist Suggestion")
        shot_list = data.get("shotlist", [])
        if shot_list:
            df = pd.DataFrame(shot_list)
            df["elements_involved"] = df["elements_involved"].apply(lambda x: ", ".join(x) if isinstance(x, list) else x)
            df.columns = ["Scene #", "Shot #", "Shot Type", "Camera Angle", "Framing & Action Description", "Elements Involved"]
            st.dataframe(df, width="stretch")
        else:
            st.warning("No shotlist generated.")

    with tab2:
        st.subheader("Unified Character Summaries")
        for char in data.get("character_summaries", []):
            st.markdown(f"**👤 {char['name']}**")
            st.write(char['summary'])
            st.markdown("---")

    with tab3:
        st.subheader("Unified Environment Summaries")
        for env in data.get("environment_summaries", []):
            st.markdown(f"**📍 {env['name']}**")
            st.write(env['summary'])
            st.markdown("---")

    with tab4:
        st.subheader("Unified Prop Breakdown")
        for prop in data.get("prop_summaries", []):
            st.markdown(f"**🎒 {prop['name']}**")
            st.write(prop['summary'])
            st.markdown("---")

    if st.button("🔄 Analyze New Script"):
        st.session_state.step = "input"
        st.session_state.raw_extraction = None
        st.session_state.confirmed_mappings = {}
        st.session_state.final_output = None
        st.rerun()
