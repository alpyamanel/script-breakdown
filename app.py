import os
import re
import json
import io
import time
from pathlib import Path
import pandas as pd
import streamlit as st
import pypdf
import docx
from pydantic import BaseModel, Field
from google import genai
from google.genai import types

# ----------------------------------------------------
# SESSION STATE
# ----------------------------------------------------
for k, v in {"step": "input", "raw_extraction": None,
             "confirmed_mappings": {}, "final_output": None}.items():
    if k not in st.session_state:
        st.session_state[k] = v

st.set_page_config(page_title="Secure Script Breakdown", layout="wide")
st.title("🎬 Secure Script Breakdown & Shotlist Generator")
st.write("De-duplicate elements, build a production breakdown, and match against the internal CCM asset library.")

model_choice = "gemini-2.5-flash"

# ----------------------------------------------------
# INTERNAL ASSET LIBRARY  (bundled app_data folder)
# ----------------------------------------------------
APP_DIR    = Path(__file__).parent
DATA_DIR   = APP_DIR / "app_data"
THUMBS_DIR = DATA_DIR / "thumbs"
INVENTORY_SHEET_URL = "https://docs.google.com/spreadsheets/d/14AgWDKuTXEepq2FfILmblmQlkuardDD1/edit"

CHAR_TYPES = {"char", "anml"}
ENV_TYPES  = {"envr", "env", "set"}
PROP_TYPES = {"prop", "food", "plnt", "plant", "veh", "vehicle", "fx"}
STOP_TOKENS = {"the", "a", "an", "of", "and", "with", "his", "her", "their",
               "default", "new", "old", "int", "ext", "day", "night", "scene"}


def _tokens(s):
    # split camelCase and non-alphanumerics into lowercase tokens
    s = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", str(s))
    parts = re.split(r"[^A-Za-z0-9]+", s)
    return [p.lower() for p in parts if len(p) >= 2 and p.lower() not in STOP_TOKENS]


@st.cache_data(show_spinner=False)
def load_inventory():
    f = DATA_DIR / "inventory.csv"
    if not f.exists():
        return None
    df = pd.read_csv(f, dtype=str).fillna("")
    # aggregate to one entry per (type, asset-name), collecting years + an image
    index = {}
    for _, r in df.iterrows():
        key = (r["type"].lower(), "".join(c for c in r["name"].lower() if c.isalnum()))
        e = index.setdefault(key, {"name": r["name"], "type": r["type"],
                                    "category": r["category"], "years": set(),
                                    "image": "", "tokens": set()})
        e["years"].add(r["year"])
        if not e["image"] and r["image"]:
            e["image"] = r["image"]
        e["tokens"].update(_tokens(r["name"]))
        e["tokens"].update(_tokens(r["category"]))
    return list(index.values())


def kind_types(kind):
    k = kind.lower()
    if k.startswith("char"):
        return CHAR_TYPES
    if k.startswith("env") or k.startswith("loc"):
        return ENV_TYPES
    return PROP_TYPES


def match_assets(term, kind, inventory, limit=10):
    types_ok = kind_types(kind)
    tterms = set(_tokens(term))
    if not tterms:
        return []
    results = []
    for e in inventory:
        if e["type"].lower() not in types_ok:
            continue
        cat_l = e["category"].lower()
        name_l = e["name"].lower()
        cat_toks = set(_tokens(e["category"]))
        score = 0
        for tt in tterms:
            if tt == cat_l:
                score += 120
            elif tt in cat_toks:
                score += 80
            elif tt in e["tokens"]:
                score += 50
            elif tt in name_l or tt in cat_l:
                score += 20
        score += 8 * len(tterms & e["tokens"])
        if score >= 50:
            results.append((score, e))
    results.sort(key=lambda x: (-x[0], x[1]["name"].lower()))
    return [e for _, e in results[:limit]]


# ----------------------------------------------------
# GEMINI CLIENT
# ----------------------------------------------------
def get_gemini_client():
    try:
        key = os.environ.get("GEMINI_API_KEY") or st.secrets.get("GEMINI_API_KEY")
        if not key:
            st.error("Missing Gemini API Key. Please add it to Streamlit Secrets.")
            return None
        return genai.Client(api_key=key)
    except Exception as e:
        st.error(f"Failed to initialize client: {e}")
        return None


# ----------------------------------------------------
# DOCUMENT PARSING
# ----------------------------------------------------
def extract_text_from_pdf(file_bytes):
    reader = pypdf.PdfReader(io.BytesIO(file_bytes))
    return "\n".join(p.extract_text() for p in reader.pages if p.extract_text())


def extract_text_from_docx(file_bytes):
    doc = docx.Document(io.BytesIO(file_bytes))
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())


# ----------------------------------------------------
# SCHEMAS
# ----------------------------------------------------
class DuplicateGroup(BaseModel):
    category: str = Field(description="Must be 'character', 'environment', or 'prop'")
    items: list[str] = Field(description="Items identified as duplicates of each other")
    suggested_canonical_name: str = Field(description="Recommended single name for these items")

class ExtractionResponse(BaseModel):
    all_characters: list[str]
    all_environments: list[str]
    all_props: list[str]
    potential_duplicates: list[DuplicateGroup]

class Shot(BaseModel):
    scene_num: int
    shot_num: int
    shot_type: str
    camera_angle: str
    action_description: str
    elements_involved: list[str]

class SummaryItem(BaseModel):
    name: str
    summary: str

class FinalBreakdownResponse(BaseModel):
    character_summaries: list[SummaryItem]
    environment_summaries: list[SummaryItem]
    prop_summaries: list[SummaryItem]
    shotlist: list[Shot]


# ----------------------------------------------------
# EXCEL EXPORT  (uses the EDITED shotlist + sourcing)
# ----------------------------------------------------
SHOT_COLS = ["Scene #", "Shot #", "Shot Type", "Camera Angle",
             "Framing & Action Description", "Elements Involved"]


def generate_excel(shot_df, data, sourcing_rows):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        (shot_df if not shot_df.empty else pd.DataFrame(columns=SHOT_COLS)) \
            .to_excel(writer, sheet_name="Shot List", index=False)

        sheet = "Breakdown Elements"
        chars = pd.DataFrame(data.get("character_summaries", []) or [], columns=["name", "summary"])
        envs  = pd.DataFrame(data.get("environment_summaries", []) or [], columns=["name", "summary"])
        props = pd.DataFrame(data.get("prop_summaries", []) or [], columns=["name", "summary"])
        chars.columns = ["Character Name", "Description & Context"]
        envs.columns  = ["Environment / Location", "Description"]
        props.columns = ["Prop Name", "Description & Context"]

        pd.DataFrame([["CHARACTERS"]]).to_excel(writer, sheet_name=sheet, startrow=0, header=False, index=False)
        chars.to_excel(writer, sheet_name=sheet, startrow=1, index=False)
        e0 = len(chars) + 4
        pd.DataFrame([["ENVIRONMENTS & LOCATIONS"]]).to_excel(writer, sheet_name=sheet, startrow=e0, header=False, index=False)
        envs.to_excel(writer, sheet_name=sheet, startrow=e0 + 1, index=False)
        p0 = e0 + len(envs) + 4
        pd.DataFrame([["PROPS & OBJECTS"]]).to_excel(writer, sheet_name=sheet, startrow=p0, header=False, index=False)
        props.to_excel(writer, sheet_name=sheet, startrow=p0 + 1, index=False)

        src = pd.DataFrame(sourcing_rows, columns=["Element", "Kind", "Status", "Suggested internal matches"]) \
            if sourcing_rows else pd.DataFrame(columns=["Element", "Kind", "Status", "Suggested internal matches"])
        src.to_excel(writer, sheet_name="Asset Sourcing", index=False)

    output.seek(0)
    return output.getvalue()


# ====================================================
# STEP 1: INPUT
# ====================================================
if st.session_state.step == "input":
    st.header("Step 1: Upload or Paste Your Screenplay Script")
    input_method = st.radio("Choose Input Method:", ["📁 Upload Document (PDF, DOCX, TXT)", "✍️ Paste Script Text"])
    script_text = ""

    if input_method.startswith("📁"):
        uploaded_file = st.file_uploader("Drag and drop your script file here",
                                         type=["pdf", "docx", "txt"])
        if uploaded_file is not None:
            fb = uploaded_file.read()
            if uploaded_file.name.endswith(".pdf"):
                script_text = extract_text_from_pdf(fb)
            elif uploaded_file.name.endswith(".docx"):
                script_text = extract_text_from_docx(fb)
            else:
                script_text = fb.decode("utf-8")
            st.success(f"Loaded '{uploaded_file.name}' ({len(script_text)} characters)")
    else:
        default_script = """SCENE 1 - INT. JJ'S HOUSE - DAY
JJ plays in the kitchen with a red cup and a balloon.
Cody runs in holding a toy bus."""
        script_text = st.text_area("Script Text:", value=default_script, height=300)

    if st.button("Analyze Script & Detect Duplicates", disabled=(not script_text)):
        client = get_gemini_client()
        if client:
            pb = st.progress(0); status = st.empty()
            status.text("🔄 Parsing text..."); pb.progress(20); time.sleep(0.3)
            status.text("🧠 Running deep analysis with Gemini..."); pb.progress(60)
            try:
                prompt = f"""
                Analyze the following film script. Extract:
                1. All Characters
                2. All Environments/Locations
                3. All Props
                Additionally, flag any potential duplicates/variations (e.g., "John" vs "Johnny").
                Script:
                {script_text}
                """
                response = client.models.generate_content(
                    model=model_choice, contents=prompt,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_schema=ExtractionResponse, temperature=0.1))
                pb.progress(100); status.text("✅ Analysis complete!"); time.sleep(0.3)
                if response.parsed:
                    st.session_state.raw_extraction = response.parsed.model_dump()
                    st.session_state.script_text = script_text
                    st.session_state.step = "duplicate_check"
                    st.rerun()
                else:
                    st.error("No data parsed. Please check model and input.")
            except Exception as e:
                st.error(f"Analysis failed: {e}")

# ====================================================
# STEP 2: DUPLICATE VERIFICATION
# ====================================================
elif st.session_state.step == "duplicate_check":
    st.header("🔍 Step 2: Review and Resolve Duplicates")
    duplicates = st.session_state.raw_extraction.get("potential_duplicates", [])
    user_decisions = {}

    if not duplicates:
        st.success("No duplicates detected! Click proceed to build your breakdown.")
    else:
        for idx, dup in enumerate(duplicates):
            st.subheader(f"Group #{idx+1}: {dup['category'].upper()}")
            st.info("AI suggests matching: **" + ", ".join(f"'{i}'" for i in dup["items"]) + "**")
            choice = st.radio(f"Action for Group #{idx+1}:",
                              [f"Merge all into '{dup['suggested_canonical_name']}'",
                               "Keep them separate", "Merge into a custom name..."],
                              key=f"choice_{idx}")
            custom_name = ""
            if choice == "Merge into a custom name...":
                custom_name = st.text_input("Custom consolidated name:",
                                            value=dup["suggested_canonical_name"], key=f"custom_{idx}")
            user_decisions[idx] = {"items": dup["items"], "choice": choice,
                                   "custom_name": custom_name,
                                   "suggested": dup["suggested_canonical_name"]}
            st.markdown("---")

    if st.button("Confirm Mappings & Generate Final Breakdown"):
        mappings = {}
        for idx, d in user_decisions.items():
            if "Merge all into" in d["choice"]:
                tgt = d["suggested"]
            elif d["choice"] == "Merge into a custom name...":
                tgt = d["custom_name"]
            else:
                tgt = None
            for item in d["items"]:
                mappings[item] = tgt if tgt else item
        st.session_state.confirmed_mappings = mappings

        client = get_gemini_client()
        if client:
            pb = st.progress(0); status = st.empty()
            status.text("📝 Generating unified breakdown + shotlist..."); pb.progress(50)
            try:
                final_prompt = f"""
                Produce a complete production breakdown and shotlist.
                CRITICAL: map any duplicate to the unified name per this table:
                {json.dumps(st.session_state.confirmed_mappings, indent=2)}
                Script to analyze:
                {st.session_state.script_text}
                """
                response = client.models.generate_content(
                    model=model_choice, contents=final_prompt,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_schema=FinalBreakdownResponse, temperature=0.2))
                pb.progress(100); status.text("🎉 Done!"); time.sleep(0.3)
                if response.parsed:
                    st.session_state.final_output = response.parsed.model_dump()
                    st.session_state.step = "final"
                    st.rerun()
                else:
                    st.error("Could not construct breakdown. Please try again.")
            except Exception as e:
                st.error(f"Failed generation step: {e}")

# ====================================================
# STEP 3: RESULTS
# ====================================================
elif st.session_state.step == "final":
    st.header("🚀 Production Breakdown & Shotlist")
    data = st.session_state.final_output
    inventory = load_inventory()

    # ---------- EDITABLE SHOTLIST ----------
    st.subheader("🎥 Shotlist (editable)")
    st.caption("Edit any cell, add or delete rows — the Excel export and table below update automatically.")
    shot_list = data.get("shotlist", [])
    if shot_list:
        df = pd.DataFrame(shot_list)
        df["elements_involved"] = df["elements_involved"].apply(
            lambda x: ", ".join(x) if isinstance(x, list) else x)
        df.columns = SHOT_COLS
    else:
        df = pd.DataFrame(columns=SHOT_COLS)

    edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True,
                               key="shot_editor")

    # ---------- INTERNAL ASSET RECOMMENDATIONS ----------
    st.markdown("---")
    st.subheader("🗂️ Internal Asset Recommendations")
    if inventory is None:
        st.warning("Asset library not found. Make sure the **app_data** folder "
                   "(inventory.csv + thumbs) is committed alongside the app.")
        sourcing_rows = []
    else:
        st.caption(f"Matched against {len(inventory)} internal CCM assets · "
                   f"[source inventory]({INVENTORY_SHEET_URL})")
        groups = [("Characters", "character", data.get("character_summaries", [])),
                  ("Environments", "environment", data.get("environment_summaries", [])),
                  ("Props", "prop", data.get("prop_summaries", []))]
        sourcing_rows = []
        for title, kind, items in groups:
            if not items:
                continue
            st.markdown(f"### {title}")
            for it in items:
                elem = it["name"]
                matches = match_assets(elem, kind, inventory)
                if matches:
                    st.markdown(f"**{elem}**  —  ✅ {len(matches)} internal option(s)")
                    cols = st.columns(5)
                    for i, m in enumerate(matches):
                        with cols[i % 5]:
                            img_path = THUMBS_DIR / m["image"] if m["image"] else None
                            if img_path and img_path.exists():
                                st.image(str(img_path), use_container_width=True)
                            else:
                                st.markdown("`no image`")
                            yrs = ", ".join(sorted(m["years"]))
                            st.caption(f"{m['name']}\n\n{m['type']} · {yrs}")
                    sourcing_rows.append([elem, title.rstrip("s"), "HAVE (internal)",
                                          ", ".join(m["name"] for m in matches[:5])])
                else:
                    st.markdown(f"**{elem}**  —  ⚠️ not in library · **request from vendor**")
                    sourcing_rows.append([elem, title.rstrip("s"), "REQUEST (vendor)", ""])
                st.markdown("")

        have = sum(1 for r in sourcing_rows if r[2].startswith("HAVE"))
        need = len(sourcing_rows) - have
        st.info(f"**{have}** elements have internal options · **{need}** to request from the vendor")

    # ---------- EXPORT (built from edited shotlist) ----------
    st.markdown("---")
    excel_file = generate_excel(edited_df, data, sourcing_rows)
    st.download_button("📥 Download Production Board (Excel .xlsx)", data=excel_file,
                       file_name="production_breakdown.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)

    # ---------- SUMMARY TABS ----------
    st.markdown("---")
    t1, t2, t3 = st.tabs(["👥 Characters", "📍 Locations", "🎒 Key Props"])
    with t1:
        for c in data.get("character_summaries", []):
            st.markdown(f"**👤 {c['name']}**"); st.write(c["summary"]); st.markdown("---")
    with t2:
        for e in data.get("environment_summaries", []):
            st.markdown(f"**📍 {e['name']}**"); st.write(e["summary"]); st.markdown("---")
    with t3:
        for p in data.get("prop_summaries", []):
            st.markdown(f"**🎒 {p['name']}**"); st.write(p["summary"]); st.markdown("---")

    if st.button("🔄 Analyze New Script"):
        st.session_state.step = "input"
        st.session_state.raw_extraction = None
        st.session_state.confirmed_mappings = {}
        st.session_state.final_output = None
        st.rerun()
