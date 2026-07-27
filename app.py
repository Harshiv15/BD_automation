# Streamlit Web App for Bulk BD Automation
# To run this locally:
# 1. Install dependencies: pip install streamlit openpyxl google-genai pydantic pdfplumber
# 2. Run command: streamlit run app.py

import streamlit as st
import os
import json
import tempfile
import pdfplumber
from pydantic import BaseModel, Field
from typing import List, Optional

from google import genai
from google.genai import types

# Import your working Excel scripts
try:
    from scripts import fill_factsheet
    from scripts import safe_insert_rows
except ImportError as e:
    st.error(f"Failed to import required scripts. Ensure 'fill_factsheet.py' and 'safe_insert_rows.py' are accessible. Error: {e}")
    st.stop()

st.set_page_config(page_title="EY TP BD Automator", layout="wide")

# =====================================================================
# 1. DEFINE THE STRICT JSON SCHEMA (PYDANTIC)
# =====================================================================

class FieldBase(BaseModel):
    value: str | int | float | None = Field(description="The extracted value. Use None if not found.")
    source: str = Field(description="Exact page number and section where this was found.")
    flag: Optional[str] = Field(None, description="Any judgment calls, assumptions, or notes for the human reviewer.")

class AERevenueSplitValue(BaseModel):
    domestic_ae_fy25: float | int | None
    export_ae_fy25: float | int | None
    domestic_third_party_fy25: float | int | None
    export_third_party_fy25: float | int | None

class AERevenueSplit(BaseModel):
    value: AERevenueSplitValue
    source: str
    flag: Optional[str] = None

class ShareholderRow(BaseModel):
    name: str = Field(description="Name of the shareholder/promoter category")
    shares_fy25: int | float | None
    shares_fy24: int | float | None

class Shareholding(BaseModel):
    rows: List[ShareholderRow]
    total_shares_fy25: int | float | None
    total_shares_fy24: int | float | None
    source: str

class RPTItem(BaseModel):
    label: str = Field(description="Exact transaction label from the AR")
    value_fy25: float | int | None = Field(description="Positive for outflow/cost, Negative for inflow/receivable")

class RelatedPartyTransactions(BaseModel):
    items: List[RPTItem]
    source: str

class LitigationItem(BaseModel):
    nature_of_dues: str
    amount_demanded_lakhs: float | int | None
    amount_paid_lakhs: float | int | None
    period: str
    forum: str

class Litigation(BaseModel):
    items: List[LitigationItem]
    source: str

class ExtractedFields(BaseModel):
    hq_india_entity: FieldBase
    hq_group: FieldBase
    company_description: FieldBase
    group_description: FieldBase
    standalone_turnover_fy25_lakhs: FieldBase
    standalone_turnover_fy24_lakhs: FieldBase
    standalone_total_cost_fy25_lakhs: FieldBase
    standalone_total_cost_fy24_lakhs: FieldBase
    consolidated_summary: FieldBase
    statutory_auditors: FieldBase
    ae_revenue_split: AERevenueSplit
    cash_fy25_lakhs: FieldBase
    cash_fy24_lakhs: FieldBase
    ae_trade_receivables_fy25_lakhs: FieldBase
    ae_trade_receivables_fy24_lakhs: FieldBase
    pe_investment: FieldBase
    shareholding: Shareholding
    related_party_transactions_lakhs: RelatedPartyTransactions
    countries_presence: FieldBase
    litigation: Litigation
    website: FieldBase
    linkedin: FieldBase

class CompanyData(BaseModel):
    entity: str = Field(description="Name of the company")
    fields: ExtractedFields

# =====================================================================
# 2. THE LLM EXTRACTION FUNCTION (TEXT-ONLY FAILSAFE)
# =====================================================================

def extract_bd_data(pdf_path: str, api_key: str, company_name: str) -> dict:
    """Extracts text locally, sends pure text to Gemini, and returns a validated dict."""
    
    # 1. Initialize Client
    client = genai.Client(api_key=api_key)
    
    # 2. Extract text locally (Bypasses Google File API to avoid Multimodal 429s)
    extracted_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            extracted_text += (page.extract_text() or "") + "\n"
    
    # 3. Read the SKILL.md prompt
    try:
        with open("SKILL.md", "r") as f:
            system_prompt = f.read()
    except FileNotFoundError:
        raise Exception("SKILL.md not found in the same directory.")

    # 4. Configure the API call (Search tool removed to avoid Free Tier 429s)
    config = types.GenerateContentConfig(
        system_instruction=system_prompt,
        temperature=0.0,
        response_mime_type="application/json",
        response_schema=CompanyData,
    )

    user_prompt = f"Please extract the Business Description data for {company_name} from the following Annual Report text:\n\n{extracted_text}"

    # 5. Call the Model
    # REPLACE THIS STRING WITH THE ONE YOU FOUND IN check_models.py
    response = client.models.generate_content(
        model='gemini-flash-lite-latest', 
        contents=user_prompt,
        config=config
    )

    return json.loads(response.text)

# =====================================================================
# 3. STREAMLIT UI & PROCESSING PIPELINE
# =====================================================================

st.title("📊 TP Business Description Automator")
st.markdown("Upload Annual Reports to extract financial data and populate the BD Excel format.")

with st.sidebar:
    st.header("Settings")
    api_key = st.text_input("Enter Gemini API Key", type="password")

col1, col2 = st.columns(2)
with col1:
    excel_template = st.file_uploader("Upload BD Format (.xlsx)", type=["xlsx"])
with col2:
    pdf_files = st.file_uploader("Upload ARs (.pdf)", type=["pdf"], accept_multiple_files=True)

if st.button("Start Bulk Extraction", type="primary"):
    if not api_key:
        st.error("Please enter a Gemini API key.")
    elif not excel_template or not pdf_files:
        st.error("Please upload the Excel template and at least one PDF.")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        with tempfile.TemporaryDirectory() as temp_dir:
            template_path = os.path.join(temp_dir, "template.xlsx")
            with open(template_path, "wb") as f:
                f.write(excel_template.getvalue())

            for i, pdf_file in enumerate(pdf_files):
                company_name = pdf_file.name.replace('.pdf', '')
                
                temp_pdf_path = os.path.join(temp_dir, f"{company_name}.pdf")
                with open(temp_pdf_path, "wb") as f:
                    f.write(pdf_file.getvalue())
                
                status_text.text(f"Processing {company_name}... Extracting text & Querying LLM...")
                try:
                    structured_data = extract_bd_data(temp_pdf_path, api_key, company_name)
                    
                    json_path = os.path.join(temp_dir, f"{company_name}_extracted.json")
                    with open(json_path, "w") as f:
                        json.dump(structured_data, f, indent=2)
                        
                except Exception as e:
                    st.error(f"Error during API extraction for {company_name}: {e}")
                    continue
                
                status_text.text(f"Processing {company_name}... Populating Excel...")
                output_excel_path = os.path.join(temp_dir, f"{company_name}_Filled_BD.xlsx")
                
                try:
                    template_to_use = template_path
                    
                    if len(structured_data["fields"]["shareholding"]["rows"]) > 2: 
                        status_text.text(f"Safely expanding Shareholding table...")
                        expanded_template_path = os.path.join(temp_dir, "expanded_template.xlsx")
                        
                        required_rows = len(structured_data["fields"]["shareholding"]["rows"])
                        delta = max(0, required_rows - 2) 
                        
                        safe_insert_rows.insert_rows_safe(
                            src_path=template_path, 
                            out_path=expanded_template_path, 
                            sheet_name="Copy of Thomas Cook", 
                            threshold=43, 
                            delta=delta
                        )
                        template_to_use = expanded_template_path

                    fill_factsheet.fill(
                        data_path=json_path, 
                        template_path=template_to_use, 
                        out_path=output_excel_path, 
                        sheet_name="Copy of Thomas Cook" 
                    )
                    
                except Exception as e:
                    st.error(f"Error populating Excel for {company_name}: {e}")
                    continue
                
                with open(output_excel_path, "rb") as f:
                    final_excel_bytes = f.read()
                    
                st.download_button(
                    label=f"Download Populated BD Format ({company_name})",
                    data=final_excel_bytes,
                    file_name=f"{company_name}_Populated_BD_Format.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_btn_{i}"
                )
                
                progress_bar.progress((i + 1) / len(pdf_files))
                
            status_text.text("✅ Processing Complete!")