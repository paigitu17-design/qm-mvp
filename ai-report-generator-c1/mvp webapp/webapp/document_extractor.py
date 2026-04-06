"""
DOCUMENT EXTRACTOR MODULE
Extracts text from PDF and image documents, then uses GPT to extract structured data.
Also provides Vision API integration for damage image analysis.
"""

import os
import json
import base64
from PyPDF2 import PdfReader
import pdfplumber
from PIL import Image
from openai import OpenAI
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Initialize OpenAI client
client = None
try:
    api_key = os.getenv('OPENAI_API_KEY')
    if api_key:
        client = OpenAI(api_key=api_key)
except Exception as e:
    print(f"Warning: OpenAI client not initialized: {e}")


def extract_text_from_pdf(file_path):
    """Extract text from PDF file"""
    text = ""
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n\n"
    except Exception as e:
        print(f"pdfplumber failed, trying PyPDF2: {e}")
        try:
            reader = PdfReader(file_path)
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n\n"
        except Exception as e2:
            print(f"PyPDF2 also failed: {e2}")
            text = ""
    return text.strip()


def extract_text_from_image(file_path):
    """Extract text from image file (placeholder for OCR)"""
    try:
        Image.open(file_path)
        return f"[Image file: {os.path.basename(file_path)}]"
    except Exception as e:
        return f"[Unable to process image: {e}]"


def extract_text_from_document(file_path):
    """Extract text from various document types"""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.pdf':
        return extract_text_from_pdf(file_path)
    elif ext in ['.jpg', '.jpeg', '.png']:
        return extract_text_from_image(file_path)
    elif ext in ['.doc', '.docx']:
        return "[Word document - extraction not yet implemented]"
    else:
        return "[Unsupported file type]"


def extract_shipping_data_with_gpt(document_texts):
    """
    Use GPT to extract ALL structured shipping/survey data from documents.
    Extracts comprehensive fields needed for the ISA Certificate of Survey report.
    """
    if not client:
        return {
            'error': 'OpenAI API not configured. Please set OPENAI_API_KEY environment variable.',
            'extracted': False
        }

    combined_text = ""
    for doc_type, text in document_texts.items():
        combined_text += f"\n\n=== {doc_type.upper()} ===\n{text}"

    prompt = f"""You are a marine surveyor assistant. Extract ALL available information from the provided shipping and inspection documents (Bill of Lading, Commercial Invoice, Packing List, iAuditor/SafetyCulture reports).

IMPORTANT: Return a FLAT JSON object (no nesting). Each key must be a direct field name. Use "N/A" for fields not found.

Required fields (return ALL as top-level keys):

case_reference, principal_reference,
surveyor_in_charge, attending_surveyor, survey_date, survey_location, other_parties,
container_number, container_type, number_of_containers,
bl_number, bl_issue_place, bl_issue_date,
carrier_name, vessel_name, voyage_number,
origin_port, origin_country, discharge_port, destination_country,
transhipment_port, arrival_date, delivery_date, delivery_location,
gate_out_date, container_return_date, vessel_loading_date,
shipment_terms, incoterms, has_transhipment,
goods_description, number_of_packages, gross_weight, net_weight,
shipper, consignee,
container_condition_description, container_exterior_condition,
container_interior_condition, container_damages_found,
packaging_description, cargo_condition_description,
quantities_offloaded, quantities_inspected, damage_details,
silver_nitrate_test, light_test, other_tests,
cause_of_loss, cause_explanation, discussions, damage_discovery,
invoice_number, invoice_date, claim_currency, claim_amount,
notice_of_loss_date, consignee_contact, survey_discussion,
delivery_type, collection_date,
transit_routing (IMPORTANT: extract the FULL multi-leg vessel routing if available - each leg as "M/V VESSEL VOY. XXX - FROM PORT, COUNTRY TO PORT, COUNTRY" separated by newlines. Include all transhipment legs.)

Example format (FLAT, no grouping):
{{"case_reference": "QM5497-2024", "attending_surveyor": "John Doe", "container_number": "HASU4154103", "transit_routing": "M/V \\"POLAR PERU\\" VOY. 402N - Santiago, Chile to Balboa, Panama\\nM/V \\"LICA MAERSK\\" VOY. 405E - Manzanillo, Mexico to Tangier, Morocco", ...}}

DOCUMENT TEXT:
{combined_text[:12000]}

Return ONLY a valid flat JSON object.
"""

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are a precise data extraction assistant specializing in marine survey and shipping documents. Extract ALL available data and return only valid JSON. Pay special attention to iAuditor/SafetyCulture reports which contain surveyor details, container conditions, cargo conditions, and test results."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
            max_tokens=3000
        )

        content = response.choices[0].message.content.strip()

        try:
            data = json.loads(content)
            data['extracted'] = True
            return data
        except json.JSONDecodeError:
            if "```json" in content:
                json_str = content.split("```json")[1].split("```")[0].strip()
                data = json.loads(json_str)
                data['extracted'] = True
                return data
            elif "```" in content:
                json_str = content.split("```")[1].split("```")[0].strip()
                data = json.loads(json_str)
                data['extracted'] = True
                return data
            else:
                raise

    except Exception as e:
        return {
            'error': f'GPT extraction failed: {str(e)}',
            'extracted': False
        }


def analyze_damage_image(file_path, category="cargo"):
    """
    Use OpenAI Vision API to analyze a damage image and return a description.
    """
    if not client:
        return {
            'error': 'OpenAI API not configured. Please set OPENAI_API_KEY environment variable.',
            'success': False
        }

    try:
        with open(file_path, "rb") as image_file:
            base64_image = base64.b64encode(image_file.read()).decode("utf-8")

        ext = os.path.splitext(file_path)[1].lower()
        mime_type = "image/png" if ext == ".png" else "image/jpeg"

        category_context = {
            "cargo": "cargo/goods inside a shipping container",
            "container": "shipping container (exterior or interior structure)",
            "vessel": "vessel/ship"
        }
        context = category_context.get(category, "shipping/maritime equipment")

        prompt = f"""You are an expert marine surveyor inspecting {context}.
Analyze this image and provide a concise damage assessment (3-5 sentences max) suitable for an official inspection report.

Cover: what is shown, type of damage observed, severity (minor/moderate/severe), and likely cause.
Write in professional factual tone. If no damage is visible, state the {category} appears in sound condition."""

        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:{mime_type};base64,{base64_image}",
                                "detail": "high"
                            }
                        }
                    ]
                }
            ],
            max_tokens=300
        )

        description = response.choices[0].message.content.strip()

        return {
            'success': True,
            'description': description,
            'category': category
        }

    except Exception as e:
        return {
            'error': f'Image analysis failed: {str(e)}',
            'success': False
        }


def process_uploaded_documents(upload_folder, document_files):
    """Process all uploaded documents and extract data"""
    document_texts = {}

    for doc_type, filename in document_files.items():
        if filename:
            file_path = os.path.join(upload_folder, f"{doc_type}_{filename}")
            if os.path.exists(file_path):
                text = extract_text_from_document(file_path)
                if text and not text.startswith("["):
                    document_texts[doc_type] = text

    if document_texts:
        return extract_shipping_data_with_gpt(document_texts)
    else:
        return {
            'error': 'No text could be extracted from uploaded documents',
            'extracted': False
        }
