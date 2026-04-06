"""
DOCUMENT EXTRACTOR MODULE
Extracts text from PDF and image documents, then uses GPT to extract structured data.
Handles both text-based and scanned (image) PDFs using Vision API.
Also provides Vision API integration for damage image analysis.
"""

import os
import json
import base64
import io
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
    return text.strip()


def pdf_page_to_base64(file_path, page_num=0):
    """Convert a PDF page to a base64 image using pdfplumber."""
    try:
        with pdfplumber.open(file_path) as pdf:
            if page_num < len(pdf.pages):
                page = pdf.pages[page_num]
                img = page.to_image(resolution=200)
                buf = io.BytesIO()
                img.save(buf, format='PNG')
                buf.seek(0)
                return base64.b64encode(buf.read()).decode('utf-8')
    except Exception as e:
        print(f"Failed to convert PDF page to image: {e}")
    return None


def is_scanned_pdf(file_path):
    """Check if PDF is scanned (image-based with no extractable text)."""
    text = extract_text_from_pdf(file_path)
    return len(text.strip()) < 50


def extract_text_from_document(file_path):
    """Extract text from various document types"""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.pdf':
        return extract_text_from_pdf(file_path)
    elif ext in ['.jpg', '.jpeg', '.png']:
        return ""
    elif ext in ['.doc', '.docx']:
        return ""
    else:
        return ""


def extract_shipping_data_with_gpt(document_texts, document_images=None):
    """
    Use GPT-4o to extract ALL structured shipping/survey data.
    Supports both text and scanned document images via Vision API.
    """
    if not client:
        return {
            'error': 'OpenAI API not configured. Please set OPENAI_API_KEY environment variable.',
            'extracted': False
        }

    prompt = """You are a marine surveyor assistant. Extract ALL available information from the provided shipping and inspection documents (Bill of Lading, Commercial Invoice, Packing List, iAuditor/SafetyCulture reports).

Some documents are provided as text, others as images (scanned PDFs). Read ALL of them carefully.

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
transit_routing (IMPORTANT: extract the FULL multi-leg vessel routing if available - each leg as "M/V VESSEL VOY. XXX - FROM PORT, COUNTRY TO PORT, COUNTRY" separated by newlines)

CRITICAL FIELDS from Commercial Invoice / Packing List:
- shipper: Full company name and address of the shipper
- consignee: Full company name and address of the consignee
- gross_weight: Gross weight in KGS (look for GW, Gross Weight, Gross Kg)
- net_weight: Net weight in KGS (look for NW, Net Weight, Net Kg)
- number_of_packages: Total number and type of packages
- invoice_number: Commercial invoice number
- invoice_date: Date on the commercial invoice
- container_type: Container size and type (e.g., 40' HC, 20' GP)

"""

    # Build content array for GPT-4o Vision
    content_parts = [{"type": "text", "text": prompt}]

    # Add text-based documents
    if document_texts:
        combined_text = ""
        for doc_type, text in document_texts.items():
            combined_text += f"\n\n=== {doc_type.upper()} (TEXT) ===\n{text}"
        content_parts.append({
            "type": "text",
            "text": f"DOCUMENT TEXT:\n{combined_text[:12000]}"
        })

    # Add scanned document images
    if document_images:
        for doc_type, img_base64 in document_images.items():
            content_parts.append({
                "type": "text",
                "text": f"\n=== {doc_type.upper()} (SCANNED IMAGE) ==="
            })
            content_parts.append({
                "type": "image_url",
                "image_url": {
                    "url": f"data:image/png;base64,{img_base64}",
                    "detail": "high"
                }
            })

    content_parts.append({
        "type": "text",
        "text": "Return ONLY a valid flat JSON object with all fields above."
    })

    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are a precise data extraction assistant specializing in marine survey and shipping documents. Extract ALL available data from both text and images. Return only valid flat JSON."},
                {"role": "user", "content": content_parts}
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
    """Use OpenAI Vision API to analyze a damage image."""
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
        return {'success': True, 'description': description, 'category': category}

    except Exception as e:
        return {'error': f'Image analysis failed: {str(e)}', 'success': False}


def process_uploaded_documents(upload_folder, document_files):
    """
    Process all uploaded documents — extract text from text-based PDFs
    and convert scanned PDFs to images for Vision API processing.
    """
    document_texts = {}
    document_images = {}

    for doc_type, filename in document_files.items():
        if filename:
            file_path = os.path.join(upload_folder, f"{doc_type}_{filename}")
            if not os.path.exists(file_path):
                continue

            ext = os.path.splitext(file_path)[1].lower()

            if ext == '.pdf':
                text = extract_text_from_pdf(file_path)
                if len(text.strip()) > 50:
                    # Text-based PDF — use text
                    document_texts[doc_type] = text
                else:
                    # Scanned PDF — convert first page to image for Vision
                    print(f"  {doc_type}: Scanned PDF detected, using Vision API")
                    img_b64 = pdf_page_to_base64(file_path, page_num=0)
                    if img_b64:
                        document_images[doc_type] = img_b64
            elif ext in ['.jpg', '.jpeg', '.png']:
                # Direct image upload
                try:
                    with open(file_path, 'rb') as f:
                        img_b64 = base64.b64encode(f.read()).decode('utf-8')
                    document_images[doc_type] = img_b64
                except Exception:
                    pass

    if document_texts or document_images:
        return extract_shipping_data_with_gpt(document_texts, document_images)
    else:
        return {
            'error': 'No data could be extracted from uploaded documents',
            'extracted': False
        }
