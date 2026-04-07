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

CRITICAL FORMATTING RULES:
- ALL dates must be in "DD MONTH YYYY" format e.g. "08 APRIL 2024", "21 JANUARY 2024", "28 MARCH 2024". Never use ISO format.
- other_parties: Include FULL name AND designation for each person, e.g. "Mr Rajesh Gurung : Warehouse Supervisor, Mohebi Logistics; Mr Hari Krishna : Surveyor representing Amin Technical on behalf of Carrier Maersk"
- delivery_location: The physical address/location where goods were delivered to consignee. If not explicitly stated, use the survey location.

CRITICAL FIELD - packaging_description:
In the iAuditor document, look under section "4. Condition of container & goods" > "Description of Cargo". It contains these specific fields:
- "How the cargo was packed?" (e.g. Jumbo bag)
- "Nature of cargo stowage inside the container" (e.g. Slipsheet)
- "Quantities per stack" (e.g. 1)
- "Number of stacks per container" (e.g. 32)
- "Method of stowage in the container" (e.g. two high two across)
- "Was the stowage intact?" (e.g. Yes)
- "Was the cargo offloaded from the container prior to your attendance?" (e.g. No)
Compose a FULL professional paragraph using ALL of these values. Example: "Each unit of Full Cream Milk Powder was packed in jumbo bags placed on slipsheets. The bags were stowed two high, two across, with 1 bag per stack and 32 stacks per container. The stowage was found to be intact at the time of inspection. The cargo had not been offloaded prior to our attendance. No lashing straps or additional securing materials were observed."

CRITICAL FIELDS from Commercial Invoice / Packing List:
- shipper: Full company name and address of the shipper
- consignee: Full company name and address of the consignee
- gross_weight: Gross weight in KGS (look for GW, Gross Weight, Gross Kg)
- net_weight: Net weight in KGS (look for NW, Net Weight, Net Kg)
- number_of_packages: Total number and type of packages
- invoice_number: Commercial invoice number
- invoice_date: Date on the commercial invoice (in DD MONTH YYYY format)
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


def _heading_to_category(heading):
    """Map an iAuditor heading string to a report category bucket."""
    if not heading:
        return None
    h = heading.lower()
    # Skip media summary section (duplicate thumbnails)
    if 'media summary' in h:
        return 'skip'
    if 'description of cargo' in h:
        return 'cargo_packaging'
    if 'container exterior' in h or 'exterior of container' in h:
        return 'container_exterior'
    if 'container interior' in h or 'interior of container' in h:
        return 'container_interior'
    if 'condition of cargo' in h or 'package identification' in h or \
       'nature of damage' in h or 'damaged' in h or 'cargo condition' in h:
        return 'cargo_condition'
    if 'silver nitrate' in h or 'tests conducted' in h or 'moisture' in h or \
       'light test' in h or 'temperature' in h:
        return 'testing'
    return None


def _build_photo_section_map(pdf_path):
    """
    Walk the iAuditor PDF top-to-bottom and assign each "Photo N" label to
    the heading immediately above it.

    Returns:
        photo_sections: dict {photo_number(int): {'heading': str, 'category': str, 'page': int}}
        page_photo_order: dict {page_index(int): [photo_number, ...]}  (in vertical order)
    """
    import re
    photo_sections = {}
    page_photo_order = {}
    current_heading = None
    current_category = None

    heading_re = re.compile(r'^[0-9]+(\.[0-9]+)*\s*[\.\)]?\s*[A-Z]')
    photo_re = re.compile(r'^Photo\s*(\d+)\b', re.IGNORECASE)

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                # Extract words with positions, sort top-to-bottom, left-to-right
                try:
                    words = page.extract_words(use_text_flow=True, keep_blank_chars=False)
                except Exception:
                    words = []
                if not words:
                    continue

                # Group into lines by approximate y position
                lines = {}
                for w in words:
                    key = round(w['top'] / 4.0) * 4
                    lines.setdefault(key, []).append(w)
                sorted_keys = sorted(lines.keys())

                page_photo_order[page_idx] = []
                for k in sorted_keys:
                    line_words = sorted(lines[k], key=lambda x: x['x0'])
                    line_text = ' '.join(w['text'] for w in line_words).strip()
                    if not line_text:
                        continue

                    # Heading detection: numbered headings or known section words
                    cat = _heading_to_category(line_text)
                    if cat is not None:
                        current_heading = line_text
                        current_category = cat
                        continue
                    # Also recognize bare numbered headings like "4. Condition of cargo"
                    if heading_re.match(line_text) and len(line_text) < 120:
                        cat2 = _heading_to_category(line_text)
                        if cat2:
                            current_heading = line_text
                            current_category = cat2
                        continue

                    m = photo_re.match(line_text)
                    if m:
                        pnum = int(m.group(1))
                        photo_sections[pnum] = {
                            'heading': current_heading or '',
                            'category': current_category,
                            'page': page_idx + 1,
                        }
                        page_photo_order[page_idx].append(pnum)
    except Exception as e:
        print(f"Failed to build photo/section map: {e}")

    return photo_sections, page_photo_order


def extract_iauditor_images(pdf_path, output_folder):
    """
    Extract photos from an iAuditor PDF and assign each to the heading
    immediately above its "Photo N" label, per the report structure.

    Returns a list of dicts:
        {filename, category, heading, photo_number, page, width, height}
    """
    os.makedirs(output_folder, exist_ok=True)
    extracted = []

    photo_sections, page_photo_order = _build_photo_section_map(pdf_path)

    try:
        reader = PdfReader(pdf_path)
        for i, page in enumerate(reader.pages):
            if '/Resources' not in page or '/XObject' not in page['/Resources']:
                continue

            xobj = page['/Resources']['/XObject'].get_object()
            # Collect candidate images on this page (filter logos/icons)
            page_images = []
            for name in sorted(xobj.keys()):
                obj = xobj[name]
                if obj.get('/Subtype') != '/Image':
                    continue
                w = obj['/Width']
                h = obj['/Height']
                if w < 100 or h < 100:
                    continue
                page_images.append((name, obj, w, h))

            if not page_images:
                continue

            photo_nums_on_page = page_photo_order.get(i, [])

            for idx, (name, obj, w, h) in enumerate(page_images):
                # Match Nth image on page to Nth "Photo N" label on page
                if idx < len(photo_nums_on_page):
                    pnum = photo_nums_on_page[idx]
                    info = photo_sections.get(pnum, {})
                    category = info.get('category')
                    heading = info.get('heading', '')
                else:
                    pnum = None
                    category = None
                    heading = ''

                # Skip images we cannot place under any known section,
                # or that fall under media-summary duplicates.
                if not category or category == 'skip':
                    continue

                filter_type = obj.get('/Filter')
                try:
                    data = obj.get_data()
                except Exception as e:
                    print(f"Failed to read image data {name} on page {i+1}: {e}")
                    continue

                filename = f"iauditor_{category}_photo{pnum or f'p{i+1}_{idx}'}.jpg"
                filepath = os.path.join(output_folder, filename)

                try:
                    if filter_type == '/DCTDecode':
                        tmp_img = Image.open(io.BytesIO(data))
                        tmp_img.save(filepath, 'JPEG', quality=95)
                    else:
                        mode = 'RGB'
                        if obj.get('/ColorSpace') == '/DeviceGray':
                            mode = 'L'
                        tmp_img = Image.frombytes(mode, (w, h), data)
                        tmp_img.save(filepath, 'JPEG', quality=95)

                    extracted.append({
                        'filename': filename,
                        'category': category,
                        'heading': heading,
                        'photo_number': pnum,
                        'page': i + 1,
                        'width': w,
                        'height': h,
                    })
                except Exception as e:
                    print(f"Failed to save image {name} from page {i+1}: {e}")
    except Exception as e:
        print(f"Failed to extract iAuditor images: {e}")

    # Preserve order by photo_number when present, then by page
    extracted.sort(key=lambda x: (x.get('photo_number') or 9999, x.get('page', 0)))
    return extracted


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
