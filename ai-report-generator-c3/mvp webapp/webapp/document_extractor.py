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


def _respace(text):
    """Re-insert spaces into iAuditor's despaced text for human display."""
    import re
    if not text:
        return text
    # Split before lowercase connector words glued onto preceding word
    for w in ['of', 'and', 'in', 'on', 'to', 'for', 'the', 'with']:
        text = re.sub(r'([A-Za-z])' + w + r'([A-Z])', r'\1 ' + w + r' \2', text)
    # Insert space between lowercase/digit followed by uppercase
    s = re.sub(r'([a-z0-9])([A-Z])', r'\1 \2', text)
    # Insert space between letter and digit
    s = re.sub(r'([A-Za-z])(\d)', r'\1 \2', s)
    s = re.sub(r'(\d)([A-Za-z])', r'\1 \2', s)
    return s.strip()


def _heading_to_category(heading):
    """Map an iAuditor heading string to a report category bucket.
    Accepts either spaced or despaced text — match against the despaced lower form.
    """
    if not heading:
        return None
    h = heading.replace(' ', '').lower()
    if 'mediasummary' in h:
        return 'skip'
    if 'descriptionofcargo' in h:
        return 'cargo_packaging'
    if 'conditionofcontainerexterior' in h or 'exteriorofcontainer' in h:
        return 'container_exterior'
    if 'conditionofcontainerinterior' in h or 'interiorofcontainer' in h:
        return 'container_interior'
    if 'conditionofcargo' in h or 'packageidentification' in h or \
       'natureofdamage' in h:
        return 'cargo_condition'
    if 'silvernitratetest' in h or 'testsconducted' in h or \
       'lighttest' in h or 'moisturetest' in h:
        return 'testing'
    return None


def _build_photo_section_map(pdf_path):
    """
    Walk the iAuditor PDF top-to-bottom and assign each "Photo N" label to
    the section it appears under, capturing surrounding answer text as a
    description.

    Returns:
        photo_sections: {photo_number(int): {category, heading, sub_heading, description, page}}
        page_photo_order: {page_index(int): [photo_number, ...]} in document order
        skip_pages: set of 0-based page indices to skip during image extraction
                    (Media Summary duplicates).
    """
    import re
    photo_sections = {}
    page_photo_order = {}
    skip_pages = set()

    current_category = None
    current_heading = ''      # top-level section name (e.g. Condition of Container Exterior)
    current_sub = ''          # sub-section name   (e.g. Left Sidewall)
    buffer_lines = []         # text lines under the current sub-section
    pending_photos = []       # photos found in current sub-section (assigned description on flush)

    photo_re = re.compile(r'Photo\s*(\d+)', re.IGNORECASE)
    numbered_heading_re = re.compile(r'^\s*\d+(?:\.\d+)*\.?\s*(.+)$')

    def flush_sub():
        """Assign accumulated description text to all pending photos in this sub-section."""
        if not pending_photos:
            return
        # Build a clean human-readable description from buffer
        cleaned = []
        for ln in buffer_lines:
            # Drop pure photo-label lines
            if re.fullmatch(r'(?:Photo\s*\d+\s*)+', ln.strip()):
                continue
            # Re-space and keep
            cleaned.append(_respace(ln.strip()))
        description = ' '.join(c for c in cleaned if c).strip()
        for pnum in pending_photos:
            # Don't overwrite an earlier (real) section assignment
            if pnum in photo_sections and photo_sections[pnum].get('category'):
                continue
            photo_sections[pnum] = {
                'category': current_category,
                'heading': _respace(current_heading),
                'sub_heading': _respace(current_sub),
                'description': description,
                'page': page_idx_holder[0] + 1,
            }

    page_idx_holder = [0]

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                page_idx_holder[0] = page_idx
                text = page.extract_text() or ''
                if not text:
                    continue

                # Detect & mark Media Summary pages — skip during image extraction
                compact_page = text.replace(' ', '').lower()
                if 'mediasummary' in compact_page:
                    skip_pages.add(page_idx)
                    # still walk lines so we don't pollute current section,
                    # but don't capture photos from here
                    continue

                page_photo_order[page_idx] = []

                for raw_line in text.split('\n'):
                    line = raw_line.strip()
                    if not line:
                        continue
                    if 'safetyculture.com' in line.lower():
                        continue
                    if re.match(r'^\d+/\d+$', line):  # page numbers like 5/20
                        continue
                    if line.lower().startswith('private'):
                        continue

                    compact = line.replace(' ', '')
                    cat = _heading_to_category(compact)

                    # Numbered top-level headings like "4.1.1.2.ConditionofContainerExterior"
                    m_num = numbered_heading_re.match(line)
                    is_numbered_heading = bool(m_num) and len(line) < 120 and not photo_re.search(line)

                    if cat is not None and is_numbered_heading:
                        # New top-level / category section — flush previous
                        flush_sub()
                        pending_photos = []
                        buffer_lines = []
                        current_category = cat
                        current_heading = m_num.group(1) if m_num else line
                        current_sub = current_heading
                        continue

                    if is_numbered_heading:
                        # Sub-section change but same category (e.g. LeftSidewall)
                        flush_sub()
                        pending_photos = []
                        buffer_lines = []
                        current_sub = m_num.group(1) if m_num else line
                        continue

                    # Collect any Photo N references on this line
                    photos_on_line = [int(x) for x in photo_re.findall(line)]
                    if photos_on_line:
                        for pnum in photos_on_line:
                            pending_photos.append(pnum)
                            page_photo_order[page_idx].append(pnum)
                        # Don't add the photo-label line to description buffer
                        continue

                    # Otherwise it's content text under current sub-section
                    buffer_lines.append(line)

            # Final flush at end of document
            flush_sub()
    except Exception as e:
        print(f"Failed to build photo/section map: {e}")

    return photo_sections, page_photo_order, skip_pages


def extract_iauditor_images(pdf_path, output_folder):
    """
    Extract photos from an iAuditor PDF and assign each to the heading
    immediately above its "Photo N" label, per the report structure.

    Returns a list of dicts:
        {filename, category, heading, photo_number, page, width, height}
    """
    os.makedirs(output_folder, exist_ok=True)
    extracted = []

    photo_sections, page_photo_order, skip_pages = _build_photo_section_map(pdf_path)

    try:
        reader = PdfReader(pdf_path)
        for i, page in enumerate(reader.pages):
            if i in skip_pages:
                continue
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
                    sub_heading = info.get('sub_heading', '')
                    description = info.get('description', '')
                else:
                    pnum = None
                    category = None
                    heading = ''
                    sub_heading = ''
                    description = ''

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

                filename = f"iauditor_{category}_photo{pnum or 'x'}_p{i+1}_{idx}.jpg"
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
                        'sub_heading': sub_heading,
                        'description': description,
                        'photo_number': pnum,
                        'page': i + 1,
                        'width': w,
                        'height': h,
                    })
                except Exception as e:
                    print(f"Failed to save image {name} from page {i+1}: {e}")
    except Exception as e:
        print(f"Failed to extract iAuditor images: {e}")

    # Drop content-duplicate files first (same image saved twice on a page)
    import hashlib
    seen_hashes = {}
    deduped = []
    for item in extracted:
        fp = os.path.join(output_folder, item['filename'])
        try:
            with open(fp, 'rb') as f:
                h = hashlib.md5(f.read()).hexdigest()
        except Exception:
            h = item['filename']
        if h in seen_hashes:
            try:
                os.remove(fp)
            except Exception:
                pass
            continue
        seen_hashes[h] = item['filename']
        deduped.append(item)
    extracted = deduped

    # Dedupe by photo_number — keep the largest (full-resolution) image
    by_num = {}
    for item in extracted:
        pnum = item.get('photo_number')
        if pnum is None:
            continue
        area = (item.get('width') or 0) * (item.get('height') or 0)
        existing = by_num.get(pnum)
        if existing is None or area > (existing.get('width') or 0) * (existing.get('height') or 0):
            by_num[pnum] = item
        else:
            # Remove the smaller duplicate file from disk
            try:
                os.remove(os.path.join(output_folder, item['filename']))
            except Exception:
                pass

    # Also remove disk files for the displaced "existing" entries when a larger replaced them
    final = list(by_num.values())
    final_files = {f['filename'] for f in final}
    for item in extracted:
        if item.get('filename') not in final_files:
            try:
                os.remove(os.path.join(output_folder, item['filename']))
            except Exception:
                pass

    final.sort(key=lambda x: (x.get('photo_number') or 9999, x.get('page', 0)))
    return final


def extract_iauditor_file_summary(pdf_path):
    """
    Extract the numbered "File summary" list from the end of an iAuditor PDF
    (the documents collected as enclosures). Returns a list of strings.
    """
    import re
    docs = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = ''
            for page in pdf.pages:
                t = page.extract_text() or ''
                full_text += '\n' + t
        # Locate the "File summary" block
        m = re.search(r'File\s*summary(.*)', full_text, re.IGNORECASE | re.DOTALL)
        if not m:
            return docs
        block = m.group(1)
        # Stop at the iAuditor footer or end of document
        block = re.split(r'safetyculture\.com|Private\s*&\s*confidential|\n\d+/\d+\b', block)[0]
        # Match "1.Filename" or "1. Filename" up to end of line
        for line in block.split('\n'):
            line = line.strip()
            if not line:
                continue
            mm = re.match(r'^(\d+)\s*[\.\)]\s*(.+)$', line)
            if mm:
                name = mm.group(2).strip()
                if name and len(name) < 200:
                    docs.append(name)
    except Exception as e:
        print(f"Failed to extract file summary: {e}")
    # Dedupe preserving order
    seen = set()
    result = []
    for d in docs:
        if d not in seen:
            seen.add(d)
            result.append(d)
    return result


def narrate_iauditor_sections(images):
    """
    Convert each iAuditor sub-section's captured Q&A-style text into a clean
    professional narrative paragraph using GPT-4o, suitable for an insurance
    report. Mutates the image dicts in place, replacing 'description'.
    """
    if not images or not client:
        return images

    # Group unique (category, sub_heading) -> raw_text
    groups = {}
    for img in images:
        key = (img.get('category') or '', img.get('sub_heading') or '')
        raw = (img.get('description') or '').strip()
        if not raw:
            continue
        if key not in groups or len(raw) > len(groups[key]):
            groups[key] = raw
    if not groups:
        return images

    # Build a single prompt asking GPT to narrate each section
    items = []
    for i, ((cat, sub), raw) in enumerate(groups.items()):
        items.append(f"[{i}] section=\"{sub}\" category=\"{cat}\"\nraw: {raw[:1500]}")
    listing = "\n\n".join(items)

    prompt = f"""You are a senior marine cargo surveyor writing a formal inspection report for an insurance claim.

For each numbered iAuditor section below, rewrite the raw checklist Q&A text into ONE clean, professional narrative paragraph (3-5 sentences) describing the observation in past tense, third-person, factual surveyor tone. Do NOT echo questions, do NOT use bullet points, do NOT include the words "Photo", "Yes/No", or numeric IDs from the form. Mention concrete observations: condition, damages, locations, measurements, packaging method, stowage, test results.

Return STRICT JSON: {{"0": "paragraph...", "1": "paragraph...", ...}} with one entry per item index.

Sections:
{listing}
"""

    try:
        resp = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are a precise marine surveyor who writes professional insurance-grade narratives."},
                {"role": "user", "content": prompt},
            ],
            temperature=0.2,
            max_tokens=2000,
        )
        content = resp.choices[0].message.content.strip()
        if "```json" in content:
            content = content.split("```json")[1].split("```")[0].strip()
        elif "```" in content:
            content = content.split("```")[1].split("```")[0].strip()
        narratives = json.loads(content)
    except Exception as e:
        print(f"narrate_iauditor_sections failed: {e}")
        return images

    # Map back: index -> key -> narration
    keys = list(groups.keys())
    key_to_text = {}
    for i, k in enumerate(keys):
        v = narratives.get(str(i)) or narratives.get(i)
        if v:
            key_to_text[k] = v.strip()

    for img in images:
        key = (img.get('category') or '', img.get('sub_heading') or '')
        if key in key_to_text:
            img['description'] = key_to_text[key]

    return images


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
