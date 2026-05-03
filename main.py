import base64
import datetime as dt
import json
import os
import time
import subprocess
import sys
import urllib.parse
from pathlib import Path


def ensure_import(module_name, package_name=None):
    try:
        return __import__(module_name, fromlist=["*"])
    except ModuleNotFoundError:
        package_to_install = package_name or module_name
        print(f"Installing missing package '{package_to_install}' for {sys.executable}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package_to_install])
        return __import__(module_name, fromlist=["*"])


requests = ensure_import("requests")
openai_module = ensure_import("openai")
docx_module = ensure_import("docx", "python-docx")
openpyxl_module = ensure_import("openpyxl")
httpx_module = ensure_import("httpx")
supabase_module = ensure_import("supabase")
OpenAI = openai_module.OpenAI
RateLimitError = openai_module.RateLimitError
Document = docx_module.Document
Workbook = openpyxl_module.Workbook
load_workbook = openpyxl_module.load_workbook
HTTPXClient = httpx_module.Client
create_client = supabase_module.create_client
Pt = docx_module.shared.Pt
RGBColor = docx_module.shared.RGBColor
WD_PARAGRAPH_ALIGNMENT = docx_module.enum.text.WD_PARAGRAPH_ALIGNMENT

try:
    import winreg
except ImportError:
    winreg = None


BASE_URL = "https://rest.arbeitsagentur.de/jobboerse/jobsuche-service/pc/v4"
HEADERS = {"X-API-Key": "jobboerse-jobsuche"}
REQUEST_TIMEOUT = 15
REQUEST_RETRIES = 1
REQUEST_RETRY_DELAY_SECONDS = 0
SEARCH_PAGE_SIZE = 20
SEARCH_MAX_PAGES = 2
MIN_SCORE_TO_PRINT = 2
MAX_COVER_LETTERS = 20
MIN_AI_MATCH_SCORE = 7
CV_PATH = Path("about-ai.txt")
TEMPLATE_PATH = Path("cover_letter_template.docx")
OPENAI_MODEL = "gpt-5.4-mini"
AI_SCORED_JOBS_PATH = Path("ai_scored_jobs.xlsx")
AI_WRITTEN_JOBS_PATH = Path("ai_cover_letters.xlsx")
LETTER_FONT_NAME = "Helvetica"
LETTER_FONT_SIZE = 9
LETTER_SPACE_AFTER_PT = 6
LETTER_TEXT_COLOR = RGBColor(29, 39, 49)
LETTER_LINE_SPACING = 1.5
SUPABASE_TABLE = "jobs"
DEFAULT_APPLICATION_STATUS = "not_applied"
SEARCH_TERMS_PATH = Path("search_terms.json")


def encode_refnr(refnr):
    return base64.b64encode(refnr.encode("utf-8")).decode("utf-8")


def build_job_page_url(refnr):
    encoded_refnr = urllib.parse.quote(refnr, safe="")
    return (
        "https://www.arbeitsagentur.de/jobsuche/jobdetail/"
        f"{encoded_refnr}"
    )


def discovery_timestamp(base_time, offset):
    return (base_time + dt.timedelta(microseconds=offset)).isoformat()


def get_json(url, *, params=None):
    last_exc = None

    for attempt in range(1, REQUEST_RETRIES + 1):
        try:
            res = requests.get(
                url,
                headers=HEADERS,
                params=params,
                timeout=REQUEST_TIMEOUT,
            )
            res.raise_for_status()
            return res.json()
        except requests.RequestException as exc:
            last_exc = exc
            if attempt == REQUEST_RETRIES:
                break
            if REQUEST_RETRY_DELAY_SECONDS > 0:
                time.sleep(REQUEST_RETRY_DELAY_SECONDS * attempt)

    raise last_exc


def search_jobs(term):
    jobs = []
    page = 1
    max_results = None

    while page <= SEARCH_MAX_PAGES:
        params = {
            "was": term,
            "size": SEARCH_PAGE_SIZE,
            "page": page,
        }
        data = get_json(f"{BASE_URL}/jobs", params=params)
        page_jobs = data.get("stellenangebote", [])
        if not page_jobs:
            break

        jobs.extend(page_jobs)
        max_results = data.get("maxErgebnisse")
        if max_results is not None and len(jobs) >= int(max_results):
            break
        if len(page_jobs) < SEARCH_PAGE_SIZE:
            break

        page += 1

    return jobs


def get_job_details(refnr):
    encoded = encode_refnr(refnr)
    return get_json(f"{BASE_URL}/jobdetails/{encoded}")


def score_job(job, description):
    keywords = [
        # Migration, integration, and refugee-support roles
        "Migration",
        "Integration",
        "Migrationsforschung",
        "Integrationsforschung",
        "Migrationsberatung",
        "Migrationsberater",
        "Migrationsberaterin",
        "MBE",
        "Zugewanderte",
        "Zuwanderung",
        "Zuwanderer",
        "Migrationshintergrund",
        "Geflüchtete",
        "Flüchtlinge",
        "Flüchtlingshilfe",
        "Flüchtlingsarbeit",
        "Asylsuchende",
        "Duldung",
        "Asyl",
        "Arbeitsmarktintegration",
        "Arbeitsmarktforschung",
        "internationale Arbeitsmarktforschung",
        "Migrantensozialarbeit",
        "interkulturell",
        "interkulturelle Kompetenz",
        "kultursensibel",
        "Integrationsmanagement",
        "Case Management",
        "Integrationskurs",
        "Orientierungskurs",
        "Erstorientierung",
        "Erstorientierungskurs",
        "Sprachkurse",
        "Anerkennungsverfahren",
        "Anerkennungsberatung",
        "Aufenthaltsrecht",
        "Sozialrecht",
        "Netzwerkarbeit",
        "Kooperationspartner",
        "Öffentlichkeitsarbeit",
        "Projektarbeit",
        "Einzelfallhilfe",
        "Begleitung",
        "Verweisberatung",
        "Clearing",
        "Berufliche Integration",
        "soziale Integration",
        "gesellschaftliche Integration",
        "Teilhabe",
        "Arbeitsmarkt",
        "Integration begleiten",

        # Research and social science roles aligned with your CV
        "empirische Sozialforschung",
        "Sozialforschung",
        "quantitative Sozialforschung",
        "qualitative Sozialforschung",
        "Sozialwissenschaftler",
        "Soziologe",
        "Soziologie",
        "Sozialwissenschaften",
        "Wissenschaftlicher Mitarbeiter",
        "wissenschaftliche Mitarbeit",
        "wissenschaftliche Hilfskraft",
        "Research Assistant",
        "Forschungsassistenz",
        "Forschungszentrum",
        "Datenerhebung",
        "Datenaufbereitung",
        "Datenanalyse",
        "Literaturrecherche",
        "Literaturüberblick",
        "Survey-Management",
        "Fragebogen",
        "Befragung",
        "Tabellen und Grafiken",
        "Auswertung",
        "Evaluierung",
        "Evaluation",
        "Politikberatung",
        "Interview",
        "Interviews",
        "Feldforschung",
        "Monitoring",
        "Mixed Methods",
        "MAXQDA",
        "SPSS",
        "Statistik",

        # Counseling, support, and community-facing terms
        "Sozialarbeiter",
        "Sozialpädagoge",
        "Sozialbetreuer",
        "Betreuungshelfer",
        "Beratung",
        "Migrationsberatung für Erwachsene",
        "Beratungsstelle für Zugewanderte",
        "Beratung für Geflüchtete",
        "psychosozial",
        "Konfliktbewältigung",
        "soziale Teilhabe",
        "Bildung und Arbeit",
        "Gemeinwesenarbeit",
        "Teilhabemanagement",

        # Language and cultural fit signals from current postings and your profile
        "Türkisch",
        "Farsi",
        "Dari",
        "Persisch",
        "Paschto",
        "Mehrsprachigkeit",
    ]

    score = 0
    text = (job.get("titel", "") + " " + (description or "")).lower()

    for kw in keywords:
        if kw.lower() in text:
            score += 1

    return score


def load_cv_text():
    return CV_PATH.read_text(encoding="utf-8")


def get_openai_api_key():
    api_key = os.environ.get("OPENAI_API_KEY")
    if api_key:
        return api_key

    if winreg is not None:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment") as key:
                saved_value, _ = winreg.QueryValueEx(key, "OPENAI_API_KEY")
                if saved_value:
                    os.environ["OPENAI_API_KEY"] = saved_value
                    return saved_value
        except OSError:
            pass

    env_file = Path(".env")
    if env_file.exists():
        for line in env_file.read_text(encoding="utf-8").splitlines():
            if line.startswith("OPENAI_API_KEY="):
                saved_value = line.split("=", 1)[1].strip().strip('"').strip("'")
                if saved_value:
                    os.environ["OPENAI_API_KEY"] = saved_value
                    return saved_value

    return None


def get_supabase_client():
    url = os.environ.get("SUPABASE_URL")
    key = os.environ.get("SUPABASE_ANON_KEY")
    if not url or not key:
        return None
    return create_client(url, key)


def upsert_job_in_supabase(client, payload):
    if client is None:
        return
    client.table(SUPABASE_TABLE).upsert(payload, on_conflict="refnr").execute()


def load_existing_jobs_from_supabase(client):
    if client is None:
        return {}

    rows = []
    page_size = 1000
    offset = 0
    while True:
        response = (
            client.table(SUPABASE_TABLE)
            .select(
                "refnr,job_id,date,keyword_score,ai_match_score,title,employer,city,reason,job_url,has_cover_letter,cover_letter_text,application_status,application_method,application_result,applied_at,note,created_at,updated_at,cover_letter_path,job_description_path,job_description_text,employer_street,employer_postal_code,employer_city"
            )
            .range(offset, offset + page_size - 1)
            .execute()
        )
        batch = response.data or []
        rows.extend(batch)
        if len(batch) < page_size:
            break
        offset += page_size

    existing = {}
    for row in rows:
        refnr = str(row.get("refnr") or "").strip()
        if not refnr:
            continue
        existing[refnr] = row
    return existing


def has_ai_match_score(row):
    value = row.get("ai_match_score") if row else None
    return value is not None and str(value).strip() != ""


def has_keyword_score(row):
    value = row.get("keyword_score") if row else None
    return value is not None and str(value).strip() != ""


def numeric_score(value):
    if value is None or str(value).strip() == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def has_generated_cover_letter(row):
    if not row:
        return False
    return bool(row.get("has_cover_letter") or row.get("cover_letter_text"))


def parse_job_id_number(job_id):
    raw = str(job_id or "").strip()
    if not raw or "-" not in raw:
        return None
    prefix = raw.split("-", 1)[0]
    try:
        return int(prefix)
    except ValueError:
        return None


def build_job_id(number, employer):
    cleaned_employer = " ".join(str(employer or "Unknown Employer").split())
    if not cleaned_employer:
        cleaned_employer = "Unknown Employer"
    return f"{number:04d}-{cleaned_employer}"


def format_log_table(rows):
    label_width = max(len(label) for label, _ in rows)
    value_width = max(len(str(value)) for _, value in rows)
    border = f"+-{'-' * label_width}-+-{'-' * value_width}-+"
    lines = [border]
    for label, value in rows:
        lines.append(f"| {label.ljust(label_width)} | {str(value).ljust(value_width)} |")
    lines.append(border)
    return "\n".join(lines)


def print_job_summary_table(job_id, employer, title, keyword_score, ai_match_score, search_term):
    ai_value = ai_match_score if ai_match_score not in (None, "") else "—"
    search_value = search_term or "—"
    print(
        format_log_table(
            [
                ("Job ID", job_id),
                ("Employer", employer or "Unknown employer"),
                ("Job Title", title or "Untitled"),
                ("Keyword Score", keyword_score),
                ("AI Matching Score", ai_value),
                ("Search Term", search_value),
            ]
        )
    )


def next_cover_letter_job_id_number(client):
    if client is None:
        return None

    response = client.table(SUPABASE_TABLE).select(
        "job_id"
    ).execute()
    rows = response.data or []
    used_numbers = [
        number
        for number in (parse_job_id_number(row.get("job_id")) for row in rows)
        if number is not None
    ]
    return (max(used_numbers) + 1) if used_numbers else 1


def is_job_id_prefix_in_use(client, number):
    if client is None:
        return False
    prefix = f"{number:04d}-"
    response = client.table(SUPABASE_TABLE).select("refnr").like("job_id", f"{prefix}%").limit(1).execute()
    rows = response.data or []
    return bool(rows)


def assign_job_id_if_missing(client, refnr, employer, current_job_id):
    existing_job_id = str(current_job_id or "").strip()
    if existing_job_id:
        return existing_job_id
    if client is None:
        return None

    candidate = next_cover_letter_job_id_number(client)
    if candidate is None:
        return None

    # Keep numeric prefixes globally unique (0001, 0002, ...) even if employer suffixes differ.
    while is_job_id_prefix_in_use(client, candidate):
        candidate += 1

    return build_job_id(candidate, employer)


def build_cover_letter_prompt(cv_text, job, description):
    location = job.get("arbeitsort", {}) or {}
    city = location.get("ort") or "Unknown city"
    employer = job.get("arbeitgeber") or "Unknown employer"
    title = job.get("titel") or "Unknown title"

    return f"""
Write a tailored professional cover letter in German for the following job application.

Candidate CV:
{cv_text}

Job information:
- Title: {title}
- Employer: {employer}
- City: {city}
- Description:
{description or "No description provided."}

Requirements:
- Use only facts supported by the CV.
- Tailor the letter to the role and employer.
- Emphasize relevant research, migration, sociology, multilingual, and fieldwork experience when appropriate.
- Keep the tone professional and natural.
- Keep the length around 300 to 450 words.
- Do not invent degrees, publications, software skills, or achievements.
- Write the letter with concrete details instead of generic claims.
- Write it for a German professional setting.
- Do not include a date line.
- Do not include a subject line.
""".strip()


def build_match_assessment_prompt(cv_text, job, description):
    location = job.get("arbeitsort", {}) or {}
    city = location.get("ort") or "Unknown city"
    employer = job.get("arbeitgeber") or "Unknown employer"
    title = job.get("titel") or "Unknown title"

    return f"""
Assess how well this candidate matches the job.

Candidate CV:
{cv_text}

Job information:
- Title: {title}
- Employer: {employer}
- City: {city}
- Description:
{description or "No description provided."}

Instructions:
- Score the match from 0 to 10.
- Base the score only on evidence in the CV and the job description.
- Be strict and realistic.
- If the job is in a very different field from the candidate profile, score it low.
- Respond in exactly this format:
MATCH_SCORE: <integer 0-10>
REASON: <one short paragraph>
""".strip()


def parse_match_score(text):
    for line in text.splitlines():
        if line.upper().startswith("MATCH_SCORE:"):
            raw_value = line.split(":", 1)[1].strip()
            try:
                return max(0, min(10, int(raw_value)))
            except ValueError:
                return None
    return None


def summarize_match_reason(ai_match_summary):
    for line in ai_match_summary.splitlines():
        if line.upper().startswith("REASON:"):
            return line.split(":", 1)[1].strip()
    return ai_match_summary.strip()


def assess_job_match(client, cv_text, job, description):
    response = client.responses.create(
        model=OPENAI_MODEL,
        input=build_match_assessment_prompt(cv_text, job, description),
    )
    output_text = response.output_text.strip()
    match_score = parse_match_score(output_text)
    if match_score is None:
        raise ValueError(f"Could not parse MATCH_SCORE from response: {output_text}")
    return match_score, output_text


def generate_cover_letter(client, cv_text, job, description):
    response = client.responses.create(
        model=OPENAI_MODEL,
        input=build_cover_letter_prompt(cv_text, job, description),
    )
    return response.output_text.strip()


def safe_slug(text):
    slug = "".join(ch.lower() if ch.isalnum() else "-" for ch in text)
    while "--" in slug:
        slug = slug.replace("--", "-")
    return slug.strip("-") or "job"


def german_date_string():
    months = [
        "Januar",
        "Februar",
        "März",
        "April",
        "Mai",
        "Juni",
        "Juli",
        "August",
        "September",
        "Oktober",
        "November",
        "Dezember",
    ]
    today = dt.date.today()
    return f"{today.day}. {months[today.month - 1]} {today.year}"


def style_paragraph(paragraph, *, space_after_pt=LETTER_SPACE_AFTER_PT):
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(space_after_pt)
    paragraph.paragraph_format.line_spacing = LETTER_LINE_SPACING
    for run in paragraph.runs:
        run.font.name = LETTER_FONT_NAME
        run.font.size = Pt(LETTER_FONT_SIZE)
        run.font.color.rgb = LETTER_TEXT_COLOR


def style_table_paragraph(paragraph):
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = LETTER_LINE_SPACING
    for run in paragraph.runs:
        run.font.name = LETTER_FONT_NAME
        run.font.size = Pt(LETTER_FONT_SIZE)
        run.font.color.rgb = LETTER_TEXT_COLOR


def replace_paragraph_text(paragraph, text, *, bold=None, space_after_pt=LETTER_SPACE_AFTER_PT):
    paragraph.text = ""
    run = paragraph.add_run(text)
    if bold is not None:
        run.bold = bold
    style_paragraph(paragraph, space_after_pt=space_after_pt)


def insert_paragraph_after(paragraph, text):
    new_paragraph = paragraph.insert_paragraph_before("")
    paragraph._p.addnext(new_paragraph._p)
    new_paragraph.style = paragraph.style
    if paragraph.alignment is not None:
        new_paragraph.alignment = paragraph.alignment
    new_paragraph.add_run(text)
    style_paragraph(new_paragraph)
    return new_paragraph


def extract_employer_address(job, details):
    employer = job.get("arbeitgeber") or ""
    address = {}
    locations = details.get("stellenlokationen") or []
    if locations:
        address = (locations[0] or {}).get("adresse") or {}

    street = " ".join(
        part for part in [address.get("strasse"), address.get("hausnummer")] if part
    ).strip()
    city_line = " ".join(part for part in [address.get("plz"), address.get("ort")] if part).strip()

    return {
        "employer": employer,
        "street_line": street,
        "postal_code": address.get("plz") or "",
        "address_city": address.get("ort") or "",
        "city_line": city_line,
    }


def load_search_terms():
    default_terms = [
        "Türkisch",
        "Migration",
        "Geflüchtete",
        "Sociology",
        "Integration",
        "Farsi",
        "Sozialbetreuer",
        "Persisch",
        "Soziologie",
        "Internationale Beziehungen",
        "Sozialwissenschaftler",
        "Flüchtlinge",
        "Mehrsprachigkeit",
        "Zuwanderer",
    ]
    raw_env = os.environ.get("SEARCH_TERMS", "").strip()
    if raw_env:
        terms = [term.strip() for term in raw_env.replace(";", "\n").replace(",", "\n").splitlines()]
        return [term for term in terms if term]
    if SEARCH_TERMS_PATH.exists():
        try:
            data = json.loads(SEARCH_TERMS_PATH.read_text(encoding="utf-8"))
            raw_terms = data.get("terms", data) if isinstance(data, dict) else data
            terms = [str(term).strip() for term in raw_terms]
            terms = [term for term in terms if term]
            if terms:
                return terms
        except (OSError, json.JSONDecodeError, TypeError):
            pass
    return default_terms


def fill_template_document(document, subject_text, body_text, employer_address):
    body_paragraphs = [part.strip() for part in body_text.split("\n\n") if part.strip()]

    for paragraph in document.paragraphs:
        if paragraph.text == "{{DATE}}":
            replace_paragraph_text(paragraph, german_date_string())
        elif paragraph.text == "{{SUBJECT}}":
            replace_paragraph_text(paragraph, subject_text, bold=True)
        elif paragraph.text == "{{BODY}}":
            replace_paragraph_text(paragraph, body_paragraphs[0] if body_paragraphs else "")
            current = paragraph
            for extra in body_paragraphs[1:]:
                current = insert_paragraph_after(current, extra)

    placeholder_map = {
        "{{EMPLOYER}}": employer_address.get("employer", ""),
        "{{STREET, HOUSE NUMBER}}": employer_address.get("street_line", ""),
        "{{POSTAL NUMBER, CITY}}": employer_address.get("city_line", ""),
    }

    for table in document.tables:
        for row in table.rows:
            row.height = None
            row.height_rule = None
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    text = paragraph.text.strip()
                    if text in placeholder_map:
                        replace_paragraph_text(
                            paragraph,
                            placeholder_map[text],
                            space_after_pt=0,
                        )
                        style_table_paragraph(paragraph)
                    else:
                        style_table_paragraph(paragraph)


def build_job_description_text(refnr, job, details, description, posted):
    location = job.get("arbeitsort", {}) or {}
    address = extract_employer_address(job, details)
    metadata_lines = [
        f"Referenznummer: {refnr}",
        f"Arbeitgeber: {job.get('arbeitgeber') or 'Unbekannt'}",
        f"Veröffentlicht: {posted or 'Unbekannt'}",
        f"Ort: {location.get('ort') or 'Unbekannt'}",
        f"PLZ: {location.get('plz') or 'Unbekannt'}",
        f"Job-URL: {build_job_page_url(refnr)}",
    ]

    if address.get("street_line"):
        metadata_lines.append(f"Adresse: {address['street_line']}")
    if address.get("city_line"):
        metadata_lines.append(f"Adresse 2: {address['city_line']}")

    lines = list(metadata_lines)
    body_paragraphs = [part.strip() for part in (description or "").split("\n\n") if part.strip()]
    if not body_paragraphs:
        body_paragraphs = ["Keine Stellenbeschreibung verfügbar."]

    lines.extend(body_paragraphs)
    return "\n\n".join(lines)


def ensure_workbook(path, headers):
    if path.exists():
        return

    workbook = Workbook()
    sheet = workbook.active
    sheet.append(headers)
    workbook.save(path)


def load_logged_refnrs(path):
    if not path.exists():
        return set()

    workbook = load_workbook(path)
    sheet = workbook.active
    refnrs = set()
    for row in sheet.iter_rows(min_row=2, values_only=True):
        refnr = row[0]
        if refnr:
            refnrs.add(str(refnr))
    return refnrs


def append_row_to_workbook(path, row):
    workbook = load_workbook(path)
    sheet = workbook.active
    sheet.append(row)
    workbook.save(path)


def main():
    if not get_openai_api_key():
        raise RuntimeError("OPENAI_API_KEY is not set.")

    client = OpenAI(http_client=HTTPXClient(trust_env=False))
    supabase = get_supabase_client()
    supabase_jobs = load_existing_jobs_from_supabase(supabase)
    cv_text = load_cv_text()

    ensure_workbook(
        AI_SCORED_JOBS_PATH,
        [
            "refnr",
            "date",
            "keyword_score",
            "ai_match_score",
            "title",
            "employer",
            "city",
            "reason",
        ],
    )
    ensure_workbook(
        AI_WRITTEN_JOBS_PATH,
        [
            "refnr",
            "date",
            "keyword_score",
            "ai_match_score",
            "title",
            "employer",
            "city",
            "job_url",
            "cover_letter_path",
        ],
    )

    scored_refnrs = load_logged_refnrs(AI_SCORED_JOBS_PATH)
    written_refnrs = load_logged_refnrs(AI_WRITTEN_JOBS_PATH)

    supabase_scored_refnrs = {
        refnr
        for refnr, row in supabase_jobs.items()
        if has_ai_match_score(row)
    }
    supabase_written_refnrs = {
        refnr
        for refnr, row in supabase_jobs.items()
        if has_generated_cover_letter(row)
    }

    scored_refnrs |= supabase_scored_refnrs
    written_refnrs |= supabase_written_refnrs

    search_terms = [
        "Türkisch",
        "Migration",
        "Geflüchtete",
        "Sociology",
        "Integration",
        "Farsi",
        "Sozialbetreuer",
        "Persisch",
        "Soziologie",
        "Internationale Beziehungen",
        "Sozialwissenschaftler",
        "Flüchtlinge",
        "Mehrsprachigkeit",
        "Zuwanderer",
    ]

    search_terms = load_search_terms()

    all_jobs = []
    job_terms = {}
    first_seen_at = {}
    discovery_base_time = dt.datetime.now(dt.timezone.utc)
    discovery_counter = 0

    for term in search_terms:
        try:
            jobs = search_jobs(term)
        except requests.RequestException as exc:
            print(f"Search failed for '{term}': {exc}")
            continue
        for job in jobs:
            refnr = job.get("refnr")
            if refnr:
                job_terms.setdefault(refnr, set()).add(term)
                existing_row = supabase_jobs.get(refnr) or {}
                if refnr not in first_seen_at and not existing_row.get("created_at"):
                    first_seen_at[refnr] = discovery_timestamp(discovery_base_time, discovery_counter)
                    discovery_counter += 1
        all_jobs.extend(jobs)

    unique_jobs = {}
    for job in all_jobs:
        refnr = job.get("refnr")
        if refnr:
            unique_jobs[refnr] = job

    jobs = list(unique_jobs.values())
    scored_jobs = []
    skipped_already_scored = 0
    skipped_existing_letters = 0

    for job in jobs:
        refnr = job.get("refnr")
        if not refnr:
            continue
        existing_supabase_job = dict(supabase_jobs.get(refnr) or {})
        if existing_supabase_job.get("created_at"):
            continue

        location = job.get("arbeitsort", {}) or {}
        created_at = first_seen_at.get(refnr) or dt.datetime.now(dt.timezone.utc).isoformat()
        upsert_job_in_supabase(
            supabase,
            {
                "refnr": refnr,
                "job_id": existing_supabase_job.get("job_id"),
                "date": job.get("aktuelleVeroeffentlichungsdatum"),
                "keyword_score": existing_supabase_job.get("keyword_score"),
                "ai_match_score": existing_supabase_job.get("ai_match_score"),
                "title": job.get("titel"),
                "employer": job.get("arbeitgeber"),
                "city": location.get("ort"),
                "reason": existing_supabase_job.get("reason"),
                "job_url": build_job_page_url(refnr),
                "cover_letter_path": existing_supabase_job.get("cover_letter_path"),
                "job_description_path": existing_supabase_job.get("job_description_path"),
                "cover_letter_text": existing_supabase_job.get("cover_letter_text"),
                "job_description_text": existing_supabase_job.get("job_description_text"),
                "has_cover_letter": bool(existing_supabase_job.get("has_cover_letter")),
                "application_status": existing_supabase_job.get("application_status") or DEFAULT_APPLICATION_STATUS,
                "application_result": existing_supabase_job.get("application_result") or "",
                "note": existing_supabase_job.get("note") or "",
                "created_at": created_at,
                "updated_at": existing_supabase_job.get("updated_at") or dt.datetime.now().isoformat(),
            },
        )
        existing_supabase_job.update(
            {
                "refnr": refnr,
                "job_id": existing_supabase_job.get("job_id"),
                "ai_match_score": existing_supabase_job.get("ai_match_score"),
                "has_cover_letter": bool(existing_supabase_job.get("has_cover_letter")),
                "cover_letter_text": existing_supabase_job.get("cover_letter_text"),
                "application_status": existing_supabase_job.get("application_status") or DEFAULT_APPLICATION_STATUS,
                "application_result": existing_supabase_job.get("application_result") or "",
                "note": existing_supabase_job.get("note") or "",
                "created_at": created_at,
                "updated_at": existing_supabase_job.get("updated_at") or dt.datetime.now().isoformat(),
                "reason": existing_supabase_job.get("reason"),
                "cover_letter_path": existing_supabase_job.get("cover_letter_path"),
                "job_description_path": existing_supabase_job.get("job_description_path"),
                "job_description_text": existing_supabase_job.get("job_description_text"),
            }
        )
        supabase_jobs[refnr] = existing_supabase_job

    for job in jobs:
        refnr = job.get("refnr")
        if not refnr:
            continue
        existing_supabase_job = supabase_jobs.get(refnr) or {}
        if has_generated_cover_letter(existing_supabase_job):
            skipped_existing_letters += 1
            continue

        existing_ai_score = numeric_score(existing_supabase_job.get("ai_match_score"))
        existing_keyword_score = numeric_score(existing_supabase_job.get("keyword_score"))
        if existing_ai_score is not None and existing_ai_score < MIN_AI_MATCH_SCORE:
            skipped_already_scored += 1
            continue
        if refnr in scored_refnrs and existing_ai_score is None:
            skipped_already_scored += 1
            continue

        description = existing_supabase_job.get("job_description_text") or ""
        details = {}
        if not description:
            try:
                details = get_job_details(refnr)
            except requests.RequestException as exc:
                print(f"Details failed for '{refnr}': {exc}")
                continue
            description = details.get("stellenangebotsBeschreibung")

        score = int(existing_keyword_score) if existing_keyword_score is not None else score_job(job, description)
        posted = job.get("aktuelleVeroeffentlichungsdatum")
        scored_jobs.append((posted, score, job, description, refnr, details))

    scored_jobs.sort(
        key=lambda x: (x[1], x[0] or ""),
        reverse=True,
    )

    number = 0
    generated_count = 0
    ai_available = True

    for posted, score, job, description, refnr, details in scored_jobs:
        if score < MIN_SCORE_TO_PRINT:
            continue

        number += 1
        existing_supabase_job = dict(supabase_jobs.get(refnr) or {})
        job_id = existing_supabase_job.get("job_id")
        created_at = (
            existing_supabase_job.get("created_at")
            or first_seen_at.get(refnr)
            or dt.datetime.now(dt.timezone.utc).isoformat()
        )

        location = job.get("arbeitsort", {}) or {}
        city = location.get("ort")
        postal = location.get("plz")
        matched_terms = sorted(job_terms.get(refnr, set()))
        search_term_text = ", ".join(matched_terms)
        current_application_status = (
            existing_supabase_job.get("application_status") or DEFAULT_APPLICATION_STATUS
        )
        current_application_method = existing_supabase_job.get("application_method") or ""
        current_application_result = existing_supabase_job.get("application_result") or ""
        current_applied_at = existing_supabase_job.get("applied_at")
        current_note = existing_supabase_job.get("note") or ""
        current_ai_match_score = existing_supabase_job.get("ai_match_score")
        existing_ai_match_score = numeric_score(current_ai_match_score)
        employer_address = extract_employer_address(job, details)
        if not employer_address.get("street_line"):
            employer_address["street_line"] = existing_supabase_job.get("employer_street") or ""
        if not employer_address.get("postal_code"):
            employer_address["postal_code"] = existing_supabase_job.get("employer_postal_code") or ""
        if not employer_address.get("address_city"):
            employer_address["address_city"] = existing_supabase_job.get("employer_city") or ""

        upsert_job_in_supabase(
            supabase,
            {
                "refnr": refnr,
                "job_id": job_id,
                "date": posted,
                "keyword_score": score,
                "ai_match_score": current_ai_match_score,
                "title": job.get("titel"),
                "employer": job.get("arbeitgeber"),
                "city": city,
                "employer_street": employer_address.get("street_line"),
                "employer_postal_code": employer_address.get("postal_code") or postal,
                "employer_city": employer_address.get("address_city") or city,
                "reason": existing_supabase_job.get("reason"),
                "job_url": build_job_page_url(refnr),
                "cover_letter_path": existing_supabase_job.get("cover_letter_path"),
                "job_description_path": existing_supabase_job.get("job_description_path"),
                "cover_letter_text": existing_supabase_job.get("cover_letter_text"),
                "job_description_text": description or existing_supabase_job.get("job_description_text") or "",
                "has_cover_letter": bool(existing_supabase_job.get("has_cover_letter")),
                "application_status": current_application_status,
                "application_method": current_application_method,
                "application_result": current_application_result,
                "applied_at": current_applied_at,
                "note": current_note,
                "created_at": created_at,
                "updated_at": dt.datetime.now().isoformat(),
            },
        )
        existing_supabase_job.update(
            {
                "refnr": refnr,
                "job_id": job_id,
                "ai_match_score": current_ai_match_score,
                "keyword_score": score,
                "has_cover_letter": bool(existing_supabase_job.get("has_cover_letter")),
                "cover_letter_text": existing_supabase_job.get("cover_letter_text"),
                "job_description_text": description or existing_supabase_job.get("job_description_text") or "",
                "application_status": current_application_status,
                "application_result": current_application_result,
                "note": current_note,
                "created_at": created_at,
            }
        )
        supabase_jobs[refnr] = existing_supabase_job

        print("-----------------------------------------------------------------------------\n")
        print(f"Job Refnr: {refnr}")
        print(f"Job ID: {job_id or '—'}")

        print(f"{city}, {postal}, {location}")
        print(f"no. {number} | date: {posted} | score: {score}")
        if matched_terms:
            print(f"Search term: {search_term_text}")
        print(job.get("titel"), "-", job.get("arbeitgeber"))
        print(description[:5000] if description else "No description")

        if generated_count >= MAX_COVER_LETTERS:
            print("Reached cover letter limit for this run.")
            print_job_summary_table(
                job_id or "—",
                job.get("arbeitgeber"),
                job.get("titel"),
                score,
                current_ai_match_score,
                search_term_text,
            )
            continue

        if not ai_available:
            print("AI cover letter generation is unavailable for this run.")
            print_job_summary_table(
                job_id or "—",
                job.get("arbeitgeber"),
                job.get("titel"),
                score,
                current_ai_match_score,
                search_term_text,
            )
            continue

        try:
            if existing_ai_match_score is not None:
                ai_match_score = existing_ai_match_score
                ai_match_summary = existing_supabase_job.get("reason") or (
                    f"MATCH_SCORE: {ai_match_score:g}/10\n"
                    "REASON: Existing AI match score reused from a previous run."
                )
                print("Reusing existing AI match score; generating missing cover letter.")
                print(ai_match_summary)
                current_ai_match_score = ai_match_score
            else:
                print("Assessing profile match...")
                ai_match_score, ai_match_summary = assess_job_match(
                    client,
                    cv_text,
                    job,
                    description,
                )
                print(ai_match_summary)

                append_row_to_workbook(
                    AI_SCORED_JOBS_PATH,
                    [
                        refnr,
                        posted,
                        score,
                        ai_match_score,
                        job.get("titel"),
                        job.get("arbeitgeber"),
                        city,
                        summarize_match_reason(ai_match_summary),
                    ],
                )
                upsert_job_in_supabase(
                    supabase,
                    {
                        "refnr": refnr,
                        "job_id": job_id,
                        "date": posted,
                        "keyword_score": score,
                        "ai_match_score": ai_match_score,
                        "title": job.get("titel"),
                        "employer": job.get("arbeitgeber"),
                        "city": city,
                        "employer_street": employer_address.get("street_line"),
                        "employer_postal_code": employer_address.get("postal_code") or postal,
                        "employer_city": employer_address.get("address_city") or city,
                        "reason": summarize_match_reason(ai_match_summary),
                        "job_url": build_job_page_url(refnr),
                        "cover_letter_path": None,
                        "job_description_path": None,
                        "cover_letter_text": None,
                        "job_description_text": description or "",
                        "has_cover_letter": False,
                        "application_status": current_application_status,
                        "application_method": current_application_method,
                        "application_result": current_application_result,
                        "applied_at": current_applied_at,
                        "note": current_note,
                        "created_at": created_at,
                        "updated_at": dt.datetime.now().isoformat(),
                    },
                )
                supabase_jobs[refnr] = {
                    "refnr": refnr,
                    "job_id": job_id,
                    "keyword_score": score,
                    "ai_match_score": ai_match_score,
                    "has_cover_letter": False,
                    "cover_letter_text": None,
                    "job_description_text": description or "",
                    "application_status": current_application_status,
                    "application_result": current_application_result,
                    "note": current_note,
                    "created_at": created_at,
                }
                scored_refnrs.add(refnr)
                current_ai_match_score = ai_match_score

            if ai_match_score < MIN_AI_MATCH_SCORE:
                print(
                    f"Skipping cover letter because AI match score is "
                    f"{ai_match_score}/10, below {MIN_AI_MATCH_SCORE}/10."
                )
                print_job_summary_table(
                    job_id or "—",
                    job.get("arbeitgeber"),
                    job.get("titel"),
                    score,
                    current_ai_match_score,
                    search_term_text,
                )
                continue

            print("Generating cover letter...")
            cover_letter = generate_cover_letter(client, cv_text, job, description)
        except RateLimitError as exc:
            print(f"OpenAI quota error: {exc}")
            print("Disabling AI cover letter generation for the rest of this run.")
            ai_available = False
            print_job_summary_table(
                job_id or "—",
                job.get("arbeitgeber"),
                job.get("titel"),
                score,
                current_ai_match_score,
                search_term_text,
            )
            continue
        except Exception as exc:
            print(f"Cover letter generation failed for '{refnr}': {exc}")
            print_job_summary_table(
                job_id or "—",
                job.get("arbeitgeber"),
                job.get("titel"),
                score,
                current_ai_match_score,
                search_term_text,
            )
            continue

        generated_count += 1

        if refnr not in written_refnrs:
            append_row_to_workbook(
                AI_WRITTEN_JOBS_PATH,
                [
                    refnr,
                    posted,
                    score,
                    ai_match_score,
                    job.get("titel"),
                    job.get("arbeitgeber"),
                    city,
                    build_job_page_url(refnr),
                    "",
                ],
            )
            written_refnrs.add(refnr)

        job_id = assign_job_id_if_missing(
            supabase,
            refnr,
            job.get("arbeitgeber"),
            job_id,
        )

        upsert_job_in_supabase(
            supabase,
            {
                "refnr": refnr,
                "job_id": job_id,
                "date": posted,
                "keyword_score": score,
                "ai_match_score": ai_match_score,
                "title": job.get("titel"),
                "employer": job.get("arbeitgeber"),
                "city": city,
                "employer_street": employer_address.get("street_line"),
                "employer_postal_code": employer_address.get("postal_code") or postal,
                "employer_city": employer_address.get("address_city") or city,
                "reason": summarize_match_reason(ai_match_summary),
                "job_url": build_job_page_url(refnr),
                "cover_letter_path": None,
                "job_description_path": None,
                "cover_letter_text": cover_letter,
                "job_description_text": description or "",
                "has_cover_letter": True,
                "application_status": current_application_status,
                "application_method": current_application_method,
                "application_result": current_application_result,
                "applied_at": current_applied_at,
                "note": current_note,
                "created_at": created_at,
                "updated_at": dt.datetime.now().isoformat(),
            },
        )
        supabase_jobs[refnr] = {
            "refnr": refnr,
            "job_id": job_id,
            "ai_match_score": ai_match_score,
            "has_cover_letter": True,
            "cover_letter_text": cover_letter,
            "application_status": current_application_status,
            "application_result": current_application_result,
            "note": current_note,
            "created_at": created_at,
        }

        print(cover_letter)
        print_job_summary_table(
            job_id or "—",
            job.get("arbeitgeber"),
            job.get("titel"),
            score,
            current_ai_match_score,
            search_term_text,
        )

    print("-----------------------------------------------------------------------------")
    print(f"Raw search hits collected: {len(all_jobs)}")
    print(f"Unique jobs after deduplication: {len(jobs)}")
    print(f"Jobs passing local score threshold: {number}")
    print(f"Jobs skipped because already detail-called/scored: {skipped_already_scored}")
    print(f"Jobs skipped because a cover letter already exists: {skipped_existing_letters}")
    print(f"New cover letters generated in this run: {generated_count}")


if __name__ == "__main__":
    main()
