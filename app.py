import streamlit as st
import json
import os
import io
import re
import unicodedata
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
os.environ["STREAMLIT_SERVER_RUN_ON_SAVE"] = "false"

from datetime import datetime
from zoneinfo import ZoneInfo
import pandas as pd

VILNIUS_TZ = ZoneInfo("Europe/Vilnius")


def now_vilnius():
    return datetime.now(VILNIUS_TZ)


from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import cm

import gspread
from google.oauth2.service_account import Credentials

# =========================
# CONFIG — pakeisk čia, kad pritaikytum šį testą
# =========================
TEST_TITLE = "SCL-90"                  # rodoma antraštė / PDF pavadinimas / el. laiško tema
QUESTIONS_FILENAME = "klausimai.txt"   # klausimų failo pavadinimas repo šaknyje
KEYS_FILENAME = "raktai.json"          # raktų failo pavadinimas repo šaknyje
INSTRUCTIONS_FILENAME = "instrukcijos.txt"  # tekstas, rodomas prieš pradedant testą (neprivalomas)
WORKSHEET_NAME = "Progresas"           # Google Sheets lapo (tab) pavadinimas šiam testui

# Atsakymų skalė (naudojama kaip slankiklis testo puslapyje)
# Sąrašo indeksas (0-4) yra balas, naudojamas calculate_scores() funkcijoje.
SCALE_LABELS = ["Visiškai ne", "Šiek tiek", "Vidutiniškai", "Gana daug", "Labai stipriai"]

QUESTION_HEADER = "Kiek jus vargino:"  # antraštė, rodoma virš kiekvieno klausimo

# Jei kada nors abu testai naudotų TĄ PATĮ Google Sheet dokumentą, čia reikėtų
# skirtingo WORKSHEET_NAME kiekvienam testui. Kadangi šiam testui kuriamas
# atskiras Sheet (kitas spreadsheet_id secrets faile), pakanka to paties
# pavadinimo "Progresas" — jie negalės susimaišyti, nes yra skirtinguose failuose.

st.set_page_config(layout="wide")
st.markdown("""
<style>
html, body, .stApp {
    background-color: #14171c !important;
    color: #f2f2f2 !important;
}
#MainMenu {visibility: hidden;}
header {visibility: hidden;}
footer {visibility: hidden;}
.block-container {
    padding: 2rem;
    max-width: 100%;
}
html, body { font-size: 34px; }
h1 { font-size: 80px !important; color: #f2f2f2 !important; }
.question-header {
    font-size: 32px !important;
    font-weight: 400;
    color: #b8b8b8;
    margin-bottom: 10px;
}
.question-text {
    font-size: 50px !important;
    font-weight: 700;
    margin-bottom: 25px;
    color: #ffffff !important;
}
div[role="radiogroup"] label {
    font-size: 50px !important;
    color: #f2f2f2 !important;
}
input[type="radio"] {
    transform: scale(3);
    margin-right: 15px;
}
div.stButton > button {
    font-size: 28px !important;
    height: 65px !important;
    width: 100%;
}
</style>
""", unsafe_allow_html=True)

# =========================
# PATHS
# =========================
APP_DIR = os.path.dirname(os.path.abspath(__file__))


def find_data_file(filename):
    candidates = [
        os.path.join(APP_DIR, filename),
        os.path.join(APP_DIR, "Duomenys", filename),
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    raise FileNotFoundError(
        f"Nerastas failas '{filename}'. Patikrinkite, ar jis įkeltas šalia app.py "
        f"(arba į Duomenys/ aplanką). Ieškota: {candidates}"
    )


QUESTIONS_PATH = find_data_file(QUESTIONS_FILENAME)
KEYS_PATH = find_data_file(KEYS_FILENAME)

# =========================
# UNICODE ŠRIFTAS PDF ATASKAITAI (LIETUVIŠKOS RAIDĖS)
# =========================
_DEJAVU_CANDIDATES = [
    os.path.join(APP_DIR, "fonts", "DejaVuSans.ttf"),
    os.path.join(APP_DIR, "DejaVuSans.ttf"),
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    "/usr/local/share/fonts/DejaVuSans.ttf",
]
_DEJAVU_BOLD_CANDIDATES = [
    os.path.join(APP_DIR, "fonts", "DejaVuSans-Bold.ttf"),
    os.path.join(APP_DIR, "DejaVuSans-Bold.ttf"),
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf",
    "/usr/local/share/fonts/DejaVuSans-Bold.ttf",
]


def register_unicode_font():
    regular = next((p for p in _DEJAVU_CANDIDATES if os.path.exists(p)), None)
    bold = next((p for p in _DEJAVU_BOLD_CANDIDATES if os.path.exists(p)), None)
    if not regular:
        return "Helvetica", "Helvetica-Bold", None
    pdfmetrics.registerFont(TTFont("DejaVuSans", regular))
    if bold:
        pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", bold))
    else:
        pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", regular))
    pdfmetrics.registerFontFamily(
        "DejaVuSans",
        normal="DejaVuSans",
        bold="DejaVuSans-Bold",
        italic="DejaVuSans",
        boldItalic="DejaVuSans-Bold",
    )
    return "DejaVuSans", "DejaVuSans-Bold", regular


PDF_FONT, PDF_FONT_BOLD, PDF_FONT_PATH_USED = register_unicode_font()


@st.cache_data
def load_questions():
    with open(QUESTIONS_PATH, "r", encoding="utf-8") as f:
        return [q.strip() for q in f if q.strip()]


@st.cache_data
def load_keys():
    with open(KEYS_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


QUESTIONS = load_questions()
KEYS = load_keys()

try:
    INSTRUCTIONS_PATH = find_data_file(INSTRUCTIONS_FILENAME)
    with open(INSTRUCTIONS_PATH, "r", encoding="utf-8") as f:
        INSTRUCTIONS_TEXT = f.read()
except FileNotFoundError:
    INSTRUCTIONS_TEXT = None

# =========================
# GOOGLE SHEETS (PROGRESO SAUGOJIMAS)
# =========================
SHEET_HEADERS = ["test_id", "vardas", "pavarde", "el_pastas", "q_index", "answers_json", "statusas", "atnaujinta"]


@st.cache_resource
def get_worksheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=scopes
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(st.secrets["spreadsheet_id"])
    try:
        ws = sh.worksheet(WORKSHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=WORKSHEET_NAME, rows=2000, cols=len(SHEET_HEADERS))
        ws.append_row(SHEET_HEADERS)
    return ws


def make_test_id(vardas, pavarde):
    slug = f"{vardas.strip().lower()}_{pavarde.strip().lower()}"
    slug = re.sub(r"\s+", "_", slug)
    return slug


def find_progress(test_id):
    """Grąžina (eilutės_nr, įrašas) arba (None, None), jei nerasta."""
    ws = get_worksheet()
    try:
        cell = ws.find(test_id)
    except gspread.exceptions.CellNotFound:
        cell = None
    if cell is None:
        return None, None
    row_num = cell.row
    row_values = ws.row_values(row_num)
    headers = ws.row_values(1)
    record = dict(zip(headers, row_values))
    return row_num, record


def save_progress(test_id, vardas, pavarde, el_pastas, q_index, answers, statusas="vykdoma"):
    ws = get_worksheet()
    row_num = st.session_state.get("sheet_row_num")
    if row_num is None:
        row_num, _ = find_progress(test_id)
    answers_json = json.dumps(answers, ensure_ascii=False)
    atnaujinta = now_vilnius().strftime("%Y-%m-%d %H:%M:%S")
    row_values = [test_id, vardas, pavarde, el_pastas, str(q_index), answers_json, statusas, atnaujinta]
    if row_num:
        ws.update(range_name=f"A{row_num}:H{row_num}", values=[row_values])
    else:
        ws.append_row(row_values)
        row_num, _ = find_progress(test_id)
    st.session_state.sheet_row_num = row_num


def is_valid_email(email):
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email.strip()))


def normalize_text(text):
    """Sujungia suskaidytus Unicode simbolius į standartinę, sulipdytą formą,
    kad PDF šriftas juos atpažintų teisingai."""
    if not text:
        return text
    return unicodedata.normalize("NFC", text)


# =========================
# EL. PAŠTO SIUNTIMAS
# =========================
def send_email_with_results(to_email, vardas, pavarde, pdf_bytes, excel_bytes):
    smtp_user = st.secrets["email"]["smtp_user"]
    smtp_password = st.secrets["email"]["smtp_password"]

    msg = MIMEMultipart()
    msg["Subject"] = f"{TEST_TITLE} ataskaita – {vardas} {pavarde}"
    msg["From"] = smtp_user
    msg["To"] = to_email
    msg.attach(MIMEText(
        f"Sveiki,\n\nPridedama {vardas} {pavarde} {TEST_TITLE} testo ataskaita (PDF) ir "
        f"atsakymai (Excel).\n\nAutomatinis pranešimas.",
        "plain", "utf-8"
    ))

    part_pdf = MIMEApplication(pdf_bytes, Name="ataskaita.pdf")
    part_pdf["Content-Disposition"] = 'attachment; filename="ataskaita.pdf"'
    msg.attach(part_pdf)

    part_xlsx = MIMEApplication(excel_bytes, Name="atsakymai.xlsx")
    part_xlsx["Content-Disposition"] = 'attachment; filename="atsakymai.xlsx"'
    msg.attach(part_xlsx)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(smtp_user, smtp_password)
        server.send_message(msg)


# =========================
# SESSION STATE
# =========================
for key, default in [
    ("page", "start"), ("q_index", 0), ("answers", {}),
    ("name", ""), ("surname", ""), ("email", ""), ("test_id", ""), ("sheet_row_num", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default


# =========================
# SCORING
# =========================
def calculate_scores(answers):
    """SCL-90-R tipo balų skaičiavimas pagal raktai.json struktūrą:
    - KEYS["scales"][kodas] = {"name", "items": [...], "subtract": dalinys}
        -> skalės balas = (atsakymų suma tų items) / subtract  (vidurkis)
    - KEYS["total"] = {"items": [1..90], "subtract": 90} -> bendra suma / 90 = GSI
    - KEYS["derived_scores"]: GSI, PST (kiek items > 0), PSDI (suma / PST)
    - KEYS["reverse_items"]: klausimų numeriai, kurių balas apverčiamas
      (output_max - reikšmė), taikoma prieš sumuojant
    """
    output_range = KEYS.get("output_range", [0, 4])
    output_max = output_range[1]
    reverse_set = {str(x) for x in KEYS.get("reverse_items", [])}

    def get_value(item_id):
        key = str(item_id)
        val = answers.get(key)
        if val is None:
            return 0
        val = int(val)
        if key in reverse_set:
            val = output_max - val
        return val

    scores = {}

    # Devynios klinikinės skalės
    for scale_def in KEYS.get("scales", {}).values():
        items = scale_def.get("items", [])
        divisor = scale_def.get("subtract") or len(items) or 1
        raw_sum = sum(get_value(q) for q in items)
        label = scale_def.get("name", "Skalė")
        scores[label] = round(raw_sum / divisor, 2)

    # Bendras balas (visi 90 klausimų) — naudojamas GSI/PST/PSDI
    total_def = KEYS.get("total", {})
    total_items = total_def.get("items", [])
    total_divisor = total_def.get("subtract") or len(total_items) or 1
    total_sum = sum(get_value(q) for q in total_items)

    gsi = (total_sum / total_divisor) if total_divisor else 0
    pst = sum(1 for q in total_items if get_value(q) > 0)
    psdi = (total_sum / pst) if pst else 0

    scores["Bendras sunkumo indeksas (GSI)"] = round(gsi, 2)
    scores["Teigiamų simptomų kiekis (PST)"] = pst
    scores["Simptomų distreso indeksas (PSDI)"] = round(psdi, 2)

    return scores


# =========================
# START
# =========================
if st.session_state.page == "start":
    st.title(TEST_TITLE)
    col1, col2 = st.columns(2)
    with col1:
        name = st.text_input("Paciento vardas", value=st.session_state.name)
    with col2:
        surname = st.text_input("Paciento pavardė", value=st.session_state.surname)

    email = st.text_input(
        "El. paštas ataskaitai gauti",
        value=st.session_state.email,
        placeholder="pvz. gydytojas@pastas.lt"
    )

    existing_progress = None
    existing_row_num = None
    if name.strip() and surname.strip():
        try:
            tid = make_test_id(name, surname)
            existing_row_num, existing_progress = find_progress(tid)
        except Exception as e:
            st.warning(f"Nepavyko patikrinti ankstesnio progreso (Google Sheets ryšio klaida): {e}")
            existing_progress = None

    if existing_progress and existing_progress.get("statusas") == "vykdoma":
        answered = existing_progress.get("q_index", "0")
        st.info(f"Rastas neužbaigtas šio paciento testas ties {answered} klausimu. Ką norite daryti?")
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("Tęsti testą"):
                st.session_state.name = normalize_text(name.strip())
                st.session_state.surname = normalize_text(surname.strip())
                st.session_state.email = email.strip()
                st.session_state.test_id = make_test_id(name, surname)
                st.session_state.answers = json.loads(existing_progress.get("answers_json") or "{}")
                st.session_state.q_index = int(existing_progress.get("q_index") or 0)
                st.session_state.sheet_row_num = existing_row_num
                st.session_state.page = "test"
                st.rerun()
        with col_b:
            if st.button("Pradėti iš naujo"):
                if is_valid_email(email):
                    st.session_state.name = normalize_text(name.strip())
                    st.session_state.surname = normalize_text(surname.strip())
                    st.session_state.email = email.strip()
                    st.session_state.test_id = make_test_id(name, surname)
                    st.session_state.answers = {}
                    st.session_state.q_index = 0
                    st.session_state.sheet_row_num = existing_row_num
                    st.session_state.page = "test"
                    st.rerun()
                else:
                    st.warning("Įveskite teisingą el. pašto adresą")
    else:
        if st.button("Kitas"):
            if not (name.strip() and surname.strip()):
                st.warning("Įveskite vardą ir pavardę")
            elif not is_valid_email(email):
                st.warning("Įveskite teisingą el. pašto adresą")
            else:
                st.session_state.name = normalize_text(name.strip())
                st.session_state.surname = normalize_text(surname.strip())
                st.session_state.email = email.strip()
                st.session_state.test_id = make_test_id(name, surname)
                st.session_state.answers = {}
                st.session_state.q_index = 0
                st.session_state.page = "instrukcijos"
                try:
                    save_progress(
                        st.session_state.test_id, st.session_state.name,
                        st.session_state.surname, st.session_state.email,
                        0, {}, statusas="vykdoma"
                    )
                except Exception as e:
                    st.error(f"Nepavyko pradėti — Google Sheets ryšio klaida: {e}")
                    st.stop()
                st.rerun()

# =========================
# INSTRUKCIJOS (rodoma pacientui prieš pradedant klausimus)
# =========================
elif st.session_state.page == "instrukcijos":
    st.title("Instrukcijos")

    # Čia gali įdėti paveikslėlį arba trumpą vaizdo įrašą vietoje/šalia teksto:
    #   st.image("instrukcijos.png")               -- vietinis failas repo šaknyje
    #   st.image("https://.../paveikslelis.png")    -- nuoroda
    #   st.video("instrukcijos.mp4")                -- vietinis vaizdo failas
    #   st.video("https://youtu.be/...")             -- YouTube ar kita nuoroda
    if INSTRUCTIONS_TEXT:
        st.markdown(INSTRUCTIONS_TEXT.replace("\n", "  \n"))
    else:
        st.info("Instrukcijų tekstas nerastas (instrukcijos.txt).")

    if st.button("Pradėti testą"):
        st.session_state.page = "test"
        st.rerun()

# =========================
# TEST
# =========================
elif st.session_state.page == "test":
    i = st.session_state.q_index
    total = len(QUESTIONS)
    st.progress(i / total)
    st.markdown(f"<div class='question-header'>{QUESTION_HEADER}</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='question-text'>{i+1}. {QUESTIONS[i]}</div>", unsafe_allow_html=True)

    current = st.session_state.answers.get(str(i + 1))
    index = current if isinstance(current, int) and 0 <= current < len(SCALE_LABELS) else None
    selected_label = st.radio(" ", SCALE_LABELS, index=index, key=f"q_{i}")
    ans = SCALE_LABELS.index(selected_label) if selected_label is not None else None

    col1, col2, col3, col4 = st.columns(4)

    def persist():
        """Kas klausimą išsaugo progresą į Google Sheets, kad testą būtų galima tęsti bet kada vėliau."""
        try:
            save_progress(
                st.session_state.test_id, st.session_state.name,
                st.session_state.surname, st.session_state.email,
                st.session_state.q_index, st.session_state.answers,
                statusas="vykdoma"
            )
        except Exception as e:
            st.warning(f"Progreso išsaugoti nepavyko (bus laikoma tik šioje sesijoje): {e}")

    with col1:
        if st.button("Atgal"):
            st.session_state.answers[str(i + 1)] = ans
            if st.session_state.q_index > 0:
                st.session_state.q_index -= 1
            persist()
            st.rerun()
    with col2:
        if st.button("Kitas"):
            if ans is None:
                st.warning("Pasirinkite atsakymą")
            else:
                st.session_state.answers[str(i + 1)] = ans
                st.session_state.q_index += 1
                if st.session_state.q_index >= total:
                    st.session_state.page = "finish"
                persist()
                st.rerun()
    with col3:
        if st.button("Daryti pertrauką"):
            if ans is not None:
                st.session_state.answers[str(i + 1)] = ans
            persist()
            st.session_state.page = "paused"
            st.rerun()
    with col4:
        if st.button("Baigti testą"):
            if ans is not None:
                st.session_state.answers[str(i + 1)] = ans
            st.session_state.page = "finish"
            persist()
            st.rerun()

# =========================
# PERTRAUKA (progresas išsaugotas, ataskaita NESIUNČIAMA)
# =========================
elif st.session_state.page == "paused":
    st.title("Testas pristabdytas")
    st.success(
        f"Progresas išsaugotas ties {st.session_state.q_index}/{len(QUESTIONS)} klausimu. "
        f"Ataskaita NEBUVO išsiųsta — ji siunčiama tik atsakius į visus klausimus."
    )
    st.info(
        f"Norėdami tęsti vėliau, pradžios lange įveskite tą patį vardą ir pavardę: "
        f"**{st.session_state.name} {st.session_state.surname}**"
    )
    if st.button("Grįžti į pradžią"):
        for key, default in [
            ("page", "start"), ("q_index", 0), ("answers", {}),
            ("name", ""), ("surname", ""), ("email", ""), ("test_id", ""), ("sheet_row_num", None),
        ]:
            st.session_state[key] = default
        st.rerun()
    st.stop()

# =========================
# FINISH
# =========================
elif st.session_state.page == "finish":
    name = st.session_state.name
    surname = st.session_state.surname
    email = st.session_state.email

    # Apsauga nuo pakartotinio siuntimo: jei šioje sesijoje jau siųsta,
    # arba Google Sheets jau pažymėta "baigta" (pvz. po naršyklės atnaujinimo,
    # kuris sukuria naują sesiją, bet Sheets įrašas išlieka), nebekartojam.
    already_done = st.session_state.get("email_sent", False)
    if not already_done:
        try:
            _, existing = find_progress(st.session_state.test_id)
            if existing and existing.get("statusas") == "baigta":
                already_done = True
        except Exception:
            pass  # jei Sheets nepasiekiamas, tęsiam įprastai (geriau parodyti nei blokuoti)

    if already_done:
        st.success("Testas baigtas, ačiū už atsakymus.")
        st.stop()

    scores = calculate_scores(st.session_state.answers)

    # EXCEL
    excel_buffer = io.BytesIO()
    pd.DataFrame(
        list(st.session_state.answers.items()),
        columns=["Klausimas", "Atsakymas"]
    ).to_excel(excel_buffer, index=False, engine="openpyxl")
    excel_buffer.seek(0)

    # PDF
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        pdf_buffer, leftMargin=3 * cm, rightMargin=3 * cm,
        topMargin=3 * cm, bottomMargin=3 * cm
    )
    elements = []
    title_style = ParagraphStyle(name="Title", fontName=PDF_FONT_BOLD, fontSize=20, alignment=TA_CENTER, spaceAfter=20)
    normal_style = ParagraphStyle(name="Normal", fontName=PDF_FONT, fontSize=12, leading=18)
    score_style = ParagraphStyle(name="Score", fontName=PDF_FONT, fontSize=12, leading=20, spaceAfter=10)

    elements.append(Paragraph(f"{TEST_TITLE} ATASKAITA", title_style))
    elements.append(Paragraph(f"Atlikimo data: {now_vilnius().strftime('%Y-%m-%d %H:%M:%S')}", normal_style))
    elements.append(Spacer(1, 20))
    elements.append(Paragraph(f"Vardas: {name}", normal_style))
    elements.append(Paragraph(f"Pavarde: {surname}", normal_style))
    elements.append(Spacer(1, 20))
    elements.append(Paragraph("Rezultatai:", normal_style))
    elements.append(Spacer(1, 10))
    for k, v in scores.items():
        elements.append(Paragraph(f"{k}: <b>{v}</b>", score_style))

    doc.build(elements)
    pdf_buffer.seek(0)

    try:
        save_progress(
            st.session_state.test_id, name, surname, email,
            st.session_state.q_index, st.session_state.answers, statusas="baigta"
        )
    except Exception:
        pass  # klaidos pacientui nerodomos

    st.success("Testas baigtas, ačiū už atsakymus.")

    if email:
        try:
            send_email_with_results(
                email, name, surname,
                pdf_buffer.getvalue(), excel_buffer.getvalue()
            )
        except Exception:
            pass  # klaidos pacientui nerodomos; jei reikia, žr. Google Sheets "statusas" stulpelį

    st.session_state.email_sent = True
    st.stop()