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
TEST_TITLE = "SCL-90"          # rodoma antraštė / PDF pavadinimas / el. laiško tema
QUESTIONS_FILENAME = "Klausimai.txt"  # klausimų failo pavadinimas repo šaknyje
KEYS_FILENAME = "Raktai.json"         # raktų failo pavadinimas repo šaknyje
WORKSHEET_NAME = "Progresas"          # Google Sheets lapo (tab) pavadinimas šiam testui

# Jei kada nors abu testai naudotų TĄ PATĮ Google Sheet dokumentą, čia reikėtų
# skirtingo WORKSHEET_NAME kiekvienam testui. Kadangi šiam testui kuriamas
# atskiras Sheet (kitas spreadsheet_id secrets faile), pakanka to paties
# pavadinimo "Progresas" — jie negalės susimaišyti, nes yra skirtinguose failuose.

st.set_page_config(layout="wide")
st.markdown("""
<style>
#MainMenu {visibility: hidden;}
header {visibility: hidden;}
footer {visibility: hidden;}
.block-container {
    padding: 2rem;
    max-width: 100%;
}
html, body { font-size: 34px; }
h1 { font-size: 80px !important; }
.question-text {
    font-size: 50px !important;
    font-weight: 700;
    margin-bottom: 25px;
}
div[role="radiogroup"] label {
    font-size: 50px !important;
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
def normalize(ans):
    if ans == "Sutinku":
        return "T"
    if ans == "Nesutinku":
        return "K"
    return None


def calculate_scores(answers):
    scores = {k: 0 for k in KEYS.keys()}
    for q_idx, ans in answers.items():
        ans = normalize(ans)
        if not ans:
            continue
        for scale, items in KEYS.items():
            if str(q_idx) in items and items[str(q_idx)] == ans:
                scores[scale] += 1
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
        if st.button("Pradėti"):
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
                st.session_state.page = "test"
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
# TEST
# =========================
elif st.session_state.page == "test":
    i = st.session_state.q_index
    total = len(QUESTIONS)
    st.progress(i / total)
    st.markdown(f"<div class='question-text'>{i+1}. {QUESTIONS[i]}</div>", unsafe_allow_html=True)

    current = st.session_state.answers.get(str(i + 1))
    index = 0 if current == "Sutinku" else 1 if current == "Nesutinku" else None
    ans = st.radio(" ", ["Sutinku", "Nesutinku"], index=index, key=f"q_{i}")

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
    except Exception as e:
        st.warning(f"Nepavyko pažymėti testo kaip baigto Google Sheets: {e}")

    st.success("Testas baigtas, ačiū už atsakymus.")

    if email:
        try:
            send_email_with_results(
                email, name, surname,
                pdf_buffer.getvalue(), excel_buffer.getvalue()
            )
            st.success(f"Ataskaita automatiškai išsiųsta į {email}")
        except Exception as e:
            st.error(f"Nepavyko išsiųsti el. laiško automatiškai: {e}")

    pdf_buffer.seek(0)
    excel_buffer.seek(0)

    st.markdown("### Taip pat galite atsisiųsti rezultatus čia")
    col1, col2 = st.columns(2)
    filename_base = f"{name}_{surname}_{now_vilnius().strftime('%Y%m%d_%H%M%S')}"
    with col1:
        st.download_button(
            "Atsisiųsti PDF ataskaitą", data=pdf_buffer,
            file_name=f"{filename_base}_ataskaita.pdf", mime="application/pdf",
        )
    with col2:
        st.download_button(
            "Atsisiųsti Excel atsakymus", data=excel_buffer,
            file_name=f"{filename_base}_atsakymai.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    st.stop()
