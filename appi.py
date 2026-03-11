import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import requests
import xml.etree.ElementTree as ET
import pandas as pd
import base64
import io
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                 Table, TableStyle, HRFlowable)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Rekisteröi DejaVu Sans – täysi Unicode-tuki ÅÄÖ mukaan lukien
_DJVU_REG = False
def _register_fonts():
    global _DJVU_REG
    if _DJVU_REG:
        return
    try:
        pdfmetrics.registerFont(TTFont("DJV",  "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
        pdfmetrics.registerFont(TTFont("DJV-B","/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))
        pdfmetrics.registerFont(TTFont("DJV-I","/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf"))
        _DJVU_REG = True
    except Exception:
        pass  # fallback to Helvetica if fonts missing

def _font(bold=False):
    return "DJV-B" if (bold and _DJVU_REG) else ("Helvetica-Bold" if bold else ("DJV" if _DJVU_REG else "Helvetica"))

# -------------------------------------------------
# PAGE SETTINGS
# -------------------------------------------------
st.set_page_config(page_title="Sustaina", layout="wide", initial_sidebar_state="collapsed")

# -------------------------------------------------
# BACKGROUND
# -------------------------------------------------
def get_base64(file_path):
    try:
        with open(file_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except:
        return ""

bg_image = get_base64("background.jpg")
bg_css = (
    f".stApp {{ background: linear-gradient(rgba(0,0,0,0.65),rgba(0,0,0,0.65)),"
    f"url('data:image/jpg;base64,{bg_image}'); background-size:cover;"
    f"background-position:center; background-attachment:fixed; }}"
    if bg_image else
    ".stApp { background: linear-gradient(135deg, #0f2027, #203a43, #2c5364); }"
)

st.markdown(f"""
<style>
{bg_css}
body, p, div, span, label {{ color: white !important; }}
h1, h2, h3, h4, h5, h6 {{ color: white !important; }}
input, textarea {{ color: black !important; }}
div[data-baseweb="select"] > div {{
    background-color: rgba(0,0,0,0.6) !important; color: white !important; border-radius: 8px !important;
}}
div[data-baseweb="popover"] {{ background-color: #111 !important; }}
ul[role="listbox"] {{ background-color: #111 !important; }}
li[role="option"] {{ background-color: #111 !important; color: white !important; }}
li[role="option"]:hover {{ background-color: #333 !important; }}
[data-testid="stMetricLabel"], [data-testid="stMetricValue"] {{ color: white !important; }}
.stAlert {{ background: rgba(255,255,255,0.08) !important; color: white !important; border-radius: 10px; }}
.stButton > button {{
    background-color: #1a3a2a !important; color: white !important;
    border-radius: 8px; border: 1px solid #2d6a4f; font-weight: 600; padding: 0.5rem 1.5rem;
}}
.stButton > button:hover {{ background-color: #2d6a4f !important; }}
.stDownloadButton > button {{
    background-color: #2d6a4f !important; color: white !important;
    border-radius: 8px; border: 2px solid #52b788 !important;
    font-weight: 700; padding: 0.6rem 2rem; font-size: 1rem;
    width: 100%;
}}
.stDownloadButton > button:hover {{ background-color: #1b4332 !important; color: white !important; }}
.badge {{
    display: inline-block; background: rgba(45,106,79,0.4);
    border: 1px solid #52b788; border-radius: 20px; padding: 3px 12px;
    font-size: 12px; color: #52b788 !important; margin-bottom: 8px;
}}
.sources-box {{
    background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.12);
    border-radius: 12px; padding: 20px 24px; margin-top: 32px; font-size: 13px;
}}
.disclaimer {{
    background: rgba(255,200,0,0.08); border: 1px solid rgba(255,200,0,0.25);
    border-radius: 8px; padding: 12px 16px; margin-top: 12px;
    font-size: 12px; color: #ffd166 !important;
}}
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------
# HEADER – pelkkä teksti, ei logo-riippuvuutta
# -------------------------------------------------
st.markdown("""
<div style="padding: 8px 0 20px 0; border-bottom: 1px solid rgba(255,255,255,0.12); margin-bottom: 16px;">
    <h1 style="margin:0 0 4px 0; font-size:2.2rem; letter-spacing:-0.5px;">Sustaina</h1>
    <p style="margin:0; opacity:0.65; font-size:0.95rem;">
        Yrityksen hiilijalanjälkilaskuri &ndash; GHG Protocol Scope 1, 2 &amp; 3
    </p>
</div>
""", unsafe_allow_html=True)

# -------------------------------------------------
# PAASTOKERTOIMET – ajantasaiset, lähdeviitatut
# -------------------------------------------------
# Liikenne: EMEP/EEA Guidebook 2023 & UK DESNZ GHG Conversion Factors 2024
# (VTT LIPASTO poistunut kaytosta 2023 – Tilastokeskus LIIKE korvaa)
EMISSION_FACTORS = {
    "car_bensa":  0.170,   # Bensiiniauto Euro 6 ka., EMEP/EEA 2023
    "car_diesel": 0.195,   # Diesel-auto Euro 6 ka., EMEP/EEA 2023
    "car_ev":     0.053,   # Sähköauto, Suomen sähkömix 60 g/kWh ~15 kWh/100km
    "train":      0.035,   # Juna henkilökm, VR ympäristöraportti 2023
    "bus":        0.089,   # Bussi henkilökm, EMEP/EEA 2023 / UK DESNZ 2024
    "flights":    0.255,   # Lento henkilökm incl. RFI 1.9, ICAO / UK DESNZ 2024
    "water":      0.344,   # m3, Suomen vesilaitos ka. (energia + kemikaalit)
    "waste":      0.210,   # Sekajäte kaatopaikka, SYKE 2022
    "recycling":  0.021,   # Kierrätys (hyötykäyttö), SYKE / GHG Protocol
    "bio":        0.100,   # Biojäte kompostointi, SYKE 2022
    "paper":      1.29,    # Paperi/kartonki LCA, SYKE / ecoinvent v3.9
    "plastic":    2.53,    # Muovi (sekalainen) LCA, ecoinvent v3.9 / SYKE
    # --- Scope 3 lisäkategoriat ---
    "commute_car":    0.170,   # Pendelöinti henkilöauto (ka.), EMEP/EEA 2023
    "commute_pt":     0.050,   # Pendelöinti joukkoliikenne (ka. juna+bussi)
    "commute_ev":     0.053,   # Pendelöinti sähköauto
    "biztravel_car":  0.170,   # Liikematka auto, EMEP/EEA 2023
    "biztravel_taxi": 0.170,   # Taksi (sama kuin henkilöauto)
    "hotel":          8.0,     # Hotelliyö kg CO2e/yö, Cornell Hotel Sustainability Benchmark 2023
    "it_equipment":   300.0,   # IT-laite (kannettava) kg CO2e/kpl, Dell/Apple LCA-raportit
    "office_equip":   150.0,   # Toimistolaite (näyttö tms.) kg CO2e/kpl
    "purchased_goods":0.50,    # Ostetut tavarat/palvelut kg CO2e/€ – Exiobase / EEIO-kerroin (ka.)
}

# Lammitys kg CO2e/kWh – Motiva.fi / Tilastokeskus (paivitetty 2024)
HEATING_FACTORS = {
    "Kaukolämpö (Suomen ka.)": 0.130,  # Motiva/Tilastokeskus 3v ka. 2021-2023 = 130 g CO2/kWh (energiamenetelmä)
    "Kevyt polttoöljy":         0.267,  # Tilastokeskus polttoaineluokitus
    "Maakaasu":                 0.198,  # Tilastokeskus polttoaineluokitus
    "Sähkölämmitys (verkko)":   0.060,  # Sama kuin sähkön avg kerroin, Motiva 2024
    "Maalämpöpumppu":           0.020,  # COP ~3.0–3.5, perustuu sähkökertoimen kautta
    "Pelletit/biomassa":        0.025,  # Motiva – uusiutuva biomassa, lähes hiillineutraali
}

DEFAULT_ELECTRIC_FACTOR = 0.060  # Motiva/Tilastokeskus 3v liukuva ka. 2021-2023 = 60 g CO2/kWh
# Huom: Energiaviraston 2023 vuosikerroin on 47.2 g CO2e/kWh (trendi laskeva)
# GHG Protocol suosittelee liukuvaa keskiarvoa – käytetään 60 g konservatiivisena arvona

SCOPE_MAP = {
    "Sähkö": "Scope 2", "Lämmitys": "Scope 1/2",
    "Bensa-auto": "Scope 1", "Diesel-auto": "Scope 1", "Sähköauto": "Scope 2",
    "Juna": "Scope 3", "Bussi": "Scope 3", "Lennot": "Scope 3", "Vesi": "Scope 3",
    "Sekajäte": "Scope 3", "Kierrätys": "Scope 3", "Biojäte": "Scope 3",
    "Paperi": "Scope 3", "Muovi": "Scope 3",
    # Scope 3 lisä
    "Pendelöinti (auto)":   "Scope 3", "Pendelöinti (julkinen)": "Scope 3",
    "Pendelöinti (sähköauto)": "Scope 3",
    "Liikematka (auto)":    "Scope 3", "Taksimatkat":            "Scope 3",
    "Hotelliöt":            "Scope 3", "IT-laitteet":            "Scope 3",
    "Toimistolaitteet":     "Scope 3", "Ostetut tavarat/palvelut": "Scope 3",
}

INDUSTRY_DEFAULTS = {
    "Toimisto / Hallinto": {
        "electricity": 25000, "heating": 40000, "water": 300,
        "car_bensa": 5000, "car_diesel": 2000, "car_ev": 2000,
        "train": 8000, "bus": 3000, "flights": 15000,
        "waste": 800, "recycling": 300, "bio": 200, "paper": 150, "plastic": 80
    },
    "Teollisuus / Valmistus": {
        "electricity": 200000, "heating": 300000, "water": 5000,
        "car_bensa": 20000, "car_diesel": 30000, "car_ev": 5000,
        "train": 5000, "bus": 2000, "flights": 30000,
        "waste": 10000, "recycling": 3000, "bio": 500, "paper": 500, "plastic": 2000
    },
    "Logistiikka / Kuljetus": {
        "electricity": 50000, "heating": 60000, "water": 400,
        "car_bensa": 10000, "car_diesel": 150000, "car_ev": 5000,
        "train": 2000, "bus": 1000, "flights": 20000,
        "waste": 2000, "recycling": 500, "bio": 200, "paper": 100, "plastic": 300
    },
    "Vähittäiskauppa": {
        "electricity": 80000, "heating": 50000, "water": 800,
        "car_bensa": 8000, "car_diesel": 10000, "car_ev": 3000,
        "train": 3000, "bus": 2000, "flights": 10000,
        "waste": 5000, "recycling": 2000, "bio": 1000, "paper": 500, "plastic": 1000
    },
    "IT / Teknologia": {
        "electricity": 60000, "heating": 35000, "water": 250,
        "car_bensa": 4000, "car_diesel": 1000, "car_ev": 5000,
        "train": 12000, "bus": 4000, "flights": 40000,
        "waste": 600, "recycling": 400, "bio": 150, "paper": 80, "plastic": 60
    },
    "Muu / Mukautettu": {
        "electricity": 3000, "heating": 12000, "water": 150,
        "car_bensa": 10000, "car_diesel": 0, "car_ev": 5000,
        "train": 2000, "bus": 500, "flights": 1000,
        "waste": 400, "recycling": 50, "bio": 100, "paper": 20, "plastic": 15
    },
}

# -------------------------------------------------
# DATA FETCHING
# -------------------------------------------------
@st.cache_data
def get_finland_emission_trend():
    url = "https://pxdata.stat.fi/PXWeb/api/v1/fi/StatFin/khki/statfin_khki_pxt_11ic.px"
    query = {
        "query": [{"code": "Sektori", "selection": {"filter": "item", "values": ["Yhteensa"]}}],
        "response": {"format": "json-stat"}
    }
    try:
        r = requests.post(url, json=query, timeout=10)
        data = r.json()
        values = data["value"]
        years  = list(data["dimension"]["Vuosi"]["category"]["label"].values())
        df = pd.DataFrame({"Vuosi": years, "Mt_CO2e": values})
        df["kg_per_person"] = (df["Mt_CO2e"] * 1_000_000 * 1000) / 5_550_000
        return df
    except:
        return None

@st.cache_data
def get_current_temperature():
    url = (
        "https://opendata.fmi.fi/wfs?service=WFS&version=2.0.0&request=GetFeature"
        "&storedquery_id=fmi::observations::weather::simple"
        "&place=Helsinki&parameters=t2m&maxlocations=1"
    )
    try:
        r = requests.get(url, timeout=8)
        root = ET.fromstring(r.content)
        for elem in root.iter():
            if "ParameterValue" in elem.tag:
                return float(elem.text)
    except:
        return None

# -------------------------------------------------
# APP STATE
# -------------------------------------------------
pages = ["Yritysinfo", "Toimitilat", "Liikkuminen", "Jatteet", "Scope 3", "Yhteenveto"]

for key, default in [
    ("page_index", 0), ("data", {}), ("industry", "Muu / Mukautettu"),
    ("employees", 10), ("company_name", ""),
    ("heating_type", list(HEATING_FACTORS.keys())[0])
]:
    if key not in st.session_state:
        st.session_state[key] = default

def next_page():
    if st.session_state.page_index < len(pages) - 1:
        st.session_state.page_index += 1

def prev_page():
    if st.session_state.page_index > 0:
        st.session_state.page_index -= 1

def reset_app():
    st.session_state.page_index = 0
    st.session_state.data = {}

current_page = pages[st.session_state.page_index]
st.progress((st.session_state.page_index + 1) / len(pages))
st.caption(f"Vaihe {st.session_state.page_index + 1} / {len(pages)}")

# -------------------------------------------------
# LASKENTA
# -------------------------------------------------
def calculate_emissions(data, electric_factor):
    hf = data.get("heating_factor", list(HEATING_FACTORS.values())[0])
    return_dict = {
        "Sähkö":       data.get("electricity", 0) * electric_factor,
        "Lämmitys":    data.get("heating",     0) * hf,
        "Bensa-auto":  data.get("car_bensa",   0) * EMISSION_FACTORS["car_bensa"],
        "Diesel-auto": data.get("car_diesel",  0) * EMISSION_FACTORS["car_diesel"],
        "Sähköauto":   data.get("car_ev",      0) * EMISSION_FACTORS["car_ev"],
        "Juna":        data.get("train",       0) * EMISSION_FACTORS["train"],
        "Bussi":       data.get("bus",         0) * EMISSION_FACTORS["bus"],
        "Lennot":      data.get("flights",     0) * EMISSION_FACTORS["flights"],
        "Vesi":        data.get("water",       0) * EMISSION_FACTORS["water"],
        "Sekajäte":    data.get("waste",       0) * EMISSION_FACTORS["waste"],
        "Kierrätys":   data.get("recycling",   0) * EMISSION_FACTORS["recycling"],
        "Biojäte":     data.get("bio",         0) * EMISSION_FACTORS["bio"],
        "Paperi":      data.get("paper",       0) * EMISSION_FACTORS["paper"],
        "Muovi":       data.get("plastic",     0) * EMISSION_FACTORS["plastic"],
        # Scope 3 lisäkategoriat
        "Pendelöinti (auto)":      data.get("commute_car",   0) * EMISSION_FACTORS["commute_car"],
        "Pendelöinti (julkinen)":  data.get("commute_pt",    0) * EMISSION_FACTORS["commute_pt"],
        "Pendelöinti (sähköauto)": data.get("commute_ev",    0) * EMISSION_FACTORS["commute_ev"],
        "Liikematka (auto)":       data.get("biztravel_car", 0) * EMISSION_FACTORS["biztravel_car"],
        "Taksimatkat":             data.get("biztravel_taxi",0) * EMISSION_FACTORS["biztravel_taxi"],
        "Hotelliöt":               data.get("hotel_nights",  0) * EMISSION_FACTORS["hotel"],
        "IT-laitteet":             data.get("it_devices",    0) * EMISSION_FACTORS["it_equipment"],
        "Toimistolaitteet":        data.get("office_devices",0) * EMISSION_FACTORS["office_equip"],
        "Ostetut tavarat/palvelut":data.get("purchased_goods_eur", 0) * EMISSION_FACTORS["purchased_goods"],
    }
    return sum(return_dict.values()), return_dict

def suggest_savings(emissions, employees):
    emp = max(employees, 1)
    tips = []
    if emissions.get("Lennot", 0) / emp > 50:
        tips.append("Lentomatkat: harkitse videokokouksia tai junaa – lennot ovat suurin yksittäinen päästölähde per henkilö")
    if emissions.get("Bensa-auto", 0) / emp > 150:
        tips.append("Bensiiniajoneuvokalusto: siirry sähköautoihin tai vähennä ajomäärää")
    if emissions.get("Diesel-auto", 0) / emp > 200:
        tips.append("Diesel-kalusto: sähköinen logistiikka vähentää päästöjä merkittävästi")
    if emissions.get("Sähkö", 0) / emp > 200:
        tips.append("Sähkö: vaihda sertifioituun uusiutuvaan – kerroin laskee lähes nollaan (0.010 kg CO2e/kWh)")
    if emissions.get("Vesi", 0) / emp > 10:
        tips.append("Vesi: optimoi prosessit ja vähennä kulutusta")
    if emissions.get("Sekajäte", 0) / emp > 20:
        tips.append("Jätteet: lisää kierrätystä – sekajätteen kerroin on 10x kierrätystä korkeampi")
    if emissions.get("Muovi", 0) / emp > 50:
        tips.append("Muovi: vaihda kierrätettäviin pakkausmateriaaleihin (muovin LCA-kerroin 2.53 kg CO2e/kg)")
    if emissions.get("Pendelöinti (auto)", 0) / emp > 200:
        tips.append("Pendelöinti: kannusta julkiseen liikenteeseen tai sähköautoihin – työmatkaliikkuminen on merkittävä Scope 3 -lähde")
    if emissions.get("Hotelliöt", 0) / emp > 100:
        tips.append("Hotellit: valitse sertifioituja vihreitä hotelleja (Green Key / ISO 14001) – säästö jopa 50%")
    if emissions.get("Ostetut tavarat/palvelut", 0) / emp > 500:
        tips.append("Toimitusketju: tee hankinnoista päästöarvio ja suosi toimittajia joilla on SBTi-tavoite")
    if emissions.get("IT-laitteet", 0) / emp > 100:
        tips.append("IT-laitteet: pidennä laitteiden elinkaarta – uuden kannettavan hiilijalanjälki on ~300 kg CO2e")
    return tips

# -------------------------------------------------
# PDF-RAPORTTI – siisti, ei erikoismerkkeja
# -------------------------------------------------
def generate_pdf_report(data, emissions, total, employees, company_name, industry, scenario):
    _register_fonts()
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        rightMargin=2.2*cm, leftMargin=2.2*cm,
        topMargin=2*cm, bottomMargin=2*cm
    )

    GREEN     = colors.HexColor("#1b4332")
    MID_GREEN = colors.HexColor("#2d6a4f")
    GREY_BG   = colors.HexColor("#f4f7f5")
    DARK      = colors.HexColor("#111111")
    MID_GREY  = colors.HexColor("#666666")
    WHITE     = colors.white

    base_style = getSampleStyleSheet()["Normal"]

    def S(name, **kw):
        return ParagraphStyle(name, parent=base_style, **kw)

    styles = {
        "title":  S("T",  fontSize=22, textColor=GREEN,     fontName=_font(True),  spaceAfter=0,  spaceBefore=0, leading=26),
        "sub":    S("Su", fontSize=10, textColor=MID_GREY,  fontName=_font(False), spaceAfter=14, spaceBefore=0, leading=14),
        "h2":     S("H2", fontSize=12, textColor=MID_GREEN, fontName=_font(True),  spaceBefore=14, spaceAfter=5),
        "body":   S("B",  fontSize=9,  textColor=DARK,      fontName=_font(False), spaceAfter=3,  leading=13),
        "small":  S("Sm", fontSize=8,  textColor=MID_GREY,  fontName=_font(False), spaceAfter=2,  leading=11),
        "bullet": S("Bu", fontSize=9,  textColor=DARK,      fontName=_font(False), spaceAfter=3,  leftIndent=10, leading=13),
        "cap":    S("Ca", fontSize=7.5,textColor=MID_GREY,  fontName=_font(False), spaceAfter=1),
    }

    story = []

    # Otsikko – eksplisiittinen Spacer estää päällekkäisyyden
    story.append(Paragraph("Sustaina", styles["title"]))
    story.append(Spacer(1, 6))
    story.append(Paragraph("Hiilijalanjälkiraportti  |  GHG Protocol Scope 1, 2 & 3", styles["sub"]))
    story.append(HRFlowable(width="100%", thickness=1.5, color=MID_GREEN, spaceAfter=12))

    # Metatiedot
    meta = [
        ["Yritys",      company_name or "Ei määritelty", "Toimiala",   industry],
        ["Henkilöstö",  str(employees),                   "Päivämäärä", datetime.now().strftime("%d.%m.%Y")],
        ["Skenaario",   scenario,                         "Laskuri",    "Sustaina v2.0"],
    ]
    tm = Table(meta, colWidths=[3*cm, 5.5*cm, 3*cm, 5.5*cm])
    tm.setStyle(TableStyle([
        ("FONTNAME",  (0,0), (0,-1), _font(True)),
        ("FONTNAME",  (2,0), (2,-1), _font(True)),
        ("FONTNAME",  (1,0), (1,-1), _font(False)),
        ("FONTNAME",  (3,0), (3,-1), _font(False)),
        ("FONTSIZE",  (0,0), (-1,-1), 8.5),
        ("TEXTCOLOR", (0,0), (0,-1), MID_GREY),
        ("TEXTCOLOR", (2,0), (2,-1), MID_GREY),
        ("ROWBACKGROUNDS", (0,0), (-1,-1), [WHITE, GREY_BG]),
        ("GRID",    (0,0), (-1,-1), 0.3, colors.HexColor("#dddddd")),
        ("PADDING", (0,0), (-1,-1), 5),
    ]))
    story.append(tm)
    story.append(Spacer(1, 14))

    # Yhteenveto
    story.append(Paragraph("Yhteenveto", styles["h2"]))
    per_person = total / max(employees, 1)
    tonnes     = total / 1000
    sum_rows = [
        ["Mittari",                    "Arvo",               "Yksikkö"],
        ["Kokonaispäästöt (vuosi)",    f"{total:,.0f}",      "kg CO2e"],
        ["Kokonaispäästöt (vuosi)",    f"{tonnes:,.2f}",     "t CO2e"],
        ["Päästöt per henkilö",        f"{per_person:,.0f}", "kg CO2e / hlö / vuosi"],
        ["Skenaario",                  scenario,             ""],
    ]
    ts = Table(sum_rows, colWidths=[7*cm, 5*cm, 5*cm])
    ts.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), MID_GREEN),
        ("TEXTCOLOR",  (0,0), (-1,0), WHITE),
        ("FONTNAME",   (0,0), (-1,0), _font(True)),
        ("FONTNAME",   (0,1), (-1,-1), _font(False)),
        ("FONTSIZE",   (0,0), (-1,-1), 9),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [WHITE, GREY_BG]),
        ("GRID",   (0,0), (-1,-1), 0.3, colors.HexColor("#cccccc")),
        ("PADDING",(0,0), (-1,-1), 6),
        ("ALIGN",  (1,0), (1,-1), "RIGHT"),
    ]))
    story.append(ts)
    story.append(Spacer(1, 14))

    # Kategoriat
    story.append(Paragraph("Päästöt kategorioittain", styles["h2"]))
    bd = [["Kategoria", "Scope", "kg CO2e", "Osuus (%)", "Per hlö (kg)"]]
    for cat, val in sorted(emissions.items(), key=lambda x: -x[1]):
        pct   = (val / total * 100) if total > 0 else 0
        scope = SCOPE_MAP.get(cat, "-")
        bd.append([cat, scope, f"{val:,.0f}", f"{pct:.1f}%", f"{val/max(employees,1):,.0f}"])
    tb = Table(bd, colWidths=[4*cm, 2.2*cm, 3.5*cm, 2.5*cm, 4.8*cm])
    tb.setStyle(TableStyle([
        ("BACKGROUND",     (0,0), (-1,0), MID_GREEN),
        ("TEXTCOLOR",      (0,0), (-1,0), WHITE),
        ("FONTNAME",       (0,0), (-1,0), _font(True)),
        ("FONTNAME",       (0,1), (-1,-1), _font(False)),
        ("FONTSIZE",       (0,0), (-1,-1), 8.5),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [WHITE, GREY_BG]),
        ("GRID",           (0,0), (-1,-1), 0.3, colors.HexColor("#cccccc")),
        ("PADDING",        (0,0), (-1,-1), 5),
        ("ALIGN",          (2,0), (-1,-1), "RIGHT"),
    ]))
    story.append(tb)
    story.append(Spacer(1, 14))

    # Säästösuositukset
    tips = suggest_savings(emissions, employees)
    if tips:
        story.append(Paragraph("Säästösuositukset", styles["h2"]))
        for t in tips:
            story.append(Paragraph(f"  \u2022  {t}", styles["bullet"]))
        story.append(Spacer(1, 10))

    # Käytetyt kertoimet
    story.append(Paragraph("Käytetyt päästökertoimet", styles["h2"]))
    ef_rows = [["Kategoria", "Kerroin", "Yksikkö", "Lähde"]] + [
        ("Bensa-auto",         "0.170", "kg CO2e/km",  "EMEP/EEA Guidebook 2023"),
        ("Diesel-auto",        "0.195", "kg CO2e/km",  "EMEP/EEA Guidebook 2023"),
        ("Sähköauto",          "0.053", "kg CO2e/km",  "Motiva + Tilastokeskus 2023"),
        ("Juna",               "0.035", "kg CO2e/km",  "VR ympäristöraportti 2023"),
        ("Bussi",              "0.089", "kg CO2e/km",  "EMEP/EEA 2023 / UK DESNZ 2024"),
        ("Lennot (sis. RFI)",  "0.255", "kg CO2e/km",  "ICAO + UK DESNZ GHG CF 2024"),
        ("Kaukolämpö ka.",     "0.130", "kg CO2e/kWh", "Motiva / Tilastokeskus 2021-23 ka."),
        ("Sähkö (Suomi ka.)",  "0.060", "kg CO2e/kWh", "Motiva / Tilastokeskus 2021-23 (60 g); 2023: 47 g"),
        ("Sekajäte",           "0.210", "kg CO2e/kg",  "SYKE 2022"),
        ("Muovi",              "2.53",  "kg CO2e/kg",  "ecoinvent v3.9 / SYKE"),
        ("Paperi",             "1.29",  "kg CO2e/kg",  "SYKE / ecoinvent v3.9"),
        ("Vesi",               "0.344", "kg CO2e/m3",  "Suomen vesilaitos ka."),
    ]
    tef = Table(ef_rows, colWidths=[3.5*cm, 2*cm, 3*cm, 8.5*cm])
    tef.setStyle(TableStyle([
        ("BACKGROUND",     (0,0), (-1,0), MID_GREEN),
        ("TEXTCOLOR",      (0,0), (-1,0), WHITE),
        ("FONTNAME",       (0,0), (-1,0), _font(True)),
        ("FONTNAME",       (0,1), (-1,-1), _font(False)),
        ("FONTSIZE",       (0,0), (-1,-1), 7.5),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [WHITE, GREY_BG]),
        ("GRID",           (0,0), (-1,-1), 0.3, colors.HexColor("#cccccc")),
        ("PADDING",        (0,0), (-1,-1), 4),
        ("ALIGN",          (1,0), (1,-1), "RIGHT"),
    ]))
    story.append(tef)
    story.append(Spacer(1, 16))

    # Lähteet
    story.append(HRFlowable(width="100%", thickness=0.5,
                             color=colors.HexColor("#cccccc"), spaceAfter=8))
    story.append(Paragraph("Lähteet", styles["h2"]))
    sources = [
        "Sähkö (Suomi): Motiva.fi / Tilastokeskus – Energia ja päästöt 2024, 3v ka. 2021–2023 = 60 g CO2/kWh",
        "Lämmitys: Motiva.fi / Energiateollisuus ry – Kaukolämpötilasto 2024; Tilastokeskus polttoaineet",
        "Liikenne: EMEP/EEA Air Pollutant Emission Inventory Guidebook 2023",
        "  Huom: VTT LIPASTO poistunut käytöstä 2023 – korvaaja Tilastokeskus LIIKE-järjestelmä",
        "Lennot (RFI): ICAO Carbon Emissions Calculator + UK DESNZ GHG Conversion Factors 2024",
        "Jätteet: Suomen ympäristökeskus SYKE – Jätteiden kasvihuonekaasupäästöt 2022",
        "Muovi & paperi LCA: ecoinvent v3.9 / SYKE",
        "Suomen kokonaispäästöt: Tilastokeskus – statfin_khki_pxt_11ic.px",
        "Säätiedot: Ilmatieteen laitos FMI Open Data (opendata.fmi.fi)",
        "Laskentakehikko: GHG Protocol Corporate Standard – ghgprotocol.org",
    ]
    for s in sources:
        story.append(Paragraph(f"  {s}", styles["small"]))

    story.append(Spacer(1, 8))
    story.append(Paragraph(
        "Huomio: Tämä laskuri antaa suuntaa-antavan arvion. Virallinen GHG-raportointi suositellaan "
        "tehtäväksi toimialakohtaisilla primäärikerroimilla (ISO 14064 / GHG Protocol). "
        "Päästökertoimet päivitetään vuosittain.",
        styles["small"]
    ))
    story.append(Spacer(1, 6))
    story.append(Paragraph(
        f"Raportti generoitu: {datetime.now().strftime('%d.%m.%Y klo %H:%M')}  |  Sustaina v2.0",
        styles["cap"]
    ))

    doc.build(story)
    buffer.seek(0)
    return buffer

# =====================================================================
# SIVUT
# =====================================================================

if current_page == "Yritysinfo":
    st.header("Yrityksen perustiedot")
    st.markdown('<span class="badge">GHG Protocol Scope 1, 2 & 3</span>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        company_name = st.text_input("Yrityksen nimi", value=st.session_state.company_name)
        st.session_state.company_name = company_name
    with col2:
        employees = st.number_input("Henkilöstömäärä", 1, 100000, st.session_state.employees)
        st.session_state.employees = employees

    industry = st.selectbox(
        "Toimiala", list(INDUSTRY_DEFAULTS.keys()),
        index=list(INDUSTRY_DEFAULTS.keys()).index(st.session_state.industry)
    )
    st.session_state.industry = industry

    if st.session_state.data.get("_industry") != industry:
        for key, val in INDUSTRY_DEFAULTS[industry].items():
            st.session_state.data[key] = val
        st.session_state.data["_industry"] = industry

    st.info(f"Toimialavalinta '{industry}' täyttää oletusarvot automaattisesti. Voit muuttaa kaikkia arvoja.")
    st.markdown("""
<div class="disclaimer">
Laskuri kattaa GHG Protocol Scope 1 (suorat päästöt), Scope 2 (ostettu energia) ja
Scope 3 (pendelöinti, liikematkat, jätteet, hankitut tavarat, pääomahyödykkeet).
Päästöt ovat suuntaa-antavia – virallinen GHG-raportointi vaatii toimittajakohtaiset primääridatat.
</div>
""", unsafe_allow_html=True)
    st.button("Seuraava ->", on_click=next_page)

elif current_page == "Toimitilat":
    st.header("Toimitilat & Energia")
    st.markdown('<span class="badge">Scope 1 & 2</span>', unsafe_allow_html=True)

    electric_factor = st.slider(
        "Sähkön päästökerroin (kg CO2e/kWh)", 0.0, 0.30,
        float(st.session_state.data.get("electric_factor", DEFAULT_ELECTRIC_FACTOR)),
        0.001,
        help="Suomen ka. 2021–2023: 0.060 (Motiva/Tilastokeskus). Sertifioitu uusiutuva: ~0.010"
    )

    col1, col2 = st.columns(2)
    with col1:
        electricity = st.number_input("Sähkö (kWh/vuosi)", 0, 10_000_000,
                                       int(st.session_state.data.get("electricity", 3000)))
        heating     = st.number_input("Lämmitys (kWh/vuosi)", 0, 20_000_000,
                                       int(st.session_state.data.get("heating", 12000)))
    with col2:
        hkeys       = list(HEATING_FACTORS.keys())
        saved_ht    = st.session_state.heating_type
        ht_index    = hkeys.index(saved_ht) if saved_ht in hkeys else 0
        heating_type = st.selectbox("Lämmitystapa", hkeys, index=ht_index)
        st.session_state.heating_type = heating_type
        hf = HEATING_FACTORS[heating_type]
        st.caption(f"Päästökerroin: {hf:.3f} kg CO2e/kWh  |  Lähde: Motiva.fi / Tilastokeskus 2024")
        water = st.number_input("Vesi (m³/vuosi)", 0, 500_000,
                                 int(st.session_state.data.get("water", 150)))

    st.session_state.data.update({
        "electricity": electricity, "heating": heating,
        "heating_factor": hf, "water": water, "electric_factor": electric_factor
    })
    col1, col2 = st.columns(2)
    col1.button("<- Edellinen", on_click=prev_page)
    col2.button("Seuraava ->", on_click=next_page)

elif current_page == "Liikkuminen":
    st.header("Liikkuminen & Matkat")
    st.markdown('<span class="badge">Scope 1 & 3</span>', unsafe_allow_html=True)
    st.info("Syötä yrityksen kaikkien ajoneuvojen ja työmatkojen vuotuiset kilometrit yhteensä.")

    col1, col2 = st.columns(2)
    with col1:
        car_bensa  = st.number_input("Bensiiniauto (km/vuosi)",  0, 10_000_000, int(st.session_state.data.get("car_bensa",  10000)))
        car_diesel = st.number_input("Diesel-auto (km/vuosi)",   0, 10_000_000, int(st.session_state.data.get("car_diesel", 0)))
        car_ev     = st.number_input("Sähköauto (km/vuosi)",     0, 10_000_000, int(st.session_state.data.get("car_ev",     5000)))
    with col2:
        train   = st.number_input("Juna (km/vuosi)",   0, 5_000_000,  int(st.session_state.data.get("train",   2000)))
        bus     = st.number_input("Bussi (km/vuosi)",  0, 5_000_000,  int(st.session_state.data.get("bus",     500)))
        flights = st.number_input("Lennot (km/vuosi)", 0, 10_000_000, int(st.session_state.data.get("flights", 1000)),
                                   help="Sisältää RFI-kertoimen 1.9 – kattaa lentoliikenteen kokonaisilmastovaikutuksen")

    st.session_state.data.update({
        "car_bensa": car_bensa, "car_diesel": car_diesel, "car_ev": car_ev,
        "train": train, "bus": bus, "flights": flights
    })
    col1, col2 = st.columns(2)
    col1.button("<- Edellinen", on_click=prev_page)
    col2.button("Seuraava ->", on_click=next_page)

elif current_page == "Jatteet":
    st.header("Jätteet & Kierrätys")
    st.markdown('<span class="badge">Scope 3</span>', unsafe_allow_html=True)
    st.info("Kierrätysluvun kerroin (0.021) on huomattavasti pienempi kuin sekajätteen (0.210) – lajittelu kannattaa.")

    waste     = st.number_input("Sekajäte (kg/vuosi)",  0, 1_000_000, int(st.session_state.data.get("waste",     400)))
    recycling = st.number_input("Kierrätys (kg/vuosi)", 0, 500_000,   int(st.session_state.data.get("recycling", 50)))
    bio       = st.number_input("Biojäte (kg/vuosi)",   0, 500_000,   int(st.session_state.data.get("bio",       100)))
    paper     = st.number_input("Paperi (kg/vuosi)",    0, 200_000,   int(st.session_state.data.get("paper",     20)))
    plastic   = st.number_input("Muovi (kg/vuosi)",     0, 200_000,   int(st.session_state.data.get("plastic",   15)))

    st.session_state.data.update({
        "waste": waste, "recycling": recycling, "bio": bio, "paper": paper, "plastic": plastic
    })
    col1, col2 = st.columns(2)
    col1.button("<- Edellinen", on_click=prev_page)
    col2.button("Seuraava ->", on_click=next_page)

elif current_page == "Scope 3":
    st.header("Scope 3 – Lisäkategoriat")
    st.markdown('<span class="badge">Scope 3 – Epäsuorat päästöt</span>', unsafe_allow_html=True)
    st.info("Nämä kategoriat täydentävät GHG Protocol Scope 3 -raportointia. Syötä 0 jos ei koske yritystäsi.")

    st.subheader("Työntekijöiden pendelöinti")
    st.caption("Syötä kaikkien työntekijöiden yhteenlasketut vuotuiset pendelöintikilometrit")
    col1, col2, col3 = st.columns(3)
    with col1:
        commute_car = st.number_input("Henkilöauto (km/vuosi)", 0, 10_000_000,
                                       int(st.session_state.data.get("commute_car", 0)),
                                       help="Kaikki työntekijät yhteensä")
    with col2:
        commute_pt = st.number_input("Joukkoliikenne (km/vuosi)", 0, 10_000_000,
                                      int(st.session_state.data.get("commute_pt", 0)))
    with col3:
        commute_ev = st.number_input("Sähköauto (km/vuosi)", 0, 10_000_000,
                                      int(st.session_state.data.get("commute_ev", 0)))

    st.subheader("Liikematkat & majoitus")
    st.caption("Yrityksen liikematkustaminen (erilliset työmatkaliikennelennot kirjataan Liikkuminen-sivulla)")
    col1, col2, col3 = st.columns(3)
    with col1:
        biztravel_car = st.number_input("Liikematka auto (km/vuosi)", 0, 5_000_000,
                                         int(st.session_state.data.get("biztravel_car", 0)))
    with col2:
        biztravel_taxi = st.number_input("Taksimatkat (km/vuosi)", 0, 1_000_000,
                                          int(st.session_state.data.get("biztravel_taxi", 0)))
    with col3:
        hotel_nights = st.number_input("Hotelliöt (yötä/vuosi)", 0, 10_000,
                                        int(st.session_state.data.get("hotel_nights", 0)),
                                        help="8.0 kg CO2e/yö – Cornell Hotel Sustainability Benchmark 2023")

    st.subheader("Pääomahyödykkeet (laitteet)")
    st.caption("Uusien laitteiden valmistuksen hiilijalanjälki (LCA). Jaa luku käyttövuosilla jos haluat vuosiarvion.")
    col1, col2 = st.columns(2)
    with col1:
        it_devices = st.number_input("IT-laitteet (kpl/vuosi)", 0, 10_000,
                                      int(st.session_state.data.get("it_devices", 0)),
                                      help="~300 kg CO2e/kpl koko elinkaari (Dell/Apple LCA). Syötä vuosittain hankittava määrä.")
    with col2:
        office_devices = st.number_input("Toimistolaitteet, näytöt tms. (kpl/vuosi)", 0, 10_000,
                                          int(st.session_state.data.get("office_devices", 0)),
                                          help="~150 kg CO2e/kpl koko elinkaari. Syötä vuosittain hankittava määrä.")

    st.subheader("Ostetut tavarat & palvelut")
    st.caption("Karkea arvio toimitusketjun päästöistä spend-pohjaisen EEIO-kertoimen avulla (0.50 kg CO2e/€)")
    purchased_goods_eur = st.number_input(
        "Hankintamenot (€/vuosi)", 0, 100_000_000,
        int(st.session_state.data.get("purchased_goods_eur", 0)),
        help="Exiobase / EEIO toimialan keskikerroin. Tarkempi laskenta vaatii toimittajakohtaiset tiedot."
    )

    # Näytetään reaaliaikainen esikatselutulos
    s3_total = (
        commute_car    * EMISSION_FACTORS["commute_car"] +
        commute_pt     * EMISSION_FACTORS["commute_pt"] +
        commute_ev     * EMISSION_FACTORS["commute_ev"] +
        biztravel_car  * EMISSION_FACTORS["biztravel_car"] +
        biztravel_taxi * EMISSION_FACTORS["biztravel_taxi"] +
        hotel_nights   * EMISSION_FACTORS["hotel"] +
        it_devices     * EMISSION_FACTORS["it_equipment"] +
        office_devices * EMISSION_FACTORS["office_equip"] +
        purchased_goods_eur * EMISSION_FACTORS["purchased_goods"]
    )
    if s3_total > 0:
        st.success(f"Tämän sivun Scope 3 -päästöt yhteensä: **{s3_total/1000:.2f} t CO2e/vuosi**")

    st.session_state.data.update({
        "commute_car": commute_car, "commute_pt": commute_pt, "commute_ev": commute_ev,
        "biztravel_car": biztravel_car, "biztravel_taxi": biztravel_taxi,
        "hotel_nights": hotel_nights, "it_devices": it_devices,
        "office_devices": office_devices, "purchased_goods_eur": purchased_goods_eur,
    })

    col1, col2 = st.columns(2)
    col1.button("<- Edellinen", on_click=prev_page)
    col2.button("Seuraava ->", on_click=next_page)

elif current_page == "Yhteenveto":
    st.header("Yhteenveto")

    electric_factor = st.session_state.data.get("electric_factor", DEFAULT_ELECTRIC_FACTOR)
    employees       = st.session_state.employees
    total_base, emissions_base = calculate_emissions(st.session_state.data, electric_factor)

    SCENARIOS = {
        "Nykytilanne":                    "Ei muutoksia – näytetään syötetyt luvut sellaisenaan",
        "✈️  50% vähemmän lentoa":         "Lentomatkat puolittuvat (videokokoukset, juna)",
        "🚗  Sähköauto koko kalusto":      "Kaikki bensa- ja dieselautot vaihdetaan sähköautoihin",
        "⚡  Uusiutuva sähkö":             "Sähkösopimus sertifioituun uusiutuvaan (0.010 kg CO2e/kWh)",
        "🏠  Vaihda lämmitysmuoto":        "Lämmitys vaihdetaan maalämpöön (0.020 kg CO2e/kWh)",
        "🌿  Kaukolämpö hiilineutraaliksi":"Kaukolämpö siirtyy bio/lämpöpumppuun – kerroin 0.040",
        "🗑️  Kierrätys maksimoitu":        "Sekajäte −70%, paperi ja muovi kierrätykseen",
        "🚌  Pendelöinti joukkoliikenteeseen": "Kaikki autopendelöinti siirtyy julkiseen liikenteeseen",
        "🏨  Vihreät hotellit":            "Hotelliöiden päästöt −50% (Green Key / ISO 14001 -sertifiointi)",
        "💻  Pidennetty laitekierto":      "IT- ja toimistolaitteiden hankinta −50% (pidempi elinkaari)",
        "🛒  Vihreät hankinnat":           "Ostettujen tavaroiden/palveluiden kerroin −30% (vastuulliset toimittajat)",
        "🏆  Kaikki yhdessä":             "Kaikki yllä olevat toimet yhdistettynä – maksimipotentiaali",
    }

    scenario = st.selectbox(
        "Valitse skenaario",
        list(SCENARIOS.keys()),
        format_func=lambda x: x
    )
    st.caption(f"ℹ️ {SCENARIOS[scenario]}")

    emissions_s = emissions_base.copy()

    if scenario == "✈️  50% vähemmän lentoa":
        emissions_s["Lennot"] *= 0.5

    elif scenario == "🚗  Sähköauto koko kalusto":
        km = st.session_state.data.get("car_bensa", 0) + st.session_state.data.get("car_diesel", 0)
        emissions_s["Bensa-auto"]  = 0
        emissions_s["Diesel-auto"] = 0
        emissions_s["Sähköauto"]   = (st.session_state.data.get("car_ev", 0) + km) * EMISSION_FACTORS["car_ev"]

    elif scenario == "⚡  Uusiutuva sähkö":
        emissions_s["Sähkö"]            = st.session_state.data.get("electricity", 0) * 0.010
        emissions_s["Sähkölämmitys"]    = emissions_s.get("Sähkölämmitys", 0) * (0.010 / 0.060)

    elif scenario == "🏠  Vaihda lämmitysmuoto":
        kwh = st.session_state.data.get("heating", 0)
        emissions_s["Lämmitys"] = kwh * HEATING_FACTORS["Maalämpöpumppu"]  # 0.020

    elif scenario == "🌿  Kaukolämpö hiilineutraaliksi":
        kwh = st.session_state.data.get("heating", 0)
        emissions_s["Lämmitys"] = kwh * 0.040  # bio/lämpöpumppu-mix tavoite 2030

    elif scenario == "🗑️  Kierrätys maksimoitu":
        waste_kg    = st.session_state.data.get("waste", 0)
        paper_kg    = st.session_state.data.get("paper", 0)
        plastic_kg  = st.session_state.data.get("plastic", 0)
        emissions_s["Sekajäte"]  = waste_kg * 0.3 * EMISSION_FACTORS["waste"]   # 70% pois sekajätteestä
        emissions_s["Kierrätys"] = (waste_kg * 0.7 + paper_kg + plastic_kg) * EMISSION_FACTORS["recycling"]
        emissions_s["Paperi"]    = 0
        emissions_s["Muovi"]     = 0

    elif scenario == "🚌  Pendelöinti joukkoliikenteeseen":
        commute_km = st.session_state.data.get("commute_car", 0)
        emissions_s["Pendelöinti (auto)"]     = 0
        emissions_s["Pendelöinti (julkinen)"] = (
            emissions_s.get("Pendelöinti (julkinen)", 0)
            + commute_km * EMISSION_FACTORS["commute_pt"]
        )

    elif scenario == "🏨  Vihreät hotellit":
        emissions_s["Hotelliöt"] = emissions_s.get("Hotelliöt", 0) * 0.5

    elif scenario == "💻  Pidennetty laitekierto":
        emissions_s["IT-laitteet"]       = emissions_s.get("IT-laitteet", 0) * 0.5
        emissions_s["Toimistolaitteet"]  = emissions_s.get("Toimistolaitteet", 0) * 0.5

    elif scenario == "🛒  Vihreät hankinnat":
        emissions_s["Ostetut tavarat/palvelut"] = emissions_s.get("Ostetut tavarat/palvelut", 0) * 0.7

    elif scenario == "🏆  Kaikki yhdessä":
        # Energia
        emissions_s["Sähkö"]    = st.session_state.data.get("electricity", 0) * 0.010
        emissions_s["Lämmitys"] = st.session_state.data.get("heating", 0) * HEATING_FACTORS["Maalämpöpumppu"]
        # Liikenne
        emissions_s["Lennot"] *= 0.5
        km = st.session_state.data.get("car_bensa", 0) + st.session_state.data.get("car_diesel", 0)
        emissions_s["Bensa-auto"]  = 0
        emissions_s["Diesel-auto"] = 0
        emissions_s["Sähköauto"]   = (st.session_state.data.get("car_ev", 0) + km) * EMISSION_FACTORS["car_ev"]
        # Jätteet
        waste_kg   = st.session_state.data.get("waste", 0)
        paper_kg   = st.session_state.data.get("paper", 0)
        plastic_kg = st.session_state.data.get("plastic", 0)
        emissions_s["Sekajäte"]  = waste_kg * 0.3 * EMISSION_FACTORS["waste"]
        emissions_s["Kierrätys"] = (waste_kg * 0.7 + paper_kg + plastic_kg) * EMISSION_FACTORS["recycling"]
        emissions_s["Paperi"]    = 0
        emissions_s["Muovi"]     = 0
        # Scope 3
        commute_km = st.session_state.data.get("commute_car", 0)
        emissions_s["Pendelöinti (auto)"]     = 0
        emissions_s["Pendelöinti (julkinen)"] = (
            emissions_s.get("Pendelöinti (julkinen)", 0)
            + commute_km * EMISSION_FACTORS["commute_pt"]
        )
        emissions_s["Hotelliöt"]                = emissions_s.get("Hotelliöt", 0) * 0.5
        emissions_s["IT-laitteet"]              = emissions_s.get("IT-laitteet", 0) * 0.5
        emissions_s["Toimistolaitteet"]         = emissions_s.get("Toimistolaitteet", 0) * 0.5
        emissions_s["Ostetut tavarat/palvelut"] = emissions_s.get("Ostetut tavarat/palvelut", 0) * 0.7

    total      = sum(emissions_s.values())
    per_person = total / max(employees, 1)
    saving     = total_base - total

    trend_df    = get_finland_emission_trend()
    temperature = get_current_temperature()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Kokonaispäästöt", f"{total/1000:.2f} t CO2e")
    c2.metric("Per henkilö", f"{per_person:,.0f} kg CO2e")
    if trend_df is not None:
        c3.metric("Suomen ka./hlö", f"{trend_df['kg_per_person'].iloc[-1]:,.0f} kg CO2e")
    if temperature is not None:
        c4.metric("Helsinki nyt", f"{temperature} °C")

    if saving > 0 and scenario != "Nykytilanne":
        pct = saving / total_base * 100 if total_base > 0 else 0
        st.success(f"💚 Skenaario säästää **{saving/1000:.2f} t CO2e/vuosi** ({pct:.1f}% vähennys nykytilasta)")

    st.subheader("Päästöjakauma")
    pie_data = {k: v for k, v in emissions_s.items() if v > 0}
    total_pie = sum(pie_data.values())
    fig_pie = px.pie(
        names=list(pie_data.keys()),
        values=list(pie_data.values()),
        hole=0.32
    )
    fig_pie.update_traces(
        textposition="inside",
        textinfo="percent",
        insidetextorientation="radial",
        textfont=dict(size=12, color="white"),
        # Piilota teksti alle 3% siivuista – ne näkyvät legendassa
        texttemplate="%{percent:.0%}",
    )
    fig_pie.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="white", size=13),
        # Legenda vasemmalle, piiras oikealle – ei päällekkäisyyttä
        legend=dict(
            font=dict(color="white", size=13),
            bgcolor="rgba(20,20,20,0.75)",
            bordercolor="rgba(255,255,255,0.2)",
            borderwidth=1,
            orientation="v",
            x=1.02,
            xanchor="left",
            y=0.5,
            yanchor="middle",
            itemsizing="constant",
            tracegroupgap=4,
        ),
        margin=dict(l=10, r=220, t=20, b=20),
        height=420,
    )
    st.plotly_chart(fig_pie, use_container_width=True)

    st.subheader("Nykytilanne vs. skenaario")
    fig_bar = go.Figure(data=[
        go.Bar(name="Nykytilanne", x=list(emissions_base.keys()),
               y=list(emissions_base.values()), marker_color="#ef476f"),
        go.Bar(name=scenario,      x=list(emissions_s.keys()),
               y=list(emissions_s.values()),    marker_color="#06d6a0"),
    ])
    fig_bar.update_layout(
        barmode="group",
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font_color="white",
        xaxis=dict(tickfont=dict(color="white")),
        yaxis=dict(tickfont=dict(color="white"), title="kg CO2e"),
        legend=dict(font=dict(color="white"))
    )
    st.plotly_chart(fig_bar, use_container_width=True)

    if trend_df is not None:
        st.subheader("Suomen kokonaispäästöt (Tilastokeskus)")
        fig_line = px.line(trend_df, x="Vuosi", y="kg_per_person",
                            labels={"kg_per_person": "kg CO2e / hlö"})
        fig_line.update_layout(paper_bgcolor="rgba(0,0,0,0)",
                                plot_bgcolor="rgba(0,0,0,0)", font_color="white")
        st.plotly_chart(fig_line, use_container_width=True)

    st.subheader("Säästösuositukset")
    tips = suggest_savings(emissions_s, employees)
    if tips:
        for t in tips:
            st.write("•", t)
    else:
        st.success("Erittäin matala päästötaso – yrityksen toimet ovat tehokkaita!")

    st.subheader("Lataa raportti")
    pdf_buf = generate_pdf_report(
        st.session_state.data, emissions_s, total,
        employees, st.session_state.company_name,
        st.session_state.industry, scenario
    )
    st.download_button(
        label="Lataa PDF-raportti",
        data=pdf_buf,
        file_name=f"sustaina_hiilijalanjälki_{datetime.now().strftime('%Y%m%d')}.pdf",
        mime="application/pdf"
    )

    col1, col2 = st.columns(2)
    col1.button("<- Edellinen", on_click=prev_page)
    col2.button("Aloita alusta", on_click=reset_app)

    st.markdown("""
<div class="sources-box">
<h4 style="color:white; margin-top:0; font-size:14px;">Lähteet ja menetelmät</h4>

<b style="color:#52b788">Energia &amp; lämmitys</b><br>
&bull; Sähkön päästökerroin (Suomi): <a href="https://www.motiva.fi/en/databank/energy-use-in-finland/co2-emission-factors/" style="color:#74c69d">Motiva.fi</a> / Tilastokeskus – 3v ka. 2021–2023 = 60 g CO2/kWh<br>
&bull; Lämmityksen kertoimet: <a href="https://www.motiva.fi/ratkaisut/energiankaytto_suomessa/co2-paastokertoimet" style="color:#74c69d">Motiva.fi</a> / Energiateollisuus ry – Kaukolämpötilasto 2024<br>
&bull; Polttoaineiden kertoimet: <a href="https://stat.fi" style="color:#74c69d">Tilastokeskus</a> – Polttoaineluokitus<br><br>

<b style="color:#52b788">Liikenne</b><br>
&bull; Ajoneuvojen ja joukkoliikenteen kertoimet: <a href="https://www.eea.europa.eu/publications/emep-eea-guidebook-2023" style="color:#74c69d">EMEP/EEA Guidebook 2023</a><br>
&bull; Lennot (sis. RFI 1.9): <a href="https://www.icao.int/environmental-protection/Carbonoffset" style="color:#74c69d">ICAO</a> / <a href="https://www.gov.uk/government/publications/greenhouse-gas-reporting-conversion-factors-2024" style="color:#74c69d">UK DESNZ GHG Conversion Factors 2024</a><br>
&bull; <em>Huom: VTT LIPASTO poistunut käytöstä 2023.</em> Korvaaja: <a href="https://stat.fi/meta/tilastoinnin-kehittaminen/liike-paastolaskentajarjestelma.html" style="color:#74c69d">Tilastokeskus LIIKE</a><br><br>

<b style="color:#52b788">Jätteet</b><br>
&bull; <a href="https://www.syke.fi" style="color:#74c69d">SYKE</a> – Jätteiden kasvihuonekaasupäästöt 2022; ecoinvent v3.9 (muovi, paperi)<br><br>

<b style="color:#52b788">Scope 3 – lisäkategoriat</b><br>
&bull; Pendelöinti & liikematkat: EMEP/EEA Guidebook 2023 (sama kerroin kuin ajoneuvoliikenne)<br>
&bull; Hotelliöt: <a href="https://ecomedes.com/blog/cornell-hotel-sustainability-benchmarking" style="color:#74c69d">Cornell Hotel Sustainability Benchmarking Index 2023</a> – 8.0 kg CO2e/yö<br>
&bull; IT-laitteet: Dell/Apple Product LCA -raportit 2022–2023 (~300 kg CO2e/kannettava)<br>
&bull; Ostetut tavarat/palvelut: <a href="https://www.exiobase.eu" style="color:#74c69d">Exiobase v3.8</a> / EEIO spend-pohjainen kerroin 0.50 kg CO2e/€ (toimialan ka.)<br><br>

<b style="color:#52b788">Tilastot &amp; kehikko</b><br>
&bull; Suomen kokonaispäästöt: <a href="https://stat.fi" style="color:#74c69d">Tilastokeskus</a> – statfin_khki_pxt_11ic.px<br>
&bull; Säätiedot: <a href="https://opendata.fmi.fi" style="color:#74c69d">Ilmatieteen laitos FMI Open Data</a><br>
&bull; Laskentakehikko: <a href="https://ghgprotocol.org/corporate-standard" style="color:#74c69d">GHG Protocol Corporate Standard</a><br><br>

<small style="color:#aaa">
Laskuri antaa suuntaa-antavan arvion. Virallinen GHG-raportointi tehdään toimialakohtaisilla primäärikerroimilla
(ISO 14064 / GHG Protocol). Päästökertoimet päivitetään vuosittain.
</small>
</div>
""", unsafe_allow_html=True)