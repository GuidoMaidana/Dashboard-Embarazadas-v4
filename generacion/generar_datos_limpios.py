# -*- coding: utf-8 -*-
"""
Genera el CSV final optimizado a partir de la carpeta "Excel datos".

Salida (en la raiz de este proyecto, fuera de "Excel datos"):
  - datos_salud_optimizados.csv   (consultas + centros en una sola tabla; lo genera este script)
  - LIMPIEZA_DATOS.md             (informe de limpieza y trazabilidad; documento aparte)

Fuentes:
  - Consultas_embarazadas.xlsx            (consultas mensuales agregadas)
  - datos_salud_combinados.xlsx           (SAPS completo: 37 registros; zonas)
  - Saps.xlsx                             (SAPS incompleto: 34; descartado)
  - Caps.xlsx                             (14 CAPS)
  - Hospitales.xlsx                       (8 hospitales)
  - centros de salud.xlsx                 (40 centros privados)
  - centros de salud (1).xlsx             (copia identica: descartada)
"""
from __future__ import annotations

import csv
import hashlib
import json
import math
import re
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "Excel datos"

# --------------------------------------------------------------------------
# 1. Nombres canonicos de SAPS (por gid de la fuente de establecimientos)
# --------------------------------------------------------------------------
SAPS_CANON = {
    1: "SAP XXI - Cañada Quiroz",
    2: "SAP XII - Bº Esperanza",
    3: "SAP XVII - Dr. Benjamín Serrano (Bº Parque Cadenas Norte)",
    4: "SAP XXVIII - Dr. José De Los Reyes Vidal",
    5: "SAPS XXVII - Dr. Juan Carlos Benítez Suárez (SAPS Lomas)",
    6: "SAP I - Dr. Guillermo Rawson",
    7: "SAP XXIII - Dr. Juan Chercoff",
    8: "SAP XV - Dr. Juan Carlos Semper",
    9: "SAP XIV - Dr. Manuel A. González",
    10: "SAP III - Dr. Oscar E. Pirchi (Bº Anahí)",
    11: "SAP VI - Bº Pujol",
    12: "SAP V - Bº Quinta Ferré",
    13: "SAP IV - Bº Itatí",
    14: "SAP XXXV - Dr. Romilio Monzón (Bº Cichero)",
    15: "SAP II - Bº Villa García",
    16: "SAP XI - Bº Villa Chiquita",
    17: "SAP XXXVII - Bº Güemes",
    18: "SAP VII - Bº San Marcos",
    19: "SAP XIX - Bº Primera Junta",
    20: "SAP X - Dr. Tomás Korimblum",
    21: "SAP XX - Dr. Aníbal Malvido",
    22: "SAP XXVI - Dr. Santiago Lorenzo",
    23: "SAP XXV - Dr. Alberto Lifschitz",
    24: "SAP IX - Dr. Manuel Casusso",
    25: "SAP XXXI - Bº Río Paraná",
    26: "SAP VIII - Dr. Miguel Rossi Candia",
    27: "SAPS XXIX - Bº Santa Marta",
    28: "SAP XXX - Dr. Saúl Calmanash",
    29: "SAP XXXIII - Dr. Pedro Blugerman",
    30: "SAPS XXXIV - Dr. Elías Galván",
    31: "SAP XXXII - Dr. Miguel Sussini",  # numeral corregido (fuente: XXXIII duplicado)
    32: "SAP XXXVI - Centro de Vacunación",
    33: "SAPS XXIV - Dr. Mauricio Open",
    34: "SAP XIII - Bº Santa Margarita",
    35: "SAPS XVIII - Bº Sapucay",
    37: "SAPS XVI - Bº Paso Martínez",
    38: "SAPS Santa Catalina",
}

SAPS_BARRIO = {
    1: "Cañada Quiroz", 2: "Esperanza", 3: "Parque Cadenas", 4: "Santa Rita",
    5: "Lomas", 6: "San Gerónimo", 7: "Yecohá", 8: "Víctor Colas",
    9: "Lomas del Mirador", 10: "Anahí", 11: "Pujol", 12: "Quinta Ferré",
    13: "Itatí", 14: "Cichero", 15: "Villa García", 16: "Villa Chiquita",
    17: "Güemes", 18: "San Marcos", 19: "Primera Junta", 20: "Ongay",
    21: "Pirayuí", 22: "Barrio Nuevo", 23: "San Antonio Este", 24: "San Antonio",
    25: "Río Paraná", 26: "Independencia", 27: "3 de Abril", 28: "Quilmes",
    29: "Virgen de los Dolores", 30: "Galván", 31: "Pedro Ferré",
    32: "Deportes", 33: "Paraje Perichón (Aeropuerto)", 34: "Santa Margarita",
    35: "Sapucay", 37: "Paso Martínez", 38: "Santa Catalina",
}

SAPS_DIR = {
    1: "R.P. N° 3 Km",
    2: "114 Viviendas - Mz. B - Calle 2, entre 4 y 5",
    3: "Ruta 5 - Parque Cadenas Norte S/N",
    4: "Barrio Santa Rita - Calle Pérez Rueda S/N",
    5: "Bº Flier - Mza. 17 Casa 2",
    6: "Av. Pdte. Dr. Raúl Alfonsín y Av. Laprida Narciso",
    7: "Ruta 12 - Km 10",
    8: "Argerich y Calle N° 41",
    9: "Sicilia y Linares",
    10: "Cristo Obrero N° 310",
    11: "Av. Armenia y Yerbal",
    12: "Zacarías Sánchez N° 300",
    13: "Amado Bonpland N° 100",
    14: "Spano entre Vélez Sarsfield y Sgto. Del Estero",
    15: "Pje. París N° 4000",
    16: "Pirovano N° 1988",
    17: "Gascón entre Las Heras y Lamadrid",
    18: "Roque Pérez S/N",
    19: "Ntra. Sra. de la Asunción y Ex Vías",
    20: "Güemes y Berazategui",
    21: "Suecia S/N",
    22: "Ramos Mejía y Túpac Amaru",
    23: "Roldán Belizario y Calle N° 482",
    24: "Av. Alta Gracia y Los Tobas",
    25: "Mocoretá S/N",
    26: "Dr. L. F. Leloir y Calle N° 459",
    27: "Av. José Vidal S/N",
    28: "Ortiz de Zárate N° 3400",
    29: "Montes de Oca S/N, casi Gutnisky",
    30: "Av. Alberdi y Pizarro",
    31: "3 de Abril y Chaco",
    32: "Av. Costanera Gral. San Martín y Salta",
    33: "Ruta 12 Km 8 (Control Policial)",
    34: "Barrio Santa Margarita (frente a la entrada del puente Pexoa)",
    35: "Barrio Sapucay - Calle N° 564 y Calle N° 563",
    37: "Barrio Paso Martínez",
    38: "Abelardo Héctor Dimotta y Ramón Sixto Ríos",
}

NUMERO_SAPS = {
    1: "XXI", 2: "XII", 3: "XVII", 4: "XXVIII", 5: "XXVII", 6: "I",
    7: "XXIII", 8: "XV", 9: "XIV", 10: "III", 11: "VI", 12: "V",
    13: "IV", 14: "XXXV", 15: "II", 16: "XI", 17: "XXXVII", 18: "VII",
    19: "XIX", 20: "X", 21: "XX", 22: "XXVI", 23: "XXV", 24: "IX",
    25: "XXXI", 26: "VIII", 27: "XXIX", 28: "XXX", 29: "XXXIII",
    30: "XXXIV", 31: "XXXII", 32: "XXXVI", 33: "XXIV", 34: "XIII",
    35: "XVIII", 37: "XVI", 38: "",
}

# Alias de la base de consultas -> gid del catalogo de SAPS
CONSULTAS_ALIAS = {
    "CICHERO": 14,
    "DR MANUEL A. GONZALEZ": 9,
    "DR. ALBERTO LIFSCHITZ": 23,
    "DR. ANIBAL MALVIDO": 21,
    "DR. BENJAMIN SERRANO": 3,
    "DR. BLUGERMAN": 29,
    "DR. CALMANASH": 28,
    "DR. CHERCOFF": 7,
    "DR. FLIER": 5,
    "DR. GUILLERMO RAWSON": 6,
    "DR. KORIMBLUM": 20,
    "DR. MANUEL CASSUSO": 24,
    "DR. MAURICIO OPEN": 33,
    "DR. PEDRO BLUGERMAN": 29,
    "DR. PIRCHI": 10,
    "DR. ROSSI CANDIA": 26,
    "DR. SANTIAGO LORENZO": 22,
    "DR. SEMPER": 8,
    "DR. SUSSINI": 31,
    "ESPERANZA": 2,
    "GDOR. ELIAS GALVAN": 30,
    "GÜEMES": 17,
    "ITATI": 13,
    "PRIMERA JUNTA": 19,
    "QUINTA FERRE": 12,
    "RIO PARANA": 25,
    "SAN MARCOS": 18,
    "SANTA MARGARITA": 34,
    "SANTA MARTA": 27,
    "SAPUCAY": 35,
    "VILLA CHIQUITA": 16,
}
# lookup insensible a mayusculas (el origen mezcla 'DR. CHERCOFF' y 'Dr. CALMANASH')
CONSULTAS_ALIAS_UP = {k.upper(): v for k, v in CONSULTAS_ALIAS.items()}

SAPS_OBS = {
    3: "En consultas figura como 'DR. BENJAMIN SERRANO'; mismo domicilio (Ruta 5 km 2,5, Bº Parque Cadenas).",
    5: "En consultas figura como 'DR. FLIER' (SAPS Lomas, Mancini y Juan Pablo II).",
    10: "En consultas figura como 'DR. PIRCHI' (Cristo Obrero 310, Bº Anahí).",
    24: "En consultas figura como 'DR. MANUEL CASSUSO'; el catálogo usa 'Dr. M. Casusso'.",
    26: "En consultas figura como 'DR. ROSSI CANDIA'.",
    27: "En consultas figura como 'SANTA MARTA'.",
    29: "Se unificaron 'DR. BLUGERMAN' y 'DR. PEDRO BLUGERMAN': mismo establecimiento (antes partido en dos).",
    31: "Numeral corregido: la fuente repetía 'XXXIII' (también usado por Blugerman); por secuencia y gid corresponde a SAP XXXII.",
    32: "Sin registros de consultas en ene-sep 2020.",
    38: "Sin número romano en la fuente.",
}
# SAPS sin consultas en el periodo
SAPS_SIN_CONSULTAS = {1, 11, 15, 32, 37, 4, 38}

CAPS_CANON = {
    "CAPS I - B San José": "CAPS I - Barrio San José",
    "CAPS II - Dr. Emilio Coni": "CAPS II - Dr. Emilio Coni",
    "CAPS III": "CAPS III",
    "CAPS IV - Dr. Graciano Lugo": "CAPS IV - Dr. Graciano Lugo",
    "CAPS V - B Dr. Montana": "CAPS V - Barrio Dr. Montaña",
    "CAPS VI - Dr. Tiburcio Foncesa": "CAPS VI - Dr. Tiburcio Fonseca",
    "CAPS VII - B Laguna Brava": "CAPS VII - Barrio Laguna Brava",
    "CAPS VIII - B Santa Teresita": "CAPS VIII - Barrio Santa Teresita",
    "CAPS IX - B 17 de Agosto": "CAPS IX - Barrio 17 de Agosto",
    "CAPS X - B Molina Punta": "CAPS X - Barrio Molina Punta",
    "CAPS XI - B Quintana": "CAPS XI - Barrio Quintana",
    "CAPS XII - B Ponce": "CAPS XII - Barrio Ponce",
    "CAPS XIII": "CAPS XIII",
    "CAPS XIV": "CAPS XIV",
}
CAPS_BARRIO = {
    "CAPS I - Barrio San José": "San José",
    "CAPS V - Barrio Dr. Montaña": "Dr. Montaña",
    "CAPS VII - Barrio Laguna Brava": "Laguna Brava",
    "CAPS VIII - Barrio Santa Teresita": "Santa Teresita",
    "CAPS IX - Barrio 17 de Agosto": "17 de Agosto",
    "CAPS X - Barrio Molina Punta": "Molina Punta",
    "CAPS XI - Barrio Quintana": "Quintana",
    "CAPS XII - Barrio Ponce": "Ponce",
}
HOSP_CANON = {
    "HOSPITAL PSIQUIATRICO": "Hospital Psiquiátrico",
    "HOSPITAL NEONATAL": "Hospital Neonatal",
    "HOSPITAL VIDAL": "Hospital Vidal",
    "HOSPITAL GERIATRICO": "Hospital Geriátrico",
    "HOSPITAL ESCUELA": "Hospital Escuela",
    "HOSPITAL PEDIATRICO JUAN PABLO II": "Hospital Pediátrico Juan Pablo II",
    "HOSPITAL LLANO": "Hospital Llano",
    "HOSPITAL DE CAMPAÑA": "Hospital de Campaña",
}
PRIV_CANON = {
    "CLINICA MODELO": "Clínica Modelo",
    "CLINICA MATERNAL IBERA": "Clínica Maternal Iberá",
    "CENTRO MEDICO": "Centro Médico",
    "CLINICA MAYO S.A": "Clínica Mayo S.A.",
    "INSTITUTO DE LA MUJER": "Instituto de la Mujer",
    "INSTITUTO DE DIAGNOSTICO Y TRATAMIENTO S.A.": "Instituto de Diagnóstico y Tratamiento S.A.",
    "CENTRO RENAL REVIDATTI": "Centro Renal Revidatti",
    "CLINICA DE CIRUJIA AMBULATORIA CENTRO LASER REG.": "Clínica de Cirugía Ambulatoria Centro Láser Reg.",
    "CLINICA SAGRADO CORAZON": "Clínica Sagrado Corazón",
    "CLINICA CORRIENTES": "Clínica Corrientes",
    "CLINICA DRES. ESTRAGO": "Clínica Dres. Estrago",
    "CLINICA DE CIRUJIA MINI INVASIVA S.A.": "Clínica de Cirugía Mini Invasiva S.A.",
    "UNIDAD RENAL CORRIENTES": "Unidad Renal Corrientes",
    "SANATORIO SAN JUAN": "Sanatorio San Juan",
    "SANATORIO DEL NORTE": "Sanatorio del Norte",
    "CARDIOCENTRO": "Cardiocentro",
    "CLINICA DEL SOL": "Clínica del Sol",
    "INSTITUTO DE ORTOPEDIA Y TRAUMATOLOGIA": "Instituto de Ortopedia y Traumatología",
    "CENTRO NEUROLOGICO COMPUTARIZADO": "Centro Neurológico Computarizado",
    "CENTRO MEDICO LABORAL CERKIDE": "Centro Médico Laboral Cerkide",
    "INSTITUTO DE CARDIOLOGIA DE CORRIENTES": "Instituto de Cardiología de Corrientes",
    "CONSTRUIR SALUD": "Construir Salud",
    "CENTRO DE COMPUTACION MEDICA DEL NORDESTE": "Centro de Computación Médica del Nordeste",
    "CEREBRO INSTITUTO NEUROPSIQUIATRICO": "Cerebro Instituto Neuropsiquiátrico",
    "CENTRO RESPIRATORIO CORRIENTES": "Centro Respiratorio Corrientes",
    "PSICOSALUD S.R.L.": "Psicosalud S.R.L.",
    "INTITUTO DE FLEBOLOGIA DEL NORDESTE": "Instituto de Flebología del Nordeste",
    "INSTITUTO DE HEMOTERAPIA": "Instituto de Hemoterapia",
    "ICAL CORRIENTES": "ICAL Corrientes",
    "MEDINIÑOS": "Mediniños",
    "ASISTIR SERVICIOS MEDICOS": "Asistir Servicios Médicos",
    "INSTITUTO DE LA HIPERTENSION Y DEL CORAZON": "Instituto de la Hipertensión y del Corazón",
    "INSTITUTO PRIVADO DE ASMA, ALERGIA E INMUNOLOGIA": "Instituto Privado de Asma, Alergia e Inmunología",
    "CENTRO PSICONEUROCOGNITIVO SAN JOAQUIN": "Centro Psiconeurocognitivo San Joaquín",
    "INST. TEC. DENTAL SANTA ROSA": "Instituto Técnico Dental Santa Rosa",
    "INST. DE KINESIOLOGIA A.C.I.C.": "Instituto de Kinesiología A.C.I.C.",
    "INST.CORRENTINO DE RADIODIAGNOSTICO ODONT.": "Instituto Correntino de Radiodiagnóstico Odontológico",
    "EME Policonsultorios": "EME Policonsultorios",
}

# --------------------------------------------------------------------------
# 2. Normalizacion de direcciones (CATS, hospitales y centros privados)
# --------------------------------------------------------------------------
KEEP_UPPER = {
    "CAPS", "SAPS", "SAP", "EME", "ICAL", "S.A.", "S.R.L.", "A.C.I.C.",
    "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X",
    "XI", "XII", "XIII", "XIV", "XV", "XVI", "XVII", "XVIII", "XIX",
    "XX", "XXI", "XXII", "XXIII", "XXIV", "XXV", "XXVI", "XXVII",
    "XXVIII", "XXIX", "XXX", "XXXI", "XXXII", "XXXIII", "XXXIV",
    "XXXV", "XXXVI", "XXXVII", "S/N", "N°", "Bº",
}
CONNECTORS = {"de", "del", "la", "y", "e", "a", "al", "en", "entre",
              "casi", "esq."}
WORD_FIX = {
    "cordoba": "Córdoba", "martin": "Martín", "junin": "Junín",
    "bolivar": "Bolívar", "paragua": "Paraguay", "maipu": "Maipú",
    "velez": "Vélez", "sarfiel": "Sarsfield", "sarsfield": "Sarsfield",
    "raul": "Raúl", "alfonsin": "Alfonsín", "bompland": "Bonpland",
    "paris": "París", "gascon": "Gascón", "perez": "Pérez",
    "asuncion": "Asunción", "vias": "Vías", "mejia": "Mejía",
    "tupac": "Túpac", "amaru": "Amaru", "roldan": "Roldán",
    "zacarias": "Zacarías", "sanches": "Sánchez", "mocoreta": "Mocoretá",
    "zarate": "Zárate", "hector": "Héctor", "ramon": "Ramón",
    "rios": "Ríos", "martinez": "Martínez", "cosquin": "Cosquín",
    "paysandu": "Paysandú", "agustin": "Agustín", "britez": "Brítez",
    "pellegrino": "Pellegrini", "gutemberg": "Gutenberg",
    "escalada": "Escalada", "escala": "Escalada", "peron": "Perón",
    "ferreyra": "Ferreyra", "fernandez": "Fernández", "jose": "José",
    "pexoa": "Pexoa", "gutnisky": "Gutnisky", "costanera": "Costanera",
    "saenz": "Sáenz", "turin": "Turín", "joaquin": "Joaquín",
    "medica": "Médica", "neuropsiquiatrico": "Neuropsiquiátrico",
    "km": "Km", "r.p.": "R.P.", "j.d.": "J. D.", "l.f.": "L. F.",
    "gral": "Gral.", "pdte": "Pdte.", "pje": "Pje.", "mza": "Mza.",
    "sgto": "Sgto.", "sto": "Sto.", "ntra": "Ntra.", "sra": "Sra.",
    "av": "Av.", "av.": "Av.", "gdor": "Gdor.", "dr": "Dr.",
    "dr.": "Dr.", "dra": "Dra.", "dra.": "Dra.", "esq": "esq.",
}


def fix_text(text: str) -> str:
    """Normaliza mayusculas, tildes y abreviaturas de una direccion."""
    if text is None:
        return ""
    s = re.sub(r"\s+", " ", str(text)).strip()
    s = re.sub(r"\bN\s*[°º]\s*", "N° ", s)
    s = re.sub(r"\bN\s*[°º]", "N°", s)
    s = re.sub(r"\bS\s*/\s*N\b", "S/N", s, flags=re.I)
    s = re.sub(r"\bMz\b\.?", "Mz.", s, flags=re.I)
    s = re.sub(r"\bMza\b\.?", "Mza.", s, flags=re.I)
    tokens = []
    for i, tok in enumerate(s.split(" ")):
        if not tok:
            continue
        m = re.match(r"^([^\wÁÉÍÓÚÑáéíóúñ]*)([\wÁÉÍÓÚÑáéíóúñ'/\.]+)([^\wÁÉÍÓÚÑáéíóúñ]*)$", tok)
        if not m:
            tokens.append(tok)
            continue
        pre, core, post = m.groups()
        key = core.lower()
        if core.upper() in KEEP_UPPER:
            new = core.upper()
        elif key in WORD_FIX:
            new = WORD_FIX[key]
        elif re.search(r"\d", core):
            new = core
        elif i > 0 and key in CONNECTORS:
            new = key
        else:
            new = core[0].upper() + core[1:].lower()
        tokens.append(pre + new + post)
    s = " ".join(tokens)
    s = re.sub(r"\besq\.\s*", "esq. ", s)
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r"\bSto\. C\.?\s*$", "Sto. C.", s)
    return s


def load(fn: str, sheet: str):
    wb = openpyxl.load_workbook(DATA / fn, data_only=True)
    ws = wb[sheet]
    return list(ws.iter_rows(values_only=True))[1:]


# --------------------------------------------------------------------------
# 3. Zonas (point-in-polygon)
# --------------------------------------------------------------------------
def load_zones():
    rows = load("datos_salud_combinados.xlsx", "_select_gid_zona_type_polygon_c")
    zones = []
    for gid, zona, geom in rows:
        g = json.loads(geom)
        name = {"AREA I": "Área I", "AREA II": "Área II", "AREA III": "Área III"}.get(
            str(zona).strip().upper(), str(zona).strip().title())
        zones.append((name, g["coordinates"][0]))
    return zones


def point_in_poly(x, y, ring):
    inside = False
    n = len(ring)
    j = n - 1
    for i in range(n):
        xi, yi = ring[i]
        xj, yj = ring[j]
        if ((yi > y) != (yj > y)) and (x < (xj - xi) * (y - yi) / (yj - yi) + xi):
            inside = not inside
        j = i
    return inside


def find_zone(zones, lng, lat):
    hits = [name for name, ring in zones if point_in_poly(lng, lat, ring)]
    return hits[0] if hits else ""


ZONES = load_zones()


def zone_of(lng, lat):
    return find_zone(ZONES, float(lng), float(lat))


# --------------------------------------------------------------------------
# 4. Catalogo de centros
# --------------------------------------------------------------------------
def build_centers():
    centers = []  # dicts
    saps = load("datos_salud_combinados.xlsx", "_vw_saps__202510011235")
    caps = load("Caps.xlsx", "_vw_caps__202510011216")
    hosp = load("Hospitales.xlsx", "vw_hospitales_202510011225")
    priv = load("centros de salud.xlsx", "vw_centros_de_salud_20251001123")

    for gid, nombre, barrio, direccion, lng, lat in saps:
        gid = int(gid)
        centers.append({
            "tipo": "SAPS", "sector": "Público", "gid": gid,
            "id": None,  # se asigna luego
            "nombre": SAPS_CANON[gid],
            "original": str(nombre).strip(),
            "numero": NUMERO_SAPS[gid],
            "barrio": SAPS_BARRIO[gid],
            "direccion": SAPS_DIR[gid],
            "lat": round(float(lat), 7), "lng": round(float(lng), 7),
            "fuente": "datos_salud_combinados.xlsx / '_vw_saps__202510011235'",
            "obs": SAPS_OBS.get(gid, ""),
        })
    for gid, nombre, direccion, lng, lat in caps:
        canon = CAPS_CANON[str(nombre).strip()]
        centers.append({
            "tipo": "CAPS", "sector": "Público", "gid": int(gid), "id": None,
            "nombre": canon, "original": str(nombre).strip(), "numero": "",
            "barrio": CAPS_BARRIO.get(canon, ""),
            "direccion": fix_text(direccion),
            "lat": round(float(lat), 7), "lng": round(float(lng), 7),
            "fuente": "Caps.xlsx / '_vw_caps__202510011216'",
            "obs": "Nombre normalizado." if canon != str(nombre).strip() else "",
        })
    for gid, nombre, tipo, direccion, lng, lat in hosp:
        canon = HOSP_CANON[str(nombre).strip()]
        centers.append({
            "tipo": "Hospital público", "sector": "Público", "gid": int(gid), "id": None,
            "nombre": canon, "original": str(nombre).strip(), "numero": "",
            "barrio": "", "direccion": fix_text(direccion),
            "lat": round(float(lat), 7), "lng": round(float(lng), 7),
            "fuente": "Hospitales.xlsx / 'vw_hospitales_202510011225'",
            "obs": "Nombre normalizado." if canon != str(nombre).strip() else "",
        })
    for gid, nombre, direccion, lng, lat in priv:
        raw = str(nombre).strip()
        if raw == "CLINICA DEL NIÑO":
            canon = f"Clínica del Niño ({fix_text(direccion)})"
            obs = "Posible duplicado: mismo nombre en direcciones contiguas (24 m). Se conservan ambos registros."
        else:
            canon = PRIV_CANON[raw]
            obs = "Nombre normalizado." if canon != raw else ""
        centers.append({
            "tipo": "Centro privado", "sector": "Privado", "gid": int(gid), "id": None,
            "nombre": canon, "original": raw, "numero": "",
            "barrio": "", "direccion": fix_text(direccion),
            "lat": round(float(lat), 7), "lng": round(float(lng), 7),
            "fuente": "centros de salud.xlsx / 'vw_centros_de_salud_20251001123'",
            "obs": obs,
        })

    # ids secuenciales por tipo (ordenados por gid)
    counters = defaultdict(int)
    prefixes = {"SAPS": "SAPS", "CAPS": "CAPS",
                "Hospital público": "HOSP", "Centro privado": "PRIV"}
    for c in centers:
        counters[c["tipo"]] += 1
        c["id"] = f'{prefixes[c["tipo"]]}-{counters[c["tipo"]]:02d}'
    return centers


# --------------------------------------------------------------------------
# 5. Consultas limpias
# --------------------------------------------------------------------------
AGE_MAP = {
    "12 - 15 ANIOS": ("12-15 años", 12, 15),
    "16 - 20 ANIOS": ("16-20 años", 16, 20),
    "21 - 39 ANIOS": ("21-39 años", 21, 39),
    ">=40 ANIOS": ("40 años o más", 40, ""),
}
MESES = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]


def build_consultas():
    rows = load("Consultas_embarazadas.xlsx", "consultas_embarazadas")
    # 5.1 descartar filas sin cantidad o sin grupo etario (valores NULL)
    validas = [r for r in rows
               if r[3] not in (None, "", "NULL") and r[4] not in (None, "", "NULL")]
    vacias = len(rows) - len(validas)
    # 5.2 agrupar por (SAPS, periodo, grupo etario)
    groups = defaultdict(lambda: {"cantidades": [], "raw": set()})
    for saps_raw, fecha, servicio, cantidad, rango in validas:
        key_raw = str(saps_raw).strip()
        gid = CONSULTAS_ALIAS_UP[key_raw.upper()]
        date = datetime.strptime(str(fecha).strip(), "%d-%m-%Y")
        key = (gid, date.strftime("%Y-%m"), str(rango).strip())
        g = groups[key]
        g["cantidades"].append(float(cantidad))
        g["raw"].add(key_raw)
    out = []
    duplicadas = 0
    for (gid, periodo, rango_raw), g in groups.items():
        grupo, emin, emax = AGE_MAP[rango_raw]
        anio, mes = int(periodo[:4]), int(periodo[5:])
        if mes <= 2:
            etapa = "Pre-ASPO"
        elif mes <= 6:
            etapa = "ASPO"
        else:
            etapa = "Recuperación"
        cantidades = g["cantidades"]
        n = len(cantidades)
        obs = ""
        if n > 1 and len(set(cantidades)) == 1:
            # filas exactamente duplicadas en el origen: se conserva una sola
            cantidad = int(round(cantidades[0]))
            duplicadas += n - 1
            obs = (f"Fila duplicada en origen ({n} registros idénticos): "
                   f"se conservó un solo registro.")
        elif n > 1:
            # registros distintos para la misma celda: se agregan
            cantidad = int(round(sum(cantidades)))
            obs = (f"Se agregaron {n} registros distintos del mismo periodo "
                   f"y grupo etario.")
        else:
            cantidad = int(round(cantidades[0]))
        if gid == 29:
            obs = (obs + " " if obs else "") + \
                ("Nombre normalizado: 'DR. BLUGERMAN' y 'DR. PEDRO BLUGERMAN' "
                 "son el mismo SAPS (se unificaron).")
        out.append({
            "gid": gid, "periodo": periodo, "anio": anio, "mes": mes,
            "fecha": f"{periodo}-01", "mes_nombre": MESES[mes],
            "etapa": etapa, "grupo": grupo, "emin": emin, "emax": emax,
            "cantidad": cantidad,
            "filas": n,
            "raw": " / ".join(sorted(g["raw"])),
            "obs": obs,
        })
    out.sort(key=lambda r: (r["periodo"], r["gid"], r["emin"]))
    return out, len(validas), vacias, duplicadas


# --------------------------------------------------------------------------
# 6. Escritura del CSV final
# --------------------------------------------------------------------------
HEADERS = [
    "tipo_registro", "id_registro", "id_centro", "tipo_centro", "sector",
    "nombre_centro", "nombre_original", "numero_saps", "barrio", "zona",
    "direccion", "latitud", "longitud", "tiene_coordenadas",
    "datos_consultas_disponibles", "anio", "mes", "periodo",
    "fecha_inicio_periodo", "mes_nombre", "etapa_pandemia", "servicio",
    "grupo_etario", "edad_min", "edad_max", "cantidad_consultas_mes",
    "consultas_totales_2020", "meses_con_registros_2020", "fuente",
    "observaciones",
]


def build_rows(centers, consultas):
    by_gid = {c["gid"]: c for c in centers if c["tipo"] == "SAPS"}
    # totales por SAPS
    tot = defaultdict(int)
    meses = defaultdict(set)
    for r in consultas:
        tot[r["gid"]] += r["cantidad"]
        meses[r["gid"]].add(r["periodo"])

    rows = []
    for i, r in enumerate(consultas, 1):
        c = by_gid[r["gid"]]
        rows.append({
            "tipo_registro": "consulta", "id_registro": f"CONS-{i:04d}",
            "id_centro": c["id"], "tipo_centro": "SAPS", "sector": "Público",
            "nombre_centro": c["nombre"], "nombre_original": r["raw"],
            "numero_saps": c["numero"], "barrio": c["barrio"], "zona": c["zona"],
            "direccion": c["direccion"], "latitud": c["lat"], "longitud": c["lng"],
            "tiene_coordenadas": "SI" if c["lat"] and c["lng"] else "NO",
            "datos_consultas_disponibles": "SI",
            "anio": r["anio"], "mes": r["mes"], "periodo": r["periodo"],
            "fecha_inicio_periodo": r["fecha"], "mes_nombre": r["mes_nombre"],
            "etapa_pandemia": r["etapa"], "servicio": "Embarazadas del Área Programática",
            "grupo_etario": r["grupo"], "edad_min": r["emin"], "edad_max": r["emax"],
            "cantidad_consultas_mes": r["cantidad"],
            "consultas_totales_2020": "", "meses_con_registros_2020": "",
            "fuente": "Consultas_embarazadas.xlsx / 'consultas_embarazadas'",
            "observaciones": r["obs"],
        })
    order = {"SAPS": 0, "CAPS": 1, "Hospital público": 2, "Centro privado": 3}
    for c in sorted(centers, key=lambda x: (order[x["tipo"]], x["id"])):
        es_saps = c["tipo"] == "SAPS"
        obs = c["obs"]
        if es_saps and not tot.get(c["gid"]) and "Sin registros" not in obs:
            obs = (obs + " " if obs else "") + \
                "Sin registros de consultas en ene-sep 2020."
        rows.append({
            "tipo_registro": "centro", "id_registro": c["id"], "id_centro": c["id"],
            "tipo_centro": c["tipo"], "sector": c["sector"],
            "nombre_centro": c["nombre"], "nombre_original": c["original"],
            "numero_saps": c["numero"], "barrio": c["barrio"], "zona": c["zona"],
            "direccion": c["direccion"], "latitud": c["lat"], "longitud": c["lng"],
            "tiene_coordenadas": "SI" if c["lat"] and c["lng"] else "NO",
            "datos_consultas_disponibles": "SI" if es_saps else "NO",
            "anio": "", "mes": "", "periodo": "", "fecha_inicio_periodo": "",
            "mes_nombre": "", "etapa_pandemia": "", "servicio": "",
            "grupo_etario": "", "edad_min": "", "edad_max": "",
            "cantidad_consultas_mes": "",
            "consultas_totales_2020": tot.get(c["gid"], 0) if es_saps else "",
            "meses_con_registros_2020": len(meses.get(c["gid"], set())) if es_saps else "",
            "fuente": c["fuente"], "observaciones": obs,
        })
    return rows


def write_csv(rows, path):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=HEADERS, delimiter=";",
                           quoting=csv.QUOTE_MINIMAL, lineterminator="\r\n")
        w.writeheader()
        for r in rows:
            w.writerow(r)


def main():
    # zonas a cada centro
    centers = build_centers()
    for c in centers:
        c["zona"] = zone_of(c["lng"], c["lat"])
    consultas, filas_validas, filas_vacias, duplicadas = build_consultas()
    rows = build_rows(centers, consultas)
    out = ROOT / "datos_salud_optimizados.csv"
    write_csv(rows, out)

    total = sum(r["cantidad_consultas_mes"] for r in rows
                if r["tipo_registro"] == "consulta")
    n_cons = sum(1 for r in rows if r["tipo_registro"] == "consulta")
    n_cent = sum(1 for r in rows if r["tipo_registro"] == "centro")
    total_cent = sum(r["consultas_totales_2020"] for r in rows
                     if r["tipo_registro"] == "centro" and r["consultas_totales_2020"] != "")
    print(f"CSV : {out}")
    print(f"Consultas: {n_cons} filas | {total} consultas")
    print(f"Centros  : {n_cent} filas | {total_cent} consultas (SAPS)")
    print(f"Duplicados removidos: {duplicadas} | filas vacias removidas: {filas_vacias}")
    return total, n_cons, n_cent, total_cent, duplicadas, filas_vacias


if __name__ == "__main__":
    main()
