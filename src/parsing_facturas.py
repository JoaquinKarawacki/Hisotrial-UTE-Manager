"""parsing_facturas.py -- Extraccion y parseo de datos desde los PDF de
facturas UTE (tipos TH y GS). Funciones puras: reciben un pdf_path o texto
y devuelven dicts. No dependen de Playwright ni de la base de datos.

Separado de historial_facturas_ute.py (scraping) sin cambios de logica.
"""
from __future__ import annotations

import os
import re

import pdfplumber


# =========================
# Utils generales
# =========================
def guardar_debug(path: str, content: str):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)
 
 
def es_pdf_real(path: str) -> bool:
    try:
        with open(path, "rb") as f:
            return f.read(5) == b"%PDF-"
    except Exception:
        return False
 
 
def contar_paginas_pdf(path: str) -> int:
    try:
        with pdfplumber.open(path) as pdf:
            return len(pdf.pages)
    except Exception:
        return 0


def detectar_tipo_factura(pdf_path: str) -> str:
    """'GS' si la factura tiene una sola hoja, 'TH' si tiene dos o mas."""
    return "GS" if contar_paginas_pdf(pdf_path) == 1 else "TH"


def _norm_spaces(s: str) -> str:
    return " ".join((s or "").split())
 
 
def _find_first(pattern: str, text: str, flags=re.I) -> str:
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else ""
 
 
def _norm_money(s: str) -> str:
    if not s:
        return ""
    s = str(s).strip().replace(" ", "")
    s = re.sub(r"[^0-9\.,\-]", "", s)
    return s


def _to_excel_number(value):
    """Convierte strings numericos estilo UY a int/float para exportacion Excel."""
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return value

    raw = str(value).strip()
    if not raw:
        return None

    # No convertir rangos o textos compuestos como "100 - 100".
    if re.search(r"[A-Za-z]", raw) or " - " in raw or "–" in raw:
        return value

    cleaned = _norm_money(raw)
    if not cleaned:
        return value

    # Si tiene coma decimal, quitar separador de miles con punto.
    if "," in cleaned:
        cleaned = cleaned.replace(".", "").replace(",", ".")
    else:
        # Si solo tiene puntos y hay mas de uno, asumir miles.
        if cleaned.count(".") > 1:
            cleaned = cleaned.replace(".", "")

    if not re.fullmatch(r"-?\d+(?:\.\d+)?", cleaned):
        return value

    try:
        number = float(cleaned)
        return int(number) if number.is_integer() else number
    except Exception:
        return value
 
 
def _read_pdf_text(pdf_path: str) -> str:
    parts: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for p in pdf.pages:
            parts.append(p.extract_text() or "")
    return "\n".join(parts)
 
 
# =========================
# Constantes de parsing
# =========================
DATE_PAT = r"(\d{2}[/-]\d{2}[/-]\d{4})"
NUM_UY = r"\d{1,3}(?:\.\d{3})*(?:,\d+)?|\d+(?:,\d+)?"
 
 
# =========================
# Helpers de reconstruccion de texto pegado
# =========================
def _reconstruct_tension(raw: str) -> str:
    raw = re.sub(r"(\d+)(V)", r"\1 \2", raw)
    raw = re.sub(r"(V)(en)", r"\1 \2", raw, flags=re.I)
    raw = re.sub(r"(en)(Baja|Alta|Media)", r"\1 \2", raw, flags=re.I)
    raw = re.sub(r"(Baja|Alta|Media)(tensi)", r"\1 \2", raw, flags=re.I)
    return _norm_spaces(raw)
 
 
def _reconstruct_zona(raw: str) -> str:
    raw = re.sub(r"(ADT)(\d)", r"\1 \2", raw)
    raw = re.sub(r"(\d)([-\u2013])([A-Z])", r"\1 \2 \3", raw)
    raw = re.sub(r"([a-z\xe1\xe9\xed\xf3\xfa\xfc\xf1])([A-Z\xc1\xc9\xcd\xd3\xda\xdc\xd1])", r"\1 \2", raw)
    raw = re.sub(r"(densidad)(alta|baja|media)", r"\1 \2", raw, flags=re.I)
    raw = re.sub(r"(Urbana)(densidad)", r"\1 \2", raw, flags=re.I)
    return _norm_spaces(raw)
 
 
# =========================
# Parsing PDF — cuadro superior
# =========================
def _extract_page2_text(pdf_path: str) -> str:
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[1] if len(pdf.pages) > 1 else pdf.pages[0]
        return page.extract_text() or ""


def _agrupar_words_en_lineas(words: list[dict], y_tolerance: float = 4.0) -> list[str]:
    if not words:
        return []

    lineas: list[list[dict]] = []
    for word in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if not lineas or abs(word["top"] - lineas[-1][0]["top"]) > y_tolerance:
            lineas.append([word])
        else:
            lineas[-1].append(word)

    return [
        _norm_spaces(" ".join(w["text"] for w in sorted(linea, key=lambda w: w["x0"])))
        for linea in lineas
    ]


def _limpiar_linea_encabezado_cliente(linea: str) -> str:
    linea = _norm_spaces(linea)
    linea = re.sub(r"\bHoja\s+\d+\s+de\s+\d+\b", "", linea, flags=re.I)
    linea = re.sub(r"\bDETALLE\s+DE\b", "", linea, flags=re.I)
    linea = re.sub(r"\bFACTURA\b", "", linea, flags=re.I)
    linea = re.sub(r"\b\d{10}\b", "", linea)
    linea = re.sub(r"\b\d{2}/\d{2}/\d{4}\b", "", linea)
    return _norm_spaces(linea)


def extract_cliente_ubicacion(pdf_path: str) -> dict:
    p2 = _extract_page2_text(pdf_path)

    stop_markers = (
        "DETALLE DE FACTURA",
        "MEDIDOR Nro",
        "MEDIDOR NRO",
        "OFICINA COMERCIAL",
        "Acuerdo de Servicio",
        "Tarifa Aplicada",
        "Direccion del",
        "Período de",
        "Periodo de",
    )
    skip_patterns = [
        re.compile(r"^UTE\b", re.I),
        re.compile(r"^ADMINISTRACI", re.I),
        re.compile(r"^PARAGUAY\s+\d+", re.I),
        re.compile(r"^R\.?U\.?T\.?", re.I),
        re.compile(r"^(098|0800|\+?\d{2,})"),
        re.compile(r"^Hoja\s+\d+", re.I),
        re.compile(r"^\d{2}/\d{2}/\d{4}$"),
        re.compile(r"^\d{10}$"),
    ]

    encabezado: list[str] = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[1] if len(pdf.pages) > 1 else pdf.pages[0]
            words = page.extract_words(use_text_flow=False, keep_blank_chars=False)

        # El bloque del cliente aparece en el centro superior. Filtrar por
        # coordenadas evita mezclar logo, telefonos, nro. de factura y fecha.
        x_min = page.width * 0.28
        x_max = page.width * 0.78
        y_max = page.height * 0.24
        words_cliente = [
            w for w in words
            if x_min <= w["x0"] <= x_max and w["top"] <= y_max
        ]

        for linea in _agrupar_words_en_lineas(words_cliente):
            if any(marker.lower() in linea.lower() for marker in stop_markers):
                break
            if any(p.search(linea) for p in skip_patterns):
                continue
            linea = _limpiar_linea_encabezado_cliente(linea)
            if linea:
                encabezado.append(linea)
    except Exception:
        encabezado = []

    if not encabezado:
        lineas = [_norm_spaces(linea) for linea in p2.splitlines()]
        lineas = [linea for linea in lineas if linea]

        for linea in lineas:
            if any(marker.lower() in linea.lower() for marker in stop_markers):
                break
            if any(p.search(linea) for p in skip_patterns):
                continue
            linea = _limpiar_linea_encabezado_cliente(linea)
            if linea:
                encabezado.append(linea)

    encabezado = [
        linea for linea in encabezado
        if not re.fullmatch(r"(DETALLE|DE|FACTURA)(\s+(DETALLE|DE|FACTURA))*", linea, re.I)
    ]

    nombre_cliente = ""
    direccion_cliente = ""
    localidad_cliente = ""
    departamento_cliente = ""

    if encabezado:
        nombre_cliente = encabezado[0]
    if len(encabezado) > 1:
        direccion_cliente = encabezado[1]

    idx_localidad = None
    for i, linea in enumerate(encabezado[2:], start=2):
        if re.search(r"\s-\s*CP\b|\bCP\s*\d", linea, re.I):
            idx_localidad = i
            break

    if idx_localidad is not None:
        localidad_cliente = encabezado[idx_localidad]
        posteriores = encabezado[idx_localidad + 1:]
        if posteriores:
            departamento_cliente = posteriores[-1]
    else:
        if len(encabezado) > 2:
            localidad_cliente = encabezado[2]
        if len(encabezado) > 3:
            departamento_cliente = encabezado[-1]

    if localidad_cliente:
        localidad_cliente = re.split(r"\s-\s*CP\b|\bCP\s*\d", localidad_cliente, maxsplit=1, flags=re.I)[0]
        localidad_cliente = _norm_spaces(localidad_cliente)

    if not departamento_cliente and localidad_cliente.upper() == "MONTEVIDEO":
        departamento_cliente = "MONTEVIDEO"

    return {
        "nombre_cliente": nombre_cliente,
        "direccion_cliente": direccion_cliente,
        "localidad_cliente": localidad_cliente,
        "departamento_cliente": departamento_cliente,
    }


def extract_total_detalle_facturacion(pdf_path: str) -> str:
    p2 = _extract_page2_text(pdf_path)
    all_text = _read_pdf_text(pdf_path)

    fuentes = [p2, all_text]
    patrones = [
        re.compile(
            r"DETALLE\s+DE\s+FACTURACION[\s\S]{0,2500}?\bTOTAL\b\s*(" + NUM_UY + r")",
            re.I,
        ),
        re.compile(
            r"DETALLE\s+DE\s+FACTURACION[\s\S]{0,2500}?\nTOTAL\s+(" + NUM_UY + r")",
            re.I,
        ),
    ]

    for text in fuentes:
        for patron in patrones:
            m = patron.search(text)
            if m:
                return _norm_money(m.group(1))
    return ""


def extract_cuadro_superior(pdf_path: str, archivo: str) -> dict:
    all_text = _read_pdf_text(pdf_path)
    p2 = _extract_page2_text(pdf_path)
    cliente_ubicacion = extract_cliente_ubicacion(pdf_path)
    total_detalle_facturacion = extract_total_detalle_facturacion(pdf_path)
 
    # Numero de cuenta
    nro_cuenta = _find_first(r"\b(\d{10})\b", p2)
    if not nro_cuenta:
        nro_cuenta = _find_first(r"\b(\d{10})\b", all_text)
 
    # Acuerdo de servicio
    acuerdo = _find_first(r"Acuerdo\s+de\s+[Ss]ervicio\s+(\d{7,12})", p2)
    if not acuerdo:
        acuerdo = _find_first(r"Acuerdo\s+de\s+[Ss]ervicio[:\s]+(\d{7,12})", all_text)
 
    # Nro. Medidor
    nro_medidor = _find_first(r"MEDIDOR\s+Nro\.?\s*(\d+)", p2)
    if not nro_medidor:
        nro_medidor = _find_first(r"Nro\.\s*Medidor\s+(\d+)", p2)
    if not nro_medidor:
        nro_medidor = _find_first(r"MEDIDOR\s+Nro\.?\s*(\d+)", all_text)
 
    # Nro. Factura, Fecha Emision, Proximo Vencimiento
    nro_factura = fecha_emision = prox_venc = ""
    ml = re.search(r"\b([A-Z])\s*(\d{6,10})\s+" + DATE_PAT + r"\s+" + DATE_PAT + r"\b", all_text)
    if ml:
        nro_factura = f"{ml.group(1)} {ml.group(2)}"
        fecha_emision = ml.group(3)
        prox_venc = ml.group(4)
    else:
        mnf = re.search(r"\b([A-Z])\s*(\d{6,10})\b", all_text)
        if mnf:
            nro_factura = f"{mnf.group(1)} {mnf.group(2)}"
 
    # Tarifa Aplicada
    tarifa = ""
    m_tar = re.search(
        r"([A-Za-z\xc0-\xff][^\n]{3,})\n\s*Tarifa\s+Aplicada\s*\n([^\n]*)",
        p2, re.I
    )
    if m_tar:
        parte1 = _norm_spaces(m_tar.group(1))
        parte2 = re.split(r"Direcci|Potencia|Consumo", m_tar.group(2))[0]
        tarifa = _norm_spaces((parte1 + " " + parte2).strip(" -"))
    if not tarifa:
        tarifa = _norm_spaces(_find_first(r"Tarifa:\s*(.+?)(?=\n|Potencia|$)", all_text))
 
    # Consumo Activo (kWh)
    consumo_activo = ""
    m_ca = re.search(r"\(kWh\)\s+(\d[\d\.,]*)", p2)
    if m_ca:
        consumo_activo = _norm_money(m_ca.group(1))
    if not consumo_activo:
        m_ca2 = re.search(r"\(kWh\)\s*\n\s*(\d[\d\.,]*)", p2)
        if m_ca2:
            consumo_activo = _norm_money(m_ca2.group(1))
    if not consumo_activo:
        consumo_activo = _norm_money(_find_first(r"Consumo\s+Activo[^\d]*(\d[\d\.,]*)", all_text))
 
    # Consumo Reactivo (kVArh)
    # Estructura: "CR_VAL PV_VAL DD/MM/YYYY...\n(kVArh) Valle..."
    # CR_VAL = primer numero de la linea que precede a "(kVArh)"
    consumo_reactivo = ""
    m_cr = re.search(
        r"(\d[\d\.,]*)\s+\d[\d\.,]*\s+\d{2}/\d{2}/\d{4}[^\n]*\n[^\n]*\(kVArh\)",
        p2
    )
    if m_cr:
        consumo_reactivo = _norm_money(m_cr.group(1))
    if not consumo_reactivo:
        m_cr2 = re.search(r"(\d[\d\.,]*)\s*\n\s*\(kVArh\)", p2)
        if m_cr2:
            consumo_reactivo = _norm_money(m_cr2.group(1))
    if not consumo_reactivo:
        m_cr3 = re.search(r"\(kVArh\)\s+(\d[\d\.,]*)", p2)
        if m_cr3:
            consumo_reactivo = _norm_money(m_cr3.group(1))
    if consumo_reactivo == "":
        consumo_reactivo = "0"
 
    # Potencia Contratada Punta-Llano (kW)
    PATRON_PAR = r"\d+(?:[,\.]\d+)?\s*[-\u2013]\s*\d+(?:[,\.]\d+)?"
    potencia_punta_llano = ""
    m_ppl = re.search(r"Punta\s*[-\u2013]\s*Llano\s*\(kW\)\s*\n\s*(" + PATRON_PAR + r")", p2, re.I)
    if m_ppl:
        potencia_punta_llano = _norm_spaces(m_ppl.group(1))
    if not potencia_punta_llano:
        m_ppl2 = re.search(r"Punta\s*[-\u2013]\s*Llano\s*\(kW\)\s+(" + PATRON_PAR + r")", p2, re.I)
        if m_ppl2:
            potencia_punta_llano = _norm_spaces(m_ppl2.group(1))
    if not potencia_punta_llano:
        m_ppl3 = re.search(r"(" + PATRON_PAR + r")[^\n]*\n[^\n]*Punta\s*[-\u2013]\s*Llano", p2, re.I)
        if m_ppl3:
            potencia_punta_llano = _norm_spaces(m_ppl3.group(1))
    if not potencia_punta_llano:
        m_ppl4 = re.search(r"Punta\s*[-\u2013]\s*Llano\s*\(kW\)[:\s]*(" + PATRON_PAR + r")", all_text, re.I)
        if m_ppl4:
            potencia_punta_llano = _norm_spaces(m_ppl4.group(1))
 
    # Potencia Contratada Valle (kW)
    potencia_valle = ""
    m_pv = re.search(r"Potencia\s+Contratada\s*\n?\s*Valle\s*\(kW\)\s+(\d[\d\.,]*)", p2, re.I)
    if m_pv:
        potencia_valle = _norm_money(m_pv.group(1))
    if not potencia_valle:
        m_pv2 = re.search(r"Valle\s*\(kW\)\s+(\d[\d\.,]*)", p2, re.I)
        if m_pv2:
            potencia_valle = _norm_money(m_pv2.group(1))
    if not potencia_valle:
        m_pv3 = re.search(
            r"\d[\d\.,]*\s+(\d[\d\.,]*)\s+\d{2}/\d{2}/\d{4}[a\s]+\d{2}/\d{2}/\d{4}"
            r"[^\n]*\n[^\n]*\(kVArh\)[^\n]*Valle",
            p2, re.I
        )
        if m_pv3:
            potencia_valle = _norm_money(m_pv3.group(1))
    if potencia_valle == "":
        potencia_valle = "0"
 
    # Tension
    tension = ""
    m_ten = re.search(r"Tensi[o\xf3]n\s+([\d]+V[^\n]+?)(?=\s*(?:Tipo\s+de\s+Zona|ADT\d|Rural|Fases|\n))", p2, re.I)
    if m_ten:
        tension = _reconstruct_tension(m_ten.group(1))
    if not tension:
        m_ten2 = re.search(r"Tensi[o\xf3]n\s+(.+?)(?=\n)", p2, re.I)
        if m_ten2:
            tension = _reconstruct_tension(m_ten2.group(1))
    if not tension:
        m_ten3 = re.search(r"Tensi[o\xf3]n\s+([\d]+\s*V\b[^\n]+)", p2, re.I)
        if m_ten3:
            tension = _reconstruct_tension(m_ten3.group(1))
 
    # Fases
    fases = _norm_spaces(_find_first(r"Fases?\s+([\w]+)", p2))
    if not fases:
        fases = _norm_spaces(_find_first(r"Fases?\s*[:\-]?\s*([^\n]+)", all_text))
 
    # Direccion del Servicio
    direccion = ""
    m_dir_p1 = re.search(
        r"Acuerdo\s+de\s+[Ss]ervicio:\s*\d+[^\n]*\n([\s\S]+?)(?=\nTarifa:)",
        all_text, re.I
    )
    if m_dir_p1:
        bloque = m_dir_p1.group(1)
        lineas = [_norm_spaces(l) for l in bloque.splitlines() if _norm_spaces(l)]
        dir_parts = []
        for l in lineas:
            tiene_numero = bool(re.search(r'\d', l))
            es_ciudad = bool(re.match(r'^[A-Z\xc1\xc9\xcd\xd3\xda\xd1\s]+$', l) and 3 < len(l) < 30)
            if tiene_numero and not re.match(r'^(R\.U\.T\.|---+|\.\.\.)', l, re.I):
                if not re.search(r'(Total|Importe|Concepto|\$|CAE|Rango|Fecha\s+de\s+vto)', l, re.I):
                    dir_parts.append(l)
            elif es_ciudad:
                dir_parts.append(l)
            if len(dir_parts) == 3:
                break
        if dir_parts:
            direccion = ", ".join(dir_parts)
    if not direccion:
        m_dir2 = re.search(r"Direcci[o\xf3]n\s+del[^\n]*?(CL[^\s][^\n]+)", p2, re.I)
        if m_dir2:
            direccion = _norm_spaces(m_dir2.group(1))
    if not direccion:
        m_dir3 = re.search(
            r"Direcci[o\xf3]n\s+del\s*\n?\s*[Ss]ervicio\s*\n\s*([^\n]+)\n\s*([^\n]+)",
            p2, re.I
        )
        if m_dir3:
            partes = [_norm_spaces(g) for g in [m_dir3.group(1), m_dir3.group(2)]
                      if _norm_spaces(g) and len(_norm_spaces(g)) > 4]
            if partes:
                direccion = ", ".join(partes)
 
    # Periodo de Consumo
    periodo = ""
    m_per = re.search(
        r"Per[i\xed]odo\s+de\s*\n?\s*[Cc]onsumo\s+"
        r"(\d{2}[/\-]\d{2}[/\-]\d{4})\s*a\s*(\d{2}[/\-]\d{2}[/\-]\d{4})",
        p2, re.I
    )
    if m_per:
        periodo = f"{m_per.group(1)} a {m_per.group(2)}"
    if not periodo:
        m_per2 = re.search(r"(\d{2}/\d{2}/\d{4})a(\d{2}/\d{2}/\d{4})", p2)
        if m_per2:
            periodo = f"{m_per2.group(1)} a {m_per2.group(2)}"
    if not periodo:
        m_per3 = re.search(DATE_PAT + r"\s*a\s*" + DATE_PAT, all_text)
        if m_per3:
            periodo = f"{m_per3.group(1)} a {m_per3.group(2)}"
 
    # Tipo de Zona Electrica
    zona = ""
    m_zona = re.search(r"Tipo\s+de\s*\n?\s*Zona\s*\n?\s*El[e\xe9]ctrica\s*\n?\s*([^\n]+)", p2, re.I)
    if m_zona:
        zona = _reconstruct_zona(m_zona.group(1))
    if not zona:
        m_zona2 = re.search(r"(ADT\s*\d+[-\u2013][A-Za-z][^\n]+?)(?=\n|\Z)", p2)
        if m_zona2:
            zona = _reconstruct_zona(m_zona2.group(1))
    if not zona:
        zona = _norm_spaces(_find_first(
            r"(?:Tipo\s+de\s+Zona|Zona\s+El[e\xe9]ctrica)\s*[:\-]?\s*([^\n]+)", all_text
        ))
 
    return {
        "archivo":                            archivo,
        "nro_cuenta":                         nro_cuenta,
        "nro_factura":                        nro_factura,
        "fecha_emision":                      fecha_emision,
        "prox_vencimiento":                   prox_venc,
        "acuerdo_servicio":                   acuerdo,
        "nro_medidor":                        nro_medidor,
        "tarifa_aplicada":                    tarifa,
        "potencia_contratada_punta_llano_kw": potencia_punta_llano,
        "potencia_contratada_valle_kw":       potencia_valle,
        "consumo_activo_kwh":                 consumo_activo,
        "consumo_reactivo_kvarh":             consumo_reactivo,
        "tension":                            tension,
        "fases":                              fases,
        "direccion_servicio":                 direccion,
        "periodo_consumo":                    periodo,
        "zona_electrica":                     zona,
        "total_detalle_facturacion":          total_detalle_facturacion,
        "nombre_cliente":                     cliente_ubicacion["nombre_cliente"],
        "direccion_cliente":                  direccion or cliente_ubicacion["direccion_cliente"],
        "localidad_cliente":                  cliente_ubicacion["localidad_cliente"],
        "departamento_cliente":               cliente_ubicacion["departamento_cliente"],
    }
 
 
# =========================
# Parsing PDF — lecturas pivoteadas
# =========================
def extract_lecturas_pivotadas(pdf_path: str) -> dict:
    all_text = _read_pdf_text(pdf_path)
    with pdfplumber.open(pdf_path) as pdf:
        page_text = pdf.pages[1].extract_text() if len(pdf.pages) > 1 else (pdf.pages[0].extract_text() or "")
 
    TIPO_PAT = (
        r"Potencia(?:\s+(?:Punta|Valle|Llano))?"
        r"|Energ[i\xed]a\s+(?:Punta|Valle|Llano|Reactiva(?:\s+Q4)?|sal\.?\s*(?:Punta|Valle|Llano))"
        r"|Energ[i\xed]a\s+[Ss]al\.\s*(?:Punta|Valle|Llano)"
        r"|Punta(?:\s+NO?\s+[Hh][a\xe1]bil(?:es)?)?"
        r"|Fuera\s+de\s+Punta"
        r"|Valle|Llano"
        r"|Reactiva(?:\s+Q4)?"
    )
    row_pat = re.compile(
        rf"(?P<tipo>{TIPO_PAT})\s+"
        rf"(?P<ant>{NUM_UY})\s+"
        rf"(?P<act>{NUM_UY})\s+"
        rf"(?P<factor>\d+)\s+"
        rf"(?P<total>{NUM_UY})\s+"
        rf"(?P<tlec>[A-Za-z\u00C0-\u00FF]+)\b",
        re.I | re.M
    )
 
    rows = []
    seen = set()
    for text_src in [page_text, all_text]:
        if rows:
            break
        for m in row_pat.finditer(text_src):
            tipo = _norm_spaces(m.group("tipo"))
            key = (tipo.lower(), m.group("ant"), m.group("act"))
            if key in seen:
                continue
            seen.add(key)
            rows.append({
                "tipo":     tipo,
                "lect_ant": _norm_money(m.group("ant")),
                "lect_act": _norm_money(m.group("act")),
                "factor":   m.group("factor"),
                "total":    _norm_money(m.group("total")),
                "tipo_lec": _norm_spaces(m.group("tlec")),
            })
 
    result = {}
    for r in rows:
        col_base = re.sub(r"\s+", "_", r["tipo"].lower())
        col_base = re.sub(r"[^a-z0-9_]", "", col_base)
        result[f"{col_base}_lect_ant"] = r["lect_ant"]
        result[f"{col_base}_lect_act"] = r["lect_act"]
        result[f"{col_base}_factor"]   = r["factor"]
        result[f"{col_base}_total"]    = r["total"]
        result[f"{col_base}_tipo_lec"] = r["tipo_lec"]
 
    return result


# =========================
# Parsing PDF — factura GS (una hoja)
# =========================
# Layout distinto al de TH: una sola pagina en vez de dos, tarifa por escalones
# en vez de Punta/Valle/Llano, y un solo valor de Potencia Contratada (no
# punta-llano/valle). Las etiquetas "Tarifa Aplicada" / "Direccion del Servicio"
# / "Tipo de Zona Electrica" SI existen, solo que comprimidas de a dos boxes por
# fila (ej. "Direccion del Servicio" y "Acuerdo de Servicio" en la misma fila) —
# eso hace que pdfplumber.extract_text() a veces mezcle en una misma linea el
# texto de la columna izquierda y de la derecha cuando comparten altura, por lo
# que la extraccion de estos campos no puede asumir adyacencia estricta entre
# etiqueta y valor (ver acuerdo_servicio/periodo_consumo/tarifa_aplicada abajo).
# Se mantiene separado de extract_cuadro_superior/extract_lecturas_pivotadas (TH)
# para no arriesgar regresiones en el flujo ya probado en produccion.
#
# IMPORTANTE (confirmado con texto crudo real de facturas GS de produccion):
# en varias facturas GS, las etiquetas "Tarifa Aplicada" / "Acuerdo de
# Servicio" / "Direccion del Servicio" / "Vencimiento" NO existen como texto
# extraible (se renderizan como graficos/pills sin texto embebido) — solo el
# valor queda en el texto, sin ninguna etiqueta cerca. Por eso varias
# extracciones de esta funcion y de extract_cuadro_superior_gs usan la
# POSICION relativa de los valores (linea siguiente a tal otro dato) en vez
# de anclarse a una etiqueta, con fallback a un anclaje por etiqueta cuando
# si esta presente en otras variantes de plantilla.
DEPARTAMENTOS_UY = {
    "MONTEVIDEO", "CANELONES", "MALDONADO", "COLONIA", "SAN JOSE", "SORIANO",
    "RIO NEGRO", "PAYSANDU", "SALTO", "ARTIGAS", "RIVERA", "TACUAREMBO",
    "DURAZNO", "FLORIDA", "FLORES", "LAVALLEJA", "TREINTA Y TRES", "ROCHA",
    "CERRO LARGO",
}

# Fallback para cuando la factura no imprime el departamento como linea
# propia (o esa linea no matchea DEPARTAMENTOS_UY): mapeo estatico de
# localidad -> departamento. Geografia de Uruguay, no cambia con el tiempo.
# Alcance inicial: capitales departamentales + localidades comunes; si
# aparece una localidad no mapeada, el campo queda vacio como antes (sin
# regresion), y se puede agregar la entrada cuando se detecte.
DEPARTAMENTOS_POR_LOCALIDAD = {
    "MONTEVIDEO": "MONTEVIDEO",
    "CANELONES": "CANELONES", "PANDO": "CANELONES", "LAS PIEDRAS": "CANELONES",
    "SAN RAMON": "CANELONES", "SANTA LUCIA": "CANELONES",
    "MALDONADO": "MALDONADO", "PUNTA DEL ESTE": "MALDONADO", "SAN CARLOS": "MALDONADO",
    "COLONIA": "COLONIA", "CARMELO": "COLONIA", "NUEVA HELVECIA": "COLONIA", "JUAN LACAZE": "COLONIA",
    "SAN JOSE": "SAN JOSE", "LIBERTAD": "SAN JOSE",
    "MERCEDES": "SORIANO", "DOLORES": "SORIANO",
    "FRAY BENTOS": "RIO NEGRO", "YOUNG": "RIO NEGRO",
    "PAYSANDU": "PAYSANDU",
    "SALTO": "SALTO",
    "ARTIGAS": "ARTIGAS", "BELLA UNION": "ARTIGAS",
    "RIVERA": "RIVERA", "TRANQUERAS": "RIVERA",
    "TACUAREMBO": "TACUAREMBO",
    "DURAZNO": "DURAZNO",
    "FLORIDA": "FLORIDA", "SARANDI GRANDE": "FLORIDA",
    "TRINIDAD": "FLORES",
    "MINAS": "LAVALLEJA",
    "TREINTA Y TRES": "TREINTA Y TRES",
    "ROCHA": "ROCHA", "CHUY": "ROCHA", "LA PALOMA": "ROCHA",
    "MELO": "CERRO LARGO",
}


def extract_cliente_ubicacion_gs(pdf_path: str, all_text: str) -> dict:
    lineas = [_norm_spaces(l) for l in all_text.splitlines() if _norm_spaces(l)]
    nombre_cliente = lineas[0] if lineas else ""

    localidad_cliente = ""
    departamento_cliente = ""

    # El departamento (cuando la factura lo imprime) aparece como linea
    # propia inmediatamente despues de "LOCALIDAD - CP nnnnn-nnnnn"; se
    # valida contra DEPARTAMENTOS_UY para no capturar por error otra linea
    # (ej. en facturas donde esa linea sea una fecha u otro dato).
    m_loc_depto = re.search(
        r"\n([A-Z\xc1\xc9\xcd\xd3\xda\xd1 ]{3,})\s*-\s*CP\s*[\d-]+[^\n]*\n([A-Z\xc1\xc9\xcd\xd3\xda\xd1 ]{3,})\n",
        all_text
    )
    if m_loc_depto:
        localidad_cliente = _norm_spaces(m_loc_depto.group(1))
        posible_depto = _norm_spaces(m_loc_depto.group(2)).upper()
        if posible_depto in DEPARTAMENTOS_UY:
            departamento_cliente = posible_depto

    m_loc = re.search(r"\n([A-Z\xc1\xc9\xcd\xd3\xda\xd1 ]{3,})\s*-\s*CP\s*[\d-]+", all_text)
    if not localidad_cliente and m_loc:
        localidad_cliente = _norm_spaces(m_loc.group(1))

    # Direccion del suministro: segundo bloque nombre+direccion, ubicado despues
    # del R.U.T. (formato: <direccion>\n<10 digitos>\n<localidad>). Se prioriza
    # sobre la direccion de facturacion porque el caso de uso central de la app
    # es un cliente con N subcuentas en direcciones fisicas distintas.
    #
    # Se extrae por palabras/coordenadas (columna izquierda, x0 < 260) en vez
    # de texto plano porque pdfplumber.extract_text() fusiona en una misma
    # linea palabras de la columna izquierda (bloque cliente) y de la derecha
    # (detalle de facturacion) cuando comparten altura — esto ocurre en
    # facturas con potencia contratada prorrateada (cambio de tarifa a mitad
    # de periodo), donde aparece una fila extra de detalle a la misma altura
    # que la direccion/localidad.
    direccion_cliente = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
        col_izq = [w for w in words if w["x0"] < 260]
        lineas_col = _agrupar_words_en_lineas(col_izq)
        idx_rut = next(
            (i for i, l in enumerate(lineas_col) if re.match(r"^R\.U\.T\.", l, re.I)),
            None
        )
        if idx_rut is not None:
            posteriores = lineas_col[idx_rut + 1:]
            # Caso mas comun: la direccion y el nro. de cuenta quedan en la
            # misma linea agrupada porque estan a alturas casi identicas
            # (ej. "CL ENTRE RIOS 1098 0591771164").
            m_embed = re.search(r"^(.*?)\s*\b(\d{10})\b\s*$", posteriores[0]) if posteriores else None
            if m_embed:
                direccion_cliente = _norm_spaces(m_embed.group(1))
                if len(posteriores) > 1:
                    localidad_cliente = _norm_spaces(posteriores[1])
            else:
                # Fallback: direccion, nro. de cuenta y localidad en lineas separadas.
                idx_cuenta = next(
                    (i for i, l in enumerate(posteriores) if re.fullmatch(r"\d{10}", l)),
                    None
                )
                if idx_cuenta:
                    direccion_cliente = _norm_spaces(posteriores[idx_cuenta - 1])
                    if idx_cuenta + 1 < len(posteriores):
                        localidad_cliente = _norm_spaces(posteriores[idx_cuenta + 1])
    except Exception:
        pass

    if not departamento_cliente and localidad_cliente.upper() == "MONTEVIDEO":
        departamento_cliente = "MONTEVIDEO"
    if not departamento_cliente:
        departamento_cliente = DEPARTAMENTOS_POR_LOCALIDAD.get(localidad_cliente.upper(), "")

    return {
        "nombre_cliente": nombre_cliente,
        "direccion_cliente": direccion_cliente,
        "localidad_cliente": localidad_cliente,
        "departamento_cliente": departamento_cliente,
    }


def extract_cuadro_superior_gs(pdf_path: str, archivo: str) -> dict:
    all_text = _read_pdf_text(pdf_path)
    cliente_ubicacion = extract_cliente_ubicacion_gs(pdf_path, all_text)

    nro_cuenta = _find_first(r"\b(\d{10})\b", all_text)

    # Nro. Medidor: tolerar que no haya un espacio real entre "MEDIDOR" y
    # "Nro." en el texto extraido, y si el patron principal no matchea, caer
    # a la tabla "Informacion de Medidores" (mismo numero sin ceros a la
    # izquierda, ej. "Med 812505240" para "MEDIDOR Nro.00812505240").
    nro_medidor = _find_first(r"MEDIDOR[\s\S]{0,20}?(\d{6,})", all_text)
    if not nro_medidor:
        nro_medidor = _find_first(r"\bMed\.?\s+(\d{5,})\b", all_text)

    # Acuerdo de Servicio: en algunas plantillas la etiqueta "Acuerdo de
    # Servicio" existe como texto (misma fila que "Direccion del Servicio",
    # dos boxes por fila, por lo que el valor puede no quedar pegado a la
    # etiqueta). En otras plantillas la etiqueta NO existe como texto en
    # absoluto (confirmado con texto crudo real) y el valor aparece como
    # linea propia, sin ninguna etiqueta, inmediatamente despues de la
    # direccion del servicio (2do bloque direccion, ver
    # extract_cliente_ubicacion_gs) y antes de la localidad del servicio.
    acuerdo = _find_first(r"Acuerdo\s+de\s+[Ss]ervicio[\s\S]{0,150}?\b(\d{7,12})\b", all_text)
    if not acuerdo and cliente_ubicacion["direccion_cliente"]:
        idx_dir = all_text.rfind(cliente_ubicacion["direccion_cliente"])
        if idx_dir != -1:
            ventana = all_text[idx_dir: idx_dir + 400]
            m_acuerdo = re.search(r"^\s*(\d{7,12})\s*$", ventana, re.M)
            if m_acuerdo:
                acuerdo = m_acuerdo.group(1)

    # Nro. Factura / Fecha de Emision: el valor de "Proxima Vencimiento" de
    # esta misma fila puede venir vacio (ver prox_venc mas abajo), asi que no
    # se puede exigir una segunda fecha para no perder los tres campos juntos.
    nro_factura = fecha_emision = prox_venc = ""
    ml = re.search(r"\b([A-Z])\s*(\d{6,10})\s+" + DATE_PAT + r"\b", all_text)
    if ml:
        nro_factura = f"{ml.group(1)} {ml.group(2)}"
        fecha_emision = ml.group(3)

    # Proximo Vencimiento: se muestra en una placa separada (sin etiqueta de
    # texto "Vencimiento" extraible), como linea propia inmediatamente
    # despues de "LOCALIDAD - CP nnnnn-nnnnn" y, si existe, del departamento.
    m_venc = re.search(
        r"-\s*CP\s*[\d-]+[^\n]*\n(?:[A-Z\xc1\xc9\xcd\xd3\xda\xd1 ]+\n)?" + DATE_PAT,
        all_text
    )
    if m_venc:
        prox_venc = m_venc.group(1)

    # Periodo de Consumo: anclado a su propia etiqueta ("Periodo de Consumo
    # Electrico"), independiente de la Tarifa Aplicada — antes ambos se
    # extraian con un solo regex que asumia que quedaban pegados en la misma
    # linea (misma fila de dos boxes que Direccion/Acuerdo de Servicio), y si
    # no matcheaba se perdian los dos valores a la vez.
    periodo = ""
    m_per = re.search(
        r"Per[i\xed]odo\s+de\s+[Cc]onsumo(?:\s+[Ee]l[e\xe9]ctrico)?[\s\S]{0,60}?"
        + DATE_PAT + r"\s*a\s*" + DATE_PAT,
        all_text, re.I
    )
    if m_per:
        periodo = f"{m_per.group(1)} a {m_per.group(2)}"
    if not periodo:
        m_per2 = re.search(DATE_PAT + r"\s*a\s*" + DATE_PAT, all_text)
        if m_per2:
            periodo = f"{m_per2.group(1)} a {m_per2.group(2)}"

    # Tarifa Aplicada: primero se intenta anclada a su etiqueta (tolerando que
    # el valor venga pegado al rango de fechas del periodo, ej. "Tarifa
    # Aplicada\n18/06/2026 a 18/07/2026 General Simple."); si no matchea, se
    # cae al regex combinado original.
    tarifa = ""
    m_tar = re.search(
        r"Tarifa\s+Aplicada[\s\S]{0,60}?\n\s*(?:" + DATE_PAT + r"\s*a\s*" + DATE_PAT + r"\s+)?(?P<tarifa>[^\n]+?)\.?\s*\n",
        all_text, re.I
    )
    if m_tar:
        tarifa = _norm_spaces(m_tar.group("tarifa"))
    if not tarifa:
        # Punto final opcional: en plantillas sin etiqueta "Tarifa Aplicada"
        # el valor viene sin punto de cierre (ej. "... Residencial Simple\n").
        m_per_tar = re.search(
            DATE_PAT + r"\s*a\s*" + DATE_PAT + r"\s+([^\n]+?)\.?\s*\n",
            all_text
        )
        if m_per_tar:
            tarifa = _norm_spaces(m_per_tar.group(3))

    # Potencia Contratada: un solo valor (GS no distingue punta-llano/valle).
    # Cuando la potencia se factura prorrateada (cambio de tarifa a mitad de
    # periodo), la linea siguiente al encabezado arranca con "R.U.T.: <digitos>"
    # antes del valor en kW, asi que no se puede exigir ausencia de digitos.
    potencia_contratada_kw = ""
    m_pot = re.search(r"CARGO\s+POTENCIA\s+CONTRATADA[\s\S]{0,120}?([\d\.,]+)\s*kW", all_text, re.I)
    if m_pot:
        potencia_contratada_kw = _norm_money(m_pot.group(1))

    # Tipo de Zona Electrica: dos fragmentos no adyacentes ("ADT n - Tipo" y,
    # unas lineas mas abajo, "densidad NIVEL").
    zona = ""
    m_zona = re.search(r"(ADT\s*\d+\s*-\s*\w+)[\s\S]{0,60}?\n\s*(densidad\s+\w+)", all_text, re.I)
    if m_zona:
        zona = _norm_spaces(f"{m_zona.group(1)} {m_zona.group(2)}")

    # Total antes de impuestos (equivalente al TOTAL de "DETALLE DE FACTURACION" en TH).
    total_detalle_facturacion = ""
    m_tot = re.search(r"Importe\s+Gravado\s+\d+%\s*([\d\.,]+)", all_text, re.I)
    if m_tot:
        total_detalle_facturacion = _norm_money(m_tot.group(1))

    # Si algun campo "riesgoso" (mezclado de columnas, o extraido por
    # posicion en vez de etiqueta) sigue vacio, volcar el texto crudo para
    # poder ajustar el regex con datos reales sin depender de otra captura
    # de pantalla ni de otra corrida.
    if not (periodo and nro_medidor and acuerdo and nro_factura and tarifa and prox_venc):
        try:
            guardar_debug(
                os.path.join(os.path.dirname(pdf_path), "_debug_gs", f"{archivo}.txt"),
                all_text,
            )
        except Exception:
            pass

    return {
        "archivo":                            archivo,
        "nro_cuenta":                         nro_cuenta,
        "nro_factura":                        nro_factura,
        "fecha_emision":                      fecha_emision,
        "prox_vencimiento":                   prox_venc,
        "acuerdo_servicio":                   acuerdo,
        "nro_medidor":                        nro_medidor,
        "tarifa_aplicada":                    tarifa,
        "potencia_contratada_kw":             potencia_contratada_kw,
        "potencia_contratada_punta_llano_kw": "",
        "potencia_contratada_valle_kw":       "",
        "consumo_activo_kwh":                 "",
        "consumo_reactivo_kvarh":             "",
        "tension":                            "",
        "fases":                              "",
        "direccion_servicio":                 cliente_ubicacion["direccion_cliente"],
        "periodo_consumo":                    periodo,
        "zona_electrica":                     zona,
        "total_detalle_facturacion":          total_detalle_facturacion,
        "nombre_cliente":                     cliente_ubicacion["nombre_cliente"],
        "direccion_cliente":                  cliente_ubicacion["direccion_cliente"],
        "localidad_cliente":                  cliente_ubicacion["localidad_cliente"],
        "departamento_cliente":               cliente_ubicacion["departamento_cliente"],
    }


def extract_lecturas_pivotadas_gs(pdf_path: str) -> dict:
    all_text = _read_pdf_text(pdf_path)

    TIPO_PAT_GS = r"Activa\s+Saliente|Activa|Reactiva"
    row_pat = re.compile(
        rf"(?P<tipo>{TIPO_PAT_GS})\s+"
        rf"(?P<ant>{NUM_UY})\s+"
        rf"(?P<act>{NUM_UY})\s+"
        rf"(?P<total>{NUM_UY})\s+"
        rf"(?P<tlec>[A-Za-zÀ-ÿ]+)\b",
        re.I | re.M
    )

    # Se normaliza el tipo detectado a una clave base fija ("activa" /
    # "reactiva") independientemente de la variante de texto que matcheo
    # (ej. "Activa Saliente" en medidores con generacion/exportacion), para
    # que activa_total/reactiva_total (y sus columnas hermanas) sean siempre
    # las mismas claves en el Excel consolidado en vez de variar segun la
    # etiqueta literal de cada factura puntual.
    rows = []
    seen = set()
    for m in row_pat.finditer(all_text):
        tipo_raw = _norm_spaces(m.group("tipo"))
        col_base = "activa" if tipo_raw.lower().startswith("activa") else "reactiva"
        key = (col_base, m.group("ant"), m.group("act"))
        if key in seen:
            continue
        seen.add(key)
        rows.append({
            "col_base": col_base,
            "lect_ant": _norm_money(m.group("ant")),
            "lect_act": _norm_money(m.group("act")),
            "factor":   "1",
            "total":    _norm_money(m.group("total")),
            "tipo_lec": _norm_spaces(m.group("tlec")),
        })

    # Siempre presentes, aunque esta factura puntual no reporte ese tipo de
    # lectura (evita columnas activa_total/reactiva_total inconsistentes
    # entre facturas del mismo consolidado).
    result = {}
    for base in ("activa", "reactiva"):
        result[f"{base}_lect_ant"] = ""
        result[f"{base}_lect_act"] = ""
        result[f"{base}_factor"]   = ""
        result[f"{base}_total"]    = ""
        result[f"{base}_tipo_lec"] = ""

    for r in rows:
        col_base = r["col_base"]
        result[f"{col_base}_lect_ant"] = r["lect_ant"]
        result[f"{col_base}_lect_act"] = r["lect_act"]
        result[f"{col_base}_factor"]   = r["factor"]
        result[f"{col_base}_total"]    = r["total"]
        result[f"{col_base}_tipo_lec"] = r["tipo_lec"]

    return result


# =========================
# Parse completo
# =========================
def parse_factura_historial(pdf_path: str) -> dict:
    archivo = os.path.basename(pdf_path)
    tipo_factura = detectar_tipo_factura(pdf_path)

    if tipo_factura == "GS":
        cuadro = extract_cuadro_superior_gs(pdf_path, archivo)
        lecturas = extract_lecturas_pivotadas_gs(pdf_path)
        cuadro["consumo_activo_kwh"] = lecturas.get("activa_total") or cuadro.get("consumo_activo_kwh", "")
        cuadro["consumo_reactivo_kvarh"] = lecturas.get("reactiva_total") or "0"
    else:
        cuadro = extract_cuadro_superior(pdf_path, archivo)
        lecturas = extract_lecturas_pivotadas(pdf_path)

    return {**cuadro, **lecturas, "tipo_factura": tipo_factura}
 
 
