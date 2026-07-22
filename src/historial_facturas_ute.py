from __future__ import annotations
 
import os
import re
import time
from datetime import datetime
from urllib.parse import urljoin
 
import pandas as pd
import pdfplumber
import requests
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
 
# =========================
# CONFIG
# =========================
UTE_LOGIN = "https://identityserver.ute.com.uy/Account/Login"
SERVICIOS_EN_LINEA = "https://www.ute.com.uy/clientes/tramites-y-servicios/servicios-en-linea"
 
MAX_FACTURAS_VALIDAS = 24
 
 
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
# Descarga PDF con cookies
# =========================
def cookies_to_requests(context) -> requests.Session:
    s = requests.Session()
    for c in context.cookies():
        s.cookies.set(c["name"], c["value"], domain=c.get("domain"), path=c.get("path"))
    s.headers.update({
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/pdf,application/octet-stream;q=0.9,*/*;q=0.8",
        "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
        "Connection": "keep-alive",
    })
    return s
 
 
def descargar_pdf_con_cookies(context, url: str, out_pdf: str) -> str | None:
    s = cookies_to_requests(context)
    try:
        r = s.get(url, allow_redirects=True, timeout=60)
        if r.status_code != 200:
            return None
        data = r.content
        if data[:5] != b"%PDF-":
            return None
        os.makedirs(os.path.dirname(out_pdf), exist_ok=True)
        with open(out_pdf, "wb") as f:
            f.write(data)
        return out_pdf
    except Exception:
        return None
 
 
def extraer_posibles_urls_pdf(html: str, base_url: str) -> list[str]:
    candidates = set()
    for m in re.findall(r"""href=["']([^"']+\.pdf[^"']*)["']""", html, flags=re.I):
        candidates.add(urljoin(base_url, m))
    for m in re.findall(r"""src=["']([^"']+\.pdf[^"']*)["']""", html, flags=re.I):
        candidates.add(urljoin(base_url, m))
    for m in re.findall(r"""(https?://[^\s"']+\.pdf[^\s"']*)""", html, flags=re.I):
        candidates.add(m)
    return list(candidates)
 
 
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
# Layout distinto al de TH: una sola pagina, sin etiquetas "Tarifa Aplicada" /
# "Direccion del Servicio" / "Tipo de Zona Electrica", tarifa por escalones en vez
# de Punta/Valle/Llano, y un solo valor de Potencia Contratada (no punta-llano/valle).
# Se mantiene separado de extract_cuadro_superior/extract_lecturas_pivotadas (TH)
# para no arriesgar regresiones en el flujo ya probado en produccion.
def extract_cliente_ubicacion_gs(pdf_path: str, all_text: str) -> dict:
    lineas = [_norm_spaces(l) for l in all_text.splitlines() if _norm_spaces(l)]
    nombre_cliente = lineas[0] if lineas else ""

    localidad_cliente = ""
    departamento_cliente = ""
    m_loc = re.search(r"\n([A-Z\xc1\xc9\xcd\xd3\xda\xd1 ]{3,})\s*-\s*CP\s*[\d-]+", all_text)
    if m_loc:
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
    nro_medidor = _find_first(r"MEDIDOR\s+Nro\.?\s*(\d+)", all_text)

    nro_factura = fecha_emision = prox_venc = ""
    ml = re.search(r"\b([A-Z])\s*(\d{6,10})\s+" + DATE_PAT + r"\s+" + DATE_PAT + r"\b", all_text)
    if ml:
        nro_factura = f"{ml.group(1)} {ml.group(2)}"
        fecha_emision = ml.group(3)
        prox_venc = ml.group(4)

    # Periodo de Consumo y Tarifa aparecen pegados en la misma linea, sin
    # etiqueta "Tarifa Aplicada" (ej: "22/05/2026 a 22/06/2026 General Simple.").
    periodo = ""
    tarifa = ""
    m_per_tar = re.search(
        DATE_PAT + r"\s*a\s*" + DATE_PAT + r"\s+([^\n]+?)\.\s*\n",
        all_text
    )
    if m_per_tar:
        periodo = f"{m_per_tar.group(1)} a {m_per_tar.group(2)}"
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

    return {
        "archivo":                            archivo,
        "nro_cuenta":                         nro_cuenta,
        "nro_factura":                        nro_factura,
        "fecha_emision":                      fecha_emision,
        "prox_vencimiento":                   prox_venc,
        "acuerdo_servicio":                   "",
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

    rows = []
    seen = set()
    for m in row_pat.finditer(all_text):
        tipo = _norm_spaces(m.group("tipo"))
        key = (tipo.lower(), m.group("ant"), m.group("act"))
        if key in seen:
            continue
        seen.add(key)
        rows.append({
            "tipo":     tipo,
            "lect_ant": _norm_money(m.group("ant")),
            "lect_act": _norm_money(m.group("act")),
            "factor":   "1",
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
 
 
# =========================
# Navegacion UTE — helpers
# =========================
def esperar_frame_autoservicio(page):
    patrones_url = [
        re.compile(r"autoservicio\.ute\.com\.uy", re.I),
        re.compile(r"selfservice", re.I),
    ]
    for _ in range(160):
        for fr in page.frames:
            fr_url = (fr.url or "").strip()
            if fr_url and any(p.search(fr_url) for p in patrones_url):
                return fr
        for sel in [
            "iframe[src*='autoservicio']",
            "iframe[src*='SelfService']",
            "iframe[name*='autoservicio' i]",
            "iframe[id*='autoservicio' i]",
            "iframe[src*='ute.com.uy']",
        ]:
            loc = page.locator(sel)
            if loc.count() > 0:
                try:
                    fr = loc.first.content_frame()
                    if fr:
                        return fr
                except Exception:
                    pass
        page.wait_for_timeout(250)
    return None
 
 
def _resumen_frames_debug(page, max_items: int = 8) -> str:
    frames = []
    for fr in page.frames:
        try:
            u = (fr.url or "").strip()
        except Exception:
            u = ""
        if u:
            frames.append(u)
    vistos = list(dict.fromkeys(frames))
    if len(vistos) > max_items:
        return " | ".join(vistos[:max_items]) + f" | ... (+{len(vistos) - max_items})"
    return " | ".join(vistos) if vistos else "(sin frames con URL)"
 
 
def _es_error_frame_detached(exc: Exception) -> bool:
    msg = str(exc).lower()
    return (
        "frame was detached" in msg
        or "execution context was destroyed" in msg
        or "target page, context or browser has been closed" in msg
    )
 
 
def _reenganchar_frame(page, frame):
    try:
        if frame and not frame.is_detached():
            return frame
    except Exception:
        pass
    return esperar_frame_autoservicio(page)


def _esperar_tabla_libre(frame, timeout_ms: int = 12000):
    """Espera que el overlay de loading de jTable desaparezca antes de hacer click."""
    try:
        frame.locator("div.jtable-busy-panel-background").wait_for(state="hidden", timeout=timeout_ms)
    except Exception:
        pass
 
 
def recolectar_todas_las_cuentas(frame, page) -> list[str]:
    cuentas = []
    pagina = 1
    reintentos_detach = 0
    while True:
        try:
            frame = _reenganchar_frame(page, frame)
            if not frame:
                break
            frame.locator("tr.jtable-data-row").first.wait_for(state="visible", timeout=15000)
            filas = frame.locator("tr.jtable-data-row")
            for i in range(filas.count()):
                numero = filas.nth(i).locator("td").nth(0).inner_text().strip()
                if numero and numero not in cuentas:
                    cuentas.append(numero)
            btn_next = frame.locator("span.jtable-page-number-next")
            if btn_next.count() == 0:
                break
            clase = btn_next.first.get_attribute("class") or ""
            if "disabled" in clase:
                break
            _esperar_tabla_libre(frame)
            btn_next.first.click()
            page.wait_for_timeout(1500)
            frame.locator("tr.jtable-data-row").first.wait_for(state="visible", timeout=10000)
            pagina += 1
            reintentos_detach = 0
        except Exception as e:
            if _es_error_frame_detached(e) and reintentos_detach < 4:
                reintentos_detach += 1
                page.wait_for_timeout(900)
                frame = esperar_frame_autoservicio(page)
                continue
            raise
    return cuentas
 
 
def _seleccionar_cuenta_en_tabla(frame, page, cuenta_actual: str, max_paginas: int = 80) -> bool:
    paginas_recorridas = 0
    reintentos_detach = 0
    while paginas_recorridas < max_paginas:
        try:
            frame = _reenganchar_frame(page, frame)
            if not frame:
                return False
            frame.locator("tr.jtable-data-row").first.wait_for(state="visible", timeout=12000)
        except PlaywrightTimeoutError:
            return False
        except Exception as e:
            if _es_error_frame_detached(e) and reintentos_detach < 4:
                reintentos_detach += 1
                page.wait_for_timeout(900)
                frame = esperar_frame_autoservicio(page)
                continue
            return False
 
        filas = frame.locator("tr.jtable-data-row")
        try:
            total_filas = filas.count()
        except Exception:
            return False
 
        for j in range(total_filas):
            try:
                txt = filas.nth(j).locator("td").nth(0).inner_text(timeout=3000).strip()
            except Exception:
                continue
            if txt == cuenta_actual:
                try:
                    filas.nth(j).click(timeout=5000)
                    return True
                except Exception:
                    return False
 
        btn_next = frame.locator("span.jtable-page-number-next")
        if btn_next.count() == 0:
            return False
        clase = btn_next.first.get_attribute("class") or ""
        if "disabled" in clase:
            return False
        _esperar_tabla_libre(frame)
        btn_next.first.click()
        page.wait_for_timeout(900)
        paginas_recorridas += 1
        reintentos_detach = 0
 
    return False
 
 
def _esperar_contexto_subcuenta(page, frame, timeout_ms: int = 30000) -> bool:
    intentos = max(1, timeout_ms // 500)
    for _ in range(intentos):
        try:
            frame = _reenganchar_frame(page, frame)
            if not frame:
                page.wait_for_timeout(500)
                continue
            if frame.locator("text=/Informaci[oó]n\\s+de\\s+la\\s+cuenta/i").count() > 0:
                return True
            if frame.locator("text=/N[uú]mero\\s+de\\s+cuenta/i").count() > 0:
                return True
            if frame.locator("text=/Mis\\s+Facturas/i").count() > 0:
                return True
        except Exception as e:
            if not _es_error_frame_detached(e):
                pass
        page.wait_for_timeout(500)
    return False
 
 
def _goto_estable(page, url: str, wait_until: str = "domcontentloaded", timeout: int = 60000, reintentos: int = 5) -> bool:
    for _ in range(reintentos):
        try:
            page.goto(url, wait_until=wait_until, timeout=timeout)
            return True
        except Exception as e:
            msg = str(e).lower()
            if "interrupted by another navigation" in msg or "net::err_abort" in msg:
                try:
                    page.wait_for_load_state("domcontentloaded", timeout=10000)
                except Exception:
                    pass
                page.wait_for_timeout(1200)
                continue
            raise
    return False
 
 
def _detectar_modo_multicuentas(page, frame, timeout_ms: int = 25000) -> bool:
    intentos = max(1, timeout_ms // 500)
    for _ in range(intentos):
        try:
            frame = _reenganchar_frame(page, frame)
            if not frame:
                page.wait_for_timeout(500)
                continue
            filas = frame.locator("tr.jtable-data-row")
            if filas.count() > 0:
                return True
            hay_numero_cuenta = frame.locator("text=/N[uú]mero\\s+de\\s+cuenta/i").count() > 0
            hay_mis_facturas = frame.locator("text=/Mis\\s+Facturas/i").count() > 0
            if hay_numero_cuenta and hay_mis_facturas:
                return False
        except Exception as e:
            if not _es_error_frame_detached(e):
                pass
        page.wait_for_timeout(500)
    return True
 
 
# =========================
# Mis Facturas — Historial completo
# =========================

def abrir_mis_facturas_historial(page, frame, log):
    log("  Haciendo clic en 'Mis Facturas'...")
    frame.get_by_text("Mis Facturas").click()
    page.wait_for_timeout(1000)
 
    activado = False
    try:
        sel = frame.locator("select#getAll, select[name='getAll']")
        sel1 = frame.locator("Select#tipoDoc, select[name='tipoDoc']")
        sel.wait_for(state="visible", timeout=8000)
        sel1.wait_for(state="visible", timeout=8000)
        sel.select_option(value="true")
        sel1.select_option(label="Factura")
        page.wait_for_timeout(600)
        log("Solo facturas y Historial completo activado.")
        activado = True
    except PlaywrightTimeoutError:
        pass
    except Exception as e:
        log(f"  [WARN] select#getAll: {e}")
 
    if not activado:
        try:
            sel = frame.locator("select").filter(has_text=re.compile(r"[Hh]istorial"))
            sel1 = frame.locator("select").filter(has=frame.locator("option", has_text=re.compile(r"Factura", re.I)))
            if sel.count() > 0:
                sel.first.select_option(value="true")
                sel1.first.select_option(label="Factura")
                page.wait_for_timeout(600)
                log("Solo factursa y Historial completo activado (fallback).")
                activado = True
        except Exception:
            pass
 
    if not activado:
        log("  [WARN] No se pudo activar el filtrado de solo facturas y historial completo.")
 
    try:
        btn_buscar = frame.locator("#btnBuscar, button#btnBuscar")
        btn_buscar.first.wait_for(state="visible", timeout=8000)
        btn_buscar.first.click()
        page.wait_for_timeout(1500)
    except Exception:
        try:
            frame.get_by_role("button", name=re.compile("Buscar", re.I)).click()
            page.wait_for_timeout(1500)
        except Exception:
            pass
 
 
def _scroll_tabla_historial(frame) -> bool:
    candidates = [
        ".jtable-main-container",
        ".jtable-main-container > div",
        ".table-responsive",
        ".jtable",
        "#tblHistorialFacturas",
    ]
    for sel in candidates:
        try:
            box = frame.locator(sel)
            if box.count() == 0:
                continue
            moved = box.first.evaluate(
                """
                (el) => {
                    const maxTop = el.scrollHeight - el.clientHeight;
                    if (maxTop <= 0) return false;
                    const prev = el.scrollTop;
                    const step = Math.max(220, Math.floor(el.clientHeight * 0.8));
                    el.scrollTop = Math.min(maxTop, prev + step);
                    return el.scrollTop > prev;
                }
                """
            )
            if moved:
                return True
        except Exception:
            continue
    try:
        moved = frame.evaluate(
            """
            () => {
                const doc = document.scrollingElement || document.documentElement || document.body;
                if (!doc) return false;
                const maxTop = doc.scrollHeight - doc.clientHeight;
                if (maxTop <= 0) return false;
                const prev = doc.scrollTop;
                const step = Math.max(220, Math.floor(doc.clientHeight * 0.8));
                doc.scrollTop = Math.min(maxTop, prev + step);
                return doc.scrollTop > prev;
            }
            """
        )
        return bool(moved)
    except Exception:
        return False
 
 
def obtener_links_facturas_historial(frame, min_links: int = MAX_FACTURAS_VALIDAS) -> list[str]:
    try:
        frame.locator("a[href*='cmviewbill']").first.wait_for(state="attached", timeout=30000)
    except PlaywrightTimeoutError:
        return []
 
    hrefs = []
    sin_cambios = 0
    for _ in range(35):
        links = frame.locator("a[href*='cmviewbill']").filter(has_text=re.compile(r"\d"))
        before = len(hrefs)
        for i in range(links.count()):
            href = links.nth(i).get_attribute("href")
            if href:
                full = urljoin(frame.url, href)
                if full not in hrefs:
                    hrefs.append(full)
        if len(hrefs) >= min_links:
            break
        if len(hrefs) == before:
            sin_cambios += 1
        else:
            sin_cambios = 0
        if sin_cambios >= 3:
            break
        if not _scroll_tabla_historial(frame):
            break
        frame.wait_for_timeout(280)
 
    return hrefs
 
 
def hay_mas_paginas_historial(frame) -> bool:
    try:
        candidatos = [
            "span.jtable-page-number-next",
            "a.jtable-page-number-next",
            "li.next a",
            "a[aria-label='Next']",
            "a[title='Next']",
        ]
        for sel in candidatos:
            btn = frame.locator(sel)
            if btn.count() == 0:
                continue
            clase = (btn.first.get_attribute("class") or "").lower()
            aria_disabled = (btn.first.get_attribute("aria-disabled") or "").lower()
            parent_class = ""
            try:
                parent_class = (btn.first.evaluate("el => (el.parentElement && el.parentElement.className) || ''") or "").lower()
            except Exception:
                pass
            if "disabled" in clase or "disabled" in parent_class or aria_disabled == "true":
                continue
            return True
        return False
    except Exception:
        return False
 
 
def ir_a_siguiente_pagina_historial(frame, page, primer_href_anterior: str = "") -> bool:
    try:
        pagina_actual = None
        try:
            act = frame.locator(".jtable-page-number-active, .pagination .active")
            if act.count() > 0:
                pagina_actual = (act.first.inner_text() or "").strip()
        except Exception:
            pagina_actual = None
 
        if not primer_href_anterior:
            try:
                first_link = frame.locator("a[href*='cmviewbill']")
                if first_link.count() > 0:
                    primer_href_anterior = first_link.first.get_attribute("href") or ""
            except Exception:
                primer_href_anterior = ""
 
        clickeado = False
        for sel in [
            "span.jtable-page-number-next",
            "a.jtable-page-number-next",
            "li.next a",
            "a[aria-label='Next']",
            "a[title='Next']",
        ]:
            btn = frame.locator(sel)
            if btn.count() == 0:
                continue
            clase = (btn.first.get_attribute("class") or "").lower()
            aria_disabled = (btn.first.get_attribute("aria-disabled") or "").lower()
            parent_class = ""
            try:
                parent_class = (btn.first.evaluate("el => (el.parentElement && el.parentElement.className) || ''") or "").lower()
            except Exception:
                pass
            if "disabled" in clase or "disabled" in parent_class or aria_disabled == "true":
                continue
            _esperar_tabla_libre(frame)
            btn.first.click(timeout=6000)
            clickeado = True
            break
 
        if not clickeado:
            return False
 
        for _ in range(20):
            page.wait_for_timeout(300)
            frame = _reenganchar_frame(page, frame)
            if not frame:
                continue
            cambio_pagina = False
            try:
                if pagina_actual is not None:
                    act = frame.locator(".jtable-page-number-active, .pagination .active")
                    if act.count() > 0:
                        pagina_nueva = (act.first.inner_text() or "").strip()
                        if pagina_nueva and pagina_nueva != pagina_actual:
                            cambio_pagina = True
            except Exception:
                pass
            cambio_href = False
            try:
                first_link = frame.locator("a[href*='cmviewbill']")
                if first_link.count() > 0:
                    href_nuevo = first_link.first.get_attribute("href") or ""
                    if href_nuevo and href_nuevo != primer_href_anterior:
                        cambio_href = True
            except Exception:
                pass
            if cambio_pagina or cambio_href:
                return True
 
        return False
    except Exception:
        return False
 
 
# =========================
# Procesamiento por subcuenta
# =========================
def descargar_factura_por_url(context, page, bill_url: str, out_pdf: str, log) -> str | None:
    os.makedirs(os.path.dirname(out_pdf), exist_ok=True)
    pdf_path = descargar_pdf_con_cookies(context, bill_url, out_pdf)
    if pdf_path:
        return pdf_path
    try:
        page.goto(bill_url, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(1200)
        html = page.content()
        for u in extraer_posibles_urls_pdf(html, page.url):
            pdf_path = descargar_pdf_con_cookies(context, u, out_pdf)
            if pdf_path:
                return pdf_path
    except Exception as e:
        log(f"    [WARN] {e}")
    return None


def _safe_path_fragment(value: str) -> str:
    """Normaliza texto para usarlo en nombres de archivo/carpeta en Windows."""
    raw = str(value or "").strip()
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", raw)
    return cleaned or "sin_nombre"


def _renombrar_factura(pdf_path: str, cuenta_actual, datos: dict, n: int) -> str:
    """Renombra 'factura_NNN.pdf' a '{cuenta} {GS|TH}_{nro_factura}.pdf'.

    Si no se pudo parsear nro_factura, o el nombre resultante ya existe
    (colision), cae al correlativo n para garantizar que nunca se pise
    un archivo existente.
    """
    tipo_factura = datos.get("tipo_factura") or "TH"
    nro_factura_raw = (datos.get("nro_factura") or "").replace(" ", "")
    ident = _safe_path_fragment(nro_factura_raw) if nro_factura_raw else f"{n:03d}"
    out_dir = os.path.dirname(pdf_path)
    nuevo_nombre = f"{_safe_path_fragment(str(cuenta_actual))} {tipo_factura}_{ident}.pdf"
    nuevo_path = os.path.join(out_dir, nuevo_nombre)

    if os.path.exists(nuevo_path) and os.path.abspath(nuevo_path) != os.path.abspath(pdf_path):
        nuevo_path = os.path.join(out_dir, f"{nuevo_nombre[:-4]}_{n:03d}.pdf")

    try:
        os.replace(pdf_path, nuevo_path)
        return nuevo_path
    except OSError:
        return pdf_path
 
 
def procesar_subcuenta_historial(context, page, frame, cuenta_actual, output_dir, log) -> list[dict]:
    abrir_mis_facturas_historial(page, frame, log)
 
    cuenta_nombre = _safe_path_fragment(str(cuenta_actual))
    out_dir = os.path.join(output_dir, f"historial_{cuenta_nombre}")
    os.makedirs(out_dir, exist_ok=True)
 
    validas: list[dict] = []
    filas_excel: list[dict] = []
    descartadas = 0
    pagina = 1
    inicio = time.monotonic()
    max_segundos = 12 * 60
    max_intentos_descarga = 180
    intentos_descarga = 0
    urls_vistas: set[str] = set()
    paginas_sin_nuevos = 0
    firma_pagina_anterior: tuple[str, ...] = ()
    repeticiones_misma_pagina = 0
 
    while len(validas) < MAX_FACTURAS_VALIDAS:
        if (time.monotonic() - inicio) > max_segundos:
            log(f"  [{cuenta_actual}] Corte por tiempo. Se exporta parcial.")
            break
        if intentos_descarga >= max_intentos_descarga:
            log(f"  [{cuenta_actual}] Corte por intentos. Se exporta parcial.")
            break
 
        frame = _reenganchar_frame(page, frame)
        if not frame:
            break
 
        links = obtener_links_facturas_historial(frame)
        if not links:
            log(f"  [{cuenta_actual}] Sin links en pagina {pagina}.")
            break

        firma_pagina_actual = tuple(links[:5])
        if firma_pagina_actual and firma_pagina_actual == firma_pagina_anterior:
            repeticiones_misma_pagina += 1
        else:
            repeticiones_misma_pagina = 0
            firma_pagina_anterior = firma_pagina_actual

        if repeticiones_misma_pagina >= 1:
            log(f"  [{cuenta_actual}] Fin del historial detectado: la pagina no cambia al avanzar.")
            break

        links_nuevos = [u for u in links if u not in urls_vistas]
        if not links_nuevos:
            log(f"  [{cuenta_actual}] Pag {pagina}: sin links nuevos.")
            paginas_sin_nuevos += 1
        else:
            log(f"  [{cuenta_actual}] Pag {pagina}: {len(links_nuevos)} facturas nuevas.")
            paginas_sin_nuevos = 0

        for bill_url in links_nuevos:
            if len(validas) >= MAX_FACTURAS_VALIDAS:
                break
            if (time.monotonic() - inicio) > max_segundos:
                break
            if intentos_descarga >= max_intentos_descarga:
                break
 
            urls_vistas.add(bill_url)
            intentos_descarga += 1
 
            n = len(validas) + descartadas + 1
            out_pdf = os.path.join(out_dir, f"factura_{n:03d}.pdf")
 
            pdf_path = descargar_factura_por_url(context, page, bill_url, out_pdf, log)
            if not pdf_path or not es_pdf_real(pdf_path):
                descartadas += 1
                continue
 
            npags = contar_paginas_pdf(pdf_path)
            if npags <= 0:
                log(f"    [{cuenta_actual}] PDF invalido — descartada.")
                descartadas += 1
                continue

            try:
                datos = parse_factura_historial(pdf_path)
                parse_ok = True
            except Exception as e:
                log(f"    [{cuenta_actual}] Error parseando ({npags} pag(s)): {e}")
                datos = {"tipo_factura": detectar_tipo_factura(pdf_path)}
                parse_ok = False

            # El renombrado no depende de que el parseo haya sido exitoso: aun
            # si parse_factura_historial fallo, se dispone del tipo (TH/GS) y
            # el archivo no debe quedar con el nombre temporal "factura_NNN.pdf".
            datos["nro_cuenta"] = datos.get("nro_cuenta") or cuenta_actual
            datos["npags_pdf"] = npags
            pdf_path = _renombrar_factura(pdf_path, cuenta_actual, datos, n)
            datos["archivo"] = os.path.basename(pdf_path)

            if parse_ok:
                validas.append(datos)
                filas_excel.append(datos)
                log(f"    [{cuenta_actual}] {len(validas)}/{MAX_FACTURAS_VALIDAS} validas ({npags} pag(s)).")
            else:
                descartadas += 1

        if len(validas) >= MAX_FACTURAS_VALIDAS:
            break
        if (time.monotonic() - inicio) > max_segundos:
            break
        if intentos_descarga >= max_intentos_descarga:
            break

        if paginas_sin_nuevos >= 2:
            log(f"  [{cuenta_actual}] Fin del historial detectado: no aparecen facturas nuevas.")
            break

        primer_href_actual = links[0] if links else ""
        if not hay_mas_paginas_historial(frame):
            log(f"  [{cuenta_actual}] Fin del historial detectado: no hay mas paginas.")
            break

        if not ir_a_siguiente_pagina_historial(frame, page, primer_href_actual):
            log(f"  [{cuenta_actual}] No se pudo avanzar a la siguiente pagina del historial.")
            break

        pagina += 1
 
    log(
        f"  [{cuenta_actual}] Validas: {len(validas)} | Descartadas: {descartadas} | "
        f"Intentos: {intentos_descarga}"
    )
    return filas_excel
 
 
# =========================
# Guardar Excel consolidado
# =========================
def guardar_excel_consolidado(filas: list[dict], cuenta: str, output_dir: str, log):
    if not filas:
        log(f"  [{cuenta}] Sin datos para exportar.")
        return None
 
    df = pd.DataFrame(filas)
 
    COLS_TEXTO = {"nro_cuenta", "nro_medidor", "acuerdo_servicio", "nro_factura"}
    for col in COLS_TEXTO:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
            df[col] = df[col].str.replace(r"\.0$", "", regex=True)

    COLS_NO_NUMERICAS = COLS_TEXTO | {
        "archivo",
        "tipo_factura",
        "fecha_emision",
        "prox_vencimiento",
        "tarifa_aplicada",
        "potencia_contratada_punta_llano_kw",
        "tension",
        "fases",
        "direccion_servicio",
        "nombre_cliente",
        "direccion_cliente",
        "localidad_cliente",
        "departamento_cliente",
        "periodo_consumo",
        "zona_electrica",
    }
    for col in df.columns:
        if col in COLS_NO_NUMERICAS or col.endswith("_tipo_lec"):
            continue
        df[col] = df[col].apply(_to_excel_number)
 
    cols_fijas = [
        "archivo", "tipo_factura", "nro_cuenta", "nro_factura", "fecha_emision", "prox_vencimiento",
        "acuerdo_servicio", "nro_medidor", "tarifa_aplicada",
        "nombre_cliente", "direccion_cliente", "localidad_cliente", "departamento_cliente",
        "potencia_contratada_punta_llano_kw", "potencia_contratada_valle_kw", "potencia_contratada_kw",
        "consumo_activo_kwh", "consumo_reactivo_kvarh",
        "tension", "fases", "direccion_servicio", "periodo_consumo", "zona_electrica",
        "total_detalle_facturacion",
    ]
    cols_lecturas = sorted(c for c in df.columns if c not in cols_fijas)
    df = df[[c for c in cols_fijas if c in df.columns] + cols_lecturas]
 
    # Sufijo TH/GS segun el tipo de factura predominante del suministro (2
    # hojas o 1 hoja). Todas las facturas de una misma cuenta deberian ser del
    # mismo tipo; se usa el mas frecuente por si alguna quedo mal detectada.
    tipos_factura = [str(f.get("tipo_factura") or "").strip().upper() for f in filas]
    tipos_factura = [t for t in tipos_factura if t in ("TH", "GS")]
    tipo_factura_cuenta = max(set(tipos_factura), key=tipos_factura.count) if tipos_factura else "TH"

    cuenta_nombre = f"{_safe_path_fragment(str(cuenta))} {tipo_factura_cuenta}"
    excel_out = os.path.join(output_dir, f"{cuenta_nombre}.xlsx")
    os.makedirs(os.path.dirname(excel_out), exist_ok=True)

    # Si ya existe y está bloqueado, eliminarlo antes de escribir
    if os.path.exists(excel_out):
        try:
            os.remove(excel_out)
        except PermissionError:
            # Archivo abierto en Windows: renombrar el viejo y escribir el nuevo
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_viejo = os.path.join(output_dir, f"{cuenta_nombre}_OLD_{ts}.xlsx")
            try:
                os.rename(excel_out, excel_viejo)
                log(f"  [{cuenta}] [INFO] Archivo anterior renombrado a: {os.path.basename(excel_viejo)}")
            except Exception:
                pass  # Si tampoco se puede renombrar, intentar escribir igual
 
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.styles.numbers import FORMAT_TEXT
 
    with pd.ExcelWriter(excel_out, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Historial", index=False)
        ws = writer.sheets["Historial"]
 
        col_idx = {cell.value: cell.column for cell in ws[1]}
        for col_name in COLS_TEXTO:
            if col_name in col_idx:
                c = col_idx[col_name]
                for row in ws.iter_rows(min_row=2, min_col=c, max_col=c):
                    for cell in row:
                        cell.number_format = FORMAT_TEXT
                        if cell.value is not None:
                            cell.value = str(cell.value).replace(".0", "")
 
        hf = PatternFill("solid", start_color="1F4E79", end_color="1F4E79")
        hfont = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        dfont = Font(name="Arial", size=10)
 
        for cell in ws[1]:
            cell.font = hfont
            cell.fill = hf
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
 
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font = dfont
                if isinstance(cell.value, int):
                    cell.alignment = Alignment(horizontal="right")
                    cell.number_format = "0"
                elif isinstance(cell.value, float):
                    cell.alignment = Alignment(horizontal="right")
                    cell.number_format = "0.00"
                else:
                    cell.alignment = Alignment(horizontal="left")
 
        for col in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)
 
        ws.row_dimensions[1].height = 35
 
    log(f"  [{cuenta}] Excel guardado: {excel_out}")
    return excel_out
 
 
# =========================
# FUNCION PRINCIPAL
# =========================
def run_historial(
    ute_user: str,
    ute_pass: str,
    cuentas_seleccionadas: list[str],
    output_dir: str,
    log_fn=None,
):
    def log(msg):
        if log_fn:
            log_fn(msg)
        else:
            print(msg)
 
    if not ute_user or not ute_pass:
        log("ERROR: credenciales vacias.")
        return
 
    os.makedirs(output_dir, exist_ok=True)
    log(f"Guardando en: {output_dir}")
 
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(viewport={"width": 1600, "height": 900}, accept_downloads=True)
        context.set_default_timeout(15000)
        context.set_default_navigation_timeout(60000)
        page = context.new_page()
        page.set_default_timeout(15000)
        page.set_default_navigation_timeout(60000)
 
        try:
            log("Iniciando sesion en UTE...")
            if not _goto_estable(page, UTE_LOGIN, wait_until="domcontentloaded", timeout=60000):
                log("No se pudo abrir login de UTE.")
                return
            page.wait_for_selector("input#Username", timeout=60000)
            page.locator("input#Username").fill(ute_user)
            page.locator("input#Password").fill(ute_pass)
            btn = page.get_by_role("button", name=re.compile(r"^Ingresar$", re.I))
            btn.wait_for(state="visible", timeout=30000)
            # FIX: no_wait_after=True para evitar timeout por navegacion post-login
            btn.click(timeout=20000, no_wait_after=True)
            try:
                page.wait_for_load_state("networkidle", timeout=8000)
            except PlaywrightTimeoutError:
                pass
            page.wait_for_timeout(1200)
 
            if not _goto_estable(page, SERVICIOS_EN_LINEA, wait_until="domcontentloaded", timeout=60000):
                log("No se pudo abrir Servicios en linea.")
                return
            page.wait_for_timeout(3000)
 
            frame = esperar_frame_autoservicio(page)
            if not frame:
                log("No se encontro iframe de autoservicio.")
                log(f"Frames detectados: {_resumen_frames_debug(page)}")
                return
 
            multi = _detectar_modo_multicuentas(page, frame, timeout_ms=25000)
 
            if multi:
                log("Modo MULTI-CUENTAS.")
                todas = recolectar_todas_las_cuentas(frame, page)
                log(f"Subcuentas en portal: {len(todas)}")
                if not todas:
                    log("No se pudieron listar subcuentas. Probando CUENTA UNICA.")
                    multi = False
 
            if multi:
                a_procesar = (
                    [c for c in todas if c in cuentas_seleccionadas]
                    if cuentas_seleccionadas else todas
                )
                log(f"A procesar: {len(a_procesar)}")
 
                for cuenta_actual in a_procesar:
                    log(f"\n{'='*50}\n[{cuenta_actual}] Iniciando...")
                    datos: list[dict] = []
                    try:
                        if not _goto_estable(page, SERVICIOS_EN_LINEA, wait_until="domcontentloaded", timeout=60000):
                            log(f"[{cuenta_actual}] No se pudo abrir Servicios en linea.")
                            continue
                        page.wait_for_timeout(2500)
                        frame = esperar_frame_autoservicio(page)
                        if not frame:
                            log(f"[{cuenta_actual}] Sin iframe.")
                            continue
                        if not _seleccionar_cuenta_en_tabla(frame, page, cuenta_actual):
                            log(f"[{cuenta_actual}] No encontrada en tabla.")
                            continue
                        if not _esperar_contexto_subcuenta(page, frame, timeout_ms=30000):
                            log(f"[{cuenta_actual}] Timeout esperando carga.")
                            continue
                        page.wait_for_timeout(500)
                        frame = _reenganchar_frame(page, frame)
                        datos = procesar_subcuenta_historial(
                            context, page, frame, cuenta_actual, output_dir, log
                        )
                    except PlaywrightTimeoutError:
                        log(f"[{cuenta_actual}] Timeout.")
                    except Exception as e:
                        if _es_error_frame_detached(e):
                            log(f"[{cuenta_actual}] Frame recargado.")
                        else:
                            log(f"[{cuenta_actual}] Error: {e}")
                    finally:
                        try:
                            guardar_excel_consolidado(datos, cuenta_actual, output_dir, log)
                        except Exception as e:
                            log(f"[{cuenta_actual}] Error guardando Excel: {e}")
 
            if not multi:
                log("Modo CUENTA UNICA.")
                try:
                    cuenta_actual = (
                        frame.locator("td:has-text('Numero de cuenta') + td")
                        .first.inner_text().strip()
                    )
                except Exception:
                    cuenta_actual = ute_user
                if not _esperar_contexto_subcuenta(page, frame, timeout_ms=30000):
                    log("Timeout esperando carga de cuenta unica.")
                    return
                page.wait_for_timeout(500)
                datos: list[dict] = []
                try:
                    datos = procesar_subcuenta_historial(
                        context, page, frame, cuenta_actual, output_dir, log
                    )
                finally:
                    guardar_excel_consolidado(datos, cuenta_actual, output_dir, log)
 
            log("\nProceso finalizado.")
 
        except Exception as e:
            if _es_error_frame_detached(e):
                log("Iframe recargado al finalizar.")
            else:
                log(f"Error inesperado: {e}")
        finally:
            browser.close()
 
 
# =========================
# MAIN — consola
# =========================
def _debug_parsear_pdfs(paths: list[str]):
    """Corre parse_factura_historial sobre PDFs locales e imprime el resultado.

    Util para validar cambios de parsing (GS/TH) sin correr el scraping real.
    Uso: python historial_facturas_ute.py debug "archivo1.pdf" "archivo2.pdf"
    """
    for path in paths:
        print("=" * 70)
        print(path)
        if not os.path.exists(path):
            print("  [ERROR] archivo no encontrado")
            continue
        datos = parse_factura_historial(path)
        for k, v in datos.items():
            print(f"  {k}: {v!r}")


if __name__ == "__main__":
    import sys as _sys
    if len(_sys.argv) > 1 and _sys.argv[1] == "debug":
        _debug_parsear_pdfs(_sys.argv[2:])
    else:
        ute_user = input("Usuario UTE: ").strip()
        ute_pass = input("Contrasena UTE: ").strip()
        output_dir = os.path.join(os.getcwd(), ute_user)
        run_historial(ute_user, ute_pass, cuentas_seleccionadas=[], output_dir=output_dir)
