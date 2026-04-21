import re
import io
import pdfplumber

_DOC_RE = re.compile(r"\b(CC|CE|PAS|TI|RC|TE|CD|SC|NIT)\s*[:\-]?\s*(\d{5,15})\b", re.IGNORECASE)
_DATE_RE = re.compile(r"\b(\d{4}-\d{2}-\d{2})\b")
_MONEY_RE = re.compile(r"\b(\d{1,3}(?:\.\d{3})*,\d{2})\b")
_ESTADO_RE = re.compile(r"\b(Autorizada|Rechazada|Pendiente)\b", re.IGNORECASE)

# Repeated sub-header text that SIIF embeds inside each person's merged cell
_SUB_HEADER_RE = re.compile(
    r"Fecha\s+Inicial\s+Comisi[oó]n\s+Fecha\s+final\s+Comisi[oó]n"
    r"[\s\S]*?"
    r"Porcentaj[ae]\s+Pernocta",
    re.IGNORECASE,
)

# Location pattern: "DEPT D.C. / CITY" — stops before next ALL-CAPS token or digit
_LOC_RE = re.compile(
    r"([A-ZÁÉÍÓÚÜÑ][A-ZÁÉÍÓÚÜÑ\.\s]*?)\s*/\s*([A-ZÁÉÍÓÚÜÑ][A-ZÁÉÍÓÚÜÑ\.\s]*?)"
    r"(?=\s+[A-ZÁÉÍÓÚÜÑ]|\s+\d|$)",
    re.MULTILINE,
)


# Labels SIIF repeats as sub-headers inside table cells
_CELL_LABELS_RE = re.compile(
    r"Regi[oó]n\s+o\s+Depto\s+Origen|Ciudad\s+o\s+Muni\s+Destino"
    r"|Fecha\s+Inicial\s+Comisi[oó]n|Fecha\s+final\s+Comisi[oó]n"
    r"|N[°o][\.°]?\s*D[ií]as|Pernocta\s+[ÚU]ltimo\s+d[ií]a\s+Comisi[oó]n"
    r"|Porcentaj[ae]\s+Pernocta",
    re.IGNORECASE,
)


def _clean(text):
    if not text:
        return ""
    return " ".join(str(text).split()).strip()


def _clean_cell(text):
    """Remove SIIF sub-header labels that appear inside merged table cells."""
    if not text:
        return ""
    return _clean(_CELL_LABELS_RE.sub(" ", str(text)))


def _find_solicitud_no(text):
    m = re.search(r"Solicitud\s+de\s+Comisi[oó]n\s+No[\.:]?\s*(\d+)", text, re.IGNORECASE)
    return m.group(1) if m else ""


def _find_person_section(text):
    """Return the slice of text that contains the comisionado rows."""
    start_m = re.search(r"Objeto\s+de\s+la\s+Comisi[oó]n\s+por\s+Tercero", text, re.IGNORECASE)
    end_m = re.search(r"Totales\s+Solicitud\s+de\s+Comisi[oó]n", text, re.IGNORECASE)
    if start_m and end_m and start_m.end() < end_m.start():
        return text[start_m.end(): end_m.start()]
    # Fallback: whole text (person detection will still filter correctly)
    return text


def _extract_locations(text_after_dates):
    """Return (origen, destino) from the text immediately following the dates."""
    matches = list(_LOC_RE.finditer(text_after_dates.strip()))
    if len(matches) >= 2:
        origen = _clean(matches[0].group(1) + " / " + matches[0].group(2))
        destino = _clean(matches[1].group(1) + " / " + matches[1].group(2))
    elif len(matches) == 1:
        origen = _clean(matches[0].group(1) + " / " + matches[0].group(2))
        destino = ""
    else:
        origen = ""
        destino = ""
    return origen, destino


def _parse_person_block(block, solicitud_no):
    """Parse one comisionado from a text block anchored at its doc-type match."""
    doc_m = _DOC_RE.search(block)
    if not doc_m:
        return None

    doc_type = doc_m.group(1).upper()
    doc_num = doc_m.group(2)

    # Name: last all-caps sequence before the doc match
    pre_doc = block[: doc_m.start()].strip()
    name_m = re.search(r"([A-ZÁÉÍÓÚÜÑ][A-ZÁÉÍÓÚÜÑ\s\-]+)$", pre_doc)
    nombre = _clean(name_m.group(1)) if name_m else ""

    post_doc = block[doc_m.end():]

    # Cargo is between the doc number and the estado word
    estado_m = _ESTADO_RE.search(post_doc)
    if estado_m:
        cargo = _clean(post_doc[: estado_m.start()])
        after_estado = post_doc[estado_m.end():]
    else:
        cargo = _clean(post_doc[:60])
        after_estado = post_doc[60:]

    # Strip repeated SIIF sub-headers
    after_estado = _SUB_HEADER_RE.sub(" ", after_estado)

    # Dates
    dates = _DATE_RE.findall(after_estado)
    fecha_ini = dates[0] if len(dates) > 0 else ""
    fecha_fin = dates[1] if len(dates) > 1 else ""

    # Text after both dates → locations + numbers + object
    date_matches = list(_DATE_RE.finditer(after_estado))
    if len(date_matches) >= 2:
        after_dates_text = after_estado[date_matches[1].end():]
    elif date_matches:
        after_dates_text = after_estado[date_matches[0].end():]
    else:
        after_dates_text = after_estado

    origen, destino = _extract_locations(after_dates_text)

    # Money values — last one is "Valor total a pagar"
    money_vals = _MONEY_RE.findall(after_estado)
    valor = money_vals[-1] if money_vals else ""

    # Object: text after the last money value
    objeto = ""
    if valor:
        last_pos = after_estado.rfind(valor)
        if last_pos >= 0:
            raw_obj = after_estado[last_pos + len(valor):]
            raw_obj = re.split(r"\s+(?:Totales|OBJETO|ORDENADOR|Firma|P[áa]gina)\b", raw_obj, flags=re.IGNORECASE)[0]
            objeto = _clean(raw_obj)

    if not (nombre or doc_num):
        return None

    return {
        "Solicitud de Comisión No.": solicitud_no,
        "Nombre": nombre,
        "Tipo de Documento": doc_type,
        "Número de Documento": doc_num,
        "Cargo": cargo,
        "Fecha Inicial Comisión": fecha_ini,
        "Fecha Final Comisión": fecha_fin,
        "Dpto. / Municipio Origen": origen,
        "Dpto. / Municipio Destino": destino,
        "Valor Total a Pagar": valor,
        "Objeto de la Comisión": objeto,
    }


def _parse_via_text(full_text, solicitud_no):
    """Text-based extraction (primary or fallback)."""
    section = _find_person_section(full_text)
    section = _SUB_HEADER_RE.sub(" ", section)

    doc_matches = list(_DOC_RE.finditer(section))
    if not doc_matches:
        return []

    persons = []
    for i, dm in enumerate(doc_matches):
        # Block: from 300 chars before this doc match to just before the next one
        look_back = max(0, dm.start() - 300)
        if i > 0:
            look_back = max(look_back, doc_matches[i - 1].end())
        block_end = doc_matches[i + 1].start() if i + 1 < len(doc_matches) else len(section)

        block = section[look_back:block_end]
        person = _parse_person_block(block, solicitud_no)
        if person:
            persons.append(person)

    return persons


def _col_map_from_header(header):
    """Build a dict of column indices from a header row (list of lowercase strings)."""
    def find(*kws):
        for i, h in enumerate(header):
            if all(k in h for k in kws):
                return i
        return None

    return {
        "nombre":    0,
        "doc":       find("documento") or find("tipo") or 1,
        "cargo":     find("cargo") or 2,
        "fecha_ini": find("inicial"),
        "fecha_fin": find("final") or find("fin "),
        "origen":    find("origen"),
        "destino":   find("destino"),
        "valor":     find("pagar"),
        "objeto":    find("objeto"),
    }


def _person_from_row(row, cols, solicitud_no):
    """Extract one person dict from a table row using a column-index map."""
    row_text = " ".join(str(c or "") for c in row)
    if not (_DATE_RE.search(row_text) or _DOC_RE.search(row_text)):
        return None

    def get(key):
        idx = cols.get(key)
        if idx is not None and idx < len(row) and row[idx]:
            return _clean(str(row[idx]))
        return ""

    doc_cell = get("doc")
    doc_m = _DOC_RE.search(doc_cell)
    doc_type = doc_m.group(1).upper() if doc_m else ""
    doc_num  = doc_m.group(2) if doc_m else ""

    nombre = _clean_cell(get("nombre"))
    cargo  = _clean_cell(get("cargo"))

    fecha_ini = (_DATE_RE.findall(get("fecha_ini")) or [""])[0]
    fecha_fin = (_DATE_RE.findall(get("fecha_fin")) or [""])[0]
    if not fecha_ini or not fecha_fin:
        all_dates = _DATE_RE.findall(row_text)
        fecha_ini = fecha_ini or (all_dates[0] if all_dates else "")
        fecha_fin = fecha_fin or (all_dates[1] if len(all_dates) > 1 else "")

    origen  = _clean_cell(get("origen"))
    destino = _clean_cell(get("destino"))

    valor = get("valor")
    if not valor:
        mv = _MONEY_RE.findall(row_text)
        valor = mv[-1] if mv else ""

    objeto = get("objeto")

    if not (nombre and (doc_type or fecha_ini)):
        return None

    return {
        "Solicitud de Comisión No.": solicitud_no,
        "Nombre": nombre,
        "Tipo de Documento": doc_type,
        "Número de Documento": doc_num,
        "Cargo": cargo,
        "Fecha Inicial Comisión": fecha_ini,
        "Fecha Final Comisión": fecha_fin,
        "Dpto. / Municipio Origen": origen,
        "Dpto. / Municipio Destino": destino,
        "Valor Total a Pagar": valor,
        "Objeto de la Comisión": objeto,
    }


def _parse_via_tables(pdf, solicitud_no):
    """Table-based extraction using pdfplumber.

    The column map learned from a header row is kept across pages so that
    continuation tables (page 2+, no repeated header) are processed correctly.
    """
    persons = []
    HEADER_KWS = ["nombre", "cargo", "fecha"]
    saved_cols = None   # remembered across pages

    for page in pdf.pages:
        for table in page.extract_tables() or []:
            if not table:
                continue

            # Try to find a header row in this table
            header_idx = None
            for i, row in enumerate(table):
                row_text = " ".join(str(c or "").lower() for c in row)
                if all(kw in row_text for kw in HEADER_KWS):
                    header_idx = i
                    break

            if header_idx is not None:
                saved_cols = _col_map_from_header(
                    [str(c or "").lower() for c in table[header_idx]]
                )
                data_rows = table[header_idx + 1:]
            elif saved_cols is not None:
                # Continuation page — reuse the column map from the previous header
                data_rows = table
            else:
                continue

            for row in data_rows:
                person = _person_from_row(row, saved_cols, solicitud_no)
                if person:
                    persons.append(person)

    return persons


def extract_commission_data(pdf_bytes, filename=""):
    """
    Extract one dict per comisionado from a SIIF PDF.
    Returns (list[dict], list[str] warnings).
    """
    warnings = []
    persons = []

    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)
            solicitud_no = _find_solicitud_no(full_text)
            if not solicitud_no:
                warnings.append(f"{filename}: No se encontró 'Solicitud de Comisión No.'")

            # Try table extraction first; fall back to text parsing
            persons = _parse_via_tables(pdf, solicitud_no)
            if not persons:
                persons = _parse_via_text(full_text, solicitud_no)

            if not persons:
                warnings.append(f"{filename}: No se encontraron comisionados en el documento.")

    except Exception as exc:
        warnings.append(f"{filename}: Error al procesar — {exc}")

    return persons, warnings
