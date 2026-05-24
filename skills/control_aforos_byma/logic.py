"""
Control Aforos BYMA vs Gallo
Compara el PDF de especies aceptadas como garantía (circular BYMA)
con el maestro de especies ESPECIES.XLS.
"""

import io
import re
import xlrd
import pdfplumber
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─── Tablas Lista <-> Aforo BYMA ─────────────────────────────────────────────
TABLAS = {
    "Renta Variable": {
        1: 85, 2: 80, 3: 75, 4: 70, 5: 60, 6: 50, 7: 40, 8: 30,
    },
    "Renta Fija Publicos": {
        11: 90, 12: 85, 13: 80, 14: 75, 15: 70, 16: 65, 17: 60,
    },
    "Renta Fija Privados": {
        22: 85, 23: 80, 24: 75, 25: 70, 26: 65, 27: 60,
    },
    "Letras y Bonos del tesoro": {
        85: 85, 90: 90, 95: 95,
    },
}

SECCION_A_TABLA = {
    ("Renta Fija",     "Titulos Publicos"):           "Renta Fija Publicos",
    ("Renta Fija",     "Obligaciones Negociables"):   "Renta Fija Privados",
    ("Renta Fija",     "Letras del Tesoro Nacional"): "Letras y Bonos del tesoro",
    ("Renta Variable", "Acciones Locales"):            "Renta Variable",
    ("Renta Variable", "Cedears"):                     "Renta Variable",
}

# Identificadores de sección normalizados (ASCII, minúsculas)
SECCION_HEADERS = {
    "renta fija":                  ("tipo",    "Renta Fija"),
    "ttulos pblicos":              ("subtipo", "Titulos Publicos"),   # "Títulos Públicos" mangled
    "obligaciones negociables":    ("subtipo", "Obligaciones Negociables"),
    "letras del tesoro nacional":  ("subtipo", "Letras del Tesoro Nacional"),
    "renta variable":              ("tipo",    "Renta Variable"),
    "acciones locales":            ("subtipo", "Acciones Locales"),
    "cedears":                     ("subtipo", "Cedears"),
}

STOP_PATTERNS = ["cuotapartes de", "fondos comunes de"]

# Regex para fila de especie: TICKER  CVSA  AFORO% [...]
ESPECIE_RE = re.compile(r"^([A-Za-z0-9]+)\s+(\d+)\s+(\d+)%")


def _ascii_only(s):
    return re.sub(r"[^a-zA-Z0-9 ]", "", s).lower().strip()


def _aforo_para_lista(tabla, lista):
    return TABLAS.get(tabla, {}).get(lista)


def _lista_para_aforo(tabla, aforo_pct):
    for lista, aforo in TABLAS.get(tabla, {}).items():
        if aforo == aforo_pct:
            return lista
    return "REVISAR"


def _xls_int(cell):
    if cell.ctype == xlrd.XL_CELL_NUMBER:
        return int(cell.value)
    if cell.ctype == xlrd.XL_CELL_TEXT:
        v = str(cell.value).strip().lstrip("'")
        try:
            return int(float(v))
        except ValueError:
            return None
    return None


def generar_control(especies_file, pdf_file):
    """
    Parámetros
    ----------
    especies_file : file-like  — ESPECIES.XLS
    pdf_file      : file-like  — PDF aforos BYMA

    Retorna
    -------
    xlsx_bytes       : bytes          Excel con diferencias
    txt_faltantes    : str | None     Texto FALTANTE en GALLO (None si no hay)
    resumen          : dict           Conteos para la UI
    advertencias     : list[str]      Avisos de procesamiento
    """
    advertencias = []

    # ── 1. Leer ESPECIES.XLS ─────────────────────────────────────────────────
    especies_file.seek(0)
    wb_xls = xlrd.open_workbook(file_contents=especies_file.read())
    hoja = ("Datos_Fijos_Especies"
            if "Datos_Fijos_Especies" in wb_xls.sheet_names()
            else wb_xls.sheet_names()[0])
    ws_xls = wb_xls.sheet_by_name(hoja)

    especies_dict  = {}   # cvsa_int -> lista_int
    codigos_vistos = {}
    for row_idx in range(1, ws_xls.nrows):
        cell_a = ws_xls.cell(row_idx, 0)   # Código CVSA
        cell_f = ws_xls.cell(row_idx, 5)   # Lista
        if cell_a.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
            continue
        cvsa_int = _xls_int(cell_a)
        if cvsa_int is None:
            advertencias.append(f"Código CVSA no numérico fila {row_idx+1} — omitida.")
            continue
        if cvsa_int in codigos_vistos:
            advertencias.append(
                f"Código CVSA duplicado: {cvsa_int} "
                f"filas {codigos_vistos[cvsa_int]+1} y {row_idx+1}."
            )
        codigos_vistos[cvsa_int] = row_idx
        lista_val = _xls_int(cell_f)
        if lista_val is None:
            lista_val = 0
        especies_dict[cvsa_int] = lista_val

    # ── 2. Parsear PDF línea a línea ─────────────────────────────────────────
    pdf_file.seek(0)
    pdf_name    = getattr(pdf_file, "name", "aforos.pdf")
    fecha_match = re.search(r"(\d{4}-\d{2}-\d{2})", pdf_name)
    ref_circular = fecha_match.group(1) if fecha_match else pdf_name

    especies_pdf   = []
    tipo_actual    = None
    subtipo_actual = None

    with pdfplumber.open(pdf_file) as pdf:
        stop = False
        for page in pdf.pages:
            if stop:
                break
            text       = page.extract_text() or ""
            text_ascii = _ascii_only(text)
            for pat in STOP_PATTERNS:
                if pat in text_ascii:
                    stop = True
                    break

            for line in text.split("\n"):
                line_s    = line.strip()
                line_norm = _ascii_only(line_s)
                if not line_s:
                    continue

                # Detectar encabezado de sección
                if line_norm in SECCION_HEADERS:
                    campo, valor = SECCION_HEADERS[line_norm]
                    if campo == "tipo":
                        tipo_actual    = valor
                        subtipo_actual = None
                    else:
                        subtipo_actual = valor
                    continue

                # Stop pattern en línea individual
                for pat in STOP_PATTERNS:
                    if pat in line_norm:
                        stop = True
                        break
                if stop:
                    break

                # Intentar parsear como especie
                m = ESPECIE_RE.match(line_s)
                if not m:
                    continue
                ticker     = m.group(1)
                cvsa_int   = int(m.group(2))
                aforo_byma = int(m.group(3))

                if not tipo_actual or not subtipo_actual:
                    continue
                tabla_nombre = SECCION_A_TABLA.get((tipo_actual, subtipo_actual))
                if tabla_nombre is None:
                    advertencias.append(
                        f"Sección no reconocida: '{tipo_actual}'/'{subtipo_actual}'. "
                        f"'{ticker}' ({cvsa_int}) omitida."
                    )
                    continue
                especies_pdf.append({
                    "nombre":     ticker,
                    "cvsa":       cvsa_int,
                    "tipo":       tipo_actual,
                    "subtipo":    subtipo_actual,
                    "tabla":      tabla_nombre,
                    "aforo_byma": aforo_byma,
                })

    # ── 3. Faltantes en Gallo ────────────────────────────────────────────────
    faltantes = [esp for esp in especies_pdf if esp["cvsa"] not in especies_dict]

    txt_faltantes = None
    if faltantes:
        lines = [
            f"Control de Especies - Circular BYMA {ref_circular}",
            "=" * 60, "",
            "Las siguientes especies aparecen en el PDF de BYMA pero NO",
            "estan registradas en el maestro ESPECIES.XLS de Gallo:", "",
            f"{'Especie':<15} {'Tipo':<15} {'Subtipo':<25} {'Codigo CVSA':>12}",
            "-" * 70,
        ]
        for esp in faltantes:
            lines.append(
                f"{esp['nombre']:<15} {esp['tipo']:<15} {esp['subtipo']:<25} {esp['cvsa']:>12}"
            )
        lines += [
            "", "-" * 70,
            f"Total faltantes: {len(faltantes)}", "",
            "ACCION REQUERIDA: Dar de alta estas especies en Gallo",
            "y actualizar ESPECIES.XLS.",
        ]
        txt_faltantes = "\n".join(lines)

    # ── 4. Control de aforos ─────────────────────────────────────────────────
    diferencias         = []
    no_encontradas_ctrl = []

    for esp in especies_pdf:
        cvsa = esp["cvsa"]
        if cvsa not in especies_dict:
            no_encontradas_ctrl.append(esp)
            continue
        lista_actual = especies_dict[cvsa]
        tabla_nombre = esp["tabla"]
        aforo_byma   = esp["aforo_byma"]

        # Lista 0 o no mapeada → especie sin lista asignada, no es diferencia
        if lista_actual == 0 or lista_actual not in TABLAS.get(tabla_nombre, {}):
            continue

        aforo_sailing = _aforo_para_lista(tabla_nombre, lista_actual)
        if aforo_sailing != aforo_byma:
            lista_sugerida = _lista_para_aforo(tabla_nombre, aforo_byma)
            diferencia_pp  = aforo_byma - aforo_sailing
            diferencias.append({
                "nombre":         esp["nombre"],
                "tipo":           esp["tipo"],
                "subtipo":        esp["subtipo"],
                "cvsa":           cvsa,
                "aforo_byma":     aforo_byma,
                "aforo_sailing":  aforo_sailing,
                "lista_actual":   lista_actual,
                "diferencia_pp":  diferencia_pp,
                "lista_sugerida": lista_sugerida,
            })
            if lista_sugerida == "REVISAR":
                advertencias.append(
                    f"Lista sugerida REVISAR: '{esp['nombre']}' ({cvsa}), "
                    f"aforo BYMA={aforo_byma}%, no existe en tabla '{tabla_nombre}'."
                )

    # ── 5. Generar Excel ─────────────────────────────────────────────────────
    AZUL_FILL   = PatternFill("solid", fgColor="4472C4")
    ROJO_FILL   = PatternFill("solid", fgColor="C00000")
    VERDE_FILL  = PatternFill("solid", fgColor="00B050")
    ROJO_CELDA  = PatternFill("solid", fgColor="FF0000")
    BLANCO_FONT = Font(name="Arial", bold=True, color="FFFFFF")
    NORMAL_FONT = Font(name="Arial", size=11)
    BORDER_THIN = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin"),
    )
    CENTER = Alignment(horizontal="center", vertical="center")

    def aplicar_encabezado(ws, headers, fill=None):
        f = fill or AZUL_FILL
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.fill = f
            cell.font = BLANCO_FONT
            cell.alignment = CENTER
            cell.border = BORDER_THIN

    def aplicar_fila(ws, fila, valores, fills=None, fonts=None):
        for col, v in enumerate(valores, 1):
            cell = ws.cell(row=fila, column=col, value=v)
            cell.font      = (fonts or {}).get(col, NORMAL_FONT)
            cell.alignment = CENTER
            cell.border    = BORDER_THIN
            if fills and fills.get(col):
                cell.fill = fills[col]

    wb = openpyxl.Workbook()

    # Hoja "Diferencias Aforo"
    ws_dif = wb.active
    ws_dif.title = "Diferencias Aforo"
    headers_dif = [
        "Especie", "Tipo", "Subtipo", "Codigo CVSA",
        "Aforo BYMA (PDF) %", "Aforo Sailing (Lista) %",
        "Lista Actual", "Diferencia (pp)", "LISTA SUGERIDA",
    ]
    aplicar_encabezado(ws_dif, headers_dif)
    ws_dif.cell(1, 9).fill = VERDE_FILL  # col LISTA SUGERIDA en verde

    for i, d in enumerate(diferencias, 2):
        diff_str = (f"+{d['diferencia_pp']}" if d["diferencia_pp"] > 0
                    else str(d["diferencia_pp"])) + "pp"
        aplicar_fila(ws_dif, i, [
            d["nombre"], d["tipo"], d["subtipo"], d["cvsa"],
            d["aforo_byma"], d["aforo_sailing"],
            d["lista_actual"], diff_str, d["lista_sugerida"],
        ], fills={6: ROJO_CELDA, 9: VERDE_FILL},
           fonts={9: Font(name="Arial", color="FFFFFF", bold=True)})

    col_widths = [15, 15, 25, 14, 20, 22, 14, 16, 16]
    for col, w in enumerate(col_widths, 1):
        ws_dif.column_dimensions[get_column_letter(col)].width = w

    # Hoja "No Encontradas" (especies del PDF no presentes en Gallo)
    if no_encontradas_ctrl:
        ws_nf = wb.create_sheet("No Encontradas")
        headers_nf = [
            "Especie", "Tipo", "Subtipo",
            "Codigo CVSA", "Aforo BYMA (PDF) %", "LISTA SUGERIDA",
        ]
        aplicar_encabezado(ws_nf, headers_nf, fill=ROJO_FILL)
        ws_nf.cell(1, 6).fill = VERDE_FILL
        for i, esp in enumerate(no_encontradas_ctrl, 2):
            lista_sug = _lista_para_aforo(esp["tabla"], esp["aforo_byma"])
            aplicar_fila(ws_nf, i, [
                esp["nombre"], esp["tipo"], esp["subtipo"],
                esp["cvsa"], esp["aforo_byma"], lista_sug,
            ], fills={6: VERDE_FILL},
               fonts={6: Font(name="Arial", color="FFFFFF", bold=True)})
        col_widths_nf = [15, 15, 25, 14, 22, 16]
        for col, w in enumerate(col_widths_nf, 1):
            ws_nf.column_dimensions[get_column_letter(col)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    xlsx_bytes = buf.getvalue()

    resumen = {
        "pdf_name":    pdf_name,
        "total_pdf":   len(especies_pdf),
        "faltantes":   len(faltantes),
        "diferencias": len(diferencias),
    }

    return xlsx_bytes, txt_faltantes, resumen, advertencias
