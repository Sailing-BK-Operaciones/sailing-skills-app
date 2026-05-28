"""
Control Aforos BYMA vs Gallo
Lee haircut desde col 26 (Aforo) de ESPECIES.XLS — informado por la API BYMA via Gallo.
Aforo BYMA = 100 - haircut  (ej. haircut 15% -> aforo 85%)
El PDF de circular BYMA ya no se utiliza.
"""

import io
import re
import datetime

import xlrd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─── Tablas Lista <-> Aforo BYMA ─────────────────────────────────────────────
# Aforo = 100 - haircut
TABLAS = {
    "Renta Variable": {
        1: 85, 2: 80, 3: 75, 4: 70, 5: 60, 6: 50, 7: 40, 8: 30,
    },
    "Renta Fija Publicos": {
        10: 95, 11: 90, 12: 85, 13: 80, 14: 75, 15: 70, 16: 65, 17: 60,
    },
    "Renta Fija Privados": {
        22: 85, 23: 80, 24: 75, 25: 70, 26: 65, 27: 60,
    },
    "Letras y Bonos del tesoro": {
        85: 85, 90: 90, 95: 95,
    },
}


def _tabla_para_lista(lista):
    if 1 <= lista <= 8:       return "Renta Variable"
    if 10 <= lista <= 17:     return "Renta Fija Publicos"
    if 22 <= lista <= 27:     return "Renta Fija Privados"
    if lista in (85, 90, 95): return "Letras y Bonos del tesoro"
    return None


def _tabla_para_tipo(tipo_col2, tipo_activo):
    t2 = (tipo_col2 or "").strip().upper()
    ta = (tipo_activo or "").lower()
    if t2 == "LETR":
        return "Letras y Bonos del tesoro"
    if t2 == "O/N":
        return "Renta Fija Privados"
    if t2 == "PUBL":
        if any(x in ta for x in ["letra", "tesoro", "17-letra"]):
            return "Letras y Bonos del tesoro"
        return "Renta Fija Publicos"
    if t2 in ("PRIV", "FDO."):
        return "Renta Variable"
    if t2 == "FIDE":
        return "Renta Fija Publicos"
    if any(x in ta for x in ["obliga", "negoci", "05-obli"]):
        return "Renta Fija Privados"
    if any(x in ta for x in ["letra", "tesoro"]):
        return "Letras y Bonos del tesoro"
    if any(x in ta for x in ["publi", "bono", "provincial", "03-titulo", "titulos de deuda", "19-titulo"]):
        return "Renta Fija Publicos"
    return "Renta Variable"


def _aforo_para_lista(tabla, lista):
    return TABLAS.get(tabla, {}).get(lista)


def _lista_para_aforo(tabla, aforo_pct):
    for l, a in TABLAS.get(tabla, {}).items():
        if a == aforo_pct:
            return l
    return "REVISAR"


# ─── Helpers xlrd ────────────────────────────────────────────────────────────
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


def _xls_str(cell):
    if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        return ""
    return str(cell.value).strip()


_DATE_RE = re.compile(r'\b(\d{2})[/\-](\d{2})[/\-](\d{2,4})\b')


def _fecha_vencimiento(wb_xls, ws_xls, row_idx):
    """Devuelve date de vencimiento o None."""
    cell_v = ws_xls.cell(row_idx, 15)
    if cell_v.ctype == xlrd.XL_CELL_DATE:
        try:
            return xlrd.xldate_as_datetime(cell_v.value, wb_xls.datemode).date()
        except Exception:
            pass
    elif cell_v.ctype == xlrd.XL_CELL_NUMBER and cell_v.value > 1000:
        try:
            return xlrd.xldate_as_datetime(cell_v.value, wb_xls.datemode).date()
        except Exception:
            pass
    elif cell_v.ctype == xlrd.XL_CELL_TEXT:
        s = str(cell_v.value).strip()
        for fmt in ("%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y"):
            try:
                return datetime.datetime.strptime(s, fmt).date()
            except ValueError:
                pass
    nombre = _xls_str(ws_xls.cell(row_idx, 1))
    m = _DATE_RE.search(nombre)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        try:
            return datetime.date(y, mo, d)
        except ValueError:
            pass
    return None


def generar_control(especies_file):
    """
    Parámetros
    ----------
    especies_file : file-like  — ESPECIES.XLS

    Retorna
    -------
    xlsx_bytes   : bytes     Excel con diferencias
    resumen      : dict      Conteos para la UI
    advertencias : list[str] Avisos de procesamiento
    """
    HOY = datetime.date.today()
    advertencias = []

    # ── Leer ESPECIES.XLS ────────────────────────────────────────────────────
    especies_file.seek(0)
    wb_xls = xlrd.open_workbook(file_contents=especies_file.read())
    hoja = ("Datos_Fijos_Especies"
            if "Datos_Fijos_Especies" in wb_xls.sheet_names()
            else wb_xls.sheet_names()[0])
    ws_xls = wb_xls.sheet_by_name(hoja)

    diferencias    = []   # haircut > 0, lista > 0, pero lista no coincide
    sin_lista      = []   # haircut > 0, lista = 0  → necesita lista asignada en Gallo
    lista_sin_byma = []   # lista > 0, haircut = 0, NO vencido → informacional
    vencidos_filtrados = 0
    total_byma = 0
    total_ok   = 0

    for r in range(1, ws_xls.nrows):
        cell_a = ws_xls.cell(r, 0)
        if cell_a.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
            continue

        cvsa = _xls_int(cell_a)
        if cvsa is None:
            continue

        ticker_norm = _xls_str(ws_xls.cell(r, 9))   # col Norm.
        nombre      = _xls_str(ws_xls.cell(r, 1))    # Nombre_de_la_Especie
        tipo_col2   = _xls_str(ws_xls.cell(r, 2))    # Tipo
        tipo_activo = _xls_str(ws_xls.cell(r, 18))   # Tipo_de_Activo
        lista       = _xls_int(ws_xls.cell(r, 5)) or 0

        # Col 26 = haircut BYMA API
        try:
            raw26 = ws_xls.cell(r, 26)
            haircut = _xls_int(raw26) or 0
        except IndexError:
            haircut = 0

        ticker = ticker_norm or str(cvsa)

        if haircut > 0:
            total_byma += 1
            aforo_byma = 100 - haircut

            if lista == 0:
                tabla_inf = _tabla_para_tipo(tipo_col2, tipo_activo)
                lista_sug = _lista_para_aforo(tabla_inf, aforo_byma) if tabla_inf else "REVISAR"
                sin_lista.append({
                    "cvsa": cvsa, "ticker": ticker, "nombre": nombre,
                    "tipo_activo": tipo_activo, "tipo_col2": tipo_col2,
                    "haircut": haircut, "aforo_byma": aforo_byma,
                    "lista_sugerida": lista_sug,
                })
            else:
                tabla = _tabla_para_lista(lista)
                if tabla is None:
                    advertencias.append(
                        f"Lista {lista} fuera de rango — CVSA {cvsa} ({ticker}), "
                        f"haircut={haircut}%. Especie omitida del control."
                    )
                    continue

                aforo_sailing = _aforo_para_lista(tabla, lista)
                if aforo_sailing is None:
                    advertencias.append(
                        f"Lista {lista} no mapeada en tabla '{tabla}' — CVSA {cvsa}. Omitida."
                    )
                    continue

                if aforo_sailing == aforo_byma:
                    total_ok += 1
                else:
                    lista_sug = _lista_para_aforo(tabla, aforo_byma)
                    diferencias.append({
                        "cvsa": cvsa, "ticker": ticker, "nombre": nombre,
                        "tipo_activo": tipo_activo, "tabla": tabla,
                        "haircut": haircut, "aforo_byma": aforo_byma,
                        "lista_actual": lista, "aforo_sailing": aforo_sailing,
                        "diferencia_pp": aforo_byma - aforo_sailing,
                        "lista_sugerida": lista_sug,
                    })
                    if lista_sug == "REVISAR":
                        advertencias.append(
                            f"Lista sugerida REVISAR — CVSA {cvsa} ({ticker}), "
                            f"aforo BYMA={aforo_byma}%, tabla '{tabla}'."
                        )

        elif lista > 0:
            fvenc = _fecha_vencimiento(wb_xls, ws_xls, r)
            if fvenc is not None and fvenc < HOY:
                vencidos_filtrados += 1
                continue
            lista_sin_byma.append({
                "cvsa": cvsa, "ticker": ticker, "nombre": nombre,
                "tipo_activo": tipo_activo, "lista": lista,
            })

    # ─── Estilos ──────────────────────────────────────────────────────────────
    AZUL_FILL    = PatternFill("solid", fgColor="4472C4")
    ROJO_FILL    = PatternFill("solid", fgColor="C00000")
    NARANJA_FILL = PatternFill("solid", fgColor="ED7D31")
    VERDE_FILL   = PatternFill("solid", fgColor="00B050")
    ROJO_CELDA   = PatternFill("solid", fgColor="FF0000")
    BLANCO_FONT  = Font(name="Arial", bold=True, color="FFFFFF")
    NORMAL_FONT  = Font(name="Arial", size=11)
    BORDER_THIN  = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin"),
    )
    CENTER = Alignment(horizontal="center", vertical="center")

    def _aplicar_encabezado(ws, headers, fill=None):
        f = fill or AZUL_FILL
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.fill = f
            cell.font = BLANCO_FONT
            cell.alignment = CENTER
            cell.border = BORDER_THIN

    def _aplicar_fila(ws, fila, valores, fills=None, fonts=None):
        for col, v in enumerate(valores, 1):
            cell = ws.cell(row=fila, column=col, value=v)
            cell.font      = (fonts or {}).get(col, NORMAL_FONT)
            cell.alignment = CENTER
            cell.border    = BORDER_THIN
            if fills and fills.get(col):
                cell.fill = fills[col]

    def _autofit(ws, col_widths):
        for col, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = w

    # ─── Generar Excel ────────────────────────────────────────────────────────
    wb = openpyxl.Workbook()

    # Hoja 1: Diferencias Aforo
    ws_dif = wb.active
    ws_dif.title = "Diferencias Aforo"
    headers_dif = [
        "Ticker", "Nombre", "Tipo Activo", "Codigo CVSA",
        "Aforo BYMA %", "Lista Actual (Gallo)", "Aforo Lista Actual %",
        "Diferencia (pp)", "LISTA SUGERIDA",
    ]
    _aplicar_encabezado(ws_dif, headers_dif)
    ws_dif.cell(1, 9).fill = VERDE_FILL

    for i, d in enumerate(diferencias, 2):
        diff_str = (f"+{d['diferencia_pp']}" if d["diferencia_pp"] > 0
                    else str(d["diferencia_pp"])) + "pp"
        _aplicar_fila(ws_dif, i, [
            d["ticker"], d["nombre"], d["tipo_activo"], d["cvsa"],
            d["aforo_byma"], d["lista_actual"], d["aforo_sailing"],
            diff_str, d["lista_sugerida"],
        ], fills={7: ROJO_CELDA, 9: VERDE_FILL},
           fonts={9: Font(name="Arial", color="FFFFFF", bold=True)})

    _autofit(ws_dif, [12, 30, 25, 14, 15, 22, 22, 16, 16])

    # Hoja 2: Sin Lista en Gallo
    ws_sl = wb.create_sheet("Sin Lista Gallo")
    headers_sl = [
        "Ticker", "Nombre", "Tipo Activo", "Codigo CVSA",
        "Haircut BYMA %", "Aforo BYMA %", "LISTA SUGERIDA",
    ]
    _aplicar_encabezado(ws_sl, headers_sl, fill=ROJO_FILL)
    ws_sl.cell(1, 7).fill = VERDE_FILL

    for i, esp in enumerate(sin_lista, 2):
        _aplicar_fila(ws_sl, i, [
            esp["ticker"], esp["nombre"], esp["tipo_activo"], esp["cvsa"],
            esp["haircut"], esp["aforo_byma"], esp["lista_sugerida"],
        ], fills={7: VERDE_FILL},
           fonts={7: Font(name="Arial", color="FFFFFF", bold=True)})

    _autofit(ws_sl, [12, 30, 25, 14, 16, 15, 16])

    # Hoja 3: Lista Sin BYMA (informacional, vencidos excluidos)
    ws_lb = wb.create_sheet("Lista Sin BYMA")
    headers_lb = ["Ticker", "Nombre", "Tipo Activo", "Codigo CVSA", "Lista Gallo"]
    _aplicar_encabezado(ws_lb, headers_lb, fill=NARANJA_FILL)

    for i, esp in enumerate(lista_sin_byma, 2):
        _aplicar_fila(ws_lb, i, [
            esp["ticker"], esp["nombre"], esp["tipo_activo"], esp["cvsa"], esp["lista"],
        ])

    _autofit(ws_lb, [12, 30, 25, 14, 14])

    buf = io.BytesIO()
    wb.save(buf)
    xlsx_bytes = buf.getvalue()

    resumen = {
        "total_byma":    total_byma,
        "total_ok":      total_ok,
        "diferencias":   len(diferencias),
        "sin_lista":     len(sin_lista),
        "lista_sin_byma": len(lista_sin_byma),
        "vencidos_filtrados": vencidos_filtrados,
    }

    return xlsx_bytes, resumen, advertencias
