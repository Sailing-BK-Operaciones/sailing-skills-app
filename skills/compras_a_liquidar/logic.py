"""
Compras a Liquidar — versión web (Streamlit).

Identifica compras pendientes de liquidar para comitentes con saldo deudor
vencido > $99.900 (que van a recibir toma de caución).

Cruza: SALPESO.XLS (saldos), OPEVEN.XLS (ops a vencer), CONTBOLE.XLS (boletos),
       ESPECIES.XLS (ticker map).
"""

from io import BytesIO
from datetime import datetime, date

import xlrd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


UMBRAL_DEUDOR = 99_900.0

HEADER_FILL   = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT   = Font(bold=True, color="FFFFFF", size=10)
SUBTOTAL_FILL = PatternFill("solid", fgColor="D6E4F0")
SUBTOTAL_FONT = Font(bold=True, size=10)
ORANGE_FILL   = PatternFill("solid", fgColor="FFE0B2")
THIN_BORDER   = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)
FMT_MONEY = "#,##0.00"
FMT_NUM   = "#,##0"


def _read_xls(file_obj, sheet_name):
    """Lee un XLS desde un file-like; devuelve lista de dicts {header: value}."""
    file_obj.seek(0)
    wb = xlrd.open_workbook(file_contents=file_obj.read())
    sh = wb.sheet_by_name(sheet_name)
    headers = [str(sh.cell_value(0, c)).strip("'") for c in range(sh.ncols)]
    out = []
    for r in range(1, sh.nrows):
        out.append({headers[c]: sh.cell_value(r, c) for c in range(sh.ncols)})
    return out


def _clean_ctte(val):
    return str(val).strip().rstrip(".0").strip()


def _parse_date_str(s):
    s = str(s).strip()
    for fmt in ("%d/%m/%y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def _apply_header(ws, headers):
    ws.append(headers)
    for cell in ws[1]:
        cell.font      = HEADER_FONT
        cell.fill      = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = THIN_BORDER


def _style_data_row(ws, r, money_cols=(), num_cols=(), fill=None):
    for ci in range(1, ws.max_column + 1):
        cell = ws.cell(r, ci)
        cell.border    = THIN_BORDER
        cell.alignment = Alignment(
            horizontal="right" if ci in money_cols or ci in num_cols else "left",
            vertical="center",
        )
        if fill:
            cell.fill = fill
    for ci in money_cols:
        ws.cell(r, ci).number_format = FMT_MONEY
    for ci in num_cols:
        ws.cell(r, ci).number_format = FMT_NUM


def generar_reporte(salpeso_file, opeven_file, contbole_file, especies_file):
    """
    Parámetros
    ----------
    salpeso_file   : UploadedFile — SALPESO.XLS (saldos vencidos por comitente)
    opeven_file    : UploadedFile — OPEVEN.XLS (operaciones a vencer)
    contbole_file  : UploadedFile — CONTBOLE.XLS (boletos del día)
    especies_file  : UploadedFile — ESPECIES.XLS (maestro de especies)

    Devuelve (BytesIO con Excel, dict resumen).
    """
    # ── 1. ESPECIES.XLS → mapa código → ticker ───────────────────────────────
    ticker_map = {}   # codigo_5dig -> ticker (Norm. col 9)
    nombre_map = {}   # codigo_5dig -> nombre completo
    esp_rows = _read_xls(especies_file, "Datos_Fijos_Especies")
    for row in esp_rows:
        cod    = str(row.get("Codigo", "")).strip("'").strip().zfill(5)
        ticker = str(row.get("Norm.", "")).strip("'").strip()
        nombre = str(row.get("Nombre_de_la_Especie", "")).strip("'").strip()
        if cod:
            ticker_map[cod] = ticker
            nombre_map[cod] = nombre

    # ── 2. SALPESO → deudores > umbral ───────────────────────────────────────
    salpeso_rows = _read_xls(salpeso_file, "Listado_de_Saldos")
    deudores = {}   # ctte -> {"nombre": str, "saldo": float}
    for row in salpeso_rows:
        ctte = _clean_ctte(row.get("Numero", ""))
        if not ctte:
            continue
        try:
            saldo = float(row.get("Saldo Vencido", 0) or 0)
        except (ValueError, TypeError):
            saldo = 0.0
        if saldo > UMBRAL_DEUDOR:
            nombre = str(row.get("Nombre Comitente", "")).strip()
            deudores[ctte] = {"nombre": nombre, "saldo": saldo}

    # ── 3. OPEVEN → compras (CPRA con Fec.Liq.) de deudores ─────────────────
    opeven_rows = _read_xls(opeven_file, "Operaciones_Vencer")
    compras_opeven = []
    for row in opeven_rows:
        ctte   = _clean_ctte(row.get("Numero", ""))
        concep = str(row.get("Concepto", "")).strip()
        fecliq = str(row.get("Fec.Liq.", "")).strip()
        if not fecliq or concep != "CPRA":
            continue
        if ctte not in deudores:
            continue
        try:
            importe  = float(row.get("Importe Neto", 0) or 0)
            cantidad = float(row.get("Cantidad", 0) or 0)
        except (ValueError, TypeError):
            importe = cantidad = 0.0
        codigo = str(row.get("Codigo", "")).strip().zfill(5)
        compras_opeven.append({
            "ctte":     ctte,
            "nombre":   deudores[ctte]["nombre"],
            "origen":   "OPEVEN",
            "codigo":   codigo,
            "ticker":   ticker_map.get(codigo, ""),
            "cantidad": cantidad,
            "importe":  importe,
            "fec_liq":  fecliq,
            "especie":  str(row.get("Nombre de la Especie", "")).strip(),
            "concepto": concep,
            "fec_ope":  str(row.get("Fec.Ope.", "")).strip(),
            "boleto":   str(row.get("Boleto", "")).strip(),
        })

    # ── 4. CONTBOLE → compras CI de deudores ─────────────────────────────────
    compras_ci = []
    cb_rows = _read_xls(contbole_file, "Control_de_Boletos")
    for row in cb_rows:
        operacion = str(row.get("peracion", "")).strip()
        if operacion != "CPRA":
            continue
        ctte = _clean_ctte(row.get("Comitente", ""))
        if ctte not in deudores:
            continue
        fec_ope = str(row.get("Fec_Ope", "")).strip()
        fec_liq = str(row.get("Fec_Liq", "")).strip()
        d_ope = _parse_date_str(fec_ope)
        d_liq = _parse_date_str(fec_liq)
        if d_ope is None or d_liq is None or d_ope != d_liq:
            continue
        try:
            importe  = float(row.get("Total_Neto", 0) or 0)
            cantidad = float(row.get("Valor_Nominal", 0) or 0)
        except (ValueError, TypeError):
            importe = cantidad = 0.0
        especie_cb = str(row.get("Especie", "")).strip()
        cod_cb = next((k for k, v in ticker_map.items() if v == especie_cb), "")
        compras_ci.append({
            "ctte":     ctte,
            "nombre":   deudores[ctte]["nombre"],
            "origen":   "CONTBOLE CI",
            "codigo":   cod_cb,
            "ticker":   especie_cb,
            "cantidad": cantidad,
            "importe":  importe,
            "fec_liq":  fec_liq,
            "especie":  nombre_map.get(cod_cb, especie_cb),
            "concepto": "CPRA CI",
            "fec_ope":  fec_ope,
            "boleto":   str(row.get("Boleto", "")).strip(),
        })

    # ── 5. Resumen por comitente ─────────────────────────────────────────────
    resumen = {}
    for c, d in deudores.items():
        resumen[c] = {"nombre": d["nombre"], "saldo_deudor": d["saldo"],
                      "total_opeven": 0.0, "total_ci": 0.0}
    for op in compras_opeven:
        resumen[op["ctte"]]["total_opeven"] += op["importe"]
    for op in compras_ci:
        resumen[op["ctte"]]["total_ci"] += op["importe"]

    all_compras = compras_opeven + compras_ci

    # ── 6. Excel ─────────────────────────────────────────────────────────────
    wb = Workbook()

    # Hoja Resumen
    ws_res = wb.active
    ws_res.title = "Resumen"
    _apply_header(ws_res, ["Comitente", "Nombre", "Saldo Deudor",
                            "Compras OPEVEN", "Compras CI", "Total Compras"])
    for ctte in sorted(resumen.keys()):
        d     = resumen[ctte]
        total = d["total_opeven"] + d["total_ci"]
        ws_res.append([ctte, d["nombre"], d["saldo_deudor"],
                       d["total_opeven"], d["total_ci"], total])
        r    = ws_res.max_row
        fill = ORANGE_FILL if d["total_ci"] > 0 else None
        _style_data_row(ws_res, r, money_cols=(3, 4, 5, 6), fill=fill)

    tot_saldo  = sum(d["saldo_deudor"] for d in resumen.values())
    tot_opeven = sum(d["total_opeven"] for d in resumen.values())
    tot_ci     = sum(d["total_ci"]     for d in resumen.values())
    ws_res.append(["", "TOTAL", tot_saldo, tot_opeven, tot_ci, tot_opeven + tot_ci])
    r = ws_res.max_row
    for ci in range(1, 7):
        cell = ws_res.cell(r, ci)
        cell.font   = SUBTOTAL_FONT
        cell.fill   = SUBTOTAL_FILL
        cell.border = THIN_BORDER
        cell.alignment = Alignment(
            horizontal="right" if ci >= 3 else "left", vertical="center"
        )
        if ci >= 3:
            cell.number_format = FMT_MONEY

    ws_res.column_dimensions["A"].width = 12
    ws_res.column_dimensions["B"].width = 30
    for col in ["C", "D", "E", "F"]:
        ws_res.column_dimensions[col].width = 20
    ws_res.freeze_panes = "A2"
    if ws_res.max_row > 2:
        ws_res.auto_filter.ref = f"A1:F{ws_res.max_row - 1}"

    # Hoja Detalle
    ws_det = wb.create_sheet("Detalle")
    _apply_header(ws_det, [
        "Comitente", "Nombre", "Saldo Deudor", "Origen",
        "Especie", "ticker", "Cantidad", "Importe",
        "Fec. Liq.", "Nombre Especie", "Concepto", "Fec. Ope.", "Boleto",
    ])
    for op in sorted(all_compras, key=lambda x: (x["ctte"], x["origen"], x["fec_liq"])):
        saldo = resumen[op["ctte"]]["saldo_deudor"]
        ws_det.append([
            op["ctte"],    op["nombre"],  saldo,         op["origen"],
            op["codigo"],  op["ticker"],  op["cantidad"], op["importe"],
            op["fec_liq"], op["especie"], op["concepto"], op["fec_ope"], op["boleto"],
        ])
        r    = ws_det.max_row
        fill = ORANGE_FILL if op["origen"] == "CONTBOLE CI" else None
        _style_data_row(ws_det, r, money_cols=(3, 8), num_cols=(7,), fill=fill)

    col_widths = {
        "A": 12, "B": 30, "C": 18, "D": 14, "E": 10, "F": 12,
        "G": 14, "H": 18, "I": 12, "J": 35, "K": 12, "L": 12, "M": 10,
    }
    for col, w in col_widths.items():
        ws_det.column_dimensions[col].width = w
    ws_det.freeze_panes = "A2"
    if ws_det.max_row >= 1:
        ws_det.auto_filter.ref = f"A1:M{max(ws_det.max_row, 1)}"

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    cttes_con_ci = sorted({op["ctte"] for op in compras_ci})
    resumen_ui = {
        "fecha":            date.today().strftime("%d-%m-%Y"),
        "n_deudores":       len(deudores),
        "n_ops_opeven":     len(compras_opeven),
        "n_ops_ci":         len(compras_ci),
        "total_opeven":     tot_opeven,
        "total_ci":         tot_ci,
        "total_compras":    tot_opeven + tot_ci,
        "total_deudor":     tot_saldo,
        "cttes_con_ci":     cttes_con_ci,
        "detalle_deudores": sorted(
            [{"ctte": c, **d, "total": d["total_opeven"] + d["total_ci"]}
             for c, d in resumen.items()],
            key=lambda x: x["ctte"],
        ),
    }
    return output, resumen_ui
