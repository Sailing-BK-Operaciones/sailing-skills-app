"""
Control diario Op-SDIB — versión web (Streamlit).
Concilia operaciones PPT BYMA + SENEBI vs CONTBOLE (Gallo) del día.
Recibe UploadedFiles; devuelve (BytesIO con el Excel, dict con resumen).
"""

from io import BytesIO
from datetime import datetime
from collections import defaultdict, Counter
import xlrd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

TOL      = 0.02
C_GREEN  = "C6EFCE"
C_RED    = "FFC7CE"
C_YELLOW = "FFEB9C"
C_BLUE   = "DDEEFF"
C_HEADER = "4472C4"
C_ORANGE = "FFE0B2"

DNC_EXCEPTIONS = {
    "DNC5D": ("DNC5O", "DOLAR MEP"),
    "DNC7D": ("DNC7O", "DOLAR MEP"),
    "DNCAD": ("DNCAO", "DOLAR MEP"),
    "DNCAC": ("DNCAO", "USD"),
}

GALLO_CAU  = {"CCCD", "CCTE", "TCCD", "TOCT", "CTCD", "CTTD", "CCCP", "CCFP"}
GALLO_TRAS = {"VTTR", "COTR"}
GALLO_SNB  = {"CRCN", "VRCN"}
GALLO_MAV  = {"VCHM", "CCHM"}

GALLO_SEN = {
    "CPRA": "C", "CPU$": "C", "CPUC": "C", "CSU$": "C",
    "CRCN": "C", "COPR": "C", "CTCD": "C", "CTTD": "C",
    "CCHM": "C",
    "VTAS": "V", "VTU$": "V", "VTUC": "V", "VSU$": "V",
    "VRCN": "V", "VCHM": "V", "VTPR": "V",
    "COTR": "C", "VTTR": "V",
    "CCCD": "V", "CCTE": "V",
    "TCCD": "C", "TOCT": "C",
    "CCCP": "V", "CCFP": "V",
}

GALLO_MON = {
    "CPRA": "Pesos",     "VTAS": "Pesos",
    "CRCN": "Pesos",     "VRCN": "Pesos",
    "VCHM": "Pesos",     "CCHM": "Pesos",
    "COPR": "Pesos",     "VTPR": "Pesos",
    "CPU$": "DOLAR MEP", "VTU$": "DOLAR MEP",
    "CSU$": "DOLAR MEP", "VSU$": "DOLAR MEP",
    "CTCD": "DOLAR MEP", "CTTD": "DOLAR MEP",
    "CPUC": "USD",       "VTUC": "USD",
    "VTTR": "Pesos",     "COTR": "Pesos",
    "CCCD": "Pesos",     "CCTE": "Pesos",
    "TCCD": "Pesos",     "TOCT": "Pesos",
    "CCCP": "DOLAR MEP", "CCFP": "DOLAR MEP",
}

GALLO_CAU_CAP = {"CCCD", "TCCD", "CTCD", "CCCP"}
CTTES_CARTERA = {"1000", "1001", "1002", "1003"}


# ── Helpers de parseo ──────────────────────────────────────────────────────────

def _map_byma_esp(esp_raw, gallo_esp):
    esp = esp_raw.strip()
    if esp == "PESOS":
        return ("PESOS", "Pesos")
    if esp in DNC_EXCEPTIONS:
        return DNC_EXCEPTIONS[esp]
    if len(esp) > 1 and esp.endswith("C") and esp[:-1] in gallo_esp:
        return (esp[:-1], "USD")
    if len(esp) > 1 and esp.endswith("D") and esp[:-1] in gallo_esp:
        return (esp[:-1], "DOLAR MEP")
    if len(esp) > 1 and esp.endswith("D") and (esp[:-1] + "O") in gallo_esp:
        return (esp[:-1] + "O", "DOLAR MEP")
    if len(esp) > 1 and esp.endswith("C") and (esp[:-1] + "O") in gallo_esp:
        return (esp[:-1] + "O", "USD")
    return (esp, "Pesos")


def _merge_dicts(a, b):
    result = defaultdict(lambda: [0., 0.])
    for k, v in a.items():
        result[k][0] += v[0]
        result[k][1] += v[1]
    for k, v in b.items():
        result[k][0] += v[0]
        result[k][1] += v[1]
    return dict(result)


def _parse_byma(lines, gallo_esp):
    """Parsea OPERSECEXT_GARA.DAT (segmento garantizado PPT BYMA)."""
    mer = defaultdict(lambda: [0., 0.])
    cau = defaultdict(lambda: [0., 0.])
    for line in lines:
        p = line.split('"')
        if len(p) < 28:
            continue
        sen  = p[2].strip()
        tipo = p[6].strip()
        ctte = p[17].strip()
        try:
            vn  = float(p[12])
            imp = float(p[19])
        except Exception:
            continue
        if sen not in ("V", "C"):
            continue
        if tipo == "U" or p[7].strip() == "PESOS":
            mon_cau = "DOLAR MEP" if p[7].strip() == "DOLAR" else "Pesos"
            cau[(ctte, sen, mon_cau)][0] += vn
            cau[(ctte, sen, mon_cau)][1] += imp
        else:
            esp, mon = _map_byma_esp(p[7], gallo_esp)
            mer[(ctte, esp, mon, sen)][0] += vn
            mer[(ctte, esp, mon, sen)][1] += imp
    return dict(mer), dict(cau)


def _parse_senebi(lines, gallo_esp):
    """Parsea OPERBILEXT_GARA.DAT (segmento bilateral SENEBI).
    Registra lado cliente e infiere posición de contraparte (sentido invertido).
    """
    snb = defaultdict(lambda: [0., 0.])
    for line in lines:
        p = line.split('"')
        if len(p) < 30:
            continue
        sen_cli  = p[4].strip()
        ctte_cli = p[7].strip()
        esp_raw  = p[9].strip()
        cparty   = p[29].strip()
        try:
            vn  = float(p[13])
            imp = float(p[15])
        except Exception:
            continue
        if sen_cli not in ("V", "C"):
            continue
        esp, mon = _map_byma_esp(esp_raw, gallo_esp)
        snb[(ctte_cli, esp, mon, sen_cli)][0] += vn
        snb[(ctte_cli, esp, mon, sen_cli)][1] += imp
        sen_cp = "V" if sen_cli == "C" else "C"
        snb[(cparty, esp, mon, sen_cp)][0] += vn
        snb[(cparty, esp, mon, sen_cp)][1] += imp
    return dict(snb)


def _parse_gallo(file_bytes):
    """Parsea CONTBOLE.XLS desde bytes."""
    wb = xlrd.open_workbook(file_contents=file_bytes)
    ws = wb.sheet_by_name("Control_de_Boletos")
    esp_set = set()
    mer = defaultdict(lambda: [0., 0.])
    cau = defaultdict(lambda: {"capital": 0., "total": 0.})
    tra = defaultdict(lambda: [0., 0.])
    snb = defaultdict(lambda: [0., 0.])
    mav = []
    arancel_alerts = []
    for i in range(1, ws.nrows):
        op  = str(ws.cell_value(i, 1)).strip()
        raw = ws.cell_value(i, 4)
        if not op or not raw:
            continue
        try:
            ctte = str(int(float(raw)))
        except Exception:
            continue
        esp     = str(ws.cell_value(i, 6)).strip()
        imp     = float(ws.cell_value(i, 7)) if ws.cell_value(i, 7) else 0.
        vn      = float(ws.cell_value(i, 8)) if ws.cell_value(i, 8) else 0.
        arancel = float(ws.cell_value(i, 9)) if ws.cell_value(i, 9) else 0.
        esp_set.add(esp)
        sen = GALLO_SEN.get(op)
        mon = GALLO_MON.get(op)
        if ctte in CTTES_CARTERA and abs(arancel) > 0.001:
            boleto = str(ws.cell_value(i, 0)).strip()
            arancel_alerts.append({
                "boleto": boleto, "ctte": ctte, "op": op,
                "esp": esp, "arancel": arancel,
            })
        if op in GALLO_CAU:
            mon_cau = GALLO_MON.get(op, "Pesos")
            if op in GALLO_CAU_CAP:
                cau[(ctte, sen, mon_cau)]["capital"] += imp
            else:
                cau[(ctte, sen, mon_cau)]["total"] += imp
        elif op in GALLO_TRAS:
            if sen and mon:
                tra[(ctte, esp, mon, sen)][0] += vn
                tra[(ctte, esp, mon, sen)][1] += imp
        elif op in GALLO_MAV:
            mav.append({"ctte": ctte, "op": op, "esp": esp,
                        "sen": sen or "V", "mon": mon or "Pesos",
                        "vn": vn, "imp": imp})
        elif op in GALLO_SNB:
            if sen and mon:
                snb[(ctte, esp, mon, sen)][0] += vn
                snb[(ctte, esp, mon, sen)][1] += imp
        elif sen and mon:
            mer[(ctte, esp, mon, sen)][0] += vn
            mer[(ctte, esp, mon, sen)][1] += imp
    return esp_set, dict(mer), dict(cau), dict(tra), dict(snb), mav, arancel_alerts


# ── Comparadores ──────────────────────────────────────────────────────────────

def _get_st(in_b, in_g, difs):
    if in_b and in_g:
        return "OK" if all(abs(d) <= TOL for d in difs) else "DIFERENCIA"
    return "SOLO BYMA" if in_b else "SOLO GALLO"


def _compare_ops(b, g, ppt_keys_b, ppt_keys_g, snb_keys_b, snb_keys_g):
    rows = []
    for k in sorted(set(b) | set(g)):
        ctte, esp, mon, sen = k
        bv = b.get(k, [0., 0.])
        gv = g.get(k, [0., 0.])
        dv = bv[0] - gv[0]
        di = bv[1] - gv[1]
        in_ppt = k in ppt_keys_b or k in ppt_keys_g
        in_snb = k in snb_keys_b or k in snb_keys_g
        seg = ("PPT+SENEBI" if in_ppt and in_snb
               else ("PPT" if in_ppt else "SENEBI"))
        estado = _get_st(k in b, k in g, [dv, di])
        verificar = ""
        if ctte == "1002":
            verificar = "CTA 1002"
        elif estado == "SOLO BYMA":
            verificar = "FALTA BOLETO"
        elif estado not in ("OK", "TRADING INTRADAY"):
            verificar = "VERIFICAR"
        rows.append({
            "ctte": ctte, "especie": esp, "moneda": mon, "sentido": sen,
            "segmento": seg,
            "vn_b": bv[0], "vn_g": gv[0], "dif_vn": dv,
            "imp_b": bv[1], "imp_g": gv[1], "dif_imp": di,
            "estado": estado, "verificar": verificar,
        })
    return rows


def _second_pass_ticker(rows):
    """Re-empareja SOLO BYMA vs SOLO GALLO donde el ticker difiere solo en el
    último carácter (convención de moneda O/D/C). Mismo ctte, sentido, VN e IMP."""
    solo_b = [(i, r) for i, r in enumerate(rows) if r["estado"] == "SOLO BYMA"]
    solo_g = [(i, r) for i, r in enumerate(rows) if r["estado"] == "SOLO GALLO"]
    absorbed_g = set()
    for ib, rb in solo_b:
        esp_b = rb["especie"]
        if len(esp_b) <= 1:
            continue
        for ig, rg in solo_g:
            if ig in absorbed_g:
                continue
            esp_g = rg["especie"]
            if len(esp_g) <= 1:
                continue
            if (rb["ctte"]    == rg["ctte"] and
                rb["sentido"] == rg["sentido"] and
                esp_b[:-1]    == esp_g[:-1] and
                esp_b         != esp_g and
                abs(rb["vn_b"]  - rg["vn_g"])  <= TOL and
                abs(rb["imp_b"] - rg["imp_g"]) <= TOL):
                rows[ib]["vn_g"]    = rg["vn_g"]
                rows[ib]["imp_g"]   = rg["imp_g"]
                rows[ib]["dif_vn"]  = rb["vn_b"] - rg["vn_g"]
                rows[ib]["dif_imp"] = rb["imp_b"] - rg["imp_g"]
                rows[ib]["especie"] = rg["especie"]
                rows[ib]["moneda"]  = rg["moneda"]
                rows[ib]["estado"]  = "OK"
                rows[ib]["verificar"] = ""
                absorbed_g.add(ig)
                break
    return [r for i, r in enumerate(rows) if i not in absorbed_g]


def _compare_cau(b, g):
    rows = []
    for k in sorted(set(b) | set(g)):
        ctte, sen, mon = k
        bv = b.get(k, [0., 0.])
        gv = g.get(k, {"capital": 0., "total": 0.})
        dc = bv[0] - gv["capital"]
        dt = bv[1] - gv["total"]
        rows.append({
            "ctte": ctte, "sentido": sen, "moneda": mon,
            "vn_b": bv[0], "cap_g": gv["capital"], "dif_cap": dc,
            "imp_b": bv[1], "tot_g": gv["total"],  "dif_tot": dt,
            "estado": _get_st(k in b, k in g, [dc, dt]),
        })
    return rows


def _compare_tra(g):
    rows = []
    for k in sorted(g):
        ctte, esp, mon, sen = k
        gv = g[k]
        rows.append({
            "ctte": ctte, "especie": esp, "moneda": mon, "sentido": sen,
            "vn_b": 0., "vn_g": gv[0], "dif_vn": -gv[0],
            "imp_b": 0., "imp_g": gv[1], "dif_imp": -gv[1],
            "estado": "TRADING INTRADAY",
        })
    return rows


def _analyze_1002(ops_rows):
    rows_1002 = {(r["especie"], r["moneda"], r["sentido"]): r
                 for r in ops_rows if r["ctte"] == "1002"}
    results = []
    for k, r1002 in rows_1002.items():
        esp, mon, sen = k
        dif_1002 = r1002["dif_vn"]
        otros = [(r["ctte"], r["dif_vn"])
                 for r in ops_rows
                 if r["ctte"] != "1002"
                 and r["especie"] == esp
                 and r["moneda"] == mon
                 and r["sentido"] == sen
                 and abs(r["dif_vn"]) > TOL]
        dif_otros = sum(v for _, v in otros)
        neto = dif_1002 + dif_otros
        results.append({
            "especie":     esp,
            "moneda":      mon,
            "sentido":     sen,
            "estado_1002": r1002["estado"],
            "vn_b_1002":   r1002["vn_b"],
            "vn_g_1002":   r1002["vn_g"],
            "dif_1002":    dif_1002,
            "otros_cttes": ", ".join(c for c, _ in otros) or "-",
            "dif_otros":   dif_otros,
            "neto":        neto,
            "balance":     "OK" if abs(neto) <= TOL else "REVISAR",
        })
    return results


def _apply_1002_compensation(ops_rows, analisis_1002):
    for a in analisis_1002:
        if a["balance"] != "OK":
            continue
        esp, mon, sen = a["especie"], a["moneda"], a["sentido"]
        otros_set = {c.strip() for c in a["otros_cttes"].split(",")
                     if c.strip() and c.strip() != "-"}
        for r in ops_rows:
            if r["especie"] != esp or r["moneda"] != mon or r["sentido"] != sen:
                continue
            if r["estado"] == "OK":
                continue
            if r["ctte"] == "1002" or r["ctte"] in otros_set:
                r["estado"]    = "OK"
                r["verificar"] = "CTA 1002"
    return ops_rows


# ── Helpers Excel ──────────────────────────────────────────────────────────────

def _fl(c):  return PatternFill("solid", fgColor=c)
def _hf():   return Font(bold=True, color="FFFFFF")
def _bf():   return Font(bold=True)


def _wh(ws, hdrs, row=1):
    for c, h in enumerate(hdrs, 1):
        cell = ws.cell(row, c, h)
        cell.font = _hf()
        cell.fill = _fl(C_HEADER)
        cell.alignment = Alignment(horizontal="center")


def _nf(ws, r, c, v):
    cell = ws.cell(row=r, column=c, value=v)
    cell.number_format = "#,##0.00"
    return cell


def _rf(ws, r, n, color):
    for c in range(1, n + 1):
        ws.cell(r, c).fill = _fl(color)


def _af(ws):
    for col in ws.columns:
        w = max((len(str(c.value or "")) for c in col), default=8)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(w + 2, 42)


def _st_color(estado):
    return {"OK": C_GREEN, "DIFERENCIA": C_RED,
            "SOLO BYMA": C_YELLOW, "SOLO GALLO": C_BLUE,
            "TRADING INTRADAY": C_ORANGE}.get(estado, "FFFFFF")


# ── Escritura de hojas ─────────────────────────────────────────────────────────

def _write_ops_sheet(ws, rows):
    _wh(ws, ["CTTE", "ESPECIE", "MONEDA", "SENTIDO", "SEGMENTO",
             "VN BYMA", "VN GALLO", "DIF VN",
             "IMP BYMA", "IMP GALLO", "DIF IMP", "ESTADO", "VERIFICAR"])
    for i, r in enumerate(rows, 2):
        co = _st_color(r["estado"])
        ws.cell(i,  1, r["ctte"])
        ws.cell(i,  2, r["especie"])
        ws.cell(i,  3, r["moneda"])
        ws.cell(i,  4, r["sentido"])
        ws.cell(i,  5, r["segmento"])
        _nf(ws, i,  6, r["vn_b"])
        _nf(ws, i,  7, r["vn_g"])
        _nf(ws, i,  8, r["dif_vn"])
        _nf(ws, i,  9, r["imp_b"])
        _nf(ws, i, 10, r["imp_g"])
        _nf(ws, i, 11, r["dif_imp"])
        ws.cell(i, 12, r["estado"])
        ws.cell(i, 13, r["verificar"])
        _rf(ws, i, 13, co)
    ws.auto_filter.ref = f"A1:M{len(rows)+1}"
    _af(ws)


def _write_cau_sheet(ws, rows):
    _wh(ws, ["CTTE", "SENTIDO", "MONEDA",
             "VN BYMA (capital)", "CAPITAL GALLO", "DIF CAPITAL",
             "IMP BYMA (total)", "TOTAL GALLO", "DIF TOTAL", "ESTADO"])
    for i, r in enumerate(rows, 2):
        co = _st_color(r["estado"])
        ws.cell(i, 1, r["ctte"])
        ws.cell(i, 2, r["sentido"])
        ws.cell(i, 3, r["moneda"])
        _nf(ws, i, 4, r["vn_b"])
        _nf(ws, i, 5, r["cap_g"])
        _nf(ws, i, 6, r["dif_cap"])
        _nf(ws, i, 7, r["imp_b"])
        _nf(ws, i, 8, r["tot_g"])
        _nf(ws, i, 9, r["dif_tot"])
        ws.cell(i, 10, r["estado"])
        _rf(ws, i, 10, co)
    ws.auto_filter.ref = f"A1:J{len(rows)+1}"
    _af(ws)


def _write_tra_sheet(ws, rows):
    _wh(ws, ["CTTE", "ESPECIE", "MONEDA", "SENTIDO",
             "VN BYMA", "VN GALLO", "DIF VN",
             "IMP BYMA", "IMP GALLO", "DIF IMP", "ESTADO"])
    for i, r in enumerate(rows, 2):
        co = _st_color(r["estado"])
        ws.cell(i, 1, r["ctte"])
        ws.cell(i, 2, r["especie"])
        ws.cell(i, 3, r["moneda"])
        ws.cell(i, 4, r["sentido"])
        _nf(ws, i, 5, r["vn_b"])
        _nf(ws, i, 6, r["vn_g"])
        _nf(ws, i, 7, r["dif_vn"])
        _nf(ws, i, 8, r["imp_b"])
        _nf(ws, i, 9, r["imp_g"])
        _nf(ws, i, 10, r["dif_imp"])
        ws.cell(i, 11, r["estado"])
        _rf(ws, i, 11, co)
    ws.auto_filter.ref = f"A1:K{len(rows)+1}"
    _af(ws)


def _write_mav_sheet(ws, mav_rows):
    _wh(ws, ["CTTE", "ESPECIE", "SENTIDO", "MONEDA",
             "VN GALLO", "IMP GALLO", "OP GALLO"])
    for i, r in enumerate(mav_rows, 2):
        ws.cell(i, 1, r["ctte"])
        ws.cell(i, 2, r["esp"])
        ws.cell(i, 3, r["sen"])
        ws.cell(i, 4, r["mon"])
        _nf(ws, i, 5, r["vn"])
        _nf(ws, i, 6, r["imp"])
        ws.cell(i, 7, r["op"])
        _rf(ws, i, 7, C_BLUE)
    ws.auto_filter.ref = f"A1:G{len(mav_rows)+1}"
    _af(ws)


def _write_res(ws, ops_rows, cr, tr, mav_rows, analisis_1002, arancel_alerts=None):
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 12
    for col in "CDEFGHIJK":
        ws.column_dimensions[col].width = 16

    row = 1
    ws.cell(row, 1, "RESUMEN CONCILIACION").font = _bf()
    row += 1
    ws.cell(row, 1, f"Fecha: {datetime.today().strftime('%d/%m/%Y')}")
    row += 2

    est_std = [("OK", C_GREEN), ("DIFERENCIA", C_RED),
               ("SOLO BYMA", C_YELLOW), ("SOLO GALLO", C_BLUE),
               ("TRADING INTRADAY", C_ORANGE)]

    def bloque(titulo, rows, estados):
        nonlocal row
        ws.cell(row, 1, titulo).fill = _fl(C_HEADER)
        ws.cell(row, 1).font = _hf()
        ws.cell(row, 2, "Cant").fill = _fl(C_HEADER)
        ws.cell(row, 2).font = _hf()
        row += 1
        cnt = Counter(x["estado"] for x in rows)
        shown = False
        for st, co in estados:
            n = cnt.get(st, 0)
            if n == 0:
                continue
            ws.cell(row, 1, st).fill = _fl(co)
            ws.cell(row, 2, n).fill = _fl(co)
            row += 1
            shown = True
        if not shown:
            ws.cell(row, 1, "(sin datos)")
            row += 1
        row += 1

    bloque("OPERACIONES (PPT BYMA + SENEBI)", ops_rows, est_std)
    bloque("CAUCIONES",                       cr,       est_std)
    bloque("TRADING INTRADAY",                tr,       est_std)

    n_falta  = sum(1 for r in ops_rows if r["verificar"] == "FALTA BOLETO")
    co_falta = C_RED if n_falta > 0 else C_GREEN
    label_falta = (f"FALTA BOLETO EN GALLO — {n_falta} op(s)" if n_falta > 0
                   else "SIN FALTANTES DE BOLETO EN GALLO")
    ws.cell(row, 1, label_falta).fill = _fl(co_falta)
    ws.cell(row, 1).font = Font(bold=True, color="FFFFFF" if n_falta > 0 else "000000")
    ws.cell(row, 2, n_falta).fill = _fl(co_falta)
    ws.cell(row, 2).font = Font(bold=True)
    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width,
                                          len(label_falta) + 2)
    row += 2

    ws.cell(row, 1, "MAV").fill = _fl(C_HEADER)
    ws.cell(row, 1).font = _hf()
    ws.cell(row, 2, "Cant").fill = _fl(C_HEADER)
    ws.cell(row, 2).font = _hf()
    row += 1
    ws.cell(row, 1, "Ops MAV Gallo (sin contr. BYMA)").fill = _fl(C_BLUE)
    ws.cell(row, 2, len(mav_rows)).fill = _fl(C_BLUE)
    row += 2

    n_ar   = len(arancel_alerts) if arancel_alerts else 0
    co_ar  = C_RED if n_ar > 0 else C_GREEN
    label_ar = (f"*** ARANCEL EN CUENTA CARTERA PROPIA — {n_ar} boleto(s) ***"
                if n_ar > 0 else "SIN ARANCEL EN CUENTAS CARTERA PROPIA (1000-1003)")
    ws.cell(row, 1, label_ar).fill = _fl(co_ar)
    ws.cell(row, 1).font = Font(bold=True, color="FFFFFF" if n_ar > 0 else "000000")
    ws.cell(row, 2, n_ar).fill = _fl(co_ar)
    ws.cell(row, 2).font = Font(bold=True)
    ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width,
                                          len(label_ar) + 2)
    row += 1
    if arancel_alerts:
        for c, h in enumerate(["Boleto", "CTTE", "Operacion", "Especie", "Arancel"], 1):
            cell = ws.cell(row, c, h)
            cell.font = Font(bold=True)
            cell.fill = _fl("D9D9D9")
        row += 1
        for a in arancel_alerts:
            ws.cell(row, 1, a["boleto"])
            ws.cell(row, 2, a["ctte"])
            ws.cell(row, 3, a["op"])
            ws.cell(row, 4, a["esp"])
            cell_ar = ws.cell(row, 5, a["arancel"])
            cell_ar.number_format = "#,##0.00"
            cell_ar.fill = _fl(C_RED)
            cell_ar.font = Font(bold=True, color="FFFFFF")
            row += 1
    row += 2

    if analisis_1002:
        ws.cell(row, 1, "ANALISIS CTA 1002").font = _hf()
        ws.cell(row, 1).fill = _fl(C_HEADER)
        row += 1
        hdrs_1002 = ["ESPECIE", "MONEDA", "SENTIDO", "ESTADO 1002",
                     "VN BYMA 1002", "VN GALLO 1002", "DIF VN 1002",
                     "OTROS CTTES", "DIF VN OTROS", "NETO VN", "BALANCE"]
        for c, h in enumerate(hdrs_1002, 1):
            cell = ws.cell(row, c, h)
            cell.font = _bf()
            cell.fill = _fl("D9D9D9")
        row += 1
        for a in sorted(analisis_1002, key=lambda x: x["especie"]):
            co = C_GREEN if a["balance"] == "OK" else C_RED
            ws.cell(row,  1, a["especie"])
            ws.cell(row,  2, a["moneda"])
            ws.cell(row,  3, a["sentido"])
            ws.cell(row,  4, a["estado_1002"])
            _nf(ws, row,  5, a["vn_b_1002"])
            _nf(ws, row,  6, a["vn_g_1002"])
            _nf(ws, row,  7, a["dif_1002"])
            ws.cell(row,  8, a["otros_cttes"])
            _nf(ws, row,  9, a["dif_otros"])
            _nf(ws, row, 10, a["neto"])
            ws.cell(row, 11, a["balance"]).fill = _fl(co)
            row += 1


# ── Punto de entrada ──────────────────────────────────────────────────────────

def generar_reporte(dat_file, xls_file, bil_file=None):
    """
    dat_file : UploadedFile — OPERSECEXT_GARA.DAT
    xls_file : UploadedFile — CONTBOLE.XLS
    bil_file : UploadedFile or None — OPERBILEXT_GARA.DAT (opcional)

    Devuelve (BytesIO, resumen_dict).
    """
    dat_lines = dat_file.read().decode("latin-1").splitlines()
    xls_bytes = xls_file.read()
    bil_lines = bil_file.read().decode("latin-1").splitlines() if bil_file else []

    esp_g, mer_g, cau_g, tra_g, snb_g, mav_g, arancel_alerts = _parse_gallo(xls_bytes)
    mer_b, cau_b = _parse_byma(dat_lines, esp_g)
    snb_b = _parse_senebi(bil_lines, esp_g) if bil_lines else {}

    ppt_keys_b = set(mer_b.keys())
    ppt_keys_g = set(mer_g.keys())
    snb_keys_b = set(snb_b.keys())
    snb_keys_g = set(snb_g.keys())

    combined_b = _merge_dicts(mer_b, snb_b)
    combined_g = _merge_dicts(mer_g, snb_g)

    ops_rows = _compare_ops(combined_b, combined_g,
                            ppt_keys_b, ppt_keys_g, snb_keys_b, snb_keys_g)
    ops_rows = _second_pass_ticker(ops_rows)

    for r in ops_rows:
        if r["estado"] == "SOLO BYMA":
            k = (r["ctte"], r["especie"], r["moneda"], r["sentido"])
            if k in tra_g:
                r["estado"]    = "TRADING INTRADAY"
                r["verificar"] = ""

    cr    = _compare_cau(cau_b, cau_g)
    tr    = _compare_tra(tra_g)
    a1002 = _analyze_1002(ops_rows)
    ops_rows = _apply_1002_compensation(ops_rows, a1002)

    wb     = openpyxl.Workbook()
    ws_res = wb.active;           ws_res.title = "Resumen"
    ws_ops = wb.create_sheet("Operaciones")
    ws_cau = wb.create_sheet("Cauciones")
    ws_tra = wb.create_sheet("Trading Intraday")
    ws_mav = wb.create_sheet("MAV")
    _write_res(ws_res, ops_rows, cr, tr, mav_g, a1002, arancel_alerts)
    _write_ops_sheet(ws_ops, ops_rows)
    _write_cau_sheet(ws_cau, cr)
    _write_tra_sheet(ws_tra, tr)
    _write_mav_sheet(ws_mav, mav_g)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    cnt_ops = Counter(r["estado"] for r in ops_rows)
    cnt_cau = Counter(r["estado"] for r in cr)

    resumen = {
        "fecha":               datetime.today().strftime("%d-%m-%Y"),
        "n_ops_ok":            cnt_ops.get("OK", 0),
        "n_ops_dif":           cnt_ops.get("DIFERENCIA", 0),
        "n_solo_byma":         cnt_ops.get("SOLO BYMA", 0),
        "n_solo_gallo":        cnt_ops.get("SOLO GALLO", 0),
        "n_ti":                cnt_ops.get("TRADING INTRADAY", 0),
        "n_cau_ok":            cnt_cau.get("OK", 0),
        "n_cau_dif":           cnt_cau.get("DIFERENCIA", 0),
        "n_mav":               len(mav_g),
        "n_falta_boleto":      sum(1 for r in ops_rows if r["verificar"] == "FALTA BOLETO"),
        "n_arancel":           len(arancel_alerts),
        "arancel_alerts":      arancel_alerts,
        "falta_boleto_detail": [r for r in ops_rows if r["verificar"] == "FALTA BOLETO"],
        "con_senebi":          bool(bil_lines),
    }
    return output, resumen
