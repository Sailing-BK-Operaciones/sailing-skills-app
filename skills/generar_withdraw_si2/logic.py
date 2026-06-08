"""
Genera SI2 NASDAQ — Withdraw (BARRIDO) y/o Deposit (TRANSFERENCIA).

Lógica SI2 (separado por ';'):
  col 2  = cuenta (SecuritiesAccount)
  col 3  = CVSA
  col 7  = nodo (SecuritiesAccountOfCounterparty)
  col 10 = VN

Corrección automática: cuenta 233/1000 pertenece al nodo HOUSE pero
Gallo la exporta bajo CLIENT. El script corrige el nodo en el SI2 de
salida. Cuando Gallo lo corrija en origen, la condición no tendrá efecto.
"""

from collections import defaultdict
import xlrd

NODE_LABELS = {
    "80233/555555555": "CLIENT",
    "80233/222222222": "HOUSE",
}
NODE_ORDER = ["80233/555555555", "80233/222222222"]

CUENTAS_FORZAR_HOUSE = {"233/1000"}
NODO_HOUSE  = "80233/222222222"
NODO_CLIENT = "80233/555555555"


def _load_ticker_map(especies_file):
    """Devuelve dict {cvsa_str: ticker} desde ESPECIES.XLS col 0/9."""
    especies_file.seek(0)
    wb = xlrd.open_workbook(file_contents=especies_file.read())
    ws = wb.sheet_by_name("Datos_Fijos_Especies")
    ticker_map = {}
    for r in range(1, ws.nrows):
        cod  = str(ws.cell_value(r, 0)).strip().strip("'").zfill(5)
        norm = str(ws.cell_value(r, 9)).strip().strip("'")
        if cod and norm:
            ticker_map[cod] = norm
    return ticker_map


def _fix_si2(si2_file):
    """
    Lee el SI2, corrige el nodo de las cuentas en CUENTAS_FORZAR_HOUSE
    y retorna (corrected_bytes, reasignadas_count).
    """
    si2_file.seek(0)
    raw = si2_file.read()
    text = raw.decode("utf-8", errors="replace")
    lines = text.splitlines(keepends=True)

    corrected = []
    reasignadas = 0

    for i, line in enumerate(lines):
        stripped = line.rstrip("\n")
        if i == 0 or not stripped:
            corrected.append(line)
            continue
        parts = stripped.split(";")
        if len(parts) < 11:
            corrected.append(line)
            continue
        cuenta = parts[2].strip()
        nodo   = parts[7].strip()
        if cuenta in CUENTAS_FORZAR_HOUSE and nodo == NODO_CLIENT:
            parts[7] = NODO_HOUSE
            reasignadas += 1
            corrected.append(";".join(parts) + "\n")
        else:
            corrected.append(line)

    corrected_bytes = "".join(corrected).encode("utf-8")
    return corrected_bytes, reasignadas


def _parse_node_data(si2_bytes):
    """Agrupa VN por nodo e instrumento a partir del SI2 corregido."""
    node_data = defaultdict(lambda: defaultdict(int))
    text = si2_bytes.decode("utf-8", errors="replace")
    for line in text.splitlines()[1:]:   # skip header
        line = line.strip()
        if not line:
            continue
        parts = line.split(";")
        if len(parts) < 11:
            continue
        nodo = parts[7].strip()
        cvsa = parts[3].strip().zfill(5)
        try:
            qty = int(parts[10])
        except ValueError:
            continue
        node_data[nodo][cvsa] += qty
    return node_data


def _build_resumen_text(node_data, ticker_map) -> str:
    lines = []
    for nodo in NODE_ORDER:
        if nodo not in node_data:
            continue
        label = NODE_LABELS.get(nodo, nodo)
        lines.append(f"NODO {label} {nodo}")
        lines.append("Codigo CVSA;Ticker;VN")
        for cvsa in sorted(node_data[nodo]):
            ticker = ticker_map.get(cvsa, "???")
            qty    = node_data[nodo][cvsa]
            lines.append(f"{cvsa};{ticker};{qty}")
        lines.append("")
    return "\n".join(lines)


def _procesar_one(si2_file, ticker_map) -> dict:
    """Procesa un único SI2 y retorna dict con métricas y contenidos."""
    si2_bytes, reasignadas = _fix_si2(si2_file)
    node_data = _parse_node_data(si2_bytes)
    resumen_text = _build_resumen_text(node_data, ticker_map)

    n_sin_ticker = sum(
        1 for nodo in node_data
        for cvsa in node_data[nodo]
        if ticker_map.get(cvsa, "???") == "???"
    )

    def _stats(nodo):
        d = node_data.get(nodo, {})
        return len(d), sum(d.values())

    n_cl, vn_cl = _stats(NODO_CLIENT)
    n_hs, vn_hs = _stats(NODO_HOUSE)

    return {
        "si2_bytes":    si2_bytes,
        "txt_content":  resumen_text,
        "n_client":     n_cl,
        "vn_client":    vn_cl,
        "n_house":      n_hs,
        "vn_house":     vn_hs,
        "n_sin_ticker": n_sin_ticker,
        "reasignadas":  reasignadas,
    }


def procesar(barrido_file, transferencia_file, especies_file):
    """
    Parámetros
    ----------
    barrido_file       : file-like or None — BARRIDO*.SI2 (Withdraw)
    transferencia_file : file-like or None — TRANSFERENCIA*.SI2 (Deposit)
    especies_file      : file-like          — ESPECIES.XLS

    Retorna
    -------
    withdraw_result : dict or None
    deposit_result  : dict or None
    """
    ticker_map = _load_ticker_map(especies_file)

    withdraw_result = _procesar_one(barrido_file,       ticker_map) if barrido_file       else None
    deposit_result  = _procesar_one(transferencia_file, ticker_map) if transferencia_file else None

    return withdraw_result, deposit_result
