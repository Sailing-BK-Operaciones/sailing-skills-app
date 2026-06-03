"""
Arreglo Dev Gara Gallo — adaptación web (Streamlit).

Dos pasos en una sola ejecución:
1. Convierte el archivo exportado por Gallo (BARRIDO.SI2-XXXX.XXXXXXX, extensión
   incorrecta) al formato SI2 válido para subir a NASDAQ (BARRIDO.SI2).
2. Genera Withdraw.txt agrupando VN por nodo CLIENT / HOUSE con tickers de ESPECIES.XLS.

Formato SI2: separado por ';'
  col 3  = CVSA  (zero-pad a 5 dígitos)
  col 7  = nodo  (80233/555555555 = CLIENT | 80233/222222222 = HOUSE)
  col 10 = VN
"""

from collections import defaultdict
import io
import xlrd

NODE_LABELS = {
    "80233/555555555": "CLIENT",
    "80233/222222222": "HOUSE",
}
NODE_ORDER = ["80233/555555555", "80233/222222222"]


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


def _parse_si2(content: bytes):
    """Parsea contenido SI2 y devuelve {nodo: {cvsa: vn_total}}."""
    node_data = defaultdict(lambda: defaultdict(int))
    text = content.decode("utf-8", errors="replace")
    for line in text.splitlines()[1:]:   # skip header
        line = line.strip()
        if not line:
            continue
        parts = line.split(";")
        if len(parts) < 11:
            continue
        nodo = parts[7]
        cvsa = parts[3].strip().zfill(5)
        try:
            qty = int(parts[10])
        except ValueError:
            continue
        node_data[nodo][cvsa] += qty
    return node_data


def _build_withdraw_text(node_data, ticker_map) -> str:
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


def procesar(barrido_file, especies_file):
    """
    Parámetros
    ----------
    barrido_file  : file-like — archivo BARRIDO exportado por Gallo (cualquier extensión)
    especies_file : file-like — ESPECIES.XLS

    Retorna
    -------
    si2_bytes      : bytes    contenido del SI2 corregido (BARRIDO.SI2)
    withdraw_text  : str      contenido de Withdraw.txt
    resumen        : dict     {n_client, vn_client, n_house, vn_house, n_sin_ticker}
    """
    barrido_file.seek(0)
    si2_bytes = barrido_file.read()

    ticker_map = _load_ticker_map(especies_file)
    node_data  = _parse_si2(si2_bytes)

    withdraw_text = _build_withdraw_text(node_data, ticker_map)

    n_sin_ticker = sum(
        1 for nodo in node_data
        for cvsa in node_data[nodo]
        if ticker_map.get(cvsa, "???") == "???"
    )

    def _stats(nodo):
        d = node_data.get(nodo, {})
        return len(d), sum(d.values())

    n_cl, vn_cl = _stats("80233/555555555")
    n_hs, vn_hs = _stats("80233/222222222")

    resumen = {
        "n_client":    n_cl,
        "vn_client":   vn_cl,
        "n_house":     n_hs,
        "vn_house":    vn_hs,
        "n_sin_ticker": n_sin_ticker,
    }

    return si2_bytes, withdraw_text, resumen
