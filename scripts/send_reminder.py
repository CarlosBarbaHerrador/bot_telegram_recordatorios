import urllib.request
import urllib.parse
import openpyxl
import os
import sys
import math
import tempfile

EXCEL_URL = "https://docs.google.com/spreadsheets/d/1fL7CZ_Td2JAb0Q5RlL_plOMLqrZGl6Ax1zbXJa_kGaE/export?format=xlsx"
SHEET_NAME = " Nigun Grid Luin"

def download_excel(filepath):
    urllib.request.urlretrieve(EXCEL_URL, filepath)

def parse_undead(ws):
    guardia = []
    basura = []

    for row in ws.iter_rows(min_row=13, max_row=18, values_only=True):
        row_list = list(row)

        # Guardia de No Muertos: col 28 (idx 27)=qty, col 29 (idx 28)=name
        if len(row_list) > 28:
            qty = row_list[27]
            name = row_list[28]
            if qty is not None and name is not None:
                name_s = str(name).strip()
                if name_s not in ('Total', 'Velocidad de Reaccion', 'Aceleracion de Reaccion',
                                   'Vista', 'Vista máxima', 'm/s', 'm', 'Multiplicador de Defensa',
                                   'Multiplicador de Evasion', 'Fuerza de stats', 'Fuerza de Cuerpo',
                                   'Kilos', 'Largo') and isinstance(qty, (int, float)) and not math.isnan(qty):
                    guardia.append((int(qty), name_s))

        # Basura: col 38 (idx 37)=qty, col 39 (idx 38)=name
        if len(row_list) > 38:
            bq = row_list[37]
            bn = row_list[38]
            if bq is not None and bn is not None:
                bn_s = str(bn).strip()
                if bn_s not in ('Total',) and isinstance(bq, (int, float)) and not math.isnan(bq):
                    basura.append((int(bq), bn_s))

    # Row 20 additional basura
    for row in ws.iter_rows(min_row=20, max_row=20, values_only=True):
        row_list = list(row)
        if len(row_list) > 38:
            bq = row_list[37]
            bn = row_list[38]
            if bq is not None and bn is not None:
                bn_s = str(bn).strip()
                if bn_s not in ('Total',) and isinstance(bq, (int, float)) and not math.isnan(bq):
                    if all(b[1] != bn_s for b in basura):
                        basura.append((int(bq), bn_s))

    return guardia, basura

def build_message(guardia, basura):
    total = sum(q for q, _ in guardia)
    lines = ["Tienes que hacer Periodo  ", "Periodo de Ningun", f"No muertos: {total}"]
    for qty, name in guardia:
        lines.append(f"- {qty}x {name}")
    if basura:
        for qty, name in basura:
            lines.append(f"Basura: {qty}x {name}")
    return "\n".join(lines)

def send_telegram(token, chat_id, message):
    data = urllib.parse.urlencode({'chat_id': chat_id, 'text': message}).encode()
    req = urllib.request.Request(
        f"https://api.telegram.org/bot{token}/sendMessage",
        data=data,
        headers={'Content-Type': 'application/x-www-form-urlencoded'}
    )
    with urllib.request.urlopen(req) as resp:
        return resp.status

def main():
    token = os.environ.get('TELEGRAM_TOKEN')
    chat_id = os.environ.get('TELEGRAM_CHAT_ID')
    if not token or not chat_id:
        print("Faltan TELEGRAM_TOKEN o TELEGRAM_CHAT_ID", file=sys.stderr)
        sys.exit(1)

    filepath = os.path.join(tempfile.gettempdir(), "ningun.xlsx")
    try:
        download_excel(filepath)
    except Exception as e:
        print(f"Error descargando Excel: {e}", file=sys.stderr)
        sys.exit(1)

    try:
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        ws = wb[SHEET_NAME]
        guardia, basura = parse_undead(ws)
        wb.close()
    except Exception as e:
        print(f"Error leyendo Excel: {e}", file=sys.stderr)
        sys.exit(1)

    message = build_message(guardia, basura)
    print("Mensaje generado:")
    print(message)

    status = send_telegram(token, chat_id, message)
    print(f"Telegram response: {status}")

if __name__ == '__main__':
    main()
