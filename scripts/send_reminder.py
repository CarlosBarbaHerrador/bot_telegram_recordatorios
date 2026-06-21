import urllib.request
import urllib.parse
import openpyxl
import os
import sys
import math
import tempfile

EXCEL_URL = "https://docs.google.com/spreadsheets/d/1fL7CZ_Td2JAb0Q5RlL_plOMLqrZGl6Ax1zbXJa_kGaE/export?format=xlsx"
SHEET_NAME = " Nigun Grid Luin"

SKIP_NAMES = {'Total', 'Velocidad de Reaccion', 'Aceleracion de Reaccion',
              'Vista', 'Vista máxima', 'm/s', 'm', 'Multiplicador de Defensa',
              'Multiplicador de Evasion', 'Fuerza de stats', 'Fuerza de Cuerpo',
              'Kilos', 'Largo'}

def download_excel(filepath):
    urllib.request.urlretrieve(EXCEL_URL, filepath)

def is_valid_number(v):
    return isinstance(v, (int, float)) and not math.isnan(v)

def parse_undead(ws):
    entries = []
    basura = []

    for row in ws.iter_rows(min_row=13, max_row=18, values_only=True):
        row_list = list(row)

        qty = row_list[27] if len(row_list) > 27 else None
        guardia_name = row_list[28] if len(row_list) > 28 else None
        apostata_name = row_list[33] if len(row_list) > 33 else None
        bq = row_list[37] if len(row_list) > 37 else None
        bn = row_list[38] if len(row_list) > 38 else None

        # Guardia + Apostata combined
        if is_valid_number(qty):
            names = []
            if guardia_name is not None:
                gs = str(guardia_name).strip()
                if gs not in SKIP_NAMES:
                    names.append(gs)
            if apostata_name is not None:
                astr = str(apostata_name).strip()
                if astr not in SKIP_NAMES:
                    names.append(astr)
            if names:
                entries.append((int(qty), " / ".join(names)))

        # Basura
        if is_valid_number(bq) and bn is not None:
            bns = str(bn).strip()
            if bns not in SKIP_NAMES:
                basura.append((int(bq), bns))

    # Row 20 additional basura
    for row in ws.iter_rows(min_row=20, max_row=20, values_only=True):
        row_list = list(row)
        if len(row_list) > 38:
            bq = row_list[37]
            bn = row_list[38]
            if is_valid_number(bq) and bn is not None:
                bns = str(bn).strip()
                if bns not in SKIP_NAMES and all(b[1] != bns for b in basura):
                    basura.append((int(bq), bns))

    return entries, basura

def build_message(entries, basura):
    total = sum(q for q, _ in entries)
    lines = ["Tienes que hacer Periodo  ", "Periodo de Ningun", f"No muertos: {total}"]
    for qty, name in entries:
        lines.append(f"- {qty} {name}")
    if basura:
        for qty, name in basura:
            lines.append(f"Basura: {qty} {name}")
    lines.append("----------------")
    lines.append("+1 Esqueleto arquero basico")
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
        entries, basura = parse_undead(ws)
        wb.close()
    except Exception as e:
        print(f"Error leyendo Excel: {e}", file=sys.stderr)
        sys.exit(1)

    message = build_message(entries, basura)
    print("Mensaje generado:")
    print(message)

    status = send_telegram(token, chat_id, message)
    print(f"Telegram response: {status}")

if __name__ == '__main__':
    main()
