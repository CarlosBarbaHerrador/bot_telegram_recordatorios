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
              'Vista', 'Vista maxima', 'Vista máxima', 'm/s', 'm',
              'Multiplicador de Defensa', 'Multiplicador de Evasion',
              'Fuerza de stats', 'Fuerza de Cuerpo', 'Kilos', 'Largo'}

def download_excel(filepath):
    urllib.request.urlretrieve(EXCEL_URL, filepath)

def is_valid_number(v):
    return isinstance(v, (int, float)) and not math.isnan(v)

def parse_sections(ws):
    guardia = []
    apostatas = []
    basura = []

    for row in ws.iter_rows(min_row=13, max_row=18, values_only=True):
        row_list = list(row)

        g_qty = row_list[27] if len(row_list) > 27 else None
        g_name = row_list[28] if len(row_list) > 28 else None
        a_name = row_list[33] if len(row_list) > 33 else None
        b_qty = row_list[37] if len(row_list) > 37 else None
        b_name = row_list[38] if len(row_list) > 38 else None

        # Guardia de No Muertos
        if is_valid_number(g_qty) and g_name is not None:
            gs = str(g_name).strip()
            if gs not in SKIP_NAMES:
                guardia.append((int(g_qty), gs))

        # Apostatas de Myrkull
        if a_name is not None:
            astr = str(a_name).strip()
            if astr not in SKIP_NAMES:
                apostatas.append(astr)

        # Basura
        if is_valid_number(b_qty) and b_name is not None:
            bns = str(b_name).strip()
            if bns not in SKIP_NAMES:
                basura.append((int(b_qty), bns))

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

    return guardia, apostatas, basura

def build_message(guardia, apostatas, basura):
    total = sum(q for q, _ in guardia)
    lines = ["Tienes que hacer Periodo  ", "Periodo de Ningun", ""]

    lines.append(f"Guardia de No Muertos: {total}")
    for qty, name in guardia:
        lines.append(f"- {qty} {name}")

    lines.append("")
    lines.append("Apostatas de Myrkull:")
    for name in apostatas:
        lines.append(f"- {name}")

    lines.append("")
    lines.append("Basura:")
    if basura:
        for qty, name in basura:
            lines.append(f"- {qty} {name}")
    else:
        lines.append("- 0")

    lines.append("")
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
        guardia, apostatas, basura = parse_sections(ws)
        wb.close()
    except Exception as e:
        print(f"Error leyendo Excel: {e}", file=sys.stderr)
        sys.exit(1)

    message = build_message(guardia, apostatas, basura)
    print("Mensaje generado:")
    print(message)

    status = send_telegram(token, chat_id, message)
    print(f"Telegram response: {status}")

if __name__ == '__main__':
    main()
