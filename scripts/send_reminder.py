import urllib.request
import urllib.parse
import openpyxl
import os
import sys
import math
import tempfile

EXCEL_URL = "https://docs.google.com/spreadsheets/d/1fL7CZ_Td2JAb0Q5RlL_plOMLqrZGl6Ax1zbXJa_kGaE/export?format=xlsx"
SHEET_NAME = " Nigun Grid Luin"

SECTIONS = {
    'Guardia de No Muertos': 'guardia',
    'Apostatas de Myrkull': 'apostatas',
    'Basura': 'basura',
}

SKIP_NAMES = {'Total'}

def download_excel(filepath):
    urllib.request.urlretrieve(EXCEL_URL, filepath)

def is_valid_number(v):
    return isinstance(v, (int, float)) and not math.isnan(v)

def parse_sections(ws):
    data = {'guardia': [], 'apostatas': [], 'basura': []}
    current_section = None

    for row in ws.iter_rows(min_row=1, values_only=True):
        row_list = list(row)
        c23 = row_list[22] if len(row_list) > 22 else None

        c23_str = str(c23).strip() if c23 is not None else None

        # Check for section headers
        if c23_str in SECTIONS:
            current_section = SECTIONS[c23_str]
            continue

        # If not in a section, skip
        if current_section is None:
            continue

        # If col 23 is a number, it's a data row
        if is_valid_number(c23):
            qty = int(c23)
            name = str(row_list[23]).strip() if len(row_list) > 23 and row_list[23] is not None else ''
            if name not in SKIP_NAMES and name:
                data[current_section].append((qty, name))
        else:
            # Non-numeric in col 23 means section is over
            # But could be a description row, only end if we hit next section header
            if c23_str in SECTIONS:
                current_section = SECTIONS[c23_str]
            else:
                # Description row or other text - just skip, section continues
                pass

    return data['guardia'], data['apostatas'], data['basura']

def build_message(guardia, apostatas, basura):
    guardia_total = sum(q for q, _ in guardia)
    apostatas_total = sum(q for q, _ in apostatas)
    basura_total = sum(q for q, _ in basura)
    gran_total = guardia_total + apostatas_total + basura_total

    lines = ["Tienes que hacer Periodo  ", "Periodo de Ningun", ""]

    lines.append("Guardia de No Muertos")
    for qty, name in guardia:
        lines.append(f"{qty} {name}")
    lines.append(f"Total: {guardia_total}")

    lines.append("")
    lines.append("Apostatas de Myrkull")
    for qty, name in apostatas:
        lines.append(f"{qty} {name}")
    lines.append(f"Total: {apostatas_total}")

    lines.append("")
    lines.append("Basura")
    for qty, name in basura:
        lines.append(f"{qty} {name}")
    lines.append(f"Total: {basura_total}")

    lines.append("")
    lines.append(f"Total general: {gran_total}")
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
