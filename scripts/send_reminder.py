import urllib.request
import urllib.parse
import openpyxl
import os
import sys
import math
import tempfile
import time
import hashlib
from datetime import date

EXCEL_URL = "https://docs.google.com/spreadsheets/d/1fL7CZ_Td2JAb0Q5RlL_plOMLqrZGl6Ax1zbXJa_kGaE/export?format=xlsx"
SHEET_NAME = " Nigun Grid Luin"

SECTIONS = {
    'Guardia de No Muertos': 'guardia',
    'Apostatas de Myrkull': 'apostatas',
    'Basura': 'basura',
}

SKIP_NAMES = {'Total'}

HASH_FILE = "sheet_hash.txt"

def download_excel(filepath):
    urllib.request.urlretrieve(EXCEL_URL, filepath)

def is_valid_number(v):
    return isinstance(v, (int, float)) and not math.isnan(v)

def compute_sheet_hash(ws):
    h = hashlib.sha256()
    for row in ws.iter_rows(min_row=1, max_row=250, values_only=True):
        for cell in row:
            if cell is not None:
                h.update(str(cell).encode("utf-8"))
    return h.hexdigest()

def parse_sections(ws):
    data = {'guardia': [], 'apostatas': [], 'basura': []}
    current_section = None

    for row in ws.iter_rows(min_row=1, values_only=True):
        row_list = list(row)
        c23 = row_list[22] if len(row_list) > 22 else None
        c23_str = str(c23).strip() if c23 is not None else None

        if c23_str in SECTIONS:
            current_section = SECTIONS[c23_str]
            continue

        if current_section is None:
            continue

        if is_valid_number(c23):
            qty = int(c23)
            name = str(row_list[23]).strip() if len(row_list) > 23 and row_list[23] is not None else ''
            if name not in SKIP_NAMES and name:
                data[current_section].append((qty, name))
        else:
            if c23_str in SECTIONS:
                current_section = SECTIONS[c23_str]
            else:
                pass

    return data['guardia'], data['apostatas'], data['basura']

def read_ganancia():
    path = os.path.join(os.path.dirname(__file__), "..", "ganancia.txt")
    with open(path, encoding="utf-8") as f:
        return f.read().strip()

def read_previous_hash():
    try:
        if os.path.exists(HASH_FILE):
            with open(HASH_FILE) as f:
                return f.read().strip()
    except Exception:
        pass
    return None

def write_hash(h):
    try:
        with open(HASH_FILE, "w") as f:
            f.write(h)
    except Exception:
        pass

def build_message(guardia, apostatas, basura, ganancia, changed):
    guardia_total = sum(q for q, _ in guardia)
    apostatas_total = sum(q for q, _ in apostatas)
    basura_total = sum(q for q, _ in basura)
    gran_total = guardia_total + apostatas_total + basura_total

    sep = "-" * 30

    today = date.today().strftime("%d/%m/%Y (%A)")
    dias = {"Monday": "Lunes", "Tuesday": "Martes", "Wednesday": "Miércoles",
            "Thursday": "Jueves", "Friday": "Viernes", "Saturday": "Sábado", "Sunday": "Domingo"}
    for en, es in dias.items():
        today = today.replace(en, es)
    lines = []
    if changed:
        lines.append("*⚠️ Cambios detectados en la hoja Nigun desde el último periodo*")
        lines.append("")
    lines.append(f"Periodo de Ningun - {today}")
    lines.append("")

    lines.append(f"*Guardia de No Muertos ---> {guardia_total}*")
    for qty, name in guardia:
        lines.append(f"{qty} {name}")
    lines.append(sep)

    lines.append(f"*Apostatas de Myrkull ---> {apostatas_total}*")
    for qty, name in apostatas:
        lines.append(f"{qty} {name}")
    lines.append(sep)

    lines.append(f"*Basura ---> {basura_total}*")
    for qty, name in basura:
        lines.append(f"{qty} {name}")
    lines.append(sep)

    lines.append(f"*Total general: {gran_total}*")
    lines.append("=" * 40)
    lines.append(ganancia)
    return "\n".join(lines)

def send_telegram(token, chat_id, message):
    data = urllib.parse.urlencode({'chat_id': chat_id, 'text': message, 'parse_mode': 'Markdown'}).encode()
    req = urllib.request.Request(
        f"https://api.telegram.org/bot{token}/sendMessage",
        data=data,
        headers={'Content-Type': 'application/x-www-form-urlencoded'}
    )
    with urllib.request.urlopen(req) as resp:
        return resp.status

ERROR_LOCK = os.path.join(tempfile.gettempdir(), "periodo_error.lock")
ERROR_COOLDOWN = 3600

def can_send_error():
    try:
        if os.path.exists(ERROR_LOCK):
            with open(ERROR_LOCK) as f:
                last = float(f.read().strip())
            if time.time() - last < ERROR_COOLDOWN:
                return False
        return True
    except Exception:
        return True

def mark_error_sent():
    try:
        with open(ERROR_LOCK, "w") as f:
            f.write(str(time.time()))
    except Exception:
        pass

def send_error_alert(token, chat_id, error_msg):
    if not can_send_error():
        print("Error alert omitido (cooldown de 1h)", file=sys.stderr)
        return
    text = f"*⚠️ Error en Periodo Bot:*\n{error_msg}"
    try:
        data = urllib.parse.urlencode({'chat_id': chat_id, 'text': text, 'parse_mode': 'Markdown'}).encode()
        req = urllib.request.Request(
            f"https://api.telegram.org/bot{token}/sendMessage",
            data=data,
            headers={'Content-Type': 'application/x-www-form-urlencoded'}
        )
        with urllib.request.urlopen(req, timeout=15):
            pass
        mark_error_sent()
    except Exception:
        pass

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
        msg = f"No pude descargar el Excel: {e}"
        print(msg, file=sys.stderr)
        send_error_alert(token, chat_id, msg)
        sys.exit(1)

    try:
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        ws = wb[SHEET_NAME]

        current_hash = compute_sheet_hash(ws)
        previous_hash = read_previous_hash()
        changed = previous_hash is not None and current_hash != previous_hash

        guardia, apostatas, basura = parse_sections(ws)
        wb.close()
    except Exception as e:
        msg = f"No pude leer el Excel (cambio de estructura?): {e}"
        print(msg, file=sys.stderr)
        send_error_alert(token, chat_id, msg)
        sys.exit(1)

    if not guardia and not apostatas and not basura:
        msg = "No encontré ninguna sección de no-muertos en el Excel. ¿Cambiaste la estructura?"
        print(msg, file=sys.stderr)
        send_error_alert(token, chat_id, msg)
        sys.exit(1)

    write_hash(current_hash)

    message = build_message(guardia, apostatas, basura, read_ganancia(), changed)
    print("Mensaje generado:")
    print(message)

    status = send_telegram(token, chat_id, message)
    print(f"Telegram response: {status}")

if __name__ == '__main__':
    main()
