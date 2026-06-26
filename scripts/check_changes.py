import urllib.request
import urllib.parse
import openpyxl
import os
import sys
import hashlib
import tempfile
import time

EXCEL_URL = "https://docs.google.com/spreadsheets/d/1fL7CZ_Td2JAb0Q5RlL_plOMLqrZGl6Ax1zbXJa_kGaE/export?format=xlsx"
SHEET_NAME = " Nigun Grid Luin"
HASH_FILE = "sheet_check_hash.txt"

def download_excel(filepath, retries=3, delay=5):
    for attempt in range(retries):
        try:
            urllib.request.urlretrieve(EXCEL_URL, filepath)
            return
        except Exception as e:
            if attempt < retries - 1:
                wait = delay * (2 ** attempt)
                print(f"Error descargando (intento {attempt+1}/{retries}): {e}. Reintentando en {wait}s...", file=sys.stderr)
                time.sleep(wait)
            else:
                raise

def compute_sheet_hash(ws):
    h = hashlib.sha256()
    for row in ws.iter_rows(min_row=1, max_row=250, values_only=True):
        for cell in row:
            if cell is not None:
                h.update(str(cell).encode("utf-8"))
    return h.hexdigest()

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

def send_notification(token, chat_id):
    text = "🔔 La hoja de personaje está teniendo cambios"
    data = urllib.parse.urlencode({'chat_id': chat_id, 'text': text}).encode()
    req = urllib.request.Request(
        f"https://api.telegram.org/bot{token}/sendMessage",
        data=data,
        headers={'Content-Type': 'application/x-www-form-urlencoded'}
    )
    with urllib.request.urlopen(req, timeout=15) as resp:
        return resp.status

def main():
    token = os.environ.get('TELEGRAM_TOKEN')
    chat_id = os.environ.get('TELEGRAM_CHAT_ID')
    if not token or not chat_id:
        print("Faltan TELEGRAM_TOKEN o TELEGRAM_CHAT_ID", file=sys.stderr)
        sys.exit(1)

    filepath = os.path.join(tempfile.gettempdir(), "ningun_check.xlsx")
    try:
        download_excel(filepath)
    except Exception as e:
        print(f"Error descargando: {e}", file=sys.stderr)
        sys.exit(1)

    try:
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        ws = wb[SHEET_NAME]
        current_hash = compute_sheet_hash(ws)
        wb.close()
    except Exception as e:
        print(f"Error leyendo: {e}", file=sys.stderr)
        sys.exit(1)

    previous_hash = read_previous_hash()

    if previous_hash is None:
        write_hash(current_hash)
        print("Primera ejecucion, hash guardado. Sin notificacion.")
        sys.exit(0)

    if current_hash != previous_hash:
        write_hash(current_hash)
        status = send_notification(token, chat_id)
        print(f"Cambio detectado, notificacion enviada. Status: {status}")
    else:
        print("Sin cambios.")

if __name__ == '__main__':
    main()
