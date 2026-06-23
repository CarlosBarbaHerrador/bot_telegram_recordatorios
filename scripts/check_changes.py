import json
import os
import sys
import urllib.parse
import urllib.request
from google.auth.transport.requests import AuthorizedSession
from google.oauth2.service_account import Credentials

SCOPES = ['https://www.googleapis.com/auth/drive.activity.readonly']
FILE_ID = "1fL7CZ_Td2JAb0Q5RlL_plOMLqrZGl6Ax1zbXJa_kGaE"
MY_EMAIL = "nadie@nadie.com"
LAST_TIME_FILE = "last_activity_time.txt"

def get_credentials():
    creds_json = os.environ.get('GOOGLE_CREDENTIALS')
    if not creds_json:
        print("Falta GOOGLE_CREDENTIALS", file=sys.stderr)
        sys.exit(1)
    return Credentials.from_service_account_info(
        json.loads(creds_json), scopes=SCOPES
    )

def read_last_time():
    try:
        if os.path.exists(LAST_TIME_FILE):
            with open(LAST_TIME_FILE) as f:
                return f.read().strip()
    except Exception:
        pass
    return None

def write_last_time(t):
    try:
        with open(LAST_TIME_FILE, "w") as f:
            f.write(t)
    except Exception:
        pass

def send_telegram(token, chat_id, text):
    data = urllib.parse.urlencode({
        'chat_id': chat_id,
        'text': text,
        'parse_mode': 'Markdown'
    }).encode()
    req = urllib.request.Request(
        f"https://api.telegram.org/bot{token}/sendMessage",
        data=data,
        headers={'Content-Type': 'application/x-www-form-urlencoded'}
    )
    with urllib.request.urlopen(req, timeout=15) as resp:
        return resp.status

def send_error_alert(token, chat_id, error_msg):
    text = f"*⚠️ Error en monitor de cambios:*\n{error_msg}"
    send_telegram(token, chat_id, text)

def main():
    token = os.environ.get('TELEGRAM_TOKEN')
    chat_id = os.environ.get('TELEGRAM_CHAT_ID')
    if not token or not chat_id:
        print("Faltan TELEGRAM_TOKEN o TELEGRAM_CHAT_ID", file=sys.stderr)
        sys.exit(1)

    creds = get_credentials()
    session = AuthorizedSession(creds)

    last_time = read_last_time()
    first_run = last_time is None

    body = {'itemName': f'items/{FILE_ID}'}
    if last_time:
        body['filter'] = f'time > "{last_time}"'

    try:
        resp = session.post(
            'https://driveactivity.googleapis.com/v2/activity:query',
            json=body
        )
        resp.raise_for_status()
        data = resp.json()
    except Exception as e:
        msg = f"Error consultando Drive Activity API: {e}"
        print(msg, file=sys.stderr)
        send_error_alert(token, chat_id, msg)
        sys.exit(1)

    activities = data.get('activities', [])

    if not activities:
        print("Sin actividad nueva.")
        sys.exit(0)

    import json as _json
    print("DEBUG primer activity:", _json.dumps(activities[0], indent=2)[:500])

    latest_time = None
    for act in activities:
        ts = act.get('timestamp', '')
        if ts and (latest_time is None or ts > latest_time):
            latest_time = ts

    if latest_time:
        write_last_time(latest_time)

    if first_run:
        print(f"Primera ejecucion. {len(activities)} eventos registrados, sin notificar.")
        sys.exit(0)

    notified = False
    for act in activities:
        action = act.get('primaryActionDetail', {})
        if 'edit' not in action and 'create' not in action:
            continue

        actors = act.get('actors', [])
        who = None
        for actor in actors:
            user = actor.get('user', {})
            known = user.get('knownUser', {})
            email = known.get('emailAddress')
            if email:
                who = email
                break
            if any(k in user for k in ('unknownUser', 'anonymousUser', 'unauthenticated')):
                who = None
                break

        if who == MY_EMAIL:
            print(f"Edit detectado de mi usuario ({who}), omitido.")
            continue

        if who:
            msg = f"👤 *{who}* ha modificado la hoja de personaje"
        else:
            msg = "🔔 Alguien (anónimo) ha modificado la hoja de personaje"

        print(f"Notificando: {msg}")
        send_telegram(token, chat_id, msg)
        notified = True

    if not notified:
        print("Actividad detectada pero solo de mi usuario o sin relevancia.")

if __name__ == '__main__':
    main()
