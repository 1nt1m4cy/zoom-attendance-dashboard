#!/usr/bin/env python3
"""
Attendance Companion Dashboard
Professional web dashboard for Zoom attendance collection.

Credentials are loaded from environment variables or a .env file.
See .env.example for the required variables.
"""

import re, time, signal, sys, os, platform, json, threading, queue, random, socket, webbrowser, gc, csv
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

for pkg, mod in [('flask','flask'),('openpyxl','openpyxl'),('playwright','playwright.sync_api'),('python-dotenv','dotenv')]:
    try: __import__(mod)
    except ImportError: os.system(f"{sys.executable} -m pip install {pkg}")

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from flask import Flask, Response, request, send_file, jsonify, render_template_string
from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ── CREDENTIALS (from environment -- never hardcoded) ────────────────────────
BOT_NAME = os.environ.get('BOT_NAME', 'Attendance Bot')
_ZOOM_DOMAIN = 'zoom.us'

def _load_creds():
    """Load credentials once, return (email, password). Exits if missing."""
    e = os.environ.get('ZOOM_EMAIL', '')
    p = os.environ.get('ZOOM_PASSWORD', '')
    if not e or not p:
        print("\n  ERROR: Zoom credentials not found.")
        print("  Set them via environment variables or a .env file:\n")
        print("    export ZOOM_EMAIL='you@example.com'")
        print("    export ZOOM_PASSWORD='your_password'\n")
        print("  Or create a .env file -- see .env.example\n")
        sys.exit(1)
    return e, p

def _secure_fill(page, selector, value):
    """Fill a field via JS value injection -- no keystrokes emitted, no visible
    typing in headful mode, resistant to keystroke loggers. Triggers input/change
    events so React forms register the value."""
    page.evaluate('''([sel, val]) => {
        const el = document.querySelector(sel);
        if (!el) return;
        const nativeSetter = Object.getOwnPropertyDescriptor(
            window.HTMLInputElement.prototype, 'value').set;
        nativeSetter.call(el, val);
        el.dispatchEvent(new Event('input', {bubbles: true}));
        el.dispatchEvent(new Event('change', {bubbles: true}));
    }''', [selector, value])
    page.wait_for_timeout(random.randint(200, 500))

def _wipe(varname_holder):
    """Overwrite a string in memory and trigger garbage collection."""
    try:
        if isinstance(varname_holder, list) and varname_holder:
            varname_holder[0] = 'x' * len(varname_holder[0])
            varname_holder.clear()
    except: pass
    gc.collect()

def _load_smtp():
    """Load SMTP credentials from environment. Returns (email, password) or (None, None)."""
    e = os.environ.get('SMTP_EMAIL', '')
    p = os.environ.get('SMTP_PASSWORD', '')
    return (e, p) if e and p else (None, None)

def send_attendance_email(recipient, filepath, filename):
    """Send the attendance Excel file as an email attachment."""
    smtp_email, smtp_pass = _load_smtp()
    if not smtp_email or not smtp_pass:
        return False, "SMTP credentials not configured. Set SMTP_EMAIL and SMTP_PASSWORD."

    msg = MIMEMultipart()
    msg['From'] = smtp_email
    msg['To'] = recipient
    msg['Subject'] = f"Attendance Report - {filename}"

    body = (
        f"Attendance report attached: {filename}.xlsx\n\n"
        f"Generated on {datetime.now().strftime('%Y-%m-%d at %H:%M')} "
        f"by Attendance Companion.\n"
    )
    msg.attach(MIMEText(body, 'plain'))

    with open(filepath, 'rb') as f:
        part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{filename}.xlsx"')
    msg.attach(part)

    try:
        with smtplib.SMTP('smtp.gmail.com', 587, timeout=15) as s:
            s.starttls()
            s.login(smtp_email, smtp_pass)
            s.send_message(msg)
        return True, "Sent"
    except smtplib.SMTPAuthenticationError:
        return False, "SMTP authentication failed. Check SMTP_EMAIL and SMTP_PASSWORD."
    except Exception as ex:
        return False, str(ex)

# ── NOISE FILTERS ────────────────────────────────────────────────────────────
NOISE_WORDS = [
    'attendance bot', '(me)', '(host)', 'zoom', 'invite', 'participants',
    'mute', 'unmute', 'chat', 'share', 'react', 'leave', 'more', 'video',
    'audio', 'settings', 'options', 'view', 'screen', 'copyright', 'privacy',
    'legal', 'send report', 'acceleration', 'learn more', 'search', 'close',
    'pop out', 'waiting room', 'breakout', 'recording', 'raise hand',
    'lower hand', 'rename', 'put in waiting', 'remove', 'report',
    'make host', 'make co-host', 'allow to', 'everyone', 'to everyone',
    'joined the meeting', 'left the meeting', 'has been removed',
    'is sharing', 'stopped sharing', 'you are', 'browser settings',
    'joining', 'remaining', 'in meeting', 'connecting', 'waiting',
    'end meeting', 'meeting has ended', 'this meeting has been ended',
    'sign in', 'sign up', 'join audio', 'computer audio', 'phone call',
    'got it', 'i agree', 'accept', 'decline', 'cancel', 'confirm',
    'enter meeting info', 'your name', 'meeting info', 'start video',
    'stop video', 'join meeting', 'meeting id', 'passcode', 'password',
    'remember my name', 'future meeting', 'for future', 'attendee',
    'reclaim host', 'continue the conversation', 'who can see your messages',
    'no one has sent', 'send a message', 'type message here', 'meeting chat',
    'in-meeting chat', 'direct message', 'to:', 'from:', 'new message',
    'you can chat with', 'everyone can see', 'only visible to',
    'products', 'solutions', 'resources', 'plans', 'pricing', 'contact sales',
    'request a demo', 'support', 'download', 'marketplace', 'developer',
    'enterprise', 'small business', 'industries', 'healthcare', 'education',
    'financial', 'government', 'virtual events', 'webinars', 'whiteboard',
    'phone system', 'ai companion', 'workvivo', 'meeting link is invalid',
    'does not exist', 'this meeting is not available', 'has expired',
    'try again', 'go to zoom', 'explore zoom', 'platform', 'communication',
    'collaboration', 'schedule a meeting', 'host a meeting', 'cookie',
    'terms of service', 'about us', 'blog', 'careers', 'partners',
    'investors', 'trust center', 'accessibility',
]

CHAT_PATTERNS = [
    re.compile(r'^From\s+.+\s+to\s+', re.IGNORECASE),
    re.compile(r'has joined the meeting', re.IGNORECASE),
    re.compile(r'has left the meeting', re.IGNORECASE),
    re.compile(r'is waiting', re.IGNORECASE),
    re.compile(r'^\d{1,2}:\d{2}\s*(AM|PM)?', re.IGNORECASE),
    re.compile(r'https?://'),
]

INVALID_MEETING_PHRASES = [
    'this meeting link is invalid', 'this meeting does not exist',
    'meeting is not available', 'this meeting id is not valid',
    'meeting has expired', 'invalid meeting id',
    'the meeting has not started', 'check your meeting link',
]

# ── SHARED STATE ─────────────────────────────────────────────────────────────
lock = threading.Lock()
event_q = queue.Queue()
seen = {}
groups_found = set()
MY_GROUP = 'A'

roster = []
roster_mode = False

state = {
    'phase': 'idle',
    'meeting_url': '',
    'filename': '',
    'total_in_meeting': 0,
    'stop': False,
    'collected': [],
    'ignored': [],
}

# ── UTILITY FUNCTIONS ────────────────────────────────────────────────────────
def extract_id(text):
    m = re.search(r'[\s_(\-/|,]*(\d{7,13})[\s_)\-/|,]*', text)
    return m.group(1) if m else None

def extract_group(text):
    m = re.search(r'(?:^|[\s(\-_])[Gg](?:roup|rp|r|p)?\s*([A-Za-z])\b', text, re.IGNORECASE)
    if m:
        return m.group(1).upper()
    m2 = re.search(r'\d{7,13}\s*\)?\s*([A-Za-z])\s*$', text)
    if m2 and len(m2.group(1)) == 1:
        return m2.group(1).upper()
    return None

def clean_name(text):
    text = re.sub(r'[_(\-/|,\s]*\d{7,13}[_)\-/|,\s]*', ' ', text)
    text = re.sub(r'(?:^|[\s(\-_])[Gg](?:roup|rp|r|p)?\s*[A-Za-z]\b', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\(Me\)', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\(Host\)', '', text, flags=re.IGNORECASE)
    text = re.sub(r'has joined the meeting', '', text, flags=re.IGNORECASE)
    text = re.sub(r'has left the meeting', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\s+[A-Za-z]\s*$', '', text)
    text = re.sub(r'[^A-Za-z\s\'-]', '', text)
    text = text.strip(' -')
    text = re.sub(r'^([A-Z]{1,3})\s+(?=[A-Z][a-z])', '', text)
    return ' '.join(text.split())

def normalize(name):
    return re.sub(r'\s+', ' ', name.strip().lower())

def is_noise(text):
    low = text.lower()
    if any(w in low for w in NOISE_WORDS):
        return True
    if any(p.search(text) for p in CHAT_PATTERNS):
        return True
    if re.match(r'^[A-Z]{1,3}$', text.strip()):
        return True
    return False

def parse_roster(content, filename):
    entries = []
    ext = os.path.splitext(filename)[1].lower()
    lines = []
    if ext == '.json':
        try:
            data = json.loads(content)
            if isinstance(data, list):
                for item in data:
                    if isinstance(item, dict):
                        n = item.get('name', item.get('Name', item.get('student_name', '')))
                        sid = str(item.get('student_id', item.get('studentId', item.get('StudentID', item.get('id', item.get('ID', item.get('index', item.get('index_number', ''))))))))
                        if n: entries.append({'name': n.strip(), 'student_id': sid.strip()})
                    elif isinstance(item, str):
                        lines.append(item)
        except json.JSONDecodeError:
            lines = content.splitlines()
    elif ext == '.xlsx':
        from io import BytesIO
        from openpyxl import load_workbook
        wb = load_workbook(BytesIO(content if isinstance(content, bytes) else content.encode()), read_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return []
        header = [str(c).lower().strip() if c else '' for c in rows[0]]
        name_col = next((i for i, h in enumerate(header) if h in ('name', 'student_name', 'full_name', 'fullname')), 0)
        id_col = next((i for i, h in enumerate(header) if h in ('id', 'student_id', 'index', 'index_number', 'student id', 'index number')), None)
        start = 1 if any(h in ('name', 'student_name', 'full_name', 'fullname', 'id', 'student_id', 'index') for h in header) else 0
        for row in rows[start:]:
            n = str(row[name_col]).strip() if name_col < len(row) and row[name_col] else ''
            sid = str(row[id_col]).strip() if id_col is not None and id_col < len(row) and row[id_col] else ''
            if n and n.lower() != 'none':
                entries.append({'name': n, 'student_id': sid if sid.lower() != 'none' else ''})
        wb.close()
        return entries
    else:
        lines = content.splitlines() if isinstance(content, str) else content.decode('utf-8', errors='ignore').splitlines()

    for line in lines:
        raw = line.strip()
        if re.match(r'^#{1,6}\s', raw):
            continue
        line = re.sub(r'^[\-\*\d\.]+\s*', '', raw)
        line = line.strip('- *#>')
        if not line or len(line) < 3:
            continue
        sid = extract_id(line)
        name = re.sub(r'[_(\-/|,\s]*\d{7,13}[_)\-/|,\s]*', ' ', line).strip()
        name = re.sub(r'[^A-Za-z\s\'-]', '', name).strip()
        if name and len(name) >= 3:
            entries.append({'name': ' '.join(name.split()), 'student_id': sid or ''})
    return entries

def roster_match(zoom_name, zoom_sid):
    zn = normalize(zoom_name)
    zt = set(zn.split())
    best = None
    best_score = 0
    for r in roster:
        rn = normalize(r['name'])
        rt = set(rn.split())
        if zn == rn or (zoom_sid and r['student_id'] and zoom_sid == r['student_id']):
            return r
        if zt and rt:
            common = zt & rt
            smaller = min(len(zt), len(rt))
            if len(common) == smaller and smaller >= 2:
                score = len(common) * 10 + (10 if zoom_sid and r['student_id'] and zoom_sid == r['student_id'] else 0)
                if score > best_score:
                    best = r; best_score = score
            elif len(common) >= 2:
                score = len(common)
                if score > best_score:
                    best = r; best_score = score
    if best and best_score >= 2:
        return best
    return None

def emit(etype, **kwargs):
    event_q.put({'type': etype, **kwargs})

def add_entry(raw_text):
    raw_text = raw_text.strip()
    if not raw_text or len(raw_text) < 3:
        return
    if is_noise(raw_text):
        return
    sid = extract_id(raw_text)
    grp = extract_group(raw_text)
    name = clean_name(raw_text)
    if not name or len(name) < 3:
        return
    if re.match(r'^[A-Z]{1,3}$', name):
        return

    nk = normalize(name)
    key = sid if sid else nk

    with lock:
        if grp:
            groups_found.add(grp)

        if key in seen:
            old = seen[key]
            if sid and not old['student_id']:
                old['student_id'] = sid
            if grp and not old['group']:
                old['group'] = grp
            if len(name) > len(old['name']):
                old['name'] = name
            return

        skip = False
        for k, v in list(seen.items()):
            vn = normalize(v['name'])
            if nk in vn or vn in nk:
                if len(name) > len(v['name']):
                    v['name'] = name
                if sid and not v['student_id']:
                    v['student_id'] = sid
                if grp and not v['group']:
                    v['group'] = grp
                skip = True
                break

        if not skip:
            seen[key] = {"name": name, "student_id": sid or "", "group": grp or ""}
            mixed = bool(len(groups_found) > 1 or (groups_found and groups_found != {MY_GROUP}))
            tag = f"ID:{sid}" if sid else "NO-ID"
            reason = None
            if mixed:
                if grp and grp != MY_GROUP:
                    reason = f"not in Group {MY_GROUP} (Grp {grp})"
                elif not grp and not sid:
                    reason = "no group tag + no student ID"
            else:
                if not sid:
                    reason = "no student ID"

            if reason:
                emit('name', name=name, student_id=sid or '', group=grp or '',
                     status='ignored', reason=reason)
                emit('log', text=f"soul captured but ignored -- {name} ({tag}) | reason: {reason}")
            else:
                emit('name', name=name, student_id=sid or '', group=grp or '',
                     status='collected', reason='')
                gtag = f" [Grp {grp}]" if grp else ""
                emit('log', text=f"caught: {name} ({tag}){gtag}")

            _rebuild_lists()

def _rebuild_lists():
    collected = []
    ignored = []

    if roster_mode and roster:
        matched_roster = set()
        for v in sorted(seen.values(), key=lambda x: x['name'].lower()):
            match = roster_match(v['name'], v.get('student_id', ''))
            if match:
                rn = normalize(match['name'])
                if rn not in matched_roster:
                    matched_roster.add(rn)
                    collected.append({
                        'name': match['name'],
                        'student_id': match['student_id'] or v.get('student_id', ''),
                        'group': v.get('group', ''),
                    })
            else:
                ignored.append({**v, 'reason': 'not on roster'})
        absent = []
        for r in roster:
            if normalize(r['name']) not in matched_roster:
                absent.append({'name': r['name'], 'student_id': r.get('student_id', ''), 'group': '', 'reason': 'absent'})
        state['absent'] = absent
        mixed = False
    else:
        state['absent'] = []
        mixed = bool(len(groups_found) > 1 or (groups_found and groups_found != {MY_GROUP}))
        for v in sorted(seen.values(), key=lambda x: x['name'].lower()):
            if mixed:
                if v['group'] == MY_GROUP:
                    collected.append(v)
                elif v['group'] and v['group'] != MY_GROUP:
                    ignored.append({**v, 'reason': f"not in Group {MY_GROUP} (Grp {v['group']})"})
                elif v['student_id']:
                    collected.append(v)
                else:
                    ignored.append({**v, 'reason': "no group tag + no student ID"})
            else:
                if v['student_id']:
                    collected.append(v)
                else:
                    ignored.append({**v, 'reason': "no student ID"})

    state['collected'] = collected
    state['ignored'] = ignored
    emit('stats', collected=len(collected), ignored=len(ignored),
         total=state['total_in_meeting'],
         groups=sorted(list(groups_found)) if groups_found else [],
         mixed=mixed,
         absent=len(state.get('absent', [])))
    emit('rebuild',
         collected=[{'name':v['name'],'student_id':v.get('student_id',''),'group':v.get('group','')} for v in collected],
         ignored=[{'name':v['name'],'reason':v.get('reason',''),'student_id':v.get('student_id',''),'group':v.get('group','')} for v in ignored],
         absent=[{'name':v['name'],'student_id':v.get('student_id','')} for v in state.get('absent', [])])

def get_desktop():
    home = os.path.expanduser("~")
    if platform.system() == "Windows":
        d = os.path.join(home, "Desktop")
        if not os.path.isdir(d):
            d = os.path.join(home, "OneDrive", "Desktop")
    elif platform.system() == "Darwin":
        d = os.path.join(home, "Desktop")
    else:
        d = os.path.join(home, "Desktop")
        if not os.path.isdir(d):
            d = home
    os.makedirs(d, exist_ok=True)
    return d

def save_excel(filepath):
    with lock:
        collected = list(state['collected'])
        absent = list(state.get('absent', []))
    with_id = [v for v in collected if v['student_id']]
    without_id = [v for v in collected if not v['student_id']]

    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance"
    hf = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    hfont = Font(bold=True, size=12, color="FFFFFF")
    tb = Border(left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin'))
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 18
    ws.append(['No', 'Name', 'Student ID'])
    for c in ws[1]:
        c.font = hfont; c.fill = hf; c.alignment = Alignment(horizontal='center'); c.border = tb
    n = 1
    for v in with_id:
        ws.append([n, v['name'], v['student_id']]); n += 1
        for c in ws[ws.max_row]: c.border = tb; c.alignment = Alignment(horizontal='left')
    if without_id:
        ws.append([])
        ws.append(['', '--- No Student ID ---', ''])
        for c in ws[ws.max_row]: c.font = Font(bold=True, italic=True, color="FF0000")
        for v in without_id:
            ws.append([n, v['name'], '']); n += 1
            for c in ws[ws.max_row]: c.border = tb
    if absent:
        ws.append([])
        ws.append([])
        af = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
        afont = Font(bold=True, size=12, color="FFFFFF")
        ws.append(['No', 'Name (Absent)', 'Student ID'])
        for c in ws[ws.max_row]:
            c.font = afont; c.fill = af; c.alignment = Alignment(horizontal='center'); c.border = tb
        for i, v in enumerate(absent, 1):
            ws.append([i, v['name'], v.get('student_id', '')])
            for c in ws[ws.max_row]: c.border = tb; c.alignment = Alignment(horizontal='left')
    wb.save(filepath)
    return filepath

def save_csv(filepath):
    with lock:
        collected = list(state['collected'])
        absent = list(state.get('absent', []))
    with open(filepath, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(['No', 'Name', 'Student ID'])
        for i, v in enumerate(collected, 1):
            w.writerow([i, v['name'], v['student_id']])
        if absent:
            w.writerow([])
            w.writerow(['--- ABSENT ---', '', ''])
            w.writerow(['No', 'Name (Absent)', 'Student ID'])
            for i, v in enumerate(absent, 1):
                w.writerow([i, v['name'], v.get('student_id', '')])
    return filepath

# ── PLAYWRIGHT AUTOMATION THREAD ─────────────────────────────────────────────
def run_bot(meeting_url):
    emit('phase', phase='launching')
    emit('log', text='booting up the shadow client...')

    try:
        with sync_playwright() as p:
            br = p.chromium.launch(
                headless=True,
                args=['--no-sandbox', '--disable-dev-shm-usage',
                      '--disable-blink-features=AutomationControlled',
                      '--disable-features=AutomationControlled',
                      '--disable-infobars', '--window-size=1920,1080',
                      '--use-fake-ui-for-media-stream',
                      '--use-fake-device-for-media-stream'])
            ctx = br.new_context(
                viewport={'width': 1920, 'height': 1080},
                user_agent='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
                locale='en-US', permissions=['microphone', 'camera'])
            pg = ctx.new_page()
            pg.add_init_script("""
                Object.defineProperty(navigator,'webdriver',{get:()=>false});
                delete navigator.__proto__.webdriver;
                window.chrome={runtime:{},loadTimes:function(){},csi:function(){}};
                Object.defineProperty(navigator,'plugins',{get:()=>[1,2,3,4,5]});
                Object.defineProperty(navigator,'languages',{get:()=>['en-US','en']});
                Object.defineProperty(navigator,'platform',{get:()=>'Linux x86_64'});
                Object.defineProperty(navigator,'hardwareConcurrency',{get:()=>8});
                Object.defineProperty(navigator,'deviceMemory',{get:()=>8});
                window.outerWidth=window.innerWidth;
                window.outerHeight=window.innerHeight+85;
            """)

            def human_type(el, text):
                el.click(); pg.wait_for_timeout(random.randint(200, 400))
                for ch in text:
                    el.type(ch, delay=random.randint(40, 90))
                    if random.random() < 0.1:
                        pg.wait_for_timeout(random.randint(100, 300))
                pg.wait_for_timeout(random.randint(300, 600))

            def dismiss_popups():
                for label in ["Got it","OK","Later","Skip","Dismiss","Close","Not Now",
                              "I Agree","Join Audio by Computer","Join with Computer Audio","Accept"]:
                    try:
                        btn = pg.query_selector(f'button:has-text("{label}")')
                        if btn and btn.is_visible(): btn.click(); pg.wait_for_timeout(300)
                    except: pass

            # ── SECURE LOGIN ──────────────────────────────────────────
            # Credentials loaded into short-lived locals, used once, then wiped.
            # Password injected via JS setter (no keystroke events emitted) so it
            # cannot be captured by keystroke loggers, Playwright traces, screen
            # recordings in headful mode, or DevTools event listeners.
            _cred_email, _cred_pw = _load_creds()
            _pw_holder = [_cred_pw]
            del _cred_pw

            emit('phase', phase='logging_in')
            emit('log', text='slipping through the front door... shhh')
            pg.goto("https://app.zoom.us/signin#/login", wait_until='domcontentloaded', timeout=30000)
            pg.wait_for_timeout(random.randint(3000, 5000))

            # Domain gate: refuse to enter credentials on anything except zoom.us
            if _ZOOM_DOMAIN not in pg.url:
                emit('phase', phase='error')
                emit('error', text='Unexpected domain. Credentials withheld.')
                _wipe(_pw_holder); br.close(); return

            pg.mouse.move(random.randint(400, 800), random.randint(300, 500))
            pg.wait_for_timeout(random.randint(500, 1000))
            try:
                cb = pg.query_selector('#onetrust-accept-btn-handler')
                if cb and cb.is_visible(): cb.click(); pg.wait_for_timeout(500)
            except: pass

            # Email: human-typed (not sensitive, aids anti-bot stealth)
            ei = pg.query_selector('input[name="account"]')
            if ei: human_type(ei, _cred_email)
            del _cred_email
            pg.wait_for_timeout(random.randint(300, 600))
            try: pg.get_by_role("button", name="Next").first.click()
            except: pass
            pg.wait_for_timeout(random.randint(4000, 6000))

            # Domain re-check before password entry
            if _ZOOM_DOMAIN not in pg.url:
                emit('phase', phase='error')
                emit('error', text='Redirected off Zoom. Password withheld.')
                _wipe(_pw_holder); br.close(); return

            # Password: injected via JS value setter -- zero keystroke events,
            # invisible in headful mode, not capturable by input event listeners.
            pi = pg.query_selector('input[name="password"]')
            if pi and pi.is_visible():
                _secure_fill(pg, 'input[name="password"]', _pw_holder[0])
                pg.wait_for_timeout(random.randint(300, 700))
                try: pg.get_by_role("button", name="Sign in").first.click()
                except: pass

            # Wipe password from memory immediately after submission
            _wipe(_pw_holder)
            # Clear the password field in the DOM so DevTools can't read it
            pg.evaluate('()=>{let p=document.querySelector("input[name=password]");if(p)p.value="";}')

            pg.wait_for_timeout(12000)
            dismiss_popups()
            if 'passkey' in pg.url.lower():
                pg.goto("https://app.zoom.us/wc", wait_until='domcontentloaded', timeout=15000)
                pg.wait_for_timeout(3000)

            url = pg.url
            if 'myhome' in url or 'meetings' in url or '/wc' in url:
                emit('log', text="we're in. zero alarms triggered.")
                emit('phase', phase='logged_in')
            else:
                emit('log', text="login felt off, pushing through anyway")

            # JOIN
            emit('phase', phase='joining')
            emit('log', text='pulling up to the lecture...')
            wc_url = meeting_url.replace('/j/', '/wc/join/')
            pg.goto(wc_url, wait_until='domcontentloaded', timeout=30000)
            pg.wait_for_timeout(6000)
            dismiss_popups()

            ni = pg.query_selector('input#inputname, input[placeholder*="name" i]')
            if ni: ni.fill(BOT_NAME); pg.wait_for_timeout(500)

            emit('log', text='knocking on the door...')
            pg.evaluate('()=>{document.querySelectorAll("[class*=preview]").forEach(e=>e.style.pointerEvents="none")}')
            pg.wait_for_timeout(300)

            clicked = pg.evaluate('''() => {
                let btns = document.querySelectorAll('button');
                for(let b of btns) {
                    let t = b.textContent.trim();
                    if(t === 'Join' || t === 'Join Meeting') {
                        b.style.pointerEvents = 'auto';
                        b.dispatchEvent(new MouseEvent('click',{bubbles:true,cancelable:true}));
                        return 'js';
                    }
                }
                let pb = document.querySelector('button.zm-btn--primary, button.preview-join-button');
                if(pb) { pb.dispatchEvent(new MouseEvent('click',{bubbles:true,cancelable:true})); return 'primary'; }
                return null;
            }''')
            if not clicked:
                try:
                    pg.locator('button:has-text("Join")').first.click(force=True, timeout=5000)
                except:
                    pg.keyboard.press("Enter")

            pg.wait_for_timeout(15000)
            dismiss_popups()

            body_lc = pg.inner_text('body')[:2000].lower()
            if any(ph in body_lc for ph in INVALID_MEETING_PHRASES):
                emit('phase', phase='error')
                emit('log', text='dead link. this meeting is ghosting us.')
                emit('error', text='Invalid meeting link. Check the URL and try again.')
                br.close(); return

            body_check = pg.inner_text('body')[:500]
            if 'Enter Meeting Info' in body_check or 'Your Name' in body_check:
                emit('log', text='retrying join...')
                ni2 = pg.query_selector('input#inputname, input[placeholder*="name" i]')
                if ni2: ni2.fill(BOT_NAME)
                pg.evaluate('''()=>{let btns=document.querySelectorAll('button');
                    for(let b of btns){if(b.textContent.trim()==='Join'){b.click();return;}}}''')
                pg.wait_for_timeout(12000)
                dismiss_popups()

            pg.wait_for_timeout(10000)
            dismiss_popups()
            pg.wait_for_timeout(3000)
            dismiss_popups()

            body_lc2 = pg.inner_text('body')[:2000].lower()
            if any(ph in body_lc2 for ph in INVALID_MEETING_PHRASES):
                emit('phase', phase='error')
                emit('log', text="link's cooked. meeting doesn't exist.")
                emit('error', text='Invalid meeting. Check URL.')
                br.close(); return

            # Mute mic + camera
            pg.evaluate('''()=>{let btns=document.querySelectorAll('button');
                for(let b of btns){let al=(b.getAttribute('aria-label')||'').toLowerCase();
                if(al==='mute'||al.includes('mute my audio'))b.click();
                if(al==='stop video'||al.includes('stop my video'))b.click();}}''')
            pg.wait_for_timeout(500)

            # Verify in meeting
            body = pg.inner_text('body')[:600].lower()
            in_meeting = False
            if any(k in body for k in ['participant', 'mute', 'unmute', 'chat', 'leave']):
                emit('log', text="WE'RE INSIDE. ghost mode activated.")
                in_meeting = True
            elif 'waiting' in body:
                emit('log', text='chilling in the waiting room...')
                emit('phase', phase='waiting_room')
                for _ in range(120):
                    if state['stop']: break
                    pg.wait_for_timeout(5000)
                    body = pg.inner_text('body')[:300].lower()
                    if any(k in body for k in ['participant', 'mute', 'leave']):
                        emit('log', text="host said bet. WE'RE IN.")
                        in_meeting = True; break
            else:
                cur = pg.url.lower()
                if '/wc/' in cur and 'join' not in cur:
                    in_meeting = True
                    emit('log', text='slid in via URL check.')

            if not in_meeting:
                emit('phase', phase='error')
                emit('error', text="Couldn't get into the meeting. Host may need to start it.")
                br.close(); return

            emit('phase', phase='in_meeting')
            emit('log', text='door opened. walked right in.')

            # Open participants + read count
            pg.evaluate('''()=>{let btns=document.querySelectorAll('button');
                for(let b of btns){let al=(b.getAttribute('aria-label')||'').toLowerCase();
                if(al.includes('participant')){b.click();return;}}}''')
            time.sleep(2)
            pg.wait_for_timeout(3000)

            try:
                ct = pg.evaluate('''()=>{
                    let btns=document.querySelectorAll('button');
                    for(let b of btns){let al=(b.getAttribute('aria-label')||'');
                    if(al.toLowerCase().includes('participant')){
                    let m=b.textContent.match(/(\\d+)/);if(m)return parseInt(m[1]);
                    let m2=al.match(/(\\d+)/);if(m2)return parseInt(m2[1]);}}
                    let sp=document.querySelector('.panel-shell-header__title span');
                    if(sp){let m3=sp.textContent.match(/(\\d+)/);if(m3)return parseInt(m3[1]);}
                    return 0;}''')
                state['total_in_meeting'] = ct or 0
                if ct:
                    emit('log', text=f'{ct} souls detected in the matrix.')
                    emit('stats', total=ct, collected=0, ignored=0, groups=[], mixed=False)
            except: pass

            emit('phase', phase='collecting')
            emit('log', text='harvesting names... sit back and watch.')

            # COLLECTION LOOP
            poll = 0
            while not state['stop']:
                poll += 1
                dismiss_popups()

                # Open participants
                pg.evaluate('''()=>{let btns=document.querySelectorAll('button');
                    for(let b of btns){let al=(b.getAttribute('aria-label')||'').toLowerCase();
                    if(al.includes('participant')){b.click();return;}}}''')
                pg.wait_for_timeout(1500)

                # Scroll participant list
                dims = pg.evaluate('''()=>{
                    let l=document.querySelector('.ReactVirtualized__List.participants-l,.ReactVirtualized__List');
                    if(!l)return null;
                    return{sh:l.scrollHeight,ch:l.clientHeight};}''')
                if dims:
                    sh, ch = dims['sh'], dims['ch']
                    step = max(int(ch * 0.7), 100)
                    pos = 0
                    while pos <= sh:
                        pg.evaluate(f'()=>{{let l=document.querySelector(".ReactVirtualized__List");if(l)l.scrollTop={pos};}}')
                        pg.wait_for_timeout(300)
                        texts = pg.evaluate('''()=>{
                            let l=document.querySelector('.ReactVirtualized__List');if(!l)return[];
                            let o=new Set();l.querySelectorAll('span,div').forEach(e=>{
                            let t=e.textContent.trim();if(t&&t.length>2&&t.length<200)o.add(t);});
                            return Array.from(o);}''')
                        for t in texts:
                            for line in t.split('\n'):
                                add_entry(line)
                        pos += step
                    pg.evaluate('()=>{let l=document.querySelector(".ReactVirtualized__List");if(l)l.scrollTop=0;}')

                panel = pg.query_selector('#wc-container-right')
                if panel:
                    txt = panel.inner_text()
                    for line in txt.split('\n'):
                        add_entry(line)

                # Body scan
                try:
                    bd = pg.inner_text('body')
                    bd_low = bd[:500].lower()
                    if any(k in bd_low for k in ['participant', 'mute', 'unmute', 'leave']):
                        for line in bd.split('\n'):
                            line = line.strip()
                            if not line or len(line) < 4 or len(line) > 200: continue
                            if re.search(r'\b\d{7,13}\b', line):
                                add_entry(line)
                            elif re.match(r'^[A-Za-z]', line) and 4 < len(line) < 80:
                                add_entry(line)
                except: pass

                # Chat every 4th poll
                if poll % 4 == 0:
                    pg.evaluate('''()=>{let btns=document.querySelectorAll('button');
                        for(let b of btns){let al=(b.getAttribute('aria-label')||'').toLowerCase();
                        if(al.includes('chat')&&!al.includes('setting')){b.click();return;}}}''')
                    time.sleep(2)
                    cp = pg.query_selector('#wc-container-right')
                    if cp:
                        for line in cp.inner_text().split('\n'):
                            add_entry(line)
                    pg.evaluate('''()=>{let btns=document.querySelectorAll('button');
                        for(let b of btns){let al=(b.getAttribute('aria-label')||'').toLowerCase();
                        if(al.includes('participant')){b.click();return;}}}''')
                    pg.wait_for_timeout(1500)

                # Re-read attendee count
                try:
                    ct2 = pg.evaluate('''()=>{let btns=document.querySelectorAll('button');
                        for(let b of btns){let al=(b.getAttribute('aria-label')||'');
                        if(al.toLowerCase().includes('participant')){
                        let m=b.textContent.match(/(\\d+)/);if(m)return parseInt(m[1]);
                        let m2=al.match(/(\\d+)/);if(m2)return parseInt(m2[1]);}}return 0;}''')
                    if ct2: state['total_in_meeting'] = ct2
                except: pass

                with lock:
                    _rebuild_lists()

                if poll % 3 == 0:
                    with lock:
                        emit('log', text=f"scan #{poll} | caught {len(state['collected'])} | ignored {len(state['ignored'])}")

                try:
                    blo = pg.inner_text('body').lower()[:500]
                    if 'meeting has ended' in blo or 'this meeting has been ended' in blo:
                        emit('log', text='meeting wrapped. gg everyone.')
                        emit('phase', phase='done')
                        break
                except: pass

                time.sleep(15)

            emit('phase', phase='done')
            emit('log', text='collection finished.')
            try: br.close()
            except: pass

    except Exception as e:
        emit('phase', phase='error')
        emit('error', text=str(e))
        emit('log', text=f'error: {e}')


# ── FLASK APP ────────────────────────────────────────────────────────────────
app = Flask(__name__)

TEMPLATE = r'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Attendance Companion</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#0b0d13;--bg2:#111318;--card:#171a23;--card-hover:#1c2029;
  --border:#252832;--border-light:#2f3340;
  --text:#e2e5eb;--text2:#9ba1ae;--text3:#5e6373;
  --accent:#4f8ff7;--accent2:#3b7be0;
  --green:#2dd4a0;--green-bg:rgba(45,212,160,.08);
  --red:#f87171;--red-bg:rgba(248,113,113,.08);
  --amber:#fbbf24;--amber-bg:rgba(251,191,36,.08);
  --blue-bg:rgba(79,143,247,.06);
  --radius:10px;--radius-sm:6px;
}
html{font-size:14px}
body{font-family:'Inter',system-ui,-apple-system,sans-serif;background:var(--bg);color:var(--text);min-height:100vh;overflow-x:hidden}

/* LANDING */
#landing{display:flex;align-items:center;justify-content:center;min-height:100vh;padding:24px;background:var(--bg)}
.landing-wrap{width:100%;max-width:440px}
.landing-card{background:var(--card);border:1px solid var(--border);border-radius:16px;padding:48px 40px;text-align:center}
.landing-card h1{font-size:1.6rem;font-weight:700;letter-spacing:-.02em;margin-bottom:4px}
.landing-card .sub{color:var(--text2);font-size:.85rem;margin-bottom:32px}
.landing-card label{display:block;text-align:left;font-size:.75rem;font-weight:600;color:var(--text2);text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px;margin-top:20px}
.landing-card input{width:100%;padding:12px 14px;background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius-sm);color:var(--text);font-size:.9rem;outline:none;transition:border .2s}
.landing-card input:focus{border-color:var(--accent)}
.landing-card input::placeholder{color:var(--text3)}
.btn-primary{width:100%;padding:14px;margin-top:28px;background:var(--accent);color:#fff;border:none;border-radius:var(--radius-sm);font-size:.95rem;font-weight:600;cursor:pointer;transition:background .2s}
.btn-primary:hover{background:var(--accent2)}
.btn-primary:disabled{opacity:.5;cursor:not-allowed}

/* DASHBOARD */
#dashboard{display:none;min-height:100vh;padding:0}
.topbar{display:flex;align-items:center;justify-content:space-between;padding:16px 28px;background:var(--bg2);border-bottom:1px solid var(--border);position:sticky;top:0;z-index:100}
.topbar-left{display:flex;align-items:center;gap:12px}
.topbar-left h1{font-size:1.05rem;font-weight:700;letter-spacing:-.01em}
.status-pill{display:inline-flex;align-items:center;gap:6px;padding:4px 12px;border-radius:20px;font-size:.72rem;font-weight:600;text-transform:uppercase;letter-spacing:.04em}
.status-pill.connecting{background:var(--amber-bg);color:var(--amber)}
.status-pill.live{background:var(--green-bg);color:var(--green)}
.status-pill.error{background:var(--red-bg);color:var(--red)}
.status-pill.done{background:var(--blue-bg);color:var(--accent)}
.dot{width:7px;height:7px;border-radius:50%;background:currentColor}
.dot.pulse{animation:pulse 1.5s infinite}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.3}}
.topbar-right{display:flex;align-items:center;gap:16px}
.timer{font-size:.85rem;font-weight:600;color:var(--text2);font-variant-numeric:tabular-nums}

.dash-body{padding:24px 28px;max-width:1400px;margin:0 auto}

/* STEPS */
.steps-panel{background:var(--card);border:1px solid var(--border);border-radius:var(--radius);padding:32px;margin-bottom:24px}
.steps-list{display:flex;flex-direction:column;gap:14px}
.step-row{display:flex;align-items:center;gap:14px;font-size:.9rem}
.step-icon{width:28px;height:28px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:.7rem;flex-shrink:0;transition:all .3s}
.step-icon.pending{border:2px solid var(--border-light);color:var(--text3)}
.step-icon.active{background:var(--accent);color:#fff;animation:pulse 1.5s infinite}
.step-icon.done{background:var(--green);color:#0b0d13}
.step-icon.error{background:var(--red);color:#fff}
.step-text{color:var(--text2)}.step-text.active{color:var(--text);font-weight:500}

/* STATS */
.stats-row{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:24px}
.stat-card{background:var(--card);border:1px solid var(--border);border-radius:var(--radius);padding:20px 22px;transition:border .2s}
.stat-card:hover{border-color:var(--border-light)}
.stat-val{font-size:2rem;font-weight:700;line-height:1;margin-bottom:4px;font-variant-numeric:tabular-nums}
.stat-label{font-size:.72rem;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:.06em}
.stat-val.green{color:var(--green)}.stat-val.red{color:var(--red)}.stat-val.amber{color:var(--amber)}.stat-val.blue{color:var(--accent)}

/* TABLES */
.tables-row{display:grid;grid-template-columns:2fr 1fr;gap:16px;margin-bottom:24px}
.table-card{background:var(--card);border:1px solid var(--border);border-radius:var(--radius);overflow:hidden;display:flex;flex-direction:column;max-height:460px}
.table-header{display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid var(--border)}
.table-header h2{font-size:.9rem;font-weight:600}
.table-header .count{background:var(--bg);padding:2px 10px;border-radius:12px;font-size:.75rem;font-weight:600;color:var(--text2)}
.table-search{padding:8px 20px;border-bottom:1px solid var(--border)}
.table-search input{width:100%;padding:8px 10px;background:var(--bg);border:1px solid var(--border);border-radius:var(--radius-sm);color:var(--text);font-size:.8rem;outline:none}
.table-search input:focus{border-color:var(--accent)}
.table-scroll{overflow-y:auto;flex:1}
table{width:100%;border-collapse:collapse}
thead{position:sticky;top:0;background:var(--card);z-index:1}
th{padding:10px 16px;font-size:.7rem;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;text-align:left;border-bottom:1px solid var(--border)}
td{padding:10px 16px;font-size:.85rem;border-bottom:1px solid var(--border)}
tr:hover td{background:var(--card-hover)}
.row-new{animation:rowIn .4s ease}
@keyframes rowIn{from{opacity:0;transform:translateX(-6px)}to{opacity:1;transform:translateX(0)}}
.badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:.7rem;font-weight:600}
.badge-green{background:var(--green-bg);color:var(--green)}
.badge-red{background:var(--red-bg);color:var(--red)}

/* ACTIVITY LOG */
.log-card{background:var(--card);border:1px solid var(--border);border-radius:var(--radius);overflow:hidden;margin-bottom:24px}
.log-card .table-header h2{font-size:.9rem}
.log-scroll{max-height:180px;overflow-y:auto;padding:12px 20px;font-family:'SF Mono','Fira Code',monospace;font-size:.78rem;line-height:1.7;color:var(--text2)}
.log-line{opacity:0;animation:logIn .3s ease forwards}
@keyframes logIn{to{opacity:1}}
.log-line .ts{color:var(--text3);margin-right:8px}

/* ACTIONS BAR */
.actions-bar{display:flex;gap:12px;justify-content:flex-end;padding:16px 0}
.btn-stop{padding:10px 24px;background:transparent;border:1px solid var(--red);color:var(--red);border-radius:var(--radius-sm);font-size:.85rem;font-weight:600;cursor:pointer;transition:all .2s}
.btn-stop:hover{background:var(--red-bg)}
.btn-save{padding:10px 28px;background:var(--green);color:#0b0d13;border:none;border-radius:var(--radius-sm);font-size:.85rem;font-weight:700;cursor:pointer;transition:all .2s;position:relative;overflow:hidden;z-index:0}
.btn-save:hover{background:#25bf8e}
.btn-save:disabled{opacity:.5;cursor:not-allowed}
.btn-primary{position:relative;overflow:hidden;z-index:0}
.btn-stop{position:relative;overflow:hidden;z-index:0}
.btn-loading{pointer-events:none;border-color:transparent!important}
.btn-loading::before{content:'';position:absolute;top:50%;left:50%;width:calc(100% + 6px);height:calc(100% + 6px);margin-top:calc(-50% - 3px);margin-left:calc(-50% - 3px);border-radius:inherit;background:conic-gradient(transparent 55%,var(--glow-color,#1dd882) 75%,transparent 100%);animation:border-spin 1.2s linear infinite;z-index:-2}
.btn-loading::after{content:'';position:absolute;inset:2px;border-radius:inherit;z-index:-1}
.btn-save.btn-loading::after{background:var(--green)}
.btn-stop.btn-loading::after{background:var(--bg)}
.btn-primary.btn-loading::after{background:var(--accent)}
.email-actions .btn-send.btn-loading::after{background:var(--accent)}
.btn-email.btn-loading::after{background:var(--bg)}
@keyframes border-spin{to{transform:rotate(360deg)}}
.btn-email{padding:10px 24px;background:transparent;border:1px solid var(--accent);color:var(--accent);border-radius:var(--radius-sm);font-size:.85rem;font-weight:600;cursor:pointer;transition:all .2s;position:relative;overflow:hidden;z-index:0}
.btn-email:hover{background:var(--blue-bg)}.btn-email:disabled{opacity:.5;cursor:not-allowed}

#email-modal{display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.7);z-index:9000;align-items:center;justify-content:center}
#email-modal.active{display:flex}
.email-card{background:var(--card);border:1px solid var(--border);border-radius:16px;padding:36px 32px;width:100%;max-width:420px;text-align:center}
.email-card h2{font-size:1.15rem;font-weight:700;margin-bottom:4px}
.email-card .esub{color:var(--text2);font-size:.8rem;margin-bottom:22px}
.email-card label{display:block;text-align:left;font-size:.72rem;font-weight:600;color:var(--text2);text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px}
.email-card input{width:100%;padding:11px 14px;background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius-sm);color:var(--text);font-size:.88rem;outline:none;transition:border .2s;margin-bottom:16px}
.email-card input:focus{border-color:var(--accent)}
.email-actions{display:flex;gap:10px;justify-content:flex-end}
.email-actions .btn-cancel{padding:9px 20px;background:transparent;border:1px solid var(--border);color:var(--text2);border-radius:var(--radius-sm);font-size:.85rem;cursor:pointer}
.email-actions .btn-send{padding:9px 24px;background:var(--accent);color:#fff;border:none;border-radius:var(--radius-sm);font-size:.85rem;font-weight:600;cursor:pointer;position:relative;overflow:hidden;z-index:0}
.email-actions .btn-send:disabled{opacity:.5;cursor:not-allowed}

/* SCAN OVERLAY */
#scan-overlay{position:fixed;top:0;left:0;right:0;bottom:0;background:#000;z-index:10000;display:none;align-items:center;justify-content:center;flex-direction:column;overflow:hidden}
#scan-overlay.active{display:flex;animation:scanBgFade 3.5s ease forwards}
@keyframes scanBgFade{0%{opacity:1}80%{opacity:1}100%{opacity:0}}
.scan-line{position:absolute;left:0;right:0;height:3px;background:linear-gradient(90deg,transparent 0%,#2dd4a0 20%,#4fffcf 50%,#2dd4a0 80%,transparent 100%);box-shadow:0 0 30px 10px rgba(45,212,160,.4),0 0 80px 20px rgba(45,212,160,.15);animation:sweepDown 2s ease-in-out forwards;top:0}
@keyframes sweepDown{0%{top:-2%;opacity:0}5%{opacity:1}95%{opacity:1}100%{top:102%;opacity:0}}
.scan-line2{position:absolute;left:0;right:0;height:2px;background:linear-gradient(90deg,transparent,#2dd4a0,transparent);box-shadow:0 0 15px rgba(45,212,160,.3);animation:sweepDown 2s .3s ease-in-out forwards;top:0;opacity:0}
.scan-grid{position:absolute;top:0;left:0;right:0;bottom:0;background-image:linear-gradient(rgba(45,212,160,.03) 1px,transparent 1px),linear-gradient(90deg,rgba(45,212,160,.03) 1px,transparent 1px);background-size:40px 40px;animation:gridFade 3s ease forwards}
@keyframes gridFade{0%{opacity:0}20%{opacity:1}80%{opacity:1}100%{opacity:0}}
.scan-corners{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);width:320px;height:180px;opacity:0;animation:cornersIn 2s .2s ease forwards}
@keyframes cornersIn{0%{opacity:0;width:400px;height:220px}30%{opacity:1}70%{opacity:1}100%{opacity:0;width:280px;height:160px}}
.scan-corners::before,.scan-corners::after{content:'';position:absolute;border:2px solid rgba(45,212,160,.5)}
.scan-corners::before{top:0;left:0;width:40px;height:40px;border-right:none;border-bottom:none}
.scan-corners::after{bottom:0;right:0;width:40px;height:40px;border-left:none;border-top:none}
.scan-corner-br,.scan-corner-tl{position:absolute;border:2px solid rgba(45,212,160,.5)}
.scan-corner-tl{top:0;right:0;width:40px;height:40px;border-left:none;border-bottom:none}
.scan-corner-br{bottom:0;left:0;width:40px;height:40px;border-right:none;border-top:none}
.scan-text{color:#2dd4a0;font-size:1.3rem;font-weight:700;letter-spacing:.25em;text-transform:uppercase;text-shadow:0 0 30px rgba(45,212,160,.6),0 0 60px rgba(45,212,160,.2);opacity:0;animation:textIn 2.5s .5s ease forwards;z-index:1}
@keyframes textIn{0%{opacity:0;letter-spacing:.4em}20%{opacity:1;letter-spacing:.25em}80%{opacity:1}100%{opacity:0}}
.scan-sub{color:rgba(45,212,160,.6);font-size:.7rem;font-weight:600;letter-spacing:.15em;margin-top:8px;opacity:0;animation:textIn 2s 1s ease forwards;z-index:1}
.scan-bars{position:absolute;bottom:60px;display:flex;gap:3px;align-items:flex-end;height:40px;opacity:0;animation:barsIn 2s .3s ease forwards}
@keyframes barsIn{0%{opacity:0}20%{opacity:.6}80%{opacity:.6}100%{opacity:0}}
.scan-bar{width:4px;background:#2dd4a0;border-radius:1px;animation:barPulse .6s ease infinite alternate}
@keyframes barPulse{from{opacity:.3}to{opacity:1}}

/* TOAST */
.toast{position:fixed;bottom:24px;right:24px;padding:14px 20px;border-radius:var(--radius-sm);font-size:.85rem;font-weight:500;color:#fff;z-index:9999;animation:toastIn .3s ease,toastOut .3s ease 2.7s forwards}
.toast.success{background:#1a7f5a}.toast.error{background:#b91c1c}
@keyframes toastIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
@keyframes toastOut{to{opacity:0;transform:translateY(10px)}}

/* RESPONSIVE */
@media(max-width:900px){.tables-row{grid-template-columns:1fr}.stats-row{grid-template-columns:repeat(2,1fr)}}
@media(max-width:600px){.stats-row{grid-template-columns:1fr}}

::-webkit-scrollbar{width:6px}::-webkit-scrollbar-track{background:transparent}::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px}
</style>
</head>
<body>

<!-- SCAN OVERLAY -->
<div id="scan-overlay">
  <div class="scan-grid"></div>
  <div class="scan-line"></div>
  <div class="scan-line2"></div>
  <div class="scan-corners">
    <div class="scan-corner-tl"></div>
    <div class="scan-corner-br"></div>
  </div>
  <div class="scan-text" id="scan-text">SCANNING</div>
  <div class="scan-sub">ghost mode activated</div>
  <div class="scan-bars" id="scan-bars"></div>
</div>

<!-- LANDING -->
<div id="landing">
<div class="landing-wrap">
  <div class="landing-card">
    <h1>Attendance Companion</h1>
    <p class="sub">Drop the link, we handle the rest</p>
    <label>Meeting URL</label>
    <input id="inp-url" placeholder="https://us05web.zoom.us/j/..." autofocus>
    <div style="display:flex;gap:12px;margin-top:20px">
      <div style="flex:1;text-align:left">
        <label style="margin-top:0">Output Filename</label>
        <input id="inp-file" placeholder="e.g. ITPM_Attendance">
      </div>
      <div style="width:72px;text-align:left;flex-shrink:0">
        <label style="margin-top:0">Group</label>
        <input id="inp-group" placeholder="A" maxlength="1" style="text-transform:uppercase;text-align:center">
      </div>
    </div>
    <p style="font-size:.72rem;color:var(--text3);margin-top:6px;text-align:left">If your class shares a meeting with other groups, only your group is collected. Default: A.</p>

    <!-- ROSTER TOGGLE -->
    <div style="display:flex;align-items:center;gap:10px;margin-top:20px;text-align:left">
      <label style="margin:0;cursor:pointer;flex-shrink:0" for="roster-toggle">Use Class Roster</label>
      <label style="position:relative;width:40px;height:22px;margin:0;flex-shrink:0">
        <input type="checkbox" id="roster-toggle" onchange="toggleRoster()" style="opacity:0;width:0;height:0">
        <span style="position:absolute;inset:0;background:var(--border);border-radius:11px;cursor:pointer;transition:.2s"></span>
        <span id="roster-knob" style="position:absolute;top:3px;left:3px;width:16px;height:16px;background:#fff;border-radius:50%;transition:.2s"></span>
      </label>
      <span style="font-size:.72rem;color:var(--text3)">Match participants against an uploaded name list</span>
    </div>
    <div id="roster-box" style="display:none;margin-top:12px;text-align:left">
      <p style="font-size:.72rem;color:var(--text3);margin-bottom:8px">Upload a file with student names (one per line). Accepted: .txt, .json, .md, .xlsx</p>
      <input type="file" id="roster-file" accept=".txt,.json,.md,.xlsx" onchange="uploadRoster()" style="font-size:.82rem;color:var(--text2)">
      <div id="roster-status" style="margin-top:8px;font-size:.8rem;display:none"></div>
      <div id="roster-preview" style="margin-top:8px;max-height:150px;overflow-y:auto;font-size:.78rem;color:var(--text2);display:none;background:var(--bg2);border-radius:8px;padding:8px 12px"></div>
    </div>

    <button class="btn-primary" id="btn-start" onclick="startTracking()">Start Tracking</button>
  </div>
  <p style="text-align:center;color:var(--text3);font-size:.72rem;margin-top:16px;letter-spacing:.02em">Created by Osman Mohammed (TheNooB)</p>
</div>
</div>

<!-- DASHBOARD -->
<div id="dashboard">

<div class="topbar">
  <div class="topbar-left">
    <h1>Attendance Harvester</h1>
    <div class="status-pill connecting" id="status-pill"><span class="dot pulse"></span><span id="status-text">Connecting</span></div>
  </div>
  <div class="topbar-right">
    <span class="timer" id="timer">00:00:00</span>
  </div>
</div>

<div class="dash-body">

<!-- STEPS -->
<div class="steps-panel" id="steps-panel">
  <div class="steps-list">
    <div class="step-row" id="step-launch"><div class="step-icon pending" id="si-launch">1</div><span class="step-text">Booting shadow client</span></div>
    <div class="step-row" id="step-login"><div class="step-icon pending" id="si-login">2</div><span class="step-text">Slipping through the front door</span></div>
    <div class="step-row" id="step-join"><div class="step-icon pending" id="si-join">3</div><span class="step-text">Pulling up to the lecture</span></div>
    <div class="step-row" id="step-collect"><div class="step-icon pending" id="si-collect">4</div><span class="step-text">Harvesting names</span></div>
  </div>
</div>

<!-- STATS -->
<div class="stats-row">
  <div class="stat-card"><div class="stat-val blue" id="sv-total">0</div><div class="stat-label">In Meeting</div></div>
  <div class="stat-card"><div class="stat-val green" id="sv-collected">0</div><div class="stat-label">Collected</div></div>
  <div class="stat-card"><div class="stat-val red" id="sv-ignored">0</div><div class="stat-label">Ignored</div></div>
  <div class="stat-card"><div class="stat-val amber" id="sv-groups">-</div><div class="stat-label">Groups</div></div>
</div>

<!-- TABLES -->
<div class="tables-row">
  <div class="table-card">
    <div class="table-header"><h2>Collected</h2><span class="count" id="tc-collected">0</span></div>
    <div class="table-search"><input id="search-collected" placeholder="Search names..."></div>
    <div class="table-scroll">
      <table><thead><tr><th style="width:50px">#</th><th>Name</th><th>Student ID</th></tr></thead>
      <tbody id="tbody-collected"></tbody></table>
    </div>
  </div>
  <div class="table-card">
    <div class="table-header"><h2>Ignored</h2><span class="count" id="tc-ignored">0</span></div>
    <div class="table-search"><input id="search-ignored" placeholder="Search ignored..."></div>
    <div class="table-scroll">
      <table><thead><tr><th>Name</th><th>Reason</th></tr></thead>
      <tbody id="tbody-ignored"></tbody></table>
    </div>
  </div>
</div>

<!-- ABSENT (roster mode only) -->
<div class="tables-grid" id="absent-section" style="display:none">
  <div class="table-card" style="grid-column:1/-1">
    <div class="table-header"><h2>Absent <span id="absent-count" style="color:var(--red);font-weight:400"></span></h2></div>
    <div class="table-search"><input id="search-absent" placeholder="Search absent..."></div>
    <div class="table-scroll">
      <table><thead><tr><th>Name</th><th>Student ID</th></tr></thead>
      <tbody id="tbody-absent"></tbody></table>
    </div>
  </div>
</div>

<!-- LOG -->
<div class="log-card">
  <div class="table-header"><h2>Activity</h2></div>
  <div class="log-scroll" id="log-scroll"></div>
</div>

<!-- ACTIONS -->
<div class="actions-bar">
  <button class="btn-stop" id="btn-stop" onclick="stopCollection()">Stop Collection</button>
  <button class="btn-save" id="btn-save" onclick="saveExport()" disabled>Save &amp; Export</button>
  <button class="btn-stop" id="btn-csv" onclick="downloadCSV()" style="display:none;border-color:var(--accent);color:var(--accent)">Download CSV</button>
  <button class="btn-email" id="btn-email" onclick="openEmailModal()" disabled style="display:none">Email Attendance</button>
</div>

</div></div>

<!-- EMAIL MODAL -->
<div id="email-modal">
  <div class="email-card">
    <h2>Send Attendance via Email</h2>
    <p class="esub">The saved attendance file will be sent as an attachment.</p>
    <label>Recipient Email</label>
    <input id="email-to" type="email" placeholder="lecturer@university.edu">
    <div class="email-actions">
      <button class="btn-cancel" onclick="closeEmailModal()">Cancel</button>
      <button class="btn-send" id="btn-send-email" onclick="sendEmail()">Send</button>
    </div>
  </div>
</div>

<script>
let timerInterval, startTime;
let collectedRows = [], ignoredRows = [], absentRows = [];
let rosterMode = false;

function toggleRoster() {
  const cb = document.getElementById('roster-toggle');
  const box = document.getElementById('roster-box');
  const knob = document.getElementById('roster-knob');
  const grpInput = document.getElementById('inp-group');
  if (cb.checked) {
    box.style.display = 'block';
    knob.style.left = '20px'; knob.style.background = 'var(--green)';
    if (grpInput) { grpInput.disabled = true; grpInput.style.opacity = '0.4'; }
  } else {
    box.style.display = 'none';
    knob.style.left = '2px'; knob.style.background = 'var(--text3)';
    if (grpInput) { grpInput.disabled = false; grpInput.style.opacity = '1'; }
    document.getElementById('roster-status').textContent = '';
    document.getElementById('roster-preview').innerHTML = '';
    fetch('/clear-roster', {method:'POST'});
    rosterMode = false;
  }
}

function uploadRoster() {
  const inp = document.getElementById('roster-file');
  const statusEl = document.getElementById('roster-status');
  const previewEl = document.getElementById('roster-preview');
  if (!inp.files || !inp.files[0]) return;
  statusEl.textContent = 'Uploading...';
  statusEl.style.color = 'var(--text2)';
  previewEl.innerHTML = '';
  const fd = new FormData();
  fd.append('file', inp.files[0]);
  fetch('/upload-roster', {method:'POST', body: fd})
  .then(r => r.json()).then(d => {
    if (d.ok) {
      rosterMode = true;
      statusEl.textContent = d.count + ' students loaded';
      statusEl.style.color = 'var(--green)';
      let html = '<div style="max-height:120px;overflow-y:auto;margin-top:6px;font-size:.78rem;color:var(--text2)">';
      (d.names || []).forEach((n, i) => {
        html += '<div style="padding:2px 0;border-bottom:1px solid var(--border)">' +
          (i+1) + '. ' + esc(n.name) + (n.student_id ? ' <span style="color:var(--accent)">[' + esc(n.student_id) + ']</span>' : '') + '</div>';
      });
      html += '</div>';
      previewEl.innerHTML = html;
    } else {
      statusEl.textContent = d.error || 'Upload failed';
      statusEl.style.color = 'var(--red)';
      rosterMode = false;
    }
  }).catch(() => { statusEl.textContent = 'Network error'; statusEl.style.color = 'var(--red)'; });
}

function startTracking() {
  const url = document.getElementById('inp-url').value.trim();
  const fname = document.getElementById('inp-file').value.trim();
  const grp = document.getElementById('inp-group').value.trim().toUpperCase() || 'A';
  if (!url) { document.getElementById('inp-url').style.borderColor='var(--red)'; return; }
  const btn = document.getElementById('btn-start');
  btn.disabled = true;
  btn.textContent = 'Launching...';
  btn.classList.add('btn-loading');
  btn.style.setProperty('--glow-color','var(--accent)');

  fetch('/start', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({url: url, filename: fname, group: grp})
  }).then(r => r.json()).then(d => {
    btn.classList.remove('btn-loading');
    if (d.ok) {
      document.getElementById('landing').style.display = 'none';
      document.getElementById('dashboard').style.display = 'block';
      startTimer();
      listenSSE();
    }
  });
}

function startTimer() {
  startTime = Date.now();
  timerInterval = setInterval(() => {
    const e = Date.now() - startTime;
    const h = String(Math.floor(e/3600000)).padStart(2,'0');
    const m = String(Math.floor((e%3600000)/60000)).padStart(2,'0');
    const s = String(Math.floor((e%60000)/1000)).padStart(2,'0');
    document.getElementById('timer').textContent = h+':'+m+':'+s;
  }, 1000);
}

function listenSSE() {
  const es = new EventSource('/stream');
  es.onmessage = function(e) {
    const d = JSON.parse(e.data);
    handleEvent(d);
  };
  es.onerror = function() { setTimeout(() => listenSSE(), 2000); };
}

const stepMap = {
  launching: ['launch','active',['login','pending'],['join','pending'],['collect','pending']],
  logging_in: ['launch','done',['login','active'],['join','pending'],['collect','pending']],
  logged_in: ['launch','done',['login','done'],['join','pending'],['collect','pending']],
  joining: ['launch','done',['login','done'],['join','active'],['collect','pending']],
  waiting_room: ['launch','done',['login','done'],['join','active'],['collect','pending']],
  in_meeting: ['launch','done',['login','done'],['join','done'],['collect','pending']],
  collecting: ['launch','done',['login','done'],['join','done'],['collect','active']],
  done: ['launch','done',['login','done'],['join','done'],['collect','done']],
  error: ['launch','done',['login','done'],['join','error'],['collect','pending']],
};

function setStep(id, cls) {
  const el = document.getElementById('si-'+id);
  el.className = 'step-icon ' + cls;
  el.textContent = cls==='done' ? '✓' : cls==='error' ? '✗' : el.textContent.replace(/[^0-9]/g,'') || '';
  const row = document.getElementById('step-'+id);
  row.querySelector('.step-text').className = 'step-text' + (cls==='active' ? ' active' : '');
}

function handleEvent(d) {
  if (d.type === 'phase') {
    const p = d.phase;
    const pill = document.getElementById('status-pill');
    const ptxt = document.getElementById('status-text');
    if (p==='in_meeting') { triggerScan(); pill.className='status-pill live'; ptxt.textContent='Live'; }
    else if (p==='collecting') { pill.className='status-pill live'; ptxt.textContent='Live'; }
    else if (p==='done') { pill.className='status-pill done'; ptxt.textContent='Done'; clearInterval(timerInterval); document.getElementById('btn-save').disabled=false; document.getElementById('btn-stop').style.display='none'; }
    else if (p==='error') { pill.className='status-pill error'; ptxt.textContent='Error'; clearInterval(timerInterval); document.getElementById('btn-save').disabled=false; }
    else if (p==='waiting_room') { pill.className='status-pill connecting'; ptxt.textContent='Waiting Room'; }
    else { pill.className='status-pill connecting'; ptxt.textContent='Connecting'; }

    const steps = ['launch','login','join','collect'];
    const map = {launching:['active','pending','pending','pending'],logging_in:['done','active','pending','pending'],logged_in:['done','done','pending','pending'],joining:['done','done','active','pending'],waiting_room:['done','done','active','pending'],in_meeting:['done','done','done','pending'],collecting:['done','done','done','active'],done:['done','done','done','done'],error:['done','done','error','pending']};
    const st = map[p] || map.launching;
    steps.forEach((s,i) => setStep(s, st[i]));
    if (p==='collecting'||p==='done') document.getElementById('steps-panel').style.display='none';
  }

  if (d.type === 'stats') {
    document.getElementById('sv-total').textContent = d.total || 0;
    document.getElementById('sv-collected').textContent = d.collected || 0;
    document.getElementById('sv-ignored').textContent = d.ignored || 0;
    document.getElementById('sv-groups').textContent = d.groups && d.groups.length ? d.groups.join(', ') : '-';
    document.getElementById('tc-collected').textContent = d.collected || 0;
    document.getElementById('tc-ignored').textContent = d.ignored || 0;
    if (typeof d.absent === 'number' && d.absent > 0) {
      document.getElementById('absent-count').textContent = '(' + d.absent + ')';
    }
  }

  if (d.type === 'name') {
    if (d.status === 'collected') {
      const exists = collectedRows.find(r => r.name === d.name);
      if (!exists) {
        collectedRows.push(d);
        const tbody = document.getElementById('tbody-collected');
        const tr = document.createElement('tr');
        tr.className = 'row-new';
        tr.innerHTML = `<td>${collectedRows.length}</td><td>${esc(d.name)}</td><td><span class="badge badge-green">${d.student_id||'N/A'}</span></td>`;
        tbody.appendChild(tr);
      }
    } else {
      const exists = ignoredRows.find(r => r.name === d.name);
      if (!exists) {
        ignoredRows.push(d);
        const tbody = document.getElementById('tbody-ignored');
        const tr = document.createElement('tr');
        tr.className = 'row-new';
        tr.innerHTML = `<td>${esc(d.name)}</td><td><span class="badge badge-red">${esc(d.reason)}</span></td>`;
        tbody.appendChild(tr);
      }
    }
  }

  if (d.type === 'rebuild') {
    rebuildTables(d.collected || [], d.ignored || [], d.absent || []);
  }

  if (d.type === 'log') {
    const log = document.getElementById('log-scroll');
    const now = new Date().toLocaleTimeString('en-US',{hour12:false});
    const div = document.createElement('div');
    div.className = 'log-line';
    div.innerHTML = `<span class="ts">${now}</span>${esc(d.text)}`;
    log.appendChild(div);
    log.scrollTop = log.scrollHeight;
  }

  if (d.type === 'error') {
    showToast(d.text, 'error');
  }
}

function esc(s) { const d=document.createElement('div'); d.textContent=s; return d.innerHTML; }

function stopCollection() {
  const btn=document.getElementById('btn-stop');
  btn.classList.add('btn-loading');btn.style.setProperty('--glow-color','var(--red)');
  btn.textContent='Stopping...';
  fetch('/stop', {method:'POST'}).then(r=>r.json()).then(d=>{
    btn.classList.remove('btn-loading');
    if(d.ok) showToast('Collection stopped','success');
    btn.textContent='Stop Collection';
  });
}

function saveExport() {
  const btn = document.getElementById('btn-save');
  btn.disabled = true; btn.textContent = 'Saving...';
  btn.classList.add('btn-loading');btn.style.setProperty('--glow-color','var(--green)');
  fetch('/save',{method:'POST'}).then(r=>r.json()).then(d=>{
    btn.classList.remove('btn-loading');
    if(d.ok){
      showToast('Saved to '+d.path,'success');
      window.location.href='/download';
      document.getElementById('btn-csv').style.display='inline-block';
      setTimeout(()=>{btn.textContent='Saved!';},500);
      setTimeout(()=>{btn.textContent='Save & Export';btn.disabled=false;},3000);
      fetch('/smtp-status').then(r=>r.json()).then(s=>{
        if(s.configured){
          const eb=document.getElementById('btn-email');
          eb.style.display='inline-block';eb.disabled=false;
        }
      });
    } else {
      showToast(d.error||'Save failed','error');
      btn.textContent='Save & Export';btn.disabled=false;
    }
  });
}

function downloadCSV(){window.location.href='/download-csv';}

function openEmailModal(){document.getElementById('email-modal').classList.add('active');document.getElementById('email-to').focus();}
function closeEmailModal(){document.getElementById('email-modal').classList.remove('active');}

function sendEmail(){
  const to=document.getElementById('email-to').value.trim();
  if(!to||!to.includes('@')){document.getElementById('email-to').style.borderColor='var(--red)';return;}
  const btn=document.getElementById('btn-send-email');
  btn.disabled=true;btn.textContent='Sending...';
  btn.classList.add('btn-loading');btn.style.setProperty('--glow-color','var(--accent)');
  fetch('/email',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({to:to})})
  .then(r=>r.json()).then(d=>{
    btn.classList.remove('btn-loading');
    if(d.ok){
      showToast('Email sent to '+to,'success');
      closeEmailModal();
      btn.textContent='Send';btn.disabled=false;
    } else {
      showToast(d.error||'Send failed','error');
      btn.textContent='Send';btn.disabled=false;
    }
  }).catch(()=>{btn.classList.remove('btn-loading');showToast('Network error','error');btn.textContent='Send';btn.disabled=false;});
}

document.getElementById('email-to').addEventListener('keydown',function(e){if(e.key==='Enter')sendEmail();});
document.getElementById('email-modal').addEventListener('click',function(e){if(e.target===this)closeEmailModal();});

function showToast(msg, type) {
  const t = document.createElement('div');
  t.className = 'toast ' + type;
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(()=>t.remove(), 3000);
}

document.getElementById('search-collected').addEventListener('input', function(e){
  const q = e.target.value.toLowerCase();
  document.querySelectorAll('#tbody-collected tr').forEach(r=>{
    r.style.display = r.textContent.toLowerCase().includes(q) ? '' : 'none';
  });
});
document.getElementById('search-ignored').addEventListener('input', function(e){
  const q = e.target.value.toLowerCase();
  document.querySelectorAll('#tbody-ignored tr').forEach(r=>{
    r.style.display = r.textContent.toLowerCase().includes(q) ? '' : 'none';
  });
});
document.getElementById('search-absent').addEventListener('input', function(e){
  const q = e.target.value.toLowerCase();
  document.querySelectorAll('#tbody-absent tr').forEach(r=>{
    r.style.display = r.textContent.toLowerCase().includes(q) ? '' : 'none';
  });
});

function rebuildTables(coll, ign, abs) {
  collectedRows = coll;
  ignoredRows = ign;
  absentRows = abs || [];
  const tc = document.getElementById('tbody-collected');
  tc.innerHTML = '';
  coll.forEach((d, i) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${i+1}</td><td>${esc(d.name)}</td><td><span class="badge badge-green">${d.student_id||'N/A'}</span></td>`;
    tc.appendChild(tr);
  });
  const ti = document.getElementById('tbody-ignored');
  ti.innerHTML = '';
  ign.forEach(d => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${esc(d.name)}</td><td><span class="badge badge-red">${esc(d.reason||'')}</span></td>`;
    ti.appendChild(tr);
  });
  document.getElementById('tc-collected').textContent = coll.length;
  document.getElementById('tc-ignored').textContent = ign.length;
  const abSec = document.getElementById('absent-section');
  const ta = document.getElementById('tbody-absent');
  if (absentRows.length > 0) {
    abSec.style.display = '';
    ta.innerHTML = '';
    absentRows.forEach(d => {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${esc(d.name)}</td><td>${d.student_id ? esc(d.student_id) : '-'}</td>`;
      ta.appendChild(tr);
    });
    document.getElementById('absent-count').textContent = '(' + absentRows.length + ')';
  } else {
    abSec.style.display = 'none';
    ta.innerHTML = '';
    document.getElementById('absent-count').textContent = '';
  }
}

// Restore state on page load / refresh
fetch('/state').then(r=>r.json()).then(d=>{
  if (d.roster_mode) {
    rosterMode = true;
    const cb = document.getElementById('roster-toggle');
    if (cb) { cb.checked = true; toggleRoster(); }
    const statusEl = document.getElementById('roster-status');
    if (statusEl && d.roster_count) { statusEl.textContent = d.roster_count + ' students loaded'; statusEl.style.color = 'var(--green)'; }
  }
  if (d.phase && d.phase !== 'idle') {
    document.getElementById('landing').style.display='none';
    document.getElementById('dashboard').style.display='block';
    document.getElementById('steps-panel').style.display='none';
    rebuildTables(d.collected || [], d.ignored || [], d.absent || []);
    document.getElementById('sv-total').textContent = d.total || 0;
    document.getElementById('sv-collected').textContent = (d.collected||[]).length;
    document.getElementById('sv-ignored').textContent = (d.ignored||[]).length;
    document.getElementById('sv-groups').textContent = d.groups && d.groups.length ? d.groups.join(', ') : '-';
    document.getElementById('tc-collected').textContent = (d.collected||[]).length;
    document.getElementById('tc-ignored').textContent = (d.ignored||[]).length;
    const pill = document.getElementById('status-pill');
    const ptxt = document.getElementById('status-text');
    if (d.phase==='done') { pill.className='status-pill done'; ptxt.textContent='Done'; document.getElementById('btn-save').disabled=false; document.getElementById('btn-stop').style.display='none'; }
    else if (d.phase==='error') { pill.className='status-pill error'; ptxt.textContent='Error'; document.getElementById('btn-save').disabled=false; }
    else { pill.className='status-pill live'; ptxt.textContent='Live'; }
    startTimer();
    listenSSE();
  }
});

document.getElementById('inp-url').addEventListener('keydown', function(e){
  if(e.key==='Enter') document.getElementById('inp-file').focus();
});
document.getElementById('inp-file').addEventListener('keydown', function(e){
  if(e.key==='Enter') startTracking();
});

// Generate scan bars
(function(){
  const bars = document.getElementById('scan-bars');
  for(let i=0;i<24;i++){
    const b = document.createElement('div');
    b.className='scan-bar';
    const h = 8 + Math.random()*28;
    b.style.height=h+'px';
    b.style.animationDelay=(Math.random()*0.5)+'s';
    b.style.animationDuration=(0.3+Math.random()*0.5)+'s';
    bars.appendChild(b);
  }
})();

function triggerScan(){
  const overlay = document.getElementById('scan-overlay');
  overlay.classList.add('active');
  setTimeout(()=>{
    overlay.classList.remove('active');
    overlay.style.display='none';
  }, 3500);
}
</script>
<footer style="text-align:center;padding:24px 0 16px;color:var(--text3);font-size:.72rem;letter-spacing:.02em">Created by Osman Mohammed (TheNooB)</footer>
</body>
</html>'''


@app.route('/')
def index():
    return render_template_string(TEMPLATE)


@app.route('/stream')
def stream():
    def gen():
        while True:
            try:
                ev = event_q.get(timeout=2)
                yield f"data: {json.dumps(ev)}\n\n"
            except queue.Empty:
                yield ": keepalive\n\n"
    return Response(gen(), mimetype='text/event-stream',
                    headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'})


@app.route('/upload-roster', methods=['POST'])
def upload_roster():
    global roster, roster_mode
    f = request.files.get('file')
    if not f or not f.filename:
        return jsonify(ok=False, error='No file')
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ('.txt', '.json', '.md', '.xlsx'):
        return jsonify(ok=False, error='Unsupported format. Use .txt, .json, .md, or .xlsx')
    try:
        raw = f.read()
        content = raw if ext == '.xlsx' else raw.decode('utf-8', errors='ignore')
        entries = parse_roster(content, f.filename)
        if not entries:
            return jsonify(ok=False, error='No names found in file')
        roster = entries
        roster_mode = True
        return jsonify(ok=True, count=len(entries),
                       names=[{'name': e['name'], 'student_id': e.get('student_id', '')} for e in entries])
    except Exception as e:
        return jsonify(ok=False, error=str(e))

@app.route('/clear-roster', methods=['POST'])
def clear_roster():
    global roster, roster_mode
    roster = []
    roster_mode = False
    state['absent'] = []
    return jsonify(ok=True)

@app.route('/start', methods=['POST'])
def start():
    global MY_GROUP
    data = request.json
    url = data.get('url', '').strip()
    fname = data.get('filename', '').strip()
    my_grp = data.get('group', '').strip().upper()
    if not url:
        return jsonify(ok=False, error='No URL')
    if not fname:
        fname = f"attendance_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    fname = re.sub(r'\.xlsx$', '', fname, flags=re.IGNORECASE)
    fname = re.sub(r'[/\\:*?"<>|]', '_', fname)

    if my_grp and len(my_grp) == 1 and my_grp.isalpha():
        MY_GROUP = my_grp

    state['meeting_url'] = url
    state['filename'] = fname
    state['stop'] = False

    t = threading.Thread(target=run_bot, args=(url,), daemon=True)
    t.start()
    return jsonify(ok=True)


@app.route('/stop', methods=['POST'])
def stop():
    state['stop'] = True
    return jsonify(ok=True)


_saved_path = [None]
_saved_csv = [None]

@app.route('/save', methods=['POST'])
def save():
    fname = state.get('filename', 'attendance')
    desktop = get_desktop()
    xlsx_path = os.path.join(desktop, fname + '.xlsx')
    csv_path = os.path.join(desktop, fname + '.csv')
    try:
        save_excel(xlsx_path)
        save_csv(csv_path)
        _saved_path[0] = xlsx_path
        _saved_csv[0] = csv_path
        return jsonify(ok=True, path=xlsx_path)
    except Exception as e:
        return jsonify(ok=False, error=str(e))


@app.route('/download')
def download():
    if _saved_path[0] and os.path.isfile(_saved_path[0]):
        return send_file(_saved_path[0], as_attachment=True)
    return "No file", 404

@app.route('/download-csv')
def download_csv():
    if _saved_csv[0] and os.path.isfile(_saved_csv[0]):
        return send_file(_saved_csv[0], as_attachment=True)
    return "No file", 404

@app.route('/state')
def get_state():
    c = state.get('collected', [])
    ig = state.get('ignored', [])
    mixed_val = bool(len(groups_found) > 1 or (groups_found and groups_found != {MY_GROUP}))
    ab = state.get('absent', [])
    return jsonify(
        phase=state.get('phase', 'idle'),
        total=state.get('total_in_meeting', 0),
        collected=[{'name':v['name'],'student_id':v.get('student_id',''),'group':v.get('group','')} for v in c],
        ignored=[{'name':v['name'],'reason':v.get('reason',''),'student_id':v.get('student_id',''),'group':v.get('group','')} for v in ig],
        absent=[{'name':v['name'],'student_id':v.get('student_id','')} for v in ab],
        groups=sorted(list(groups_found)) if groups_found else [],
        mixed=mixed_val,
        my_group=MY_GROUP,
        roster_mode=roster_mode,
        roster_count=len(roster),
    )

@app.route('/email', methods=['POST'])
def email_attendance():
    data = request.json
    recipient = (data.get('to', '') or '').strip()
    if not recipient or '@' not in recipient:
        return jsonify(ok=False, error='Enter a valid email address.')
    if not _saved_path[0] or not os.path.isfile(_saved_path[0]):
        return jsonify(ok=False, error='Save attendance first before sending.')
    fname = state.get('filename', 'attendance')
    ok, msg = send_attendance_email(recipient, _saved_path[0], fname)
    return jsonify(ok=ok, message=msg if ok else None, error=msg if not ok else None)

@app.route('/smtp-status')
def smtp_status():
    e, p = _load_smtp()
    return jsonify(configured=bool(e and p))


# ── MAIN ─────────────────────────────────────────────────────────────────────
def find_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('127.0.0.1', 0))
        return s.getsockname()[1]


if __name__ == '__main__':
    port = find_port()
    url = f'http://127.0.0.1:{port}'
    print(f"\n  Attendance Companion Dashboard")
    print(f"  Open: {url}\n")
    threading.Timer(1.5, lambda: webbrowser.open(url)).start()
    app.run(host='127.0.0.1', port=port, debug=False, threaded=True)
